# -*- coding: utf-8 -*-
"""
bidding_backtest_v3_fixed7.py  (실패도 기록하는 run_backtest 패치 포함 완성본)

- 낙찰정보리스트(입찰 결과 리스트) 엑셀을 읽어서
- history_root 아래 히스토리(발주기관/업종 기준) 파일을 매칭하고
- bidding_cal_regime_shrink.py (또는 호환 cal 스크립트)를 subprocess로 호출하여
  예가/기초 사정률 예측(y_pred)을 만들고
- 실제 낙찰 결과(y_true=예가/기초(100%) 등)와 비교하는 backtest 결과를 xlsx로 저장합니다.

핵심 개선
1) 입력 엑셀: 여러 시트/헤더행(0~3) 자동 스캔 -> required cols 만족/행 수 최대 시트 자동 선택
2) df0가 비어 있으면 즉시 중단(0행 저장 방지) + 선택된 시트/헤더 로그 출력
3) 실패(no_match/pred_error)도 detail에 무조건 남김 (사람 눈으로 안 찾아도 됨)
4) pred_error/no_match 원인을 자동 분류하여 Top-N 집계
5) cal 호출 결과 캐시(동일 history_file + open_month 조합은 1회만 실행) -> 응답 속도 개선
6) pivot 저장 시 MultiIndex 방지(엑셀 NotImplementedError 회피)
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import difflib
import numpy as np
import pandas as pd


# -----------------------------
# Util: robust excel reader
# -----------------------------
def _read_excel_one(path: str, sheet_name=None, header: int = 0) -> pd.DataFrame | Dict[str, pd.DataFrame]:
    ext = Path(path).suffix.lower()
    kwargs = dict(sheet_name=sheet_name, header=header)

    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return pd.read_excel(path, engine="openpyxl", **kwargs)
    if ext == ".xls":
        # .xls는 손상/확장자 위장 케이스가 많지만, 여기선 사용자가 이미 xlsx로 맞췄다고 가정
        return pd.read_excel(path, engine="xlrd", **kwargs)
    return pd.read_excel(path, **kwargs)


@dataclass
class ExcelPickInfo:
    sheet: str
    header: int
    n_rows: int
    matched_cols: int


def read_excel_any_best(path: str, required_cols: List[str], header_max: int = 3) -> Tuple[pd.DataFrame, ExcelPickInfo]:
    """
    - sheet_name=None으로 모든 시트를 dict로 읽고
    - header=0..header_max 스캔
    - required_cols를 가장 많이 포함하고, 행 수가 많은 시트/헤더 조합을 선택
    """
    required_cols = [c for c in required_cols if c and str(c).strip()]
    best_df: Optional[pd.DataFrame] = None
    best_info: Optional[ExcelPickInfo] = None
    best_score = (-1, -1, -1)  # (matched_cols, n_rows, n_cols)

    # sheet list 확보
    sheet_names: List[str] = []
    sheets_dict = None
    for h in range(0, header_max + 1):
        try:
            sheets_dict = _read_excel_one(path, sheet_name=None, header=h)
            if isinstance(sheets_dict, dict) and len(sheets_dict) > 0:
                sheet_names = list(sheets_dict.keys())
                break
        except Exception:
            continue

    if not sheet_names:
        df = _read_excel_one(path, sheet_name=0, header=0)  # type: ignore
        df.columns = [str(c).strip() for c in df.columns]
        info = ExcelPickInfo(sheet="0", header=0, n_rows=int(df.shape[0]), matched_cols=sum(c in df.columns for c in required_cols))
        return df, info

    # 스캔
    for sh in sheet_names:
        for h in range(0, header_max + 1):
            try:
                df = _read_excel_one(path, sheet_name=sh, header=h)  # type: ignore
                if df is None:
                    continue
                df.columns = [str(c).strip() for c in df.columns]
                n_rows = int(df.shape[0])
                if n_rows <= 0:
                    continue
                cols_set = set(df.columns)
                matched = sum(1 for c in required_cols if c in cols_set) if required_cols else 0
                score = (matched, n_rows, int(df.shape[1]))
                if score > best_score:
                    best_score = score
                    best_df = df
                    best_info = ExcelPickInfo(sheet=str(sh), header=h, n_rows=n_rows, matched_cols=matched)
            except Exception:
                continue

    if best_df is None or best_info is None:
        df = _read_excel_one(path, sheet_name=sheet_names[0], header=0)  # type: ignore
        df.columns = [str(c).strip() for c in df.columns]
        info = ExcelPickInfo(sheet=str(sheet_names[0]), header=0, n_rows=int(df.shape[0]), matched_cols=sum(c in df.columns for c in required_cols))
        return df, info

    return best_df, best_info


# -----------------------------
# Util: datetime parse
# -----------------------------
def to_datetime_series(s: pd.Series, tag: str = "") -> pd.Series:
    dtv = pd.to_datetime(s, errors="coerce")
    nat_ratio = float(dtv.isna().mean()) if len(dtv) else 1.0
    print(f"[date parse{(':'+tag) if tag else ''}] NaT ratio = {nat_ratio:.2%}")
    return dtv


# -----------------------------
# y scale align
# -----------------------------
def align_y_scale(y_raw: pd.Series, mode: str = "auto") -> Tuple[pd.Series, str]:
    """
    낙찰정보리스트의 y_col이 99.xx 형태(=100% 기준)로 들어오는데
    예측기는 보통 -0.xx(=사정률) 형태를 쓰는 경우가 많아 자동 정렬.

    mode:
      - "none": 그대로 사용
      - "minus100": y-100
      - "auto": median이 50보다 크면 minus100 선택
    """
    y = pd.to_numeric(y_raw, errors="coerce")
    if mode == "none":
        return y, "none"

    med = float(np.nanmedian(y.values)) if np.isfinite(np.nanmedian(y.values)) else np.nan
    if mode == "minus100":
        return y - 100.0, "minus100"

    # auto
    if np.isfinite(med) and med > 50.0:
        eff = "minus100"
        y2 = y - 100.0
    else:
        eff = "none"
        y2 = y

    print(f"[y_scale] mode=auto | median_raw={med:.6f} -> mode_eff={eff}")
    return y2, eff


# -----------------------------
# History indexing + matching
# -----------------------------
def norm_text(s: Any) -> str:
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return ""
    s = str(s).strip()
    s = s.replace("\u3000", " ")
    # 괄호와 그 안의 내용, 특수문자 등을 제거하여 핵심 키워드만 남김
    s = re.sub(r"[\s_()\[\]]+", "", s) 
    s = s.replace("㈜", "").replace("주식회사", "")
    return s


def issuer_level1(issuer: str) -> str:
    toks = [t for t in norm_text(issuer).split("_") if t]
    return toks[0] if toks else ""


def issuer_level12(issuer: str) -> str:
    toks = [t for t in norm_text(issuer).split("_") if t]
    if not toks:
        return ""
    if len(toks) == 1:
        return toks[0]
    return toks[0] + "_" + toks[1]


def trade_norm(trade: str) -> str:
    """
    업종: 토목 계열/조경 계열 표준화
    - 사용자 규칙: 토목=토목/토건/지반조성 포함, 조경=조경/조경식재/조경시설물 포함
    """
    t = norm_text(trade)
    if any(k in t for k in ["조경", "식재", "시설물"]):
        return "조경"
    if any(k in t for k in ["토목", "토건", "지반"]):
        return "토목"
    return t  # 기타는 원문 유지


@dataclass
class HistoryFile:
    path: str
    folder: str
    stem: str
    issuer_prefix: str
    trades: List[str]


def parse_history_filename(p: Path) -> Tuple[str, List[str]]:
    """
    예: 성남시_조경_토목_예가분석_20251211.xlsx
        issuer_prefix = 성남시
        trades = ['조경','토목']
    예: 농어촌공사_여주이천지사_토목_예가분석_20251217.xlsx
        issuer_prefix = 농어촌공사_여주이천지사
        trades = ['토목']
    """
    stem = p.stem
    toks = stem.split("_")
    if "예가분석" in toks:
        idx = toks.index("예가분석")
        core = toks[:idx]
    else:
        core = toks

    trade_tokens = []
    for tk in core[1:]:
        tn = trade_norm(tk)
        if tn in ("토목", "조경"):
            trade_tokens.append(tn)

    issuer_parts = []
    for tk in core:
        tn = trade_norm(tk)
        if tn in ("토목", "조경"):
            break
        issuer_parts.append(tk)
    issuer_prefix = "_".join([norm_text(x) for x in issuer_parts if x])
    if not issuer_prefix and core:
        issuer_prefix = norm_text(core[0])

    if not trade_tokens:
        trade_tokens = [""]

    return issuer_prefix, sorted(list(set(trade_tokens)))


def index_history(history_root: str) -> List[HistoryFile]:
    root = Path(history_root)
    files = []
    for p in root.rglob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        issuer_prefix, trades = parse_history_filename(p)
        files.append(
            HistoryFile(
                path=str(p),
                folder=norm_text(p.parent.name),
                stem=norm_text(p.stem),
                issuer_prefix=issuer_prefix,
                trades=trades,
            )
        )
    return files


def fuzzy_best(target: str, candidates: List[str], cutoff: float = 0.70, gap: float = 0.08) -> Tuple[Optional[str], float, float]:
    target = norm_text(target)
    if not target or not candidates:
        return None, 0.0, 0.0

    scored = []
    for c in candidates:
        sc = difflib.SequenceMatcher(None, target, c).ratio()
        scored.append((c, sc))
    scored.sort(key=lambda x: x[1], reverse=True)
    best_c, best_sc = scored[0]
    second = scored[1][1] if len(scored) >= 2 else 0.0
    if best_sc >= cutoff and (best_sc - second) >= gap:
        return best_c, best_sc, (best_sc - second)
    return None, best_sc, (best_sc - second)


def match_history(
    issuer: str,
    trade: str,
    hist: List[HistoryFile],
    fuzzy_cutoff: float,
    fuzzy_gap: float,
    fallback_level: int,
) -> Tuple[Optional[HistoryFile], Dict]:
    """
    fallback_level:
      0: issuer(L1~L2) + trade 고려
      1: issuer(L1) + trade 고려
      2: issuer(L1~L2)만( trade 무시 )
      3: folder명 보조키(issuer) + trade 고려
      4: folder명 보조키만( trade 무시 )
    """
    issuer_raw = str(issuer or "")
    trade_raw = str(trade or "")
    tn = trade_norm(trade_raw)

    key12 = issuer_level12(issuer_raw)
    key1 = issuer_level1(issuer_raw)
    folder_key = norm_text(issuer_raw)

    meta = {"issuer_raw": issuer_raw, "trade_raw": trade_raw, "trade_norm": tn, "fallback_level": fallback_level}

    def _filter_candidates(by_folder: bool, ignore_trade: bool, use_l1_only: bool) -> List[HistoryFile]:
        res = []
        for hf in hist:
            issuer_ok = False
            if by_folder:
                if folder_key and (folder_key in hf.folder or folder_key in hf.issuer_prefix):
                    issuer_ok = True
            else:
                if use_l1_only:
                    issuer_ok = bool(key1) and (hf.issuer_prefix.startswith(key1))
                else:
                    issuer_ok = bool(key12) and (hf.issuer_prefix.startswith(key12))

            if not issuer_ok:
                continue

            if ignore_trade:
                res.append(hf)
            else:
                # 1. 완전 일치 확인
                is_match = (tn in hf.trades or "" in hf.trades)
                
                # 2. 부분 일치 확인 (파일명의 단어가 업종명에 포함되는지)
                if not is_match:
                    for ht in hf.trades:
                        if len(ht) >= 2 and (ht in trade_raw or trade_raw in ht):
                            is_match = True
                            break
                
                if is_match:
                    res.append(hf)
        return res

    if fallback_level == 0:
        cands = _filter_candidates(by_folder=False, ignore_trade=False, use_l1_only=False)
    elif fallback_level == 1:
        cands = _filter_candidates(by_folder=False, ignore_trade=False, use_l1_only=True)
    elif fallback_level == 2:
        cands = _filter_candidates(by_folder=False, ignore_trade=True, use_l1_only=False)
    elif fallback_level == 3:
        cands = _filter_candidates(by_folder=True, ignore_trade=False, use_l1_only=False)
    else:
        cands = _filter_candidates(by_folder=True, ignore_trade=True, use_l1_only=False)

    if not cands:
        meta["match_reason"] = "no_candidate_after_filter"
        return None, meta

    cand_keys = []
    cand_map = {}
    for hf in cands:
        ck = norm_text(hf.issuer_prefix)
        cand_keys.append(ck)
        cand_map[ck] = hf

    best_key, best_sc, best_gap = fuzzy_best(issuer_raw, cand_keys, cutoff=fuzzy_cutoff, gap=fuzzy_gap)
    meta.update({"fuzzy_best_score": best_sc, "fuzzy_gap": best_gap, "fuzzy_cutoff": fuzzy_cutoff})

    if best_key is None:
        meta["match_reason"] = "fuzzy_reject"
        return None, meta

    meta["match_reason"] = "ok"
    return cand_map[best_key], meta


# -----------------------------
# Cal engine (subprocess + parse) + cache
# -----------------------------
@dataclass
class CalResult:
    ok: bool
    y_pred: float = np.nan
    stdout: str = ""
    stderr: str = ""
    rc: int = -1
    reason: str = ""


class CalEngine:
    """
    cal_path(=bidding_cal_regime_shrink.py)를 subprocess로 실행.
    stdout에서 예측값을 파싱.
    동일 (history_file, open_month) 호출은 캐시로 1회만 실행.
    """

    def __init__(
        self,
        cal_path: str,
        weight_mode: str,
        stack_scope: str,
        stack_loss: str,
        stack_lambda: float,
        stack_min_samples: int,
        hard_min_history: int,
        regime_on: str,
        shrink_on: str,
    ):
        self.cal_path = cal_path
        self.weight_mode = weight_mode
        self.stack_scope = stack_scope
        self.stack_loss = stack_loss
        self.stack_lambda = stack_lambda
        self.stack_min_samples = stack_min_samples
        self.hard_min_history = hard_min_history
        self.regime_on = regime_on
        self.shrink_on = shrink_on
        self.cache: Dict[Tuple[str, int], CalResult] = {}

    def predict(self, history_file: str, open_month: int) -> CalResult:
        key = (history_file, int(open_month))
        if key in self.cache:
            return self.cache[key]
        # 특수문자가 제거된 안전한 업종명 생성 (인자 전달 오류 방지)
        safe_trade = re.sub(r'[^\w\s]', '', str(key[0])).strip()

        cmd = [
            sys.executable,
            self.cal_path,
            "--file", str(history_file),
            "--open_month", str(int(open_month)),
            "--trade", safe_trade,
            "--weight_mode", self.weight_mode,
            "--stack_scope", self.stack_scope,
            "--stack_loss", self.stack_loss,
            "--stack_lambda", str(self.stack_lambda),
            "--stack_min_samples", str(int(self.stack_min_samples)),
            "--hard_min_history", str(int(self.hard_min_history)),
            "--regime_on", self.regime_on,
            "--shrink_on", self.shrink_on,
        ]

        try:
            p = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                shell=False,
            )
            out = (p.stdout or "").strip()
            err = (p.stderr or "").strip()
            rc = int(p.returncode)
        except Exception as e:
            r = CalResult(ok=False, rc=-1, reason=f"subprocess_exception:{type(e).__name__}:{e}")
            self.cache[key] = r
            return r

        if rc != 0:
            first = err.splitlines()[0] if err else ""
            r = CalResult(ok=False, rc=rc, stdout=out, stderr=err, reason=f"cal_rc{rc}:{first[:120]}")
            self.cache[key] = r
            return r

        y_pred = np.nan

        # 1) "권장" 라인 우선
        m = re.search(r"권장.*?([-+]?\d+(?:\.\d+)?)", out)
        if m:
            try:
                y_pred = float(m.group(1))
            except Exception:
                y_pred = np.nan

        # 2) 없으면 전체에서 abs가 작은 숫자 우선
        if not np.isfinite(y_pred):
            nums = re.findall(r"[-+]?\d+\.\d+|[-+]?\d+", out)
            cand = []
            for s in nums:
                try:
                    v = float(s)
                    # N=117 같은 큰 수는 배제하려고 abs<50 우선순위
                    cand.append(v)
                except Exception:
                    continue
            if cand:
                cand.sort(key=lambda v: abs(v))
                y_pred = float(cand[0])

        if not np.isfinite(y_pred):
            r = CalResult(ok=False, rc=0, stdout=out, stderr=err, reason="parse_fail_no_float")
            self.cache[key] = r
            return r

        r = CalResult(ok=True, rc=0, y_pred=float(y_pred), stdout=out, stderr=err, reason="ok")
        self.cache[key] = r
        return r


# -----------------------------
# Backtest core
# -----------------------------
def open_month_from_date(d: pd.Timestamp) -> int:
    if pd.isna(d):
        return 0
    return int(d.month)


def classify_pred_error(reason: str) -> str:
    if not reason:
        return "unknown"
    r = reason.lower()
    if "no_candidate_after_filter" in r or "fuzzy_reject" in r:
        return "match_fail"
    if "parse_fail" in r:
        return "parse_fail"
    if "cal_rc" in r:
        return "cal_runtime"
    if "history" in r and "min" in r:
        return "history_short"
    if "argument" in r or "unrecognized" in r:
        return "arg_error"
    return "other"


def run_backtest(
    df0: pd.DataFrame,
    issuer_col: str,
    trade_col: str,
    date_col: str,
    y_col: str,
    y_scale: str,
    history_root: str,
    cal_path: str,
    delta: float,
    fuzzy_cutoff: float,
    fuzzy_gap: float,
    fallback_on: bool,
    auto_tune: bool,
    auto_tune_passes: int,
    # cal opts
    weight_mode: str,
    stack_scope: str,
    stack_loss: str,
    stack_lambda: float,
    stack_min_samples: int,
    hard_min_history: int,
    regime_on: str,
    shrink_on: str,
) -> Tuple[pd.DataFrame, pd.DataFrame]:

    required = [issuer_col, trade_col, date_col, y_col]
    for c in required:
        if c not in df0.columns:
            raise RuntimeError(f"입력 엑셀에 컬럼이 없습니다: {c} (현재 컬럼 일부={list(df0.columns)[:30]})")

    df = df0.copy()
    df[date_col] = to_datetime_series(df[date_col], tag="bidlist")
    y_aligned, y_scale_eff = align_y_scale(df[y_col], mode=y_scale)
    df["_y_true"] = y_aligned

    if len(df) <= 0:
        raise RuntimeError("입력 엑셀에서 데이터 행을 찾지 못했습니다(0행).")

    hist = index_history(history_root)
    print(f"[history] files={len(hist)} | issuers={len(set(h.issuer_prefix for h in hist))}")

    engine = CalEngine(
        cal_path=cal_path,
        weight_mode=weight_mode,
        stack_scope=stack_scope,
        stack_loss=stack_loss,
        stack_lambda=stack_lambda,
        stack_min_samples=stack_min_samples,
        hard_min_history=hard_min_history,
        regime_on=regime_on,
        shrink_on=shrink_on,
    )

    cutoff_eff = float(fuzzy_cutoff)
    gap_eff = float(fuzzy_gap)
    fallback_level_used = 0

    def _one_pass(cutoff_now: float, gap_now: float, base_level: int) -> pd.DataFrame:
        rows = []
        for _, r in df.iterrows():
            issuer = r.get(issuer_col, "")
            trade = r.get(trade_col, "")
            dtt = r.get(date_col, pd.NaT)
            y_true = r.get("_y_true", np.nan)
            om = open_month_from_date(dtt)

            matched_hf = None
            meta = {}
            levels = [base_level]
            if fallback_on:
                levels = [base_level, 1, 2, 3, 4]

            for lv in levels:
                hf, meta_lv = match_history(
                    issuer=str(issuer),
                    trade=str(trade),
                    hist=hist,
                    fuzzy_cutoff=cutoff_now,
                    fuzzy_gap=gap_now,
                    fallback_level=lv,
                )
                if hf is not None:
                    matched_hf = hf
                    meta = meta_lv
                    break
                meta = meta_lv

            # (패치1) no_match도 detail에 반드시 기록
            if matched_hf is None:
                rows.append(
                    {
                        "issuer": issuer,
                        "trade": trade,
                        "date": dtt,
                        "open_month": om,
                        "y_true": y_true,
                        "y_pred": np.nan,
                        "abs_err": np.nan,
                        "hit": np.nan,
                        "status": "no_match",
                        "pred_reason": meta.get("match_reason", "no_match"),
                        "match_meta": json.dumps(meta, ensure_ascii=False),
                        "history_file": "",
                        "cal_rc": np.nan,
                        "cal_err_head": "",
                    }
                )
                continue

            # cal 호출
            cres = engine.predict(matched_hf.path, om)

            # (패치2) pred_error도 detail에 반드시 기록(+rc/에러헤더)
            if not cres.ok:
                err_head = (cres.stderr.splitlines()[0] if cres.stderr else "")[:180]
                rows.append(
                    {
                        "issuer": issuer,
                        "trade": trade,
                        "date": dtt,
                        "open_month": om,
                        "y_true": y_true,
                        "y_pred": np.nan,
                        "abs_err": np.nan,
                        "hit": np.nan,
                        "status": "pred_error",
                        "pred_reason": cres.reason,
                        "match_meta": json.dumps(meta, ensure_ascii=False),
                        "history_file": matched_hf.path,
                        "cal_rc": cres.rc,
                        "cal_err_head": err_head,
                    }
                )
                continue

            y_pred = float(cres.y_pred)
            abs_err = float(abs(y_pred - float(y_true))) if np.isfinite(y_true) else np.nan
            hit = 1.0 if (np.isfinite(abs_err) and abs_err <= delta) else 0.0

            rows.append(
                {
                    "issuer": issuer,
                    "trade": trade,
                    "date": dtt,
                    "open_month": om,
                    "y_true": y_true,
                    "y_pred": y_pred,
                    "abs_err": abs_err,
                    "hit": hit,
                    "status": "ok",
                    "pred_reason": "ok",
                    "match_meta": json.dumps(meta, ensure_ascii=False),
                    "history_file": matched_hf.path,
                    "cal_rc": 0,
                    "cal_err_head": "",
                }
            )

        return pd.DataFrame(rows)

    # auto_tune 루프: n_pred 최대화 방향으로 cutoff/gap/레벨 완화
    detail_best = None
    best_pred = -1
    best_tuple = (cutoff_eff, gap_eff, fallback_level_used)

    passes = int(auto_tune_passes) if auto_tune else 1
    for _ in range(passes):
        detail = _one_pass(cutoff_eff, gap_eff, fallback_level_used)
        n_pred = int(np.isfinite(detail["y_pred"]).sum())
        if n_pred > best_pred:
            best_pred = n_pred
            detail_best = detail
            best_tuple = (cutoff_eff, gap_eff, fallback_level_used)

        if not auto_tune:
            break

        no_match_ratio = float((detail["status"] == "no_match").mean())
        if no_match_ratio > 0.30:
            cutoff_eff = max(0.55, cutoff_eff - 0.05)
            gap_eff = max(0.02, gap_eff - 0.02)
            fallback_level_used = min(2, fallback_level_used + 1)
        else:
            break

    detail = detail_best if detail_best is not None else _one_pass(fuzzy_cutoff, fuzzy_gap, 0)
    cutoff_eff, gap_eff, fallback_level_used = best_tuple

    # summary
    n_rows = int(len(detail))
    n_pred = int(np.isfinite(detail["y_pred"]).sum())
    fail_total = int((detail["status"] != "ok").sum())
    mae = float(np.nanmean(detail["abs_err"].values)) if n_pred > 0 else np.nan
    hit_rate = float(np.nanmean(detail["hit"].values)) if n_pred > 0 else np.nan

    pred_err = detail.loc[detail["status"] == "pred_error", "pred_reason"].fillna("").astype(str)
    no_match = detail.loc[detail["status"] == "no_match", "pred_reason"].fillna("").astype(str)

    pred_err_group = pred_err.apply(classify_pred_error).value_counts().to_dict()
    no_match_group = no_match.value_counts().to_dict()

    summary = pd.DataFrame(
        [
            {
                "n_rows": n_rows,
                "n_pred": n_pred,
                "fail_total": fail_total,
                "mae": mae,
                "hit@delta": hit_rate,
                "delta": float(delta),
                "y_scale_eff": y_scale_eff,
                "fuzzy_cutoff_eff": float(cutoff_eff),
                "fuzzy_gap_eff": float(gap_eff),
                "fallback_level_used": int(fallback_level_used),
                "fallback_used": bool(fallback_on),
                "auto_tune": bool(auto_tune),
                "history_files": int(len(hist)),
                "status_ok": int((detail["status"] == "ok").sum()),
                "status_pred_error": int((detail["status"] == "pred_error").sum()),
                "status_no_match": int((detail["status"] == "no_match").sum()),
                "pred_error_top": json.dumps(pred_err_group, ensure_ascii=False),
                "no_match_top": json.dumps(no_match_group, ensure_ascii=False),
                "run_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ]
    )

    return detail, summary


# -----------------------------
# Excel output (no MultiIndex)
# -----------------------------
def save_excel(out_xlsx: str, detail: pd.DataFrame, summary: pd.DataFrame) -> None:
    out_path = Path(out_xlsx)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        summary.to_excel(w, index=False, sheet_name="summary")
        detail.to_excel(w, index=False, sheet_name="detail")

        if len(detail) > 0:
            piv1 = (
                detail.pivot_table(
                    index="issuer",
                    values=["abs_err", "hit"],
                    aggfunc={"abs_err": "mean", "hit": "mean"},
                )
                .reset_index()
            )
            piv1.to_excel(w, index=False, sheet_name="pivot_issuer")

            piv2 = (
                detail.pivot_table(
                    index="trade",
                    values=["abs_err", "hit"],
                    aggfunc={"abs_err": "mean", "hit": "mean"},
                )
                .reset_index()
            )
            piv2.to_excel(w, index=False, sheet_name="pivot_trade")


# -----------------------------
# CLI
# -----------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--history_root", required=True)
    ap.add_argument("--cal_path", required=True)
    ap.add_argument("--infile", required=True)
    ap.add_argument("--out_xlsx", required=True)

    ap.add_argument("--issuer_col", default="발주기관")
    ap.add_argument("--trade_col", default="업종")
    ap.add_argument("--date_col", default="개찰일")
    ap.add_argument("--y_col", default="예가/기초(100%)")

    ap.add_argument("--y_scale", default="auto", choices=["auto", "none", "minus100"])
    ap.add_argument("--delta", type=float, default=0.03)

    ap.add_argument("--fuzzy_cutoff", type=float, default=0.70)
    ap.add_argument("--fuzzy_gap", type=float, default=0.08)

    ap.add_argument("--fallback_on", default="on", choices=["on", "off"])
    ap.add_argument("--auto_tune", default="on", choices=["on", "off"])
    ap.add_argument("--auto_tune_passes", type=int, default=3)

    # cal opts
    ap.add_argument("--weight_mode", default="stack")
    ap.add_argument("--stack_scope", default="all")
    ap.add_argument("--stack_loss", default="huber")
    ap.add_argument("--stack_lambda", type=float, default=0.6)
    ap.add_argument("--stack_min_samples", type=int, default=60)
    ap.add_argument("--hard_min_history", type=int, default=10)
    ap.add_argument("--regime_on", default="on", choices=["on", "off"])
    ap.add_argument("--shrink_on", default="on", choices=["on", "off"])

    args = ap.parse_args()

    required_cols = [args.issuer_col, args.trade_col, args.date_col, args.y_col]
    df0, pick = read_excel_any_best(args.infile, required_cols=required_cols, header_max=3)
    print(f"[excel pick] sheet={pick.sheet} header={pick.header} rows={pick.n_rows} matched_cols={pick.matched_cols}")
    df0.columns = [str(c).strip() for c in df0.columns]

    if len(df0) <= 0:
        raise RuntimeError(
            "입력 엑셀에서 데이터 행을 찾지 못했습니다(0행).\n"
            "- 엑셀에 여러 시트가 있거나 상단 제목행 때문에 데이터가 다른 위치에 있을 수 있습니다.\n"
            "- 엑셀에서 데이터 시트만 남기거나, 첫 행을 컬럼 헤더로 맞춘 뒤 재시도해 주세요."
        )

    detail, summary = run_backtest(
        df0=df0,
        issuer_col=args.issuer_col,
        trade_col=args.trade_col,
        date_col=args.date_col,
        y_col=args.y_col,
        y_scale=args.y_scale,
        history_root=args.history_root,
        cal_path=args.cal_path,
        delta=float(args.delta),
        fuzzy_cutoff=float(args.fuzzy_cutoff),
        fuzzy_gap=float(args.fuzzy_gap),
        fallback_on=(args.fallback_on == "on"),
        auto_tune=(args.auto_tune == "on"),
        auto_tune_passes=int(args.auto_tune_passes),
        weight_mode=args.weight_mode,
        stack_scope=args.stack_scope,
        stack_loss=args.stack_loss,
        stack_lambda=float(args.stack_lambda),
        stack_min_samples=int(args.stack_min_samples),
        hard_min_history=int(args.hard_min_history),
        regime_on=args.regime_on,
        shrink_on=args.shrink_on,
    )

    save_excel(args.out_xlsx, detail, summary)

    print(f"\n[Backtest v3 saved] {args.out_xlsx}")
    print(summary.to_string(index=False))


if __name__ == "__main__":
    main()