# -*- coding: utf-8 -*-
"""
리더십 평가 통합 점수 산출기 (가중 통합안, –6~+6 유지)
- 구성: 초과근무(–3~+3) + 연차(–3~+3) + 직원효율화(–3~+3)
- 최종 점수 = w_ot*OT + w_lv*연차 + w_eff*효율, 기본 (0.4, 0.3, 0.3)
- 효율 산정: 업로드 파일(예: 실센터장_리더십_평가_2025년_2025-10-15.xlsx)에서
  [인원, 총 프로젝트 수, 실제 달성점수] 자동 감지 → 팀 단위 집계 → CE 산출 → Residual(%) → 점수화
"""

import os
import re
import pandas as pd
import numpy as np
from datetime import datetime

# -----------------------
# 공통 설정
# -----------------------
HOURS_PER_DAY = 8.0
BASE_MONTHLY_HOURS = 160
RANK_WEIGHTS = {
    "책임": 1.2,
    "선임": 1.0,
    "주임": 0.9,
    "원급": 0.8,
}

# -----------------------
# 공통 유틸
# -----------------------
def _norm(s):
    return str(s).strip() if pd.notna(s) else ""

def parse_hhmm_to_minutes(x):
    if pd.isna(x):
        return 0
    if isinstance(x, str) and ":" in x:
        try:
            h, m = x.strip().split(":")
            return int(h) * 60 + int(m)
        except Exception:
            return 0
    if isinstance(x, (int, float)):
        return int(x)
    return 0

def detect_months_from_dates(series):
    months = set()
    for v in series.dropna():
        try:
            dt = pd.to_datetime(v)
            months.add((dt.year, dt.month))
        except Exception:
            pass
    cnt = len(months) if months else 1
    return min(max(cnt, 1), 12)

def _min_max_norm(s):
    s = pd.to_numeric(s, errors="coerce")
    vmin, vmax = s.min(), s.max()
    if pd.isna(vmin) or pd.isna(vmax) or vmax - vmin == 0:
        return pd.Series([0.5]*len(s), index=s.index)
    return (s - vmin) / (vmax - vmin)

# 날짜/부서/시간 컬럼 자동 감지 (초과근무 쪽)
def _pick_date_col(df):
    name_candidates = [c for c in df.columns if any(k in str(c) for k in ["신청일", "신청 일", "일자", "근무일자", "기안일", "작성일"])]
    best_col, best_rate = None, -1.0
    for c in (name_candidates if name_candidates else list(df.columns)):
        parsed = pd.to_datetime(df[c], errors="coerce")
        rate = parsed.notna().mean()
        if rate > best_rate:
            best_rate, best_col = rate, c
    return best_col if best_rate > 0.2 else None

def _pick_dept_col_daily(df):
    for cand in ["기본정보", "부서", "소속"]:
        if cand in df.columns:
            return cand
    best_col, best_score = None, -1.0
    for c in df.columns:
        s = df[c].astype(str)
        score = s.apply(lambda x: x.strip() != "" and not x.strip().isdigit()).mean()
        if score > best_score:
            best_score, best_col = score, c
    return best_col

def _pick_time_col(df):
    for c in df.columns:
        if "환산" in str(c):
            return c
    best_col, best_rate = None, -1.0
    for c in df.columns:
        s = df[c].astype(str)
        rate = s.str.contains(r"^\s*\d{1,3}:\d{2}\s*$").mean()
        if rate > best_rate:
            best_rate, best_col = rate, c
    return best_col

# -----------------------
# 초과근무 → 점수(–3~+3)
# -----------------------
def residual_to_score_step(residual_pct: float) -> float:
    """
    공통 계단식 점수 매핑(OT/효율 공용): Residual(%) → –3~+3
    음수가 '좋음'(효율적/OT절감)이라는 철학(잔차가 음수일수록 가점) 유지
    """
    if residual_pct <= -10:
        return 3.0
    elif residual_pct <= -7:
        return 2.0
    elif residual_pct <= -4:
        return 1.0
    elif residual_pct <= -1:
        return 0.5
    elif residual_pct < 1:
        return 0.0
    elif residual_pct < 4:
        return -0.5
    elif residual_pct < 7:
        return -1.0
    elif residual_pct < 10:
        return -2.0
    else:
        return -3.0

def compute_overtime_score(excel_path_or_buffer, dept_filter: str | None = None):
    # Streamlit UploadedFile 대응
    if hasattr(excel_path_or_buffer, 'seek'):
        excel_path_or_buffer.seek(0)

    xls = pd.ExcelFile(excel_path_or_buffer)
    if "총계" not in xls.sheet_names or "일별현황_A" not in xls.sheet_names:
        raise ValueError("엑셀에 '총계' 또는 '일별현황_A' 시트가 없습니다.")

    # 총계
    df_sum = pd.read_excel(excel_path_or_buffer, sheet_name="총계")
    df_sum.columns = [str(c).strip() for c in df_sum.columns]
    need = ["부서", "직급", "이름"]
    miss = [c for c in need if c not in df_sum.columns]
    if miss:
        raise ValueError(f"총계 시트에 {miss} 컬럼이 없습니다.")
    sum_ot_col = "환산" if "환산" in df_sum.columns else _pick_time_col(df_sum)
    if not sum_ot_col:
        raise ValueError("총계 시트에서 초과근무 합계(환산) 열을 찾지 못했습니다.")

    tmp = df_sum.copy()
    tmp["__mins__"] = tmp[sum_ot_col].apply(parse_hhmm_to_minutes)
    tmp = tmp[pd.to_numeric(tmp["__mins__"], errors="coerce").notna()]
    tmp["부서"] = tmp["부서"].apply(_norm)
    tmp["직급"] = tmp["직급"].apply(_norm)

    # 일별현황_A
    df_daily = pd.read_excel(excel_path_or_buffer, sheet_name="일별현황_A")
    df_daily.columns = [str(c).strip() for c in df_daily.columns]
    date_col = _pick_date_col(df_daily)
    dept_col_daily = _pick_dept_col_daily(df_daily)
    ot_col_daily = _pick_time_col(df_daily)
    months_cnt = detect_months_from_dates(df_daily[date_col]) if date_col else 1
    df_daily["_부서"] = df_daily[dept_col_daily].apply(_norm)
    df_daily["_분"] = df_daily[ot_col_daily].apply(parse_hhmm_to_minutes)

    # AH(시간) = 인원 × 직급가중치 × 월160 × 월수
    tmp["_가중치"] = tmp["직급"].map(RANK_WEIGHTS).fillna(1.0)
    dept_rank = tmp.groupby(["부서", "직급"]).size().reset_index(name="인원수")
    dept_rank["_가중치"] = dept_rank["직급"].map(RANK_WEIGHTS).fillna(1.0)
    dept_rank["AH(시간)"] = dept_rank["인원수"] * dept_rank["_가중치"] * BASE_MONTHLY_HOURS * months_cnt
    dept_ah = dept_rank.groupby("부서")["AH(시간)"].sum().reset_index()

    dept_ot = tmp.groupby("부서")["__mins__"].sum().reset_index().rename(columns={"__mins__": "OT(분)"})
    dept_table = pd.merge(dept_ah, dept_ot, on="부서", how="outer").fillna(0.0)
    more_ot = df_daily.groupby("_부서")["_분"].sum().reset_index().rename(columns={"_부서": "부서", "_분": "OT_daily(분)"})
    dept_table = pd.merge(dept_table, more_ot, on="부서", how="outer").fillna(0.0)
    dept_table["OT(분)"] = dept_table[["OT(분)", "OT_daily(분)"]].max(axis=1)
    dept_table.drop(columns=["OT_daily(분)"], inplace=True)
    dept_table["OT(시간)"] = dept_table["OT(분)"] / 60.0

    # 전사 OT율/평균 AH
    org_total_ot = dept_table["OT(시간)"].sum()
    org_total_ah = dept_table["AH(시간)"].sum()
    if org_total_ah <= 0:
        raise ValueError("전사 AH(시간)가 0 이하입니다.")
    org_ot_rate = org_total_ot / (org_total_ot + org_total_ah)
    org_avg_ah = dept_table["AH(시간)"].replace(0, np.nan).mean()

    def calc_aor(ot_h, ah_h):
        denom = ot_h + ah_h
        return (ot_h / denom) if denom > 0 else 0.0

    def calc_eor(team_ah):
        if team_ah <= 0:
            return org_ot_rate * 100.0
        e = org_ot_rate * (org_avg_ah / team_ah) * 100.0
        return max(0.0, min(100.0, e))

    dept_table["AOR(%)"] = dept_table.apply(lambda r: calc_aor(r["OT(시간)"], r["AH(시간)"]) * 100.0, axis=1)
    dept_table["EOR(%)"] = dept_table["AH(시간)"].apply(calc_eor)
    dept_table["Residual(%)"] = dept_table["AOR(%)"] - dept_table["EOR(%)"]
    dept_table["OT점수"] = dept_table["Residual(%)"].apply(residual_to_score_step)

    ot_out = dept_table[["부서", "OT점수"]]
    if dept_filter:
        ot_out = ot_out[ot_out["부서"].str.contains(dept_filter, na=False)]
    return ot_out

# -----------------------
# 연차 → 점수(–3~+3)
# -----------------------
def _looks_hhmm(val: str) -> bool:
    return bool(re.match(r"^\s*\d{1,3}:\d{2}\s*$", str(val or "")))

def _to_days(x, hours_per_day=HOURS_PER_DAY):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if _looks_hhmm(s):
        h, m = s.split(":")
        hours = int(h) + int(m) / 60.0
        return hours / hours_per_day if hours_per_day > 0 else 0.0
    m = re.search(r"(\d+(\.\d+)?)", s)
    return float(m.group(1)) if m else 0.0

def compute_leave_score(excel_path_or_buffer, dept_filter: str | None = None, sheet_name: str | None = None):
    if hasattr(excel_path_or_buffer, 'seek'):
        excel_path_or_buffer.seek(0)

    xls = pd.ExcelFile(excel_path_or_buffer)
    use_sheet = sheet_name if (sheet_name and sheet_name in xls.sheet_names) else xls.sheet_names[0]

    df = pd.read_excel(excel_path_or_buffer, sheet_name=use_sheet)
    df.columns = [str(c).strip() for c in df.columns]

    def _rebuild_header_by_rowindex(idx: int):
        df2 = pd.read_excel(excel_path_or_buffer, sheet_name=use_sheet, header=idx)
        if isinstance(df2.columns, pd.MultiIndex):
            df2.columns = [" ".join([_norm(x) for x in tup if _norm(x)]).strip() for tup in df2.columns.to_list()]
        df2.columns = [c.replace("이름(ID)", "이름").strip() for c in df2.columns]
        return df2

    def _looks_broken_header(df0: pd.DataFrame) -> bool:
        unnamed_ratio = np.mean([c.startswith("Unnamed:") for c in df0.columns])
        has_keywords = any(k in " ".join(df0.columns) for k in ["부여", "사용", "잔여"])
        return (unnamed_ratio > 0.3) and (not has_keywords)

    if _looks_broken_header(df):
        raw = pd.read_excel(excel_path_or_buffer, sheet_name=use_sheet, header=None)
        candidate = None
        best_score = -1
        for i in range(min(6, len(raw))):
            row = raw.iloc[i].astype(str).fillna("")
            score = 0
            text = " ".join(row.tolist())
            if re.search(r"부서|소속|부서명|기본정보", text): score += 2
            if re.search(r"이름|성명|사원명|이름\(ID\)", text): score += 2
            if re.search(r"부여|사용|잔여|연차", text): score += 2
            unnamed_rate = np.mean([str(x).startswith("Unnamed") for x in row])
            score += max(0, 1.5 - unnamed_rate)
            if score > best_score:
                best_score, candidate = score, i
        df = _rebuild_header_by_rowindex(candidate if candidate is not None else 0)

    df.columns = [c.replace("이름(ID)", "이름").strip() for c in df.columns]

    def _pick_dept_col_leave(df_):
        for cand in ["부서", "소속", "부서명", "부서(소속)", "실/센터", "실센터", "본부", "기본정보"]:
            if cand in df_.columns: return cand
        best, score = None, -1.0
        for c in df_.columns:
            rate = df_[c].astype(str).apply(lambda x: x.strip() != "" and not x.strip().isdigit()).mean()
            if rate > score:
                score, best = rate, c
        return best

    def _pick_name_col_leave(df_):
        for cand in ["이름", "성명", "사원명", "직원명", "Name", "name", "이름(ID)"]:
            if cand in df_.columns: return cand
        best, score = None, -1
        for c in df_.columns:
            m = df_[c].astype(str).str.len().mean()
            if m > score:
                score, best = m, c
        return best

    def _pick_granted_col(df_):
        for cand in ["부여", "부여일수", "연차부여", "발생", "발생일수", "총연차", "연차(부여)", "부여(일수)"]:
            hits = [c for c in df_.columns if cand in c]
            if hits: return hits[0]
        for c in df_.columns:
            if df_[c].astype(str).head(6).str.contains("부여|발생|총연차").any():
                return c
        return None

    def _pick_used_col(df_):
        for cand in ["사용", "사용일수", "연차사용", "사용(일수)", "사용일", "사용수", "사용횟수"]:
            hits = [c for c in df_.columns if cand in c]
            if hits: return hits[0]
        for c in df_.columns:
            if df_[c].astype(str).head(6).str.contains("사용").any():
                return c
        return None

    def _pick_remaining_col(df_):
        for cand in ["잔여", "잔여일수", "잔여연차", "미사용", "미사용일수", "남은", "남은일수"]:
            hits = [c for c in df_.columns if cand in c]
            if hits: return hits[0]
        for c in df_.columns:
            if df_[c].astype(str).head(6).str.contains("잔여|미사용|남은").any():
                return c
        return None

    dept_col = _pick_dept_col_leave(df)
    name_col = _pick_name_col_leave(df)
    grant_col = _pick_granted_col(df)
    used_col = _pick_used_col(df)
    remain_col = _pick_remaining_col(df)

    have_cnt = sum([grant_col is not None, used_col is not None, remain_col is not None])
    if have_cnt < 2:
        raise ValueError("연차 시트에서 부여/사용/잔여 중 최소 2개를 찾지 못했습니다.")

    df["_부서"] = df[dept_col].apply(_norm)
    df["_이름"] = df[name_col].apply(_norm)

    if grant_col:  df["_부여"] = df[grant_col].apply(_to_days)
    if used_col:   df["_사용"] = df[used_col].apply(_to_days)
    if remain_col: df["_잔여"] = df[remain_col].apply(_to_days)

    if "_부여" not in df.columns and {"_사용", "_잔여"} <= set(df.columns):
        df["_부여"] = df["_사용"] + df["_잔여"]
    if "_사용" not in df.columns and {"_부여", "_잔여"} <= set(df.columns):
        df["_사용"] = df["_부여"] - df["_잔여"]
    if "_잔여" not in df.columns and {"_부여", "_사용"} <= set(df.columns):
        df["_잔여"] = df["_부여"] - df["_사용"]

    for c in ["_부여", "_사용", "_잔여"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
            df.loc[df[c] < 0, c] = 0.0

    grp = df.groupby("_부서", dropna=False).agg(
        부여합=("_부여", "sum"),
        사용합=("_사용", "sum"),
        잔여합=("_잔여", "sum"),
        인원수=("_이름", "count"),
    ).reset_index().rename(columns={"_부서": "부서"})

    grp["잔여율(%)"] = (grp["잔여합"] / grp["부여합"].replace(0, np.nan) * 100.0).fillna(0.0).clip(0, 100)

    def leave_score(remain_pct):
        if remain_pct <= 2: return 3.0
        elif remain_pct <= 5: return 2.0
        elif remain_pct <= 10: return 1.0
        elif remain_pct <= 20: return 0.0
        elif remain_pct <= 30: return -1.0
        elif remain_pct <= 40: return -2.0
        else: return -3.0

    grp["연차점수"] = grp["잔여율(%)"].apply(leave_score)

    if dept_filter:
        grp = grp[grp["부서"].str.contains(dept_filter, na=False)]
    return grp[["부서", "연차점수"]]

# -----------------------
# 직원효율화 → 점수(–3~+3)
# -----------------------
def _pick_efficiency_columns(df):
    # 그룹키
    dept_candidates = ["부서", "본부", "실/센터", "소속", "부서명"]
    dept_col = next((c for c in dept_candidates if c in df.columns), None)
    if dept_col is None:
        # 가장 텍스트 비율 높은 열을 부서대용
        best, score = None, -1
        for c in df.columns:
            rate = df[c].astype(str).apply(lambda x: x.strip() != "" and not x.strip().isdigit()).mean()
            if rate > score:
                score, best = rate, c
        dept_col = best

    # 인원
    ppl_candidates = ["소속 직원 수", "인원", "인원수", "직원수", "구성원수"]
    ppl_col = next((c for c in ppl_candidates if c in df.columns), None)

    # 총 프로젝트 수
    prj_candidates = ["총 프로젝트 수", "프로젝트 수", "프로젝트총계", "총프로젝트수"]
    prj_col = next((c for c in prj_candidates if c in df.columns), None)

    # 실제 달성점수
    score_candidates = ["실제 달성점수", "총점", "달성점수", "성과점수"]
    ach_col = next((c for c in score_candidates if c in df.columns), None)

    return dept_col, ppl_col, prj_col, ach_col

def _to_num(x):
    if pd.isna(x): return 0.0
    if isinstance(x, (int, float)): return float(x)
    m = re.search(r"-?\d+(\.\d+)?", str(x))
    return float(m.group()) if m else 0.0

def compute_efficiency_score(eff_file, dept_filter: str | None = None,
                             sheet_name: str | None = None,
                             mode: str = "CE", alpha: float = 0.5):
    # Streamlit UploadedFile 대응
    if hasattr(eff_file, 'seek'):
        eff_file.seek(0)

    xls = pd.ExcelFile(eff_file)
    use_sheet = sheet_name if (sheet_name and sheet_name in xls.sheet_names) else xls.sheet_names[0]
    df = pd.read_excel(eff_file, sheet_name=use_sheet)
    df.columns = [str(c).strip() for c in df.columns]

    dept_col, ppl_col, prj_col, ach_col = _pick_efficiency_columns(df)
    missing = []
    if dept_col is None: missing.append("부서/소속 구분 열")
    if ppl_col is None: missing.append("인원")
    if prj_col is None: missing.append("총 프로젝트 수")
    if ach_col is None: missing.append("실제 달성점수")
    if missing:
        raise ValueError(f"직원효율화 산정에 필요한 컬럼을 찾지 못했습니다: {missing}\n열 목록: {list(df.columns)}")

    work = pd.DataFrame({
        "부서": df[dept_col].apply(_norm),
        "인원": df[ppl_col].apply(_to_num),
        "총 프로젝트 수": df[prj_col].apply(_to_num),
        "실제 달성점수": df[ach_col].apply(_to_num),
    })
    work["인원"] = work["인원"].clip(lower=0)

    # 팀 단위 집계(부서 동일명 합산)
    grp = work.groupby("부서", dropna=False).agg(
        인원=("인원", "sum"),
        총프로젝트=("총 프로젝트 수", "sum"),
        실제달성=("실제 달성점수", "sum"),
    ).reset_index()

    # 인원 0 대비 안전 처리
    grp = grp[grp["인원"] > 0].copy()
    grp["PE"] = grp["총프로젝트"] / grp["인원"]
    grp["SE"] = grp["실제달성"] / grp["인원"]

    # 정규화
    grp["PE_norm"] = _min_max_norm(grp["PE"])
    grp["SE_norm"] = _min_max_norm(grp["SE"])

    # 효율 스코어
    mode = (mode or "CE").upper()
    if mode == "PE":
        grp["EFF_raw"] = grp["PE_norm"]
    elif mode == "SE":
        grp["EFF_raw"] = grp["SE_norm"]
    else:
        grp["EFF_raw"] = alpha * grp["PE_norm"] + (1 - alpha) * grp["SE_norm"]

    # AOR/EOR 방식(OT 철학 동일)
    org_eff_sum = grp["EFF_raw"].sum()
    org_people_sum = grp["인원"].sum()
    org_avg_people = grp["인원"].replace(0, np.nan).mean()

    def aor_eff(team_eff):
        # team_eff / (team_eff + 1) 를 확률형으로 해석 (스케일-독립적)
        denom = team_eff + 1.0
        return (team_eff / denom) if denom > 0 else 0.0

    org_eff_rate = aor_eff(org_eff_sum / max(len(grp), 1))  # 조직 평균수준으로 스케일링

    def eor_eff(team_people):
        if team_people <= 0:
            return org_eff_rate * 100.0
        e = org_eff_rate * (org_avg_people / team_people) * 100.0
        return max(0.0, min(100.0, e))

    grp["AOR_EFF(%)"] = grp["EFF_raw"].apply(lambda x: aor_eff(x) * 100.0)
    grp["EOR_EFF(%)"] = grp["인원"].apply(eor_eff)
    grp["Residual_EFF(%)"] = grp["AOR_EFF(%)"] - grp["EOR_EFF(%)"]

    # 점수화(–3~+3)
    grp["직원효율화점수"] = grp["Residual_EFF(%)"].apply(residual_to_score_step)

    if dept_filter:
        grp = grp[grp["부서"].str.contains(dept_filter, na=False)]

    # 참고치 포함 반환
    return grp[["부서", "직원효율화점수", "PE", "SE", "인원", "총프로젝트", "실제달성"]].rename(
        columns={"PE": "참고(PE)", "SE": "참고(SE)", "총프로젝트": "총 프로젝트 수", "실제달성": "실제 달성점수"}
    )

# -----------------------
# 통합 최종 계산(가중 통합)
# -----------------------
def compute_total_score(overtime_file,
                        leave_file,
                        efficiency_file=None,
                        dept_filter: str | None = None,
                        leave_sheet: str | None = None,
                        eff_sheet: str | None = None,
                        mode: str = "CE",
                        alpha: float = 0.5,
                        weights: tuple[float, float, float] = (0.4, 0.3, 0.3)):
    """
    최종 –6~+6 유지: w_ot*OT + w_lv*연차 + w_eff*효율  (각 항목 –3~+3)
    efficiency_file 미제공 시 효율=0점 처리
    """
    w_ot, w_lv, w_eff = weights
    if not np.isclose(w_ot + w_lv + w_eff, 1.0):
        raise ValueError(f"weights 합은 1이어야 합니다. 현재 합={w_ot + w_lv + w_eff}")

    ot = compute_overtime_score(overtime_file, dept_filter)
    lv = compute_leave_score(leave_file, dept_filter, sheet_name=leave_sheet)

    if efficiency_file is not None:
        eff = compute_efficiency_score(efficiency_file, dept_filter, sheet_name=eff_sheet, mode=mode, alpha=alpha)
        merged = pd.merge(pd.merge(ot, lv, on="부서", how="outer"), eff, on="부서", how="outer")
    else:
        merged = pd.merge(ot, lv, on="부서", how="outer")
        merged["직원효율화점수"] = 0.0
        merged["참고(PE)"] = np.nan
        merged["참고(SE)"] = np.nan
        merged["인원"] = np.nan
        merged["총 프로젝트 수"] = np.nan
        merged["실제 달성점수"] = np.nan

    merged = merged.fillna(0.0)

    # 가중 합산(–6~+6 유지)
    merged["최종점수(–6~+6)"] = (
        w_ot * merged["OT점수"] +
        w_lv * merged["연차점수"] +
        w_eff * merged["직원효율화점수"]
    )

    # 보기 좋게 정렬
    ordered = merged[[
        "부서", "OT점수", "연차점수", "직원효율화점수", "최종점수(–6~+6)",
        "참고(PE)", "참고(SE)", "인원", "총 프로젝트 수", "실제 달성점수"
    ]].sort_values(
        by=["최종점수(–6~+6)", "OT점수", "연차점수", "직원효율화점수"],
        ascending=[False, False, False, False]
    ).reset_index(drop=True)

    return ordered

# -----------------------
# 스크립트 직접 실행 (선택)
# -----------------------
if __name__ == "__main__":
    print("현재 작업 디렉터리:", os.getcwd())

    # 예시 파일 경로(환경에 맞게 수정)
    ot_file = r"시간외근무_현황_전체 (6월~12월).xlsx"
    lv_file = r"2025년_연차설정+정보_1423.xlsx"
    eff_file = r"실센터장_리더십_평가_2025년_2025-10-15.xlsx"  # 직원효율화 원본

    dept = None
    leave_sheet = None
    eff_sheet = None

    try:
        result = compute_total_score(
            overtime_file=ot_file,
            leave_file=lv_file,
            efficiency_file=eff_file,
            dept_filter=dept,
            leave_sheet=leave_sheet,
            eff_sheet=eff_sheet,
            mode="CE",          # "PE" | "SE" | "CE"
            alpha=0.5,          # CE 가중(프로젝트 대비 성과 비중)
            weights=(0.4,0.3,0.3)  # w_ot, w_lv, w_eff (합=1)
        )
        print("\n=== 최종 리더십 점수 (–6~+6, 가중 통합) ===")
        print(result.to_string(index=False))

        out_path = "leadership_total_results.csv"
        result.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"\n저장 완료: {os.path.abspath(out_path)}")
    except Exception as e:
        print(f"\n오류 발생: {e}")
