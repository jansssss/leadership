# -*- coding: utf-8 -*-
"""
리더십 평가 통합 점수 산출기 (전체 파일)
- OT 점수: –1 ~ +2
- 연차 점수: –1 ~ +1
- 업무적절성 점수: –1 ~ +2
- 최종 점수: –3 ~ +5 (= OT + 연차 + 업무적절성)

설계 포인트
- 초과근무: '총계' + '일별현황_A' 시트 사용 (부서/직급/이름/환산/신청일 등 자동 감지)
- AH(가용근로시간) = Σ(직급별 인원 × 직급계수 × 기준근로시간 × 월수)
- EOR(예상 초과근무율) = 전사평균OT율 × (전사평균AH ÷ 팀AH)
- AOR(실제 초과근무율) = 팀 OT ÷ (팀 AH + 팀 OT)
- Residual(%) = AOR – EOR
- 연차: 첫 시트(또는 지정 시트)에서 부서/이름/부여/사용/잔여 자동 감지
- 업무적절성: 달성률(%) 기반으로 점수 산정
"""

import os
import re
import pandas as pd
import numpy as np
from datetime import datetime

# -----------------------
# 공통 설정
# -----------------------
HOURS_PER_DAY = 8.0                # 1일 근무시간
BASE_MONTHLY_HOURS = 160           # 월 기준 근로시간
RANK_WEIGHTS = {                   # 직급 가중치(미정의는 1.0)
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

def parse_hhmm_to_minutes(x, numeric_is_hours: bool = True):
    """
    '64:00' → 분(int)
    숫자형은 기본적으로 '시간'으로 간주하여 분으로 변환(시간×60).
    numeric_is_hours=False 로 주면 숫자를 '분'으로 간주.
    """
    if pd.isna(x):
        return 0
    # 문자열 HH:MM
    if isinstance(x, str) and ":" in x:
        try:
            h, m = x.strip().split(":")
            return int(float(h)) * 60 + int(float(m))
        except Exception:
            return 0
    # 숫자형
    if isinstance(x, (int, float, np.integer, np.floating)):
        return int(round(float(x) * 60)) if numeric_is_hours else int(round(float(x)))
    return 0

def detect_months_from_dates(series):
    """날짜열에서 (연-월) 유니크 개수 → 1~12 범위 월 수 산정"""
    months = set()
    for v in series.dropna():
        try:
            dt = pd.to_datetime(v)
            months.add((dt.year, dt.month))
        except Exception:
            pass
    cnt = len(months) if months else 1
    return min(max(cnt, 1), 12)

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
    # 최선: 텍스트 비율이 높은 열
    best_col, best_score = None, -1.0
    for c in df.columns:
        s = df[c].astype(str)
        score = s.apply(lambda x: x.strip() != "" and not x.strip().isdigit()).mean()
        if score > best_score:
            best_score, best_col = score, c
    return best_col

def _pick_time_col(df):
    for c in df.columns:
        if "환산" in str(c) or "시간" in str(c):
            return c
    # HH:MM 패턴이 많은 열 선택
    best_col, best_rate = None, -1.0
    for c in df.columns:
        s = df[c].astype(str)
        rate = s.str.contains(r"^\s*\d{1,3}:\d{2}\s*$").mean()
        if rate > best_rate:
            best_rate, best_col = rate, c
    return best_col

# -----------------------
# 초과근무 → 점수(–1~+2)
# -----------------------
def residual_to_score_ot(residual_pct: float) -> float:
    """
    Residual(%) → OT 점수(–1~+2)
    - Residual = AOR - EOR
    - 음수(예상보다 적은 OT)는 가점, 양수는 감점
    """
    r = float(residual_pct)

    # 가점(최대 +2)
    if r <= -10:
        return 2.0
    elif r <= -6:
        return 1.5
    elif r <= -3:
        return 1.0
    elif r <= -1:
        return 0.5

    # 중립
    if -1 < r < 1:
        return 0.0

    # 감점(최대 -1)
    if r < 3:
        return -0.25
    elif r < 6:
        return -0.5
    elif r < 10:
        return -0.75
    else:
        return -1.0

def compute_overtime_score(excel_path_or_buffer, dept_filter: str | None = None):
    """
    '총계' + '일별현황_A' 시트 사용
    출력: 부서별 OT점수 DataFrame[부서, OT점수]
    """
    # Streamlit UploadedFile 객체 처리
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
        raise ValueError("총계 시트에서 초과근무 합계(환산/시간) 열을 찾지 못했습니다.")

    tmp = df_sum.copy()
    # 숫자형은 '시간'으로 간주하여 분으로 변환
    tmp["__mins__"] = tmp[sum_ot_col].apply(lambda v: parse_hhmm_to_minutes(v, numeric_is_hours=True))
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
    # 숫자형은 '시간'으로 간주하여 분으로 변환
    df_daily["_분"] = df_daily[ot_col_daily].apply(lambda v: parse_hhmm_to_minutes(v, numeric_is_hours=True))

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
    # 두 출처 중 큰 값을 채택
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
    dept_table["OT점수"] = dept_table["Residual(%)"].apply(residual_to_score_ot)

    ot_out = dept_table[["부서", "OT점수"]]
    if dept_filter:
        ot_out = ot_out[ot_out["부서"].str.contains(dept_filter, na=False)]
    return ot_out

# -----------------------
# 업무적절성 → 점수(–1~+2)
# -----------------------
def appropriateness_to_score(achievement_pct: float) -> float:
    """
    달성률(%) → 업무적절성 점수(–1~+2)
    - 달성률이 높을수록 가점
    """
    r = float(achievement_pct)

    # 가점(최대 +2)
    if r >= 100:
        return 2.0
    elif r >= 90:
        return 1.5
    elif r >= 80:
        return 1.0  # 우수
    elif r >= 70:
        return 0.5

    # 중립
    if 60 <= r < 70:
        return 0.0  # 보통

    # 감점(최대 -1)
    if r >= 50:
        return -0.5
    else:
        return -1.0  # 미흡

def compute_appropriateness_score(excel_path_or_buffer, dept_filter: str | None = None, sheet_name: str | None = None):
    """
    업무적절성 파일 → 부서별 업무적절성점수 DataFrame[부서, 업무적절성점수]
    """
    # Streamlit UploadedFile 객체 처리
    if hasattr(excel_path_or_buffer, 'seek'):
        excel_path_or_buffer.seek(0)

    # 시트 선택
    xls = pd.ExcelFile(excel_path_or_buffer)
    use_sheet = sheet_name if (sheet_name and sheet_name in xls.sheet_names) else xls.sheet_names[0]

    df = pd.read_excel(excel_path_or_buffer, sheet_name=use_sheet)
    df.columns = [str(c).strip() for c in df.columns]

    # 컬럼 자동 감지
    dept_col = None
    for cand in ["실/센터", "부서", "소속", "센터", "실센터"]:
        if cand in df.columns:
            dept_col = cand
            break

    achievement_col = None
    for cand in ["달성률(%)", "달성률", "달성률 (%)", "달성율(%)"]:
        if cand in df.columns:
            achievement_col = cand
            break

    if not dept_col:
        raise ValueError("업무적절성 시트에서 부서 컬럼(실/센터, 부서 등)을 찾지 못했습니다.")
    if not achievement_col:
        raise ValueError("업무적절성 시트에서 달성률 컬럼을 찾지 못했습니다.")

    # 데이터 정제
    df["_부서"] = df[dept_col].apply(_norm)
    df["_달성률"] = pd.to_numeric(df[achievement_col], errors="coerce").fillna(0.0)

    # 부서별 점수 계산
    result = df[["_부서", "_달성률"]].copy()
    result = result[result["_부서"] != ""]  # 빈 부서명 제외
    result["업무적절성점수"] = result["_달성률"].apply(appropriateness_to_score)
    result = result.rename(columns={"_부서": "부서"})

    if dept_filter:
        result = result[result["부서"].str.contains(dept_filter, na=False)]

    return result[["부서", "업무적절성점수"]]

# -----------------------
# 연차 → 점수(–1~+1)
# -----------------------
def _looks_hhmm(val: str) -> bool:
    return bool(re.match(r"^\s*\d{1,3}:\d{2}\s*$", str(val or "")))

def _to_days(x, hours_per_day=HOURS_PER_DAY):
    """숫자=일수, 'HH:MM'은 시간/8로 일수 환산, 그 외는 숫자 추출"""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return 0.0
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip()
    if _looks_hhmm(s):
        h, m = s.split(":")
        hours = float(h) + float(m) / 60.0
        return hours / hours_per_day if hours_per_day > 0 else 0.0
    m = re.search(r"(\d+(\.\d+)?)", s)
    return float(m.group(1)) if m else 0.0

def compute_leave_score(excel_path_or_buffer, dept_filter: str | None = None, sheet_name: str | None = None):
    """
    연차 사용현황 파일 → 부서별 연차점수 DataFrame[부서, 연차점수]
    - 헤더가 병합되어 'Unnamed: n'이 많은 경우를 대비해: 헤더 복구 + 상단 N행에서 '부여/사용/잔여' 마커 스캔
    """
    # Streamlit UploadedFile 객체 처리
    if hasattr(excel_path_or_buffer, 'seek'):
        excel_path_or_buffer.seek(0)
    
    # 1) 시트 선택
    xls = pd.ExcelFile(excel_path_or_buffer)
    use_sheet = sheet_name if (sheet_name and sheet_name in xls.sheet_names) else xls.sheet_names[0]

    # 2) 1차 로드(일반 헤더)
    df = pd.read_excel(excel_path_or_buffer, sheet_name=use_sheet)
    df.columns = [str(c).strip() for c in df.columns]

    # -----------------------------
    # 헤더 복구 유틸
    # -----------------------------
    def _rebuild_header_by_rowindex(idx: int):
        df2 = pd.read_excel(excel_path_or_buffer, sheet_name=use_sheet, header=idx)
        if isinstance(df2.columns, pd.MultiIndex):
            df2.columns = [" ".join([_norm(x) for x in tup if _norm(x)]).strip() for tup in df2.columns.to_list()]
        df2.columns = [c.replace("이름(ID)", "이름").strip() for c in df2.columns]
        return df2

    def _looks_broken_header(df0: pd.DataFrame) -> bool:
        unnamed_ratio = np.mean([str(c).startswith("Unnamed:") for c in df0.columns])
        has_keywords = any(k in " ".join(df0.columns) for k in ["부여", "사용", "잔여"])
        return (unnamed_ratio > 0.3) and (not has_keywords)

    # 3) 헤더가 깨진 경우 상단 6행까지 스캔하여 가장 적절한 행을 헤더로 승격
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

    # 4) 열 이름 정리
    df.columns = [c.replace("이름(ID)", "이름").strip() for c in df.columns]

    # 5) 컬럼 자동 감지 (연차용)
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

    missing = []
    if not dept_col: missing.append("부서")
    if not name_col: missing.append("이름")
    have_cnt = sum([grant_col is not None, used_col is not None, remain_col is not None])
    if have_cnt < 2:
        raise ValueError(
            "연차 시트에서 다음 정보를 찾지 못했습니다: 부여/사용/잔여 중 최소 2개\n"
            f"열 목록: {list(df.columns)}"
        )

    # 6) 수치화 및 파생 보정
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

    # 잔여율 → 점수(–1~+1)
    def leave_score(remain_pct: float) -> float:
        if remain_pct <= 2:   return 1.0
        elif remain_pct <= 5: return 0.75
        elif remain_pct <= 10:return 0.5
        elif remain_pct <= 20:return 0.25
        elif remain_pct <= 30:return 0.0
        elif remain_pct <= 40:return -0.5
        else:                 return -1.0

    grp["연차점수"] = grp["잔여율(%)"].apply(leave_score)

    if dept_filter:
        grp = grp[grp["부서"].str.contains(dept_filter, na=False)]
    return grp[["부서", "연차점수"]]

# -----------------------
# 통합 최종 계산
# -----------------------
def compute_total_score(overtime_file, leave_file, appropriateness_file=None, dept_filter: str | None = None, leave_sheet: str | None = None, appropriateness_sheet: str | None = None):
    """
    OT + 연차 + 업무적절성 통합 점수 계산
    - OT: –1 ~ +2
    - 연차: –1 ~ +1
    - 업무적절성: –1 ~ +2
    - 최종: –3 ~ +5
    """
    ot = compute_overtime_score(overtime_file, dept_filter)
    lv = compute_leave_score(leave_file, dept_filter, sheet_name=leave_sheet)

    # OT + 연차 병합
    merged = pd.merge(ot, lv, on="부서", how="outer").fillna(0.0)

    # 업무적절성 추가 (파일이 제공된 경우)
    if appropriateness_file is not None:
        ap = compute_appropriateness_score(appropriateness_file, dept_filter, sheet_name=appropriateness_sheet)
        merged = pd.merge(merged, ap, on="부서", how="outer").fillna(0.0)
    else:
        # 업무적절성 파일이 없으면 0점 처리
        merged["업무적절성점수"] = 0.0

    # 최종점수 계산 (–3~+5)
    merged["최종점수(–3~+5)"] = merged["OT점수"] + merged["연차점수"] + merged["업무적절성점수"]

    # 호환성(예전 프런트 사용 시): 옛 컬럼명도 함께 제공
    merged["최종점수(–2~+3)"] = merged["OT점수"] + merged["연차점수"]  # 업무적절성 제외
    merged["최종점수(–6~+6)"] = merged["최종점수(–3~+5)"]  # 호환성

    # 보기 좋게
    merged = merged[["부서", "OT점수", "연차점수", "업무적절성점수", "최종점수(–3~+5)"]].sort_values(
        by=["최종점수(–3~+5)", "OT점수", "연차점수", "업무적절성점수"], ascending=[False, False, False, False]
    ).reset_index(drop=True)
    return merged

# -----------------------
# 스크립트 직접 실행 (선택사항)
# -----------------------
if __name__ == "__main__":
    print("현재 작업 디렉터리:", os.getcwd())

    # ❗ 여기를 실제 파일 경로로 바꾸세요
    ot_file = r"시간외근무_현황_전체 (6월~12월).xlsx"
    lv_file = r"2025년_연차설정+정보_1423.xlsx"

    dept = None            # 예: "전략기획"
    leave_sheet = None     # 특정 시트명 지정 시

    try:
        result = compute_total_score(ot_file, lv_file, dept_filter=dept, leave_sheet=leave_sheet)
        print("\n=== 최종 리더십 점수 (–2~+3) ===")
        print(result.to_string(index=False))

        out_path = "leadership_total_results.csv"
        result.to_csv(out_path, index=False, encoding="utf-8-sig")
        print(f"\n저장 완료: {os.path.abspath(out_path)}")
    except Exception as e:
        print(f"\n오류 발생: {e}")
