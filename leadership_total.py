# -*- coding: utf-8 -*-
import io
import re
import numpy as np
import pandas as pd
from typing import Dict, Optional, Tuple

# 등급 가중치
DEFAULT_GRADE_WEIGHTS: Dict[str, float] = {
    "S": 1.0, "A": 0.8, "B": 0.65, "C": 0.5, "D": 0.3
}

def _to_df(file):
    if hasattr(file, "seek"):
        file.seek(0)
    name = getattr(file, "name", "")
    if name.lower().endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def _safe_col(df: pd.DataFrame, candidates, default=None):
    for c in candidates:
        if c in df.columns:
            return c
    return default

def _norm_grade(x: str) -> str:
    s = "".join(ch for ch in str(x).upper() if ch.isalpha())
    return s[:1] if s else ""

# ---------- OT ----------
def compute_overtime_score(overtime_file,
                           dept_filter: Optional[str]=None,
                           column_map: Optional[Dict]=None) -> pd.DataFrame:
    df = _to_df(overtime_file)
    cm = column_map or {}
    dept = cm.get("ot_dept_col") or _safe_col(df, ["부서","부 서","Department","dept"])
    name = cm.get("ot_name_col") or _safe_col(df, ["이름","성명","Name"])
    hours = cm.get("ot_hours_col") or _safe_col(df, ["OT시간","초과근무시간","인정시간","ot_hours"])

    if not all([dept, name, hours]):
        raise ValueError("OT 파일에 필요한 컬럼(부서/이름/OT시간)을 찾을 수 없습니다. 사이드바에서 매핑해 주세요.")

    df[dept] = df[dept].astype(str).str.strip()
    df[name] = df[name].astype(str).str.strip()
    df[hours] = pd.to_numeric(df[hours], errors="coerce").fillna(0.0).clip(lower=0)

    g = df.groupby(dept).agg(OT총시간=(hours,"sum"), 인원수=(name,"nunique")).reset_index()
    g["1인당OT"] = np.where(g["인원수"]>0, g["OT총시간"]/g["인원수"], 0.0)

    perc = np.percentile(g["1인당OT"], [20, 40, 50, 60, 80]) if len(g) > 0 else [0,0,0,0,0]
    def map_ot(x):
        if x <= perc[0]: return 3.0
        if x <= perc[1]: return 2.0
        if x <= perc[2]: return 1.0
        if x <= perc[3]: return 0.5
        if x <= perc[4]: return 0.0
        return -1.0
    g["OT점수"] = g["1인당OT"].apply(map_ot)

    if dept_filter:
        g = g[g[dept].str.contains(dept_filter, na=False)]

    return g.rename(columns={dept:"부서"})[["부서","OT점수","1인당OT","OT총시간","인원수"]]

# ---------- 연차 ----------
def compute_leave_score(leave_file,
                        dept_filter: Optional[str]=None,
                        column_map: Optional[Dict]=None) -> pd.DataFrame:
    df = _to_df(leave_file)
    cm = column_map or {}
    dept = cm.get("lv_dept_col") or _safe_col(df, ["부서","Department","dept"])
    name = cm.get("lv_name_col") or _safe_col(df, ["이름","성명","Name"])
    days = cm.get("lv_days_col") or _safe_col(df, ["연차사용일수","사용연차","annual_leave_used"])
    base_days = int(cm.get("lv_base_days") or 15)

    if not all([dept, name, days]):
        raise ValueError("연차 파일에 필요한 컬럼(부서/이름/연차사용일수)을 찾을 수 없습니다. 사이드바에서 매핑해 주세요.")

    df[dept] = df[dept].astype(str).str.strip()
    df[name] = df[name].astype(str).str.strip()
    df[days] = pd.to_numeric(df[days], errors="coerce").fillna(0.0).clip(lower=0)

    df["사용률"] = (df[days] / max(base_days,1)).clip(0, 1.5)
    g = df.groupby(dept).agg(연차평균사용률=("사용률","mean"), 인원수=(name,"nunique")).reset_index()

    perc = np.percentile(g["연차평균사용률"], [20, 40, 50, 60, 80]) if len(g) > 0 else [0,0,0,0,0]
    def map_leave(x):
        if x >= perc[4]: return 3.0
        if x >= perc[3]: return 2.0
        if x >= perc[2]: return 1.0
        if x >= perc[1]: return 0.5
        if x >= perc[0]: return 0.0
        return -1.0
    g["연차점수"] = g["연차평균사용률"].apply(map_leave)

    if dept_filter:
        g = g[g[dept].str.contains(dept_filter, na=False)]

    return g.rename(columns={dept:"부서"})[["부서","연차점수","연차평균사용률","인원수"]]

# ---------- 효율화(프로젝트) ----------
def efficiency_score_from_rate(r: float) -> float:
    if r >= 110: return 3.0
    if r >= 100: return 2.5
    if r >= 90:  return 2.0
    if r >= 80:  return 1.0
    if r >= 70:  return 0.5
    if r >= 60:  return 0.0
    if r >= 50:  return -0.5
    if r >= 40:  return -1.5
    return -3.0

def compute_efficiency_score(project_file,
                             dept_filter: Optional[str]=None,
                             column_map: Optional[Dict]=None,
                             grade_weights: Optional[Dict[str,float]]=None
                             ) -> pd.DataFrame:
    W = (grade_weights or DEFAULT_GRADE_WEIGHTS).copy()
    df = _to_df(project_file)
    cm = column_map or {}

    dept  = cm.get("prj_dept_col")   or _safe_col(df, ["부서","Department","dept"])
    name  = cm.get("prj_name_col")   or _safe_col(df, ["이름","성명","Name"])
    tcol  = cm.get("prj_target_col") or _safe_col(df, ["신청등급","target_grade"])
    fcol  = cm.get("prj_final_col")  or _safe_col(df, ["확정등급","final_grade"])
    scol  = cm.get("prj_score_col")  or _safe_col(df, ["avg_score","실적평가점수","평균점수"])
    stat  = cm.get("prj_status_col") or _safe_col(df, ["status","상태"])
    lvl   = cm.get("prj_level_col")  or _safe_col(df, ["level","직급","레벨"])

    if not all([dept, name, tcol, fcol, scol]):
        raise ValueError("프로젝트 파일에 필요한 컬럼(부서/이름/신청등급/확정등급/실적점수)을 찾을 수 없습니다. 사이드바에서 매핑해 주세요.")

    df[dept] = df[dept].astype(str).str.strip()
    df[name] = df[name].astype(str).str.strip()
    df[tcol] = df[tcol].map(_norm_grade)
    df[fcol] = df[fcol].map(_norm_grade)
    df[scol] = pd.to_numeric(df[scol], errors="coerce").fillna(0.0).clip(0,100)

    if stat in df.columns:
        df = df[df[stat].astype(str).str.contains("승인|완료")]
    if lvl in df.columns:
        df = df[df[lvl].astype(str).str.strip() == "1"]

    df["신청W"] = df[tcol].map(W).fillna(0.0)
    df["확정W"] = df[fcol].map(W).fillna(0.0)
    df["실제달성"] = df["확정W"] * (df[scol]/100.0)

    g = df.groupby(dept, dropna=False).agg(
        신청만점=("신청W","sum"),
        실제달성=("실제달성","sum"),
        직원수=(name,"nunique")
    ).reset_index()

    g["달성률(%)"] = np.where(g["신청만점"]>0,
                         (g["실제달성"]/g["신청만점"])*100.0, 0.0)
    g["효율화점수"] = g["달성률(%)"].apply(efficiency_score_from_rate)

    if dept_filter:
        g = g[g[dept].str.contains(dept_filter, na=False)]

    return g.rename(columns={dept:"부서"})[["부서","효율화점수","달성률(%)","신청만점","실제달성","직원수"]]

# ---------- 총합 ----------
def compute_total_score(overtime_file, leave_file, project_file,
                        dept_filter: Optional[str]=None,
                        column_map: Optional[Dict]=None,
                        grade_weights: Optional[Dict[str,float]]=None
                        ) -> Tuple[pd.DataFrame, Dict[str,pd.DataFrame]]:
    ot = compute_overtime_score(overtime_file, dept_filter, column_map)
    lv = compute_leave_score(leave_file, dept_filter, column_map)
    ef = compute_efficiency_score(project_file, dept_filter, column_map, grade_weights)

    merged = ot.merge(lv, on="부서", how="outer").merge(ef, on="부서", how="outer").fillna(0.0)
    merged["최종지수"] = merged["OT점수"] + merged["연차점수"] + merged["효율화점수"]
    merged = merged[["부서","OT점수","연차점수","효율화점수","달성률(%)","최종지수"]].sort_values("최종지수", ascending=False)
    parts = {"ot": ot, "leave": lv, "eff": ef}
    return merged, parts
