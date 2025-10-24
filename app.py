# -*- coding: utf-8 -*-
import io
import numpy as np
import pandas as pd
import streamlit as st
from leadership_total import (
    compute_total_score,
    DEFAULT_GRADE_WEIGHTS,
)

st.set_page_config(page_title="리더십 점수(OT+연차+효율화)", layout="wide")

st.title("리더십 점수 통합 대시보드")
st.caption("OT + 연차 + 효율화(프로젝트 관리 지수) 3축으로 산출합니다. 효율화 점수 범위: -3 ~ +3.")

with st.sidebar:
    st.header("① 파일 업로드")
    ot_file = st.file_uploader("OT(초과근무) 파일 (xlsx/xls/csv)", type=["xlsx","xls","csv"], key="ot")
    lv_file = st.file_uploader("연차 사용 파일 (xlsx/xls/csv)", type=["xlsx","xls","csv"], key="lv")
    prj_file = st.file_uploader("프로젝트 실적 파일 (xlsx/xls/csv)", type=["xlsx","xls","csv"], key="prj")

    st.header("② 컬럼 매핑(선택)")
    st.caption("파일 컬럼명이 다를 때 지정하세요. 지정 안 하면 기본값을 추정 시도합니다.")
    colmap = {}
    # OT
    with st.expander("OT 파일 컬럼 매핑"):
        colmap["ot_dept_col"] = st.text_input("부서 컬럼명", value="부서")
        colmap["ot_name_col"] = st.text_input("이름 컬럼명", value="이름")
        colmap["ot_hours_col"] = st.text_input("OT(인정) 시간 컬럼명", value="OT시간")
    # Leave
    with st.expander("연차 파일 컬럼 매핑"):
        colmap["lv_dept_col"] = st.text_input("부서 컬럼명 ", value="부서")
        colmap["lv_name_col"] = st.text_input("이름 컬럼명 ", value="이름")
        colmap["lv_days_col"] = st.text_input("연차 사용일수 컬럼명", value="연차사용일수")
        colmap["lv_base_days"] = st.number_input("연차 부여 기본일수(없으면 15)", min_value=1, max_value=30, value=15)
    # Project
    with st.expander("프로젝트 실적 파일 컬럼 매핑"):
        colmap["prj_dept_col"] = st.text_input("부서 컬럼명  ", value="부서")
        colmap["prj_name_col"] = st.text_input("이름 컬럼명  ", value="이름")
        colmap["prj_target_col"] = st.text_input("신청등급 컬럼명", value="신청등급")
        colmap["prj_final_col"]  = st.text_input("확정등급 컬럼명", value="확정등급")
        colmap["prj_score_col"]  = st.text_input("실적평가점수(0~100) 컬럼명", value="avg_score")
        colmap["prj_status_col"] = st.text_input("상태 컬럼명(선택, 승인/완료만 집계)", value="status")
        colmap["prj_level_col"]  = st.text_input("직급/레벨 컬럼명(선택, Level 1만 집계)", value="level")
        st.write("등급 가중치(참고):", DEFAULT_GRADE_WEIGHTS)

    st.header("③ 필터/옵션")
    dept_filter = st.text_input("부서 필터(부분일치)", value="")

    run = st.button("📊 점수 계산", type="primary",
                    disabled=not (ot_file and lv_file and prj_file))

def _show_help():
    with st.expander("계산 방식 요약", expanded=False):
        st.markdown(
            """
            - **효율화 지수**: 신청등급 가중치 합(기준만점) 대비, 확정등급 가중치×(실적점수/100)의 합(실제달성)의 **달성률(%)**로 산출.  
            - **달성률→효율화 점수(–3~+3)**  
              ≥110%:+3 / 100~109:+2.5 / 90~99:+2 / 80~89:+1 / 70~79:+0.5 / 60~69:0 / 50~59:-0.5 / 40~49:-1.5 / <40:-3
            - **최종지수** = OT점수 + 연차점수 + 효율화점수
            """
        )

_show_help()

def _dfu(d):
    st.dataframe(d, use_container_width=True, hide_index=True)

if run:
    try:
        result, parts = compute_total_score(
            overtime_file=ot_file,
            leave_file=lv_file,
            project_file=prj_file,
            dept_filter=dept_filter or None,
            column_map=colmap,
        )
        st.success("계산 완료!")

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.metric("평균 OT점수", f"{result['OT점수'].mean():.2f}")
        with c2: st.metric("평균 연차점수", f"{result['연차점수'].mean():.2f}")
        with c3: st.metric("평균 효율화점수", f"{result['효율화점수'].mean():.2f}")
        with c4: st.metric("평균 달성률(%)", f"{result['달성률(%)'].mean():.1f}")

        st.subheader("부서별 결과")
        st.dataframe(result, use_container_width=True, hide_index=True)

        st.subheader("시각화")
        st.bar_chart(result.set_index("부서")["효율화점수"])
        st.bar_chart(result.set_index("부서")["달성률(%)"])

        csv = result.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("CSV 다운로드", data=csv, file_name="leadership_scores.csv", mime="text/csv")

        with st.expander("세부 집계 (OT/연차/프로젝트)", expanded=False):
            st.markdown("**OT 집계**")
            _dfu(parts.get("ot", pd.DataFrame()))
            st.markdown("**연차 집계**")
            _dfu(parts.get("leave", pd.DataFrame()))
            st.markdown("**프로젝트 집계(효율화)**")
            _dfu(parts.get("eff", pd.DataFrame()))

    except Exception as e:
        st.exception(e)
