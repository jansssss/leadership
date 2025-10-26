import streamlit as st
import pandas as pd
from leadership_total import compute_total_score

# 페이지 설정
st.set_page_config(
    page_title="리더십 점수 자동 산출기",
    page_icon="📊",
    layout="wide"
)

# 타이틀
st.title("📊 리더십 점수 자동 산출기")
st.markdown("---")

# 설명
with st.expander("ℹ️ 사용 방법", expanded=False):
    st.markdown("""
    ### 📋 사용 방법
    1. **초과근무 파일** 업로드 (엑셀 형식, '총계'와 '일별현황_A' 시트 필요)
    2. **연차 파일** 업로드 (엑셀 형식)
    3. **업무적절성 파일** 업로드 (엑셀 형식, 선택사항)
    4. 필요시 부서 필터 입력
    5. **점수 계산** 버튼 클릭

    ### 📈 점수 기준 (개정)
    - **OT 점수**: **–1 ~ +2점** (Residual 기반)
    - **연차 점수**: **–1 ~ +1점** (잔여율 기반)
    - **업무적절성 점수**: **–1 ~ +2점** (달성률 기반)
    - **최종 점수**: **–3 ~ +5점** (OT + 연차 + 업무적절성)
    """)

# 사이드바 설정
st.sidebar.header("⚙️ 설정")

# 파일 업로드
st.sidebar.subheader("1️⃣ 파일 업로드")
ot_file = st.sidebar.file_uploader(
    "초과근무 파일",
    type=['xlsx', 'xls'],
    help="'총계'와 '일별현황_A' 시트가 포함된 엑셀 파일"
)
lv_file = st.sidebar.file_uploader(
    "연차 파일",
    type=['xlsx', 'xls'],
    help="부서/이름/부여/사용/잔여 정보가 포함된 엑셀 파일"
)
ap_file = st.sidebar.file_uploader(
    "업무적절성 파일 (선택사항)",
    type=['xlsx', 'xls'],
    help="실/센터, 달성률(%) 정보가 포함된 엑셀 파일"
)

# 옵션 설정
st.sidebar.subheader("2️⃣ 옵션 설정")
dept_filter = st.sidebar.text_input(
    "부서 필터 (선택사항)", 
    value="",
    placeholder="예: 전략기획",
    help="특정 부서만 보려면 입력하세요"
)

leave_sheet = st.sidebar.text_input(
    "연차 시트명 (선택사항)",
    value="",
    placeholder="비워두면 첫 시트 사용",
    help="특정 시트를 지정하려면 입력하세요"
)

ap_sheet = st.sidebar.text_input(
    "업무적절성 시트명 (선택사항)",
    value="",
    placeholder="비워두면 첫 시트 사용",
    help="특정 시트를 지정하려면 입력하세요"
)

# 메인 영역
col1, col2 = st.columns([3, 1])

with col1:
    st.subheader("📤 업로드된 파일")
    if ot_file:
        st.success(f"✅ 초과근무 파일: {ot_file.name}")
    else:
        st.info("⏳ 초과근무 파일을 업로드해주세요")

    if lv_file:
        st.success(f"✅ 연차 파일: {lv_file.name}")
    else:
        st.info("⏳ 연차 파일을 업로드해주세요")

    if ap_file:
        st.success(f"✅ 업무적절성 파일: {ap_file.name}")
    else:
        st.warning("⚠️ 업무적절성 파일 미업로드 (0점 처리됨)")

with col2:
    st.subheader("🎯 실행")
    calculate_btn = st.button(
        "📊 점수 계산", 
        type="primary",
        disabled=(ot_file is None or lv_file is None),
        use_container_width=True
    )

st.markdown("---")

# 계산 실행
if calculate_btn and ot_file is not None and lv_file is not None:
    try:
        with st.spinner("⏳ 계산 중... 잠시만 기다려주세요"):
            # 점수 계산
            result = compute_total_score(
                ot_file,
                lv_file,
                appropriateness_file=ap_file if ap_file else None,
                dept_filter=dept_filter if dept_filter else None,
                leave_sheet=leave_sheet if leave_sheet else None,
                appropriateness_sheet=ap_sheet if ap_sheet else None
            )

        # === 컬럼 표준화: 최종점수 컬럼명을 '최종점수'로 통일 ===
        final_candidates = ["최종점수(–3~+5)", "최종점수(–2~+3)", "최종점수(–6~+6)", "최종점수"]
        final_col = None
        for c in final_candidates:
            if c in result.columns:
                final_col = c
                break
        if final_col is None:
            raise KeyError("결과에서 '최종점수' 컬럼을 찾을 수 없습니다.")

        if final_col != "최종점수":
            result = result.rename(columns={final_col: "최종점수"})

        st.success("✅ 계산 완료!")
        
        # 결과 표시
        st.subheader("📊 최종 결과")
        
        # 통계 요약
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("전체 부서 수", len(result))
        with col2:
            avg_ot = result["OT점수"].mean()
            st.metric("평균 OT점수", f"{avg_ot:.2f}")
        with col3:
            avg_leave = result["연차점수"].mean()
            st.metric("평균 연차점수", f"{avg_leave:.2f}")
        with col4:
            if "업무적절성점수" in result.columns:
                avg_ap = result["업무적절성점수"].mean()
                st.metric("평균 업무적절성", f"{avg_ap:.2f}")
            else:
                st.metric("평균 업무적절성", "N/A")
        with col5:
            avg_total = result["최종점수"].mean()
            st.metric("평균 최종점수", f"{avg_total:.2f}")
        
        st.markdown("---")
        
        # 결과 테이블
        column_config = {
            "부서": st.column_config.TextColumn("부서", width="medium"),
            "OT점수": st.column_config.NumberColumn(
                "OT점수",
                format="%.2f",
                help="초과근무 점수 (–1 ~ +2)"
            ),
            "연차점수": st.column_config.NumberColumn(
                "연차점수",
                format="%.2f",
                help="연차 점수 (–1 ~ +1)"
            ),
            "최종점수": st.column_config.NumberColumn(
                "최종점수",
                format="%.2f",
                help="OT + 연차 + 업무적절성 = (–3 ~ +5)"
            ),
        }

        # 업무적절성 컬럼이 있으면 추가
        if "업무적절성점수" in result.columns:
            column_config["업무적절성점수"] = st.column_config.NumberColumn(
                "업무적절성점수",
                format="%.2f",
                help="업무적절성 점수 (–1 ~ +2)"
            )

        st.dataframe(
            result,
            use_container_width=True,
            height=400,
            column_config=column_config
        )
        
        # 시각화
        st.subheader("📈 점수 분포")

        # 업무적절성 파일 업로드 여부에 따라 레이아웃 조정
        if "업무적절성점수" in result.columns and ap_file:
            col1, col2, col3 = st.columns(3)

            with col1:
                st.bar_chart(
                    result.set_index("부서")["OT점수"],
                    use_container_width=True
                )
                st.caption("OT 점수 분포 (–1 ~ +2)")

            with col2:
                st.bar_chart(
                    result.set_index("부서")["연차점수"],
                    use_container_width=True
                )
                st.caption("연차 점수 분포 (–1 ~ +1)")

            with col3:
                st.bar_chart(
                    result.set_index("부서")["업무적절성점수"],
                    use_container_width=True
                )
                st.caption("업무적절성 점수 분포 (–1 ~ +2)")
        else:
            col1, col2 = st.columns(2)

            with col1:
                st.bar_chart(
                    result.set_index("부서")["OT점수"],
                    use_container_width=True
                )
                st.caption("OT 점수 분포 (–1 ~ +2)")

            with col2:
                st.bar_chart(
                    result.set_index("부서")["연차점수"],
                    use_container_width=True
                )
                st.caption("연차 점수 분포 (–1 ~ +1)")

        st.bar_chart(
            result.set_index("부서")["최종점수"],
            use_container_width=True
        )
        st.caption("최종 점수 분포 (–3 ~ +5)")
        
        # CSV 다운로드
        st.markdown("---")
        st.subheader("💾 결과 다운로드")
        
        csv = result.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label="📥 CSV 파일 다운로드",
            data=csv,
            file_name="leadership_scores.csv",
            mime="text/csv",
            use_container_width=True,
            type="primary"
        )
    
    except Exception as e:
        st.error(f"❌ 오류 발생: {str(e)}")
        
        # 디버깅 정보 (선택적으로 표시)
        with st.expander("🔍 상세 오류 정보 (개발자용)", expanded=False):
            st.exception(e)
            st.code(f"""
파일 정보:
- OT 파일: {ot_file.name if ot_file else 'None'}
- 연차 파일: {lv_file.name if lv_file else 'None'}
- 업무적절성 파일: {ap_file.name if ap_file else 'None'}
- 부서 필터: {dept_filter if dept_filter else 'None'}
- 연차 시트: {leave_sheet if leave_sheet else 'None'}
- 업무적절성 시트: {ap_sheet if ap_sheet else 'None'}
현재 컬럼: {list(result.columns) if 'result' in locals() else 'N/A'}
            """)

# 푸터
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; padding: 20px;'>
        <p>리더십 점수 자동 산출 시스템 v2.0 (업무적절성 추가)</p>
        <p>문의사항이 있으시면 관리자에게 연락해주세요.</p>
    </div>
    """,
    unsafe_allow_html=True
)
