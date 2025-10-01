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
    3. 필요시 부서 필터 입력
    4. **점수 계산** 버튼 클릭
    
    ### 📈 점수 기준
    - **OT 점수**: -3 ~ +3점 (초과근무 잔여율 기반)
    - **연차 점수**: -3 ~ +3점 (연차 잔여율 기반)
    - **최종 점수**: -6 ~ +6점 (OT + 연차)
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
                dept_filter=dept_filter if dept_filter else None,
                leave_sheet=leave_sheet if leave_sheet else None
            )
        
        st.success("✅ 계산 완료!")
        
        # 결과 표시
        st.subheader("📊 최종 결과")
        
        # 통계 요약
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("전체 부서 수", len(result))
        with col2:
            avg_ot = result["OT점수"].mean()
            st.metric("평균 OT점수", f"{avg_ot:.2f}")
        with col3:
            avg_leave = result["연차점수"].mean()
            st.metric("평균 연차점수", f"{avg_leave:.2f}")
        with col4:
            avg_total = result["최종점수(–6~+6)"].mean()
            st.metric("평균 최종점수", f"{avg_total:.2f}")
        
        st.markdown("---")
        
        # 결과 테이블
        st.dataframe(
            result,
            use_container_width=True,
            height=400,
            column_config={
                "부서": st.column_config.TextColumn("부서", width="medium"),
                "OT점수": st.column_config.NumberColumn(
                    "OT점수",
                    format="%.1f",
                    help="초과근무 점수 (-3 ~ +3)"
                ),
                "연차점수": st.column_config.NumberColumn(
                    "연차점수",
                    format="%.1f",
                    help="연차 점수 (-3 ~ +3)"
                ),
                "최종점수(–6~+6)": st.column_config.NumberColumn(
                    "최종점수",
                    format="%.1f",
                    help="OT점수 + 연차점수"
                ),
            }
        )
        
        # 시각화
        st.subheader("📈 점수 분포")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # OT 점수 막대 그래프
            st.bar_chart(
                result.set_index("부서")["OT점수"],
                use_container_width=True
            )
            st.caption("OT 점수 분포")
        
        with col2:
            # 연차 점수 막대 그래프
            st.bar_chart(
                result.set_index("부서")["연차점수"],
                use_container_width=True
            )
            st.caption("연차 점수 분포")
        
        # 최종 점수 막대 그래프
        st.bar_chart(
            result.set_index("부서")["최종점수(–6~+6)"],
            use_container_width=True
        )
        st.caption("최종 점수 분포")
        
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
- 부서 필터: {dept_filter if dept_filter else 'None'}
- 연차 시트: {leave_sheet if leave_sheet else 'None'}
            """)

# 푸터
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; padding: 20px;'>
        <p>리더십 점수 자동 산출 시스템 v1.0</p>
        <p>문의사항이 있으시면 관리자에게 연락해주세요.</p>
    </div>
    """,
    unsafe_allow_html=True
)
