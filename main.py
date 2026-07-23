import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from io import BytesIO

# ===============================
# 페이지 설정
# ===============================

st.set_page_config(
    page_title="근태 및 수당 관리",
    page_icon="🕒",
    layout="wide"
)

# ===============================
# 세션 초기화
# ===============================

if "attendance" not in st.session_state:
    st.session_state.attendance = None

if "employee_rate" not in st.session_state:
    st.session_state.employee_rate = {
        "홍길동":15000,
        "김철수":15000,
        "이영희":15000,
        "박민수":15000,
        "최수진":15000
    }

# ===============================
# 제목
# ===============================

st.title("🕒 근태 및 수당 관리")

st.caption("월별 근태 · 출장 · 시간외근무 · 수당 자동 계산")

# ===============================
# Sidebar
# ===============================

st.sidebar.header("설정")

uploaded_file = st.sidebar.file_uploader(
    "월별 엑셀 업로드",
    type=["xlsx"]
)

employee = st.sidebar.selectbox(
    "직원 선택",
    [
        "전체",
        "홍길동",
        "김철수",
        "이영희",
        "박민수",
        "최수진"
    ]
)

st.sidebar.divider()

st.sidebar.subheader("시간외 단가")

for name in st.session_state.employee_rate.keys():

    st.session_state.employee_rate[name] = st.sidebar.number_input(
        name,
        value=st.session_state.employee_rate[name],
        step=100,
        key=f"rate_{name}"
    )

# ===============================
# 데이터 읽기
# ===============================

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.session_state.attendance = df

# ===============================
# 화면 구성
# ===============================

tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs([
    "월별 요약",
    "근태 확인",
    "출장",
    "시간외",
    "지급명세서",
    "통계"
])
