import streamlit as st
import pandas as pd

st.title("🕒 근태 관리")

uploaded_file = st.file_uploader("근무현황.xlsx 업로드", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=2)
    st.success("파일 로드 성공!")
    st.dataframe(df.head(10))
else:
    st.info("Excel 파일을 업로드해주세요.")
