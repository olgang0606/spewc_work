import streamlit as st
import pandas as pd

st.title("근태 관리 테스트")

st.write("앱이 정상 로드되었습니다!")

uploaded_file = st.file_uploader("Excel 파일 업로드", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, header=2)
    st.write("데이터 로드 성공!")
    st.dataframe(df.head())
else:
    st.write("파일을 업로드 해주세요.")
