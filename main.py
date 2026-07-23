import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np

st.set_page_config(page_title="근태관리", layout="wide")
st.title("🕒 근로자 근태 및 수당 관리")

# 파일 업로드
uploaded_file = st.file_uploader("근무현황.xlsx 업로드", type=["xlsx"])

if uploaded_file is None:
    st.warning("파일을 업로드해주세요.")
    try:
        df = pd.read_excel("attachments/근무현황.xlsx", sheet_name="9월", header=2)
        st.info("기본 9월 데이터 사용")
    except:
        st.error("기본 파일도 없습니다. Excel 파일 업로드 필수")
        st.stop()
else:
    df = pd.read_excel(uploaded_file, header=2)

# 데이터 정리
df['인증일시'] = pd.to_datetime(df['인증일시'])
df['date'] = df['인증일시'].dt.date
df = df.sort_values(['이름', '인증일시'])

# 처리 함수
def process_data(df):
    data = []
    for (name, d), g in df.groupby(['이름', 'date']):
        dt = pd.to_datetime(d)
        clock_in = g[g['인증모드']=='출근']['인증일시'].min()
        clock_out = g[g['인증모드']=='퇴근']['인증일시'].max()
        
        total_h = 0
        if clock_in and clock_out:
            total_h = (clock_out - clock_in).total_seconds() / 3600
        
        is_weekend = dt.weekday() >= 5
        basic = 9.0 if not is_weekend else 0
        ot = max(0, total_h - basic) if not is_weekend else min(total_h, 8.0)
        ot = np.floor(ot * 2) / 2
        
        data.append({
            '이름': name,
            '날짜': str(d),
            '총근무': round(total_h, 2),
            '시간외': round(ot, 2),
            '비고': '주간초과주의' if ot > 8 else ''
        })
    return pd.DataFrame(data)

result_df = process_data(df)

st.header("근태 결과")
st.dataframe(result_df, use_container_width=True)

if st.button("수당 계산"):
    summary = result_df.groupby('이름').sum(numeric_only=True)
    st.dataframe(summary)
    st.success("✅ 계산 완료 (단가 적용은 수동)")

st.caption("간소화 버전입니다. 더 많은 기능 원하시면 로컬에서 실행 추천")
