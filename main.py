import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

st.set_page_config(page_title="근태 및 수당 관리", layout="wide")
st.title("🕒 근태 및 수당 관리 시스템")

uploaded_file = st.file_uploader("📁 근무현황.xlsx 업로드", type=["xlsx"])

if uploaded_file is None:
    st.warning("Excel 파일을 업로드해주세요.")
    st.stop()

df = pd.read_excel(uploaded_file, header=2)
df['인증일시'] = pd.to_datetime(df['인증일시'])
df['date'] = df['인증일시'].dt.date
df = df.sort_values(['이름', '인증일시'])

# ====================== 계산 함수 ======================
def calculate_attendance(df):
    results = []
    for (emp, date), group in df.groupby(['이름', 'date']):
        dt = pd.to_datetime(date)
        clock_in = group[group['인증모드']=='출근']['인증일시'].min()
        clock_out = group[group['인증모드']=='퇴근']['인증일시'].max()

        total_h = 0.0
        if pd.notna(clock_in) and pd.notna(clock_out):
            total_h = (clock_out - clock_in).total_seconds() / 3600

        is_weekend = dt.weekday() >= 5
        basic = 9.0 if not is_weekend else 0.0
        ot = min(total_h, 8.0) if is_weekend else max(0, total_h - basic)
        ot = np.floor(ot * 2) / 2

        # 출장시간
        business = 0.0
        out_group = group[group['인증모드'].isin(['외출', '복귀'])]
        if len(out_group) >= 2:
            business = (out_group['인증일시'].max() - out_group['인증일시'].min()).total_seconds() / 3600

        # 8시간 미만 + 출장 없으면 경고
        warning = ""
        if total_h < 8 and business < 1 and not is_weekend:
            warning = "⚠️ 8시간 미만 (연가/출장 확인 필요)"

        results.append({
            '이름': emp,
            '날짜': str(date),
            '요일': dt.strftime('%A')[:3],
            '총근무시간': round(total_h, 2),
            '시간외': round(ot, 2),
            '출장시간': round(business, 2),
            '비고': warning
        })
    
    return pd.DataFrame(results)

result_df = calculate_attendance(df)

# ====================== UI ======================
tabs = st.tabs(["전체 요약"] + sorted(df['이름'].unique().tolist()))

with tabs[0]:
    st.header("📊 전체 근태 현황")
    st.dataframe(result_df, use_container_width=True, hide_index=True)

# 사람별 탭
for i, person in enumerate(sorted(df['이름'].unique()), 1):
    with tabs[i]:
        st.header(f"👤 {person} 상세")
        person_df = result_df[result_df['이름'] == person].copy()
        
        st.subheader("일별 근태")
        st.dataframe(person_df, use_container_width=True, hide_index=True)
        
        st.subheader("📋 수당 계산")
        person_df['출장수당'] = person_df['출장시간'].apply(
            lambda x: 20000 if x >= 4 else 10000 if x > 0 else 0
        )
        
        summary = person_df.agg({
            '총근무시간': 'sum',
            '시간외': 'sum',
            '출장시간': 'sum',
            '출장수당': 'sum'
        }).round(2)
        
        col1, col2 = st.columns(2)
        with col1:
            st.dataframe(person_df[['날짜', '총근무시간', '시간외', '출장시간', '출장수당', '비고']], 
                        use_container_width=True, hide_index=True)
        with col2:
            st.metric("총 근무시간", f"{summary['총근무시간']} 시간")
            st.metric("총 시간외", f"{summary['시간외']} 시간")
            st.metric("총 출장수당", f"{int(summary['출장수당']):,} 원")
        
        st.download_button(
            label=f"{person} 상세 다운로드",
            data=person_df.to_csv(index=False).encode('utf-8'),
            file_name=f"{person}_근태.csv",
            mime="text/csv"
        )

st.success("✅ 모든 기능 적용 완료!")
