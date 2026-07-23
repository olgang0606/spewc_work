import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np

st.set_page_config(page_title="근태 및 수당 관리", layout="wide")
st.title("🕒 근로자 근태 · 수당 관리 시스템")
st.markdown("**근무현황.xlsx 파일을 업로드하세요**")

# 파일 업로드 (필수)
uploaded_file = st.file_uploader("📁 근무현황.xlsx 업로드", type=["xlsx"])

if uploaded_file is None:
    st.warning("⚠️ 근무현황.xlsx 파일을 업로드해주세요.")
    st.stop()

# 데이터 로드
df = pd.read_excel(uploaded_file, header=2)
st.success(f"✅ {uploaded_file.name} 파일 로드 완료 ({len(df)}건)")

# 데이터 정리
df['인증일시'] = pd.to_datetime(df['인증일시'])
df['date'] = df['인증일시'].dt.date
df = df.sort_values(['이름', '인증일시'])

# ====================== 계산 ======================
def calculate_attendance(df):
    results = []
    weekly_ot = {}

    for (emp, date), group in df.groupby(['이름', 'date']):
        dt = pd.to_datetime(date)
        week_key = (emp, dt.isocalendar()[1])

        if week_key not in weekly_ot:
            weekly_ot[week_key] = 0.0

        clock_in = group[group['인증모드']=='출근']['인증일시'].min()
        clock_out = group[group['인증모드']=='퇴근']['인증일시'].max()

        total_h = 0.0
        if pd.notna(clock_in) and pd.notna(clock_out):
            total_h = (clock_out - clock_in).total_seconds() / 3600

        is_weekend = dt.weekday() >= 5
        basic = 9.0 if not is_weekend else 0.0
        ot = min(total_h, 8.0) if is_weekend else max(0, total_h - basic)
        ot = np.floor(ot * 2) / 2   # 30분 미만 버림

        weekly_ot[week_key] += ot
        warning = "🚨 주 12시간 초과" if weekly_ot[week_key] > 12 else ""

        # 출장시간 (외출~복귀)
        business = 0.0
        out_times = group[group['인증모드'].isin(['외출', '복귀'])]['인증일시']
        if len(out_times) >= 2:
            business = (out_times.max() - out_times.min()).total_seconds() / 3600

        results.append({
            '이름': emp,
            '날짜': str(date),
            '총근무시간': round(total_h, 2),
            '시간외': round(ot, 2),
            '출장시간': round(business, 2),
            '비고': warning
        })
    
    return pd.DataFrame(results)

result_df = calculate_attendance(df)

# ====================== 화면 ======================
st.header("📊 근태 분석 결과")
st.dataframe(result_df, use_container_width=True, hide_index=True)

if st.button("💰 수당 계산", type="primary"):
    summary = result_df.groupby('이름').agg({
        '총근무시간': 'sum',
        '시간외': 'sum',
        '출장시간': 'sum'
    }).round(2)
    
    summary['출장수당'] = summary['출장시간'].apply(lambda x: 20000 if x >= 4 else 10000 if x > 0 else 0)
    st.dataframe(summary, use_container_width=True)

    csv = summary.to_csv().encode('utf-8')
    st.download_button("📥 수당 요약 다운로드", csv, "수당_요약.csv", "text/csv")

st.info("""
**적용 규칙**
- 평일 기본 9시간 (9~18시)
- 주말/공휴일 전체 시간외 (최대 8시간)
- 30분 미만 버림
- 주간 시간외 12시간 초과 경고
- 출장: 4시간 이상 2만원, 미만 1만원
""")
