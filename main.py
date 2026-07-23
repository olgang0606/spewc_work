import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

st.set_page_config(page_title="근태 및 수당 관리", layout="wide")
st.title("🕒 5인 근로자 근태 · 수당 관리 시스템")

# ====================== 파일 업로드 ======================
uploaded_file = st.file_uploader("근무현황.xlsx 업로드", type=["xlsx"])

if uploaded_file is None:
    st.info("기본 9월 데이터 사용 중...")
    try:
        df = pd.read_excel("/home/workdir/attachments/근무현황.xlsx", sheet_name="9월", header=2)
        st.success("✅ 기본 데이터 로드 성공")
    except:
        st.error("기본 파일을 찾을 수 없습니다. Excel 파일을 업로드하세요.")
        st.stop()
else:
    df = pd.read_excel(uploaded_file, header=2)
    st.success("✅ 업로드 파일 로드 성공")

# ====================== 데이터 준비 ======================
df['인증일시'] = pd.to_datetime(df['인증일시'])
df = df.sort_values(['이름', '인증일시']).reset_index(drop=True)
df['date'] = df['인증일시'].dt.date

# ====================== 계산 함수 ======================
def process_attendance(df):
    results = []
    weekly = {}  # 주간 시간외 추적

    for (emp, date), group in df.groupby(['이름', 'date']):
        dt = pd.to_datetime(date)
        week_key = (emp, dt.isocalendar()[1])

        if week_key not in weekly:
            weekly[week_key] = 0.0

        # 출퇴근
        clock_in = group[group['인증모드'] == '출근']['인증일시'].min()
        clock_out = group[group['인증모드'] == '퇴근']['인증일시'].max()

        total_h = 0.0
        if pd.notna(clock_in) and pd.notna(clock_out):
            total_h = (clock_out - clock_in).total_seconds() / 3600

        is_weekend = dt.weekday() >= 5
        basic = 9.0 if not is_weekend else 0.0
        
        # 시간외
        if is_weekend:
            ot = min(total_h, 8.0)
        else:
            ot = max(0, total_h - basic)

        # 30분 미만 버림
        ot = np.floor(ot * 2) / 2

        weekly[week_key] += ot
        warning = "🚨 주간 12h 초과" if weekly[week_key] > 12 else ""

        # 출장 (외출/복귀)
        business = 0.0
        outings = group[group['인증모드'].isin(['외출', '복귀'])]
        if len(outings) >= 2:
            business = (outings['인증일시'].max() - outings['인증일시'].min()).total_seconds() / 3600

        results.append({
            '이름': emp,
            '날짜': str(date),
            '총근무시간': round(total_h, 2),
            '시간외': round(ot, 2),
            '출장시간': round(business, 2),
            '비고': warning
        })

    return pd.DataFrame(results)

result_df = process_attendance(df)

# ====================== 화면 ======================
st.header("📊 근태 결과")
st.dataframe(result_df, use_container_width=True, hide_index=True)

col1, col2 = st.columns(2)
with col1:
    if st.button("💰 수당 계산", type="primary"):
        summary = result_df.groupby('이름').agg({
            '총근무시간': 'sum',
            '시간외': 'sum',
            '출장시간': 'sum'
        }).round(2)
        summary['출장수당(원)'] = (summary['출장시간'] >= 4).astype(int) * 10000 + (summary['출장시간'] > 0).astype(int) * 10000
        st.dataframe(summary, use_container_width=True)

with col2:
    st.download_button("📥 전체 결과 다운로드", 
                      result_df.to_csv(index=False).encode('utf-8'),
                      "근태_결과.csv", "text/csv")

st.caption("✅ 규칙 적용: 30분 버림 | 주말 8h 제한 | 주간 12h 경고 | 출장수당")
