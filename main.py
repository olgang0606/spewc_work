import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np

st.set_page_config(page_title="근태 및 수당 관리", layout="wide")
st.title("🕒 5인 근로자 근태 및 수당 관리 시스템")
st.markdown("**근무현황.xlsx 파일을 업로드하세요**")

# ====================== 파일 업로드 ======================
uploaded_file = st.file_uploader("📁 근무현황.xlsx 업로드", type=["xlsx"])

if uploaded_file is None:
    st.warning("⚠️ Excel 파일을 업로드해주세요.")
    st.stop()

# ====================== 데이터 로드 ======================
df = pd.read_excel(uploaded_file, header=2)
st.success(f"✅ {uploaded_file.name} 로드 완료 ({len(df)}건)")

df['인증일시'] = pd.to_datetime(df['인증일시'])
df['date'] = df['인증일시'].dt.date
df = df.sort_values(['이름', '인증일시'])

# ====================== 근태 계산 ======================
def calculate_attendance(df):
    results = []
    weekly_ot = {}

    for (emp, date), group in df.groupby(['이름', 'date']):
        dt = pd.to_datetime(date)
        week_key = (emp, dt.isocalendar()[1])

        if week_key not in weekly_ot:
            weekly_ot[week_key] = 0.0

        clock_in = group[group['인증모드'] == '출근']['인증일시'].min()
        clock_out = group[group['인증모드'] == '퇴근']['인증일시'].max()

        total_h = 0.0
        if pd.notna(clock_in) and pd.notna(clock_out):
            total_h = (clock_out - clock_in).total_seconds() / 3600

        is_weekend = dt.weekday() >= 5
        basic = 9.0 if not is_weekend else 0.0
        
        # 시간외 계산
        if is_weekend:
            ot = min(total_h, 8.0)
        else:
            ot = max(0, total_h - basic)
        
        ot = np.floor(ot * 2) / 2  # 30분 미만 버림

        weekly_ot[week_key] += ot
        warning = "🚨 주간 시간외 12시간 초과" if weekly_ot[week_key] > 12 else ""

        # 출장시간
        business = 0.0
        out_group = group[group['인증모드'].isin(['외출', '복귀'])]
        if len(out_group) >= 2:
            business = (out_group['인증일시'].max() - out_group['인증일시'].min()).total_seconds() / 3600

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

# ====================== 화면 출력 ======================
st.header("📊 근태 결과")
st.dataframe(result_df, use_container_width=True, hide_index=True)

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("💰 수당 계산", type="primary"):
        summary = result_df.groupby('이름').agg({
            '총근무시간': 'sum',
            '시간외': 'sum',
            '출장시간': 'sum'
        }).round(2)
        
        summary['출장수당(원)'] = summary['출장시간'].apply(
            lambda x: 20000 if x >= 4 else 10000 if x > 0 else 0
        )
        st.dataframe(summary, use_container_width=True)

with col2:
    st.download_button(
        label="📥 전체 결과 CSV 다운로드",
        data=result_df.to_csv(index=False).encode('utf-8'),
        file_name="근태_결과.csv",
        mime="text/csv"
    )

st.info("""
**적용된 규칙**
- 평일: 09:00~18:00 (9시간 기본)
- 주말/공휴일: 전체 시간외 (최대 8시간)
- 30분 미만 버림
- 주간 시간외 12시간 초과 경고
- 출장: 4시간 이상 2만원 / 미만 1만원
""")

st.caption("✅ 완성 버전 | 필요 시 단가 입력 기능 등 추가 가능")
