import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import numpy as np

st.set_page_config(page_title="근태 및 수당 관리", layout="wide")
st.title("🕒 근로자 근태 및 수당 관리 시스템")
st.markdown("**5명 근로자 근무시간 / 시간외 / 출장수당 관리**")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("월별 근무현황 Excel 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("📌 기본 9월 데이터 로드 중...")
    try:
        df = pd.read_excel('/home/workdir/attachments/근무현황.xlsx', sheet_name='9월', header=2)
        st.success("✅ 2025년 9월 기본 데이터 로드 완료")
    except Exception as e:
        st.error("기본 파일을 찾을 수 없습니다. Excel 파일을 업로드하세요.")
        st.stop()
else:
    df = pd.read_excel(uploaded_file, header=2)
    st.success(f"✅ {uploaded_file.name} 로드 완료")

# ====================== DATA PREP ======================
df['인증일시'] = pd.to_datetime(df['인증일시'])
df = df.sort_values(['이름', '인증일시']).reset_index(drop=True)
df['date'] = df['인증일시'].dt.date

employees = sorted(df['이름'].unique())

def is_holiday(dt):
    # 2025년 9월 공휴일 없음. 주말만 처리
    return dt.weekday() >= 5

# ====================== PROCESSING ======================
@st.cache_data
def process_attendance(df):
    results = []
    weekly_ot = {}

    for (emp, date), group in df.groupby(['이름', 'date']):
        date_dt = pd.to_datetime(date)
        week_key = (emp, date_dt.isocalendar()[1])

        if week_key not in weekly_ot:
            weekly_ot[week_key] = 0.0

        records = group.sort_values('인증일시')
        clock_in = None
        clock_out = None
        outings = []
        note = ""

        for _, row in records.iterrows():
            mode = row['인증모드']
            ts = row['인증일시']
            if mode == '출근':
                clock_in = ts
            elif mode == '퇴근':
                clock_out = ts
            elif mode in ['외출', '복귀']:
                outings.append((mode, ts))

        total_hours = 0.0
        business_hours = 0.0

        if clock_in and clock_out:
            total_hours = (clock_out - clock_in).total_seconds() / 3600

        # 출장 처리 (간단)
        if len(outings) >= 2:
            for i in range(0, len(outings)-1, 2):
                if outings[i][0] == '외출' and outings[i+1][0] == '복귀':
                    business_hours += (outings[i+1][1] - outings[i][1]).total_seconds() / 3600

        is_hol = is_holiday(date_dt)
        basic_hours = 9.0 if not is_hol and date_dt.weekday() < 5 else 0.0

        if is_hol or date_dt.weekday() >= 5:
            ot_hours = min(total_hours, 8.0)
        else:
            ot_hours = max(0, total_hours - basic_hours)

        weekly_ot[week_key] += ot_hours
        if weekly_ot[week_key] > 12:
            note = "🚨 주간 시간외 12시간 초과!"

        ot_hours = np.floor(ot_hours * 2) / 2   # 30분 미만 버림

        results.append({
            '이름': emp,
            '날짜': str(date),
            '총근무시간(h)': round(total_hours, 2),
            '기본근무(h)': basic_hours,
            '시간외(h)': round(ot_hours, 2),
            '출장시간(h)': round(business_hours, 2),
            '비고': note
        })

    return pd.DataFrame(results)

attendance_df = process_attendance(df)

# ====================== UI ======================
st.sidebar.header("필터")
selected_emp = st.sidebar.selectbox("근로자 선택", ["전체"] + employees)

if selected_emp != "전체":
    display_df = attendance_df[attendance_df['이름'] == selected_emp]
else:
    display_df = attendance_df

st.header("📅 근태 테이블")
st.dataframe(display_df, use_container_width=True, hide_index=True)

# ====================== 수당 ======================
st.header("💰 수당 계산")
if st.button("수당 계산하기", type="primary"):
    summary = display_df.groupby('이름').agg({
        '총근무시간(h)': 'sum',
        '시간외(h)': 'sum',
        '출장시간(h)': 'sum'
    }).round(2)
    
    summary['출장수당(원)'] = summary['출장시간(h)'].apply(lambda x: sum(20000 if h >= 4 else 10000 for h in [x]))
    
    st.dataframe(summary, use_container_width=True)
    st.download_button("📥 CSV 다운로드", summary.to_csv().encode('utf-8'), "수당_요약.csv")

st.info("**규칙 적용 완료**: 30분 미만 버림, 주간 12h 경고, 출장수당(4h 기준), 주말/평일 구분")
