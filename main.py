import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, time
import numpy as np
import plotly.express as px

st.set_page_config(page_title="근태 및 수당 관리", layout="wide", initial_sidebar_state="expanded")
st.title("🕒 근로자 근태 및 수당 관리 시스템")
st.markdown("**5명 근로자 (윤신혜, 박은경, 채미혜, 박인미, 조윤희)** 근무시간/시간외/출장 수당 관리")

# ====================== FILE UPLOAD ======================
uploaded_file = st.file_uploader("월별 근무현황 Excel 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("📌 **기본 데이터 (9월)** 로드 중...")
    try:
        df = pd.read_excel('/home/workdir/attachments/근무현황.xlsx', sheet_name='9월', header=2)
        st.success("✅ 2025년 9월 기본 데이터 로드 완료")
    except Exception as e:
        st.error(f"기본 파일 로드 실패: {e}")
        st.stop()
else:
    try:
        df = pd.read_excel(uploaded_file, header=2)
        st.success(f"✅ 업로드 파일 로드 완료: {uploaded_file.name}")
    except:
        st.error("파일 형식 오류 (header=2로 읽힘)")
        st.stop()

# ====================== DATA PREP ======================
df['인증일시'] = pd.to_datetime(df['인증일시'])
df = df.sort_values(['이름', '인증일시']).reset_index(drop=True)
df['date'] = df['인증일시'].dt.date
df['weekday'] = df['인증일시'].dt.weekday  # 0=월 ~ 6=일

employees = sorted(df['이름'].unique().tolist())

# 2025년 9월 공휴일 (검색 결과: 없음)
holidays_2025_sep = []  # pd.to_datetime(['2025-09-xx'])

def is_holiday(dt):
    return dt.date() in [d.date() for d in holidays_2025_sep] or dt.weekday() >= 5

# ====================== SIDEBAR ======================
st.sidebar.header("📊 필터")
selected_emp = st.sidebar.selectbox("근로자", ["전체"] + employees)
month_view = st.sidebar.selectbox("보기 모드", ["월별 요약", "주별 상세", "개인 상세"])

# ====================== ATTENDANCE PROCESSING ======================
@st.cache_data
def process_attendance(df):
    results = []
    weekly_ot = {}  # 주별 시간외 추적 (월~일)

    for (emp, date), group in df.groupby(['이름', df['date']]):
        date_dt = pd.to_datetime(date)
        is_hol = is_holiday(date_dt)
        day_str = date_dt.strftime('%Y-%m-%d')
        week_key = (emp, date_dt.isocalendar()[1])  # ISO week

        if week_key not in weekly_ot:
            weekly_ot[week_key] = 0.0

        records = group.sort_values('인증일시')
        actions = list(zip(records['인증모드'], records['인증일시']))

        # 기본 변수
        clock_in = None
        clock_out = None
        outings = []  # (외출/복귀, 시간)

        for mode, ts in actions:
            if mode == '출근':
                clock_in = ts
            elif mode == '퇴근':
                clock_out = ts
            elif mode in ['외출', '복귀']:
                outings.append((mode, ts))

        total_hours = 0.0
        ot_hours = 0.0
        business_hours = 0.0
        note = ""

        if clock_in and clock_out:
            total_delta = clock_out - clock_in
            total_hours = total_delta.total_seconds() / 3600
        else:
            note = "⚠️ 출/퇴근 기록 불완전"

        # 출장/외출 처리
        if outings:
            if len(outings) % 2 == 0:  # 짝수
                for i in range(0, len(outings), 2):
                    if outings[i][0] == '외출' and outings[i+1][0] == '복귀':
                        bh = (outings[i+1][1] - outings[i][1]).total_seconds() / 3600
                        business_hours += bh
            else:
                # 불완전: 규칙 적용
                note += " (불완전 외출)"
                if outings[0][0] == '외출':
                    assumed_out = outings[0][1]
                    assumed_in = datetime.combine(date, time(18, 0))
                    business_hours += (assumed_in - assumed_out).total_seconds() / 3600
                # etc.

        # 기본 근무시간: 평일 9~18 (9시간)
        basic_hours = 9.0 if not is_hol and date_dt.weekday() < 5 else 0.0

        # 시간외 계산
        if is_hol or date_dt.weekday() >= 5:
            # 주말/공휴일: 전체가 시간외, 최대 8시간
            ot_hours = min(total_hours, 8.0)
        else:
            # 평일: 9-18 제외
            ot_hours = max(0, total_hours - basic_hours)

        # 주간 시간외 12시간 초과 체크
        weekly_ot[week_key] += ot_hours
        if weekly_ot[week_key] > 12:
            note += " 🚨 주간 시간외 12h 초과!"

        # 30분 미만 버림 (floor to 30min unit?)
        ot_hours = np.floor(ot_hours * 2) / 2  # 30분 단위 flooring

        results.append({
            '이름': emp,
            '날짜': date,
            '요일': date_dt.strftime('%A')[:3],
            '총근무시간': round(total_hours, 2),
            '기본근무': basic_hours,
            '시간외': round(ot_hours, 2),
            '출장시간': round(business_hours, 2),
            '비고': note
        })

    return pd.DataFrame(results), weekly_ot

attendance_df, weekly_ot = process_attendance(df)

# ====================== FILTER ======================
if selected_emp != "전체":
    attendance_df = attendance_df[attendance_df['이름'] == selected_emp]

# ====================== DISPLAY ======================
st.header("📅 근태 상세")
col1, col2 = st.columns([3, 1])
with col1:
    st.dataframe(attendance_df, use_container_width=True, hide_index=True)

with col2:
    st.metric("총 기록일", len(attendance_df))
    st.metric("평균 일 근무", f"{attendance_df['총근무시간'].mean():.1f}h" if len(attendance_df)>0 else 0)

# ====================== ALLOWANCE CALC ======================
st.header("💰 수당 계산")

def calculate_allowance(att_df):
    att_df = att_df.copy()
    att_df['출장수당'] = 0
    att_df['시간외수당'] = 0  # 단가 추후 입력

    for idx, row in att_df.iterrows():
        bh = row['출장시간']
        if bh > 0:
            att_df.at[idx, '출장수당'] = 20000 if bh >= 4 else 10000

    # 월/주 요약
    summary = att_df.groupby('이름').agg({
        '총근무시간': 'sum',
        '시간외': 'sum',
        '출장시간': 'sum',
        '출장수당': 'sum'
    }).round(2)
    return att_df, summary

detailed, monthly_summary = calculate_allowance(attendance_df)

if st.button("💸 수당 계산 및 요약", type="primary"):
    st.subheader("월별 수당 요약")
    st.dataframe(monthly_summary, use_container_width=True)

    # Export
    csv = monthly_summary.to_csv().encode('utf-8')
    st.download_button("📥 월 요약 CSV 다운로드", csv, "수당요약.csv", "text/csv")

# ====================== VISUALS ======================
st.header("📈 시각화")
fig = px.bar(attendance_df, x='날짜', y='시간외', color='이름', title="일별 시간외 근무")
st.plotly_chart(fig, use_container_width=True)

# ====================== WARNINGS & NOTES ======================
st.info("""
**규칙 요약**:
- 기본: 평일 09:00~18:00 (9h)
- 시간외: 주말/공휴일 전체 (max 8h/일), 평일 외 시간
- 주간 시간외 max 12h (초과 경고)
- 출장: 4h 미만 1만, 이상 2만
- 기록 미비일: 수동 확인 (연가/종일출장)
- 30분 미만 버림 적용
""")

st.caption("Streamlit 앱 완성형. 추가 단가(시간외수당률, 휴무 보정) 입력 UI는 필요 시 확장 가능.")
