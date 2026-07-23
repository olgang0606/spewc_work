import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

# -----------------------------------------------------------------------------
# 페이지 기본 설정
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="월별 근무현황 및 수당 관리 시스템",
    page_icon="📅",
    layout="wide"
)

st.title("📅 월별 근무현황 및 수당 지급 명세서 관리 시스템")
st.caption("근로자 5인(박은경, 채미혜, 박인미, 조윤희, 성지영) 기준 근무현황 분석 및 수당 계산")

# 지정 근로자 5인 명단
TARGET_WORKERS = ["박은경", "채미혜", "박인미", "조윤희", "성지영"]

# -----------------------------------------------------------------------------
# 사이드바: 1. 월별 파일 업로드 & 샘플 양식 다운로드
# -----------------------------------------------------------------------------
st.sidebar.header("📁 월별 근무현황 파일 업로드")

# 5명 기준 샘플 데이터 생성 함수
def create_sample_data():
    dates = pd.date_range(start="2026-03-01", end="2026-03-31", freq="D")
    
    rows = []
    for w in TARGET_WORKERS:
        for d in dates:
            is_weekend = d.weekday() >= 5
            if is_weekend:
                if d.weekday() == 5 and w == "박은경":
                    rows.append({
                        "날짜": d.strftime("%Y-%m-%d"),
                        "이름": w,
                        "출근시간": "", "퇴근시간": "",
                        "외근시간": "10:00", "복귀시간": "15:00",
                        "휴무여부": "X", "공휴일여부": "O" if d.weekday() == 6 else "X",
                        "비고": "토요일 출장"
                    })
                else:
                    rows.append({
                        "날짜": d.strftime("%Y-%m-%d"),
                        "이름": w,
                        "출근시간": "", "퇴근시간": "",
                        "외근시간": "", "복귀시간": "",
                        "휴무여부": "X", "공휴일여부": "X",
                        "비고": ""
                    })
            else:
                if d.day == 10 and w == "채미혜":
                    rows.append({
                        "날짜": d.strftime("%Y-%m-%d"),
                        "이름": w,
                        "출근시간": "08:30", "퇴근시간": "",
                        "외근시간": "", "복귀시간": "",
                        "휴무여부": "X", "공휴일여부": "X",
                        "비고": "퇴근누락"
                    })
                elif d.day == 15 and w == "박인미":
                    rows.append({
                        "날짜": d.strftime("%Y-%m-%d"),
                        "이름": w,
                        "출근시간": "", "퇴근시간": "",
                        "외근시간": "", "복귀시간": "",
                        "휴무여부": "O", "공휴일여부": "X",
                        "비고": "연가"
                    })
                else:
                    rows.append({
                        "날짜": d.strftime("%Y-%m-%d"),
                        "이름": w,
                        "출근시간": "08:30", "퇴근시간": "19:30",
                        "외근시간": "", "복귀시간": "",
                        "휴무여부": "X", "공휴일여부": "X",
                        "비고": "일반근무"
                    })
    return pd.DataFrame(rows)

sample_df = create_sample_data()
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    sample_df.to_excel(writer, index=False, sheet_name='근태기록')
buffer.seek(0)

st.sidebar.download_button(
    label="📄 5인 이름 반영 샘플 엑셀 다운로드",
    data=buffer,
    file_name="근무현황_양식_5인반영.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_file = st.sidebar.file_uploader(
    "월별 근무현황 파일(Excel/CSV)", 
    type=["xlsx", "csv"]
)

# -----------------------------------------------------------------------------
# 파일 데이터 로드 및 헤더/컬럼 자동 감지
# -----------------------------------------------------------------------------
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name=0)
            
            # [스마트 헤더 감지]: 1행이 제목일 경우 '이름'이나 '날짜'가 포함된 행 찾기
            cols_str = [str(c) for c in df_raw.columns]
            if not any('이름' in c or '날짜' in c or '일자' in c for c in cols_str):
                for idx, row in df_raw.head(10).iterrows():
                    row_vals = [str(v) for v in row.values]
                    if any('이름' in v or '날짜' in v or '일자' in v for v in row_vals):
                        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=idx + 1)
                        break
                        
    except Exception as e:
        st.error(f"파일을 읽는 도중 오류가 발생했습니다: {e}")
        st.stop()
        
    # 컬럼명 특수문자(' / ") 및 공백 일괄 정돈
    df_raw.columns = [str(col).strip().strip("'").strip('"') for col in df_raw.columns]

    # 1. '이름' 관련 컬럼 자동 감지 및 변환
    name_col = None
    for c in df_raw.columns:
        if '이름' in c or '성명' in c:
            name_col = c
            break

    if name_col:
        df_raw.rename(columns={name_col: '이름'}, inplace=True)
        current_names = list(df_raw['이름'].unique())
        if current_names != TARGET_WORKERS and len(current_names) == 5:
            name_map = dict(zip(current_names, TARGET_WORKERS))
            df_raw['이름'] = df_raw['이름'].map(name_map)
            st.sidebar.warning("⚠️ 파일 내 근로자 이름을 [박은경, 채미혜, 박인미, 조윤희, 성지영]으로 자동 변환했습니다.")
    else:
        st.error(f"🚨 '이름' 열을 찾을 수 없습니다. (인식된 열 목록: {list(df_raw.columns)})")
        st.stop()

    # 2. '날짜' 관련 컬럼 자동 감지 및 변환 💡 [신규 추가]
    date_col = None
    for c in df_raw.columns:
        if '날짜' in c or '일자' in c or '근무일' in c or 'date' in c.lower():
            date_col = c
            break

    if date_col:
        df_raw.rename(columns={date_col: '날짜'}, inplace=True)
    else:
        st.error(f"🚨 '날짜' 열을 찾을 수 없습니다. (인식된 열 목록: {list(df_raw.columns)})")
        st.stop()

else:
    df_raw = sample_df.copy()
    st.info("👈 사이드바에서 근무현황 파일을 업로드해주세요.")

# 날짜 데이터 처리
df_raw['날짜'] = pd.to_datetime(df_raw['날짜'])
df_raw['요일'] = df_raw['날짜'].dt.weekday
df_raw['주차'] = df_raw['날짜'].dt.isocalendar().week

target_month_str = df_raw['날짜'].dt.strftime('%Y-%m').iloc[0]
active_workers = list(df_raw['이름'].unique())

st.success(f"📌 **대상월:** `{target_month_str}` | **적용 근로자:** `{', '.join(active_workers)}`")

# -----------------------------------------------------------------------------
# 사이드바: 2. 수당 단가 설정
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("💰 근로자별 수당 단가 설정")

worker_rates = {}
for w in active_workers:
    st.sidebar.subheader(f"👤 {w}")
    col_r1, col_r2 = st.sidebar.columns(2)
    with col_r1:
        base_rate = st.number_input(f"{w} 기본시급(1배)", value=10000, step=500, key=f"base_{w}")
    with col_r2:
        ot_rate = st.number_input(f"{w} 시간외단가", value=15000, step=500, key=f"ot_{w}")
    worker_rates[w] = {"base_rate": base_rate, "ot_rate": ot_rate}

# -----------------------------------------------------------------------------
# 시간 파싱 유틸리티 함수
# -----------------------------------------------------------------------------
def parse_time(t_str):
    if pd.isna(t_str) or not str(t_str).strip():
        return None
    try:
        t_str = str(t_str).strip()
        if len(t_str) == 5:
            return datetime.strptime(t_str, "%H:%M").time()
        elif len(t_str) == 8:
            return datetime.strptime(t_str, "%H:%M:%S").time()
    except:
        return None
    return None

# -----------------------------------------------------------------------------
# UI 탭 구성
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📝 근무현황 데이터 확인 & 직접 지정", "📊 근태 및 수당 계산 분석", "🧾 개별 수당 지급명세서"])

# -----------------------------------------------------------------------------
# TAB 1: 미기록 및 누락 확인
# -----------------------------------------------------------------------------
with tab1:
    st.subheader(f"1. {target_month_str} 근무현황 데이터 확인")

    df_processed = df_raw.copy()
    if 'user_classifications' not in st.session_state:
        st.session_state.user_classifications = {}

    for idx, row in df_processed.iterrows():
        is_weekday = row['요일'] < 5
        
        # 컬럼 존재 여부 안전 체크
        has_in = ('출근시간' in row and pd.notna(row['출근시간']) and str(row['출근시간']).strip() != "")
        has_out = ('퇴근시간' in row and pd.notna(row['퇴근시간']) and str(row['퇴근시간']).strip() != "")
        has_out_work = ('외근시간' in row and pd.notna(row['외근시간']) and str(row['외근시간']).strip() != "")
        has_ret_work = ('복귀시간' in row and pd.notna(row['복귀시간']) and str(row['복귀시간']).strip() != "")
        
        no_records = (not has_in) and (not has_out) and (not has_out_work) and (not has_ret_work)
        is_missing_commute = is_weekday and ((has_in and not has_out) or (not has_in and has_out))
        
        status = "정상"
        if no_records and is_weekday:
            status = "미기록(선택 필요)"
        elif is_missing_commute:
            status = "출퇴근 기록 누락(경고)"

        df_processed.at[idx, '상태'] = status

    display_cols = [c for c in ['날짜', '이름', '출근시간', '퇴근시간', '외근시간', '복귀시간', '휴무여부', '공휴일여부', '상태', '비고'] if c in df_processed.columns]
    st.dataframe(df_processed[display_cols], use_container_width=True)

    st.markdown("---")
    st.subheader("⚠️ 미기록 일자 구분 직접 지정 (휴가 vs 종일출장)")
    
    no_record_df = df_processed[df_processed['상태'] == "미기록(선택 필요)"]
    if not no_record_df.empty:
        for idx, row in no_record_df.iterrows():
            key_name = f"select_{row['이름']}_{row['날짜'].strftime('%Y%m%d')}"
            c1, c2, c3, c4 = st.columns([2, 2, 3, 3])
            c1.write(row['날짜'].strftime('%Y-%m-%d (%a)'))
            c2.write(row['이름'])
            selected = c3.selectbox(
                "구분", 
                ["휴가(연가)", "종일출장", "무급휴무/기타"], 
                key=key_name,
                label_visibility="collapsed"
            )
            c4.write("종일출장 선택 시 8시간 출장 인정" if selected == "종일출장" else "휴가/휴무 처리")
            st.session_state.user_classifications[f"{row['이름']}_{row['날짜'].strftime('%Y-%m-%d')}"] = selected
    else:
        st.success("🎉 미기록된 평일 근무현황 데이터가 없습니다.")

# -----------------------------------------------------------------------------
# TAB 2: 수당 분석
# -----------------------------------------------------------------------------
with tab2:
    st.subheader(f"2. {target_month_str} 수당 산출 결과")
    
    daily_results = []
    for idx, row in df_raw.iterrows():
        worker = row['이름']
        date_str = row['날짜'].strftime('%Y-%m-%d')
        date_obj = row['날짜']
        is_weekend = row['요일'] >= 5
        
        is_holiday = is_weekend or ('공휴일여부' in row and str(row['공휴일여부']).upper() == 'O')
        is_off_day = '휴무여부' in row and str(row['휴무여부']).upper() == 'O'
        
        t_in = parse_time(row.get('출근시간'))
        t_out = parse_time(row.get('퇴근시간'))
        t_out_work = parse_time(row.get('외근시간'))
        t_ret_work = parse_time(row.get('복귀시간'))
        
        warning_msg = []
        ot_seconds = 0
        trip_hours = 0.0
        
        user_choice = st.session_state.user_classifications.get(f"{worker}_{date_str}", None)
        if user_choice == "종일출장":
            trip_hours = 8.0
        elif user_choice == "휴가(연가)":
            is_off_day = True
        
        if not is_holiday and not is_off_day:
            if (t_in is None and t_out is not None) or (t_in is not None and t_out is None):
                warning_msg.append("평일 출/퇴근 누락")
        
        if t_out_work is not None or t_ret_work is not None:
            actual_out_work = t_out_work if t_out_work is not None else time(9, 0)
            actual_ret_work = t_ret_work if t_ret_work is not None else time(18, 0)
            dt_out = datetime.combine(date_obj, actual_out_work)
            dt_ret = datetime.combine(date_obj, actual_ret_work)
            if dt_ret > dt_out:
                trip_hours = (dt_ret - dt_out).total_seconds() / 3600.0

        if is_holiday:
            if t_in is not None and t_out is not None:
                dt_in = datetime.combine(date_obj, t_in)
                dt_out = datetime.combine(date_obj, t_out)
                if dt_out > dt_in:
                    dur = (dt_out - dt_in).total_seconds()
                    ot_seconds += min(dur, 8 * 3600)
        else:
            if t_in is not None and t_out is not None and ("평일 출/퇴근 누락" not in "".join(warning_msg)):
                if t_in < time(9, 0):
                    ot_seconds += (datetime.combine(date_obj, time(9, 0)) - datetime.combine(date_obj, t_in)).total_seconds()
                if t_out > time(18, 0):
                    ot_seconds += (datetime.combine(date_obj, t_out) - datetime.combine(date_obj, time(18, 0))).total_seconds()

        if is_holiday and trip_hours > 0 and ot_seconds > 0:
            trip_hours = 0.0
            warning_msg.append("출장-시간외 중복 (시간외만 인정)")

        off_hours = 8.0 if (is_off_day and not is_holiday) else 0.0

        daily_results.append({
            "주차": row['주차'],
            "날짜": date_str,
            "이름": worker,
            "요일": ["월","화","수","목","금","토","일"][row['요일']],
            "휴일여부": is_holiday,
            "시간외초": ot_seconds,
            "출장시간": trip_hours,
            "휴무시간": off_hours,
            "경고": ", ".join(warning_msg) if warning_msg else "정상"
        })

    df_res = pd.DataFrame(daily_results)

    st.markdown("### 📊 근로자별 월간 수당 집계")
    summary_rows = []
    for worker in active_workers:
        w_df = df_res[df_res['이름'] == worker]
        rates = worker_rates.get(worker, {"base_rate": 10000, "ot_rate": 15000})
        
        total_ot_seconds = w_df['시간외초'].sum()
        total_ot_hours_raw = total_ot_seconds / 3600.0
        final_ot_hours = int(total_ot_hours_raw)
        
        total_off_hours = w_df['휴무시간'].sum()
        hours_at_base = min(final_ot_hours, total_off_hours)
        hours_at_ot = max(0, final_ot_hours - hours_at_base)
        
        ot_pay = (hours_at_base * rates['base_rate']) + (hours_at_ot * rates['ot_rate'])
        
        trip_pay = 0
        trip_count_under_4 = 0
        trip_count_over_4 = 0
        
        for _, r in w_df.iterrows():
            t_h = r['출장시간']
            if t_h > 0:
                if t_h < 4.0:
                    trip_pay += 10000
                    trip_count_under_4 += 1
                else:
                    trip_pay += 20000
                    trip_count_over_4 += 1
                    
        total_pay = ot_pay + trip_pay
        
        summary_rows.append({
            "이름": worker,
            "실제시간외(시간)": round(total_ot_hours_raw, 2),
            "인정시간외(시간)": final_ot_hours,
            "휴무시간(시간)": total_off_hours,
            "1배율적용시간": hours_at_base,
            "시간외단가적용시간": hours_at_ot,
            "시간외수당(원)": ot_pay,
            "출장(4시간미만)": f"{trip_count_under_4}회",
            "출장(4시간이상)": f"{trip_count_over_4}회",
            "출장수당(원)": trip_pay,
            "총支給수당(원)": total_pay
        })

    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df.style.format({
        "시간외수당(원)": "{:,.0f}",
        "출장수당(원)": "{:,.0f}",
        "총支給수당(원)": "{:,.0f}"
    }), use_container_width=True)

# -----------------------------------------------------------------------------
# TAB 3: 지급 명세서
# -----------------------------------------------------------------------------
with tab3:
    st.subheader("3. 근로자별 수당 지급명세서")
    selected_worker = st.selectbox("명세서를 출력할 근로자를 선택하세요:", active_workers)
    
    if selected_worker:
        w_summary = next((item for item in summary_rows if item["이름"] == selected_worker), None)
        w_daily = df_res[df_res['이름'] == selected_worker]
        rates = worker_rates.get(selected_worker, {"base_rate": 10000, "ot_rate": 15000})
        
        if w_summary:
            st.markdown(f"## 📄 [{selected_worker}] 수당 지급 명세서 ({target_month_str})")
            
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("총 지급 수당", f"{w_summary['총支給수당(원)']:,} 원")
            col_m2.metric("시간외 근무 수당", f"{w_summary['시간외수당(원)']:,} 원")
            col_m3.metric("출장 수당", f"{w_summary['출장수당(원)']:,} 원")
            
            disp_df = w_daily[['날짜', '요일', '휴일여부', '시간외초', '출장시간', '휴무시간', '경고']].copy()
            disp_df['시간외근무(분)'] = (disp_df['시간외초'] / 60).astype(int)
            disp_df = disp_df.drop(columns=['시간외초'])
            
            st.dataframe(disp_df, use_container_width=True)
