import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import io

# -----------------------------------------------------------------------------
# 페이지 기본 설정
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="근태 및 수당 관리 시스템",
    page_icon="📅",
    layout="wide"
)

st.title("📅 근로자 근태 및 수당 지급 명세서 관리 시스템")
st.caption("근로자 5인의 월별 근태 기록 분석, 시간외·출장 수당 계산 및 지급명세서 생성")

# -----------------------------------------------------------------------------
# 사이드바: 1. 파일 업로드 & 샘플 다운로드
# -----------------------------------------------------------------------------
st.sidebar.header("📁 데이터 관리")

# 샘플 데이터 생성 함수
def create_sample_data():
    dates = pd.date_range(start="2026-03-01", end="2026-03-31", freq="D")
    workers = ["홍길동", "김철수", "이영희", "박민수", "정지원"]
    
    rows = []
    for w in workers:
        for d in dates:
            is_weekend = d.weekday() >= 5
            # 샘플 데이터 조합 (출퇴근/외근복귀/공휴일/휴가 등)
            if is_weekend:
                if d.weekday() == 5 and w == "홍길동":  # 토요일 시간외/출장
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
                if d.day == 10 and w == "김철수":  # 누락 데이터
                    rows.append({
                        "날짜": d.strftime("%Y-%m-%d"),
                        "이름": w,
                        "출근시간": "08:30", "퇴근시간": "",
                        "외근시간": "", "복귀시간": "",
                        "휴무여부": "X", "공휴일여부": "X",
                        "비고": "퇴근누락"
                    })
                elif d.day == 15 and w == "이영희":  # 평일 연가(휴무)
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
                        "출근시간": "08:30", "퇴근시간": "19:30",  # 1시간 30분 시간외
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
    label="📄 샘플 엑셀 양식 다운로드",
    data=buffer,
    file_name="근태기록_양식_샘플.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_file = st.sidebar.file_upload_button if hasattr(st.sidebar, 'file_upload_button') else st.sidebar.file_uploader(
    "월별 근태 파일(Excel/CSV) 업로드", type=["xlsx", "csv"]
)

# -----------------------------------------------------------------------------
# 사이드바: 2. 근로자별 단가 설정
# -----------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("💰 근로자별 수당 단가 설정")

# 기본 근로자 목록
default_workers = ["홍길동", "김철수", "이영희", "박민수", "정지원"]
worker_rates = {}

for w in default_workers:
    st.sidebar.subheader(f"👤 {w}")
    col_r1, col_r2 = st.sidebar.columns(2)
    with col_r1:
        base_rate = st.number_input(f"{w} 기본시급(1배)", value=10000, step=500, key=f"base_{w}")
    with col_r2:
        ot_rate = st.number_input(f"{w} 시간외단가", value=15000, step=500, key=f"ot_{w}")
    worker_rates[w] = {"base_rate": base_rate, "ot_rate": ot_rate}

# -----------------------------------------------------------------------------
# 시간/날짜 유틸리티 함수
# -----------------------------------------------------------------------------
def parse_time(t_str):
    if pd.isna(t_str) or not str(t_str).strip():
        return None
    try:
        t_str = str(t_str).strip()
        if len(t_str) == 5: # HH:MM
            return datetime.strptime(t_str, "%H:%M").time()
        elif len(t_str) == 8: # HH:MM:SS
            return datetime.strptime(t_str, "%H:%M:%S").time()
    except:
        return None
    return None

def calc_hours_overlap(start1, end1, start2, end2):
    """두 시간대의 겹치는 시간(시간 단위 Float)을 계산"""
    s1 = datetime.combine(datetime.min, start1)
    e1 = datetime.combine(datetime.min, end1)
    s2 = datetime.combine(datetime.min, start2)
    e2 = datetime.combine(datetime.min, end2)
    
    overlap_start = max(s1, s2)
    overlap_end = min(e1, e2)
    
    if overlap_start < overlap_end:
        return (overlap_end - overlap_start).total_seconds() / 3600.0
    return 0.0

# -----------------------------------------------------------------------------
# 메인 로직 처리
# -----------------------------------------------------------------------------
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"파일을 읽는 도중 오류가 발생했습니다: {e}")
        st.stop()
else:
    st.info("👈 왼쪽 사이드바에서 근태기록 파일(.xlsx / .csv)을 업로드해주세요. 샘플 양식을 내려받아 테스트할 수 있습니다.")
    df_raw = sample_df.copy()

# 데이터 기본 처리
df_raw['날짜'] = pd.to_datetime(df_raw['날짜'])
df_raw['요일'] = df_raw['날짜'].dt.weekday # 0:월, ..., 6:일
df_raw['주차'] = df_raw['날짜'].dt.isocalendar().week

# -----------------------------------------------------------------------------
# UI 탭 구성
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📝 근태 데이터 확인 & 구분 직접 지정", "📊 근태 및 수당 계산 분석", "🧾 개별 수당 지급명세서"])

# -----------------------------------------------------------------------------
# TAB 1: 미기록 날짜 처리 (휴가 / 종일출장 직접 지정)
# -----------------------------------------------------------------------------
with tab1:
    st.subheader("1. 근태 데이터 및 미기록/누락 확인")
    st.write("기록이 없는 날(출퇴근/외근복귀 모두 미기록)은 **휴가(연가)** 또는 **종일출장** 여부를 직접 지정해주세요.")

    # 미기록 데이터 및 평일 출퇴근 누락 찾기
    df_processed = df_raw.copy()
    
    # 세션 상태로 관리
    if 'user_classifications' not in st.session_state:
        st.session_state.user_classifications = {}

    edited_rows = []
    
    for idx, row in df_processed.iterrows():
        is_weekday = row['요일'] < 5
        has_in = pd.notna(row['출근시간']) and str(row['출근시간']).strip() != ""
        has_out = pd.notna(row['퇴근시간']) and str(row['퇴근시간']).strip() != ""
        has_out_work = pd.notna(row['외근시간']) and str(row['외근시간']).strip() != ""
        has_ret_work = pd.notna(row['복귀시간']) and str(row['복귀시간']).strip() != ""
        
        no_records = (not has_in) and (not has_out) and (not has_out_work) and (not has_ret_work)
        is_missing_commute = is_weekday and ((has_in and not has_out) or (not has_in and has_out))
        
        status = "정상"
        if no_records and is_weekday:
            status = "미기록(선택 필요)"
        elif is_missing_commute:
            status = "출퇴근 기록 누락(경고)"

        df_processed.at[idx, '상태'] = status

    # 데이터 에디터로 보여주기
    st.dataframe(df_processed[['날짜', '이름', '출근시간', '퇴근시간', '외근시간', '복귀시간', '휴무여부', '공휴일여부', '상태', '비고']], use_container_width=True)

    st.markdown("---")
    st.subheader("⚠️ 미기록 일자 구분 직접 지정 (휴가 vs 종일출장)")
    
    no_record_df = df_processed[df_processed['상태'] == "미기록(선택 필요)"]
    
    if not no_record_df.empty:
        st.warning(f"총 {len(no_record_df)}건의 미기록 날짜가 있습니다. 아래에서 성격을 지정해주세요.")
        
        grid_cols = st.columns([2, 2, 3, 3])
        grid_cols[0].write("**날짜**")
        grid_cols[1].write("**이름**")
        grid_cols[2].write("**구분 선택**")
        grid_cols[3].write("**설명**")
        
        for idx, row in no_record_df.iterrows():
            key_name = f"select_{row['이름']}_{row['날짜'].strftime('%Y%m%d')}"
            c1, c2, c3, c4 = st.columns([2, 2, 3, 3])
            c1.write(row['날짜'].strftime('%Y-%m-%d (%a)'))
            c2.write(row['이름'])
            selected = c3.selectbox(
                f"구분", 
                ["휴가(연가)", "종일출장", "무급휴무/기타"], 
                key=key_name,
                label_visibility="collapsed"
            )
            c4.write("종일출장 선택 시 8시간 출장 인정" if selected == "종일출장" else "휴가/휴무 처리")
            st.session_state.user_classifications[f"{row['이름']}_{row['날짜'].strftime('%Y-%m-%d')}"] = selected
    else:
        st.success("🎉 미기록된 평일 근태 데이터가 없습니다.")

# -----------------------------------------------------------------------------
# TAB 2: 수당 계산 & 주간/월간 분석
# -----------------------------------------------------------------------------
with tab2:
    st.subheader("2. 시간외 근무 & 출장 수당 종합 산출")
    
    # 일별 산출 상세 결과를 담을 리스트
    daily_results = []
    
    for idx, row in df_raw.iterrows():
        worker = row['이름']
        date_str = row['날짜'].strftime('%Y-%m-%d')
        date_obj = row['날짜']
        is_weekend = row['요일'] >= 5
        is_holiday = str(row['공휴일여부']).upper() == 'O' or is_weekend
        is_off_day = str(row['휴무여부']).upper() == 'O'
        
        # 수당 단가
        rates = worker_rates.get(worker, {"base_rate": 10000, "ot_rate": 15000})
        base_rate = rates['base_rate']
        ot_rate = rates['ot_rate']
        
        t_in = parse_time(row['출근시간'])
        t_out = parse_time(row['퇴근시간'])
        t_out_work = parse_time(row['외근시간'])
        t_ret_work = parse_time(row['복귀시간'])
        
        # 유효성 검사 및 보정
        warning_msg = []
        ot_seconds = 0 # 시간외 근무 초단위
        trip_hours = 0.0 # 출장 인정 시간
        
        # 1. 미기록 직접 지정 값 적용
        user_choice = st.session_state.user_classifications.get(f"{worker}_{date_str}", None)
        if user_choice == "종일출장":
            trip_hours = 8.0
        elif user_choice == "휴가(연가)":
            is_off_day = True
        
        # 2. 평일 출퇴근 누락 처리
        if not is_holiday and not is_off_day:
            if (t_in is None and t_out is not None) or (t_in is not None and t_out is None):
                warning_msg.append("평일 출/퇴근 누락 (시간외 제외)")
        
        # 3. 출장 시간 계산 (외근/복귀 보정)
        if t_out_work is not None or t_ret_work is not None:
            actual_out_work = t_out_work if t_out_work is not None else time(9, 0)
            actual_ret_work = t_ret_work if t_ret_work is not None else time(18, 0)
            
            dt_out = datetime.combine(date_obj, actual_out_work)
            dt_ret = datetime.combine(date_obj, actual_ret_work)
            
            if dt_ret > dt_out:
                trip_hours = (dt_ret - dt_out).total_seconds() / 3600.0

        # 4. 시간외 근무 계산
        if is_holiday:
            # 휴일: 24시간 중 최대 8시간
            if t_in is not None and t_out is not None:
                dt_in = datetime.combine(date_obj, t_in)
                dt_out = datetime.combine(date_obj, t_out)
                if dt_out > dt_in:
                    dur = (dt_out - dt_in).total_seconds()
                    ot_seconds += min(dur, 8 * 3600) # 8시간 제한
        else:
            # 평일: 09:00~18:00 제외한 시간
            if t_in is not None and t_out is not None and ("평일 출/퇴근 누락" not in "".join(warning_msg)):
                # 조기출근 (09:00 이전)
                if t_in < time(9, 0):
                    ot_seconds += (datetime.combine(date_obj, time(9, 0)) - datetime.combine(date_obj, t_in)).total_seconds()
                # 연장근무 (18:00 이후)
                if t_out > time(18, 0):
                    ot_seconds += (datetime.combine(date_obj, t_out) - datetime.combine(date_obj, time(18, 0))).total_seconds()

        # 5. 시간외 vs 출장 중복 시 시간외만 인정 (규칙 3번)
        if is_holiday and trip_hours > 0 and ot_seconds > 0:
            trip_hours = 0.0
            warning_msg.append("출장-시간외 중복 (시간외만 인정)")

        # 6. 휴무시간 계산 (평일 휴무 시)
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

    # -------------------------------------------------------------------------
    # 주간 12시간 초과 검증 & 월 단위 시간 계산
    # -------------------------------------------------------------------------
    st.markdown("### ⚠️ 주간 시간외 근무 12시간 초과 검증")
    
    weekly_ot = df_res.groupby(['이름', '주차'])['시간외초'].sum().reset_index()
    weekly_ot['주간시간외_시간'] = weekly_ot['시간외초'] / 3600.0
    
    over_12h = weekly_ot[weekly_ot['주간시간외_시간'] > 12.0]
    if not over_12h.empty:
        for _, r in over_12h.iterrows():
            st.error(f"🚨 **[경고]** {r['이름']} 근로자 - 주차 {r['주차']}주차 시간외 근무가 **{r['주간시간외_시간']:.1f}시간**으로 주 12시간을 초과했습니다!")
    else:
        st.success("✅ 모든 근로자의 주간 시간외 근무 시간이 12시간 이내입니다.")

    # -------------------------------------------------------------------------
    # 월 단위 수당 계산 (규칙 1: 1시간 미만 버림, 규칙 4: 휴무 1배율 차감 적용)
    # -------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 📊 근로자별 월간 수당 집계 요약")
    
    summary_rows = []
    
    for worker in df_res['이름'].unique():
        w_df = df_res[df_res['이름'] == worker]
        rates = worker_rates.get(worker, {"base_rate": 10000, "ot_rate": 15000})
        
        # 1. 시간외 총 시간 (월 합산 후 1시간 미만 버림)
        total_ot_seconds = w_df['시간외초'].sum()
        total_ot_hours_raw = total_ot_seconds / 3600.0
        final_ot_hours = int(total_ot_hours_raw) # 1시간 미만 버림 (절사)
        
        # 2. 주간 휴무에 따른 단가 변경 적용 (규칙 4번)
        # 평일 휴무시간 만큼은 1배율(기본시급) 적용, 남은 시간외는 지정 단가 적용
        total_off_hours = w_df['휴무시간'].sum()
        
        # 휴무시간 차감 및 배율 계산
        hours_at_base = min(final_ot_hours, total_off_hours)
        hours_at_ot = max(0, final_ot_hours - hours_at_base)
        
        ot_pay = (hours_at_base * rates['base_rate']) + (hours_at_ot * rates['ot_rate'])
        
        # 3. 출장 수당 계산
        # 일별로 4시간 미만 1만원, 4시간 이상 2만원
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
# TAB 3: 근로자별 수당 지급 명세서 생성 및 지출 세부분석
# -----------------------------------------------------------------------------
with tab3:
    st.subheader("3. 근로자별 수당 지급명세서 & 세부 산출 내역")
    
    selected_worker = st.selectbox("명세서를 출력할 근로자를 선택하세요:", df_res['이름'].unique())
    
    if selected_worker:
        w_summary = next(item for item in summary_rows if item["이름"] == selected_worker)
        w_daily = df_res[df_res['이름'] == selected_worker]
        rates = worker_rates.get(selected_worker, {"base_rate": 10000, "ot_rate": 15000})
        
        st.markdown("---")
        # 명세서 헤더
        st.markdown(f"## 📄 [{selected_worker}] 수당 지급 명세서")
        st.caption(f"발행일자: {datetime.now().strftime('%Y-%m-%d')} | 대상월: {df_raw['날짜'].dt.strftime('%Y-%m').iloc[0]}")
        
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("총 지급 수당", f"{w_summary['총支給수당(원)']:,} 원")
        col_m2.metric("시간외 근무 수당", f"{w_summary['시간외수당(원)']:,} 원", f"인정 {w_summary['인정시간외(시간)']}시간")
        col_m3.metric("출장 수당", f"{w_summary['출장수당(원)']:,} 원")
        
        st.markdown("### 🔍 지출 세부 적용 수당 산출 내역")
        
        # 1. 시간외 수당 세부 산출식
        st.markdown("#### 1) 시간외 근무 수당 산출")
        st.write(f"- **월 총 실제 시간외 근무합계:** {w_summary['실제시간외(시간)']} 시간")
        st.write(f"- **월 단위 인정 시간 (1시간 미만 버림):** **{w_summary['인정시간외(시간)']} 시간**")
        st.write(f"- **단가 적용 세부 내역:**")
        st.write(f"  * **1배율 적용 (평일 휴무 차감분):** {w_summary['1배율적용시간']} 시간 × {rates['base_rate']:,} 원 = **{w_summary['1배율적용시간'] * rates['base_rate']:,} 원**")
        st.write(f"  * **시간외 단가 적용:** {w_summary['시간외단가적용시간']} 시간 × {rates['ot_rate']:,} 원 = **{w_summary['시간외단가적용시간'] * rates['ot_rate']:,} 원**")
        st.info(f"💡 **시간외 수당 소계:** {w_summary['시간외수당(원)']:,} 원")

        # 2. 출장 수당 세부 산출식
        st.markdown("#### 2) 출장 수당 산출")
        st.write(f"- **4시간 미만 출장:** {w_summary['출장(4시간미만)']} × 10,000 원")
        st.write(f"- **4시간 이상 출장:** {w_summary['출장(4시간이상)']} × 20,000 원")
        st.info(f"💡 **출장 수당 소계:** {w_summary['출장수당(원)']:,} 원")

        # 3. 일자별 상세 내역 표
        st.markdown("#### 3) 일자별 세부 근태 및 수당 발생 내역")
        
        # 출력용 데이터 정리
        disp_df = w_daily[['날짜', '요일', '휴일여부', '시간외초', '출장시간', '휴무시간', '경고']].copy()
        disp_df['시간외근무(분)'] = (disp_df['시간외초'] / 60).astype(int)
        disp_df = disp_df.drop(columns=['시간외초'])
        
        st.dataframe(disp_df, use_container_width=True)
        
        # 명세서 인쇄/다운로드용 텍스트 생성
        receipt_text = f"""========================================
[수당 지급 명세서]
성명: {selected_worker}
기본시급: {rates['base_rate']:,}원 | 시간외단가: {rates['ot_rate']:,}원
----------------------------------------
1. 시간외수당: {w_summary['시간외수당(원)']:,} 원
   - 인정시간: {w_summary['인정시간외(시간)']}시간 (실제: {w_summary['실제시간외(시간)']}시간)
   - 1배율 적용({w_summary['1배율적용시간']}h) + 시간외단가({w_summary['시간외단가적용시간']}h)
2. 출장수당: {w_summary['출장수당(원)']:,} 원
   - 4시간 미만: {w_summary['출장(4시간미만)']} / 4시간 이상: {w_summary['출장(4시간이상)']}
----------------------------------------
총 지급액: {w_summary['총支給수당(원)']:,} 원
================================--------
"""
        st.download_button(
            label=f"📥 {selected_worker} 명세서 텍스트 다운로드",
            data=receipt_text,
            file_name=f"{selected_worker}_수당지급명세서.txt",
            mime="text/plain"
        )
