import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import io
import time
from datetime import datetime

# 페이지 설정
st.set_page_config(
    page_title="충전기 모델분류 자동화",
    page_icon="⚡",
    layout="centered"
)

def get_safe_value(row_data, col_letter):
    """엑셀 셀 값을 안전하게 문자열로 가져오는 헬퍼 함수"""
    val = row_data.get(col_letter)
    if val is None:
        return ""
    return str(val).strip()

def format_time(seconds):
    """초를 읽기 쉬운 형식으로 변환"""
    if seconds < 60:
        return f"{seconds:.1f}초"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{minutes}분 {secs}초"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours}시간 {minutes}분"

def classify_model(row_data, row_num):
    """엑셀 IFS 수식을 파이썬 로직으로 변환한 모델 분류 함수"""
    # 필요한 열 데이터 안전하게 추출
    AD = get_safe_value(row_data, 'AD')  # 30열
    AG = get_safe_value(row_data, 'AG')  # 33열  
    AH = get_safe_value(row_data, 'AH')  # 34열
    AJ = get_safe_value(row_data, 'AJ')  # 36열
    
    # 1단계: AH열이 "급속"인 경우의 세부 분류
    if AH == "급속":
        ag_left4 = AG[:4] if len(AG) >= 4 else AG
        
        if ag_left4 == "S0F1":
            return "급속스필_100"
        elif ag_left4 == "S0F5":
            return "급속스필_50"
        elif ag_left4 == "EVQ-" and AJ == "100":
            return "급속PNE_100"
        elif (ag_left4 == "EVQ-" or ag_left4 == "EV1-") and AJ == "50":
            return "급속PNE_50"
        elif ag_left4 == "MAXE":
            return "급속PNE_200"
        elif ag_left4 == "DP15":
            return "급속PNE_150"
        elif ag_left4 in ["A01-", "AD1-"]:
            return "급속애플망고_200"
        elif ag_left4 in ["Q081", "Q101", "Q010"]:
            return "급속SK_100"
        elif ag_left4 in ["Q071", "Q102"]:
            return "급속SK_200"
        elif ag_left4 in ["1Y25", "1Y24"]:
            return "급속코스텔_50"
        elif ag_left4 == "1911":
            return "급속중앙제어_50"
        elif ag_left4 == "1900":
            return "급속그린파워_100"
        elif ag_left4 == "19C0":
            return "급속그린파워_50"
        elif ag_left4 == "QC50":
            return "급속알박_50"
        else:
            return "급속"
    
    # 2단계: AH열이 "급속"이 아닌 경우의 일반 분류
    ag_left3 = AG[:3] if len(AG) >= 3 else AG
    ag_left4 = AG[:4] if len(AG) >= 4 else AG
    ag_left6 = AG[:6] if len(AG) >= 6 else AG
    ag_left11 = AG[:11] if len(AG) >= 11 else AG
    
    if ag_left4 == "NC07":
        return "알박구형"
    elif ag_left4 in ["23NA", "22NA", "24NA", "25NA"]:
        return "알박신형"
    elif "3J10" in AD:
        return "10kW"
    elif ag_left11 == "EVL-1C-22CQ":
        return "신형대"
    elif ag_left6 == "EVL-1C":
        return "구형대"
    elif ag_left4 == "EVL-" and "1107" in AD:
        return "신형대"
    elif ag_left4 == "EVL-":
        return "구형대"
    elif ag_left4 == "SBDA":
        return "신형대"
    elif ag_left4 == "SBAA":
        return "신형소"
    elif ag_left4 == "SBPA":
        return "F01" if "F01" in AD else "PC01"
    elif ag_left4 == "SBUA":
        return "UC01"
    elif ag_left4 == "SVI0":
        return "스필_7kW"
    elif ag_left3 == "E0C" or "CP" in AD:
        return "이카플러그"
    elif ag_left4 in ["1907", "1912"]:
        return "중앙제어_7kW"
    elif ag_left4 == "SC-P":
        return "SK_7kW"
    elif ag_left4 == "SANA":
        return "3kW"
    elif ag_left4 in ["EVS-", "007S"]:
        return "PNE_7kW"
    elif ag_left4 == "SBOA":
        return "F01" if "F01" in AD else "PC01"
    else:
        return "기타"

def process_excel_file_with_progress(file_bytes, title_container, progress_bar, status_text):
    """실시간 타이머와 함께 엑셀 파일을 처리하는 함수"""
    try:
        start_time = time.time()
        
        # 1단계: 파일 로드 (0-15%)
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        status_text.markdown("📂 **엑셀 파일을 읽는 중입니다...**")
        progress_bar.progress(5)
        
        file_stream = io.BytesIO(file_bytes)
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        progress_bar.progress(10)
        
        # BA열은 53번째 열
        BA_COLUMN = 53
        ws.cell(row=4, column=BA_COLUMN, value='모델분류')
        
        # 2단계: 데이터 분석 (15-20%)
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        status_text.markdown("🔍 **데이터 구조를 분석하는 중입니다...**")
        progress_bar.progress(15)
        
        max_row = ws.max_row
        if max_row < 5:
            max_row = 5
        
        max_col = max(ws.max_column, 40)
        total_rows = max_row - 4  # 5행부터 시작
        
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        progress_bar.progress(20)
        
        # 3단계: 분류 작업 (20-90%)
        status_text.markdown(f"⚡ **모델 분류 작업을 시작합니다... (총 {total_rows:,}개 행)**")
        
        processed_count = 0
        last_update_time = time.time()
        update_interval = 0.5  # 0.5초마다 업데이트
        
        for i, row_num in enumerate(range(5, max_row + 1)):
            # 현재 행의 모든 데이터 수집
            row_data = {}
            for col_num in range(1, max_col + 1):
                col_letter = get_column_letter(col_num)
                cell_value = ws.cell(row=row_num, column=col_num).value
                row_data[col_letter] = cell_value
            
            # 분류 로직 적용
            classification_result = classify_model(row_data, row_num)
            ws.cell(row=row_num, column=BA_COLUMN, value=classification_result)
            processed_count += 1
            
            # 실시간 업데이트 (0.5초마다 또는 50행마다)
            current_time = time.time()
            should_update = (
                (i + 1) % 50 == 0 or  # 50행마다
                (current_time - last_update_time) >= update_interval or  # 0.5초마다
                i == total_rows - 1  # 마지막 행
            )
            
            if should_update:
                elapsed_time = current_time - start_time
                progress_percent = 20 + int((processed_count / total_rows) * 70)
                progress_bar.progress(progress_percent)
                
                # 🔥 핵심: 타이틀에 실시간 타이머 표시
                title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed_time)}`")
                
                # 처리 속도 및 예상 시간 계산
                if processed_count > 0:
                    rows_per_second = processed_count / elapsed_time
                    remaining_rows = total_rows - processed_count
                    estimated_remaining = remaining_rows / rows_per_second if rows_per_second > 0 else 0
                    
                    status_text.markdown(
                        f"⚡ **분류 진행 중...** `{processed_count:,}/{total_rows:,}` 완료 "
                        f"({(processed_count/total_rows*100):.1f}%) | "
                        f"🚀 **처리 속도:** `{rows_per_second:.1f}행/초` | "
                        f"⏳ **예상 남은 시간:** `{format_time(estimated_remaining)}`"
                    )
                
                last_update_time = current_time
        
        # 4단계: 파일 저장 (90-100%)
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        status_text.markdown("💾 **결과 파일을 생성하는 중입니다...**")
        progress_bar.progress(95)
        
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        
        # 완료
        total_time = time.time() - start_time
        progress_bar.progress(100)
        
        # 🎉 완료 시 타이틀 변경
        title_container.markdown(f"### 🎉 작업 완료! `✅ 총 {format_time(total_time)}`")
        
        avg_speed = processed_count / total_time if total_time > 0 else 0
        status_text.markdown(
            f"🎊 **처리 완료!** `{processed_count:,}개 행`이 성공적으로 분류되었습니다. | "
            f"**평균 속도:** `{avg_speed:.1f}행/초`"
        )
        
        return output_stream, None, processed_count, total_time
        
    except Exception as e:
        elapsed = time.time() - start_time
        title_container.markdown(f"### ❌ 작업 중단 `⚠️ {format_time(elapsed)}`")
        status_text.markdown("❌ **오류가 발생했습니다.**")
        return None, f"파일 처리 중 오류 발생: {str(e)}", 0, 0

def main():
    # 헤더
    st.title("⚡ 충전기 모델분류 자동화")
    st.markdown("""
    엑셀 파일을 업로드하면 **BA열**에 "모델분류" 헤더와  
    각 행별 자동 분류 결과가 추가됩니다.
    """)
    
    # 파일 업로더
    uploaded_file = st.file_uploader(
        "📁 엑셀 파일을 선택하세요",
        type=['xlsx', 'xls'],
        help="최대 200MB까지 업로드 가능합니다."
    )
    
    if uploaded_file is not None:
        # 파일 정보 표시
        col1, col2 = st.columns([3, 1])
        with col1:
            st.info(f"📄 **{uploaded_file.name}**")
        with col2:
            file_size_mb = uploaded_file.size / (1024 * 1024)
            if file_size_mb >= 1:
                st.metric("파일 크기", f"{file_size_mb:.1f} MB")
            else:
                st.metric("파일 크기", f"{uploaded_file.size / 1024:.1f} KB")
        
        # 처리 버튼
        if st.button("🚀 모델분류 시작", type="primary", use_container_width=True):
            
            # 🔥 핵심: 실시간 타이머가 포함된 타이틀 컨테이너
            title_container = st.empty()
            title_container.markdown("### 📊 작업 진행 상황 `⏱️ 0.0초`")
            
            # 진행률 바
            progress_bar = st.progress(0)
            
            # 상태 텍스트
            status_text = st.empty()
            
            # 구분선
            st.markdown("---")
            
            # 파일 처리 실행
            file_bytes = uploaded_file.read()
            processed_file, error, processed_count, total_time = process_excel_file_with_progress(
                file_bytes, 
                title_container,  # 타이틀 컨테이너 전달
                progress_bar, 
                status_text
            )
            
            if error:
                st.error(f"❌ {error}")
            else:
                # 성공 메시지
                st.success(
                    f"🎊 **축하합니다!** 총 **{processed_count:,}개 행**의 모델분류가 "
                    f"**{format_time(total_time)}**만에 완료되었습니다!"
                )
                
                # 다운로드 버튼
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_name = f"모델분류_결과_{timestamp}.xlsx"
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.download_button(
                        label="📥 결과 파일 다운로드",
                        data=processed_file.getvalue(),
                        file_name=download_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type="primary"
                    )
                
                # 처리 결과 요약
                with st.expander("📋 처리 결과 상세 정보", expanded=False):
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("처리된 행 수", f"{processed_count:,}개")
                    with col2:
                        st.metric("소요 시간", format_time(total_time))
                    with col3:
                        avg_speed = processed_count / total_time if total_time > 0 else 0
                        st.metric("처리 속도", f"{avg_speed:.1f}행/초")
                    with col4:
                        st.metric("결과 열", "BA열")
    
    # ========================================================================
    # 🎯 핵심 개선: 분류 기준 정보를 체계적인 표로 정리
    # ========================================================================
    st.markdown("---")
    st.markdown("### 💡 모델분류 기준표")
    
    # 탭으로 구분하여 정보 구조화
    tab1, tab2, tab3 = st.tabs(["⚡ 급속 충전기", "🔌 완속 충전기", "📋 참조 정보"])
    
    with tab1:
        st.markdown("#### ⚡ 급속 충전기 분류 기준")
        st.info("**전제 조건:** AH열 = '급속'")
        
        # 급속 충전기 분류표 (검색 가능한 데이터프레임)
        fast_charger_data = {
            "모델분류명": [
                "급속스필_100",
                "급속스필_50", 
                "급속PNE_100",
                "급속PNE_50",
                "급속PNE_200",
                "급속PNE_150",
                "급속애플망고_200",
                "급속SK_100",
                "급속SK_200", 
                "급속코스텔_50",
                "급속중앙제어_50",
                "급속그린파워_100",
                "급속그린파워_50",
                "급속알박_50",
                "급속"
            ],
            "AG열 코드 조건": [
                "S0F1",
                "S0F5",
                "EVQ- (AJ=100)",
                "EVQ- 또는 EV1- (AJ=50)", 
                "MAXE",
                "DP15",
                "A01- 또는 AD1-",
                "Q081, Q101, Q010",
                "Q071, Q102",
                "1Y25, 1Y24",
                "1911",
                "1900", 
                "19C0",
                "QC50",
                "위 조건 해당 없음"
            ],
            "제조사": [
                "스필",
                "스필",
                "PNE", 
                "PNE",
                "PNE",
                "PNE",
                "애플망고",
                "SK",
                "SK",
                "코스텔",
                "중앙제어",
                "그린파워",
                "그린파워", 
                "알박",
                "기타 급속"
            ]
        }
        
        st.dataframe(
            fast_charger_data,
            use_container_width=True,
            hide_index=True
        )
        
        st.warning("⚠️ **중요:** 급속 충전기는 위에서 아래 순서대로 조건을 검사하며, 먼저 일치하는 조건이 적용됩니다.")
    
    with tab2:
        st.markdown("#### 🔌 완속 충전기 분류 기준") 
        st.info("**전제 조건:** AH열 ≠ '급속' (완속 또는 기타)")
        
        # 완속 충전기 분류표
        slow_charger_data = {
            "우선순위": [
                "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", 
                "11", "12", "13", "14", "15", "16", "17", "18", "19"
            ],
            "모델분류명": [
                "알박구형",
                "알박신형", 
                "10kW",
                "신형대 (EVL-1C-22CQ)",
                "구형대 (EVL-1C)", 
                "신형대 (EVL+1107)",
                "구형대 (EVL 기본)",
                "신형대 (SBDA)",
                "신형소",
                "F01 / PC01 (SBPA)",
                "UC01",
                "스필_7kW",
                "이카플러그",
                "중앙제어_7kW",
                "SK_7kW",
                "3kW", 
                "PNE_7kW",
                "F01 / PC01 (SBOA)",
                "기타"
            ],
            "판별 조건": [
                "AG열 처음 4자리 = NC07",
                "AG열 처음 4자리 = 23NA/22NA/24NA/25NA",
                "AD열에 '3J10' 포함",
                "AG열 처음 11자리 = EVL-1C-22CQ", 
                "AG열 처음 6자리 = EVL-1C",
                "AG열 처음 4자리 = EVL- AND AD열에 '1107' 포함",
                "AG열 처음 4자리 = EVL-",
                "AG열 처음 4자리 = SBDA",
                "AG열 처음 4자리 = SBAA",
                "AG열 처음 4자리 = SBPA (AD열 F01 유무로 구분)",
                "AG열 처음 4자리 = SBUA",
                "AG열 처음 4자리 = SVI0",
                "AG열 처음 3자리 = E0C OR AD열에 'CP' 포함",
                "AG열 처음 4자리 = 1907/1912", 
                "AG열 처음 4자리 = SC-P",
                "AG열 처음 4자리 = SANA",
                "AG열 처음 4자리 = EVS-/007S",
                "AG열 처음 4자리 = SBOA (AD열 F01 유무로 구분)",
                "위 모든 조건에 해당 없음"
            ],
            "제조사/시리즈": [
                "알박",
                "알박",
                "특수용량",
                "EVL",
                "EVL", 
                "EVL",
                "EVL", 
                "SB시리즈",
                "SB시리즈",
                "SB시리즈",
                "SB시리즈",
                "스필",
                "이카플러그",
                "중앙제어",
                "SK",
                "일반용량",
                "PNE",
                "SB시리즈", 
                "미분류"
            ]
        }
        
        st.dataframe(
            slow_charger_data,
            use_container_width=True,
            hide_index=True
        )
        
        st.error("🚨 **핵심:** 완속 충전기는 **우선순위** 순서대로 조건을 검사합니다. EVL 시리즈는 긴 코드부터 먼저 검사하므로 순서가 매우 중요합니다!")
    
    with tab3:
        st.markdown("#### 📋 엑셀 열 참조 정보")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **🔍 분류에 사용되는 엑셀 열**
            
            | 열 이름 | 열 번호 | 용도 |
            |---------|---------|------|
            | **AD열** | 30번째 | 모델명/설명 검색 |
            | **AG열** | 33번째 | 모델 코드 (주요 기준) |
            | **AH열** | 34번째 | 급속/완속 구분 |
            | **AJ열** | 36번째 | 용량 정보 (kW) |
            """)
        
        with col2:
            st.markdown("""
            **📝 결과 출력 위치**
            
            | 항목 | 위치 | 내용 |
            |------|------|------|
            | **헤더** | BA열 4행 | "모델분류" |
            | **결과** | BA열 5행~ | 각 행별 분류값 |
            | **BA열** | 53번째 열 | 결과 출력 열 |
            """)
        
        st.markdown("---")
        
        # 분류 로직 상세 설명
        st.markdown("**💡 분류 로직 흐름도**")
        st.markdown("""
        ```
        1️⃣ AH열 확인
           └── '급속' → 급속 충전기 분류표 적용
           └── '급속' 아님 → 완속 충전기 분류표 적용
        
        2️⃣ AG열 코드 패턴 매칭 (우선순위 순)
           └── 긴 패턴부터 검사 (11자리 → 6자리 → 4자리 → 3자리)
        
        3️⃣ AD열 키워드 검색
           └── 특정 문자열 포함 여부 확인
        
        4️⃣ AJ열 용량 확인 (급속 충전기만)
           └── EVQ-, EV1- 코드의 세부 분류
        
        5️⃣ 기본값 적용
           └── 모든 조건 불일치 시 '급속' 또는 '기타'
        ```
        """)
        
        st.success("""
        **✨ 특별 규칙:**
        - **EVL 시리즈:** 11자리 → 6자리 → 4자리 순서로 검사하여 가장 구체적인 조건 우선 적용
        - **SBPA/SBOA:** AD열에 'F01' 포함 시 'F01', 미포함 시 'PC01'로 분류  
        - **급속 충전기:** AH='급속' + AG코드 + AJ용량 조합으로 세부 분류
        - **완속 충전기:** AG코드 + AD검색어 조합으로 제조사별 분류
        """)

if __name__ == "__main__":
    main()
