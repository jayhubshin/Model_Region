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

def process_excel_file_with_progress(file_bytes, progress_bar, status_text, time_text):
    """진행률과 시간을 표시하면서 엑셀 파일을 처리하는 함수"""
    try:
        start_time = time.time()
        
        # 1단계: 파일 로드 (0-15%)
        status_text.markdown("📂 **엑셀 파일을 읽는 중입니다...**")
        progress_bar.progress(5)
        
        file_stream = io.BytesIO(file_bytes)
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        
        progress_bar.progress(10)
        
        # BA열은 53번째 열
        BA_COLUMN = 53
        
        # BA4 셀에 헤더 "모델분류" 입력
        ws.cell(row=4, column=BA_COLUMN, value='모델분류')
        
        # 2단계: 데이터 분석 (15-20%)
        status_text.markdown("🔍 **데이터 구조를 분석하는 중입니다...**")
        progress_bar.progress(15)
        
        # 데이터가 있는 마지막 행 찾기
        max_row = ws.max_row
        if max_row < 5:
            max_row = 5
        
        # 필요한 열까지 읽기
        max_col = max(ws.max_column, 40)
        
        total_rows = max_row - 4  # 5행부터 시작하므로
        progress_bar.progress(20)
        
        # 3단계: 분류 작업 (20-90%)
        status_text.markdown(f"⚡ **모델 분류 작업을 시작합니다... (총 {total_rows:,}개 행)**")
        
        processed_count = 0
        last_update_time = time.time()
        
        for i, row_num in enumerate(range(5, max_row + 1)):
            # 현재 행의 모든 데이터 수집
            row_data = {}
            for col_num in range(1, max_col + 1):
                col_letter = get_column_letter(col_num)
                cell_value = ws.cell(row=row_num, column=col_num).value
                row_data[col_letter] = cell_value
            
            # 분류 로직 적용
            classification_result = classify_model(row_data, row_num)
            
            # BA열에 분류 결과 입력
            ws.cell(row=row_num, column=BA_COLUMN, value=classification_result)
            processed_count += 1
            
            # 진행률 업데이트 (성능 최적화: 50행마다 또는 1초마다)
            current_time = time.time()
            if (i + 1) % 50 == 0 or (current_time - last_update_time) >= 1.0 or i == total_rows - 1:
                # 진행률 계산 (20% ~ 90% 구간)
                progress_percent = 20 + int((processed_count / total_rows) * 70)
                progress_bar.progress(progress_percent)
                
                # 시간 계산
                elapsed_time = current_time - start_time
                if processed_count > 0:
                    avg_time_per_row = elapsed_time / processed_count
                    remaining_rows = total_rows - processed_count
                    estimated_remaining = avg_time_per_row * remaining_rows
                else:
                    estimated_remaining = 0
                
                # 상태 업데이트
                status_text.markdown(
                    f"⚡ **분류 진행 중...** `{processed_count:,}/{total_rows:,}` 완료 "
                    f"({progress_percent-20:.1f}%)"
                )
                
                # 시간 정보 업데이트
                time_text.markdown(
                    f"⏱️ **경과시간:** `{elapsed_time:.1f}초` | "
                    f"**예상 남은 시간:** `{estimated_remaining:.1f}초`"
                )
                
                last_update_time = current_time
        
        # 4단계: 파일 저장 (90-100%)
        status_text.markdown("💾 **결과 파일을 생성하는 중입니다...**")
        progress_bar.progress(95)
        
        # 메모리 스트림에 저장
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        
        # 완료
        total_time = time.time() - start_time
        progress_bar.progress(100)
        status_text.markdown(f"✅ **처리 완료!** `{processed_count:,}개 행`이 성공적으로 분류되었습니다.")
        time_text.markdown(f"🎉 **총 소요시간:** `{total_time:.2f}초`")
        
        return output_stream, None, processed_count
        
    except Exception as e:
        status_text.markdown("❌ **오류가 발생했습니다.**")
        time_text.markdown("")
        return None, f"파일 처리 중 오류 발생: {str(e)}", 0

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
            
            # 진행률 표시 영역 생성
            st.markdown("### 📊 작업 진행 상황")
            
            # 진행률 바
            progress_bar = st.progress(0)
            
            # 상태 텍스트 (실시간 업데이트)
            status_text = st.empty()
            
            # 시간 정보 (실시간 업데이트)
            time_text = st.empty()
            
            # 구분선
            st.markdown("---")
            
            # 파일 처리 실행
            file_bytes = uploaded_file.read()
            processed_file, error, processed_count = process_excel_file_with_progress(
                file_bytes, 
                progress_bar, 
                status_text,
                time_text
            )
            
            if error:
                st.error(f"❌ {error}")
            else:
                # 성공 메시지
                st.success(f"🎊 **축하합니다!** 총 **{processed_count:,}개 행**의 모델분류가 완료되었습니다!")
                
                # 다운로드 버튼
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_name = f"모델분류_결과_{timestamp}.xlsx"
                
                # 큰 다운로드 버튼
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
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("처리된 행 수", f"{processed_count:,}개")
                    with col2:
                        st.metric("결과 열", "BA열")
                    with col3:
                        st.metric("헤더 위치", "4행")
    
    # 사용 가이드
    st.markdown("---")
    st.markdown("### 💡 분류 기준 정보")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **📌 급속 충전기 분류**
        - 스필 시리즈 (S0F1, S0F5)
        - PNE 시리즈 (EVQ-, EV1-, MAXE, DP15)
        - 애플망고 (A01-, AD1-)
        - SK 시리즈 (Q081, Q101, Q010, Q071, Q102)
        - 기타 (코스텔, 중앙제어, 그린파워, 알박)
        """)
    
    with col2:
        st.markdown("""
        **📌 완속 충전기 분류**
        - 알박 시리즈 (NC07, 23NA, 22NA, 24NA, 25NA)
        - EVL 시리즈 (구형대, 신형대)
        - SB 시리즈 (SBDA, SBAA, SBPA, SBUA, SBOA)
        - 기타 (스필, 이카플러그, 중앙제어, SK, PNE)
        """)
    
    st.markdown("""
    **🔍 참조하는 엑셀 열:**
    - **AD열 (30번째):** 모델명/설명 검색
    - **AG열 (33번째):** 모델 코드 (주요 분류 기준)
    - **AH열 (34번째):** 급속/완속 구분
    - **AJ열 (36번째):** 용량 정보 (kW)
    
    **📝 결과 출력 위치:**
    - **BA열 4행:** "모델분류" 헤더 자동 추가
    - **BA열 5행~:** 각 행별 분류 결과 자동 입력
    """)

if __name__ == "__main__":
    main()
