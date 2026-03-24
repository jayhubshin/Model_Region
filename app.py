import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import io
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

def process_excel_file(file_bytes):
    """엑셀 파일 처리 함수"""
    try:
        file_stream = io.BytesIO(file_bytes)
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        
        # BA열은 53번째 열
        BA_COLUMN = 53
        
        # BA4 셀에 헤더 "모델분류" 입력
        ws.cell(row=4, column=BA_COLUMN, value='모델분류')
        
        # 데이터가 있는 마지막 행 찾기
        max_row = ws.max_row
        if max_row < 5:
            max_row = 5
        
        # 필요한 열까지 읽기
        max_col = max(ws.max_column, 40)
        
        # 5행부터 마지막 행까지 분류 적용
        processed_count = 0
        for row_num in range(5, max_row + 1):
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
        
        # 메모리 스트림에 저장
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        
        return output_stream, None, processed_count
        
    except Exception as e:
        return None, f"파일 처리 중 오류 발생: {str(e)}", 0

# 메인 UI
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
        file_details = {
            "파일명": uploaded_file.name,
            "파일 크기": f"{uploaded_file.size / 1024:.2f} KB"
        }
        st.info(f"📄 **{file_details['파일명']}** ({file_details['파일 크기']})")
        
        # 처리 버튼
        if st.button("🚀 모델분류 시작", type="primary", use_container_width=True):
            with st.spinner("분류 작업을 진행하고 있습니다... 잠시만 기다려 주세요."):
                # 파일 처리
                file_bytes = uploaded_file.read()
                processed_file, error, processed_count = process_excel_file(file_bytes)
                
                if error:
                    st.error(f"❌ {error}")
                else:
                    st.success(f"✅ 분류 완료! **{processed_count}개 행**이 처리되었습니다.")
                    
                    # 다운로드 버튼
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    download_name = f"모델분류_결과_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="📥 결과 파일 다운로드",
                        data=processed_file.getvalue(),
                        file_name=download_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
    
    # 사용 가이드
    st.markdown("---")
    st.markdown("### 💡 분류 기준 정보")
    st.markdown("""
    - **급속 충전기:** AH열이 "급속"인 경우 세부 분류 (스필, PNE, SK, 코스텔 등)
    - **완속 충전기:** 알박, EVL, SB 시리즈 등 제조사별 분류
    - **참조 열:** AD(모델명), AG(코드), AH(타입), AJ(용량)
    - **결과 위치:** BA열 4행(헤더), 5행부터 분류 결과
    """)

if __name__ == "__main__":
    main()
