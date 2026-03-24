import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import io
import time
import re
from datetime import datetime, date
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# 페이지 설정
st.set_page_config(
    page_title="충전기 모델분류 자동화",
    page_icon="⚡",
    layout="wide"
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

def clean_and_parse_date(date_value):
    """
    🔥 핵심 기능: 날짜 데이터 정리 및 변환
    - "00:00:00" 같은 시간 부분 제거
    - 다양한 날짜 형식을 Python date 객체로 변환
    - 엑셀 시리얼 날짜도 처리
    """
    if date_value is None:
        return None
    
    # 이미 datetime 또는 date 객체인 경우
    if isinstance(date_value, datetime):
        return date_value.date()
    if isinstance(date_value, date):
        return date_value
    
    # 문자열인 경우 처리
    if isinstance(date_value, str):
        date_str = date_value.strip()
        
        # 🎯 정규식으로 다양한 시간 형식 제거
        date_str = re.sub(r'\s+\d{2}:\d{2}:\d{2}$', '', date_str)      # " 12:34:56"
        date_str = re.sub(r'\s+00:00:00$', '', date_str)               # " 00:00:00"
        date_str = re.sub(r'\s+0:00:00$', '', date_str)                # " 0:00:00"
        
        if not date_str:
            return None
        
        # 📅 다양한 날짜 형식으로 파싱 시도
        date_formats = [
            '%Y-%m-%d',           # 2024-01-15
            '%Y/%m/%d',           # 2024/01/15
            '%Y.%m.%d',           # 2024.01.15
            '%Y%m%d',             # 20240115
            '%Y-%m-%d %H:%M:%S',  # 2024-01-15 12:34:56
            '%Y/%m/%d %H:%M:%S',  # 2024/01/15 12:34:56
            '%d-%m-%Y',           # 15-01-2024
            '%d/%m/%Y',           # 15/01/2024
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
    
    # 숫자인 경우 (엑셀 시리얼 날짜)
    try:
        if isinstance(date_value, (int, float)):
            from datetime import timedelta
            excel_epoch = datetime(1899, 12, 30)  # 엑셀 기준일
            return (excel_epoch + timedelta(days=float(date_value))).date()
    except:
        pass
    
    return None

def format_date_for_excel(date_obj):
    """date 객체를 엑셀용 YYYY-MM-DD 문자열로 변환"""
    if date_obj is None:
        return None
    if isinstance(date_obj, date):
        return date_obj.strftime('%Y-%m-%d')
    return None

def parse_date_safe(date_value):
    """대시보드용 날짜 파싱 (clean_and_parse_date와 동일)"""
    return clean_and_parse_date(date_value)

def classify_region(address):
    """
    H열의 주소 데이터를 기반으로 권역을 분류하는 함수
    🔥 개선: 유연한 패턴 매칭 + 정확한 우선순위
    """
    if not address:
        return "기타"
    
    # 주소 전처리: 공백 정리
    address_clean = str(address).strip()
    
    # 🔥 핵심 개선 1: 인천을 먼저 검사 (우선순위 조정)
    if re.search(r"인천", address_clean):
        # 인천 세부 구/군 검사
        if re.search(r"계양|남동|동구|미추홀|부평|연수|서구|중구|강화", address_clean):
            return "수도권남서"
        else:
            return "인천??"
    
    # 서울/경기 검사 (인천 다음)
    elif re.search(r"서울|경기", address_clean):
        
        # 🔥 핵심 개선 2: "시/구" 없이도 매칭되도록 패턴 수정
        # 수도권북서
        if re.search(r"고양|부천|김포|파주|은평구|마포구|서대문구|양천구|강서구|용산구|중구|종로구", address_clean):
            return "수도권북서"
        
        # 수도권북동 (의정부, 남양주 등 "시" 없이도 매칭)
        elif re.search(r"도봉구|노원구|중랑구|강북구|성북구|동대문구|성동구|광진구|의정부|남양주|구리|양주|포천|동두천|가평|연천", address_clean):
            return "수도권북동"
        
        # 수도권남동
        elif re.search(r"강남구|서초구|송파구|강동구|성남|용인|하남|광주|안성|수원|평택|오산|이천|여주|양평", address_clean):
            return "수도권남동"
        
        # 수도권남서
        elif re.search(r"구로구|금천구|영등포구|동작구|관악구|의왕|광명|군포|과천|시흥|안산|안양|화성", address_clean):
            return "수도권남서"
        
        # 🔥 개선 3: 경기도 기타 지역 처리
        else:
            return "수도권기타"  # "수도권??" 대신 더 명확한 분류
    
    # 광역권 분류
    elif re.search(r"강원", address_clean):
        return "강원권"
    elif re.search(r"충청|충남|충북|세종|대전", address_clean):
        return "충청권"
    elif re.search(r"경상|경남|경북|부산|대구|울산", address_clean):
        return "경상권"
    elif re.search(r"전라|전남|전북|광주", address_clean):
        return "전라권"
    else:
        return "기타"


def classify_model(row_data, row_num):
    """엑셀 IFS 수식을 파이썬 로직으로 변환한 모델 분류 함수"""
    AD = get_safe_value(row_data, 'AD')
    AG = get_safe_value(row_data, 'AG')
    AH = get_safe_value(row_data, 'AH')
    AJ = get_safe_value(row_data, 'AJ')
    
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
    """실시간 타이머와 함께 엑셀 파일을 처리하고 대시보드용 DataFrame 생성"""
    try:
        start_time = time.time()
        
        # 1단계: 파일 로드
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        status_text.markdown("📂 **엑셀 파일을 읽는 중입니다...**")
        progress_bar.progress(5)
        
        file_stream = io.BytesIO(file_bytes)
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active
        
        # 열 번호 정의
        AR_COLUMN = 44  # 운영계약 시작일
        AS_COLUMN = 45  # 운영계약 종료일
        BA_COLUMN = 53  # 모델분류
        BB_COLUMN = 54  # 권역
        
        # 헤더 추가
        ws.cell(row=4, column=BA_COLUMN, value='모델분류')
        ws.cell(row=4, column=BB_COLUMN, value='권역')
        
        # 2단계: 데이터 분석
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        status_text.markdown("🔍 **데이터 구조를 분석하는 중입니다...**")
        progress_bar.progress(15)
        
        max_row = ws.max_row
        if max_row < 5:
            max_row = 5
        
        max_col = max(ws.max_column, 54)
        total_rows = max_row - 4
        
        dashboard_data = []
        
        # 📅 날짜 정리 통계 변수
        ar_cleaned_count = 0
        as_cleaned_count = 0
        
        # 3단계: 분류 및 날짜 정리 작업
        status_text.markdown(f"⚡ **모델분류, 권역분류 및 날짜 정리 작업을 시작합니다... (총 {total_rows:,}개 행)**")
        progress_bar.progress(20)
        
        processed_count = 0
        last_update_time = time.time()
        update_interval = 0.5
        
        for i, row_num in enumerate(range(5, max_row + 1)):
            # 현재 행의 모든 데이터 수집
            row_data = {}
            for col_num in range(1, max_col + 1):
                col_letter = get_column_letter(col_num)
                cell_value = ws.cell(row=row_num, column=col_num).value
                row_data[col_letter] = cell_value
            
            # BA열: 모델 분류
            classification_result = classify_model(row_data, row_num)
            ws.cell(row=row_num, column=BA_COLUMN, value=classification_result)
            
            # BB열: 권역 분류
            address = get_safe_value(row_data, 'H')
            region_result = classify_region(address)
            ws.cell(row=row_num, column=BB_COLUMN, value=region_result)
            
            # 🔥 핵심: AR열(운영계약 시작일) 날짜 정리
            ar_value = row_data.get('AR')
            ar_cleaned = clean_and_parse_date(ar_value)
            if ar_cleaned:
                ar_formatted = format_date_for_excel(ar_cleaned)
                ws.cell(row=row_num, column=AR_COLUMN, value=ar_formatted)
                # 엑셀 날짜 형식 지정
                ws.cell(row=row_num, column=AR_COLUMN).number_format = 'YYYY-MM-DD'
                ar_cleaned_count += 1
            
            # 🔥 핵심: AS열(운영계약 종료일) 날짜 정리
            as_value = row_data.get('AS')
            as_cleaned = clean_and_parse_date(as_value)
            if as_cleaned:
                as_formatted = format_date_for_excel(as_cleaned)
                ws.cell(row=row_num, column=AS_COLUMN, value=as_formatted)
                # 엑셀 날짜 형식 지정
                ws.cell(row=row_num, column=AS_COLUMN).number_format = 'YYYY-MM-DD'
                as_cleaned_count += 1
            
            # 대시보드용 데이터 수집
            dashboard_data.append({
                '모델분류': classification_result,
                '권역': region_result,
                '주소': address,
                '운영계약시작일': ar_value,
                '운영계약종료일': as_value,
                '운영계약시작일_cleaned': ar_cleaned,
                '운영계약종료일_cleaned': as_cleaned,
                '행번호': row_num
            })
            
            processed_count += 1
            
            # 실시간 업데이트
            current_time = time.time()
            should_update = (
                (i + 1) % 50 == 0 or
                (current_time - last_update_time) >= update_interval or
                i == total_rows - 1
            )
            
            if should_update:
                elapsed_time = current_time - start_time
                progress_percent = 20 + int((processed_count / total_rows) * 70)
                progress_bar.progress(progress_percent)
                
                title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed_time)}`")
                
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
        
        # 4단계: 파일 저장
        elapsed = time.time() - start_time
        title_container.markdown(f"### 📊 작업 진행 상황 `⏱️ {format_time(elapsed)}`")
        status_text.markdown("💾 **결과 파일을 생성하는 중입니다...**")
        progress_bar.progress(95)
        
        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        
        # DataFrame 생성 (정리된 날짜 사용)
        df = pd.DataFrame(dashboard_data)
        df['운영계약시작일_parsed'] = df['운영계약시작일_cleaned']
        df['운영계약종료일_parsed'] = df['운영계약종료일_cleaned']
        
        # 완료
        total_time = time.time() - start_time
        progress_bar.progress(100)
        
        title_container.markdown(f"### 🎉 작업 완료! `✅ 총 {format_time(total_time)}`")
        
        avg_speed = processed_count / total_time if total_time > 0 else 0
        status_text.markdown(
            f"🎊 **처리 완료!** `{processed_count:,}개 행`이 성공적으로 분류되었습니다. | "
            f"**평균 속도:** `{avg_speed:.1f}행/초` | "
            f"📅 **날짜 정리:** AR열 `{ar_cleaned_count:,}개`, AS열 `{as_cleaned_count:,}개`"
        )
        
        return output_stream, None, processed_count, total_time, df, ar_cleaned_count, as_cleaned_count
        
    except Exception as e:
        import traceback
        elapsed = time.time() - start_time
        title_container.markdown(f"### ❌ 작업 중단 `⚠️ {format_time(elapsed)}`")
        status_text.markdown("❌ **오류가 발생했습니다.**")
        error_detail = traceback.format_exc()
        return None, f"파일 처리 중 오류 발생: {str(e)}\n\n상세:\n{error_detail}", 0, 0, None, 0, 0

def show_dashboard(df):
    """대시보드 화면을 표시하는 함수"""
    st.markdown("## 📊 충전기 운영 현황 대시보드")
    
    # 날짜 필터 섹션
    st.markdown("### 🗓️ 운영계약 기간 필터")
    
    # 데이터에서 날짜 범위 계산
    valid_dates = df.dropna(subset=['운영계약시작일_parsed', '운영계약종료일_parsed'])
    
    if len(valid_dates) > 0:
        min_date = valid_dates['운영계약시작일_parsed'].min()
        max_date = valid_dates['운영계약종료일_parsed'].max()
        
        col1, col2, col3 = st.columns([2, 2, 1])
        
        with col1:
            start_date = st.date_input(
                "계약 시작일 (이후)",
                value=date(2022, 1, 1),
                min_value=min_date,
                max_value=max_date,
                help="AR열 기준 - 이 날짜 이후에 시작하는 계약"
            )
        
        with col2:
            end_date = st.date_input(
                "계약 종료일 (이전)",
                value=date(2028, 1, 1),
                min_value=min_date,
                max_value=max_date,
                help="AS열 기준 - 이 날짜 이전에 종료하는 계약"
            )
        
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            filter_applied = st.button("🔍 필터 적용", type="primary", use_container_width=True)
        
        # 필터 적용 로직
        mask = (
            (df['운영계약시작일_parsed'] < end_date) & 
            (df['운영계약종료일_parsed'] >= start_date) &
            df['운영계약시작일_parsed'].notna() &
            df['운영계약종료일_parsed'].notna()
        )
        
        filtered_df = df[mask].copy()
        
        # 필터 결과 요약
        st.info(
            f"📅 **선택 기간:** {start_date} ~ {end_date} | "
            f"**해당 기간 충전기:** {len(filtered_df):,}대 (전체 {len(df):,}대 중 {len(filtered_df)/len(df)*100:.1f}%)"
        )
        
        if len(filtered_df) == 0:
            st.warning("⚠️ 선택한 기간에 해당하는 데이터가 없습니다. 필터 조건을 조정해주세요.")
            return
        
        # 대시보드 메인 콘텐츠
        st.markdown("---")
        
        # 1행: 주요 지표 (KPI)
        st.markdown("### 📈 주요 지표")
        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
        
        with kpi_col1:
            total_chargers = len(filtered_df)
            st.metric("총 충전기 수", f"{total_chargers:,}대")
        
        with kpi_col2:
            model_count = filtered_df['모델분류'].nunique()
            st.metric("모델 종류", f"{model_count}개")
        
        with kpi_col3:
            region_count = filtered_df['권역'].nunique()
            st.metric("권역 수", f"{region_count}개")
        
        with kpi_col4:
            fast_chargers = len(filtered_df[filtered_df['모델분류'].str.contains('급속', na=False)])
            fast_ratio = (fast_chargers / total_chargers * 100) if total_chargers > 0 else 0
            st.metric("급속 충전기", f"{fast_chargers:,}대", f"{fast_ratio:.1f}%")
        
        st.markdown("---")
        
        # 2행: 모델별 현황
        st.markdown("### ⚡ 모델별 충전기 현황")
        
        col1, col2 = st.columns([3, 2])
        
        with col1:
            # 모델별 집계
            model_counts = filtered_df['모델분류'].value_counts().reset_index()
            model_counts.columns = ['모델분류', '수량']
            
            # 상위 15개 모델 막대 그래프
            fig_model = px.bar(
                model_counts.head(15),
                x='수량',
                y='모델분류',
                orientation='h',
                title='모델별 충전기 수량 (상위 15개)',
                labels={'수량': '충전기 수량 (대)', '모델분류': '모델'},
                color='수량',
                color_continuous_scale='Blues',
                text='수량'
            )
            fig_model.update_layout(height=500, showlegend=False)
            fig_model.update_traces(texttemplate='%{text}', textposition='outside')
            st.plotly_chart(fig_model, use_container_width=True)
        
        with col2:
            # 모델별 수량 테이블
            st.markdown("#### 📋 모델별 수량 상세")
            model_counts['비율'] = (model_counts['수량'] / model_counts['수량'].sum() * 100).round(1)
            model_counts['비율'] = model_counts['비율'].astype(str) + '%'
            
            st.dataframe(
                model_counts[['모델분류', '수량', '비율']],
                use_container_width=True,
                hide_index=True,
                height=450
            )
        
        st.markdown("---")
        
        # 3행: 권역별 현황
        st.markdown("### 🗺️ 권역별 충전기 현황")
        
        col1, col2 = st.columns([2, 3])
        
        with col1:
            # 권역별 집계
            region_counts = filtered_df['권역'].value_counts().reset_index()
            region_counts.columns = ['권역', '수량']
            
            # 파이 차트
            fig_region_pie = px.pie(
                region_counts,
                values='수량',
                names='권역',
                title='권역별 충전기 비율',
                hole=0.4
            )
            fig_region_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_region_pie.update_layout(height=400)
            st.plotly_chart(fig_region_pie, use_container_width=True)
        
        with col2:
            # 권역별 막대 그래프
            fig_region_bar = px.bar(
                region_counts,
                x='권역',
                y='수량',
                title='권역별 충전기 수량',
                labels={'수량': '충전기 수량 (대)', '권역': '권역'},
                color='수량',
                color_continuous_scale='Greens',
                text='수량'
            )
            fig_region_bar.update_layout(height=400, showlegend=False)
            fig_region_bar.update_traces(texttemplate='%{text}', textposition='outside')
            st.plotly_chart(fig_region_bar, use_container_width=True)
        
        st.markdown("---")
        
        # 4행: 권역별 모델 분포 히트맵
        st.markdown("### 🔥 권역별 모델 분포 히트맵")
        
        # 크로스탭 생성
        crosstab = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'])
        
        # 상위 모델만 표시
        top_models = filtered_df['모델분류'].value_counts().head(12).index
        crosstab_filtered = crosstab[top_models]
        
        fig_heatmap = px.imshow(
            crosstab_filtered.T,
            labels=dict(x="권역", y="모델분류", color="수량"),
            x=crosstab_filtered.index,
            y=crosstab_filtered.columns,
            color_continuous_scale='RdYlGn',
            aspect="auto",
            title='권역별 주요 모델 분포 (상위 12개 모델)',
            text_auto=True
        )
        fig_heatmap.update_layout(height=500)
        st.plotly_chart(fig_heatmap, use_container_width=True)
        
        st.markdown("---")
        
        # 5행: 상세 데이터 테이블
        st.markdown("### 📋 권역별 × 모델별 상세 현황")
        
        # 피벗 테이블 생성
        pivot_wide = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'], margins=True)
        
        # 스타일링된 테이블 표시
        st.dataframe(
            pivot_wide.style.background_gradient(cmap='YlOrRd', axis=None),
            use_container_width=True,
            height=400
        )
        
        # 다운로드 섹션
        st.markdown("---")
        st.markdown("### 📥 데이터 다운로드")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            csv_pivot = pivot_wide.to_csv(encoding='utf-8-sig')
            st.download_button(
                label="📊 권역×모델 현황표 CSV",
                data=csv_pivot,
                file_name=f"권역모델현황_{start_date}_{end_date}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col2:
            csv_filtered = filtered_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📋 필터링된 전체 데이터 CSV",
                data=csv_filtered,
                file_name=f"필터링데이터_{start_date}_{end_date}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col3:
            summary_report = f"""충전기 현황 요약 리포트

📅 분석 기간: {start_date} ~ {end_date}
📊 총 충전기 수: {total_chargers:,}대
🔢 모델 종류: {model_count}개
🗺️ 권역 수: {region_count}개
⚡ 급속 충전기: {fast_chargers:,}대 ({fast_ratio:.1f}%)

📈 상위 5개 모델:
{chr(10).join([f"{i+1}. {row['모델분류']}: {row['수량']:,}대 ({row['비율']})" for i, row in model_counts.head(5).iterrows()])}

🗺️ 권역별 분포:
{chr(10).join([f"• {row['권역']}: {row['수량']:,}대" for _, row in region_counts.iterrows()])}
"""
            
            st.download_button(
                label="📄 요약 리포트 TXT",
                data=summary_report,
                file_name=f"요약리포트_{start_date}_{end_date}.txt",
                mime="text/plain",
                use_container_width=True
            )
    
    else:
        st.warning("⚠️ 유효한 운영계약 날짜 데이터가 없습니다. AR열과 AS열을 확인해주세요.")

# 🔍 미분류 주소 디버깅 섹션
    unknown_regions = filtered_df[
        filtered_df['권역'].str.contains('??|기타', na=False)
    ]
    
    if len(unknown_regions) > 0:
        st.warning(f"⚠️ {len(unknown_regions)}개의 주소가 미분류되었습니다.")
        with st.expander("🔍 미분류 주소 상세 보기"):
            st.dataframe(
                unknown_regions[['주소', '권역']].head(10),
                use_container_width=True
            )

def main():
    # 세션 상태 초기화
    if 'processed_df' not in st.session_state:
        st.session_state.processed_df = None
    if 'processed_file' not in st.session_state:
        st.session_state.processed_file = None
    
    # 헤더
    st.title("⚡ 충전기 모델분류 & 운영현황 대시보드")
    st.markdown("""
    엑셀 파일을 업로드하면 **BA열**에 "모델분류", **BB열**에 "권역"을 자동으로 추가하고,  
    **AR열/AS열**의 날짜 데이터를 정리하여 운영 현황을 분석할 수 있습니다.
    """)
    
    # 탭 구성
    tab1, tab2 = st.tabs(["📁 파일 업로드 & 분류", "📊 운영현황 대시보드"])
    
    with tab1:
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
                
                # 실시간 진행 상황 표시
                title_container = st.empty()
                title_container.markdown("### 📊 작업 진행 상황 `⏱️ 0.0초`")
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                st.markdown("---")
                
                # 파일 처리 실행
                file_bytes = uploaded_file.read()
                result = process_excel_file_with_progress(
                    file_bytes, 
                    title_container,
                    progress_bar, 
                    status_text
                )
                
                processed_file, error, processed_count, total_time, df, ar_cleaned, as_cleaned = result
                
                if error:
                    st.error(f"❌ {error}")
                else:
                    # 세션 상태에 데이터 저장
                    st.session_state.processed_df = df
                    st.session_state.processed_file = processed_file
                    
                    # 성공 메시지
                    st.success(
                        f"🎊 **축하합니다!** 총 **{processed_count:,}개 행**의 모델분류 및 권역분류가 "
                        f"**{format_time(total_time)}**만에 완료되었습니다!"
                    )
                    
                    # 📅 날짜 정리 정보
                    st.info(
                        f"📅 **날짜 정리 완료:** AR열(운영계약시작일) `{ar_cleaned:,}개`, "
                        f"AS열(운영계약종료일) `{as_cleaned:,}개` - 시간 부분(00:00:00) 제거 및 YYYY-MM-DD 형식으로 변환"
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
                    
                    st.info("💡 **'운영현황 대시보드'** 탭으로 이동하여 데이터를 분석해보세요!")
                    
                    # 처리 결과 요약
                    with st.expander("📋 처리 결과 상세 정보", expanded=False):
                        col1, col2, col3, col4, col5 = st.columns(5)
                        with col1:
                            st.metric("처리된 행 수", f"{processed_count:,}개")
                        with col2:
                            st.metric("소요 시간", format_time(total_time))
                        with col3:
                            avg_speed = processed_count / total_time if total_time > 0 else 0
                            st.metric("처리 속도", f"{avg_speed:.1f}행/초")
                        with col4:
                            st.metric("AR열 정리", f"{ar_cleaned:,}개")
                        with col5:
                            st.metric("AS열 정리", f"{as_cleaned:,}개")
        
        # 분류 기준 정보
        with st.expander("💡 분류 기준 정보 보기", expanded=False):
            show_classification_info()
    
    with tab2:
        if st.session_state.processed_df is not None:
            show_dashboard(st.session_state.processed_df)
        else:
            st.info("📁 먼저 **'파일 업로드 & 분류'** 탭에서 파일을 업로드하고 분류 작업을 완료해주세요.")
            
            st.markdown("""
            ### 📊 대시보드에서 확인할 수 있는 정보
            
            **🗓️ 운영계약 기간 필터링**
            - AR열(운영계약 시작일)과 AS열(운영계약 종료일) 기준
            - 날짜 데이터 자동 정리 (00:00:00 제거, YYYY-MM-DD 형식)
            
            **📈 주요 지표, 모델별/권역별 현황**
            - 인터랙티브 차트 및 상세 테이블
            - 권역 × 모델 히트맵
            - 다양한 형식의 데이터 다운로드
            """)

def show_classification_info():
    """분류 기준 정보를 표시하는 함수"""
    st.markdown("### 💡 모델분류 기준표")
    
    subtab1, subtab2, subtab3, subtab4 = st.tabs(["⚡ 급속 충전기", "🔌 완속 충전기", "🗺️ 권역 분류", "📋 참조 정보"])
    
    with subtab1:
        st.markdown("#### ⚡ 급속 충전기 분류 기준")
        st.info("**전제 조건:** AH열 = '급속'")
        
        fast_charger_data = {
            "모델분류명": [
                "급속스필_100", "급속스필_50", "급속PNE_100", "급속PNE_50",
                "급속PNE_200", "급속PNE_150", "급속애플망고_200", "급속SK_100",
                "급속SK_200", "급속코스텔_50", "급속중앙제어_50", "급속그린파워_100",
                "급속그린파워_50", "급속알박_50", "급속"
            ],
            "AG열 코드 조건": [
                "S0F1", "S0F5", "EVQ- (AJ=100)", "EVQ- 또는 EV1- (AJ=50)",
                "MAXE", "DP15", "A01- 또는 AD1-", "Q081, Q101, Q010",
                "Q071, Q102", "1Y25, 1Y24", "1911", "1900",
                "19C0", "QC50", "위 조건 해당 없음"
            ],
            "제조사": [
                "스필", "스필", "PNE", "PNE", "PNE", "PNE", "애플망고",
                "SK", "SK", "코스텔", "중앙제어", "그린파워", "그린파워",
                "알박", "기타 급속"
            ]
        }
        
        st.dataframe(fast_charger_data, use_container_width=True, hide_index=True)
    
    with subtab2:
        st.markdown("#### 🔌 완속 충전기 분류 기준")
        st.info("**전제 조건:** AH열 ≠ '급속'")
        
        slow_charger_data = {
            "우선순위": ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
                        "11", "12", "13", "14", "15", "16", "17", "18", "19"],
            "모델분류명": [
                "알박구형", "알박신형", "10kW", "신형대 (EVL-1C-22CQ)", "구형대 (EVL-1C)",
                "신형대 (EVL+1107)", "구형대 (EVL 기본)", "신형대 (SBDA)", "신형소",
                "F01 / PC01 (SBPA)", "UC01", "스필_7kW", "이카플러그", "중앙제어_7kW",
                "SK_7kW", "3kW", "PNE_7kW", "F01 / PC01 (SBOA)", "기타"
            ]
        }
        
        st.dataframe(slow_charger_data, use_container_width=True, hide_index=True)
    
    with subtab3:
        st.markdown("#### 🗺️ 권역 분류 기준 (H열 주소 기반)")
        
        region_data = {
            "권역명": ["수도권북서", "수도권북동", "수도권남동", "수도권남서", 
                      "수도권남서(인천)", "강원권", "충청권", "경상권", "전라권", "기타"],
            "주요 지역": [
                "고양, 부천, 김포, 파주, 은평, 마포, 서대문, 양천, 강서",
                "도봉, 노원, 중랑, 강북, 성북, 동대문, 의정부, 남양주, 구리",
                "강남, 서초, 송파, 강동, 성남, 용인, 하남, 수원, 평택",
                "구로, 금천, 영등포, 동작, 관악, 의왕, 광명, 안산, 안양",
                "계양, 남동, 부평, 연수, 미추홀 등 인천 주요 구",
                "강원도 전역",
                "충청, 충남, 충북, 세종, 대전",
                "경상, 경남, 경북, 부산, 대구, 울산",
                "전라, 전남, 전북, 광주",
                "위 조건 해당 없음"
            ]
        }
        
        st.dataframe(region_data, use_container_width=True, hide_index=True)
    
    with subtab4:
        st.markdown("#### 📋 엑셀 열 참조 정보")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **🔍 분류에 사용되는 엑셀 열**
            
            | 열 이름 | 열 번호 | 용도 |
            |---------|---------|------|
            | **H열** | 8번째 | 주소 (권역 분류) |
            | **AD열** | 30번째 | 모델명/설명 검색 |
            | **AG열** | 33번째 | 모델 코드 (주요 기준) |
            | **AH열** | 34번째 | 급속/완속 구분 |
            | **AJ열** | 36번째 | 용량 정보 (kW) |
            | **AR열** | 44번째 | 운영계약 시작일 ✨정리 |
            | **AS열** | 45번째 | 운영계약 종료일 ✨정리 |
            """)
        
        with col2:
            st.markdown("""
            **📝 결과 출력 위치**
            
            | 항목 | 위치 | 내용 |
            |------|------|------|
            | **모델분류 헤더** | BA열 4행 | "모델분류" |
            | **권역 헤더** | BB열 4행 | "권역" |
            | **모델분류 결과** | BA열 5행~ | 각 행별 분류값 |
            | **권역 결과** | BB열 5행~ | 각 행별 권역값 |
            | **AR열 정리** | AR열 5행~ | YYYY-MM-DD 형식 |
            | **AS열 정리** | AS열 5행~ | YYYY-MM-DD 형식 |
            """)
        
        st.success("""
        **✨ 날짜 정리 기능:**
        - **AR열, AS열**의 "00:00:00" 같은 시간 부분 자동 제거
        - **YYYY-MM-DD** 형식으로 통일
        - 다양한 날짜 형식 자동 인식 (YYYY/MM/DD, YYYY.MM.DD 등)
        - 엑셀 날짜 형식 지정으로 정렬 및 필터링 최적화
        """)

if __name__ == "__main__":
    main()
