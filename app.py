import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import io
import os
import time
import re
import numpy as np
from datetime import datetime, date
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pytz
import folium
from streamlit_folium import st_folium
from folium.plugins import MarkerCluster
from urllib.parse import quote

st.set_page_config(
    page_title="충전기 모델분류 자동화",
    page_icon="⚡",
    layout="wide"
)

DEFAULT_DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "default_data.xlsx")

# ──────────────────────────────────────────────
# 유틸리티 함수
# ──────────────────────────────────────────────

def get_korea_time():
    korea_tz = pytz.timezone('Asia/Seoul')
    return datetime.now(korea_tz)

def format_time(seconds):
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
    if date_value is None or (isinstance(date_value, float) and np.isnan(date_value)):
        return None
    if isinstance(date_value, datetime):
        return date_value.date()
    if isinstance(date_value, date):
        return date_value
    if isinstance(date_value, pd.Timestamp):
        return date_value.date()
    if isinstance(date_value, str):
        date_str = date_value.strip()
        date_str = re.sub(r'\s+\d{1,2}:\d{2}:\d{2}$', '', date_str)
        if not date_str:
            return None
        for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d',
                     '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%d-%m-%Y', '%d/%m/%Y']:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
    try:
        if isinstance(date_value, (int, float)):
            from datetime import timedelta
            excel_epoch = datetime(1899, 12, 30)
            return (excel_epoch + timedelta(days=float(date_value))).date()
    except:
        pass
    return None

def format_date_for_excel(date_obj):
    if date_obj is None:
        return None
    if isinstance(date_obj, date):
        return date_obj.strftime('%Y-%m-%d')
    return None


# ──────────────────────────────────────────────
# ★ 벡터화 분류 함수 (핵심 성능 개선)
# ──────────────────────────────────────────────

def classify_region_vectorized(addresses: pd.Series) -> pd.Series:
    """주소 Series를 받아 권역 Series를 반환 (벡터 연산)"""
    addr = addresses.fillna('').astype(str).str.strip()
    result = pd.Series('기타', index=addr.index)

    # 인천
    incheon_mask = addr.str.contains('인천', na=False)
    incheon_detail = addr.str.contains('계양|남동|동구|미추홀|부평|연수|서구|중구|강화', na=False)
    result[incheon_mask & incheon_detail] = '수도권남서'
    result[incheon_mask & ~incheon_detail] = '인천기타'

    # 서울/경기
    sg_mask = addr.str.contains('서울|경기', na=False) & ~incheon_mask
    nw = addr.str.contains('고양|부천|김포|파주|은평구|마포구|서대문구|양천구|강서구|용산구|중구|종로구', na=False)
    ne = addr.str.contains('도봉구|노원구|중랑구|강북구|성북구|동대문구|성동구|광진구|의정부|남양주|구리|양주|포천|동두천|가평|연천', na=False)
    se = addr.str.contains('강남구|서초구|송파구|강동구|성남|용인|하남|광주|안성|수원|평택|오산|이천|여주|양평', na=False)
    sw = addr.str.contains('구로구|금천구|영등포구|동작구|관악구|의왕|광명|군포|과천|시흥|안산|안양|화성', na=False)

    result[sg_mask & nw] = '수도권북서'
    result[sg_mask & ne & ~nw] = '수도권북동'
    result[sg_mask & se & ~nw & ~ne] = '수도권남동'
    result[sg_mask & sw & ~nw & ~ne & ~se] = '수도권남서'
    result[sg_mask & ~nw & ~ne & ~se & ~sw] = '수도권기타'

    # 나머지 광역
    result[addr.str.contains('강원', na=False) & ~incheon_mask & ~sg_mask] = '강원권'
    result[addr.str.contains('충청|충남|충북|세종|대전', na=False) & ~incheon_mask & ~sg_mask] = '충청권'
    result[addr.str.contains('경상|경남|경북|부산|대구|울산', na=False) & ~incheon_mask & ~sg_mask] = '경상권'
    result[addr.str.contains('전라|전남|전북|광주', na=False) & ~incheon_mask & ~sg_mask] = '전라권'

    return result


def classify_model_vectorized(df: pd.DataFrame) -> pd.Series:
    """DataFrame의 AD, AG, AH, AJ 열을 기반으로 모델 분류 (벡터 연산)"""
    AD = df['AD'].fillna('').astype(str).str.strip()
    AG = df['AG'].fillna('').astype(str).str.strip()
    AH = df['AH'].fillna('').astype(str).str.strip()
    AJ = df['AJ'].fillna('').astype(str).str.strip()

    ag4 = AG.str[:4]
    ag3 = AG.str[:3]
    ag6 = AG.str[:6]
    ag11 = AG.str[:11]

    is_fast = (AH == '급속')

    result = pd.Series('기타', index=df.index)

    # ── 급속 분류 ──
    fast_conditions = [
        is_fast & (ag4 == 'S0F1'),
        is_fast & (ag4 == 'S0F5'),
        is_fast & (ag4 == 'EVQ-') & (AJ == '100'),
        is_fast & (ag4.isin(['EVQ-', 'EV1-'])) & (AJ == '50'),
        is_fast & (ag4 == 'MAXE'),
        is_fast & (ag4 == 'DP15'),
        is_fast & (ag4.isin(['A01-', 'AD1-'])),
        is_fast & (ag4.isin(['Q081', 'Q101', 'Q010'])),
        is_fast & (ag4.isin(['Q071', 'Q102'])),
        is_fast & (ag4.isin(['1Y25', '1Y24'])),
        is_fast & (ag4 == '1911'),
        is_fast & (ag4 == '1900'),
        is_fast & (ag4 == '19C0'),
        is_fast & (ag4 == 'QC50'),
    ]
    fast_values = [
        '급속스필_100', '급속스필_50', '급속PNE_100', '급속PNE_50',
        '급속PNE_200', '급속PNE_150', '급속애플망고_200', '급속SK_100',
        '급속SK_200', '급속코스텔_50', '급속중앙제어_50', '급속그린파워_100',
        '급속그린파워_50', '급속알박_50',
    ]

    # np.select: 조건 목록에서 첫 매칭 반환
    result = pd.Series(
        np.select(fast_conditions, fast_values, default='__PENDING__'),
        index=df.index
    )
    # 급속이지만 세부 매칭 안 된 것
    result[(result == '__PENDING__') & is_fast] = '급속'

    # ── 완속 분류 (급속이 아닌 행만) ──
    slow_mask = ~is_fast
    slow_conditions = [
        slow_mask & (ag4 == 'NC07'),
        slow_mask & (ag4.isin(['23NA', '22NA', '24NA', '25NA'])),
        slow_mask & (AD.str.contains('3J10', na=False)),
        slow_mask & (ag11 == 'EVL-1C-22CQ'),
        slow_mask & (ag6 == 'EVL-1C') & (ag11 != 'EVL-1C-22CQ'),
        slow_mask & (ag4 == 'EVL-') & (AD.str.contains('1107', na=False)) & (ag6 != 'EVL-1C'),
        slow_mask & (ag4 == 'EVL-') & (~AD.str.contains('1107', na=False)) & (ag6 != 'EVL-1C'),
        slow_mask & (ag4 == 'SBDA'),
        slow_mask & (ag4 == 'SBAA'),
        slow_mask & (ag4 == 'SBPA') & (AD.str.contains('F01', na=False)),
        slow_mask & (ag4 == 'SBPA') & (~AD.str.contains('F01', na=False)),
        slow_mask & (ag4 == 'SBUA'),
        slow_mask & (ag4 == 'SVI0'),
        slow_mask & ((ag3 == 'E0C') | (AD.str.contains('CP', na=False))),
        slow_mask & (ag4.isin(['1907', '1912'])),
        slow_mask & (ag4 == 'SC-P'),
        slow_mask & (ag4 == 'SANA'),
        slow_mask & (ag4.isin(['EVS-', '007S'])),
        slow_mask & (ag4 == 'SBOA') & (AD.str.contains('F01', na=False)),
        slow_mask & (ag4 == 'SBOA') & (~AD.str.contains('F01', na=False)),
    ]
    slow_values = [
        '알박구형', '알박신형', '10kW', '신형대', '구형대',
        '신형대', '구형대', '신형대', '신형소',
        'F01', 'PC01', 'UC01', '스필_7kW', '이카플러그',
        '중앙제어_7kW', 'SK_7kW', '3kW', 'PNE_7kW', 'F01', 'PC01',
    ]

    slow_result = np.select(slow_conditions, slow_values, default='__KEEP__')

    # 완속 결과를 반영 (급속 분류된 건 유지)
    pending_slow = (result == '__PENDING__') & slow_mask
    result[pending_slow] = pd.Series(slow_result, index=df.index)[pending_slow]
    result[result == '__PENDING__'] = '기타'
    result[result == '__KEEP__'] = '기타'  # slow_mask지만 조건 미매칭

    # slow_result가 __KEEP__이 아닌 건 덮어쓰기 (result가 아직 __PENDING__인 것만)
    # 위에서 이미 처리됨. 추가로, slow 분류가 매칭된 건 중 result가 '기타'인 것 갱신
    slow_matched = pd.Series(slow_result, index=df.index)
    override_mask = slow_mask & (slow_matched != '__KEEP__')
    result[override_mask] = slow_matched[override_mask]

    return result


# ──────────────────────────────────────────────
# ★ 고속 엑셀 로드 함수
# ──────────────────────────────────────────────

def load_excel_to_dataframe(source, header_row=3):
    """
    엑셀 파일을 pandas로 빠르게 읽어 DataFrame 반환.
    source: 파일 경로(str) 또는 BytesIO
    header_row: 헤더가 있는 행 (0-indexed). 기본 3 = 엑셀 4행
    """
    df = pd.read_excel(
        source,
        header=header_row,
        engine='openpyxl',
        dtype=str  # 모두 문자열로 읽어서 타입 문제 방지
    )
    return df


def build_dashboard_df_from_raw(raw_df):
    """
    pandas로 읽은 원본 DataFrame에서 분류를 수행하고 대시보드용 DataFrame을 생성.
    열 이름이 아니라 열 위치(인덱스)로 접근하여 헤더명에 무관하게 동작.
    """
    cols = raw_df.columns.tolist()

    # 열 위치 매핑 (0-indexed: A=0, H=7, AD=29, AG=32, AH=33, AJ=35, AM=38, AN=39, AR=43, AS=44)
    def safe_col(idx):
        if idx < len(cols):
            return cols[idx]
        return None

    col_A = safe_col(0)    # 사이트ID
    col_H = safe_col(7)    # 주소
    col_AD = safe_col(29)  # 모델명
    col_AG = safe_col(32)  # 모델코드
    col_AH = safe_col(33)  # 급속/완속
    col_AJ = safe_col(35)  # 용량
    col_AM = safe_col(38)  # 경도
    col_AN = safe_col(39)  # 위도
    col_AR = safe_col(43)  # 계약시작일
    col_AS = safe_col(44)  # 계약종료일

    # 분류용 임시 DF 구성
    classify_df = pd.DataFrame({
        'AD': raw_df[col_AD] if col_AD else '',
        'AG': raw_df[col_AG] if col_AG else '',
        'AH': raw_df[col_AH] if col_AH else '',
        'AJ': raw_df[col_AJ] if col_AJ else '',
    })

    addresses = raw_df[col_H] if col_H else pd.Series('', index=raw_df.index)

    # ★ 벡터 분류 (핵심 성능)
    model_result = classify_model_vectorized(classify_df)
    region_result = classify_region_vectorized(addresses)

    # 날짜 파싱
    ar_series = raw_df[col_AR] if col_AR else pd.Series(dtype=object)
    as_series = raw_df[col_AS] if col_AS else pd.Series(dtype=object)

    ar_parsed = pd.to_datetime(ar_series, errors='coerce', infer_datetime_format=True)
    as_parsed = pd.to_datetime(as_series, errors='coerce', infer_datetime_format=True)

    # 좌표
    lon = pd.to_numeric(raw_df[col_AM], errors='coerce') if col_AM else pd.Series(dtype=float)
    lat = pd.to_numeric(raw_df[col_AN], errors='coerce') if col_AN else pd.Series(dtype=float)

    # 사이트ID
    site_ids = raw_df[col_A].fillna('').astype(str).str.strip() if col_A else pd.Series('', index=raw_df.index)
    site_ids[site_ids == ''] = 'AUTO_' + (raw_df.index + 5).astype(str)

    dashboard_df = pd.DataFrame({
        '사이트ID': site_ids,
        '모델분류': model_result,
        '권역': region_result,
        '주소': addresses,
        '위도': lat,
        '경도': lon,
        '운영계약시작일': ar_series.values if col_AR else None,
        '운영계약종료일': as_series.values if col_AS else None,
        '운영계약시작일_parsed': ar_parsed,
        '운영계약종료일_parsed': as_parsed,
        '운영계약시작일_cleaned': ar_parsed.dt.date,
        '운영계약종료일_cleaned': as_parsed.dt.date,
        '행번호': raw_df.index + 5,
    })

    return dashboard_df


def load_default_excel(filepath):
    """default_data.xlsx를 고속으로 로드"""
    raw_df = load_excel_to_dataframe(filepath)
    # 완전히 빈 행 제거
    raw_df = raw_df.dropna(how='all').reset_index(drop=True)
    return build_dashboard_df_from_raw(raw_df)


# ──────────────────────────────────────────────
# ★ 업로드 파일 처리 (고속 버전)
# ──────────────────────────────────────────────

def process_excel_file_with_progress(file_bytes, title_container, progress_bar, status_text):
    try:
        start_time = time.time()

        # ── 1단계: pandas로 고속 읽기 ──
        status_text.markdown("**엑셀 파일을 읽는 중입니다...**")
        progress_bar.progress(5)

        file_stream = io.BytesIO(file_bytes)
        raw_df = load_excel_to_dataframe(file_stream)
        raw_df = raw_df.dropna(how='all').reset_index(drop=True)
        total_rows = len(raw_df)

        elapsed = time.time() - start_time
        title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")
        status_text.markdown(f"**{total_rows:,}개 행 로드 완료. 분류 작업 시작...**")
        progress_bar.progress(30)

        # ── 2단계: 벡터 분류 ──
        dashboard_df = build_dashboard_df_from_raw(raw_df)

        elapsed = time.time() - start_time
        title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")
        status_text.markdown("**분류 완료. 결과 파일을 생성하는 중...**")
        progress_bar.progress(70)

        # ── 3단계: openpyxl로 결과 열만 쓰기 ──
        file_stream.seek(0)
        wb = openpyxl.load_workbook(file_stream)
        ws = wb.active

        BA_COL = 53  # BA
        BB_COL = 54  # BB
        AR_COL = 44
        AS_COL = 45

        ws.cell(row=4, column=BA_COL, value='모델분류')
        ws.cell(row=4, column=BB_COL, value='권역')

        model_values = dashboard_df['모델분류'].tolist()
        region_values = dashboard_df['권역'].tolist()
        ar_dates = dashboard_df['운영계약시작일_cleaned'].tolist()
        as_dates = dashboard_df['운영계약종료일_cleaned'].tolist()

        ar_cleaned_count = 0
        as_cleaned_count = 0

        for i in range(total_rows):
            row_num = i + 5  # 엑셀 5행부터

            ws.cell(row=row_num, column=BA_COL, value=model_values[i])
            ws.cell(row=row_num, column=BB_COL, value=region_values[i])

            ar_d = ar_dates[i]
            if ar_d is not None and not (isinstance(ar_d, float) and np.isnan(ar_d)):
                formatted = format_date_for_excel(ar_d)
                if formatted:
                    ws.cell(row=row_num, column=AR_COL, value=formatted)
                    ws.cell(row=row_num, column=AR_COL).number_format = 'YYYY-MM-DD'
                    ar_cleaned_count += 1

            as_d = as_dates[i]
            if as_d is not None and not (isinstance(as_d, float) and np.isnan(as_d)):
                formatted = format_date_for_excel(as_d)
                if formatted:
                    ws.cell(row=row_num, column=AS_COL, value=formatted)
                    ws.cell(row=row_num, column=AS_COL).number_format = 'YYYY-MM-DD'
                    as_cleaned_count += 1

            # 진행상황 업데이트 (500행마다)
            if (i + 1) % 500 == 0 or i == total_rows - 1:
                elapsed = time.time() - start_time
                pct = 70 + int((i / total_rows) * 25)
                progress_bar.progress(min(pct, 95))
                title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")
                status_text.markdown(
                    f"**엑셀 쓰기 중...** `{i+1:,}/{total_rows:,}` "
                    f"({(i+1)/total_rows*100:.1f}%)"
                )

        progress_bar.progress(95)
        status_text.markdown("**파일 저장 중...**")

        output_stream = io.BytesIO()
        wb.save(output_stream)
        output_stream.seek(0)
        wb.close()

        total_time = time.time() - start_time
        progress_bar.progress(100)

        title_container.markdown(f"### 작업 완료! `총 {format_time(total_time)}`")
        avg_speed = total_rows / total_time if total_time > 0 else 0
        status_text.markdown(
            f"**처리 완료!** `{total_rows:,}개 행` 분류 성공 | "
            f"**평균 속도:** `{avg_speed:.1f}행/초` | "
            f"**날짜 정리:** AR열 `{ar_cleaned_count:,}개`, AS열 `{as_cleaned_count:,}개`"
        )

        return output_stream, None, total_rows, total_time, dashboard_df, ar_cleaned_count, as_cleaned_count

    except Exception as e:
        import traceback
        elapsed = time.time() - start_time
        title_container.markdown(f"### 작업 중단 `{format_time(elapsed)}`")
        status_text.markdown("**오류가 발생했습니다.**")
        error_detail = traceback.format_exc()
        return None, f"파일 처리 중 오류 발생: {str(e)}\n\n상세:\n{error_detail}", 0, 0, None, 0, 0


# ──────────────────────────────────────────────
# 샘플 데이터 / 지도 / 대시보드 (기존과 동일)
# ──────────────────────────────────────────────

def create_sample_data():
    sample_data = {
        '사이트ID': [f'SITE_{i:03d}' for i in range(1, 21)] + [f'SITE_{i:03d}' for i in range(1, 11)],
        '모델분류': [
            '급속스필_100', '급속PNE_100', '신형대', '알박신형', '구형대',
            '급속SK_100', 'F01', '스필_7kW', '급속그린파워_100', 'PNE_7kW',
            '급속코스텔_50', '신형소', '이카플러그', '급속애플망고_200', 'UC01',
            '급속PNE_50', '10kW', 'SK_7kW', '중앙제어_7kW', '3kW',
            '급속스필_100', '신형대', '알박신형', '급속PNE_100', '구형대',
            '급속SK_200', 'F01', '급속그린파워_50', 'PNE_7kW', '급속알박_50'
        ],
        '권역': [
            '수도권북서', '수도권남동', '수도권북동', '충청권', '경상권',
            '수도권남서', '강원권', '수도권북서', '전라권', '수도권남동',
            '수도권북동', '충청권', '수도권남서', '경상권', '강원권',
            '수도권북서', '전라권', '수도권남동', '수도권북동', '충청권',
            '수도권북서', '수도권남동', '경상권', '강원권', '수도권북서',
            '전라권', '수도권남동', '수도권북동', '충청권', '수도권남서'
        ],
        '주소': [
            '서울특별시 강서구 공항대로 지하 396', '경기도 성남시 분당구 판교역로 166', '경기도 의정부시 상금로 36', '대전광역시 유성구 대학로 99', '부산광역시 해운대구 센텀중앙로 79',
            '인천광역시 계양구 안남로 560', '강원도 춘천시 중앙로 1', '서울특별시 마포구 월드컵북로 396', '광주광역시 서구 상무민주로 61', '경기도 용인시 수지구 포은대로 435',
            '서울특별시 강북구 4.19로 100', '충청남도 천안시 동남구 병천면 충절로 1600', '인천광역시 계양구 봉오대로744번길 7', '대구광역시 달서구 달구벌대로 1095', '강원도 원주시 세계로 123',
            '서울특별시 은평구 통일로 684', '전라북도 전주시 완산구 효자로 225', '경기도 수원시 팔달구 중부대로 120', '서울특별시 중랑구 망우로 379', '세종특별자치시 한누리대로 2130',
            '서울특별시 강서구 공항대로 지하 396', '경기도 성남시 분당구 판교역로 166', '울산광역시 남구 삼산로 282', '강원도 강릉시 경강로 2021', '서울특별시 양천구 목동서로 159',
            '전라남도 목포시 평화로 32', '경기도 성남시 중원구 사기막골로 45번길 14', '서울특별시 성북구 정릉로 77', '충청북도 청주시 상당구 상당로 82', '경상남도 창원시 의창구 중앙대로 151'
        ],
        '위도': [
            37.5583, 37.3945, 37.7388, 36.3704, 35.1681,
            37.5376, 37.8813, 37.5665, 35.1595, 37.3217,
            37.6398, 36.5760, 37.5420, 35.8285, 37.3422,
            37.6176, 35.8242, 37.2636, 37.5985, 36.4801,
            37.5583, 37.3945, 35.5384, 37.7519, 37.5172,
            34.7943, 37.4201, 37.5894, 36.6424, 35.2272
        ],
        '경도': [
            126.7944, 127.1116, 127.0467, 127.3622, 129.1303,
            126.7253, 127.7298, 126.9018, 126.8526, 127.1085,
            127.0253, 127.1472, 126.7389, 128.5658, 127.9202,
            126.9227, 127.1530, 127.0286, 127.0927, 127.2890,
            126.7944, 127.1116, 129.3114, 128.8761, 126.8664,
            126.3822, 127.1266, 127.0167, 127.4890, 128.6811
        ],
        '운영계약시작일': [
            date(2022, 1, 15), date(2022, 3, 20), date(2022, 5, 10), date(2022, 7, 5), date(2022, 9, 1),
            date(2023, 1, 10), date(2023, 3, 15), date(2023, 5, 20), date(2023, 7, 8), date(2023, 9, 12),
            date(2024, 1, 5), date(2024, 3, 10), date(2024, 5, 15), date(2024, 7, 20), date(2024, 9, 5),
            date(2025, 1, 8), date(2025, 3, 12), date(2025, 5, 18), date(2025, 7, 22), date(2025, 9, 10),
            date(2022, 2, 14), date(2022, 6, 18), date(2022, 10, 22), date(2023, 2, 15), date(2023, 6, 20),
            date(2024, 2, 10), date(2024, 6, 15), date(2024, 10, 20), date(2025, 2, 12), date(2025, 6, 18)
        ],
        '운영계약종료일': [
            date(2028, 1, 14), date(2028, 3, 19), date(2028, 5, 9), date(2028, 7, 4), date(2028, 8, 31),
            date(2029, 1, 9), date(2029, 3, 14), date(2029, 5, 19), date(2029, 7, 7), date(2029, 9, 11),
            date(2030, 1, 4), date(2030, 3, 9), date(2030, 5, 14), date(2030, 7, 19), date(2030, 9, 4),
            date(2031, 1, 7), date(2031, 3, 11), date(2031, 5, 17), date(2031, 7, 21), date(2031, 9, 9),
            date(2028, 2, 13), date(2028, 6, 17), date(2028, 10, 21), date(2029, 2, 14), date(2029, 6, 19),
            date(2030, 2, 9), date(2030, 6, 14), date(2030, 10, 19), date(2031, 2, 11), date(2031, 6, 17)
        ]
    }

    df = pd.DataFrame(sample_data)
    df['운영계약시작일_parsed'] = df['운영계약시작일']
    df['운영계약종료일_parsed'] = df['운영계약종료일']
    df['운영계약시작일_cleaned'] = df['운영계약시작일']
    df['운영계약종료일_cleaned'] = df['운영계약종료일']
    df['행번호'] = range(5, 5 + len(df))

    return df


def create_charger_map(filtered_df):
    map_data = filtered_df.dropna(subset=['위도', '경도']).copy()

    if len(map_data) == 0:
        return None, "좌표 데이터가 없습니다."

    if '사이트ID' not in map_data.columns or map_data['사이트ID'].isna().all():
        map_data['사이트ID'] = [f'SITE_{i:04d}' for i in range(len(map_data))]

    grouped = map_data.groupby('사이트ID').agg({
        '위도': 'first', '경도': 'first', '주소': 'first',
        '권역': 'first', '모델분류': list,
        '운영계약시작일': 'first', '운영계약종료일': 'first'
    }).reset_index()

    site_total = map_data.groupby('사이트ID').size().reset_index(name='총충전기수')
    fast_mask = map_data['모델분류'].str.contains('급속', na=False)
    fast_counts = map_data[fast_mask].groupby('사이트ID').size().reset_index(name='급속충전기수')

    grouped = grouped.merge(site_total, on='사이트ID', how='left')
    grouped = grouped.merge(fast_counts, on='사이트ID', how='left')
    grouped['급속충전기수'] = grouped['급속충전기수'].fillna(0).astype(int)
    grouped['완속충전기수'] = grouped['총충전기수'] - grouped['급속충전기수']

    center_lat = grouped['위도'].mean()
    center_lon = grouped['경도'].mean()

    m = folium.Map(location=[center_lat, center_lon], zoom_start=8, tiles='OpenStreetMap')

    marker_cluster = MarkerCluster(
        name="충전소 클러스터", overlay=True, control=True,
        options={"disableClusteringAtZoom": 15, "maxClusterRadius": 50}
    ).add_to(m)

    region_colors = {
        '수도권북서': 'blue', '수도권북동': 'green', '수도권남동': 'red', '수도권남서': 'purple',
        '수도권기타': 'cadetblue', '인천기타': 'orange', '강원권': 'lightblue', '충청권': 'lightgreen',
        '경상권': 'pink', '전라권': 'lightgray', '기타': 'gray'
    }

    for _, row in grouped.iterrows():
        site_id = row['사이트ID']
        address = row['주소'] if row['주소'] else ''
        total_chargers = row['총충전기수']
        fast_chargers = row['급속충전기수']
        slow_chargers = row['완속충전기수']
        models = row['모델분류']
        region = row['권역'] if row['권역'] else '기타'

        encoded_address = quote(f"{address} 전기차")
        naver_map_url = f"https://map.naver.com/p/search/{encoded_address}"

        icon_name = 'flash' if fast_chargers > 0 else 'plug'
        color = region_colors.get(region, 'gray')
        tooltip_text = f"{site_id} | {total_chargers}기 | {region}"
        models_text = ', '.join(set([str(m_item) for m_item in models]))

        popup_html = f"""
        <div style="width: 320px; font-family: 'Malgun Gothic', Arial, sans-serif;">
            <h4 style="margin: 0 0 12px 0; color: #333; border-bottom: 3px solid {color}; padding-bottom: 8px;">{site_id}</h4>
            <table style="width: 100%; font-size: 13px; border-collapse: collapse; margin-bottom: 12px;">
                <tr style="background-color: #f8f9fa;"><td style="padding: 6px; font-weight: bold; width: 80px;">총 충전기</td><td style="padding: 6px; color: #0066cc; font-weight: bold;">{total_chargers}대</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">급속/완속</td><td style="padding: 6px;">{fast_chargers}대 / {slow_chargers}대</td></tr>
                <tr style="background-color: #f8f9fa;"><td style="padding: 6px; font-weight: bold;">권역</td><td style="padding: 6px;">{region}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">주소</td><td style="padding: 6px; font-size: 11px;">{address}</td></tr>
                <tr style="background-color: #f8f9fa;"><td style="padding: 6px; font-weight: bold;">모델</td><td style="padding: 6px; font-size: 11px;">{models_text}</td></tr>
            </table>
            <div style="text-align: center; margin-top: 15px;">
                <a href="{naver_map_url}" target="_blank" style="display: inline-block; padding: 10px 20px; background: linear-gradient(135deg, #03C75A, #029B47); color: white; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 13px;">네이버 지도에서 보기</a>
            </div>
        </div>
        """

        folium.Marker(
            location=[row['위도'], row['경도']],
            popup=folium.Popup(popup_html, max_width=350),
            tooltip=tooltip_text,
            icon=folium.Icon(color=color, icon=icon_name, prefix='fa')
        ).add_to(marker_cluster)

    legend_html = f'''
    <div style="position: fixed; bottom: 50px; right: 50px; border: 2px solid grey; z-index: 9999; background-color: white; padding: 15px; font-size: 12px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.3);">
        <p style="margin: 0 0 10px 0; font-weight: bold; font-size: 14px;">지도 범례</p>
        <p style="margin: 5px 0;"><i class="fa fa-flash" style="color: red;"></i> 급속 포함</p>
        <p style="margin: 5px 0;"><i class="fa fa-plug" style="color: blue;"></i> 완속만</p>
        <hr style="margin: 8px 0;">
        <p style="margin: 5px 0; font-weight: bold;">총 사이트: {len(grouped):,}개</p>
        <p style="margin: 5px 0; font-weight: bold;">총 충전기: {len(map_data):,}대</p>
        <hr style="margin: 8px 0;">'''

    for region, color in region_colors.items():
        if region in grouped['권역'].values:
            site_count = len(grouped[grouped['권역'] == region])
            legend_html += f'<p style="margin: 1px 0; font-size: 10px;"><span style="color: {color}; font-size: 14px;">●</span> {region} ({site_count}개)</p>'

    legend_html += '</div>'
    m.get_root().html.add_child(folium.Element(legend_html))

    return m, None


def show_dashboard(df):
    st.markdown("## 충전기 운영 현황 대시보드")
    st.markdown("### 운영계약 기간 필터")

    df_dates = df.copy()
    df_dates['운영계약시작일_parsed'] = pd.to_datetime(df_dates['운영계약시작일_parsed'], errors='coerce')
    df_dates['운영계약종료일_parsed'] = pd.to_datetime(df_dates['운영계약종료일_parsed'], errors='coerce')

    valid_dates = df_dates.dropna(subset=['운영계약시작일_parsed', '운영계약종료일_parsed'])

    if len(valid_dates) > 0:
        min_date = valid_dates['운영계약시작일_parsed'].min().date()
        max_date = valid_dates['운영계약종료일_parsed'].max().date()

        default_start = max(min_date, date(2022, 1, 1))
        default_end = min(max_date, date(2028, 1, 1))
        if default_start > default_end:
            default_start = min_date
            default_end = max_date

        col1, col2, col3 = st.columns([2, 2, 1])
        with col1:
            start_date = st.date_input("계약 시작일 (이후)", value=default_start, min_value=min_date, max_value=max_date)
        with col2:
            end_date = st.date_input("계약 종료일 (이전)", value=default_end, min_value=min_date, max_value=max_date)
        with col3:
            st.markdown("<br>", unsafe_allow_html=True)
            st.button("필터 적용", type="primary", use_container_width=True)

        start_ts = pd.Timestamp(start_date)
        end_ts = pd.Timestamp(end_date)

        mask = (
            (df_dates['운영계약시작일_parsed'] < end_ts) &
            (df_dates['운영계약종료일_parsed'] >= start_ts) &
            df_dates['운영계약시작일_parsed'].notna() &
            df_dates['운영계약종료일_parsed'].notna()
        )
        filtered_df = df[mask].copy()

        st.info(f"**선택 기간:** {start_date} ~ {end_date} | **해당 기간 충전기:** {len(filtered_df):,}대 (전체 {len(df):,}대 중 {len(filtered_df)/len(df)*100:.1f}%)")

        if len(filtered_df) == 0:
            st.warning("선택한 기간에 해당하는 데이터가 없습니다.")
            return

        st.markdown("---")
        st.markdown("### 주요 지표")

        total_chargers = len(filtered_df)
        unique_sites = filtered_df['사이트ID'].nunique()
        region_count = filtered_df['권역'].nunique()
        model_count = filtered_df['모델분류'].nunique()
        fast_chargers = len(filtered_df[filtered_df['모델분류'].str.contains('급속', na=False)])
        fast_ratio = (fast_chargers / total_chargers * 100) if total_chargers > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("총 충전기 수", f"{total_chargers:,}대")
        c2.metric("사이트 수", f"{unique_sites:,}개")
        c3.metric("권역 수", f"{region_count}개")
        c4.metric("급속 충전기", f"{fast_chargers:,}대", f"{fast_ratio:.1f}%")

        st.markdown("---")
        st.markdown("### 충전기 위치 지도 (사이트ID 기준)")

        has_coordinates = '위도' in filtered_df.columns and '경도' in filtered_df.columns
        if has_coordinates:
            valid_coords = filtered_df.dropna(subset=['위도', '경도'])
            coord_count = len(valid_coords)

            if coord_count > 0:
                unique_sites_map = valid_coords['사이트ID'].nunique()
                st.success(f"{unique_sites_map:,}개 사이트, {coord_count:,}개 충전기의 좌표 데이터가 있습니다.")

                charger_map, error = create_charger_map(filtered_df)
                if error:
                    st.error(error)
                else:
                    st_folium(charger_map, width=1400, height=700)
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("지도 표시 사이트", f"{unique_sites_map:,}개")
                    c2.metric("지도 표시 충전기", f"{coord_count:,}대")
                    c3.metric("좌표 보유율", f"{coord_count/len(filtered_df)*100:.1f}%")
                    c4.metric("좌표 누락", f"{len(filtered_df)-coord_count:,}대")

                    with st.expander("지도 사용 방법"):
                        st.markdown("""
                        **지도 조작:** 확대/축소(마우스 휠), 이동(드래그), 클러스터(숫자 원 클릭)
                        **마커:** 오버(간단 정보), 클릭(상세 + 네이버 지도 링크)
                        **아이콘:** 번개=급속 포함, 플러그=완속만, 색상=권역별
                        """)
            else:
                st.warning("좌표 데이터가 없습니다.")
        else:
            st.warning("좌표 데이터가 없습니다.")

        st.markdown("---")
        st.markdown("### 모델별 충전기 현황")

        col1, col2 = st.columns([3, 2])
        with col1:
            model_counts = filtered_df['모델분류'].value_counts().reset_index()
            model_counts.columns = ['모델분류', '수량']
            fig = px.bar(model_counts.head(15), x='수량', y='모델분류', orientation='h',
                         title='모델별 수량 (상위 15개)', color='수량', color_continuous_scale='Blues', text='수량')
            fig.update_layout(height=500, showlegend=False)
            fig.update_traces(texttemplate='%{text}', textposition='outside')
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            st.markdown("#### 모델별 수량 상세")
            model_counts['비율'] = (model_counts['수량'] / model_counts['수량'].sum() * 100).round(1).astype(str) + '%'
            st.dataframe(model_counts[['모델분류', '수량', '비율']], use_container_width=True, hide_index=True, height=450)

        st.markdown("---")
        st.markdown("### 권역별 충전기 현황")

        col1, col2 = st.columns([2, 3])
        with col1:
            region_counts = filtered_df['권역'].value_counts().reset_index()
            region_counts.columns = ['권역', '수량']
            fig = px.pie(region_counts, values='수량', names='권역', title='권역별 비율', hole=0.4)
            fig.update_traces(textposition='inside', textinfo='percent+label')
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            fig = px.bar(region_counts, x='권역', y='수량', title='권역별 수량', color='수량',
                         color_continuous_scale='Greens', text='수량')
            fig.update_layout(height=400, showlegend=False)
            fig.update_traces(texttemplate='%{text}', textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown("### 권역별 모델 분포 히트맵")

        crosstab = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'])
        top_models = filtered_df['모델분류'].value_counts().head(12).index
        available = [m for m in top_models if m in crosstab.columns]

        if available:
            ct = crosstab[available]
            fig = px.imshow(ct.T, labels=dict(x="권역", y="모델분류", color="수량"),
                            color_continuous_scale='RdYlGn', aspect="auto",
                            title='권역별 주요 모델 분포 (상위 12개)', text_auto=True)
            fig.update_layout(height=500)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.markdown("### 권역별 × 모델별 상세 현황")
        pivot_wide = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'], margins=True)
        st.dataframe(pivot_wide, use_container_width=True, height=400)

        st.markdown("---")
        st.markdown("### 데이터 다운로드")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.download_button("권역×모델 현황표 CSV", pivot_wide.to_csv(encoding='utf-8-sig'),
                               f"권역모델현황_{start_date}_{end_date}.csv", "text/csv", use_container_width=True)
        with c2:
            st.download_button("필터링 전체 데이터 CSV", filtered_df.to_csv(index=False, encoding='utf-8-sig'),
                               f"필터링데이터_{start_date}_{end_date}.csv", "text/csv", use_container_width=True)
        with c3:
            lines = [
                f"충전기 현황 요약", "", f"기간: {start_date} ~ {end_date}",
                f"총: {total_chargers:,}대, 사이트: {unique_sites:,}개",
                f"급속: {fast_chargers:,}대 ({fast_ratio:.1f}%)", "", "상위 5 모델:"
            ]
            for i, row in model_counts.head(5).iterrows():
                lines.append(f"  {i+1}. {row['모델분류']}: {row['수량']:,}대 ({row['비율']})")
            st.download_button("요약 리포트 TXT", "\n".join(lines),
                               f"요약리포트_{start_date}_{end_date}.txt", "text/plain", use_container_width=True)

        st.markdown("---")
        st.markdown("### 데이터 품질 체크")

        unknown = filtered_df[filtered_df['권역'].isin(['수도권기타', '인천기타', '기타'])]
        c1, c2 = st.columns(2)
        normal = len(filtered_df) - len(unknown)
        c1.metric("정상 분류", f"{normal:,}대", f"{normal/len(filtered_df)*100:.1f}%")
        c2.metric("미분류/불명확", f"{len(unknown):,}대", f"{len(unknown)/len(filtered_df)*100:.1f}%")

        if len(unknown) > 0:
            st.warning(f"{len(unknown):,}개 미분류")
            with st.expander("미분류 상세"):
                stats = unknown['권역'].value_counts()
                st.dataframe(pd.DataFrame({'권역': stats.index, '수량': stats.values}), hide_index=True)
                cols = [c for c in ['주소', '권역', '모델분류', '사이트ID'] if c in unknown.columns]
                st.dataframe(unknown[cols].head(10), hide_index=True)
        else:
            st.success("모든 주소가 정확하게 분류되었습니다!")
    else:
        st.warning("유효한 운영계약 날짜 데이터가 없습니다.")


def show_classification_info():
    st.markdown("### 모델분류 기준표")
    t1, t2, t3, t4 = st.tabs(["급속 충전기", "완속 충전기", "권역 분류", "참조 정보"])

    with t1:
        st.markdown("#### 급속 충전기 분류 기준")
        st.info("**전제 조건:** AH열 = '급속'")
        st.dataframe({
            "모델분류명": ["급속스필_100", "급속스필_50", "급속PNE_100", "급속PNE_50", "급속PNE_200", "급속PNE_150",
                       "급속애플망고_200", "급속SK_100", "급속SK_200", "급속코스텔_50", "급속중앙제어_50",
                       "급속그린파워_100", "급속그린파워_50", "급속알박_50", "급속"],
            "AG열 코드": ["S0F1", "S0F5", "EVQ-(AJ=100)", "EVQ-/EV1-(AJ=50)", "MAXE", "DP15", "A01-/AD1-",
                       "Q081/Q101/Q010", "Q071/Q102", "1Y25/1Y24", "1911", "1900", "19C0", "QC50", "기타"]
        }, hide_index=True, use_container_width=True)

    with t2:
        st.markdown("#### 완속 충전기 분류 기준")
        st.info("**전제 조건:** AH열 ≠ '급속'")
        st.dataframe({
            "우선순위": list(range(1, 20)),
            "모델분류명": ["알박구형", "알박신형", "10kW", "신형대(EVL-1C-22CQ)", "구형대(EVL-1C)",
                       "신형대(EVL+1107)", "구형대(EVL기본)", "신형대(SBDA)", "신형소", "F01/PC01(SBPA)",
                       "UC01", "스필_7kW", "이카플러그", "중앙제어_7kW", "SK_7kW", "3kW", "PNE_7kW",
                       "F01/PC01(SBOA)", "기타"]
        }, hide_index=True, use_container_width=True)

    with t3:
        st.markdown("#### 권역 분류 기준 (H열 주소)")
        st.dataframe({
            "권역명": ["수도권북서", "수도권북동", "수도권남동", "수도권남서", "수도권남서(인천)",
                     "강원권", "충청권", "경상권", "전라권", "기타"],
            "주요 지역": [
                "고양,부천,김포,파주,은평,마포,서대문,양천,강서",
                "도봉,노원,중랑,강북,성북,동대문,의정부,남양주,구리",
                "강남,서초,송파,강동,성남,용인,하남,수원,평택",
                "구로,금천,영등포,동작,관악,의왕,광명,안산,안양",
                "인천-계양,남동,부평,연수,미추홀",
                "강원도 전역", "충청,세종,대전", "경상,부산,대구,울산",
                "전라,광주", "위 조건 해당 없음"]
        }, hide_index=True, use_container_width=True)

    with t4:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("""
            **입력 열:** A(사이트ID), H(주소), AD(모델명), AG(모델코드), AH(급속/완속), AJ(용량), AM(경도), AN(위도), AR(시작일), AS(종료일)
            """)
        with c2:
            st.markdown("""
            **출력 열:** BA(모델분류), BB(권역), AR/AS(날짜 정리)
            """)


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────

def main():
    if 'processed_df' not in st.session_state:
        if os.path.exists(DEFAULT_DATA_PATH):
            try:
                with st.spinner("default_data.xlsx 로드 중..."):
                    st.session_state.processed_df = load_default_excel(DEFAULT_DATA_PATH)
                    st.session_state.is_sample_data = False
                    st.session_state.default_file_loaded = True
            except Exception as e:
                st.warning(f"default_data.xlsx 로드 실패: {e}")
                st.session_state.processed_df = create_sample_data()
                st.session_state.is_sample_data = True
                st.session_state.default_file_loaded = False
        else:
            st.session_state.processed_df = create_sample_data()
            st.session_state.is_sample_data = True
            st.session_state.default_file_loaded = False

    if 'processed_file' not in st.session_state:
        st.session_state.processed_file = None

    st.title("⚡ 충전기 모델분류 & 운영현황 대시보드")
    st.markdown("""
    엑셀 파일을 업로드하면 **BA열**에 "모델분류", **BB열**에 "권역"을 자동으로 추가하고,
    **AR열/AS열** 날짜를 정리하여 **AM열(경도)/AN열(위도)** 기반 지도에서 충전기 위치를 확인할 수 있습니다.
    """)

    tab1, tab2 = st.tabs(["파일 업로드 & 분류", "운영현황 대시보드"])

    with tab1:
        if st.session_state.get('default_file_loaded'):
            row_count = len(st.session_state.processed_df)
            st.success(f"**default_data.xlsx** 자동 로드 완료 ({row_count:,}개 행)")
        elif st.session_state.get('is_sample_data'):
            st.info("**샘플 데이터 사용 중** — 대시보드 탭에서 체험 가능")

        uploaded_file = st.file_uploader("엑셀 파일 선택", type=['xlsx', 'xls'],
                                         help="AM열(경도), AN열(위도), 사이트ID 포함 권장. 최대 200MB")

        if uploaded_file is not None:
            c1, c2 = st.columns([3, 1])
            with c1:
                st.info(f"**{uploaded_file.name}**")
            with c2:
                sz = uploaded_file.size / (1024 * 1024)
                st.metric("파일 크기", f"{sz:.1f} MB" if sz >= 1 else f"{uploaded_file.size/1024:.1f} KB")

            if st.button("모델분류 시작", type="primary", use_container_width=True):
                title_container = st.empty()
                title_container.markdown("### 작업 진행 상황 `0.0초`")
                progress_bar = st.progress(0)
                status_text = st.empty()
                st.markdown("---")

                file_bytes = uploaded_file.read()
                result = process_excel_file_with_progress(file_bytes, title_container, progress_bar, status_text)
                processed_file, error, processed_count, total_time, result_df, ar_cleaned, as_cleaned = result

                if error:
                    st.error(error)
                else:
                    st.session_state.processed_df = result_df
                    st.session_state.processed_file = processed_file
                    st.session_state.is_sample_data = False
                    st.session_state.default_file_loaded = False

                    st.success(f"**{processed_count:,}개 행** 분류 완료! ({format_time(total_time)})")
                    st.info(f"날짜 정리: AR열 `{ar_cleaned:,}개`, AS열 `{as_cleaned:,}개`")

                    timestamp = get_korea_time().strftime("%Y%m%d_%H%M%S")
                    c1, c2, c3 = st.columns([1, 2, 1])
                    with c2:
                        st.download_button("결과 파일 다운로드", processed_file.getvalue(),
                                           f"모델분류_결과_{timestamp}.xlsx",
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           use_container_width=True, type="primary")

                    st.info("**'운영현황 대시보드'** 탭으로 이동하세요!")

                    with st.expander("처리 결과 상세"):
                        c1, c2, c3, c4, c5 = st.columns(5)
                        c1.metric("처리 행 수", f"{processed_count:,}개")
                        c2.metric("소요 시간", format_time(total_time))
                        c3.metric("처리 속도", f"{processed_count/total_time:.1f}행/초" if total_time > 0 else "-")
                        c4.metric("AR열 정리", f"{ar_cleaned:,}개")
                        c5.metric("AS열 정리", f"{as_cleaned:,}개")

        with st.expander("분류 기준 정보"):
            show_classification_info()

    with tab2:
        if st.session_state.processed_df is not None:
            if st.session_state.get('default_file_loaded'):
                st.info("**default_data.xlsx** 데이터 표시 중")
            elif st.session_state.get('is_sample_data'):
                st.warning("**샘플 데이터 사용 중** — 파일 업로드로 실제 데이터 분석 가능")
            show_dashboard(st.session_state.processed_df)
        else:
            st.info("먼저 파일을 업로드해주세요.")


if __name__ == "__main__":
    main()
