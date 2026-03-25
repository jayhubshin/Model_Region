import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import io
import os
import time
import re
from datetime import datetime, date
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pytz
import folium
from streamlit_folium import st_folium
from folium.plugins import MarkerCluster
from urllib.parse import quote

# 페이지 설정
st.set_page_config(
    page_title="충전기 모델분류 자동화",
    page_icon="⚡",
    layout="wide"
)

# ============================================================
# 기본 데이터 파일 경로 설정
# 앱과 같은 폴더 또는 하위 폴더에 파일을 배치하세요
# ============================================================
DEFAULT_DATA_FILES = [
    "default_data.xlsx",
    "data/default_data.xlsx",
    "초기데이터.xlsx",
    "data/초기데이터.xlsx",
    "default_data.csv",
    "data/default_data.csv",
]


def get_korea_time():
    korea_tz = pytz.timezone('Asia/Seoul')
    return datetime.now(korea_tz)


def get_safe_value(row_data, col_letter):
    val = row_data.get(col_letter)
    if val is None:
        return ""
    return str(val).strip()


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
    if date_value is None:
        return None
    if isinstance(date_value, datetime):
        return date_value.date()
    if isinstance(date_value, date):
        return date_value
    if isinstance(date_value, str):
        date_str = date_value.strip()
        date_str = re.sub(r'\s+\d{1,2}:\d{2}:\d{2}$', '', date_str)
        if not date_str:
            return None
        date_formats = [
            '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d',
            '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
            '%d-%m-%Y',
        ]
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
    try:
        if isinstance(date_value, (int, float)):
            from datetime import timedelta
            excel_epoch = datetime(1899, 12, 30)
            return (excel_epoch + timedelta(days=float(date_value))).date()
    except Exception:
        pass
    return None


def format_date_for_excel(date_obj):
    if date_obj is None:
        return None
    if isinstance(date_obj, date):
        return date_obj.strftime('%Y-%m-%d')
    return None


def classify_region(address):
    if not address:
        return "기타"
    address_clean = str(address).strip()

    if re.search(r"인천", address_clean):
        if re.search(r"계양|남동|동구|미추홀|부평|연수|서구|중구|강화", address_clean):
            return "수도권남서"
        else:
            return "인천기타"
    elif re.search(r"서울|경기", address_clean):
        if re.search(r"고양|부천|김포|파주|은평구|마포구|서대문구|양천구|강서구|용산구|중구|종로구", address_clean):
            return "수도권북서"
        elif re.search(r"도봉구|노원구|중랑구|강북구|성북구|동대문구|성동구|광진구|의정부|남양주|구리|양주|포천|동두천|가평|연천", address_clean):
            return "수도권북동"
        elif re.search(r"강남구|서초구|송파구|강동구|성남|용인|하남|광주|안성|수원|평택|오산|이천|여주|양평", address_clean):
            return "수도권남동"
        elif re.search(r"구로구|금천구|영등포구|동작구|관악구|의왕|광명|군포|과천|시흥|안산|안양|화성", address_clean):
            return "수도권남서"
        else:
            return "수도권기타"
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
    AD = get_safe_value(row_data, 'AD')
    AG = get_safe_value(row_data, 'AG')
    AH = get_safe_value(row_data, 'AH')
    AJ = get_safe_value(row_data, 'AJ')

    if AH == "급속":
        ag4 = AG[:4] if len(AG) >= 4 else AG
        if ag4 == "S0F1": return "급속스필_100"
        elif ag4 == "S0F5": return "급속스필_50"
        elif ag4 == "EVQ-" and AJ == "100": return "급속PNE_100"
        elif (ag4 == "EVQ-" or ag4 == "EV1-") and AJ == "50": return "급속PNE_50"
        elif ag4 == "MAXE": return "급속PNE_200"
        elif ag4 == "DP15": return "급속PNE_150"
        elif ag4 in ["A01-", "AD1-"]: return "급속애플망고_200"
        elif ag4 in ["Q081", "Q101", "Q010"]: return "급속SK_100"
        elif ag4 in ["Q071", "Q102"]: return "급속SK_200"
        elif ag4 in ["1Y25", "1Y24"]: return "급속코스텔_50"
        elif ag4 == "1911": return "급속중앙제어_50"
        elif ag4 == "1900": return "급속그린파워_100"
        elif ag4 == "19C0": return "급속그린파워_50"
        elif ag4 == "QC50": return "급속알박_50"
        else: return "급속"

    ag3 = AG[:3] if len(AG) >= 3 else AG
    ag4 = AG[:4] if len(AG) >= 4 else AG
    ag6 = AG[:6] if len(AG) >= 6 else AG
    ag11 = AG[:11] if len(AG) >= 11 else AG

    if ag4 == "NC07": return "알박구형"
    elif ag4 in ["23NA", "22NA", "24NA", "25NA"]: return "알박신형"
    elif "3J10" in AD: return "10kW"
    elif ag11 == "EVL-1C-22CQ": return "신형대"
    elif ag6 == "EVL-1C": return "구형대"
    elif ag4 == "EVL-" and "1107" in AD: return "신형대"
    elif ag4 == "EVL-": return "구형대"
    elif ag4 == "SBDA": return "신형대"
    elif ag4 == "SBAA": return "신형소"
    elif ag4 == "SBPA": return "F01" if "F01" in AD else "PC01"
    elif ag4 == "SBUA": return "UC01"
    elif ag4 == "SVI0": return "스필_7kW"
    elif ag3 == "E0C" or "CP" in AD: return "이카플러그"
    elif ag4 in ["1907", "1912"]: return "중앙제어_7kW"
    elif ag4 == "SC-P": return "SK_7kW"
    elif ag4 == "SANA": return "3kW"
    elif ag4 in ["EVS-", "007S"]: return "PNE_7kW"
    elif ag4 == "SBOA": return "F01" if "F01" in AD else "PC01"
    else: return "기타"


# ============================================================
# 파일 탐색 및 로드 함수
# ============================================================

def find_default_data_file():
    """DEFAULT_DATA_FILES 경로 목록에서 존재하는 첫 번째 파일 반환"""
    for filepath in DEFAULT_DATA_FILES:
        if os.path.exists(filepath):
            return filepath
    return None


def load_from_csv(filepath):
    """CSV 파일에서 DataFrame 로드"""
    try:
        df = pd.read_csv(filepath, encoding='utf-8-sig')
    except UnicodeDecodeError:
        df = pd.read_csv(filepath, encoding='cp949')

    # 영문 컬럼명 매핑
    col_map = {
        'site_id': '사이트ID', 'model': '모델분류', 'region': '권역',
        'address': '주소', 'lat': '위도', 'lng': '경도',
        'longitude': '경도', 'latitude': '위도',
        'contract_start': '운영계약시작일', 'contract_end': '운영계약종료일',
    }
    df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

    required = ['사이트ID', '모델분류', '권역', '주소']
    missing = [c for c in required if c not in df.columns]
    if missing:
        return None, f"CSV에 필수 컬럼이 없습니다: {missing}"

    for date_col in ['운영계약시작일', '운영계약종료일']:
        if date_col in df.columns:
            df[date_col + '_parsed'] = df[date_col].apply(clean_and_parse_date)
            df[date_col + '_cleaned'] = df[date_col + '_parsed']

    for num_col in ['위도', '경도']:
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors='coerce')

    if '행번호' not in df.columns:
        df['행번호'] = range(5, 5 + len(df))

    return df, None


def load_from_excel(filepath):
    """엑셀 파일에서 분류를 수행하고 DataFrame 생성"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    max_row = ws.max_row
    if max_row < 5:
        return None, "데이터가 없습니다 (5행 미만)"

    max_col = max(ws.max_column, 54)
    dashboard_data = []

    for row_num in range(5, max_row + 1):
        row_data = {}
        for col_num in range(1, max_col + 1):
            col_letter = get_column_letter(col_num)
            row_data[col_letter] = ws.cell(row=row_num, column=col_num).value

        classification_result = classify_model(row_data, row_num)
        address = get_safe_value(row_data, 'H')
        region_result = classify_region(address)

        ar_cleaned = clean_and_parse_date(row_data.get('AR'))
        as_cleaned = clean_and_parse_date(row_data.get('AS'))

        site_id = get_safe_value(row_data, 'A')

        try:
            lon = float(row_data.get('AM')) if row_data.get('AM') else None
        except (ValueError, TypeError):
            lon = None
        try:
            lat = float(row_data.get('AN')) if row_data.get('AN') else None
        except (ValueError, TypeError):
            lat = None

        dashboard_data.append({
            '사이트ID': site_id if site_id else f'AUTO_{row_num}',
            '모델분류': classification_result,
            '권역': region_result,
            '주소': address,
            '위도': lat,
            '경도': lon,
            '운영계약시작일': row_data.get('AR'),
            '운영계약종료일': row_data.get('AS'),
            '운영계약시작일_cleaned': ar_cleaned,
            '운영계약종료일_cleaned': as_cleaned,
            '행번호': row_num
        })

    df = pd.DataFrame(dashboard_data)
    df['운영계약시작일_parsed'] = df['운영계약시작일_cleaned']
    df['운영계약종료일_parsed'] = df['운영계약종료일_cleaned']
    return df, None


def load_initial_data_from_file(filepath):
    """파일 확장자에 따라 적절한 로더 호출"""
    try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext == '.csv':
            return load_from_csv(filepath)
        elif ext in ['.xlsx', '.xls']:
            return load_from_excel(filepath)
        else:
            return None, f"지원하지 않는 형식: {ext}"
    except Exception as e:
        return None, f"파일 로드 실패: {str(e)}"


def load_initial_data_from_bytes(file_bytes, filename):
    """업로드된 파일 바이트에서 DataFrame 로드 (기본 데이터 설정용)"""
    try:
        ext = os.path.splitext(filename)[1].lower()
        if ext == '.csv':
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8-sig')
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='cp949')

            col_map = {
                'site_id': '사이트ID', 'model': '모델분류', 'region': '권역',
                'address': '주소', 'lat': '위도', 'lng': '경도',
                'longitude': '경도', 'latitude': '위도',
            }
            df = df.rename(columns={k: v for k, v in col_map.items() if k in df.columns})

            required = ['사이트ID', '모델분류', '권역', '주소']
            missing = [c for c in required if c not in df.columns]
            if missing:
                return None, f"CSV에 필수 컬럼이 없습니다: {missing}"

            for date_col in ['운영계약시작일', '운영계약종료일']:
                if date_col in df.columns:
                    df[date_col + '_parsed'] = df[date_col].apply(clean_and_parse_date)
                    df[date_col + '_cleaned'] = df[date_col + '_parsed']

            for num_col in ['위도', '경도']:
                if num_col in df.columns:
                    df[num_col] = pd.to_numeric(df[num_col], errors='coerce')

            if '행번호' not in df.columns:
                df['행번호'] = range(5, 5 + len(df))

            return df, None

        elif ext in ['.xlsx', '.xls']:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb.active
            max_row = ws.max_row
            if max_row < 5:
                return None, "데이터가 없습니다 (5행 미만)"

            max_col = max(ws.max_column, 54)
            dashboard_data = []

            for row_num in range(5, max_row + 1):
                row_data = {}
                for col_num in range(1, max_col + 1):
                    col_letter = get_column_letter(col_num)
                    row_data[col_letter] = ws.cell(row=row_num, column=col_num).value

                classification_result = classify_model(row_data, row_num)
                address = get_safe_value(row_data, 'H')
                region_result = classify_region(address)
                ar_cleaned = clean_and_parse_date(row_data.get('AR'))
                as_cleaned = clean_and_parse_date(row_data.get('AS'))
                site_id = get_safe_value(row_data, 'A')

                try:
                    lon = float(row_data.get('AM')) if row_data.get('AM') else None
                except (ValueError, TypeError):
                    lon = None
                try:
                    lat = float(row_data.get('AN')) if row_data.get('AN') else None
                except (ValueError, TypeError):
                    lat = None

                dashboard_data.append({
                    '사이트ID': site_id if site_id else f'AUTO_{row_num}',
                    '모델분류': classification_result,
                    '권역': region_result,
                    '주소': address,
                    '위도': lat,
                    '경도': lon,
                    '운영계약시작일': row_data.get('AR'),
                    '운영계약종료일': row_data.get('AS'),
                    '운영계약시작일_cleaned': ar_cleaned,
                    '운영계약종료일_cleaned': as_cleaned,
                    '행번호': row_num
                })

            df = pd.DataFrame(dashboard_data)
            df['운영계약시작일_parsed'] = df['운영계약시작일_cleaned']
            df['운영계약종료일_parsed'] = df['운영계약종료일_cleaned']
            return df, None
        else:
            return None, f"지원하지 않는 형식: {ext}"
    except Exception as e:
        return None, f"파일 로드 실패: {str(e)}"


# ============================================================
# 샘플 데이터 (파일이 없을 때 폴백)
# ============================================================

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
            '서울특별시 강서구 공항대로 396', '경기도 성남시 분당구 판교역로 166',
            '경기도 의정부시 상금로 36', '대전광역시 유성구 대학로 99',
            '부산광역시 해운대구 센텀중앙로 79', '인천광역시 계양구 안남로 560',
            '강원도 춘천시 중앙로 1', '서울특별시 마포구 월드컵북로 396',
            '광주광역시 서구 상무민주로 61', '경기도 용인시 수지구 포은대로 435',
            '서울특별시 강북구 4.19로 100', '충청남도 천안시 동남구 충절로 1600',
            '인천광역시 계양구 봉오대로 7', '대구광역시 달서구 달구벌대로 1095',
            '강원도 원주시 세계로 123', '서울특별시 은평구 통일로 684',
            '전라북도 전주시 완산구 효자로 225', '경기도 수원시 팔달구 중부대로 120',
            '서울특별시 중랑구 망우로 379', '세종특별자치시 한누리대로 2130',
            '서울특별시 강서구 공항대로 396', '경기도 성남시 분당구 판교역로 166',
            '울산광역시 남구 삼산로 282', '강원도 강릉시 경강로 2021',
            '서울특별시 양천구 목동서로 159', '전라남도 목포시 평화로 32',
            '경기도 성남시 중원구 사기막골로 14', '서울특별시 성북구 정릉로 77',
            '충청북도 청주시 상당구 상당로 82', '경상남도 창원시 의창구 중앙대로 151'
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
            date(2022, 1, 15), date(2022, 3, 20), date(2022, 5, 10),
            date(2022, 7, 5), date(2022, 9, 1), date(2023, 1, 10),
            date(2023, 3, 15), date(2023, 5, 20), date(2023, 7, 8),
            date(2023, 9, 12), date(2024, 1, 5), date(2024, 3, 10),
            date(2024, 5, 15), date(2024, 7, 20), date(2024, 9, 5),
            date(2025, 1, 8), date(2025, 3, 12), date(2025, 5, 18),
            date(2025, 7, 22), date(2025, 9, 10), date(2022, 2, 14),
            date(2022, 6, 18), date(2022, 10, 22), date(2023, 2, 15),
            date(2023, 6, 20), date(2024, 2, 10), date(2024, 6, 15),
            date(2024, 10, 20), date(2025, 2, 12), date(2025, 6, 18)
        ],
        '운영계약종료일': [
            date(2028, 1, 14), date(2028, 3, 19), date(2028, 5, 9),
            date(2028, 7, 4), date(2028, 8, 31), date(2029, 1, 9),
            date(2029, 3, 14), date(2029, 5, 19), date(2029, 7, 7),
            date(2029, 9, 11), date(2030, 1, 4), date(2030, 3, 9),
            date(2030, 5, 14), date(2030, 7, 19), date(2030, 9, 4),
            date(2031, 1, 7), date(2031, 3, 11), date(2031, 5, 17),
            date(2031, 7, 21), date(2031, 9, 9), date(2028, 2, 13),
            date(2028, 6, 17), date(2028, 10, 21), date(2029, 2, 14),
            date(2029, 6, 19), date(2030, 2, 9), date(2030, 6, 14),
            date(2030, 10, 19), date(2031, 2, 11), date(2031, 6, 17)
        ]
    }

    df = pd.DataFrame(sample_data)
    df['운영계약시작일_parsed'] = df['운영계약시작일']
    df['운영계약종료일_parsed'] = df['운영계약종료일']
    df['운영계약시작일_cleaned'] = df['운영계약시작일']
    df['운영계약종료일_cleaned'] = df['운영계약종료일']
    df['행번호'] = range(5, 5 + len(df))
    return df


# ============================================================
# 지도 생성
# ============================================================

def count_fast_chargers(model_series):
    """급속 충전기 수를 세는 함수 (lambda 대신 사용)"""
    return sum(1 for m in model_series if '급속' in str(m))


def create_charger_map(filtered_df):
    map_data = filtered_df.dropna(subset=['위도', '경도']).copy()

    if len(map_data) == 0:
        return None, "좌표 데이터가 없습니다."

    if '사이트ID' not in map_data.columns or map_data['사이트ID'].isna().all():
        map_data['사이트ID'] = [f'SITE_{i:04d}' for i in range(len(map_data))]

    # 사이트별 그룹화
    grouped = map_data.groupby('사이트ID').agg({
        '위도': 'first',
        '경도': 'first',
        '주소': 'first',
        '권역': 'first',
        '모델분류': list,
        '운영계약시작일': 'first',
        '운영계약종료일': 'first'
    }).reset_index()

    # 충전기 수 계산
    site_counts = map_data.groupby('사이트ID')['모델분류'].agg(['count', count_fast_chargers]).reset_index()
    site_counts.columns = ['사이트ID', '총충전기수', '급속충전기수']
    site_counts['완속충전기수'] = site_counts['총충전기수'] - site_counts['급속충전기수']

    grouped = grouped.merge(site_counts, on='사이트ID')

    center_lat = grouped['위도'].mean()
    center_lon = grouped['경도'].mean()

    m = folium.Map(location=[center_lat, center_lon], zoom_start=8, tiles='OpenStreetMap')

    marker_cluster = MarkerCluster(
        name="충전소 클러스터",
        options={"disableClusteringAtZoom": 15, "maxClusterRadius": 50}
    ).add_to(m)

    region_colors = {
        '수도권북서': 'blue', '수도권북동': 'green', '수도권남동': 'red',
        '수도권남서': 'purple', '수도권기타': 'cadetblue', '인천기타': 'orange',
        '강원권': 'lightblue', '충청권': 'lightgreen', '경상권': 'pink',
        '전라권': 'lightgray', '기타': 'gray'
    }

    for _, row in grouped.iterrows():
        site_id = row['사이트ID']
        address = row['주소'] if row['주소'] else ''
        total = row['총충전기수']
        fast = row['급속충전기수']
        slow = row['완속충전기수']
        models = row['모델분류']
        region = row['권역']

        encoded_addr = quote(f"{address} 전기차")
        naver_url = f"https://map.naver.com/p/search/{encoded_addr}"

        icon_name = 'flash' if fast > 0 else 'plug'
        color = region_colors.get(region, 'gray')

        tooltip_text = f"{site_id} | {total}기 | {region}"
        models_text = ', '.join(set(str(mm) for mm in models))

        popup_html = f"""
        <div style="width:300px; font-family:Arial,sans-serif;">
            <h4 style="margin:0 0 10px 0; border-bottom:3px solid {color}; padding-bottom:6px;">
                {site_id}
            </h4>
            <table style="width:100%; font-size:13px; border-collapse:collapse;">
                <tr style="background:#f8f9fa;"><td style="padding:5px;"><b>충전기</b></td>
                    <td style="padding:5px; color:#0066cc;"><b>{total}대</b> (급속 {fast} / 완속 {slow})</td></tr>
                <tr><td style="padding:5px;"><b>권역</b></td><td style="padding:5px;">{region}</td></tr>
                <tr style="background:#f8f9fa;"><td style="padding:5px;"><b>주소</b></td>
                    <td style="padding:5px; font-size:11px;">{address}</td></tr>
                <tr><td style="padding:5px;"><b>모델</b></td>
                    <td style="padding:5px; font-size:11px;">{models_text}</td></tr>
            </table>
            <div style="text-align:center; margin-top:12px;">
                <a href="{naver_url}" target="_blank" style="
                    display:inline-block; padding:8px 16px;
                    background:linear-gradient(135deg,#03C75A,#029B47);
                    color:white; text-decoration:none; border-radius:6px;
                    font-weight:bold; font-size:12px;">
                    네이버 지도에서 보기
                </a>
            </div>
        </div>
        """

        folium.Marker(
            location=[row['위도'], row['경도']],
            popup=folium.Popup(popup_html, max_width=350),
            tooltip=tooltip_text,
            icon=folium.Icon(color=color, icon=icon_name, prefix='fa')
        ).add_to(marker_cluster)

    # 범례
    legend_html = f'''
    <div style="position:fixed; bottom:50px; right:50px; border:2px solid grey;
                z-index:9999; background:white; padding:12px; font-size:12px;
                border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.3);">
        <p style="margin:0 0 8px; font-weight:bold; font-size:13px;">지도 범례</p>
        <p style="margin:3px 0;">총 사이트: {len(grouped):,}개 | 총 충전기: {len(map_data):,}대</p>
    '''
    for rgn, clr in region_colors.items():
        if rgn in grouped['권역'].values:
            cnt = len(grouped[grouped['권역'] == rgn])
            legend_html += f'<p style="margin:1px 0; font-size:10px;"><span style="color:{clr}; font-size:14px;">●</span> {rgn} ({cnt})</p>'
    legend_html += '</div>'
    m.get_root().html.add_child(folium.Element(legend_html))

    return m, None


# ============================================================
# 엑셀 파일 처리 (업로드용 - 진행 바 포함)
# ============================================================

def process_excel_file_with_progress(file_bytes, title_container, progress_bar, status_text):
    try:
        start_time = time.time()

        title_container.markdown(f"### 작업 진행 상황 `{format_time(0)}`")
        status_text.markdown("**엑셀 파일을 읽는 중...**")
        progress_bar.progress(5)

        file_stream = io.BytesIO(file_bytes)
        wb = openpyxl.load_workbook(file_stream, data_only=True)
        ws = wb.active

        BA_COL = 53
        BB_COL = 54
        AR_COL = 44
        AS_COL = 45

        ws.cell(row=4, column=BA_COL, value='모델분류')
        ws.cell(row=4, column=BB_COL, value='권역')

        status_text.markdown("**데이터 분석 중...**")
        progress_bar.progress(15)

        max_row = max(ws.max_row, 5)
        max_col = max(ws.max_column, 54)
        total_rows = max_row - 4

        dashboard_data = []
        ar_cleaned_count = 0
        as_cleaned_count = 0

        status_text.markdown(f"**분류 시작... (총 {total_rows:,}행)**")
        progress_bar.progress(20)

        processed = 0
        last_update = time.time()

        for i, row_num in enumerate(range(5, max_row + 1)):
            row_data = {}
            for col_num in range(1, max_col + 1):
                col_letter = get_column_letter(col_num)
                row_data[col_letter] = ws.cell(row=row_num, column=col_num).value

            result = classify_model(row_data, row_num)
            ws.cell(row=row_num, column=BA_COL, value=result)

            address = get_safe_value(row_data, 'H')
            region = classify_region(address)
            ws.cell(row=row_num, column=BB_COL, value=region)

            ar_val = row_data.get('AR')
            ar_clean = clean_and_parse_date(ar_val)
            if ar_clean:
                ws.cell(row=row_num, column=AR_COL, value=format_date_for_excel(ar_clean))
                ws.cell(row=row_num, column=AR_COL).number_format = 'YYYY-MM-DD'
                ar_cleaned_count += 1

            as_val = row_data.get('AS')
            as_clean = clean_and_parse_date(as_val)
            if as_clean:
                ws.cell(row=row_num, column=AS_COL, value=format_date_for_excel(as_clean))
                ws.cell(row=row_num, column=AS_COL).number_format = 'YYYY-MM-DD'
                as_cleaned_count += 1

            site_id = get_safe_value(row_data, 'A')
            try:
                lon = float(row_data.get('AM')) if row_data.get('AM') else None
            except (ValueError, TypeError):
                lon = None
            try:
                lat = float(row_data.get('AN')) if row_data.get('AN') else None
            except (ValueError, TypeError):
                lat = None

            dashboard_data.append({
                '사이트ID': site_id if site_id else f'AUTO_{row_num}',
                '모델분류': result,
                '권역': region,
                '주소': address,
                '위도': lat,
                '경도': lon,
                '운영계약시작일': ar_val,
                '운영계약종료일': as_val,
                '운영계약시작일_cleaned': ar_clean,
                '운영계약종료일_cleaned': as_clean,
                '행번호': row_num
            })

            processed += 1
            now = time.time()

            if (i + 1) % 50 == 0 or (now - last_update) >= 0.5 or i == total_rows - 1:
                elapsed = now - start_time
                pct = 20 + int((processed / total_rows) * 70)
                progress_bar.progress(pct)
                title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")

                if processed > 0:
                    speed = processed / elapsed
                    remain = (total_rows - processed) / speed if speed > 0 else 0
                    status_text.markdown(
                        f"**분류 중...** `{processed:,}/{total_rows:,}` "
                        f"({processed / total_rows * 100:.1f}%) | "
                        f"속도: `{speed:.1f}행/초` | 남은 시간: `{format_time(remain)}`"
                    )
                last_update = now

        elapsed = time.time() - start_time
        status_text.markdown("**결과 파일 생성 중...**")
        progress_bar.progress(95)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        df = pd.DataFrame(dashboard_data)
        df['운영계약시작일_parsed'] = df['운영계약시작일_cleaned']
        df['운영계약종료일_parsed'] = df['운영계약종료일_cleaned']

        total_time = time.time() - start_time
        progress_bar.progress(100)
        title_container.markdown(f"### 작업 완료! `총 {format_time(total_time)}`")

        speed = processed / total_time if total_time > 0 else 0
        status_text.markdown(
            f"**완료!** `{processed:,}행` 분류 | 속도: `{speed:.1f}행/초` | "
            f"날짜 정리: AR `{ar_cleaned_count:,}`, AS `{as_cleaned_count:,}`"
        )

        return output, None, processed, total_time, df, ar_cleaned_count, as_cleaned_count

    except Exception as e:
        import traceback
        elapsed = time.time() - start_time
        title_container.markdown(f"### 작업 중단 `{format_time(elapsed)}`")
        status_text.markdown("**오류 발생**")
        return None, f"오류: {str(e)}\n\n{traceback.format_exc()}", 0, 0, None, 0, 0


# ============================================================
# 대시보드
# ============================================================

def show_dashboard(df):
    st.markdown("## 충전기 운영 현황 대시보드")
    st.markdown("### 운영계약 기간 필터")

    valid_dates = df.dropna(subset=['운영계약시작일_parsed', '운영계약종료일_parsed'])

    if len(valid_dates) == 0:
        st.warning("유효한 운영계약 날짜 데이터가 없습니다. AR열과 AS열을 확인해주세요.")
        return

    min_date = valid_dates['운영계약시작일_parsed'].min()
    max_date = valid_dates['운영계약종료일_parsed'].max()

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

    mask = (
        (df['운영계약시작일_parsed'] < end_date) &
        (df['운영계약종료일_parsed'] >= start_date) &
        df['운영계약시작일_parsed'].notna() &
        df['운영계약종료일_parsed'].notna()
    )
    filtered_df = df[mask].copy()

    st.info(
        f"선택 기간: {start_date} ~ {end_date} | "
        f"해당 충전기: {len(filtered_df):,}대 (전체 {len(df):,}대 중 {len(filtered_df) / max(len(df), 1) * 100:.1f}%)"
    )

    if len(filtered_df) == 0:
        st.warning("선택한 기간에 해당하는 데이터가 없습니다.")
        return

    st.markdown("---")

    # 주요 지표
    total_chargers = len(filtered_df)
    unique_sites = filtered_df['사이트ID'].nunique() if '사이트ID' in filtered_df.columns else 0
    region_count = filtered_df['권역'].nunique()
    model_count = filtered_df['모델분류'].nunique()
    fast_chargers = len(filtered_df[filtered_df['모델분류'].str.contains('급속', na=False)])
    fast_ratio = (fast_chargers / total_chargers * 100) if total_chargers > 0 else 0

    st.markdown("### 주요 지표")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 충전기 수", f"{total_chargers:,}대")
    c2.metric("사이트 수", f"{unique_sites:,}개")
    c3.metric("권역 수", f"{region_count}개")
    c4.metric("급속 충전기", f"{fast_chargers:,}대", f"{fast_ratio:.1f}%")

    st.markdown("---")

    # 지도
    st.markdown("### 충전기 위치 지도")
    has_coords = '위도' in filtered_df.columns and '경도' in filtered_df.columns

    if has_coords:
        valid_coords = filtered_df.dropna(subset=['위도', '경도'])
        if len(valid_coords) > 0:
            sites_map = valid_coords['사이트ID'].nunique() if '사이트ID' in valid_coords.columns else len(valid_coords)
            st.success(f"{sites_map:,}개 사이트, {len(valid_coords):,}개 충전기 좌표 로드 완료")

            charger_map, err = create_charger_map(filtered_df)
            if err:
                st.error(err)
            else:
                st_folium(charger_map, width=1400, height=700)

                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.metric("지도 사이트", f"{sites_map:,}개")
                mc2.metric("지도 충전기", f"{len(valid_coords):,}대")
                mc3.metric("좌표 보유율", f"{len(valid_coords) / max(len(filtered_df), 1) * 100:.1f}%")
                mc4.metric("좌표 누락", f"{len(filtered_df) - len(valid_coords):,}대")
        else:
            st.warning("좌표 데이터가 없습니다.")
    else:
        st.warning("위도/경도 컬럼이 없습니다.")

    st.markdown("---")

    # 모델별 현황
    st.markdown("### 모델별 충전기 현황")
    col1, col2 = st.columns([3, 2])

    with col1:
        model_counts = filtered_df['모델분류'].value_counts().reset_index()
        model_counts.columns = ['모델분류', '수량']

        fig = px.bar(
            model_counts.head(15), x='수량', y='모델분류', orientation='h',
            title='모델별 충전기 수량 (상위 15개)', color='수량',
            color_continuous_scale='Blues', text='수량'
        )
        fig.update_layout(height=500, showlegend=False)
        fig.update_traces(texttemplate='%{text}', textposition='outside')
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("#### 모델별 수량 상세")
        model_counts['비율'] = (model_counts['수량'] / model_counts['수량'].sum() * 100).round(1).astype(str) + '%'
        st.dataframe(model_counts[['모델분류', '수량', '비율']], use_container_width=True, hide_index=True, height=450)

    st.markdown("---")

    # 권역별 현황
    st.markdown("### 권역별 충전기 현황")
    col1, col2 = st.columns([2, 3])

    with col1:
        region_counts = filtered_df['권역'].value_counts().reset_index()
        region_counts.columns = ['권역', '수량']

        fig_pie = px.pie(region_counts, values='수량', names='권역', title='권역별 비율', hole=0.4)
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie.update_layout(height=400)
        st.plotly_chart(fig_pie, use_container_width=True)

    with col2:
        fig_bar = px.bar(
            region_counts, x='권역', y='수량', title='권역별 수량',
            color='수량', color_continuous_scale='Greens', text='수량'
        )
        fig_bar.update_layout(height=400, showlegend=False)
        fig_bar.update_traces(texttemplate='%{text}', textposition='outside')
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")

    # 히트맵
    st.markdown("### 권역별 모델 분포 히트맵")
    crosstab = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'])
    top_models = filtered_df['모델분류'].value_counts().head(12).index
    crosstab_top = crosstab[[c for c in top_models if c in crosstab.columns]]

    if len(crosstab_top.columns) > 0:
        fig_heat = px.imshow(
            crosstab_top.T, labels=dict(x="권역", y="모델분류", color="수량"),
            color_continuous_scale='RdYlGn', aspect="auto",
            title='권역별 주요 모델 분포 (상위 12개)', text_auto=True
        )
        fig_heat.update_layout(height=500)
        st.plotly_chart(fig_heat, use_container_width=True)

    st.markdown("---")

    # 상세 테이블
    st.markdown("### 권역별 x 모델별 상세 현황")
    pivot = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'], margins=True)
    st.dataframe(pivot, use_container_width=True, height=400)

    # 다운로드
    st.markdown("---")
    st.markdown("### 데이터 다운로드")
    dc1, dc2, dc3 = st.columns(3)

    with dc1:
        st.download_button(
            "권역x모델 현황표 CSV", pivot.to_csv(encoding='utf-8-sig'),
            f"현황표_{start_date}_{end_date}.csv", "text/csv", use_container_width=True
        )
    with dc2:
        st.download_button(
            "필터링 데이터 CSV", filtered_df.to_csv(index=False, encoding='utf-8-sig'),
            f"필터데이터_{start_date}_{end_date}.csv", "text/csv", use_container_width=True
        )
    with dc3:
        report = f"""충전기 현황 요약
기간: {start_date} ~ {end_date}
총 충전기: {total_chargers:,}대
사이트: {unique_sites:,}개
모델 종류: {model_count}개
급속: {fast_chargers:,}대 ({fast_ratio:.1f}%)

상위 5개 모델:
{chr(10).join(f"  {r['모델분류']}: {r['수량']:,}대 ({r['비율']})" for _, r in model_counts.head(5).iterrows())}

권역별:
{chr(10).join(f"  {r['권역']}: {r['수량']:,}대" for _, r in region_counts.iterrows())}
"""
        st.download_button(
            "요약 리포트 TXT", report,
            f"요약_{start_date}_{end_date}.txt", "text/plain", use_container_width=True
        )

    # 데이터 품질
    st.markdown("---")
    st.markdown("### 데이터 품질 체크")

    problem_regions = ['수도권기타', '인천기타', '기타']
    unknown = filtered_df[filtered_df['권역'].isin(problem_regions)]

    qc1, qc2 = st.columns(2)
    normal = len(filtered_df) - len(unknown)
    qc1.metric("정상 분류", f"{normal:,}대", f"{normal / max(len(filtered_df), 1) * 100:.1f}%")
    qc2.metric("미분류/불명확", f"{len(unknown):,}대", f"{len(unknown) / max(len(filtered_df), 1) * 100:.1f}%")

    if len(unknown) > 0:
        st.warning(f"{len(unknown):,}개의 주소가 미분류/불명확합니다.")
        with st.expander("미분류 주소 상세"):
            cols = [c for c in ['주소', '권역', '모델분류', '사이트ID'] if c in unknown.columns]
            st.dataframe(unknown[cols].head(10), use_container_width=True, hide_index=True)
    else:
        st.success("모든 주소가 정상 분류되었습니다!")


# ============================================================
# 분류 기준 정보
# ============================================================

def show_classification_info():
    st.markdown("### 모델분류 기준표")
    t1, t2, t3 = st.tabs(["급속 충전기", "완속 충전기", "권역 분류"])

    with t1:
        st.info("전제 조건: AH열 = '급속'")
        st.dataframe({
            "모델분류명": [
                "급속스필_100", "급속스필_50", "급속PNE_100", "급속PNE_50",
                "급속PNE_200", "급속PNE_150", "급속애플망고_200", "급속SK_100",
                "급속SK_200", "급속코스텔_50", "급속중앙제어_50", "급속그린파워_100",
                "급속그린파워_50", "급속알박_50", "급속"
            ],
            "AG열 코드": [
                "S0F1", "S0F5", "EVQ- (AJ=100)", "EVQ-/EV1- (AJ=50)",
                "MAXE", "DP15", "A01-/AD1-", "Q081/Q101/Q010",
                "Q071/Q102", "1Y25/1Y24", "1911", "1900",
                "19C0", "QC50", "해당없음"
            ]
        }, use_container_width=True, hide_index=True)

    with t2:
        st.info("전제 조건: AH열 ≠ '급속'")
        st.dataframe({
            "모델분류명": [
                "알박구형(NC07)", "알박신형(23/22/24/25NA)", "10kW(3J10)",
                "신형대(EVL-1C-22CQ)", "구형대(EVL-1C)", "신형대(SBDA)",
                "신형소(SBAA)", "F01/PC01(SBPA)", "UC01(SBUA)",
                "스필_7kW(SVI0)", "이카플러그(E0C/CP)", "중앙제어_7kW(1907/1912)",
                "SK_7kW(SC-P)", "3kW(SANA)", "PNE_7kW(EVS-/007S)",
                "F01/PC01(SBOA)", "기타"
            ]
        }, use_container_width=True, hide_index=True)

    with t3:
        st.dataframe({
            "권역": ["수도권북서", "수도권북동", "수도권남동", "수도권남서",
                    "강원권", "충청권", "경상권", "전라권", "기타"],
            "지역": [
                "고양,부천,김포,파주,은평,마포,서대문,양천,강서",
                "도봉,노원,중랑,강북,성북,동대문,의정부,남양주,구리",
                "강남,서초,송파,강동,성남,용인,하남,수원,평택",
                "구로,금천,영등포,동작,관악,의왕,광명,안산,안양,인천",
                "강원도 전역",
                "충청,충남,충북,세종,대전",
                "경상,경남,경북,부산,대구,울산",
                "전라,전남,전북,광주",
                "분류 불가"
            ]
        }, use_container_width=True, hide_index=True)


# ============================================================
# 메인
# ============================================================

def main():
    # 세션 초기화: 파일 우선 → 샘플 폴백
    if 'processed_df' not in st.session_state:
        default_file = find_default_data_file()
        if default_file:
            df, error = load_initial_data_from_file(default_file)
            if df is not None and error is None:
                st.session_state.processed_df = df
                st.session_state.is_sample_data = False
                st.session_state.default_file_name = default_file
                st.session_state.default_file_rows = len(df)
            else:
                st.session_state.processed_df = create_sample_data()
                st.session_state.is_sample_data = True
                st.session_state.file_load_error = error
        else:
            st.session_state.processed_df = create_sample_data()
            st.session_state.is_sample_data = True

    if 'processed_file' not in st.session_state:
        st.session_state.processed_file = None

    # 헤더
    st.title("⚡ 충전기 모델분류 & 운영현황 대시보드")
    st.markdown(
        "엑셀 파일을 업로드하면 **BA열**(모델분류), **BB열**(권역)을 자동 추가하고, "
        "**AM열(경도)/AN열(위도)** 기반 지도를 생성합니다."
    )

    # 기본 파일 로드 상태
    if st.session_state.get('default_file_name'):
        st.success(
            f"기본 데이터 자동 로드: `{st.session_state.default_file_name}` "
            f"({st.session_state.get('default_file_rows', 0):,}행)"
        )
    if st.session_state.get('file_load_error'):
        st.warning(f"기본 파일 로드 실패: {st.session_state.file_load_error}")

    # 탭
    tab1, tab2 = st.tabs(["파일 업로드 & 분류", "운영현황 대시보드"])

    with tab1:
        # 기본 데이터 파일 설정
        with st.expander("기본 데이터 파일 설정", expanded=False):
            st.markdown("앱 시작 시 자동 로드할 파일 경로:")
            for path in DEFAULT_DATA_FILES:
                exists = "존재" if os.path.exists(path) else "없음"
                st.code(f"{path}  →  {exists}")

            st.markdown("---")
            st.markdown("**기본 데이터 파일 직접 업로드:**")
            default_upload = st.file_uploader(
                "기본 데이터 (xlsx/csv)", type=['xlsx', 'xls', 'csv'],
                key='default_uploader'
            )

            if default_upload is not None:
                if st.button("기본 데이터로 설정", type="primary"):
                    file_bytes = default_upload.read()
                    df, error = load_initial_data_from_bytes(file_bytes, default_upload.name)

                    if df is not None:
                        # 파일도 저장 시도 (실패해도 메모리에서는 동작)
                        save_path = DEFAULT_DATA_FILES[0]
                        try:
                            with open(save_path, 'wb') as f:
                                f.write(file_bytes)
                            st.session_state.default_file_name = save_path
                        except Exception:
                            st.session_state.default_file_name = f"메모리 ({default_upload.name})"

                        st.session_state.processed_df = df
                        st.session_state.is_sample_data = False
                        st.session_state.default_file_rows = len(df)
                        st.success(f"기본 데이터 설정 완료: {len(df):,}행 로드")
                        st.rerun()
                    else:
                        st.error(f"로드 실패: {error}")

        st.markdown("---")

        if st.session_state.get('is_sample_data', False):
            st.info("현재 샘플 데이터입니다. 위에서 기본 파일을 등록하거나 아래에서 업로드하세요.")

        uploaded_file = st.file_uploader(
            "엑셀 파일 선택", type=['xlsx', 'xls'],
            help="AM열(경도), AN열(위도) 포함 파일 권장"
        )

        if uploaded_file is not None:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.info(f"**{uploaded_file.name}**")
            with col2:
                size_mb = uploaded_file.size / (1024 * 1024)
                st.metric("크기", f"{size_mb:.1f} MB" if size_mb >= 1 else f"{uploaded_file.size / 1024:.1f} KB")

            if st.button("모델분류 시작", type="primary", use_container_width=True):
                title_c = st.empty()
                prog = st.progress(0)
                status = st.empty()
                st.markdown("---")

                result = process_excel_file_with_progress(uploaded_file.read(), title_c, prog, status)
                out_file, err, count, t_time, df, ar_c, as_c = result

                if err:
                    st.error(err)
                else:
                    st.session_state.processed_df = df
                    st.session_state.processed_file = out_file
                    st.session_state.is_sample_data = False

                    st.success(f"**{count:,}행** 분류 완료! ({format_time(t_time)})")
                    st.info(f"날짜 정리: AR열 `{ar_c:,}개`, AS열 `{as_c:,}개`")

                    korea_time = get_korea_time()
                    ts = korea_time.strftime("%Y%m%d_%H%M%S")

                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        st.download_button(
                            "결과 파일 다운로드", out_file.getvalue(),
                            f"모델분류_{ts}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True, type="primary"
                        )

        with st.expander("분류 기준 정보", expanded=False):
            show_classification_info()

    with tab2:
        if st.session_state.processed_df is not None:
            if st.session_state.get('is_sample_data', False):
                st.warning("현재 샘플 데이터입니다. 실제 분석은 파일 업로드 후 가능합니다.")
            show_dashboard(st.session_state.processed_df)
        else:
            st.info("먼저 파일을 업로드하고 분류를 완료해주세요.")


if __name__ == "__main__":
    main()
