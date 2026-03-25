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

st.set_page_config(page_title="충전기 모델분류 자동화", page_icon="⚡", layout="wide")

# ── 파일 경로 (parquet 우선, xlsx 폴백) ──
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PARQUET = os.path.join(BASE_DIR, "default_data.parquet")
DEFAULT_XLSX = os.path.join(BASE_DIR, "default_data.xlsx")


# ──────────────────────────────────────────────
# 유틸리티
# ──────────────────────────────────────────────

def get_korea_time():
    return datetime.now(pytz.timezone('Asia/Seoul'))

def format_time(seconds):
    if seconds < 60:
        return f"{seconds:.1f}초"
    elif seconds < 3600:
        return f"{int(seconds//60)}분 {int(seconds%60)}초"
    else:
        return f"{int(seconds//3600)}시간 {int((seconds%3600)//60)}분"

def format_date_for_excel(date_obj):
    if date_obj is None:
        return None
    if isinstance(date_obj, date):
        return date_obj.strftime('%Y-%m-%d')
    return None


# ──────────────────────────────────────────────
# 벡터화 분류 함수
# ──────────────────────────────────────────────

def classify_region_vectorized(addresses: pd.Series) -> pd.Series:
    addr = addresses.fillna('').astype(str).str.strip()
    result = pd.Series('기타', index=addr.index)

    incheon = addr.str.contains('인천', na=False)
    incheon_detail = addr.str.contains('계양|남동|동구|미추홀|부평|연수|서구|중구|강화', na=False)
    result[incheon & incheon_detail] = '수도권남서'
    result[incheon & ~incheon_detail] = '인천기타'

    sg = addr.str.contains('서울|경기', na=False) & ~incheon
    nw = addr.str.contains('고양|부천|김포|파주|은평구|마포구|서대문구|양천구|강서구|용산구|중구|종로구', na=False)
    ne = addr.str.contains('도봉구|노원구|중랑구|강북구|성북구|동대문구|성동구|광진구|의정부|남양주|구리|양주|포천|동두천|가평|연천', na=False)
    se = addr.str.contains('강남구|서초구|송파구|강동구|성남|용인|하남|광주|안성|수원|평택|오산|이천|여주|양평', na=False)
    sw = addr.str.contains('구로구|금천구|영등포구|동작구|관악구|의왕|광명|군포|과천|시흥|안산|안양|화성', na=False)

    result[sg & nw] = '수도권북서'
    result[sg & ne & ~nw] = '수도권북동'
    result[sg & se & ~nw & ~ne] = '수도권남동'
    result[sg & sw & ~nw & ~ne & ~se] = '수도권남서'
    result[sg & ~nw & ~ne & ~se & ~sw] = '수도권기타'
    result[addr.str.contains('강원', na=False) & ~incheon & ~sg] = '강원권'
    result[addr.str.contains('충청|충남|충북|세종|대전', na=False) & ~incheon & ~sg] = '충청권'
    result[addr.str.contains('경상|경남|경북|부산|대구|울산', na=False) & ~incheon & ~sg] = '경상권'
    result[addr.str.contains('전라|전남|전북|광주', na=False) & ~incheon & ~sg] = '전라권'
    return result


def classify_model_vectorized(df: pd.DataFrame) -> pd.Series:
    AD = df['AD'].fillna('').astype(str).str.strip()
    AG = df['AG'].fillna('').astype(str).str.strip()
    AH = df['AH'].fillna('').astype(str).str.strip()
    AJ = df['AJ'].fillna('').astype(str).str.strip()

    ag4 = AG.str[:4]; ag3 = AG.str[:3]; ag6 = AG.str[:6]; ag11 = AG.str[:11]
    is_fast = (AH == '급속')

    fast_conds = [
        is_fast & (ag4 == 'S0F1'), is_fast & (ag4 == 'S0F5'),
        is_fast & (ag4 == 'EVQ-') & (AJ == '100'),
        is_fast & (ag4.isin(['EVQ-', 'EV1-'])) & (AJ == '50'),
        is_fast & (ag4 == 'MAXE'), is_fast & (ag4 == 'DP15'),
        is_fast & (ag4.isin(['A01-', 'AD1-'])),
        is_fast & (ag4.isin(['Q081', 'Q101', 'Q010'])),
        is_fast & (ag4.isin(['Q071', 'Q102'])),
        is_fast & (ag4.isin(['1Y25', '1Y24'])),
        is_fast & (ag4 == '1911'), is_fast & (ag4 == '1900'),
        is_fast & (ag4 == '19C0'), is_fast & (ag4 == 'QC50'),
    ]
    fast_vals = [
        '급속스필_100', '급속스필_50', '급속PNE_100', '급속PNE_50',
        '급속PNE_200', '급속PNE_150', '급속애플망고_200', '급속SK_100',
        '급속SK_200', '급속코스텔_50', '급속중앙제어_50', '급속그린파워_100',
        '급속그린파워_50', '급속알박_50',
    ]
    result = pd.Series(np.select(fast_conds, fast_vals, default='__PENDING__'), index=df.index)
    result[(result == '__PENDING__') & is_fast] = '급속'

    slow = ~is_fast
    slow_conds = [
        slow & (ag4 == 'NC07'),
        slow & (ag4.isin(['23NA', '22NA', '24NA', '25NA'])),
        slow & (AD.str.contains('3J10', na=False)),
        slow & (ag11 == 'EVL-1C-22CQ'),
        slow & (ag6 == 'EVL-1C') & (ag11 != 'EVL-1C-22CQ'),
        slow & (ag4 == 'EVL-') & AD.str.contains('1107', na=False) & (ag6 != 'EVL-1C'),
        slow & (ag4 == 'EVL-') & ~AD.str.contains('1107', na=False) & (ag6 != 'EVL-1C'),
        slow & (ag4 == 'SBDA'), slow & (ag4 == 'SBAA'),
        slow & (ag4 == 'SBPA') & AD.str.contains('F01', na=False),
        slow & (ag4 == 'SBPA') & ~AD.str.contains('F01', na=False),
        slow & (ag4 == 'SBUA'), slow & (ag4 == 'SVI0'),
        slow & ((ag3 == 'E0C') | AD.str.contains('CP', na=False)),
        slow & (ag4.isin(['1907', '1912'])),
        slow & (ag4 == 'SC-P'), slow & (ag4 == 'SANA'),
        slow & (ag4.isin(['EVS-', '007S'])),
        slow & (ag4 == 'SBOA') & AD.str.contains('F01', na=False),
        slow & (ag4 == 'SBOA') & ~AD.str.contains('F01', na=False),
    ]
    slow_vals = [
        '알박구형', '알박신형', '10kW', '신형대', '구형대', '신형대', '구형대',
        '신형대', '신형소', 'F01', 'PC01', 'UC01', '스필_7kW', '이카플러그',
        '중앙제어_7kW', 'SK_7kW', '3kW', 'PNE_7kW', 'F01', 'PC01',
    ]
    slow_result = pd.Series(np.select(slow_conds, slow_vals, default='기타'), index=df.index)
    pending = (result == '__PENDING__')
    result[pending] = slow_result[pending]
    return result


# ──────────────────────────────────────────────
# 데이터 로드 함수
# ──────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def load_default_parquet(filepath):
    """Parquet 파일 고속 로드 (~0.1초)"""
    df = pd.read_parquet(filepath, engine='pyarrow')
    # 날짜 열 보장
    df['운영계약시작일'] = pd.to_datetime(df['운영계약시작일'], errors='coerce')
    df['운영계약종료일'] = pd.to_datetime(df['운영계약종료일'], errors='coerce')
    df['운영계약시작일_parsed'] = df['운영계약시작일']
    df['운영계약종료일_parsed'] = df['운영계약종료일']
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def load_default_xlsx(filepath):
    """xlsx 폴백 로드 (parquet 없을 때)"""
    raw_df = pd.read_excel(filepath, header=3, engine='openpyxl', dtype=str)
    raw_df = raw_df.dropna(how='all').reset_index(drop=True)
    return build_dashboard_df_from_raw(raw_df)


def build_dashboard_df_from_raw(raw_df):
    """pandas 원본 DF → 대시보드 DF (벡터 분류 포함)"""
    cols = raw_df.columns.tolist()
    def safe_col(idx):
        return cols[idx] if idx < len(cols) else None

    col_A = safe_col(0); col_H = safe_col(7)
    col_AD = safe_col(29); col_AG = safe_col(32); col_AH = safe_col(33); col_AJ = safe_col(35)
    col_AM = safe_col(38); col_AN = safe_col(39); col_AR = safe_col(43); col_AS = safe_col(44)

    classify_df = pd.DataFrame({
        'AD': raw_df[col_AD].fillna('').astype(str) if col_AD else '',
        'AG': raw_df[col_AG].fillna('').astype(str) if col_AG else '',
        'AH': raw_df[col_AH].fillna('').astype(str) if col_AH else '',
        'AJ': raw_df[col_AJ].fillna('').astype(str) if col_AJ else '',
    })

    addresses = raw_df[col_H].fillna('').astype(str) if col_H else pd.Series('', index=raw_df.index)
    model_result = classify_model_vectorized(classify_df)
    region_result = classify_region_vectorized(addresses)

    site_ids = raw_df[col_A].fillna('').astype(str).str.strip() if col_A else pd.Series('', index=raw_df.index)
    site_ids[site_ids == ''] = 'AUTO_' + (raw_df.index + 5).astype(str)

    lon = pd.to_numeric(raw_df[col_AM], errors='coerce') if col_AM else pd.Series(dtype=float)
    lat = pd.to_numeric(raw_df[col_AN], errors='coerce') if col_AN else pd.Series(dtype=float)

    ar_parsed = pd.to_datetime(raw_df[col_AR], errors='coerce') if col_AR else pd.Series(dtype='datetime64[ns]')
    as_parsed = pd.to_datetime(raw_df[col_AS], errors='coerce') if col_AS else pd.Series(dtype='datetime64[ns]')

    return pd.DataFrame({
        '사이트ID': site_ids.values, '모델분류': model_result.values, '권역': region_result.values,
        '주소': addresses.values, '위도': lat.values, '경도': lon.values,
        '운영계약시작일': ar_parsed.values, '운영계약종료일': as_parsed.values,
        '운영계약시작일_parsed': ar_parsed.values, '운영계약종료일_parsed': as_parsed.values,
        '행번호': raw_df.index + 5,
    })


# ──────────────────────────────────────────────
# ★ 고속 지도 생성 (사이트 미리 집계)
# ──────────────────────────────────────────────

@st.cache_data(ttl=600, show_spinner=False)
def prepare_map_data(filtered_df):
    """지도용 사이트 집계 DataFrame 생성 (캐시)"""
    map_data = filtered_df.dropna(subset=['위도', '경도']).copy()
    if len(map_data) == 0:
        return None

    if '사이트ID' not in map_data.columns or map_data['사이트ID'].isna().all():
        map_data['사이트ID'] = [f'SITE_{i:04d}' for i in range(len(map_data))]

    grouped = map_data.groupby('사이트ID').agg(
        위도=('위도', 'first'),
        경도=('경도', 'first'),
        주소=('주소', 'first'),
        권역=('권역', 'first'),
        모델목록=('모델분류', lambda x: ', '.join(sorted(set(x.astype(str))))),
        총충전기수=('모델분류', 'size'),
        급속충전기수=('모델분류', lambda x: x.astype(str).str.contains('급속').sum()),
    ).reset_index()

    grouped['완속충전기수'] = grouped['총충전기수'] - grouped['급속충전기수']
    return grouped


def create_charger_map(grouped_sites):
    """미리 집계된 사이트 DF로 지도 생성"""
    if grouped_sites is None or len(grouped_sites) == 0:
        return None, "좌표 데이터가 없습니다."

    center_lat = grouped_sites['위도'].mean()
    center_lon = grouped_sites['경도'].mean()

    m = folium.Map(location=[center_lat, center_lon], zoom_start=8, tiles='OpenStreetMap')

    marker_cluster = MarkerCluster(
        name="충전소", overlay=True, control=True,
        options={"disableClusteringAtZoom": 15, "maxClusterRadius": 50}
    ).add_to(m)

    region_colors = {
        '수도권북서': 'blue', '수도권북동': 'green', '수도권남동': 'red', '수도권남서': 'purple',
        '수도권기타': 'cadetblue', '인천기타': 'orange', '강원권': 'lightblue', '충청권': 'lightgreen',
        '경상권': 'pink', '전라권': 'lightgray', '기타': 'gray'
    }

    # ★ itertuples는 iterrows보다 ~5배 빠름
    for row in grouped_sites.itertuples(index=False):
        site_id = row.사이트ID
        address = row.주소 or ''
        total = row.총충전기수
        fast = row.급속충전기수
        slow = row.완속충전기수
        region = row.권역 or '기타'
        models_text = row.모델목록

        encoded_addr = quote(f"{address} 전기차")
        naver_url = f"https://map.naver.com/p/search/{encoded_addr}"
        color = region_colors.get(region, 'gray')
        icon_name = 'flash' if fast > 0 else 'plug'

        popup_html = f"""
        <div style="width:300px;font-family:'Malgun Gothic',Arial,sans-serif">
            <h4 style="margin:0 0 8px;border-bottom:3px solid {color};padding-bottom:6px">{site_id}</h4>
            <table style="width:100%;font-size:12px;border-collapse:collapse">
                <tr style="background:#f8f9fa"><td style="padding:4px;font-weight:bold">충전기</td><td style="padding:4px;color:#0066cc;font-weight:bold">{total}대 (급속 {fast} / 완속 {slow})</td></tr>
                <tr><td style="padding:4px;font-weight:bold">권역</td><td style="padding:4px">{region}</td></tr>
                <tr style="background:#f8f9fa"><td style="padding:4px;font-weight:bold">주소</td><td style="padding:4px;font-size:11px">{address}</td></tr>
                <tr><td style="padding:4px;font-weight:bold">모델</td><td style="padding:4px;font-size:11px">{models_text}</td></tr>
            </table>
            <div style="text-align:center;margin-top:10px">
                <a href="{naver_url}" target="_blank" style="display:inline-block;padding:8px 16px;background:linear-gradient(135deg,#03C75A,#029B47);color:white;text-decoration:none;border-radius:6px;font-weight:bold;font-size:12px">네이버 지도</a>
            </div>
        </div>"""

        folium.Marker(
            location=[row.위도, row.경도],
            popup=folium.Popup(popup_html, max_width=320),
            tooltip=f"{site_id} | {total}기 | {region}",
            icon=folium.Icon(color=color, icon=icon_name, prefix='fa')
        ).add_to(marker_cluster)

    # 범례
    total_sites = len(grouped_sites)
    total_chargers = grouped_sites['총충전기수'].sum()
    legend = f'<div style="position:fixed;bottom:50px;right:50px;border:2px solid grey;z-index:9999;background:white;padding:12px;font-size:11px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.3)">'
    legend += f'<b>사이트: {total_sites:,}개 / 충전기: {total_chargers:,}대</b><hr style="margin:6px 0">'
    for rg, cl in region_colors.items():
        cnt = len(grouped_sites[grouped_sites['권역'] == rg])
        if cnt > 0:
            legend += f'<span style="color:{cl};font-size:14px">●</span> {rg} ({cnt}개)<br>'
    legend += '</div>'
    m.get_root().html.add_child(folium.Element(legend))

    return m, None


# ──────────────────────────────────────────────
# 업로드 파일 처리
# ──────────────────────────────────────────────

def process_excel_file_with_progress(file_bytes, title_container, progress_bar, status_text):
    try:
        t0 = time.time()

        status_text.markdown("**엑셀 파일을 읽는 중...**")
        progress_bar.progress(5)

        file_stream = io.BytesIO(file_bytes)
        raw_df = pd.read_excel(file_stream, header=3, engine='openpyxl', dtype=str)
        raw_df = raw_df.dropna(how='all').reset_index(drop=True)
        total_rows = len(raw_df)

        elapsed = time.time() - t0
        title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")
        status_text.markdown(f"**{total_rows:,}행 로드 완료. 벡터 분류 중...**")
        progress_bar.progress(30)

        dashboard_df = build_dashboard_df_from_raw(raw_df)

        elapsed = time.time() - t0
        title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")
        status_text.markdown("**분류 완료. 엑셀 결과 기록 중...**")
        progress_bar.progress(60)

        # openpyxl로 결과 열만 쓰기
        file_stream.seek(0)
        wb = openpyxl.load_workbook(file_stream)
        ws = wb.active

        BA, BB, AR_COL, AS_COL = 53, 54, 44, 45
        ws.cell(row=4, column=BA, value='모델분류')
        ws.cell(row=4, column=BB, value='권역')

        models = dashboard_df['모델분류'].tolist()
        regions = dashboard_df['권역'].tolist()
        ar_dates = pd.to_datetime(dashboard_df['운영계약시작일'], errors='coerce')
        as_dates = pd.to_datetime(dashboard_df['운영계약종료일'], errors='coerce')

        ar_count = as_count = 0

        for i in range(total_rows):
            rn = i + 5
            ws.cell(row=rn, column=BA, value=models[i])
            ws.cell(row=rn, column=BB, value=regions[i])

            if pd.notna(ar_dates.iloc[i]):
                ws.cell(row=rn, column=AR_COL, value=ar_dates.iloc[i].strftime('%Y-%m-%d'))
                ws.cell(row=rn, column=AR_COL).number_format = 'YYYY-MM-DD'
                ar_count += 1
            if pd.notna(as_dates.iloc[i]):
                ws.cell(row=rn, column=AS_COL, value=as_dates.iloc[i].strftime('%Y-%m-%d'))
                ws.cell(row=rn, column=AS_COL).number_format = 'YYYY-MM-DD'
                as_count += 1

            if (i + 1) % 500 == 0 or i == total_rows - 1:
                elapsed = time.time() - t0
                pct = 60 + int((i / total_rows) * 35)
                progress_bar.progress(min(pct, 95))
                title_container.markdown(f"### 작업 진행 상황 `{format_time(elapsed)}`")
                status_text.markdown(f"**엑셀 쓰기...** `{i+1:,}/{total_rows:,}` ({(i+1)/total_rows*100:.1f}%)")

        progress_bar.progress(95)
        status_text.markdown("**파일 저장 중...**")

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        wb.close()

        total_time = time.time() - t0
        progress_bar.progress(100)
        title_container.markdown(f"### 작업 완료! `{format_time(total_time)}`")
        spd = total_rows / total_time if total_time > 0 else 0
        status_text.markdown(f"**완료!** `{total_rows:,}행` | `{spd:.0f}행/초` | AR `{ar_count:,}개`, AS `{as_count:,}개` 날짜 정리")

        return output, None, total_rows, total_time, dashboard_df, ar_count, as_count

    except Exception as e:
        import traceback
        elapsed = time.time() - t0
        title_container.markdown(f"### 작업 중단 `{format_time(elapsed)}`")
        return None, f"오류: {e}\n\n{traceback.format_exc()}", 0, 0, None, 0, 0


# ──────────────────────────────────────────────
# 샘플 데이터
# ──────────────────────────────────────────────

def create_sample_data():
    sample_data = {
        '사이트ID': [f'SITE_{i:03d}' for i in range(1, 21)] + [f'SITE_{i:03d}' for i in range(1, 11)],
        '모델분류': ['급속스필_100','급속PNE_100','신형대','알박신형','구형대','급속SK_100','F01','스필_7kW','급속그린파워_100','PNE_7kW',
                   '급속코스텔_50','신형소','이카플러그','급속애플망고_200','UC01','급속PNE_50','10kW','SK_7kW','중앙제어_7kW','3kW',
                   '급속스필_100','신형대','알박신형','급속PNE_100','구형대','급속SK_200','F01','급속그린파워_50','PNE_7kW','급속알박_50'],
        '권역': ['수도권북서','수도권남동','수도권북동','충청권','경상권','수도권남서','강원권','수도권북서','전라권','수도권남동',
               '수도권북동','충청권','수도권남서','경상권','강원권','수도권북서','전라권','수도권남동','수도권북동','충청권',
               '수도권북서','수도권남동','경상권','강원권','수도권북서','전라권','수도권남동','수도권북동','충청권','수도권남서'],
        '주소': ['서울 강서구','경기 성남시','경기 의정부시','대전 유성구','부산 해운대구','인천 계양구','강원 춘천시','서울 마포구','광주 서구','경기 용인시',
               '서울 강북구','충남 천안시','인천 계양구','대구 달서구','강원 원주시','서울 은평구','전북 전주시','경기 수원시','서울 중랑구','세종시',
               '서울 강서구','경기 성남시','울산 남구','강원 강릉시','서울 양천구','전남 목포시','경기 성남시','서울 성북구','충북 청주시','경남 창원시'],
        '위도': [37.5583,37.3945,37.7388,36.3704,35.1681,37.5376,37.8813,37.5665,35.1595,37.3217,
               37.6398,36.5760,37.5420,35.8285,37.3422,37.6176,35.8242,37.2636,37.5985,36.4801,
               37.5583,37.3945,35.5384,37.7519,37.5172,34.7943,37.4201,37.5894,36.6424,35.2272],
        '경도': [126.7944,127.1116,127.0467,127.3622,129.1303,126.7253,127.7298,126.9018,126.8526,127.1085,
               127.0253,127.1472,126.7389,128.5658,127.9202,126.9227,127.1530,127.0286,127.0927,127.2890,
               126.7944,127.1116,129.3114,128.8761,126.8664,126.3822,127.1266,127.0167,127.4890,128.6811],
        '운영계약시작일': pd.to_datetime([
            '2022-01-15','2022-03-20','2022-05-10','2022-07-05','2022-09-01','2023-01-10','2023-03-15','2023-05-20','2023-07-08','2023-09-12',
            '2024-01-05','2024-03-10','2024-05-15','2024-07-20','2024-09-05','2025-01-08','2025-03-12','2025-05-18','2025-07-22','2025-09-10',
            '2022-02-14','2022-06-18','2022-10-22','2023-02-15','2023-06-20','2024-02-10','2024-06-15','2024-10-20','2025-02-12','2025-06-18']),
        '운영계약종료일': pd.to_datetime([
            '2028-01-14','2028-03-19','2028-05-09','2028-07-04','2028-08-31','2029-01-09','2029-03-14','2029-05-19','2029-07-07','2029-09-11',
            '2030-01-04','2030-03-09','2030-05-14','2030-07-19','2030-09-04','2031-01-07','2031-03-11','2031-05-17','2031-07-21','2031-09-09',
            '2028-02-13','2028-06-17','2028-10-21','2029-02-14','2029-06-19','2030-02-09','2030-06-14','2030-10-19','2031-02-11','2031-06-17']),
    }
    df = pd.DataFrame(sample_data)
    df['운영계약시작일_parsed'] = df['운영계약시작일']
    df['운영계약종료일_parsed'] = df['운영계약종료일']
    df['행번호'] = range(5, 5 + len(df))
    return df


# ──────────────────────────────────────────────
# 대시보드
# ──────────────────────────────────────────────

def show_dashboard(df):
    st.markdown("## 충전기 운영 현황 대시보드")
    st.markdown("### 운영계약 기간 필터")

    df_dates = df.copy()
    df_dates['운영계약시작일_parsed'] = pd.to_datetime(df_dates['운영계약시작일_parsed'], errors='coerce')
    df_dates['운영계약종료일_parsed'] = pd.to_datetime(df_dates['운영계약종료일_parsed'], errors='coerce')
    valid = df_dates.dropna(subset=['운영계약시작일_parsed', '운영계약종료일_parsed'])

    if len(valid) == 0:
        st.warning("유효한 운영계약 날짜 데이터가 없습니다.")
        return

    min_d = valid['운영계약시작일_parsed'].min().date()
    max_d = valid['운영계약종료일_parsed'].max().date()
    def_start = max(min_d, date(2022, 1, 1))
    def_end = min(max_d, date(2028, 1, 1))
    if def_start > def_end:
        def_start, def_end = min_d, max_d

    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        start_date = st.date_input("계약 시작일 (이후)", value=def_start, min_value=min_d, max_value=max_d)
    with c2:
        end_date = st.date_input("계약 종료일 (이전)", value=def_end, min_value=min_d, max_value=max_d)
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("필터 적용", type="primary", use_container_width=True)

    mask = (
        (df_dates['운영계약시작일_parsed'] < pd.Timestamp(end_date)) &
        (df_dates['운영계약종료일_parsed'] >= pd.Timestamp(start_date)) &
        df_dates['운영계약시작일_parsed'].notna() &
        df_dates['운영계약종료일_parsed'].notna()
    )
    filtered_df = df[mask].copy()

    st.info(f"**기간:** {start_date} ~ {end_date} | **충전기:** {len(filtered_df):,}대 / {len(df):,}대 ({len(filtered_df)/len(df)*100:.1f}%)")

    if len(filtered_df) == 0:
        st.warning("해당 기간 데이터 없음")
        return

    st.markdown("---")

    # KPI
    st.markdown("### 주요 지표")
    total = len(filtered_df)
    sites = filtered_df['사이트ID'].nunique()
    regions = filtered_df['권역'].nunique()
    fast = len(filtered_df[filtered_df['모델분류'].str.contains('급속', na=False)])
    fast_pct = fast / total * 100 if total else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 충전기", f"{total:,}대")
    c2.metric("사이트", f"{sites:,}개")
    c3.metric("권역", f"{regions}개")
    c4.metric("급속", f"{fast:,}대", f"{fast_pct:.1f}%")

    # ── 지도 ──
    st.markdown("---")
    st.markdown("### 충전기 위치 지도")

    if '위도' in filtered_df.columns and '경도' in filtered_df.columns:
        valid_coords = filtered_df.dropna(subset=['위도', '경도'])
        if len(valid_coords) > 0:
            sites_map = valid_coords['사이트ID'].nunique()
            st.success(f"{sites_map:,}개 사이트, {len(valid_coords):,}개 충전기 좌표")

            # ★ 집계 데이터 캐시
            grouped_sites = prepare_map_data(filtered_df)
            charger_map, error = create_charger_map(grouped_sites)

            if error:
                st.error(error)
            else:
                st_folium(charger_map, width=1400, height=700)

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("지도 사이트", f"{sites_map:,}개")
                c2.metric("지도 충전기", f"{len(valid_coords):,}대")
                c3.metric("좌표 보유율", f"{len(valid_coords)/total*100:.1f}%")
                c4.metric("좌표 누락", f"{total-len(valid_coords):,}대")

                with st.expander("지도 사용법"):
                    st.markdown("확대/축소(휠), 이동(드래그), 클러스터(숫자 클릭), 마커 클릭(상세+네이버지도)")
        else:
            st.warning("좌표 데이터 없음")
    else:
        st.warning("위도/경도 열 없음")

    # ── 차트 ──
    st.markdown("---")
    st.markdown("### 모델별 현황")

    c1, c2 = st.columns([3, 2])
    with c1:
        mc = filtered_df['모델분류'].value_counts().reset_index()
        mc.columns = ['모델분류', '수량']
        fig = px.bar(mc.head(15), x='수량', y='모델분류', orientation='h', title='모델별 수량 (상위 15)',
                     color='수량', color_continuous_scale='Blues', text='수량')
        fig.update_layout(height=500, showlegend=False)
        fig.update_traces(texttemplate='%{text}', textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        mc['비율'] = (mc['수량'] / mc['수량'].sum() * 100).round(1).astype(str) + '%'
        st.dataframe(mc[['모델분류', '수량', '비율']], hide_index=True, height=450, use_container_width=True)

    st.markdown("---")
    st.markdown("### 권역별 현황")

    c1, c2 = st.columns([2, 3])
    with c1:
        rc = filtered_df['권역'].value_counts().reset_index()
        rc.columns = ['권역', '수량']
        fig = px.pie(rc, values='수량', names='권역', title='권역별 비율', hole=0.4)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(rc, x='권역', y='수량', title='권역별 수량', color='수량', color_continuous_scale='Greens', text='수량')
        fig.update_layout(height=400, showlegend=False)
        fig.update_traces(texttemplate='%{text}', textposition='outside')
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### 히트맵")
    ct = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'])
    top = [m for m in filtered_df['모델분류'].value_counts().head(12).index if m in ct.columns]
    if top:
        fig = px.imshow(ct[top].T, labels=dict(x="권역", y="모델", color="수량"),
                        color_continuous_scale='RdYlGn', aspect="auto", text_auto=True)
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.markdown("### 상세 현황")
    pivot = pd.crosstab(filtered_df['권역'], filtered_df['모델분류'], margins=True)
    st.dataframe(pivot, use_container_width=True, height=400)

    # 다운로드
    st.markdown("---")
    st.markdown("### 다운로드")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("현황표 CSV", pivot.to_csv(encoding='utf-8-sig'),
                           f"현황표_{start_date}_{end_date}.csv", "text/csv", use_container_width=True)
    with c2:
        st.download_button("전체 데이터 CSV", filtered_df.to_csv(index=False, encoding='utf-8-sig'),
                           f"데이터_{start_date}_{end_date}.csv", "text/csv", use_container_width=True)
    with c3:
        txt = f"기간: {start_date}~{end_date}\n총: {total:,}대, 사이트: {sites:,}개, 급속: {fast:,}대 ({fast_pct:.1f}%)"
        st.download_button("요약 TXT", txt, f"요약_{start_date}_{end_date}.txt", "text/plain", use_container_width=True)

    # 품질 체크
    st.markdown("---")
    st.markdown("### 데이터 품질")
    unknown = filtered_df[filtered_df['권역'].isin(['수도권기타', '인천기타', '기타'])]
    c1, c2 = st.columns(2)
    normal = total - len(unknown)
    c1.metric("정상 분류", f"{normal:,}대", f"{normal/total*100:.1f}%")
    c2.metric("미분류", f"{len(unknown):,}대", f"{len(unknown)/total*100:.1f}%")
    if len(unknown) > 0:
        with st.expander("미분류 상세"):
            cols = [c for c in ['주소', '권역', '모델분류', '사이트ID'] if c in unknown.columns]
            st.dataframe(unknown[cols].head(10), hide_index=True)
    else:
        st.success("모두 정상 분류!")


def show_classification_info():
    st.markdown("### 분류 기준표")
    t1, t2, t3 = st.tabs(["급속", "완속", "권역"])
    with t1:
        st.info("AH열 = '급속'일 때")
        st.dataframe({
            "분류": ["급속스필_100","급속스필_50","급속PNE_100","급속PNE_50","급속PNE_200","급속PNE_150",
                   "급속애플망고_200","급속SK_100","급속SK_200","급속코스텔_50","급속중앙제어_50",
                   "급속그린파워_100","급속그린파워_50","급속알박_50"],
            "AG코드": ["S0F1","S0F5","EVQ-(AJ=100)","EVQ-/EV1-(AJ=50)","MAXE","DP15","A01-/AD1-",
                     "Q081/Q101/Q010","Q071/Q102","1Y25/1Y24","1911","1900","19C0","QC50"]
        }, hide_index=True, use_container_width=True)
    with t2:
        st.info("AH열 ≠ '급속'일 때")
        st.dataframe({
            "분류": ["알박구형","알박신형","10kW","신형대","구형대","신형소","F01/PC01","UC01",
                   "스필_7kW","이카플러그","중앙제어_7kW","SK_7kW","3kW","PNE_7kW"],
            "조건": ["NC07","23NA/22NA/24NA/25NA","AD에 3J10","EVL-1C-22CQ","EVL-1C","SBAA",
                   "SBPA/SBOA","SBUA","SVI0","E0C/CP","1907/1912","SC-P","SANA","EVS-/007S"]
        }, hide_index=True, use_container_width=True)
    with t3:
        st.dataframe({
            "권역": ["수도권북서","수도권북동","수도권남동","수도권남서","강원권","충청권","경상권","전라권"],
            "지역": ["고양,부천,김포,파주,은평,마포,서대문,양천,강서","도봉,노원,중랑,강북,의정부,남양주,구리",
                   "강남,서초,송파,강동,성남,용인,수원,평택","구로,금천,영등포,동작,관악,안산,안양+인천주요구",
                   "강원도","충청,세종,대전","경상,부산,대구,울산","전라,광주"]
        }, hide_index=True, use_container_width=True)


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────

def main():
    # ★ 초기 로드: parquet > xlsx > 샘플 (우선순위)
    if 'processed_df' not in st.session_state:
        if os.path.exists(DEFAULT_PARQUET):
            try:
                with st.spinner("default_data.parquet 로드 중..."):
                    st.session_state.processed_df = load_default_parquet(DEFAULT_PARQUET)
                    st.session_state.is_sample_data = False
                    st.session_state.default_file_loaded = True
                    st.session_state.data_source = "parquet"
            except Exception as e:
                st.warning(f"parquet 로드 실패: {e}")
                st.session_state.processed_df = create_sample_data()
                st.session_state.is_sample_data = True
                st.session_state.default_file_loaded = False
        elif os.path.exists(DEFAULT_XLSX):
            try:
                with st.spinner("default_data.xlsx 로드 중..."):
                    st.session_state.processed_df = load_default_xlsx(DEFAULT_XLSX)
                    st.session_state.is_sample_data = False
                    st.session_state.default_file_loaded = True
                    st.session_state.data_source = "xlsx"
            except Exception as e:
                st.warning(f"xlsx 로드 실패: {e}")
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

    tab1, tab2 = st.tabs(["파일 업로드 & 분류", "운영현황 대시보드"])

    with tab1:
        if st.session_state.get('default_file_loaded'):
            src = st.session_state.get('data_source', 'xlsx')
            cnt = len(st.session_state.processed_df)
            if src == 'parquet':
                st.success(f"**default_data.parquet** 고속 로드 완료 ({cnt:,}행)")
            else:
                st.success(f"**default_data.xlsx** 로드 완료 ({cnt:,}행)")
                st.info("💡 `convert_to_parquet.py`를 실행하면 parquet으로 변환되어 **10배 이상** 빠르게 로드됩니다.")
        elif st.session_state.get('is_sample_data'):
            st.info("샘플 데이터 사용 중")

        uploaded = st.file_uploader("엑셀 파일 선택", type=['xlsx', 'xls'])

        if uploaded:
            c1, c2 = st.columns([3, 1])
            with c1:
                st.info(f"**{uploaded.name}**")
            with c2:
                sz = uploaded.size / (1024 * 1024)
                st.metric("크기", f"{sz:.1f} MB" if sz >= 1 else f"{uploaded.size/1024:.1f} KB")

            if st.button("모델분류 시작", type="primary", use_container_width=True):
                title = st.empty()
                title.markdown("### 작업 진행 `0.0초`")
                bar = st.progress(0)
                status = st.empty()
                st.markdown("---")

                result = process_excel_file_with_progress(uploaded.read(), title, bar, status)
                pf, err, cnt, tt, rdf, arc, asc = result

                if err:
                    st.error(err)
                else:
                    st.session_state.processed_df = rdf
                    st.session_state.processed_file = pf
                    st.session_state.is_sample_data = False
                    st.session_state.default_file_loaded = False

                    st.success(f"**{cnt:,}행** 완료 ({format_time(tt)})")
                    ts = get_korea_time().strftime("%Y%m%d_%H%M%S")
                    _, c2, _ = st.columns([1, 2, 1])
                    with c2:
                        st.download_button("결과 다운로드", pf.getvalue(), f"분류결과_{ts}.xlsx",
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           use_container_width=True, type="primary")

        with st.expander("분류 기준"):
            show_classification_info()

    with tab2:
        if st.session_state.processed_df is not None:
            if st.session_state.get('is_sample_data'):
                st.warning("샘플 데이터 — 파일 업로드로 실제 분석 가능")
            show_dashboard(st.session_state.processed_df)
        else:
            st.info("파일을 먼저 업로드하세요.")


if __name__ == "__main__":
    main()
