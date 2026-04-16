#!/usr/bin/env python3
"""
처리시설 데이터 다운로드 스크립트
저번주(2026-04-09 ~ 2026-04-15) 데이터를 API에서 수집해서 txt로 저장
"""

import requests
import json
from datetime import datetime, timedelta
import os

# API 기본 설정
BASE_URL = "https://rewater.wayble.eco/stp/api/subnote"
SESSION_COOKIE = "stpsession=OTczYzQ0NTMtM2ZhNC00YWVmLWI2NDYtM2QzODBiNWQ4YWM5; _ga_889WWPX1W1=GS2.1.s1776344282$o3$g1$t1776344289$j53$l0$h0; _ga=GA1.1.452346581.1775044184"

HEADERS = {
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'X-Requested-With': 'XMLHttpRequest',
    'Cookie': SESSION_COOKIE,
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15'
}

# 처리시설 설정
FACILITIES = {
    'SITE053': {'name': '월내/중흥', 'sites': ['PJT1020', 'PJT1021']},
    'SITE065': {'name': '4단계', 'sites': ['PJT1114']},
    'SITE066': {'name': '율촌/세풍', 'sites': ['PJT1022', 'PJT1298']}
}

# 저번주 날짜 범위 (2026-04-09 ~ 2026-04-15)
START_DATE = datetime(2026, 4, 9)
END_DATE = datetime(2026, 4, 15)

def get_date_range():
    """날짜 범위 반환"""
    dates = []
    current = START_DATE
    while current <= END_DATE:
        dates.append(current.strftime('%Y%m%d'))
        current += timedelta(days=1)
    return dates

def download_data(office_cd, input_date):
    """API에서 데이터 다운로드"""
    data = {
        'ocr': None,
        'mst': None,
        'ss': None,
        'tn': None,
        'tp': None,
        'toc': None,
        'coli': None
    }

    try:
        # 1. OCR 데이터
        resp = requests.post(
            f"{BASE_URL}/selectSubNoteOcrList",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date}
        )
        if resp.status_code == 200:
            data['ocr'] = resp.json().get('body', {}).get('subNoteList', [])

        # 2. MST (BOD) 데이터
        resp = requests.post(
            f"{BASE_URL}/selectSubNoteMst",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date}
        )
        if resp.status_code == 200:
            data['mst'] = resp.json().get('body', {}).get('mstList', {})

        # 3. SS 데이터
        resp = requests.post(
            f"{BASE_URL}/selectSsList",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date}
        )
        if resp.status_code == 200:
            data['ss'] = resp.json().get('body', {}).get('subNoteSsList', [])

        # 4. T-N 데이터
        resp = requests.post(
            f"{BASE_URL}/selectListTntp",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date, 'tntpCategory': 'TN'}
        )
        if resp.status_code == 200:
            data['tn'] = resp.json().get('body', {}).get('subNoteAddList', [])

        # 5. T-P 데이터
        resp = requests.post(
            f"{BASE_URL}/selectListTntp",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date, 'tntpCategory': 'TP'}
        )
        if resp.status_code == 200:
            data['tp'] = resp.json().get('body', {}).get('subNoteAddList', [])

        # 6. TOC 데이터
        resp = requests.post(
            f"{BASE_URL}/selectListToc",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date, 'tntpCategory': 'TOC'}
        )
        if resp.status_code == 200:
            data['toc'] = resp.json().get('body', {}).get('subNoteAddList', [])

        # 7. Coli 데이터
        resp = requests.post(
            f"{BASE_URL}/selectColiList",
            headers=HEADERS,
            data={'officeCd': office_cd, 'inputDate': input_date}
        )
        if resp.status_code == 200:
            data['coli'] = resp.json().get('body', {}).get('subNoteColiList', [])

    except Exception as e:
        print(f"❌ {office_cd} {input_date} 다운로드 실패: {e}")

    return data

def save_to_txt(filename):
    """데이터를 txt 파일로 저장"""
    dates = get_date_range()

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("=" * 100 + "\n")
        f.write(f"처리시설 데이터 다운로드 (저번주: 2026-04-09 ~ 2026-04-15)\n")
        f.write(f"다운로드 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 100 + "\n\n")

        for office_cd, facility in FACILITIES.items():
            f.write(f"\n{'='*100}\n")
            f.write(f"처리시설: {facility['name']} ({office_cd})\n")
            f.write(f"{'='*100}\n")

            for input_date in dates:
                f.write(f"\n{'─'*100}\n")
                f.write(f"날짜: {input_date}\n")
                f.write(f"{'─'*100}\n")

                print(f"다운로드 중... {office_cd} {input_date}")
                data = download_data(office_cd, input_date)

                # JSON 형태로 저장
                f.write(f"\n[OCR 데이터] ({len(data['ocr'])} 건)\n")
                f.write(json.dumps(data['ocr'], ensure_ascii=False, indent=2) + "\n")

                f.write(f"\n[마스터 데이터]\n")
                f.write(json.dumps(data['mst'], ensure_ascii=False, indent=2) + "\n")

                f.write(f"\n[SS 데이터] ({len(data['ss'])} 건)\n")
                f.write(json.dumps(data['ss'], ensure_ascii=False, indent=2) + "\n")

                f.write(f"\n[T-N 데이터] ({len(data['tn'])} 건)\n")
                f.write(json.dumps(data['tn'], ensure_ascii=False, indent=2) + "\n")

                f.write(f"\n[T-P 데이터] ({len(data['tp'])} 건)\n")
                f.write(json.dumps(data['tp'], ensure_ascii=False, indent=2) + "\n")

                f.write(f"\n[TOC 데이터] ({len(data['toc'])} 건)\n")
                f.write(json.dumps(data['toc'], ensure_ascii=False, indent=2) + "\n")

                f.write(f"\n[대장균군 데이터] ({len(data['coli'])} 건)\n")
                f.write(json.dumps(data['coli'], ensure_ascii=False, indent=2) + "\n")

if __name__ == '__main__':
    output_file = os.path.expanduser('~/Documents/ETA/Data/facility_data_last_week.txt')
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    print(f"처리시설 데이터 다운로드 시작...")
    print(f"저번주: 2026-04-09 ~ 2026-04-15")
    print(f"저장 위치: {output_file}\n")

    save_to_txt(output_file)

    print(f"\n✓ 완료! 파일: {output_file}")
