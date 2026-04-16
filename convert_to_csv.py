#!/usr/bin/env python3
"""
처리시설 데이터 JSON → CSV 변환 스크립트
"""

import json
import csv
import re
from pathlib import Path

# siteCd → 시설명 매핑
FACILITY_MAP = {
    'PJT1020': '중흥',
    'PJT1021': '월내',
    'PJT1114': '4단계',
    'PJT1022': '율촌',
    'PJT1298': '세풍',
    'SITE053': '월내/중흥(마스터)',
    'SITE065': '4단계(마스터)',
    'SITE066': '율촌/세풍(마스터)'
}

def parse_json_from_txt(txt_file):
    """txt 파일에서 JSON 블록 추출"""
    with open(txt_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # JSON 블록 찾기 ([ 부터 ] 까지)
    pattern = r'\[\s*\{.*?\}\s*\]'
    matches = re.findall(pattern, content, re.DOTALL)

    all_data = []
    for match in matches:
        try:
            data = json.loads(match)
            all_data.extend(data if isinstance(data, list) else [data])
        except:
            pass

    return all_data

def get_facility_name(siteCd):
    """siteCd로 시설명 조회"""
    return FACILITY_MAP.get(siteCd, siteCd)

def extract_ocr_data(txt_file):
    """OCR 데이터만 추출"""
    with open(txt_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # [OCR 데이터] 섹션만 추출
    ocr_pattern = r'\[OCR 데이터\][^\[]*?(\[\s*\{.*?\}\s*\])'
    matches = re.findall(ocr_pattern, content, re.DOTALL)

    all_records = []
    for match in matches:
        try:
            data = json.loads(match)
            all_records.extend(data if isinstance(data, list) else [data])
        except:
            pass

    return all_records

def convert_to_csv(txt_file, csv_file):
    """JSON 데이터를 CSV로 변환"""

    print(f"txt 파일 읽는 중: {txt_file}")
    records = extract_ocr_data(txt_file)
    print(f"추출된 레코드: {len(records)}개")

    # CSV 헤더
    headers = [
        '시설명',
        '시료명',
        '채취일자',
        '분류',
        'BOD',
        'TOC',
        'SS',
        'T-N',
        'T-P',
        '대장균군',
        '입력자',
        '입력일시'
    ]

    with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()

        for record in records:
            # null 값 필터링 (모든 항목이 null이면 스킵)
            has_data = any([
                record.get('bodD1'), record.get('bodD2'),
                record.get('tocMl'),
                record.get('ssMl'),
                record.get('tnMl'),
                record.get('tpMl'),
                record.get('coliA'), record.get('coliB')
            ])

            if not has_data and record.get('sampleCategory') in ['POS00', 'POS99']:
                continue  # 마스터 시료 스킵

            siteCd = record.get('siteCd', '')
            facility = get_facility_name(siteCd)

            # BOD 값 (D1과 D2 평균, 또는 bodNo)
            bod_d1 = record.get('bodD1')
            bod_d2 = record.get('bodD2')
            bod_value = ''
            if bod_d1 or bod_d2:
                if bod_d1 and bod_d2:
                    bod_value = f"{(float(bod_d1) + float(bod_d2)) / 2:.2f}"
                elif bod_d1:
                    bod_value = str(bod_d1)
                else:
                    bod_value = str(bod_d2)

            # TOC 값
            toc_ml = record.get('tocMl', '')
            toc_p = record.get('tocP', '')
            toc_value = f"{toc_ml}(×{toc_p})" if toc_ml else ''

            # SS 값
            ss_ml = record.get('ssMl', '')
            ss_p = record.get('ssP', '')
            ss_before = record.get('ssMgBefore', '')
            ss_after = record.get('ssMgAfter', '')
            ss_value = ''
            if ss_before and ss_after:
                ss_value = f"{ss_before}-{ss_after}"
            elif ss_ml:
                ss_value = f"{ss_ml}(×{ss_p})"

            # T-N 값
            tn_ml = record.get('tnMl', '')
            tn_p = record.get('tnP', '')
            tn_value = f"{tn_ml}(×{tn_p})" if tn_ml else ''

            # T-P 값
            tp_ml = record.get('tpMl', '')
            tp_p = record.get('tpP', '')
            tp_value = f"{tp_ml}(×{tp_p})" if tp_ml else ''

            # 대장균군 값
            coli_a = record.get('coliA', '')
            coli_b = record.get('coliB', '')
            coli_value = ''
            if coli_a or coli_b:
                coli_value = f"{coli_a}/{coli_b}" if coli_a and coli_b else (coli_a or coli_b)

            # 시료명
            sample_name = record.get('sampleCategoryNm', '')
            if record.get('siteNm') and record.get('siteNm') != '-':
                sample_name = f"{record.get('siteNm')} - {sample_name}"

            row = {
                '시설명': facility,
                '시료명': sample_name,
                '채취일자': record.get('inputDate', ''),
                '분류': record.get('sampleCategory', ''),
                'BOD': bod_value,
                'TOC': toc_value,
                'SS': ss_value,
                'T-N': tn_value,
                'T-P': tp_value,
                '대장균군': coli_value,
                '입력자': record.get('createUser', ''),
                '입력일시': record.get('createTime', '')
            }

            writer.writerow(row)

    print(f"✓ CSV 저장 완료: {csv_file}")

if __name__ == '__main__':
    txt_file = Path.home() / 'Documents/ETA/Data/facility_data_last_week.txt'
    csv_file = Path.home() / 'Documents/ETA/Data/facility_data_last_week.csv'

    convert_to_csv(str(txt_file), str(csv_file))
