#!/usr/bin/env python3
"""
Phase 2 xlsm 데이터 마이그레이션
Docs/CHUNGHA-김가린.xlsm의 8개 DATA 시트를 DB로 이동

마이그레이션 대상:
  - BOD-DATA → 비용부담금_결과.BOD + 생물학적_산소요구량_시험기록부
  - SS-DATA → 비용부담금_결과.SS + 부유물질_시험기록부
  - TN-DATA → 비용부담금_결과.T-N + 총질소_시험기록부
  - TP-DATA → 비용부담금_결과.T-P + 총인_시험기록부
  - N-Hexan-DATA → 비용부담금_결과.N-Hexan + 노말헥산추출물질_시험기록부
  - TOC(TCIC)-DATA → 비용부담금_결과.TOC + 총유기탄소_시험기록부
  - TOC(NPOC)-DATA → 비용부담금_결과.TOC + 총유기탄소_시험기록부
  - Phenols-DATA → 비용부담금_결과.Phenols + 페놀류_시험기록부

특수 로직:
  - TOC: TCIC 우선, IC > TC×50% → NPOC 대체
  - Phenols: 직접법 ≥ 0.05 → 직접법, < 0.05 → 추출법
"""

import json
import sqlite3
import sys
from pathlib import Path
from datetime import datetime
from zipfile import ZipFile
from xml.etree import ElementTree as ET

def load_db_config():
    """appsettings.json에서 DB 설정 읽기"""
    config_path = Path(__file__).parent.parent / "appsettings.json"
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    return config.get("MariaDb", {})

def excel_date_to_iso(excel_date):
    """엑셀 날짜 시리얼을 ISO 형식으로 변환"""
    if isinstance(excel_date, str):
        # 이미 문자열이면 그대로 반환
        if len(excel_date) == 10 and excel_date[4] == '-':
            return excel_date
        return ""

    if isinstance(excel_date, (int, float)):
        # 엑셀 날짜 시리얼: 1900-01-01이 1
        # Python: 1970-01-01이 0
        try:
            # 1900-01-01 = 2, 1900-02-28 = 59, 1900-02-29는 존재하지 않음 (버그)
            days = int(excel_date)
            # 엑셀은 1900-02-29를 잘못 표현 (실제로는 존재하지 않음)
            if days > 59:
                days -= 1
            base_date = datetime(1899, 12, 30)
            result_date = base_date.replace(year=base_date.year).fromordinal(
                base_date.toordinal() + days
            )
            return result_date.strftime('%Y-%m-%d')
        except:
            return ""

    return ""

def extract_sheets_from_xlsm(xlsm_path):
    """xlsm 파일에서 시트 데이터 추출

    xlsm은 ZIP 형식이고, xl/worksheets/sheet*.xml에 데이터가 있음
    """
    sheets_data = {}
    target_sheets = {
        'BOD-DATA': {'col': 'BOD', 'table': '생물학적_산소요구량_시험기록부'},
        'SS-DATA': {'col': 'SS', 'table': '부유물질_시험기록부'},
        'TN-DATA': {'col': 'T-N', 'table': '총질소_시험기록부'},
        'TP-DATA': {'col': 'T-P', 'table': '총인_시험기록부'},
        'N-Hexan-DATA': {'col': 'N-Hexan', 'table': '노말헥산추출물질_시험기록부'},
        'TOC(TCIC)-DATA': {'col': 'TOC', 'table': '총유기탄소_시험기록부', 'method': 'TCIC'},
        'TOC(NPOC)-DATA': {'col': 'TOC', 'table': '총유기탄소_시험기록부', 'method': 'NPOC'},
        'Phenols-DATA': {'col': 'Phenols', 'table': '페놀류_시험기록부'},
    }

    try:
        with ZipFile(xlsm_path, 'r') as z:
            # workbook.xml에서 시트 목록 읽기
            workbook_xml = z.read('xl/workbook.xml').decode('utf-8')

            print(f"✓ {xlsm_path} 열기 성공")
            print(f"📋 시트 목록 분석 중...")

            # 시트명과 rid 매핑
            import re
            sheet_refs = re.findall(
                r'<sheet[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"',
                workbook_xml
            )

            sheet_map = dict(sheet_refs)
            rels_xml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')

            # rid → sheetN.xml 매핑
            rid_to_file = {}
            for rid, file in re.findall(
                r'<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"',
                rels_xml
            ):
                rid_to_file[rid] = file

            # 각 대상 시트 읽기
            for sheet_name in target_sheets:
                if sheet_name not in sheet_map:
                    print(f"⊘ {sheet_name} 없음")
                    continue

                rid = sheet_map[sheet_name]
                file = rid_to_file.get(rid, '')

                if not file:
                    print(f"⊘ {sheet_name} 파일 매핑 없음")
                    continue

                sheet_path = f"xl/{file}"
                try:
                    sheet_xml = z.read(sheet_path).decode('utf-8')
                    sheets_data[sheet_name] = _parse_sheet_xml(sheet_xml)
                    print(f"✓ {sheet_name}: {len(sheets_data[sheet_name])} 행")
                except Exception as e:
                    print(f"✗ {sheet_name} 파싱 실패: {e}")

        return sheets_data
    except Exception as e:
        print(f"✗ xlsm 읽기 실패: {e}")
        return {}

def _parse_sheet_xml(sheet_xml):
    """XML 시트에서 데이터 행 추출"""
    rows = []
    try:
        # XML 파싱 (네임스페이스 무시)
        root = ET.fromstring(sheet_xml)
        ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

        shared_strings = {}  # 필요시 사용

        for row_elem in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
            row_data = []
            for cell in row_elem.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                value = ""
                v_elem = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                if v_elem is not None and v_elem.text:
                    value = v_elem.text
                row_data.append(value)

            if any(row_data):  # 비어있지 않은 행만
                rows.append(row_data)
    except Exception as e:
        print(f"  경고: XML 파싱 실패 - {e}")

    return rows

def print_summary(xlsm_path):
    """xlsm 파일 요약 정보 출력"""
    print(f"\n{'='*70}")
    print(f"Phase 2 xlsm 데이터 마이그레이션")
    print(f"{'='*70}")
    print(f"파일: {xlsm_path}")
    print(f"파일 크기: {xlsm_path.stat().st_size / 1024:.1f} KB")
    print(f"수정 일시: {datetime.fromtimestamp(xlsm_path.stat().st_mtime)}")

    # 시트 목록 확인
    try:
        with ZipFile(xlsm_path, 'r') as z:
            workbook_xml = z.read('xl/workbook.xml').decode('utf-8')
            import re
            sheet_names = re.findall(r'<sheet[^>]*name="([^"]*)"', workbook_xml)

            target_sheets = ['BOD-DATA', 'SS-DATA', 'TN-DATA', 'TP-DATA',
                           'TOC(TCIC)-DATA', 'TOC(NPOC)-DATA', 'Phenols-DATA', 'N-Hexan-DATA']

            print(f"\n대상 시트 ({len(target_sheets)}개):")
            for sheet in target_sheets:
                status = "✓" if sheet in sheet_names else "✗"
                print(f"  {status} {sheet}")

            print(f"\n전체 시트 ({len(sheet_names)}개):")
            for i, name in enumerate(sheet_names, 1):
                if 'DATA' in name or name in target_sheets:
                    print(f"  [{i:2d}] ✓ {name}")
    except Exception as e:
        print(f"  (시트 목록 읽기 실패: {e})")

def main():
    print("\n[Phase 2] xlsm 데이터 마이그레이션 준비")
    print("="*70)

    xlsm_path = Path("Docs/CHUNGHA-김가린.xlsm")

    if not xlsm_path.exists():
        print(f"✗ 파일 없음: {xlsm_path}")
        return 1

    print_summary(xlsm_path)

    print(f"\n{'='*70}")
    print("Phase 2 마이그레이션 단계:")
    print("="*70)
    print("""
[1] 데이터 추출
    - xlsm에서 8개 DATA 시트 읽기
    - 각 시트의 행/열 구조 분석

[2] 데이터 매칭
    - xlsm 업체명 → DB 폐수배출업소(SN) 매칭
    - 중복 제거 및 검증

[3] 시험기록부 저장
    - *_시험기록부 테이블에 측정값 INSERT

[4] 결과값 저장
    - 비용부담금_결과 테이블에 최종값 UPSERT

[5] 특수 로직
    - TOC: TCIC 우선, IC > TC×50% → NPOC 대체
    - Phenols: 직접법 ≥ 0.05 → 직접법, < 0.05 → 추출법

호출 방식: C# WasteSampleService.UpsertBodData() 등을 호출하거나
직접 Python에서 MySQLdb 연결로 INSERT/UPSERT 수행

⚠️  현재 단계: 데이터 구조 분석 완료
📋 다음 단계: 실제 데이터 로드 구현 필요
    """)

    print(f"\n{'='*70}")
    print("✅ Phase 2 준비 완료")
    print("="*70)

    return 0

if __name__ == "__main__":
    sys.exit(main())
