# TOC 마이그레이션 삽질 기록

## 현재 상태
- TCIC: 386건 저장 완료
- NPOC: 0건 (미완료)
- 비용부담금_결과.TOC: 3993건 업데이트 (BOD 완료 후 잘못 실행됨 — 모든 항목 완료 후 한번에 보내야 함)

---

## 삽질 목록

### 1. 한자 析 반복 삽입
- 문제: `분析일` 대신 `분析일`(析=U+6790 한자) 가 코드에 계속 들어감
- 원인: Claude 자동완성이 한자로 생성
- 해결: 바이트 레벨 치환 `bytes.fromhex('e69e90')` → `bytes.fromhex('ec849d')`
- **규칙: 코드/SQL/주석에 析 절대 금지**

### 2. 비용부담금_결과 TOC 조기 업데이트
- 문제: 모든 항목 마이그레이션 완료 전에 update_waste_result 호출
- 원인: BOD 방식 그대로 복사
- **규칙: 비용부담금_결과 UPDATE는 전체 항목(BOD/TOC/SS/TN/TP/NHexan/Phenols) 완료 후 별도 스크립트로**

### 3. 총_유기탄소_시험기록부 잘못 DROP
- 문제: 이미 7043건 저장된 테이블을 실수로 DROP
- 원인: 테이블 구조 변경 시 기존 데이터 백업 없이 DROP
- 결과: 처음부터 재작업 필요

### 4. 테이블 구조 혼란 (단일 vs 분리)
- 처음: TCIC/NPOC 한 테이블에 저장 (총_유기탄소_시험기록부)
- 중간: 유니크키 (분析일, SN) 충돌 → NPOC가 TCIC 행을 못 찾아 Duplicate 오류
- 결정: 두 테이블 분리 (총_유기탄소_TCIC_시험기록부 / 총_유기탄소_NPOC_시험기록부)

### 5. 시트명 오류
- 문제: `TOC(TCIC)-DATA` → `TOC(TCIC)`로 잘못 변경
- 실제 시트명: `TOC(TCIC)-DATA`, `TOC(NPOC)-DATA`
- 결과: 0건 저장

### 6. f-string + pymysql %s 충돌
- 문제: `f"""..."""` 안에 `%s` 사용 → Python이 먼저 % 포맷팅 시도
- 오류: `not enough arguments for format string` / `not all arguments converted`
- 해결: f-string 대신 문자열 연결(`"..." + col_date + "..."`) 사용

### 7. INSERT %s 개수 불일치
- TCIC: 컬럼 38개인데 43개 %s → 수동 계산 후 수정
- NPOC: 컬럼 26개인데 32개 %s → 수동 계산 후 수정

### 8. read_only 모드에서 r[0].row 접근 불가
- 문제: `openpyxl read_only=True`에서 `EmptyCell.row` 속성 없음
- 해결: `enumerate(ws.iter_rows(...), start=2)` 로 변경

### 9. TCIC 컬럼 구조 혼란
- 초기: ST01~05_mgL_TCIC / ST01~05_abs_TCIC 만 있음
- 실제: TC 검정곡선 따로, IC 검정곡선 따로 존재
- 추가 컬럼: ST01~05_mgL_IC, ST01~05_abs_IC, 기울기_IC, 절편_IC, R2_IC

---

## 현재 TOC 테이블 구조

### 총_유기탄소_TCIC_시험기록부
- id, 분析일, SN, 업체명, 시료명, 구분, 소스구분, 비고, 등록일시
- TCcon, ICcon, TCAU, ICAU (샘플 블록값)
- 농도_TCIC, 희석배수_TCIC, 검량선_a_TCIC, 결과_TCIC
- 기울기_TCIC, 절편_TCIC, R2_TCIC
- ST01~05_mgL_TCIC, ST01~05_abs_TCIC (TC 검정곡선)
- ST01~05_mgL_IC, ST01~05_abs_IC, 기울기_IC, 절편_IC, R2_IC (IC 검정곡선)
- UNIQUE KEY uk_tcic (분析일, SN)

### 총_유기탄소_NPOC_시험기록부
- id, 분析일, SN, 업체명, 시료명, 구분, 소스구분, 비고, 등록일시
- 시료량_NPOC, 흡광도_NPOC, 희석배수_NPOC, 검량선_a_NPOC, 농도_NPOC, 결과_NPOC
- 기울기_NPOC, 절편_NPOC, R2_NPOC
- ST01~05_mgL_NPOC, ST01~05_abs_NPOC
- UNIQUE KEY uk_npoc (분析일, SN)

---

## 엑셀 구조 (확인 완료)

### TOC(TCIC)-DATA 시트
- col1: 날짜
- col3~7: TC ST-01~05 mgL
- col8~12: TC ST-01~05 abs
- col13~15: TC 기울기/절편/R2
- col16~20: IC ST-01~05 mgL
- col21~25: IC ST-01~05 abs
- col26~28: IC 기울기/절편/R2
- col30~: 샘플 블록 (8컬럼 간격)
  - +0=시료명, +1=시료량(미저장), +2=TCcon, +3=ICcon
  - +4=농도, +5=희석배수, +6=결과, +7=SN

### TOC(NPOC)-DATA 시트
- col1: 날짜
- col3~7: ST-01~05 mgL
- col8~12: ST-01~05 abs
- col13~15: 기울기/절편/R2
- col30~: 샘플 블록 (8컬럼 간격)
  - +0=시료명, +1=시료량, +2=흡광도, +3=희석배수
  - +4=검량선_a, +5=농도, +6=결과, +7=SN

---

## 현재 migrate_toc.py 상태
- f-string 제거, 문자열 연결로 SQL 구성
- TCIC 386건 저장 완료 (정상 동작 확인)
- NPOC load_npoc() 실행됐으나 0건 — 원인 미확인
- TCAU/ICAU는 엑셀에 없는 값이므로 컬럼만 존재, 저장 안함
