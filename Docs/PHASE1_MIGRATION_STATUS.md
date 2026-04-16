# Phase 1 DB 마이그레이션 상태

## 개요
ETA의 분석결과 테이블 명명 규칙을 통일하고 레거시 *_DATA 테이블을 제거하는 Phase 1 마이그레이션이 완료되었습니다.

## 완료된 작업

### 1. 테이블명 통일 규칙
`{소스}_결과` 패턴으로 통일:
- `분석의뢰및결과` → `수질분석센터_결과`
- `폐수의뢰및결과` → `비용부담금_결과`  
- `처리시설_측정결과` → `처리시설_결과`

### 2. 코드 변경 (219개 참조 업데이트)

#### 주요 파일 (20개)
**SERVICE1**
- [AnalysisRequestService.cs](../Services/SERVICE1/AnalysisRequestService.cs) - 수질분석센터_결과 참조 변경
- [AnalysisRecordService.cs](../Services/SERVICE1/AnalysisRecordService.cs)
- [CompanyRenameService.cs](../Services/SERVICE1/CompanyRenameService.cs)
- [ContractService.cs](../Services/SERVICE1/ContractService.cs)
- [OrderRequestService.cs](../Services/SERVICE1/OrderRequestService.cs)
- [TestReportService.cs](../Services/SERVICE1/TestReportService.cs)

**SERVICE2**
- [WasteSampleService.cs](../Services/SERVICE2/WasteSampleService.cs) - 비용부담금_결과 참조 변경 (219개)
- [WasteDataService.cs](../Services/SERVICE2/WasteDataService.cs)
- [WasteRequestService.cs](../Services/SERVICE2/WasteRequestService.cs)
- [WasteTestReportService.cs](../Services/SERVICE2/WasteTestReportService.cs)
- [FacilityResultService.cs](../Services/SERVICE2/FacilityResultService.cs) - 처리시설_결과 참조 변경
- [FacilityDbMigration.cs](../Services/SERVICE2/FacilityDbMigration.cs) - 레거시 CREATE TABLE 제거

**SERVICE3**
- [MyTaskService.cs](../Services/SERVICE3/MyTaskService.cs)

**Views**
- [MainPage.axaml.cs](../Views/MainPage.axaml.cs)
- [EcotoxicityPage.axaml.cs](../Views/Pages/PAGE1/EcotoxicityPage.axaml.cs)
- [WasteAnalysisInputPage.axaml.cs](../Views/Pages/PAGE2/WasteAnalysisInputPage.axaml.cs)
- [MyTaskPage.axaml.cs](../Views/Pages/PAGE3/MyTaskPage.axaml.cs)
- [DbMigrationMappingPanel.axaml.cs](../Views/Pages/PAGE1/DbMigrationMappingPanel.axaml.cs)
- [AnalysisRequestDetailPanel.axaml.cs](../Views/Pages/PAGE1/AnalysisRequestDetailPanel.axaml.cs)

### 3. 레거시 테이블 코드 제거

[FacilityDbMigration.cs:378-507](../Services/SERVICE2/FacilityDbMigration.cs#L378-L507):
- BOD_DATA, SS_DATA, NHexan_DATA, TN_DATA, TP_DATA, Phenols_DATA, TOC_TCIC_DATA, TOC_NPOC_DATA 생성 코드 제거
- 주석으로 "레거시 *_DATA 테이블은 DROP됨 (더 이상 사용하지 않음)" 표기

### 4. 마이그레이션 실행 계획

#### Phase 1 자동 실행 로직
- 새로운 `DbPhase1Migration` 클래스 생성: [Services/Common/DbPhase1Migration.cs](../Services/Common/DbPhase1Migration.cs)
- 로그인 시 자동 실행 ([Views/Login.axaml.cs:213-238](../Views/Login.axaml.cs#L213-L238)):
  1. RENAME 테이블 (3개)
  2. DROP 테이블 (8개)

#### 마이그레이션 순서
1. **DB 변경** (로그인 시 자동 실행)
   - RENAME TABLE (테이블명 통일)
   - DROP TABLE (레거시 *_DATA)

2. **코드 변경** (완료)
   - 하드코딩된 테이블명 일괄 변경 (219개)
   - FacilityDbMigration.cs 정리

3. **xlsm 데이터 마이그레이션** (다음 단계)
   - Python 스크립트로 8개 DATA 시트 파싱
   - 각 항목별 결과 테이블에 데이터 INSERT

4. **검증**
   - 새로운 테이블명에서 데이터 조회 확인
   - 분석결과 입력 페이지 정상 작동 확인

## xlsm 파일 준비 상황

[Docs/CHUNGHA-김가린.xlsm](../Docs/CHUNGHA-김가린.xlsm):
- 총 39개 시트
- 8개 DATA 시트 (마이그레이션 대상):
  - BOD-DATA
  - SS-DATA
  - TN-DATA
  - TP-DATA
  - TOC(NPOC)-DATA
  - TOC(TCIC)-DATA
  - Phenols-DATA
  - N-Hexan-DATA
- 문자 인코딩: "(주)" → "㈜" 변환 완료 (115개 → 245개)

## Phase 2 (xlsm 마이그레이션) 예정

마이그레이션 규칙:

| 항목 | 원본 | 시험기록부 | 결과 테이블 | 특수 로직 |
|------|------|---------|----------|---------|
| BOD | BOD-DATA | 생물학적_산소요구량_시험기록부 | 비용부담금_결과.BOD | - |
| SS | SS-DATA | 부유물질_시험기록부 | 비용부담금_결과.SS | - |
| T-N | TN-DATA | 총질소_시험기록부 | 비용부담금_결과.T-N | - |
| T-P | TP-DATA | 총인_시험기록부 | 비용부담금_결과.T-P | - |
| N-Hexan | N-Hexan-DATA | 노말헥산추출물질_시험기록부 | 비용부담금_결과.N-Hexan | - |
| TOC | TOC(TCIC/NPOC)-DATA | 총유기탄소_시험기록부 | 비용부담금_결과.TOC | TCIC 우선, IC>TC×50% → NPOC 대체 |
| Phenols | Phenols-DATA | 페놀류_시험기록부 | 비용부담금_결과.Phenols | 직접법≥0.05 → 직접법, <0.05 → 추출법 |

## 빌드 상태
✓ 빌드 성공 (0개 오류, 221개 경고)

## 다음 단계
1. 애플리케이션 실행 → 로그인 시 Phase 1 마이그레이션 자동 실행
2. 마이그레이션 결과 확인 (DB 테이블명 변경 확인)
3. Phase 2 xlsm 데이터 마이그레이션 구현
