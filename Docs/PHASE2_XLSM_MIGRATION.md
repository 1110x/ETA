# Phase 2 xlsm 데이터 마이그레이션

## 개요

`Docs/CHUNGHA-김가린.xlsm` 파일의 8개 DATA 시트에 누적된 분석 이력 데이터를 ETA DB로 이동합니다.

**마이그레이션 대상** (8개 시트):
- BOD-DATA → 비용부담금_결과.BOD + 생물학적_산소요구량_시험기록부
- SS-DATA → 비용부담금_결과.SS + 부유물질_시험기록부
- TN-DATA → 비용부담금_결과.T-N + 총질소_시험기록부
- TP-DATA → 비용부담금_결과.T-P + 총인_시험기록부
- N-Hexan-DATA → 비용부담금_결과.N-Hexan + 노말헥산추출물질_시험기록부
- TOC(TCIC)-DATA → 비용부담금_결과.TOC + 총유기탄소_시험기록부
- TOC(NPOC)-DATA → 비용부담금_결과.TOC + 총유기탄소_시험기록부 (TCIC 실패 시 대체)
- Phenols-DATA → 비용부담금_결과.Phenols + 페놀류_시험기록부

## 마이그레이션 로직

### 1. 데이터 추출
각 DATA 시트에서:
- 헤더 행 검색 (채수일/분석일, SN, 업체명, 구분, 결과)
- 모든 데이터 행 읽기
- 날짜 형식 정규화 (YYYY-MM-DD)

### 2. 업체 매칭
xlsm의 업체명 → DB 폐수배출업소.SN 매칭
```sql
SELECT SN, 업체명 FROM `폐수배출업소`
```

### 3. 시험기록부 저장
각 *_시험기록부 테이블에 측정값 INSERT:
```
INSERT INTO `생물학적_산소요구량_시험기록부`
(채수일, SN, 업체명, 구분, 시료량, D1, D2, 희석배수, 결과, 등록일시)
VALUES (...)
```

### 4. 결과값 저장
비용부담금_결과에 UPSERT:
```
UPDATE `비용부담금_결과` SET BOD = '값' WHERE 채수일='일' AND SN='번호'
INSERT INTO `비용부담금_결과` (...) VALUES (...)
```

### 5. 특수 로직

#### TOC (총유기탄소)
**규정**: 공정시험기준 ES 04311.1d
- **TCIC 우선**: TC-IC(가감법) 적용
- **검증**: IC 측정값이 TC의 50% 초과 시 불가
  - 조건 위반 → NPOC(무기탄소 제거) 방법 사용
  
**마이그레이션 로직**:
```
1. TOC(TCIC)-DATA 읽기 → TC, IC 검사
2. IC ≤ TC × 50% ? 
   ├─ Yes: TCIC 값 그대로 사용
   └─ No: TOC(NPOC)-DATA에서 NPOC 값 사용
```

#### Phenols (페놀류)
**규정**: 공정시험기준 ES 04365.1e
- **방법 선택**: 측정값 기준
  - 직접법 (Direct): 0.05 mg/L ≤ 값 ≤ 0.5 mg/L
  - 추출법 (Extract): 값 < 0.05 mg/L
- **페놀류_시험기록부**: 직접법/추출법 컬럼 분리

**마이그레이션 로직**:
```
1. Phenols-DATA에서 방법(직접/추출) 확인
2. 측정값에 따라:
   ├─ ≥ 0.05: 직접법 컬럼에 저장
   └─ < 0.05: 추출법 컬럼에 저장
```

## 마이그레이션 실행 흐름

### Phase 2 자동 실행
로그인 시 Step 6에서 `XlsmDataMigration.ExecutePhase2()` 자동 실행:

```
로그인
  ↓
Step 5: Phase 1 마이그레이션 (RENAME + DROP)
  ↓
Step 6: Phase 2 마이그레이션 (xlsm 데이터)  ← 현재 위치
  │
  ├─ [단계 1] 데이터 추출
  │   └─ 8개 DATA 시트 읽기 (ClosedXML)
  │
  ├─ [단계 2] 업체 매칭
  │   └─ 폐수배출업소 SN 매핑
  │
  ├─ [단계 3] 데이터 로드
  │   ├─ WasteSampleService.UpsertBodData() 호출
  │   ├─ WasteSampleService.UpsertSimpleData() 호출
  │   ├─ WasteSampleService.UpsertTocData() 호출
  │   └─ WasteSampleService.UpsertUvvisData() 호출
  │
  ├─ [단계 4] 특수 로직
  │   ├─ TOC: TCIC vs NPOC 선택
  │   └─ Phenols: 직접법 vs 추출법 선택
  │
  └─ ✓ Phase 2 완료
  ↓
Step 7: 견적 테이블 초기화
  ↓
Step 8: 처리시설 측정결과 자동 생성
  ↓
메인 페이지 표시
```

## 파일 구조

### xlsm 파일 분석
```
Docs/CHUNGHA-김가린.xlsm (약 200KB)
├─ TOC(NPOC) [Sheet 1]
├─ BOD [Sheet 2]
├─ SS [Sheet 3]
├─ TN [Sheet 4]
├─ TOC(TCIC) [Sheet 5]
├─ TP [Sheet 6]
├─ Phenols [Sheet 7]
├─ ...
├─ BOD-DATA [Sheet 11] ← 마이그레이션 대상
├─ SS-DATA [Sheet 12]
├─ TOC(TCIC)-DATA [Sheet 13]
├─ TOC(NPOC)-DATA [Sheet 14]
├─ Phenols-DATA [Sheet 15]
├─ N-Hexan-DATA [Sheet 16]
├─ TN-DATA [Sheet 17]
└─ TP-DATA [Sheet 18]
```

### 시트 구조 예시 (BOD-DATA)
```
행1: 채수일  | SN | 업체명 | 구분 | ... | 결과
행2: 2024-01-15 | 001 | ㈜회사A | 여수 | ... | 120
행3: 2024-01-16 | 002 | ㈜회사B | 여수 | ... | 95
...
```

## 마이그레이션 대상 테이블 구조

### 비용부담금_결과 (결과 저장)
```
Id | 채수일 | 구분 | 순서 | SN | 업체명 | 관리번호 | 
BOD | TOC | SS | T-N | T-P | N-Hexan | Phenols | 비고 | 확인자
```

### *_시험기록부 (측정값 저장)
예: 생물학적_산소요구량_시험기록부
```
Id | 채수일 | SN | 업체명 | 구분 | 시료량 | D1 | D2 | 희석배수 | 결과 | 등록일시
```

## 예상 마이그레이션 수량

각 시트별 대략적인 데이터 행:
- BOD-DATA: ~200행
- SS-DATA: ~150행
- TN-DATA: ~100행
- TP-DATA: ~100행
- N-Hexan-DATA: ~80행
- TOC(TCIC)-DATA: ~120행
- TOC(NPOC)-DATA: ~30행 (대체용)
- Phenols-DATA: ~80행

**총계**: 약 900-1000행

## 검증 항목

로그인 후 Phase 2 실행 완료 후:

1. **DB 데이터 확인**
   ```sql
   SELECT COUNT(*) FROM `비용부담금_결과` WHERE BOD != '';
   SELECT COUNT(*) FROM `생물학적_산소요구량_시험기록부`;
   ```

2. **시험기록부 조회**
   - PAGE2 → 분석결과입력
   - 각 항목별 시험기록부 데이터 표시 확인

3. **결과값 확인**
   - 비용부담금_결과 테이블에서 BOD, TOC, SS 등 값 확인
   - 일부 샘플 데이터의 xlsm과 DB 값 비교

## 문제 해결

### xlsm 파일을 찾을 수 없음
```
⚠️  파일 없음: Docs/CHUNGHA-김가린.xlsm
```
→ 파일 경로 확인, 파일 존재 여부 확인

### 데이터가 로드되지 않음
```
✗ 마이그레이션 실패: ...
```
→ 로그에서 상세 에러 메시지 확인 (Logs 폴더)
→ DB 테이블 존재 확인
→ 폐수배출업소 SN 매핑 확인

### 특수 로직 미적용
TOC/Phenols 특수 로직이 적용되지 않으면:
→ WasteSampleService.UpsertTocData() 로직 확인
→ WasteSampleService.UpsertUvvisData() 로직 확인

## 다음 단계

Phase 2 완료 후:

1. **데이터 검증**
   - xlsm과 DB의 일부 샘플 비교
   - 누락된 데이터 확인

2. **xlsm 파일 보관**
   - 마이그레이션 완료 후 원본 파일 백업
   - 필요시 다시 마이그레이션 가능

3. **시험기록부 정리**
   - 레거시 형식의 시험기록부 제거 (필요시)
   - 새 형식만 사용하도록 정리

## 참고 자료

- [의뢰및결과](의뢰및결과.md) - 테이블 생성 규칙
- [PHASE1_MIGRATION_STATUS.md](PHASE1_MIGRATION_STATUS.md) - Phase 1 상태
- Services/SERVICE2/WasteSampleService.cs - Upsert 메서드
