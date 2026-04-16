# Phase 1 마이그레이션 검증 가이드

## 예상 실행 흐름

다음에 애플리케이션을 실행할 때 로그인 화면에서:

```
로그인 → 데이터베이스 초기화...
  ⏳ 사진 동기화 중...        [==== 40%]
  ⏳ 사진 동기화 완료         [==== 56%]
  ⏳ DB 마이그레이션 중... (Phase 1)  [====== 60%]
    ✓ 분석의뢰및결과 → 수질분석센터_결과
    ✓ 폐수의뢰및결과 → 비용부담금_결과
    ✓ 처리시설_측정결과 → 처리시설_결과
    ✓ DROP BOD_DATA
    ✓ DROP SS_DATA
    ✓ DROP NHexan_DATA
    ✓ DROP TN_DATA
    ✓ DROP TP_DATA
    ✓ DROP Phenols_DATA
    ✓ DROP TOC_TCIC_DATA
    ✓ DROP TOC_NPOC_DATA
  ✓ DB 마이그레이션 완료 (Phase 1) [===== 65%]
  ⏳ 테이블 초기화 중...     [======= 68%]
  ✓ 테이블 준비 완료        [======== 72%]
  ⏳ 메인 페이지 로드 중...
  → 메인 페이지 표시
```

## 검증 방법

### 1. 로그인 후 DB 연결 확인
```csharp
// SQL 실행 (예: MySQL Workbench)
SHOW TABLES LIKE '%결과';
```

**예상 결과**:
```
수질분석센터_결과
비용부담금_결과
처리시설_결과
```

### 2. 레거시 테이블 삭제 확인
```sql
SHOW TABLES LIKE '%_DATA';
```

**예상 결과**: (비어있음 — 모든 *_DATA 테이블 제거됨)

### 3. 분석결과 입력 페이지 확인
- PAGE2 → 분석결과입력
- 각 탭에서 데이터 조회/입력 정상 작동 확인

### 4. 기존 데이터 확인
```sql
SELECT COUNT(*) FROM `비용부담금_결과`;
SELECT COUNT(*) FROM `수질분석센터_결과`;
SELECT COUNT(*) FROM `처리시설_결과`;
```

**예상 결과**: 기존 데이터는 유지되어야 함 (RENAME만 수행)

## Phase 1 롤백 (필요 시)

Phase 1이 실패한 경우:

```sql
-- 테이블명을 원래대로 복원 (RENAME 실패 시)
RENAME TABLE `수질분석센터_결과` TO `분석의뢰및결과`;
RENAME TABLE `비용부담금_결과` TO `폐수의뢰및결과`;
RENAME TABLE `처리시설_결과` TO `처리시설_측정결과`;

-- 코드 변경도 이전 커밋으로 되돌려야 함
git revert <commit-hash>
```

## Phase 2 준비 (xlsm 마이그레이션)

Phase 1이 성공적으로 완료되면:

1. **DB 상태 확인**
   - 새 테이블명으로 데이터 정상 접근 가능

2. **xlsm 파일 확인**
   - [Docs/CHUNGHA-김가린.xlsm](CHUNGHA-김가린.xlsm)의 8개 DATA 시트 확인
   - 각 시트의 컬럼 구조 파악

3. **Python 마이그레이션 스크립트 작성**
   - openpyxl로 xlsm 파싱
   - 날짜 시리얼 변환, 업체명 매칭
   - 각 결과 테이블에 INSERT

4. **TOC/Phenols 특수 로직 구현**
   - TOC: TCIC 데이터 검증 (IC ≤ TC×50% 확인)
   - Phenols: 직접법 결과값에 따라 추출법 사용 여부 결정

## 예상 마이그레이션 시간

- Phase 1 (RENAME + DROP): **< 1초** (로그인 화면에서 자동 실행)
- Phase 2 (xlsm 데이터): **수분** (파싱 + DB 삽입)

## 주의 사항

⚠️ **Phase 1은 로그인 시 자동 실행됩니다**
- 취소/스킵 불가
- 실패 시 에러 메시지 표시 후 로그에 기록됨
- 재시도는 다음 로그인 시 자동 수행

⚠️ **데이터 손실 없음**
- RENAME만 수행 (데이터는 유지)
- DROP되는 테이블은 모두 레거시 (*_DATA) — 더 이상 사용 중인 테이블 아님
