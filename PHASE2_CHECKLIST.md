# Phase 2 xlsm 데이터 마이그레이션 체크리스트

## ✅ 완료된 항목

### 코드 구현
- [x] XlsmDataMigration.cs 작성 (데이터 추출 로직)
- [x] Login.axaml.cs에 Phase 2 연결 (Step 6)
- [x] ClosedXML으로 xlsm 파일 읽기
- [x] 폐수배출업소 SN 매핑 로직

### 문서화
- [x] PHASE2_XLSM_MIGRATION.md (마이그레이션 가이드)
- [x] PHASE2_CHECKLIST.md (실행 체크리스트)

### 빌드
- [x] dotnet build 성공 (0 errors)

## 📋 Phase 2 실행 단계

### 1단계: 사전 준비
- [ ] xlsm 파일 존재 확인: `Docs/CHUNGHA-김가린.xlsm`
- [ ] Phase 1 마이그레이션 완료 확인
  - DB: 수질분석센터_결과, 비용부담금_결과, 처리시설_결과 테이블 존재
  - Code: 테이블명 통일 완료

### 2단계: 애플리케이션 실행
- [ ] `dotnet run --project ETA.csproj`
- [ ] 로그인 화면 표시 확인

### 3단계: 로그인
- [ ] 직원ID/패스워드 입력
- [ ] 진행률 표시 확인:
  - "DB 마이그레이션 중... (Phase 1)" (60-65%)
  - "Phase 2 데이터 마이그레이션 중..." (75-80%)
  - "견적 데이터 준비 중..." (82-88%)

### 4단계: Phase 2 로그 확인
로그인 중 다음 메시지 확인:

```
[단계 1] 데이터 추출
  📋 BOD-DATA
  ✓ 약 200행 읽음
  📋 SS-DATA
  ✓ 약 150행 읽음
  ...
  ✓ 총 약 900-1000행 추출

[단계 2] 업체명 → SN 매칭
  ✓ 약 50개 업체 SN 로드

[단계 3] 데이터 로드 → DB
  ✓ Phase 2 준비 완료
```

### 5단계: 기능 검증
로그인 후:

- [ ] 메인 페이지 정상 표시
- [ ] PAGE2 (분석결과입력) 접근 가능
- [ ] 각 탭에서 데이터 조회 가능

### 6단계: DB 데이터 확인 (SQL)

#### 비용부담금_결과 테이블
```sql
-- 항목별 데이터 입력 현황
SELECT 'BOD' as 항목, COUNT(*) as 행수 FROM `비용부담금_결과` WHERE BOD != ''
UNION
SELECT 'SS', COUNT(*) FROM `비용부담금_결과` WHERE SS != ''
UNION
SELECT 'T-N', COUNT(*) FROM `비용부담금_결과` WHERE `T-N` != ''
UNION
SELECT 'T-P', COUNT(*) FROM `비용부담금_결과` WHERE `T-P` != ''
UNION
SELECT 'N-Hexan', COUNT(*) FROM `비용부담금_결과` WHERE `N-Hexan` != ''
UNION
SELECT 'TOC', COUNT(*) FROM `비용부담금_결과` WHERE `TOC` != ''
UNION
SELECT 'Phenols', COUNT(*) FROM `비용부담금_결과` WHERE Phenols != '';
```

**예상 결과** (대략):
| 항목 | 행수 |
|------|------|
| BOD | 200+ |
| SS | 150+ |
| T-N | 100+ |
| T-P | 100+ |
| N-Hexan | 80+ |
| TOC | 150+ |
| Phenols | 80+ |

#### 시험기록부 테이블
```sql
-- 각 시험기록부 행 수
SELECT COUNT(*) FROM `생물학적_산소요구량_시험기록부`;
SELECT COUNT(*) FROM `부유물질_시험기록부`;
SELECT COUNT(*) FROM `총질소_시험기록부`;
SELECT COUNT(*) FROM `총인_시험기록부`;
SELECT COUNT(*) FROM `노말헥산추출물질_시험기록부`;
SELECT COUNT(*) FROM `총유기탄소_시험기록부`;
SELECT COUNT(*) FROM `페놀류_시험기록부`;
```

**예상 결과**: 각각 100-200+ 행

### 7단계: 특수 로직 검증

#### TOC 검증 (TCIC vs NPOC)
```sql
-- TOC 값이 입력된 비용부담금_결과 행
SELECT 채수일, SN, 업체명, `TOC` 
FROM `비용부담금_결과` 
WHERE `TOC` != ''
LIMIT 10;
```

확인 항목:
- [ ] TOC 값이 정상적으로 입력됨
- [ ] TCIC/NPOC 전환이 규정에 따라 적용됨 (필요시)

#### Phenols 검증 (직접법 vs 추출법)
```sql
-- 페놀류 시험기록부 상세
SELECT 채수일, SN, 업체명, 직접법, 추출법
FROM `페놀류_시험기록부`
WHERE 직접법 != '' OR 추출법 != ''
LIMIT 10;
```

확인 항목:
- [ ] 직접법/추출법 값이 올바른 컬럼에 저장됨
- [ ] 측정값에 따라 방법이 선택됨 (≥0.05→직접, <0.05→추출)

## ⚠️ 문제 발생 시

### Phase 2 로그에 에러가 표시되는 경우

**증상**:
```
⚠️  Phase 2 마이그레이션 실패: ...
```

**진단**:
1. 로그 파일 확인: `Logs/XlsmMigration.log` (생성된 경우)
2. DB 연결 상태 확인
3. xlsm 파일 존재 및 형식 확인

**해결**:
- xlsm 파일이 손상된 경우: 원본 파일 다시 준비
- DB 연결 오류: appsettings.json의 DB 설정 확인
- 권한 부족: DB 사용자 권한 확인

### 데이터가 로드되지 않은 경우

**증상**:
```
Phase 2 마이그레이션이 완료되었으나 데이터가 없음
```

**진단**:
```sql
-- 로그 확인
SELECT * FROM `비용부담금_결과` WHERE BOD != '';
```

**원인별 해결**:
- xlsm 파일이 비어있는 경우: 파일 내용 확인
- SN 매칭 실패: 폐수배출업소 테이블의 업체명과 비교
- 헤더 행 파싱 실패: xlsm 시트 구조 확인 (첫 5행)

### xlsm 파일을 찾을 수 없음

**증상**:
```
⊘ 파일 없음: Docs/CHUNGHA-김가린.xlsm
```

**해결**:
1. 파일 위치 확인: `Docs/` 디렉토리
2. 파일명 정확성 확인 (대소문자, 특수문자)
3. 파일이 없으면 준비: Docs 폴더에 CHUNGHA-김가린.xlsm 파일 복사

## 📊 마이그레이션 진행도

```
Phase 1 (완료):     ████████████████ 100%
  └─ DB 테이블명 통일
  └─ 코드 변경
  └─ 자동 실행 (로그인 시)

Phase 2 (진행 중):  ████████░░░░░░░░  50%
  ├─ 구조 분석        ✓ 완료
  ├─ 코드 작성        ✓ 완료
  ├─ 빌드 검증        ✓ 완료
  └─ 실행 테스트      ⏳ 대기 중

Phase 3 (예정):     ░░░░░░░░░░░░░░░░   0%
  └─ xlsm 데이터 검증
  └─ 결과 보고
```

## 📅 타임라인

- **Phase 1 완료**: 2026-04-15
- **Phase 2 준비**: 2026-04-16 (현재)
- **Phase 2 실행**: 다음 로그인 시
- **Phase 3 (결과 검증)**: Phase 2 후

## 📞 지원

문제 발생 시 체크:
1. [PHASE2_XLSM_MIGRATION.md](Docs/PHASE2_XLSM_MIGRATION.md) - 상세 가이드
2. 애플리케이션 로그 (`Logs/` 폴더)
3. DB 로그

---

**Phase 2 준비**: ✅ 완료  
**다음 단계**: 애플리케이션 실행 → 로그인 → 자동 마이그레이션
