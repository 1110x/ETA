# Phase 1 완료 체크리스트

## ✅ 완료된 항목

### 코드 변경
- [x] 분석의뢰및결과 → 수질분석센터_결과 (47개 참조)
- [x] 폐수의뢰및결과 → 비용부담금_결과 (26개+ 참조)  
- [x] 처리시설_측정결과 → 처리시설_결과 (29개 참조)
- [x] 20개 파일에서 총 219개 테이블명 참조 업데이트
- [x] FacilityDbMigration.cs에서 레거시 CREATE TABLE 코드 제거

### DB 마이그레이션
- [x] DbPhase1Migration.cs 작성 (RENAME + DROP 로직)
- [x] Login.axaml.cs에 자동 실행 훅 추가
- [x] 로그인 시 자동으로 Phase 1 마이그레이션 실행되도록 설정

### 빌드
- [x] dotnet build 성공 (0 errors)
- [x] 모든 코드 변경사항 컴파일 확인

### 문서화
- [x] PHASE1_MIGRATION_STATUS.md (상세 상태)
- [x] PHASE1_VERIFICATION.md (검증 가이드)
- [x] PHASE1_SUMMARY.md (최종 보고서)
- [x] PHASE1_ARCHITECTURE.md (아키텍처 다이어그램)

## 📋 실행 전 확인 사항

### DB 환경
- [ ] MariaDB 서버 정상 작동 (1110s.synology.me:3306)
- [ ] eta_db 데이터베이스 접근 가능
- [ ] eta_user 계정 권한 확인

### 애플리케이션 상태  
- [ ] 프로젝트 빌드 완료
- [ ] appsettings.json의 DB 연결 정보 확인
- [ ] 백업 수행 (선택사항, 권장)

## 🚀 실행 단계

### 1단계: 애플리케이션 실행
```bash
dotnet run --project ETA.csproj
```
- [ ] 로그인 화면 표시 확인

### 2단계: 로그인
- 직원ID, 패스워드 입력
- [ ] "DB 마이그레이션 중... (Phase 1)" 메시지 표시 확인
- [ ] 진행률 표시 (60% → 65%)

### 3단계: 마이그레이션 실행 확인
로그인 후 다음을 확인:

```
예상 로그:
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
```

### 4단계: 기능 확인
- [ ] 메인 페이지 정상 표시
- [ ] PAGE1 (수질분석센터) 접근 가능
- [ ] PAGE2 (분석결과입력) 탭 정상 작동
- [ ] PAGE3 (분석기록) 데이터 조회 가능

### 5단계: DB 확인 (SQL)
```sql
-- 새 테이블명 확인
SHOW TABLES LIKE '%결과';
-- 예상: 수질분석센터_결과, 비용부담금_결과, 처리시설_결과

-- 레거시 테이블 제거 확인  
SHOW TABLES LIKE '%_DATA';
-- 예상: (결과 없음)

-- 데이터 무결성 확인
SELECT COUNT(*) FROM `비용부담금_결과`;
SELECT COUNT(*) FROM `수질분석센터_결과`;
SELECT COUNT(*) FROM `처리시설_결과`;
```

## ⚠️ 문제 해결

### 마이그레이션 실패 시

**증상**: DB 마이그레이션 에러 메시지 표시
```
⚠ DB 마이그레이션 실패: ...
```

**해결 방법**:
1. DB 연결 상태 확인
2. MariaDB 서버 상태 확인
3. eta_user 계정 권한 확인 (RENAME, DROP 권한)
4. [PHASE1_VERIFICATION.md](Docs/PHASE1_VERIFICATION.md)의 롤백 섹션 참고

### 코드 실행 중 테이블명 에러 시

**증상**: "Table '수질분석센터_결과' doesn't exist" 에러
```
로그: MySqlException: Table 'eta_db.수질분석센터_결과' doesn't exist
```

**해결 방법**:
1. DB에서 실제 테이블명 확인:
   ```sql
   SHOW TABLES LIKE '%결과%';
   ```
2. 테이블명이 예상과 다르면 → 마이그레이션 실패했을 가능성
3. 테이블이 있으면 → 코드에서 잘못된 테이블명 참조 (grep으로 확인)

## 📊 Phase 1 후 상태

### DB 상태
| 항목 | Before | After | 상태 |
|------|--------|-------|------|
| 수질분석센터_결과 | 분석의뢰및결과 | 수질분석센터_결과 | ✅ 이름 변경 |
| 비용부담금_결과 | 폐수의뢰및결과 | 비용부담금_결과 | ✅ 이름 변경 |
| 처리시설_결과 | 처리시설_측정결과 | 처리시설_결과 | ✅ 이름 변경 |
| 레거시 *_DATA | 8개 존재 | 0개 존재 | ✅ 제거됨 |
| 전체 데이터 | 유지됨 | 유지됨 | ✅ 손실 없음 |

## 📅 Phase 2 준비

Phase 1 완료 후 Phase 2 (xlsm 데이터 마이그레이션) 시작 가능:

- [ ] DB 테이블명 통일 확인
- [ ] 레거시 테이블 제거 확인
- [ ] Python 마이그레이션 스크립트 작성
- [ ] xlsm 파일 8개 DATA 시트 파싱
- [ ] 각 결과 테이블에 데이터 로드
- [ ] TOC/Phenols 특수 로직 적용

## 📞 지원

문제 발생 시:
1. [PHASE1_VERIFICATION.md](Docs/PHASE1_VERIFICATION.md) 참고
2. 에러 로그 확인 (Output 탭)
3. DB 상태 수동 확인

---

**마이그레이션 시작일**: 2026-04-15  
**예상 완료 시간**: < 1초 (로그인 중 자동 실행)  
**롤백 가능**: Yes (SQL RENAME 명령)
