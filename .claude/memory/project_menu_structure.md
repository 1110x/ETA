---
name: ETA 메뉴 구조 및 버튼 배치
description: 메뉴별 클릭 핸들러, 서브메뉴 버튼(BT1~BT7) 매핑
type: project
---

최근 메뉴 재배치 반영 (v10~v11 기준).

**Why:** 메뉴 구조가 자주 바뀌므로 현재 상태를 기록

**How to apply:** 메뉴 추가/이동 시 MainPage.axaml과 MainPage.axaml.cs BT1~BT7 핸들러 동시 수정 필요

### 현재 메뉴 구조
- **수질분석센터**: 직원정보, 계약업체, 견적/의뢰서, 시험성적서, DB Migration
- **비용부담금 관리**: 폐수배출업소, 자료 조회, 분석의뢰리스트, 분석결과입력, Uipath ERP, 시험기록부, 자료검토, 테스트
- **처리시설**: 시험기록부, RDP for Zero4
- **일반업무관리**: 보수요청, 물품구매, 계약업체 채수, 출장관리, 휴일근무
- **권한관리**: 센터명칭(WaterQualityNameReconcile), 폐수배출업소명칭(WasteNameReconcile)

### 분석의뢰리스트 특이사항
- Show1: WasteSampleListPage (날짜 트리)
- Show2: 날짜 선택 시 채수 목록
- Show4: 폐수배출업소 트리 (클릭 → Show2에 업체 추가, 중복방지)
