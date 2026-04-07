---
name: 폐수배출업소 분析항목 규칙
description: 여수/율촌 구분별 분析항목 구성 및 TOC 1법/2법 처리 규칙
type: project
---

## 분析항목 목록

| 항목명 | 코드 | 비고 |
|--------|------|------|
| BOD | BOD | |
| TOC(TC-IC) | TOC1 | TOC 1법 — 기본 사용 |
| TOC(NPOC) | TOC2 | TOC 2법 — 1법 불만족 시 대체 |
| SS | SS | |
| T-N | TN | |
| T-P | TP | |
| Phenols | Phenols | 여수만 |
| N-Hexan | NHexan | 여수만 |

## 구분별 적용 항목

- **여수**: 전 항목 (BOD, TOC1, TOC2, SS, T-N, T-P, Phenols, N-Hexan)
- **율촌**: Phenols, N-Hexan 제외 (BOD, TOC1, TOC2, SS, T-N, T-P)

## TOC 1법/2법 처리 규칙

- TOC(TC-IC) = 1법, TOC(NPOC) = 2법
- 성적서에는 둘 다 "TOC"로 통칭
- **1법 결과가 만족하면** → TC-IC 값 사용
- **1법 불만족 시** → NPOC 값으로 대체 출력
- 입력 화면에서는 TC-IC, NPOC 둘 다 입력 가능하게

**How to apply:** 분析결과입력 화면 항목 구성, 시험성적서 TOC 컬럼 출력 로직에 반영
