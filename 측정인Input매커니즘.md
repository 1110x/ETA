# 측정인 Input 매커니즘

> 측정인.kr 사이트 자동화에 관한 기술 문서  
> 최종 수정: 2026-03-29

---

## 목차
1. [개요](#1-개요)
2. [사이트 구조 — 주요 HTML 요소](#2-사이트-구조--주요-html-요소)
3. [핵심 원칙: 클릭 동작](#3-핵심-원칙-클릭-동작)
4. [CDP 연결 구조](#4-cdp-연결-구조)
5. [입력 단계별 상세](#5-입력-단계별-상세)
6. [데이터 매핑 (ETA → 측정인)](#6-데이터-매핑-eta--측정인)
7. [이벤트 패턴 레퍼런스 (VBA 원본)](#7-이벤트-패턴-레퍼런스-vba-원본)
8. [자료TO측정인 (분석결과 입력)](#8-자료to측정인-분석결과-입력)
9. [트러블슈팅](#9-트러블슈팅)

---

## 1. 개요

```
[ETA 앱]  ──CDP WebSocket──▶  [Edge 브라우저]  ──▶  [측정인.kr]
  (C#)       port 9222         (--remote-debugging-port)      (jQuery + select2)
```

**ETA 앱**은 Chrome DevTools Protocol (CDP)을 통해 Edge 브라우저에 JavaScript를 `Runtime.evaluate`로 주입하여 측정인.kr의 폼 요소를 제어한다.

### 관련 파일

| 파일 | 역할 |
|------|------|
| `Views/Pages/AnalysisRequestListPanel.axaml.cs` | 의뢰계획 작성 (측정인 전송 버튼) |
| `Views/MeasurerLoginWindow.axaml.cs` | 로그인, 계약DB 스크래핑, 분석DB 스크래핑 |
| `Views/DataToMeasurerWindow.cs` | 시험성적서 → 분석결과 입력 (PAGE4) |
| `Services/MeasurerService.cs` | 채취지점·분석항목 DB (SQLite) |
| `Services/AgentService.cs` | 인력 고유번호 조회 |

---

## 2. 사이트 구조 — 주요 HTML 요소

### 계획 작성 모달 (field_water.do → 계획 작성)

| Element ID | 유형 | 용도 | select2 |
|------------|------|------|---------|
| `add_meas_cont_no` | `<select>` | **근거 계약** 선택 | ✅ Yes |
| `cmb_emis_cmpy_plc_no` | `<select>` | **현장(사업장)** 선택 | ❌ Native |
| `add_emis_fac_no` | `<select>` | **채취지점** 선택 | ❌ Native |
| `add_env_psic_name` | `<input>` | **환경기술인** (담당자) | - |
| `add_meas_start_dt` | `<input>` | **분석시작일자** | - |
| `add_meas_item` | `<select multiple>` | **분석(의뢰)항목** 다중선택 | ✅ Yes |
| `add_emp_id` | `<select multiple>` | **측정인력** 다중선택 | ❌ Native |
| `add_meas_purpose` | `<select>` | **측정목적** (CF/SELF) | ❌ Native |
| `add_meas_car_no` | `<select>` | **측정차량** | ❌ Native |
| `add_meas_equip_no` | `<select multiple>` | **측정장비** 다중선택 | ❌ Native |
| `addFieldPlan` | `<button>` | 계획 작성 모달 열기 | - |
| `insertFieldPlanBtn` | `<button>` | 계획 저장 | - |

### 분석자료 입력 (PAGE4)

| XPath / ID | 유형 | 용도 |
|------------|------|------|
| `//tr[@role='row' and contains(@class,'rg-data-row')]` | `<tr>` | 분석항목 행 |
| `.//td[4]` | `<td>` | 항목명 표시 |
| `.//td[5]` ~ `.//td[16]` | `<td>` | 법적기준, 분석결과, 방법, 장비, 담당자, 일자 등 |
| `#$$_rg_editor151` | `<input>` | 공통 입력 위치 (텍스트) |
| `#$$_rg_editor648` | `<input>` | 공통 입력 위치 (날짜) |
| `#$$_rg_editor655` | `<input>` | 공통 입력 위치 (시료용량) |
| `//div[@role='option']//span[contains(text(),'...')]` | `<div>` | 드롭다운 리스트 선택 |
| `#set-conc_group_no-1` | `<button>` | 동시분석그룹 설정 |

### 기타

| Element ID | 유형 | 용도 |
|------------|------|------|
| `samp_vesl_desc` | `<input>` | 채취용기 설명 (PAGE2) |
| `samp_vesl_pe` / `samp_vesl_pe_cnt` | `<input>` | PE용기 크기/수량 |
| `samp_vesl_gls_brown` / `samp_vesl_gls_brown_cnt` | `<input>` | 갈색유리 크기/수량 |
| `meas_start_dt` / `meas_end_dt` | `<input>` | 시료채취 시작/종료 시간 |
| `edit_meas_loc_desc_1` | `<input>` | 채취지점 설명 |
| `fst_meas_loc_field` | `<button>` | 채취지점 현장 클릭 |

---

## 3. 핵심 원칙: 클릭 동작

### ❌ 작동하지 않는 방식

```javascript
// 직접 값 설정 — 측정인 사이트에서 인식하지 못함
element.value = 'some value';
element.dispatchEvent(new Event('input'));
```

### ✅ 작동하는 방식

```javascript
// 1. native <select>: option.selected + change 이벤트
option.selected = true;
select.dispatchEvent(new Event('change', { bubbles: true }));

// 2. select2 <select>: option선택 + jQuery trigger
option.selected = true;
option.click();  // 클릭 동작 필수
select.dispatchEvent(new Event('change'));
$('#elementId').trigger({
    type: 'select2:select',
    params: { data: { id: value, text: optionText } }
});
$('#elementId').trigger('change');

// 3. 텍스트 <input>: value설정 + blur (VBA의 SendKeys 대응)
input.value = '';         // Clear
input.value = 'text';    // 입력
input.blur();             // blur로 값 확정

// 4. 날짜 <input>: 같은 패턴 (input.value + blur)
var script = "var input = document.getElementById('add_meas_start_dt');"
           + "input.value = '2026-03-29';"
           + "input.blur();";

// 5. 다중 선택 <select multiple>: 기존 해제 → 각 option.selected
while (sel.selectedOptions.length > 0)
    sel.selectedOptions[0].selected = false;
sel.dispatchEvent(new Event('change', {bubbles: true}));
// 그 다음 개별 option 선택
option1.selected = true;
option2.selected = true;
sel.dispatchEvent(new Event('change', {bubbles: true}));
```

### 연쇄 로딩 (Cascade)

계약 선택 → **AJAX로 현장 목록 로드** → 현장 선택 → **AJAX로 채취지점 로드**

각 선택 후 **반드시 폴링 대기**가 필요하다:

```csharp
// C#에서 400ms 간격으로 폴링 (최대 5초)
for (int w = 0; w < 5000; w += 400)
{
    await Task.Delay(400, cts.Token);
    string optCount = ExtractCdpValue(await CdpEvalAsync(socket, @"(function(){
        var sel = document.getElementById('cmb_emis_cmpy_plc_no');
        if (!sel || sel.options.length <= 1) return '0';
        return String(sel.options.length);
    })()", cts.Token));
    if (int.TryParse(optCount, out int count) && count > 1) break;
}
```

---

## 4. CDP 연결 구조

### 접속 흐름

```
1. HTTP GET http://127.0.0.1:9222/json → 탭 목록 JSON
2. URL에 "sijeong" 또는 "field_water" 포함하는 page 탭 찾기
3. webSocketDebuggerUrl로 WebSocket 연결
4. Runtime.evaluate로 JS 실행
```

### CDP 메시지 형식

```json
{
    "id": 1,
    "method": "Runtime.evaluate",
    "params": {
        "expression": "(function(){ ... })()",
        "returnByValue": true
    }
}
```

### C# 헬퍼

```csharp
// AnalysisRequestListPanel: 일회성 소켓
private static async Task<string> CdpEvalAsync(
    ClientWebSocket ws, string expression, CancellationToken ct)

// MeasurerLoginWindow: 영속 소켓 (_ws 필드)
private Task<string> Evaluate(string script) => EvaluateAsync(_ws!, script);
```

### 소켓 관리

| 위치 | 수명 | 용도 |
|------|------|------|
| `MeasurerLoginWindow._ws` | 로그인~창 닫기 | 로그인, 스크래핑, 의뢰계획 자동작성 |
| `TryInjectRequestDataAsync` 내 `socket` | 메서드 내 using | 의뢰 전송 (독립적) |

---

## 5. 입력 단계별 상세

### Step 0: 데이터 준비 + 모달 열기

```
window.__etaRequestData = [레코드배열];
document.getElementById('addFieldPlan').click();
```

모달 대기: `add_meas_cont_no` 옵션 존재 확인 (최대 8초 폴링)

### Step A: 근거 계약 선택 — `add_meas_cont_no` (select2)

```
ETA DB: MeasurerService.FindBestContract(업체명, 의뢰사업장, 약칭)
  → matchedContract.계약번호 ↔ 측정인 옵션 value 매칭 (점수제)

점수 기준:
  - 계약번호 일치: +500
  - 업체명 포함:   +220
  - 사업장 포함:   +180
```

JS 패턴 (VBA CommandButton5 기반):
```javascript
var opts = sel.getElementsByTagName('option');
for (var i = 0; i < opts.length; i++) {
    if (opts[i].value == contractValue) {
        opts[i].selected = true;
        opts[i].click();       // ← 핵심: click() 호출
        break;
    }
}
sel.dispatchEvent(new Event('change'));
// select2 trigger 추가
$('#add_meas_cont_no').trigger({
    type: 'select2:select',
    params: { data: { id: contractValue, text: optionText } }
});
$('#add_meas_cont_no').trigger('change');
```

→ **AJAX 대기**: `cmb_emis_cmpy_plc_no` 옵션 > 1개 될 때까지 폴링

### Step B: 현장 선택 — `cmb_emis_cmpy_plc_no` (native select)

```
ETA 데이터: records[0].workSite (의뢰사업장)
  → 옵션 텍스트에 포함되는지 검색 → 첫 매칭 선택
  → 매칭 없으면 두 번째 옵션 fallback
```

```javascript
sel.value = matchedOption.value;
sel.dispatchEvent(new Event('change', { bubbles: true }));
```

→ **AJAX 대기**: `add_emis_fac_no` 옵션 > 1개 될 때까지 폴링

### Step C: 환경기술인 — `add_env_psic_name` (text input)

```
ETA 데이터: records[0].manager (담당자)
```

```javascript
el.value = '';               // Clear(VBA)
el.value = '담당자이름';     // SendKeys(VBA)
el.dispatchEvent(new Event('input', { bubbles: true }));
el.dispatchEvent(new Event('change', { bubbles: true }));
el.blur();
```

### Step D: 분석시작일자 — `add_meas_start_dt` (date input)

```
ETA 데이터: records[0].sampleDate 또는 date (채취일자/의뢰일)
  → "YYYY-MM-DD" 형식 변환
```

```javascript
// VBA: input.value = '2026-03-29'; input.blur();
el.value = '2026-03-29';
el.blur();
```

### Step E: 분석항목 — `add_meas_item` (select2, multiple)

```
ETA 데이터: records → analytes[] (분석항목명)
  → MeasurerService.GetAllAnalysisItems()에서 코드값 조회
  → 항목명 ↔ 측정인_분석항목.항목명 매칭 → 코드값
```

```javascript
// VBA CommandButton6 패턴:
// 1) 기존 선택 해제
Array.from(sel.options).forEach(opt => opt.selected = false);
// 2) 각 항목 선택
codes.forEach(code => {
    Array.from(sel.options).forEach(opt => {
        if (opt.value === code) opt.selected = true;
    });
});
// 3) change 이벤트 + select2 닫기
sel.dispatchEvent(new Event('change'));
$('#add_meas_item').data('select2').close();
```

### Step F: 측정인력 — `add_emp_id` (select, multiple)

```
ETA 데이터: 시료채취자-1, 시료채취자-2
  → AgentService.GetAllItems()에서 성명[..3] 매칭 → 측정인고유번호
```

```javascript
// VBA CommandButton9 패턴:
// 1) 기존 선택 전부 해제
while (sel.selectedOptions.length > 0)
    sel.selectedOptions[0].selected = false;
sel.dispatchEvent(new Event('change', { bubbles: true }));
// 2) 새 인력 선택
option1.selected = true;
option2.selected = true;
sel.dispatchEvent(new Event('change', { bubbles: true }));
```

### Step G: 측정목적 — `add_meas_purpose` (native select)

```javascript
// VBA CommandButton9 패턴:
var opt = sel.querySelector('option[value="CF"]')   // 기관측정
       || sel.querySelector('option[value="SELF"]'); // 자가측정
if (opt) opt.selected = true;
sel.dispatchEvent(new Event('change', { bubbles: true }));
```

---

## 6. 데이터 매핑 (ETA → 측정인)

### 계약 매핑

```
ETA                          측정인
─────────────────────────    ──────────────────────────
업체명 + 의뢰사업장 + 약칭  → MeasurerService.FindBestContract()
  → 계약번호                 → add_meas_cont_no option[value]
  → 업체명                   → 옵션 텍스트 검색 (부분일치)
  → 채취지점명               → 옵션 텍스트 검색 (부분일치)
```

### 분석항목 매핑

```
ETA                          측정인
─────────────────────────    ──────────────────────────
분석항목명 (ex: "BOD")       → 측정인_분석항목 테이블
  → 항목명 매칭               → 코드값 (add_meas_item option[value])
  
테이블 스키마:
  Id | 분야 | 항목구분 | 항목명 | 코드값 | select2id
```

### 인력 매핑

```
ETA                          측정인
─────────────────────────    ──────────────────────────
시료채취자-1 (ex: "김지은")  → Agent 테이블
  → 성명 앞 3글자 매칭        → 측정인고유번호 (add_emp_id option[value])
  
Agent 테이블 컬럼:
  성명 | 사번 | 측정인고유번호
```

### 스크래핑으로 DB 구축

| 버튼 | 메서드 | 수집 대상 | 저장 위치 |
|------|--------|-----------|-----------|
| 계약 DB 업데이트 | `ScrapeSamplingPointsAsync` | 계약 → 현장 → 채취지점 | `측정인_채취지점` |
| 분석 DB 업데이트 | `ScrapeAnalysisItemsAsync` | add_meas_item 옵션 전체 | `측정인_분석항목` |
| (자동) | `ScrapeEmployeeIdsAsync` | add_emp_id 옵션 전체 | Agent.측정인고유번호 |

---

## 7. 이벤트 패턴 레퍼런스 (VBA 원본)

VBA에서 `WebDriver`로 직접 제어하던 패턴을 CDP `Runtime.evaluate`로 1:1 변환한 것이다.

### CommandButton5 (계약 선택)

```vba
' option.selected → click → change → select2 trigger
jsScript = "var sel = document.getElementById('add_meas_cont_no');"
jsScript = jsScript & "var opts = sel.getElementsByTagName('option');"
jsScript = jsScript & "for (var i = 0; i < opts.length; i++) {"
jsScript = jsScript & "  if (opts[i].value == '" & contractValue & "') {"
jsScript = jsScript & "    opts[i].selected = true;"
jsScript = jsScript & "    opts[i].click();"
jsScript = jsScript & "    break;"
jsScript = jsScript & "  }"
jsScript = jsScript & "}"
jsScript = jsScript & "sel.dispatchEvent(new Event('change'));"
driver.ExecuteScript jsScript
```

### CommandButton6 (분석항목 다중선택)

```vba
' 각 항목 option.selected = true → change → select2.close
For i = 1 To ListView3.ListItems.Count
    measCode = ListView3.ListItems(i).ListSubItems(5).text
    jsScript = jsScript & "Array.from(sel.options).forEach(opt => {" & _
        " if (opt.value == '" & measCode & "') opt.selected = true; });"
Next i
jsScript = jsScript & "sel.dispatchEvent(new Event('change'));"
jsScript = jsScript & "var select2 = $('#add_meas_item').data('select2');"
jsScript = jsScript & "if (select2) { select2.close(); }"
driver.ExecuteScript jsScript
```

### CommandButton9 (인력 선택)

```vba
' 기존 해제 → 다중 선택 → change
jsScript = "var sel = document.getElementById('add_emp_id');"
jsScript = jsScript & "while (sel.selectedOptions.length > 0)"
jsScript = jsScript & "  sel.selectedOptions[0].selected = false;"
jsScript = jsScript & "sel.dispatchEvent(new Event('change', {bubbles: true}));"
driver.ExecuteScript jsScript

' 새 인력 선택
jsScript = "var sel = document.getElementById('add_emp_id');"
jsScript = jsScript & "sel.querySelector('option[value=""" & empValue1 & """]').selected = true;"
jsScript = jsScript & "sel.querySelector('option[value=""" & empValue2 & """]').selected = true;"
jsScript = jsScript & "sel.dispatchEvent(new Event('change', {bubbles: true}));"
driver.ExecuteScript jsScript
```

### CommandButton4 (분석시작일자)

```vba
' input.value 설정 + blur (React/jQuery가 반응하도록)
script1 = "var input = document.getElementById('add_meas_start_dt');"
        & "input.value = '" & DATEX & "';"
        & "input.blur();"
driver.ExecuteScript script1
```

### CommandButton11 (분석자료 입력 — PAGE4)

```vba
' 핵심: td 클릭 → 공통 입력필드에 SendKeys → 다음 td 클릭
targetTD4.Click                          ' 허용기준치 셀 클릭
targetInputField.SendKeys "999"          ' 공통 에디터에 값 입력

targetTD5.Click                          ' 분석결과 셀 클릭
targetInputField.SendKeys resultValue    ' 결과 입력

targetTD7.Click                          ' 분석방법 셀 클릭 (2번)
targetTD7.Click                          ' 드롭다운 열기
Set listItem = driver.FindElementByXPath( _
    "//div[@role='option']//span[contains(text(),'" & methodName & "')]")
listItem.Click                           ' 드롭다운에서 선택

' 동시분석그룹 버튼 (CommandButton21)
' Ctrl+Click으로 다중 행 선택 후 그룹 버튼 클릭
driver.Actions.KeyDown(driver.Keys.Control) _
              .Click(targetDiv) _
              .KeyUp(driver.Keys.Control) _
              .Perform
```

---

## 8. 자료TO측정인 (분석결과 입력)

### 대상 페이지

`field_water.do` → 데이터 열 → 각 항목 행

### 입력 순서 (VBA CommandButton11 기반)

```
1. div 높이 확장 (모든 행 로드)
   ExecuteScript "document.querySelector('div.rg-root').style.height = '5000%';"

2. 모든 tr[role=row] → Dictionary(항목명 → 행번호)

3. 각 항목 순회:
   a. td[5] 클릭 → 허용기준값 입력
   b. td[6] 클릭 → 분석결과 입력 (불검출이면 "-" + 비고에 "불검출")
   c. td[8] 더블클릭 → 분석방법 드롭다운 → span 텍스트로 선택
   d. td[9] 클릭 → 분석장비 드롭다운 → span 텍스트로 선택
   e. td[10] 클릭 → 분석담당자 드롭다운 → 분장표 기반 이름으로 선택
   f. td[11] 클릭 → 분석시작일 입력 (editor648)
   g. td[12] 클릭 → 시작시간 "09:00"
   h. td[13] 클릭 → 분석종료일 입력 (editor648)
   i. td[14] 클릭 → 종료시간 "18:00"
   j. td[15] 클릭 → 시료용량 입력 (editor655)

4. 동시분석그룹 설정 (CommandButton21)
   그룹별 행 Ctrl+Click 선택 → #set-conc_group_no-1 클릭

5. div 높이 복원
```

### 핵심: 셀 클릭 → 공통 에디터 입력

측정인.kr의 그리드는 **인라인 에디터** 방식이다. 셀을 클릭하면 공통 input 요소(`#$$_rg_editor151` 등)가 해당 셀 위치에 나타나고, 여기에 값을 입력한다.

```
td[5].click()  ───▶  #$$_rg_editor151 활성화  ───▶  SendKeys "값"
td[6].click()  ───▶  #$$_rg_editor151 활성화  ───▶  SendKeys "값"
td[8].click()  ───▶  드롭다운 렌더링           ───▶  옵션 span 클릭
```

---

## 9. 트러블슈팅

### select2가 반응하지 않을 때

```javascript
// ✅ option.click() 빠뜨리지 않기
opts[i].selected = true;
opts[i].click();   // ← 이걸 안 하면 select2 UI가 업데이트 안됨

// ✅ jQuery trigger 추가
$('#id').trigger({
    type: 'select2:select',
    params: { data: { id: value, text: text } }
});
$('#id').trigger('change');
```

### AJAX 로딩이 안 끝난 상태에서 다음 단계 실행

```
❌ 계약 선택 직후 현장 선택 → 현장 옵션이 아직 안 불러와짐
✅ 계약 선택 → 400ms 폴링으로 현장 옵션 > 1 확인 → 현장 선택
```

### 모달이 안 열릴 때

`addFieldPlan` 버튼이 DOM에 있지만 `offsetParent === null` (숨김상태)인 경우:
- 이전 단계(로그인 등)가 완료되지 않았을 수 있음
- MeasurerLoginWindow에서 로그인이 끝난 후 `LoginSucceeded = true`로 설정되어야 함

### 값이 설정되었으나 서버에서 무시될 때

`input.value = '...'`만 하면 안되고 반드시 `blur()` 또는 `change` 이벤트를 발생시켜야 한다:

```javascript
// 날짜: blur() 필수
input.value = '2026-03-29';
input.blur();

// 텍스트: input + change + blur
el.value = 'text';
el.dispatchEvent(new Event('input', {bubbles: true}));
el.dispatchEvent(new Event('change', {bubbles: true}));
el.blur();
```

### CDP 응답 65536 바이트 초과

현재 버퍼 크기가 64KB로 고정되어 있다. 대량의 데이터를 반환하는 스크립트는 문자열을 JSON.stringify 후 잘라서 반환하거나, 반환 없이 `window.__etaResult`에 저장하고 별도로 조회한다.

---

## 부록: 전체 실행 시퀀스

```
사용자: Show4에서 시료 선택 → [측정인 전송] 클릭
  │
  ├─ BtnMeasurer_Click
  │   ├─ TryInjectRequestDataAsync(parentNodes)
  │   │   ├─ HTTP GET /json → 탭 찾기
  │   │   ├─ WebSocket 연결
  │   │   ├─ records 직렬화 (JSON)
  │   │   ├─ 측정인 코드 매핑 (분석항목 → 코드값, 인력 → 고유번호)
  │   │   ├─ window.__etaRequestData 주입
  │   │   ├─ addFieldPlan.click()
  │   │   ├─ 모달 대기 (add_meas_cont_no 폴링)
  │   │   │
  │   │   ├─ [A] 근거계약 선택 + 현장 로딩 대기
  │   │   ├─ [B] 현장 선택 + 채취지점 로딩 대기
  │   │   ├─ [C] 환경기술인 입력
  │   │   ├─ [D] 분석시작일자 입력
  │   │   ├─ [E] 분석항목 다중 선택
  │   │   ├─ [F] 측정인력 다중 선택
  │   │   ├─ [G] 측정목적 설정
  │   │   │
  │   │   └─ WebSocket 닫기
  │   │
  │   └─ 실패 시 → 로그인 창 ShowDialog → 재시도
  │
  └─ SetStatus("✅ 완료")
```
