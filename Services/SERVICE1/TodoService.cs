using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
using ETA.Models;
using ETA.Views.Pages;
using ETA.Views.Pages.PAGE1;
using ETA.Views.Pages.PAGE2;
using ETA.Views.Pages.Common;
using System.Linq;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

public static class TodoService
{
    // ── 인증 정보 ─────────────────────────────────────────────────────────
    private const string ClientId     = "b8c6ef86-9c86-458e-9c39-a7fec8481d8c";
    private const string TenantId     = "consumers";
    private const string RefreshToken = "M.C560_SN1.0.U.-Cqt8KDlnnPb9WpjEad0Gto4Po7iNhWUvBiBmsy1wJ6HigzLalPXOEX5*J0f!Xnq2oXfw74sNhLP0V06whjF1uF8KtJNLHiwlA7cepj5iAvxrrBJxtJ*wSVxA47Uoq3m55u2yz9f0jF6geHiLVMPQuWyaWd3PUVz7PxSxKElLSXkmaunVd7S4n380ncTzcukrLQsAmKvLmemmHONaucH4bXk8fdAST5KU6466!zAp4Wq1KKC2lsG7SbAsrWVyUyjrLbkMRJE9XVxisFAm60OqBxOtLLd*wHRrn!uGsOA7yNAn9VzCzPOoAB6WUqcrRJGES*LbsWacmnnwS3qxAP9kvgoFubSZQ4NimAp30LcslfAUuI3whH46wwzuNxeqmrk4yQsp1bxLo!9iO06xf5qpyZo$";

    private static readonly HttpClient _http = new();
    private static string   _accessToken = "";
    private static DateTime _tokenExpiry  = DateTime.MinValue;

    // =====================================================================
    //  인증
    // =====================================================================
    private static async Task<string> GetAccessTokenAsync()
    {
        if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _tokenExpiry)
            return _accessToken;

        var body = new FormUrlEncodedContent(new[]
        {
            new KeyValuePair<string,string>("client_id",     ClientId),
            new KeyValuePair<string,string>("grant_type",    "refresh_token"),
            new KeyValuePair<string,string>("refresh_token", RefreshToken),
            new KeyValuePair<string,string>("scope",         "Tasks.ReadWrite User.Read offline_access"),
        });

        var res  = await _http.PostAsync(
            $"https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token", body);
        var json = await res.Content.ReadAsStringAsync();
        var doc  = JsonDocument.Parse(json);

        if (!doc.RootElement.TryGetProperty("access_token", out var at))
            throw new Exception("Access token 갱신 실패: " + json);

        _accessToken = at.GetString()!;
        _tokenExpiry = DateTime.UtcNow.AddSeconds(
            doc.RootElement.TryGetProperty("expires_in", out var ei) ? ei.GetInt32() - 60 : 3540);

        Debug.WriteLine("[TodoService] Access Token 갱신 완료");
        return _accessToken;
    }

    // =====================================================================
    //  기본 리스트 ID (승인용)
    // =====================================================================
    private static async Task<string> GetDefaultListIdAsync(string accessToken)
    {
        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", accessToken);

        var res  = await _http.GetAsync("https://graph.microsoft.com/v1.0/me/todo/lists");
        var json = await res.Content.ReadAsStringAsync();
        var doc  = JsonDocument.Parse(json);

        foreach (var list in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            var name = list.GetProperty("displayName").GetString() ?? "";
            if (name == "작업" || name == "Tasks" || name == "To Do")
                return list.GetProperty("id").GetString()!;
        }

        var en = doc.RootElement.GetProperty("value").EnumerateArray();
        en.MoveNext();
        return en.Current.GetProperty("id").GetString()!;
    }

    // =====================================================================
    //  승인 Task 생성
    // =====================================================================
    public static async Task<string?> AddApprovalTaskAsync(string employeeName, string employeeId)
    {
        try
        {
            var token  = await GetAccessTokenAsync();
            var listId = await GetDefaultListIdAsync(token);
            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            var existing = await FindTaskByEmpIdAsync(listId, employeeId);
            if (existing != null) { Debug.WriteLine($"[TodoService] 기존 Task 재사용: {existing}"); return existing; }

            var taskId = await CreateApprovalTaskAsync(listId, employeeName, employeeId);
            Debug.WriteLine($"[TodoService] 승인 Task 생성: {taskId}");
            return taskId;
        }
        catch (Exception ex) { Debug.WriteLine($"[TodoService] AddApprovalTaskAsync 오류: {ex.Message}"); return null; }
    }

    // =====================================================================
    //  DB ↔ To Do 양방향 동기화
    // =====================================================================
    public static async Task SyncApprovalStatusAsync()
    {
        string logDir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ETA");
        Directory.CreateDirectory(logDir);
        string logPath = Path.Combine(logDir, "Logs/eta_sync.log");
        void Log(string msg)
        {
            try { File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}{Environment.NewLine}"); } catch { }
        }

        Log("=== SyncApprovalStatusAsync() 시작 ===");
        try
        {
            var token  = await GetAccessTokenAsync();
            var listId = await GetDefaultListIdAsync(token);
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var todoTasks  = await GetAllEtaTasksAsync(listId);
            var dbAccounts = AgentService.GetAllApprovalAccounts();
            Log($"To Do ETA Task: {todoTasks.Count}개  DB 계정: {dbAccounts.Count}개");

            int approvedCount = 0, pendingCount = 0, recreatedCount = 0;

            foreach (var (name, empId, dbStatus, savedTaskId) in dbAccounts)
            {
                if (string.IsNullOrEmpty(empId)) continue;
                bool    has       = todoTasks.ContainsKey(empId);
                bool    completed = has && todoTasks[empId].isCompleted;
                string? todoId    = has ? todoTasks[empId].taskId : null;

                if (has && completed && dbStatus == "pending")
                {
                    AgentService.ApproveAccount(empId);
                    if (todoId != null) AgentService.SaveTodoTaskId(empId, todoId);
                    approvedCount++;
                    Log($"  ✅ {name} pending→approved");
                }
                else if (has && !completed && dbStatus == "approved")
                {
                    AgentService.ResetToPending(empId);
                    pendingCount++;
                    Log($"  ⚠️  {name} approved→pending");
                }
                else if (!has)
                {
                    bool comp = dbStatus == "approved";
                    AgentService.SaveTodoTaskId(empId, "");
                    var newId = await CreateApprovalTaskAsync(listId, name, empId, comp);
                    if (newId != null) { AgentService.SaveTodoTaskId(empId, newId); recreatedCount++; Log($"  🔄 {name} 재생성"); }
                }
                else if (has && completed && dbStatus == "approved")
                {
                    if (todoId != null && savedTaskId != todoId) { AgentService.SaveTodoTaskId(empId, todoId); Log($"  🔗 {name} taskId 동기화"); }
                }
            }
            Log($"=== Sync 완료 approved={approvedCount} pending={pendingCount} 재생성={recreatedCount} ===");
        }
        catch (Exception ex) { Log($"❌ SyncApprovalStatusAsync 오류: {ex.Message}\n{ex.StackTrace}"); }
    }

    // =====================================================================
    //  ★ 분석의뢰 TODO 전송 ★
    //
    //  [수질분석의뢰-측대] 리스트:
    //
    //  중간 Task (부모노드 1개당 1개):
    //    "(03/25) 보임 폐수조#APond-폐수"
    //    body: [분석의뢰-중간] + 접수번호 (중복방지)
    //
    //  최종 Task (자식노드 분석항목 1개당 1개):
    //    "(03/31) SS 김지은"
    //    - 약식명: 분장표준처리 ROW "약칭" 에서 해당 컬럼값 (SS, BOD 등)
    //    - 완료기한: 채취일 + 분장표준처리 ROW "표준처리기한" 의 처리일수(영업일)
    //    - 담당자: 분장표준처리 ROW "채취일자" 에서 해당 컬럼값
    //    - dueDateTime 설정
    //    → MS Todo 앱에서 중간 Task 안의 "다음 단계"로 표시됨
    //
    //  노드 3개 × 항목 10개 → 중간 Task 3개, 각각 checklistItem 10개
    // =====================================================================
    public static async Task SendAnalysisRequestsToTodoAsync(
        IEnumerable<(AnalysisRequestRecord rec,
                     Dictionary<string, string> row,
                     List<string> analyteNames)> items)
    {
        var token  = await GetAccessTokenAsync();
        var listId = await GetAnalysisListIdAsync(token);
        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", token);

        // 분장표준처리: 처리일수 + 약식명 (1회 조회)
        var stdInfo = await Task.Run(() => AnalysisRequestService.GetStandardDaysInfo());

        TodoLog("=== SendAnalysisRequestsToTodoAsync 시작 ===");
        TodoLog($"stdInfo 항목 수: {stdInfo.Count}");
        foreach (var kv in stdInfo)
            TodoLog($"  stdInfo: [{kv.Key}] days={kv.Value.days} short={kv.Value.shortName}");

        foreach (var (rec, row, analyteNames) in items)
        {
            // ── 채취일자 파싱 ────────────────────────────────────────────
            DateTime? sampleDt   = null;
            string    sampleDisp = "";
            string    sampleKey  = ""; // yyyy-MM-dd (분장표준처리 행 검색용)

            row.TryGetValue("채취일자", out var rawDate);
            TodoLog($"[{rec.약칭} {rec.시료명}] 채취일자 raw='{rawDate}'");

            if (!string.IsNullOrWhiteSpace(rawDate))
            {
                if (DateTime.TryParse(rawDate, out var dt))
                {
                    sampleDt   = dt;
                    sampleDisp = dt.ToString("MM/dd", System.Globalization.CultureInfo.InvariantCulture);
                    sampleKey  = dt.ToString("yyyy-MM-dd");
                }
                else
                {
                    sampleDisp = rawDate.Length >= 10 ? rawDate[5..10].Replace("-", "/") : rawDate;
                    sampleKey  = rawDate.Trim();
                }
            }
            TodoLog($"  sampleDisp='{sampleDisp}'  sampleKey='{sampleKey}'");

            // ── 분장표준처리에서 채취일자 기준 담당자 목록 조회 ──────────
            Dictionary<string, string> managers;
            if (string.IsNullOrEmpty(sampleKey))
            {
                managers = new Dictionary<string, string>();
                TodoLog("  담당자 조회 SKIP: sampleKey 없음");
            }
            else
            {
                managers = await Task.Run(() => AnalysisRequestService.GetManagersByDate(sampleKey));
                TodoLog($"  담당자 조회 결과: {managers.Count}건  (sampleKey={sampleKey})");
                foreach (var m in managers)
                    TodoLog($"    manager: [{m.Key}] = '{m.Value}'");
            }

            // ── 중간 Task 제목: "🧪 [03/25] 보임 폐수조#APond-폐수 #접수번호" ─
            // title에 #{접수번호}를 포함 → 중복 체크 시 title 검색으로 찾음
            string parentTitle = string.IsNullOrEmpty(sampleDisp)
                ? $"📋 {rec.약칭} {rec.시료명}"
                : $"📋 [{sampleDisp}] {rec.약칭} {rec.시료명}";

            // ── 중복 체크: title 전체로 검색 (채취일자+약칭+시료명 조합이 고유키)
            var existingId = await FindAnalysisTaskByTitleAsync(listId, parentTitle);
            string parentTaskId;
            if (existingId != null)
            {
                parentTaskId = existingId;
                TodoLog($"  중간 Task 재사용: [{parentTitle}]");
            }
            else
            {
                parentTaskId = await CreateAnalysisParentTaskAsync(
                    listId, parentTitle, rec) ?? "";
                TodoLog($"  중간 Task 생성: [{parentTitle}] → {parentTaskId}");
            }

            if (string.IsNullOrEmpty(parentTaskId)) { TodoLog("  !! parentTaskId 없음 — SKIP"); continue; }

            // ── checklistItem (다음 단계) — 분석항목별 ───────────────────
            foreach (string analyteFull in analyteNames)
            {
                string analyteKey = analyteFull.Trim();

                // 약식명 + 처리일수
                int    days      = 0;
                string shortName = analyteKey;
                if (stdInfo.TryGetValue(analyteKey, out var info))
                {
                    days = info.days;
                    if (!string.IsNullOrEmpty(info.shortName))
                        shortName = info.shortName;
                }
                TodoLog($"  항목: [{analyteKey}] → short='{shortName}' days={days}");

                // 완료기한: 채취일 + 처리일수(영업일)
                string    dueDateDisp = "";
                DateTime? dueDate     = null;
                if (sampleDt.HasValue && days > 0)
                {
                    dueDate     = AddBusinessDays(sampleDt.Value, days);
                    dueDateDisp = dueDate.Value.ToString("MM/dd", System.Globalization.CultureInfo.InvariantCulture);
                }

                // 담당자: 분장표준처리에서 채취일자 행의 해당 컬럼값
                managers.TryGetValue(analyteKey, out var manager);
                manager = string.IsNullOrWhiteSpace(manager) ? "" : manager.Trim();
                TodoLog($"    dueDate='{dueDateDisp}'  manager='{manager}'");

                // checklistItem 제목: "🧪 [03/31] SS 김지은"
                string childTitle = string.IsNullOrEmpty(dueDateDisp)
                    ? $"🧪 {shortName}"
                    : $"🧪 [{dueDateDisp}] {shortName}";
                if (!string.IsNullOrEmpty(manager))
                    childTitle += $" {manager}";

                TodoLog($"    → checklistItem: '{childTitle}'");
                await AddChecklistItemAsync(listId, parentTaskId, childTitle);
            }
        }
        TodoLog("=== 완료 ===");
    }

    // =====================================================================
    //  내부 헬퍼 — 승인 관련
    // =====================================================================
    private static async Task<Dictionary<string, (string taskId, bool isCompleted)>>
        GetAllEtaTasksAsync(string listId)
    {
        var result = new Dictionary<string, (string, bool)>(StringComparer.OrdinalIgnoreCase);
        try
        {
            string url = $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks?$top=200";
            while (!string.IsNullOrEmpty(url))
            {
                var res  = await _http.GetAsync(url);
                var json = await res.Content.ReadAsStringAsync();
                if (!res.IsSuccessStatusCode) break;
                var doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("value", out var tasks)) break;

                foreach (var task in tasks.EnumerateArray())
                {
                    var title = task.TryGetProperty("title", out var t) ? t.GetString() ?? "" : "";
                    if (!title.StartsWith("[ETA 승인요청]")) continue;
                    string taskId      = task.TryGetProperty("id", out var idp)  ? idp.GetString()  ?? "" : "";
                    bool   isCompleted = task.TryGetProperty("status", out var sp) && sp.GetString() == "completed";
                    string empId       = ExtractEmpIdFromTask(task);
                    if (!string.IsNullOrEmpty(empId) && !string.IsNullOrEmpty(taskId))
                        if (!result.ContainsKey(empId) || isCompleted)
                            result[empId] = (taskId, isCompleted);
                }
                url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next)
                    ? next.GetString() ?? "" : "";
            }
        }
        catch (Exception ex) { Debug.WriteLine($"[TodoService] GetAllEtaTasksAsync 오류: {ex.Message}"); }
        return result;
    }

    private static async Task<string?> FindTaskByEmpIdAsync(string listId, string empId)
    {
        var all = await GetAllEtaTasksAsync(listId);
        return all.TryGetValue(empId, out var t) ? t.taskId : null;
    }

    private static async Task<string?> CreateApprovalTaskAsync(
        string listId, string name, string empId, bool completed = false)
    {
        var task = new
        {
            title = $"[ETA 승인요청] {name} 직원 계정 승인",
            body  = new { contentType = "text",
                content = $"신규 직원 계정 승인 요청입니다.\n\n이름: {name}\n직원ID: {empId}\n요청일시: {DateTime.Now:yyyy-MM-dd HH:mm}\n\nETA 시스템에서 승인 처리해주세요." },
            importance = "high",
            status     = completed ? "completed" : "notStarted"
        };

        var res = await _http.PostAsync(
            $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks",
            new StringContent(JsonSerializer.Serialize(task), Encoding.UTF8, "application/json"));

        if (res.IsSuccessStatusCode)
        {
            var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());
            return doc.RootElement.GetProperty("id").GetString();
        }
        Debug.WriteLine($"[TodoService] 승인 Task 생성 실패: {await res.Content.ReadAsStringAsync()}");
        return null;
    }

    private static string ExtractEmpIdFromTask(JsonElement task)
    {
        if (!task.TryGetProperty("body", out var b) || !b.TryGetProperty("content", out var c)) return "";
        foreach (var line in (c.GetString() ?? "").Split('\n'))
        {
            var t = line.Trim();
            if (t.StartsWith("직원ID:")) return t.Replace("직원ID:", "").Trim();
        }
        return "";
    }

    // =====================================================================
    //  내부 헬퍼 — 분석의뢰 전용
    // =====================================================================

    // "수질분석의뢰-측대" 리스트 ID (없으면 자동 생성)
    private static async Task<string> GetAnalysisListIdAsync(string accessToken)
    {
        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", accessToken);

        var res  = await _http.GetAsync("https://graph.microsoft.com/v1.0/me/todo/lists");
        var doc  = JsonDocument.Parse(await res.Content.ReadAsStringAsync());

        foreach (var list in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            var name = list.GetProperty("displayName").GetString() ?? "";
            if (name == "수질분석의뢰-측대")
                return list.GetProperty("id").GetString()!;
        }

        var body = new JsonObject { ["displayName"] = "수질분석의뢰-측대" };
        var cr   = await _http.PostAsync(
            "https://graph.microsoft.com/v1.0/me/todo/lists",
            new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json"));
        var cd = JsonDocument.Parse(await cr.Content.ReadAsStringAsync());
        return cd.RootElement.GetProperty("id").GetString()!;
    }

    // title 전체 일치로 중간 Task 중복 검색
    // 채취일자 + 약칭 + 시료명 조합이 title에 들어가므로 고유키 역할을 함
    private static async Task<string?> FindAnalysisTaskByTitleAsync(string listId, string title)
    {
        string url = $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks?$top=200";
        while (!string.IsNullOrEmpty(url))
        {
            var res = await _http.GetAsync(url);
            if (!res.IsSuccessStatusCode) break;
            var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());
            if (!doc.RootElement.TryGetProperty("value", out var tasks)) break;

            foreach (var task in tasks.EnumerateArray())
            {
                var taskTitle = task.TryGetProperty("title", out var t) ? t.GetString() ?? "" : "";
                if (taskTitle == title)
                    return task.GetProperty("id").GetString();
            }
            url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next)
                ? next.GetString() ?? "" : "";
        }
        return null;
    }
    // 중간 Task 생성: "📋 [03/25] 보임 폐수조#APond-폐수"
    // 고유키 = 채취일자 + 약칭 + 시료명 (title 자체)
    private static async Task<string?> CreateAnalysisParentTaskAsync(
        string listId, string title, AnalysisRequestRecord rec)
    {
        var task = new
        {
            title      = title,
            importance = "high",
            status     = "notStarted",
            body       = new { contentType = "text",
                content = $"[분석의뢰]\n" +
                          $"약칭: {rec.약칭}\n" +
                          $"시료명: {rec.시료명}\n" +
                          $"접수번호(견적번호): {rec.접수번호}\n" +
                          $"의뢰일: {rec.의뢰일}\n" +
                          $"생성: ETA 시스템" }
        };

        var res = await _http.PostAsync(
            $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks",
            new StringContent(JsonSerializer.Serialize(task), Encoding.UTF8, "application/json"));

        if (res.IsSuccessStatusCode)
        {
            var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());
            return doc.RootElement.GetProperty("id").GetString();
        }
        Debug.WriteLine($"[TodoService] 중간 Task 생성 실패: {await res.Content.ReadAsStringAsync()}");
        return null;
    }
    // 중간 Task의 다음 단계(checklistItem) 추가: "(03/31) SS 김지은"
    // MS Todo 앱에서 Task 안의 "다음 단계"로 표시됨
    // ── TODO 전용 로그 (Documents/ETA/TODO.log) ──────────────────────────
    private static readonly string _todoLogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "Logs/TODO.log");

    private static void TodoLog(string msg)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_todoLogPath)!);
            File.AppendAllText(_todoLogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}{Environment.NewLine}");
        }
        catch { }
        Debug.WriteLine($"[TODO] {msg}");
    }

        private static async Task AddChecklistItemAsync(
        string listId,
        string taskId,
        string displayName)
    {
        var item = new JsonObject
        {
            ["displayName"] = displayName,
            ["isChecked"]   = false
        };

        var res = await _http.PostAsync(
            $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks/{taskId}/checklistItems",
            new StringContent(item.ToJsonString(), Encoding.UTF8, "application/json"));

        if (!res.IsSuccessStatusCode)
            Debug.WriteLine($"[TodoService] checklistItem 추가 실패: {await res.Content.ReadAsStringAsync()}");
    }

        // 영업일 계산 (토/일 건너뜀)
    private static DateTime AddBusinessDays(DateTime start, int days)
    {
        var result = start;
        int added  = 0;
        while (added < days)
        {
            result = result.AddDays(1);
            if (result.DayOfWeek != DayOfWeek.Saturday &&
                result.DayOfWeek != DayOfWeek.Sunday)
                added++;
        }
        return result;
    }
}
