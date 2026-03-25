using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Generic;
using ETA.Models;
using System.Linq;

namespace ETA.Services;

public static class TodoService
{
    // ── 하드코딩 인증 정보 ────────────────────────────────────────────────
    private const string ClientId     = "b8c6ef86-9c86-458e-9c39-a7fec8481d8c";
    private const string TenantId     = "consumers";
    private const string RefreshToken = "M.C560_SN1.0.U.-Cqt8KDlnnPb9WpjEad0Gto4Po7iNhWUvBiBmsy1wJ6HigzLalPXOEX5*J0f!Xnq2oXfw74sNhLP0V06whjF1uF8KtJNLHiwlA7cepj5iAvxrrBJxtJ*wSVxA47Uoq3m55u2yz9f0jF6geHiLVMPQuWyaWd3PUVz7PxSxKElLSXkmaunVd7S4n380ncTzcukrLQsAmKvLmemmHONaucH4bXk8fdAST5KU6466!zAp4Wq1KKC2lsG7SbAsrWVyUyjrLbkMRJE9XVxisFAm60OqBxOtLLd*wHRrn!uGsOA7yNAn9VzCzPOoAB6WUqcrRJGES*LbsWacmnnwS3qxAP9kvgoFubSZQ4NimAp30LcslfAUuI3whH46wwzuNxeqmrk4yQsp1bxLo!9iO06xf5qpyZo$";

    private static readonly HttpClient _http = new();
    private static string   _accessToken = "";
    private static DateTime _tokenExpiry  = DateTime.MinValue;

    // =========================================================================
    // 인증
    // =========================================================================

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

    // =========================================================================
    // To Do 리스트 ID
    // =========================================================================

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

    // =========================================================================
    // 회원가입 시 승인 요청 Task 생성 (중복 방지 포함)
    // =========================================================================

    /// <summary>
    /// 신규 직원 회원가입 시 호출. To Do에 동일 사번의 Task가 이미 있으면 재사용하고,
    /// 없으면 새로 생성한다. 반환값은 taskId (DB에 저장해야 함).
    /// </summary>
    public static async Task<string?> AddApprovalTaskAsync(string employeeName, string employeeId)
    {
        try
        {
            var token  = await GetAccessTokenAsync();
            var listId = await GetDefaultListIdAsync(token);

            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            // 이미 동일 사번 Task가 존재하는지 검색 (중복 방지)
            var existing = await FindTaskByEmpIdAsync(listId, employeeId);
            if (existing != null)
            {
                Debug.WriteLine($"[TodoService] 기존 Task 재사용: {existing}");
                return existing;
            }

            var taskId = await CreateTaskAsync(listId, employeeName, employeeId, completed: false);
            Debug.WriteLine($"[TodoService] 승인 요청 Task 생성 완료: {taskId}");
            return taskId;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[TodoService] AddApprovalTaskAsync 오류: {ex.Message}");
            return null;
        }
    }

    // =========================================================================
    // ★★★ 핵심: DB ↔ To Do 양방향 동기화 ★★★
    //
    // 동작 원칙 (DB가 Master):
    //   1. To Do의 ETA Task 목록을 전부 읽는다.
    //   2. DB 계정 목록(pending/approved)과 대조한다.
    //   3. To Do에서 Task가 완료(completed) → DB를 approved로 갱신
    //   4. To Do에서 Task가 미완료(notStarted) → DB가 approved면 pending으로 되돌림 (승인 취소)
    //   5. To Do에서 Task 자체가 없음(삭제됨) → DB가 approved면 pending으로 되돌림
    //   6. DB에 pending인데 To Do Task가 없음 → 새로 생성 (중복 방지)
    //   7. DB에 approved인데 To Do Task가 없음 → completed Task 재생성 (UI 표시용)
    // =========================================================================

    public static async Task SyncApprovalStatusAsync()
    {
        string logDir  = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ETA");
        Directory.CreateDirectory(logDir);
        string logPath = Path.Combine(logDir, "eta_sync.log");

        void Log(string msg)
        {
            try { File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}{Environment.NewLine}"); }
            catch { }
        }

        Log("=== SyncApprovalStatusAsync() 시작 (양방향 동기화) ===");

        try
        {
            var token  = await GetAccessTokenAsync();
            var listId = await GetDefaultListIdAsync(token);
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            // ── Step 1: To Do의 ETA Task 전체 조회 ─────────────────────────
            // key: empId, value: (taskId, isCompleted)
            var todoTasks = await GetAllEtaTasksAsync(listId);
            Log($"To Do ETA Task 수: {todoTasks.Count}개");
            foreach (var kv in todoTasks)
                Log($"  To Do [{kv.Key}] taskId={kv.Value.taskId} completed={kv.Value.isCompleted}");

            // ── Step 2: DB 계정 전체 조회 ──────────────────────────────────
            var dbAccounts = AgentService.GetAllApprovalAccounts();
            Log($"DB 계정 수: {dbAccounts.Count}개");

            int approvedCount = 0, pendingCount = 0, recreatedCount = 0;

            foreach (var (name, empId, dbStatus, savedTaskId) in dbAccounts)
            {
                if (string.IsNullOrEmpty(empId)) continue;

                bool todoHasTask  = todoTasks.ContainsKey(empId);
                bool todoCompleted = todoHasTask && todoTasks[empId].isCompleted;
                string? todoTaskId = todoHasTask ? todoTasks[empId].taskId : null;

                Log($"[계정] {name} ({empId}) | DB상태={dbStatus} | To Do 존재={todoHasTask} | To Do 완료={todoCompleted}");

                // ── Case A: To Do Task가 완료 → DB를 approved로 갱신 ──────
                if (todoHasTask && todoCompleted && dbStatus == "pending")
                {
                    AgentService.ApproveAccount(empId);
                    if (todoTaskId != null)
                        AgentService.SaveTodoTaskId(empId, todoTaskId);
                    approvedCount++;
                    Log($"  ✅ pending → approved (To Do 완료 감지)");
                }
                // ── Case B: To Do Task가 미완료로 바뀜 → approved → pending ─
                else if (todoHasTask && !todoCompleted && dbStatus == "approved")
                {
                    AgentService.ResetToPending(empId);
                    Log($"  ⚠️  approved → pending (To Do 미완료 감지 = 승인 취소)");
                    pendingCount++;
                }
                // ── Case C / D: To Do Task 자체가 없음 (삭제됐거나 미생성) ──
                //   → DB 상태를 Master로 삼아 그대로 Task 재생성
                //   → DB 상태는 변경하지 않는다 (삭제는 실수일 수 있음)
                else if (!todoHasTask)
                {
                    bool shouldBeCompleted = dbStatus == "approved";

                    AgentService.SaveTodoTaskId(empId, "");   // 기존 유효하지 않은 taskId 초기화
                    var newTaskId = await CreateTaskAsync(listId, name, empId, completed: shouldBeCompleted);

                    if (newTaskId != null)
                    {
                        AgentService.SaveTodoTaskId(empId, newTaskId);
                        recreatedCount++;
                        Log($"  🔄 Task 없음 → DB 상태({dbStatus}) 기준으로 재생성 " +
                            $"[{(shouldBeCompleted ? "completed" : "notStarted")}]: {newTaskId}");
                    }
                    else
                    {
                        Log($"  ❌ Task 재생성 실패: {name} ({empId})");
                    }
                }
                // ── Case E: To Do 완료 + DB도 approved → Task ID 동기화만 ─
                else if (todoHasTask && todoCompleted && dbStatus == "approved")
                {
                    if (todoTaskId != null && savedTaskId != todoTaskId)
                    {
                        AgentService.SaveTodoTaskId(empId, todoTaskId);
                        Log($"  🔗 taskId 동기화만 수행");
                    }
                    else
                    {
                        Log($"  ✔ 이미 동기화된 상태. 변경 없음.");
                    }
                }
            }

            Log($"=== Sync 완료 === approved={approvedCount} / pending전환={pendingCount} / 재생성={recreatedCount}");
        }
        catch (Exception ex)
        {
            Log($"❌ SyncApprovalStatusAsync 오류: {ex.Message}\n{ex.StackTrace}");
        }
    }

    // =========================================================================
    // 내부 헬퍼
    // =========================================================================

    /// <summary>
    /// To Do 리스트에서 [ETA 승인요청] 태그가 붙은 Task를 모두 읽어
    /// { empId → (taskId, isCompleted) } 딕셔너리로 반환한다.
    /// </summary>
    private static async Task<Dictionary<string, (string taskId, bool isCompleted)>> GetAllEtaTasksAsync(string listId)
    {
        var result = new Dictionary<string, (string, bool)>(StringComparer.OrdinalIgnoreCase);

        try
        {
            // Microsoft Graph는 완료된 Task를 기본적으로 숨긴다 → $filter 없이 전체 조회
            // completed Task도 가져오려면 status ne 'completed' 필터를 쓰지 않아야 함
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

                    string taskId      = task.TryGetProperty("id", out var idProp)     ? idProp.GetString()  ?? "" : "";
                    bool   isCompleted = task.TryGetProperty("status", out var sProp)  && sProp.GetString() == "completed";
                    string empId       = ExtractEmpIdFromTask(task);

                    if (!string.IsNullOrEmpty(empId) && !string.IsNullOrEmpty(taskId))
                    {
                        // 같은 empId의 Task가 여럿 있을 때: completed를 우선, 없으면 가장 최신 것
                        if (!result.ContainsKey(empId) || isCompleted)
                            result[empId] = (taskId, isCompleted);
                    }
                }

                // 페이징 처리
                url = doc.RootElement.TryGetProperty("@odata.nextLink", out var next)
                    ? next.GetString() ?? ""
                    : "";
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[TodoService] GetAllEtaTasksAsync 오류: {ex.Message}");
        }

        return result;
    }

    /// <summary>
    /// 특정 empId를 body에 포함하는 [ETA 승인요청] Task가 이미 있으면 taskId 반환, 없으면 null.
    /// </summary>
    private static async Task<string?> FindTaskByEmpIdAsync(string listId, string empId)
    {
        var all = await GetAllEtaTasksAsync(listId);
        return all.TryGetValue(empId, out var t) ? t.taskId : null;
    }

    /// <summary>
    /// To Do에 [ETA 승인요청] Task를 생성한다. completed=true면 완료 상태로 생성.
    /// </summary>
    private static async Task<string?> CreateTaskAsync(
        string listId, string name, string empId, bool completed)
    {
        var task = new
        {
            title = $"[ETA 승인요청] {name} 직원 계정 승인",
            body  = new
            {
                contentType = "text",
                content     = $"신규 직원 계정 승인 요청입니다.\n\n" +
                              $"이름: {name}\n" +
                              $"직원ID: {empId}\n" +
                              $"요청일시: {DateTime.Now:yyyy-MM-dd HH:mm}\n\n" +
                              $"ETA 시스템에서 승인 처리해주세요."
            },
            importance = "high",
            status     = completed ? "completed" : "notStarted"
        };

        var content = new StringContent(
            JsonSerializer.Serialize(task), Encoding.UTF8, "application/json");

        var res = await _http.PostAsync(
            $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks", content);

        if (res.IsSuccessStatusCode)
        {
            var json = await res.Content.ReadAsStringAsync();
            var doc  = JsonDocument.Parse(json);
            return doc.RootElement.GetProperty("id").GetString();
        }

        Debug.WriteLine($"[TodoService] Task 생성 실패: {await res.Content.ReadAsStringAsync()}");
        return null;
    }

    /// <summary>
    /// Task body에서 '직원ID:' 라인을 파싱해 사번을 추출한다.
    /// </summary>
    private static string ExtractEmpIdFromTask(JsonElement task)
    {
        if (!task.TryGetProperty("body", out var bodyProp) ||
            !bodyProp.TryGetProperty("content", out var contentProp))
            return "";

        var body = contentProp.GetString() ?? "";
        foreach (var line in body.Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("직원ID:"))
                return trimmed.Replace("직원ID:", "").Trim();
        }
        return "";
    }
}
