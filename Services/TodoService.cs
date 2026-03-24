using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ETA.Services;

public static class TodoService
{
    // ── 하드코딩 인증 정보 ────────────────────────────────────────────────
    private const string ClientId     = "b8c6ef86-9c86-458e-9c39-a7fec8481d8c";
    private const string TenantId     = "consumers";
    private const string RefreshToken = "M.C560_SN1.0.U.-Cqt8KDlnnPb9WpjEad0Gto4Po7iNhWUvBiBmsy1wJ6HigzLalPXOEX5*J0f!Xnq2oXfw74sNhLP0V06whjF1uF8KtJNLHiwlA7cepj5iAvxrrBJxtJ*wSVxA47Uoq3m55u2yz9f0jF6geHiLVMPQuWyaWd3PUVz7PxSxKElLSXkmaunVd7S4n380ncTzcukrLQsAmKvLmemmHONaucH4bXk8fdAST5KU6466!zAp4Wq1KKC2lsG7SbAsrWVyUyjrLbkMRJE9XVxisFAm60OqBxOtLLd*wHRrn!uGsOA7yNAn9VzCzPOoAB6WUqcrRJGES*LbsWacmnnwS3qxAP9kvgoFubSZQ4NimAp30LcslfAUuI3whH46wwzuNxeqmrk4yQsp1bxLo!9iO06xf5qpyZo$";

    private static readonly HttpClient _http = new();
    private static string _accessToken = "";
    private static DateTime _tokenExpiry = DateTime.MinValue;

    // ── Access Token 갱신 ─────────────────────────────────────────────────
    private static async Task<string> GetAccessTokenAsync()
    {
        if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _tokenExpiry)
            return _accessToken;

        var body = new FormUrlEncodedContent(new[]
        {
            new System.Collections.Generic.KeyValuePair<string,string>("client_id",     ClientId),
            new System.Collections.Generic.KeyValuePair<string,string>("grant_type",    "refresh_token"),
            new System.Collections.Generic.KeyValuePair<string,string>("refresh_token", RefreshToken),
            new System.Collections.Generic.KeyValuePair<string,string>("scope",         "Tasks.ReadWrite User.Read offline_access"),
        });

        var res = await _http.PostAsync(
            $"https://login.microsoftonline.com/{TenantId}/oauth2/v2.0/token", body);

        var json = await res.Content.ReadAsStringAsync();
        var doc  = JsonDocument.Parse(json);

        if (!doc.RootElement.TryGetProperty("access_token", out var at))
        {
            Debug.WriteLine($"[TodoService] 토큰 갱신 실패: {json}");
            throw new Exception("Access token 갱신 실패: " + json);
        }

        _accessToken = at.GetString()!;
        _tokenExpiry = DateTime.UtcNow.AddSeconds(
            doc.RootElement.TryGetProperty("expires_in", out var ei) ? ei.GetInt32() - 60 : 3540);

        Debug.WriteLine("[TodoService] Access Token 갱신 완료");
        return _accessToken;
    }

    // ── To Do 리스트 ID 조회 (기본 리스트 사용) ───────────────────────────
    private static async Task<string> GetDefaultListIdAsync(string accessToken)
    {
        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", accessToken);

        var res  = await _http.GetAsync("https://graph.microsoft.com/v1.0/me/todo/lists");
        var json = await res.Content.ReadAsStringAsync();
        var doc  = JsonDocument.Parse(json);

        foreach (var list in doc.RootElement.GetProperty("value").EnumerateArray())
        {
            // "작업" 또는 "Tasks" 가 기본 리스트
            var name = list.GetProperty("displayName").GetString() ?? "";
            if (name == "작업" || name == "Tasks" || name == "To Do")
                return list.GetProperty("id").GetString()!;
        }

        // 없으면 첫 번째 리스트 사용
        var first = doc.RootElement.GetProperty("value").EnumerateArray();
        first.MoveNext();
        return first.Current.GetProperty("id").GetString()!;
    }

    // ── 할 일 추가 (회원가입 승인 요청) ──────────────────────────────────
    /// <summary>
    /// 신규 직원 회원가입 시 관리자 To Do 에 승인 요청 태스크를 추가합니다.
    /// taskNote 에 직원ID 를 포함해서 승인 확인 시 DB 업데이트에 사용합니다.
    /// </summary>
    public static async Task<string?> AddApprovalTaskAsync(string employeeName, string employeeId)
    {
        try
        {
            var token  = await GetAccessTokenAsync();
            var listId = await GetDefaultListIdAsync(token);

            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            var task = new
            {
                title = $"[ETA 승인요청] {employeeName} 직원 계정 승인",
                body  = new
                {
                    contentType = "text",
                    content     = $"신규 직원 계정 승인 요청입니다.\n\n" +
                                  $"이름: {employeeName}\n" +
                                  $"직원ID: {employeeId}\n" +
                                  $"요청일시: {DateTime.Now:yyyy-MM-dd HH:mm}\n\n" +
                                  $"ETA 시스템에서 승인 처리해주세요."
                },
                importance = "high",
            };

            var content = new StringContent(
                JsonSerializer.Serialize(task),
                Encoding.UTF8,
                "application/json");

            var res  = await _http.PostAsync(
                $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks", content);
            var json = await res.Content.ReadAsStringAsync();

            if (res.IsSuccessStatusCode)
            {
                var doc    = JsonDocument.Parse(json);
                var taskId = doc.RootElement.GetProperty("id").GetString();
                Debug.WriteLine($"[TodoService] 태스크 생성 완료: {taskId}");
                return taskId;
            }

            Debug.WriteLine($"[TodoService] 태스크 생성 실패: {json}");
            return null;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[TodoService] 오류: {ex.Message}");
            return null;
        }
    }

    // ── 태스크 완료 여부 확인 ─────────────────────────────────────────────
    /// <summary>
    /// 관리자가 To Do 에서 체크했는지 확인합니다.
    /// true 이면 승인 완료 → DB 상태를 approved 로 변경해야 합니다.
    /// </summary>
    public static async Task<bool> IsTaskCompletedAsync(string listId, string taskId)
    {
        try
        {
            var token = await GetAccessTokenAsync();
            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            var res  = await _http.GetAsync(
                $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks/{taskId}");
            var json = await res.Content.ReadAsStringAsync();
            var doc  = JsonDocument.Parse(json);

            var status = doc.RootElement.GetProperty("status").GetString();
            return status == "completed";
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[TodoService] IsTaskCompleted 오류: {ex.Message}");
            return false;
        }
    }

    // ── 모든 미완료 승인요청 태스크 확인 + DB 자동 승인 ──────────────────
    /// <summary>
    /// 로그인 시도 시 호출 — To Do 에서 완료된 승인요청이 있으면 DB approved 처리
    /// </summary>
    public static async Task SyncApprovalStatusAsync()
    {
        try
        {
            var token  = await GetAccessTokenAsync();
            var listId = await GetDefaultListIdAsync(token);

            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            // 완료된 태스크 조회
            var res  = await _http.GetAsync(
                $"https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks" +
                $"?$filter=status eq 'completed'&$select=id,title,body,status");
            var json = await res.Content.ReadAsStringAsync();
            var doc  = JsonDocument.Parse(json);

            foreach (var task in doc.RootElement.GetProperty("value").EnumerateArray())
            {
                var title = task.GetProperty("title").GetString() ?? "";
                if (!title.StartsWith("[ETA 승인요청]")) continue;

                // body 에서 직원ID 파싱
                var body = task.GetProperty("body")
                               .GetProperty("content").GetString() ?? "";
                var idLine = System.Array.Find(
                    body.Split('\n'),
                    l => l.StartsWith("직원ID:"));

                if (idLine == null) continue;
                var employeeId = idLine.Replace("직원ID:", "").Trim();

                // DB 승인 처리
                AgentService.ApproveAccount(employeeId);
                Debug.WriteLine($"[TodoService] 자동 승인: {employeeId}");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[TodoService] SyncApproval 오류: {ex.Message}");
        }
    }
}
