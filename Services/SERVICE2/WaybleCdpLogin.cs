using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace ETA.Services.SERVICE2;

/// <summary>
/// Edge/Chrome CDP 를 통한 Wayble OAuth 로그인.
///   1) --remote-debugging-port=9223, --user-data-dir=AppData/ETA/WaybleBrowserSession 로 브라우저 실행
///   2) rewater.wayble.eco/stp 오픈 → Cognito 호스팅 UI 리다이렉트
///   3) userId/password 자동 입력 → 로그인 버튼 클릭
///   4) stpsession 쿠키 감지 시 반환
/// 이후 호출부는 WaybleSession.SeedStpSession(cookie) 로 HttpClient 세션에 주입한다.
/// </summary>
public static class WaybleCdpLogin
{
    private const int    CdpPort      = 9223;
    private const string StartUrl     = "https://rewater.wayble.eco/stp";
    private const string CookieName   = "stpsession";
    private const string CookieHost   = "https://rewater.wayble.eco";

    private static int _cdpId;

    public record Result(bool Ok, string Message, string? Cookie);

    public static async Task<Result> LoginAsync(
        string userId,
        string password,
        Action<string> onStatus,
        CancellationToken ct)
    {
        var browser = FindBrowser();
        if (browser is null) return new(false, "Edge 또는 Chrome 실행 파일을 찾을 수 없습니다.", null);

        var sessionDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ETA", "WaybleBrowserSession");
        Directory.CreateDirectory(sessionDir);

        onStatus("브라우저 실행 중...");
        try
        {
            if (OperatingSystem.IsMacOS() && browser.AppBundlePath is not null)
            {
                // macOS: `open -na` 로 기존 Edge 인스턴스와 독립된 새 프로세스 실행
                // (그렇지 않으면 --remote-debugging-port / --user-data-dir 이 무시됨)
                var psi = new ProcessStartInfo
                {
                    FileName        = "open",
                    UseShellExecute = false,
                };
                psi.ArgumentList.Add("-na");
                psi.ArgumentList.Add(browser.AppBundlePath);
                psi.ArgumentList.Add("--args");
                psi.ArgumentList.Add($"--remote-debugging-port={CdpPort}");
                psi.ArgumentList.Add($"--user-data-dir={sessionDir}");
                psi.ArgumentList.Add("--no-first-run");
                psi.ArgumentList.Add("--no-default-browser-check");
                psi.ArgumentList.Add(StartUrl);
                Process.Start(psi);
            }
            else
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName        = browser.ExePath,
                    Arguments       = $"--remote-debugging-port={CdpPort} " +
                                      $"--user-data-dir=\"{sessionDir}\" " +
                                      $"--no-first-run --no-default-browser-check " +
                                      $"\"{StartUrl}\"",
                    UseShellExecute = false,
                });
            }
        }
        catch (Exception ex)
        {
            return new(false, $"브라우저 실행 실패: {ex.Message}", null);
        }

        onStatus("CDP 응답 대기 중...");
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
        string tabs;
        try { tabs = await PollCdpJsonAsync(http, ct); }
        catch (Exception ex) { return new(false, ex.Message, null); }

        var pick = PickPageTarget(tabs);
        if (pick is null) return new(false, "CDP page target 을 찾을 수 없습니다.", null);
        var (wsUrl, targetId) = pick.Value;

        using var ws = new ClientWebSocket();
        await ws.ConnectAsync(new Uri(wsUrl), ct);

        await SendAsync(ws, "Network.enable", null, ct);
        await SendAsync(ws, "Page.enable",    null, ct);

        // 탭을 foreground 로 올려야 Input.dispatchMouseEvent 가 먹는 경우가 있다
        try { await SendAsync(ws, "Page.bringToFront", null, ct); } catch { }

        onStatus("Wayble 로그인 페이지 대기 중...");
        var deadline = DateTime.UtcNow.AddMinutes(3);
        bool filled  = false;
        int  probes  = 0;

        var idJson = JsonSerializer.Serialize(userId);
        var pwJson = JsonSerializer.Serialize(password);

        while (DateTime.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();
            await Task.Delay(1000, ct);
            probes++;

            // 1) 쿠키가 이미 있으면 (재로그인·세션유지 케이스) 바로 성공
            var cookieVal = await GetCookieAsync(ws, CookieName, ct);
            if (!string.IsNullOrEmpty(cookieVal))
            {
                onStatus("세션 획득 완료");
                return new(true, "ok", cookieVal);
            }

            // 2) 로그인 폼 감지 — userId 필드 존재 여부
            if (!filled)
            {
                var probe = await EvaluateAsync(ws,
                    "(function(){var u=document.getElementById('userId');var p=document.getElementById('password');return (u&&p)?'Y':'N';})()",
                    ct);

                if (probe != "Y")
                {
                    if (probes % 3 == 0) onStatus($"로그인 폼 감지 중... ({probes}s)");
                    continue;
                }

                onStatus("자격증명 자동 입력 중...");
                // ID/PW만 채움 — 클릭은 사용자가 직접 (CDP Input 이벤트가 Cognito 폼에서 동작 안 함)
                var fillScript = $@"(function(){{
                    try {{
                        var u = document.getElementById('userId');
                        var p = document.getElementById('password');
                        if(!u || !p) return 'NO_FORM';
                        var proto  = window.HTMLInputElement && window.HTMLInputElement.prototype;
                        var desc   = proto ? Object.getOwnPropertyDescriptor(proto, 'value') : null;
                        var setter = desc ? desc.set : null;
                        if (setter) {{ setter.call(u, {idJson}); setter.call(p, {pwJson}); }}
                        else        {{ u.value = {idJson};        p.value = {pwJson};        }}
                        ['input','change','blur'].forEach(function(ev){{
                            u.dispatchEvent(new Event(ev, {{bubbles:true}}));
                            p.dispatchEvent(new Event(ev, {{bubbles:true}}));
                        }});
                        return 'OK';
                    }} catch(e) {{ return 'ERR:' + (e && e.message ? e.message : e); }}
                }})()";
                var fillRes = await EvaluateAsync(ws, fillScript, ct);

                if (fillRes == "OK")
                {
                    filled = true;
                    onStatus("✅ ID/PW 입력 완료 — Chrome 창에서 [로그인] 버튼을 눌러주세요.");
                }
                else
                {
                    onStatus($"자동 입력 실패({fillRes}) — Chrome 창에서 직접 입력 후 [로그인] 눌러주세요.");
                    filled = true; // 쿠키만 계속 대기
                }
            }
        }

        return new(false, "로그인 시간 초과 (3분) — 브라우저에서 직접 로그인 후 다시 시도하세요.", null);
    }

    private record BrowserInfo(string ExePath, string? AppBundlePath);

    // ── 브라우저 탐색 ────────────────────────────────────────────────────────
    private static BrowserInfo? FindBrowser()
    {
        (string exe, string? bundle)[] candidates =
        {
            ("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                "/Applications/Microsoft Edge.app"),
            ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                "/Applications/Google Chrome.app"),
            (@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", null),
            (@"C:\Program Files\Microsoft\Edge\Application\msedge.exe",        null),
            (@"C:\Program Files\Google\Chrome\Application\chrome.exe",         null),
            (@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",   null),
        };
        foreach (var (exe, bundle) in candidates)
            if (File.Exists(exe)) return new BrowserInfo(exe, bundle);
        return null;
    }

    // ── CDP 엔드포인트 ───────────────────────────────────────────────────────
    private static async Task<string> PollCdpJsonAsync(
        HttpClient http, CancellationToken ct, int timeoutMs = 12000)
    {
        Exception? last = null;
        for (int elapsed = 0; elapsed < timeoutMs; elapsed += 400)
        {
            ct.ThrowIfCancellationRequested();
            await Task.Delay(400, ct);
            try { return await http.GetStringAsync($"http://localhost:{CdpPort}/json"); }
            catch (Exception ex) { last = ex; }
        }
        throw new Exception(
            $"Edge CDP 응답 없음 (포트 {CdpPort}). 기존 Edge 창을 모두 닫고 다시 시도하세요. " +
            $"원인: {last?.Message}");
    }

    private static (string wsUrl, string targetId)? PickPageTarget(string tabsJson)
    {
        using var doc = JsonDocument.Parse(tabsJson);
        (string, string)? fallback = null;
        foreach (var item in doc.RootElement.EnumerateArray())
        {
            if (!item.TryGetProperty("type", out var t) || t.GetString() != "page") continue;
            var url = item.TryGetProperty("url", out var uel) ? uel.GetString() ?? "" : "";
            var wsUrl = item.TryGetProperty("webSocketDebuggerUrl", out var w) ? w.GetString() : null;
            var tid   = item.TryGetProperty("id", out var ie) ? ie.GetString() : null;
            if (wsUrl is null || tid is null) continue;
            if (url.Contains("wayble", StringComparison.OrdinalIgnoreCase)) return (wsUrl, tid);
            fallback ??= (wsUrl, tid);
        }
        return fallback;
    }

    // ── CDP 쿠키 조회 ────────────────────────────────────────────────────────
    private static async Task<string?> GetCookieAsync(
        ClientWebSocket ws, string name, CancellationToken ct)
    {
        var resp = await SendAsync(ws, "Network.getCookies",
            new { urls = new[] { CookieHost } }, ct);
        using var doc = JsonDocument.Parse(resp);
        if (!doc.RootElement.TryGetProperty("result", out var res)) return null;
        if (!res.TryGetProperty("cookies", out var arr))            return null;
        foreach (var c in arr.EnumerateArray())
        {
            if (c.TryGetProperty("name", out var n) && n.GetString() == name)
                return c.TryGetProperty("value", out var v) ? v.GetString() : null;
        }
        return null;
    }

    // ── CDP Runtime.evaluate ─────────────────────────────────────────────────
    private static async Task<string> EvaluateAsync(
        ClientWebSocket ws, string script, CancellationToken ct)
    {
        var resp = await SendAsync(ws, "Runtime.evaluate",
            new { expression = script, returnByValue = true }, ct);
        try
        {
            using var doc = JsonDocument.Parse(resp);
            return doc.RootElement
                .GetProperty("result")
                .GetProperty("result")
                .GetProperty("value")
                .GetString() ?? "";
        }
        catch { return ""; }
    }

    // ── CDP 공통 Send (id 매칭, 이벤트 메시지 스킵, fragmentation 처리) ─────
    private static async Task<string> SendAsync(
        ClientWebSocket ws, string method, object? parameters, CancellationToken ct)
    {
        int id = Interlocked.Increment(ref _cdpId);
        object payload = parameters is null
            ? (object)new { id, method }
            : (object)new { id, method, @params = parameters };
        var sendBuf = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(payload));
        await ws.SendAsync(sendBuf, WebSocketMessageType.Text, true, ct);

        var ms  = new MemoryStream();
        var buf = new byte[16384];
        while (true)
        {
            ms.SetLength(0);
            while (true)
            {
                var r = await ws.ReceiveAsync(buf, ct);
                ms.Write(buf, 0, r.Count);
                if (r.EndOfMessage) break;
            }
            var text = Encoding.UTF8.GetString(ms.ToArray());
            try
            {
                using var doc = JsonDocument.Parse(text);
                if (doc.RootElement.TryGetProperty("id", out var idEl) &&
                    idEl.ValueKind == JsonValueKind.Number &&
                    idEl.GetInt32() == id)
                    return text;
            }
            catch { }
            // 이벤트 알림 / 다른 id → 스킵 후 다음 메시지 대기
        }
    }
}
