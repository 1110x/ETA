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
using ETA.Models;

namespace ETA.Services;

/// <summary>
/// 측정인.kr CDP 자동입력 서비스 — 팝업 없이 Show3 패널에서 직접 제어
/// </summary>
public static class MeasurerCdpService
{
    private static readonly object LogLock = new();
    private static readonly string LogPath = ResolveLogPath();

    private static string ResolveLogPath()
    {
        try
        {
            string d = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            if (Directory.Exists(d)) return Path.Combine(d, "측정in.log");
        }
        catch { }
        return Path.Combine(AppContext.BaseDirectory, "측정in.log");
    }

    private static void Log(string msg)
    {
        try
        {
            lock (LogLock)
            {
                var dir = Path.GetDirectoryName(LogPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch { }
        Debug.WriteLine($"[MeasurerCDP] {msg}");
    }

    // ──────────────────────────────────────────────────────────────────────
    //  페이지 확인
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// 현재 브라우저 측정인 탭의 제목을 읽어 시료와 매칭 결과를 반환.
    /// </summary>
    public static async Task<(string text, string colorHex)> CheckPageMatchAsync(
        string 약칭, string 시료명, string 채취일자)
    {
        try
        {
            string h2 = await GetEditTitleAsync();
            if (string.IsNullOrEmpty(h2))
                return ("⚠ 측정인 탭을 찾을 수 없습니다. 브라우저에서 측정인.kr을 먼저 열어주세요.", "#ffaa44");

            var parts = h2.Split(new[] { " / " }, StringSplitOptions.None);
            string pageDate = parts.Length > 2 ? parts[2].Trim() : "";
            string pageCompany = parts.Length > 3 ? parts[3].Trim() : "";
            string pageSample = parts.Length > 4 ? parts[4].Trim() : "";

            double companySim = Similarity(약칭, pageCompany);
            double sampleSim = Similarity(시료명, pageSample);
            bool dateOk = !string.IsNullOrEmpty(채취일자) && pageDate.StartsWith(채취일자);

            string text =
                $"[페이지] {h2}\n" +
                $"  {(dateOk ? "✅" : "⚠")} 날짜: {pageDate} | " +
                $"{(companySim > 0.5 ? "✅" : "⚠")} 업체: {pageCompany} ↔ {약칭} ({companySim * 100:F0}%) | " +
                $"{(sampleSim > 0.4 ? "✅" : "⚠")} 시료: {pageSample} ↔ {시료명} ({sampleSim * 100:F0}%)";

            bool allGood = dateOk && companySim > 0.5 && sampleSim > 0.4;
            return (text, allGood ? "#66cc88" : "#ffaa44");
        }
        catch (Exception ex)
        {
            return ($"⚠ 확인 오류: {ex.Message}", "#ee4444");
        }
    }

    private static async Task<string> GetEditTitleAsync()
    {
        const int port = 9222;
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
            string json = await http.GetStringAsync($"http://localhost:{port}/json");
            string? wsUrl = FindWsUrl(json);
            if (wsUrl == null) return "";

            using var socket = new ClientWebSocket();
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            var cmd = JsonSerializer.Serialize(new
            {
                id = 1,
                method = "Runtime.evaluate",
                @params = new { expression = "document.getElementById('edit_title')?.innerText||''", returnByValue = true }
            });
            await socket.SendAsync(Encoding.UTF8.GetBytes(cmd), WebSocketMessageType.Text, true, cts.Token);

            var buf = new byte[8192];
            var recv = await socket.ReceiveAsync(buf, cts.Token);
            using var doc = JsonDocument.Parse(Encoding.UTF8.GetString(buf, 0, recv.Count));
            return doc.RootElement.GetProperty("result").GetProperty("result").GetProperty("value").GetString() ?? "";
        }
        catch { return ""; }
    }

    // ──────────────────────────────────────────────────────────────────────
    //  자동입력
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// 현재 브라우저 측정인 탭에 MeasurerRow 목록을 자동 입력.
    /// onStatus: UI 스레드에서 상태 텍스트 업데이트 콜백.
    /// </summary>
    public static async Task<(bool ok, string message)> AutoInputAsync(
        IReadOnlyList<MeasurerRow> rows,
        Action<string> onStatus)
    {
        const int port = 9222;
        Log("AutoInput 시작");
        onStatus("측정인 탭 연결 중...");

        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            string json = await http.GetStringAsync($"http://localhost:{port}/json");
            string? wsUrl = FindWsUrl(json);
            if (wsUrl == null)
                return (false, "측정인 탭을 찾지 못했습니다.");

            var payload = rows.Select(r => new { analyte = r.항목명, result = r.결과값 }).ToList();
            string dataJs = JsonSerializer.Serialize(payload)
                .Replace("\\", "\\\\").Replace("'", "\\'");

            using var socket = new ClientWebSocket();
            int timeout = Math.Clamp(30 + (rows.Count * 3), 90, 600);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeout));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            int cmdId = 0;

            async Task<JsonDocument> CdpSend(string method, object parms)
            {
                var id = ++cmdId;
                var bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(
                    new Dictionary<string, object> { ["id"] = id, ["method"] = method, ["params"] = parms }));
                await socket.SendAsync(bytes, WebSocketMessageType.Text, true, cts.Token);
                for (int i = 0; i < 200; i++)
                {
                    using var ms = new MemoryStream();
                    WebSocketReceiveResult r;
                    var seg = new byte[65536];
                    do
                    {
                        r = await socket.ReceiveAsync(new ArraySegment<byte>(seg), cts.Token);
                        if (r.MessageType == WebSocketMessageType.Close)
                            throw new InvalidOperationException("CDP WebSocket closed");
                        ms.Write(seg, 0, r.Count);
                    } while (!r.EndOfMessage);
                    var doc = JsonDocument.Parse(ms.ToArray());
                    if (doc.RootElement.TryGetProperty("id", out var idProp) && idProp.GetInt32() == id)
                        return doc;
                    doc.Dispose();
                }
                throw new TimeoutException($"CDP 응답 대기 초과 id={id}");
            }

            async Task<string> JsEval(string expr, bool isAsync = false)
            {
                using var doc = await CdpSend("Runtime.evaluate",
                    new { expression = expr, returnByValue = true, awaitPromise = isAsync });
                var res = doc.RootElement.GetProperty("result");
                if (res.TryGetProperty("exceptionDetails", out var ex))
                    throw new Exception($"JS 오류: {ex}");
                var v = res.GetProperty("result");
                if (v.TryGetProperty("value", out var val))
                    return val.ValueKind == JsonValueKind.String ? (val.GetString() ?? "") : val.GetRawText();
                return "";
            }

            async Task Click(double x, double y)
            {
                await CdpSend("Input.dispatchMouseEvent", new { type = "mousePressed", x, y, button = "left", clickCount = 1 });
                await CdpSend("Input.dispatchMouseEvent", new { type = "mouseReleased", x, y, button = "left", clickCount = 1 });
            }

            async Task SendTab()
            {
                await CdpSend("Input.dispatchKeyEvent", new { type = "keyDown", key = "Tab", code = "Tab", windowsVirtualKeyCode = 9 });
                await CdpSend("Input.dispatchKeyEvent", new { type = "keyUp",   key = "Tab", code = "Tab", windowsVirtualKeyCode = 9 });
            }

            async Task sendEnter()
            {
                await CdpSend("Input.dispatchKeyEvent", new { type = "keyDown", key = "Enter", code = "Enter", windowsVirtualKeyCode = 13 });
                await CdpSend("Input.dispatchKeyEvent", new { type = "keyUp",   key = "Enter", code = "Enter", windowsVirtualKeyCode = 13 });
            }

            // 현재 포커스된 엘리먼트 JS click (드롭다운 확정 등)
            async Task Sendclick()
            {
                await JsEval(@"(function(){var e=document.getElementById('$$_rg_editor648');if(e)e.click();})()");
                await Task.Delay(30, cts.Token);
            }

            async Task SendDblclick()
            {
                await JsEval(@"(function(){
                    var e=document.getElementById('$$_rg_editor648');if(!e)return;
                    var o={bubbles:true,cancelable:true,view:window};
                    e.dispatchEvent(new MouseEvent('mousedown',o));
                    e.dispatchEvent(new MouseEvent('mouseup',o));
                    e.dispatchEvent(new MouseEvent('click',o));
                    e.dispatchEvent(new MouseEvent('mousedown',o));
                    e.dispatchEvent(new MouseEvent('mouseup',o));
                    e.dispatchEvent(new MouseEvent('click',o));
                    e.dispatchEvent(new MouseEvent('dblclick',o));
                })()");
                await Task.Delay(30, cts.Token);
            }


            // Phase 0: DOM 높이 확장
            onStatus("DOM 준비 중...");
            await JsEval(@"(function(){var g=document.getElementById('gridAnalySampAnzeDataAirItemList1');if(!g)return;var s={};s.rH=g.style.height;g.style.height='5000%';var inn=g.querySelector('.rg-exact');if(inn){s.iH=inn.style.height;inn.style.height='4200px';}var p=g.closest('div[style*=""overflow""]')||g.parentElement;if(p){s.pO=p.style.overflow;p.style.overflow='visible';}var vp=g.querySelector('.rg-viewport,.rg-body');if(vp){s.vH=vp.style.height;s.vO=vp.style.overflow;vp.style.height='99999px';vp.style.overflow='visible';}window.__etaH=s;})()");
            await Task.Delay(300, cts.Token);

            // Phase 1: JS 매칭
            onStatus("항목 매칭 중...");
            string findJs = $@"(function(){{function n(s){{return(s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,'');}}var rows=JSON.parse('{dataJs}');var tb=document.evaluate(""//*[@id='gridAnalySampAnzeDataAirItemList1']/div/div[1]/div/div[1]/div[1]/div[4]/table/tbody"",document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;var trs=[];if(tb)trs=Array.from(tb.children).filter(function(x){{return x&&x.tagName==='TR'&&x.querySelector('td:nth-child(4)')&&x.querySelector('td:nth-child(6)');}});if(!trs.length){{var sc=document.getElementById('gridAnalySampAnzeDataAirItemList1')||document;trs=Array.from(sc.querySelectorAll('table tbody tr')).filter(function(tr){{return!!(tr.querySelector('td:nth-child(4)')&&tr.querySelector('td:nth-child(6)'));}}); }}var rb={{}},rl=[];for(var i=0;i<rows.length;i++){{var k=n(rows[i].analyte);if(!k||k===n('현장측정및시료채취'))continue;if(!(k in rb)){{rb[k]=rows[i];rl.push(k);}}}}var matched=[],used={{}},unmatched=[];for(var r=0;r<trs.length;r++){{var d=trs[r].querySelector('td:nth-child(4) div')||trs[r].querySelector('td:nth-child(4)');var lk=n((d.innerText||d.textContent||'').trim());if(!lk)continue;var item=rb[lk],mk=lk;if(!item){{var fk=rl.find(function(k){{return k.indexOf(lk)>=0||lk.indexOf(k)>=0;}});if(fk){{item=rb[fk];mk=fk;}}}}if(!item||used[mk])continue;used[mk]=true;matched.push({{a:item.analyte,v:(item.result||'').toString().trim(),ri:r}});}}for(var i=0;i<rl.length;i++){{if(!used[rl[i]])unmatched.push(rb[rl[i]].analyte);}}return JSON.stringify({{matched:matched,total:rows.length,domRows:trs.length,matchedCount:matched.length,unmatched:unmatched.slice(0,5)}});}})()"; ;

            string p1json = await JsEval(findJs);
            int total = 0, domRows = 0, matchedCount = 0, filled = 0, unmatchedCount = 0;
            var commitFailed = new List<string>();
            var valuePreview = new List<string>();
            string unmatchedPreview = "";
            var matchedItems = new List<(string analyte, string value, int rowIndex)>();

            using (var p1 = JsonDocument.Parse(p1json))
            {
                var r = p1.RootElement;
                total = r.GetProperty("total").GetInt32();
                domRows = r.GetProperty("domRows").GetInt32();
                matchedCount = r.GetProperty("matchedCount").GetInt32();
                if (r.TryGetProperty("unmatched", out var um) && um.ValueKind == JsonValueKind.Array)
                {
                    var ul = um.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => x.Length > 0).ToList();
                    unmatchedCount = ul.Count;
                    unmatchedPreview = string.Join(", ", ul.Take(5));
                }
                foreach (var m in r.GetProperty("matched").EnumerateArray())
                    matchedItems.Add((m.GetProperty("a").GetString() ?? "", m.GetProperty("v").GetString() ?? "", m.GetProperty("ri").GetInt32()));
            }

            onStatus($"매칭 완료: {matchedCount}/{total}건 — 입력 중...");
            Log($"매칭 {matchedCount}/{total}, DOM행 {domRows}");

            // Phase 2: 셀별 입력
            foreach (var (analyte, value, rowIndex) in matchedItems)
            {
                string safe = analyte.Replace("\\", "\\\\").Replace("'", "\\'");

                // td[6] 좌표 조회 → CDP Click으로 결과값 셀 활성화
                string getCoordJs = $@"(function(){{function n(s){{return(s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,''); }}var t=n('{safe}');var tb=document.evaluate(""//*[@id='gridAnalySampAnzeDataAirItemList1']/div/div[1]/div/div[1]/div[1]/div[4]/table/tbody"",document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;if(!tb)return JSON.stringify({{e:'noTbody'}});var trs=Array.from(tb.children).filter(function(x){{return x&&x.tagName==='TR'&&x.querySelector('td:nth-child(4)')&&x.querySelector('td:nth-child(6)');}});function gc(tr){{var t5=tr.querySelector('td:nth-child(5)');if(!t5)return JSON.stringify({{e:'noTd5'}});t5.scrollIntoView({{block:'center',behavior:'instant'}});var r5=t5.getBoundingClientRect();return JSON.stringify({{x:r5.x+r5.width/2,y:r5.y+r5.height/2}});}}if({rowIndex}<trs.length){{var d4=trs[{rowIndex}].querySelector('td:nth-child(4) div')||trs[{rowIndex}].querySelector('td:nth-child(4)');if(n((d4.innerText||d4.textContent||'').trim())===t)return gc(trs[{rowIndex}]);}}for(var i=0;i<trs.length;i++){{var d=trs[i].querySelector('td:nth-child(4) div')||trs[i].querySelector('td:nth-child(4)');if(n((d.innerText||d.textContent||'').trim())===t)return gc(trs[i]);}}return JSON.stringify({{e:'notFound'}});}})()";
                double x, y;
                try
                {
                    string cj = await JsEval(getCoordJs);
                    using var cd = JsonDocument.Parse(cj);
                    if (cd.RootElement.TryGetProperty("e", out _)) { commitFailed.Add(analyte); continue; }
                    x = cd.RootElement.GetProperty("x").GetDouble();
                    y = cd.RootElement.GetProperty("y").GetDouble();
                    if (x <= 0 || y <= 0) { commitFailed.Add(analyte); continue; }
                }
                catch { commitFailed.Add(analyte); continue; }

                await Click(x, y); await Task.Delay(80, cts.Token);

                try
                {
                    var row = rows.FirstOrDefault(r => r.항목명 == analyte);

                    // td[6]이 활성화된 상태에서 시작 — Tab으로 순차 이동
                    await CdpSend("Input.insertText", new { text ="1" });           await SendTab();
                    await CdpSend("Input.insertText", new { text = value });        await SendTab(); // td[6]  결과값
                    await SendTab();                                                                 // td[7]  skip
                    await CdpSend("Input.insertText", new { text = row.측정방법 });  await SendTab();  // td[8]  측정방법
                    await CdpSend("Input.insertText", new { text = row.장비명 });    await SendTab();  // td[9]  장비명
                    await CdpSend("Input.insertText", new { text = row.담당자 });    await SendTab(); // td[10] 담당자
                    await SendDblclick(); await CdpSend("Input.insertText", new { text = row.시작일.Replace("-","") }); await SendTab(); // td[11] 시작일
                    await CdpSend("Input.insertText", new { text = row.시작시간 });  await SendTab();  // td[12] 시작시간
                    await SendDblclick(); await CdpSend("Input.insertText", new { text = row.종료일.Replace("-","") }); await SendTab(); // td[13] 종료일
                    await CdpSend("Input.insertText", new { text = row.종료시간 });  await SendTab(); // td[14] 종료시간
                    await CdpSend("Input.insertText", new { text = "100" });                                // td[15] 샘플사용량


                    filled++;
                    if (valuePreview.Count < 5) valuePreview.Add($"[{analyte}→'{value}']");
                    onStatus($"입력 중... {filled}/{matchedCount}건");
                    Log($"입력완료: {analyte}={value}");
                }
                catch (Exception ex) { commitFailed.Add(analyte); Log($"입력오류 {analyte}: {ex.Message}"); }
            }

            // 마지막 커밋: 첫 번째 td[4] 클릭으로 포커스 이동 (JS click)
            try { await JsEval(@"(function(){var tb=document.evaluate(""//*[@id='gridAnalySampAnzeDataAirItemList1']/div/div[1]/div/div[1]/div[1]/div[4]/table/tbody"",document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;if(!tb)return;var td=tb.querySelector('tr td:nth-child(4)');if(td)td.click();})()"); } catch { }

            // Phase 3: 높이 원복
            try { await JsEval(@"(function(){var g=document.getElementById('gridAnalySampAnzeDataAirItemList1');if(!g||!window.__etaH)return;var s=window.__etaH;g.style.height=s.rH||'';var inn=g.querySelector('.rg-exact');if(inn&&s.iH!==undefined)inn.style.height=s.iH;var p=g.closest('div[style*=""overflow""]')||g.parentElement;if(p&&s.pO!==undefined)p.style.overflow=s.pO;var vp=g.querySelector('.rg-viewport,.rg-body');if(vp){if(s.vH!==undefined)vp.style.height=s.vH;if(s.vO!==undefined)vp.style.overflow=s.vO;}delete window.__etaH;window.dispatchEvent(new Event('resize'));})()"); } catch { }

            string summary = $"전체 {total} / DOM행 {domRows} / 매칭 {matchedCount} / 입력완료 {filled}";
            if (unmatchedCount > 0) summary += $" / 미매칭 {unmatchedCount}건: {unmatchedPreview}";
            if (commitFailed.Count > 0) summary += $" / 실패 {commitFailed.Count}건: {string.Join(",", commitFailed.Take(3))}";

            Log($"완료: {summary}");
            return (filled > 0, summary);
        }
        catch (Exception ex)
        {
            Log($"AutoInput 예외: {ex}");
            return (false, ex.Message);
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    //  공용 헬퍼
    // ──────────────────────────────────────────────────────────────────────

    private static string? FindWsUrl(string cdpJson)
    {
        using var doc = JsonDocument.Parse(cdpJson);
        // 1순위: 측정인 탭
        foreach (var item in doc.RootElement.EnumerateArray())
            if (item.TryGetProperty("type", out var t) && t.GetString() == "page" &&
                item.TryGetProperty("url", out var u) && (u.GetString() ?? "").Contains("측정인") &&
                item.TryGetProperty("webSocketDebuggerUrl", out var ws))
                return ws.GetString();
        // 2순위: 첫 번째 page 탭
        foreach (var item in doc.RootElement.EnumerateArray())
            if (item.TryGetProperty("type", out var t) && t.GetString() == "page" &&
                item.TryGetProperty("webSocketDebuggerUrl", out var ws))
                return ws.GetString();
        return null;
    }

    private static double Similarity(string a, string b)
    {
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;
        a = a.Trim(); b = b.Trim();
        if (a == b) return 1.0;
        var ba = new HashSet<string>(Bigrams(a));
        var bb = new HashSet<string>(Bigrams(b));
        int common = ba.Intersect(bb).Count();
        int total = ba.Count + bb.Count;
        return total == 0 ? 0.0 : (double)common * 2 / total;
    }

    private static IEnumerable<string> Bigrams(string s)
    {
        for (int i = 0; i < s.Length - 1; i++) yield return s.Substring(i, 2);
    }
}
