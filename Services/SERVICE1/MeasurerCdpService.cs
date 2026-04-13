using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Services.SERVICE1;

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
        if (App.EnableLogging)
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
        }
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
            string pageDate    = parts.Length > 2 ? parts[2].Trim() : "";
            string pageCompany = parts.Length > 3 ? parts[3].Trim() : "";
            string pageSample  = parts.Length > 4 ? parts[4].Trim() : "";

            double companySim = Similarity(약칭, pageCompany);
            double sampleSim  = Similarity(시료명, pageSample);
            bool   dateOk     = !string.IsNullOrEmpty(채취일자) && pageDate.StartsWith(채취일자);

            string text =
                $"[페이지] {h2}\n" +
                $"  {(dateOk      ? "✅" : "⚠")} 날짜: {pageDate} | " +
                $"{(companySim > 0.5 ? "✅" : "⚠")} 업체: {pageCompany} ↔ {약칭} ({companySim * 100:F0}%) | " +
                $"{(sampleSim  > 0.4 ? "✅" : "⚠")} 시료: {pageSample}  ↔ {시료명} ({sampleSim  * 100:F0}%)";

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
            using var cts    = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            var cmd = JsonSerializer.Serialize(new
            {
                id = 1,
                method = "Runtime.evaluate",
                @params = new { expression = "document.getElementById('edit_title')?.innerText||''", returnByValue = true }
            });
            await socket.SendAsync(Encoding.UTF8.GetBytes(cmd), WebSocketMessageType.Text, true, cts.Token);

            var buf  = new byte[8192];
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
            Log($"[CDP연결] {wsUrl}");

            var payload = rows.Select(r => new { analyte = r.항목명, result = r.결과값 }).ToList();
            string dataJs = JsonSerializer.Serialize(payload)
                .Replace("\\", "\\\\").Replace("'", "\\'");

            using var socket = new ClientWebSocket();
            int timeout = Math.Clamp(30 + (rows.Count * 3), 90, 600);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeout));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            int cmdId = 0;

            // ── CDP 저수준 송수신 ──────────────────────────────────────────
            async Task<JsonDocument> CdpSend(string method, object parms)
            {
                var id    = ++cmdId;
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

            // ── JS 평가 헬퍼 ──────────────────────────────────────────────
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


            // ── Phase 1: JS 매칭 ──────────────────────────────────────────
            onStatus("항목 매칭 중...");
            string findJs = $@"(function(){{
                function n(s){{ return (s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,''); }}
                var rows = JSON.parse('{dataJs}');
                var tb   = document.evaluate(""//*[@id='gridAnalySampAnzeDataAirItemList1']/div/div[1]/div/div[1]/div[1]/div[4]/table/tbody"",
                    document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                var trs = [];
                if(tb) trs = Array.from(tb.children).filter(function(x){{
                    return x && x.tagName==='TR' && x.querySelector('td:nth-child(4)') && x.querySelector('td:nth-child(6)');
                }});
                if(!trs.length){{
                    var sc = document.getElementById('gridAnalySampAnzeDataAirItemList1') || document;
                    trs = Array.from(sc.querySelectorAll('table tbody tr')).filter(function(tr){{
                        return !!(tr.querySelector('td:nth-child(4)') && tr.querySelector('td:nth-child(6)'));
                    }});
                }}
                var rb = {{}}, rl = [];
                for(var i = 0; i < rows.length; i++){{
                    var k = n(rows[i].analyte);
                    if(!k || k === n('현장측정및시료채취')) continue;
                    if(!(k in rb)){{ rb[k] = rows[i]; rl.push(k); }}
                }}
                var matched = [], used = {{}}, unmatched = [];
                for(var r = 0; r < trs.length; r++){{
                    var d   = trs[r].querySelector('td:nth-child(4) div') || trs[r].querySelector('td:nth-child(4)');
                    var lk  = n((d.innerText || d.textContent || '').trim());
                    if(!lk) continue;
                    var item = rb[lk], mk = lk;
                    if(!item){{
                        var fk = rl.find(function(k){{ return k.indexOf(lk) >= 0 || lk.indexOf(k) >= 0; }});
                        if(fk){{ item = rb[fk]; mk = fk; }}
                    }}
                    if(!item || used[mk]) continue;
                    used[mk] = true;
                    matched.push({{ a: item.analyte, v: (item.result||'').toString().trim(), ri: r }});
                }}
                for(var i = 0; i < rl.length; i++){{ if(!used[rl[i]]) unmatched.push(rb[rl[i]].analyte); }}
                return JSON.stringify({{ matched: matched, total: rows.length, domRows: trs.length, matchedCount: matched.length, unmatched: unmatched.slice(0,5) }});
            }})()";

            string p1json = await JsEval(findJs);
            int total = 0, domRows = 0, matchedCount = 0, filled = 0, unmatchedCount = 0;
            var commitFailed  = new List<string>();
            var valuePreview  = new List<string>();
            string unmatchedPreview = "";
            var matchedItems  = new List<(string analyte, string value, int rowIndex)>();

            using (var p1 = JsonDocument.Parse(p1json))
            {
                var r = p1.RootElement;
                total        = r.GetProperty("total").GetInt32();
                domRows      = r.GetProperty("domRows").GetInt32();
                matchedCount = r.GetProperty("matchedCount").GetInt32();
                if (r.TryGetProperty("unmatched", out var um) && um.ValueKind == JsonValueKind.Array)
                {
                    var ul = um.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => x.Length > 0).ToList();
                    unmatchedCount   = ul.Count;
                    unmatchedPreview = string.Join(", ", ul.Take(5));
                }
                foreach (var m in r.GetProperty("matched").EnumerateArray())
                    matchedItems.Add((m.GetProperty("a").GetString() ?? "", m.GetProperty("v").GetString() ?? "", m.GetProperty("ri").GetInt32()));
            }

            onStatus($"매칭 완료: {matchedCount}/{total}건 — 입력 중...");
            Log($"매칭 {matchedCount}/{total}, DOM행 {domRows}");

            // ── Phase 2: RealGrid dp.setValue 직접 입력 ───────────────────
            // 전체 rows를 JS에 전달 — dp.getRowCount()로 숨겨진 행 포함 전체 처리
            var inputPayload = rows.Select(row =>
            {
                if (row == null) return null;

                string 시작일   = row.시작일;
                string 종료일   = row.종료일;
                string 시작시간 = row.시작시간;
                string 종료시간 = row.종료시간;

                if (DateTime.TryParse(row.시작일,   out var dtS)) 시작일   = dtS.ToString("yyyy-MM-dd");
                if (DateTime.TryParse(row.종료일,   out var dtE)) 종료일   = dtE.ToString("yyyy-MM-dd");
                if (DateTime.TryParse(row.시작시간, out var tmS)) 시작시간 = tmS.ToString("HH:mm");
                if (DateTime.TryParse(row.종료시간, out var tmE)) 종료시간 = tmE.ToString("HH:mm");

                return new
                {
                    analyte  = row.항목명,
                    val      = row.결과값,
                    mthd     = row.측정방법,
                    equip    = row.장비명,
                    mgr      = row.담당자,
                    startDt  = 시작일,
                    startTm  = 시작시간,
                    endDt    = 종료일,
                    endTm    = 종료시간,
                };
            }).Where(x => x != null).ToList();

            string inputJson = JsonSerializer.Serialize(inputPayload)
                .Replace("\\", "\\\\").Replace("'", "\\'");

            string p2result = await JsEval($@"(function(){{
                try {{
                    var gv = RealGrid.getGridInstance('gridAnalySampAnzeDataAirItemList1');
                    if(!gv) return JSON.stringify({{error:'no-grid'}});
                    var dp = gv.getDataSource();
                    if(!dp) return JSON.stringify({{error:'no-dp'}});

                    function norm(s){{ return (s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,''); }}

                    // anze_mthd 텍스트→코드 변환 맵
                    // 1순위: 컬럼 labels/values 배열
                    var mthdMap = {{}};
                    var mthdCol = gv.columnByName('anze_mthd');
                    if(mthdCol && mthdCol.values && mthdCol.labels){{
                        for(var mi = 0; mi < mthdCol.labels.length; mi++){{
                            if(mthdCol.labels[mi]) mthdMap[norm(mthdCol.labels[mi])] = mthdCol.values[mi];
                        }}
                    }}
                    // 2순위: dp 기존 행 데이터 (anze_mthd_lb → anze_mthd)
                    var rc = dp.getRowCount();
                    for(var ri = 0; ri < rc; ri++){{
                        var dr = dp.getJsonRow(ri);
                        if(dr.anze_mthd && dr.anze_mthd_lb)
                            mthdMap[norm(dr.anze_mthd_lb)] = dr.anze_mthd;
                    }}

                    var inputs = JSON.parse('{inputJson}');
                    var inputMap = {{}};
                    inputs.forEach(function(r){{ inputMap[norm(r.analyte)] = r; }});

                    var rowCount = dp.getRowCount();
                    var filled = [], failed = [];

                    for(var i = 0; i < rowCount; i++){{
                        var row = dp.getJsonRow(i);
                        var nm  = norm(row.anze_item_nm || '');
                        var match = inputMap[nm];
                        if(!match){{
                            var keys = Object.keys(inputMap);
                            for(var j = 0; j < keys.length; j++){{
                                if(nm.indexOf(keys[j]) >= 0 || keys[j].indexOf(nm) >= 0){{
                                    match = inputMap[keys[j]]; break;
                                }}
                            }}
                        }}
                        if(!match) continue;

                        try {{
                            dp.setValue(i, 'allow_val',     '1');
                            dp.setValue(i, 'anze_val',      match.val);
                            dp.setValue(i, 'anze_start_dt', match.startDt);
                            dp.setValue(i, 'anze_start_tm', match.startTm);
                            dp.setValue(i, 'anze_end_dt',   match.endDt);
                            dp.setValue(i, 'anze_end_tm',   match.endTm);
                            dp.setValue(i, 'smpl_vol',      '100');
                            if(match.mthd){{
                                var mthdCode = mthdMap[norm(match.mthd)];
                                if(mthdCode) {{
                                    dp.setValue(i, 'anze_mthd',    mthdCode);
                                    dp.setValue(i, 'anze_mthd_lb', match.mthd);
                                }}
                            }}
                            if(match.equip) dp.setValue(i, 'anze_equip_no', match.equip);
                            if(match.mgr)   dp.setValue(i, 'anze_login_id', match.mgr);
                            filled.push(row.anze_item_nm);
                        }} catch(e) {{
                            failed.push(row.anze_item_nm + ':' + e.message);
                        }}
                    }}

                    return JSON.stringify({{filled: filled.length, items: filled, failed: failed}});
                }} catch(ex) {{
                    return JSON.stringify({{error: ex.message}});
                }}
            }})()");

            Log($"[Phase2] {p2result}");
            using (var p2 = JsonDocument.Parse(p2result))
            {
                var r2 = p2.RootElement;
                if (r2.TryGetProperty("filled", out var fc)) filled = fc.GetInt32();
                if (r2.TryGetProperty("items",  out var it))
                    foreach (var item in it.EnumerateArray())
                    {
                        var s = item.GetString() ?? "";
                        if (valuePreview.Count < 5) valuePreview.Add(s);
                    }
                if (r2.TryGetProperty("failed", out var ft))
                    foreach (var item in ft.EnumerateArray())
                        commitFailed.Add(item.GetString() ?? "");
            }
            onStatus($"입력 완료: {filled}/{matchedCount}건");


            string summary = $"전체 {total}건 중 입력완료 {filled}건";
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
    //  분석장비 스크랩 (현재 열린 측정인 페이지에서)
    // ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// 현재 브라우저에서 장비 드롭다운(rg-dropdown-list-0)을 스크랩하여 저장.
    /// onStatus: 진행 상태 콜백 (UI 스레드에서 호출).
    /// 반환: (저장 건수, 오류 메시지 or null)
    /// </summary>
    public static async Task<(int saved, string? error)> ScrapeEquipmentsAsync(
        Action<string> onStatus)
    {
        const int port = 9222;
        const string script = @"(function(){
            function isVisible(el) {
                if (!el) return false;
                var st = getComputedStyle(el);
                return el.offsetParent !== null && st.display !== 'none' && st.visibility !== 'hidden';
            }
            var lists = Array.from(document.querySelectorAll('div.rg-dropdownlist#rg-dropdown-list-0'));
            var list = lists.find(isVisible) || lists[lists.length - 1] || null;
            if (!list) return '[]';
            var opts = Array.from(list.querySelectorAll('.rg-dropdown-item[role=""option""]'));
            var out = [];
            for (var i = 0; i < opts.length; i++) {
                var o = opts[i];
                var labelEl = o.querySelector('.rg-dropdown-label') || o;
                var name = (labelEl.innerText || labelEl.textContent || '').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
                if (!name) continue;
                var code = (o.id || '').trim();
                if (!code) code = name;
                out.push({name:name, code:code});
            }
            var seen={}, dedup=[];
            for (var j=0;j<out.length;j++){
                if(seen[out[j].code]) continue;
                seen[out[j].code]=true; dedup.push(out[j]);
            }
            return JSON.stringify(dedup);
        })()";

        try
        {
            onStatus("브라우저 연결 중...");
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
            string cdpJson = await http.GetStringAsync($"http://localhost:{port}/json");
            string? wsUrl = FindWsUrl(cdpJson);
            if (wsUrl == null)
                return (0, "측정인 탭을 찾지 못했습니다. 브라우저에서 측정인.kr을 열어주세요.");

            onStatus("페이지에서 장비 목록 읽는 중...");
            using var socket = new ClientWebSocket();
            using var cts    = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            var cmdObj = JsonSerializer.Serialize(new
            {
                id = 1, method = "Runtime.evaluate",
                @params = new { expression = script, returnByValue = true }
            });
            await socket.SendAsync(Encoding.UTF8.GetBytes(cmdObj), WebSocketMessageType.Text, true, cts.Token);

            var buf = new byte[131072];
            var recv = await socket.ReceiveAsync(buf, cts.Token);
            string raw = Encoding.UTF8.GetString(buf, 0, recv.Count);

            // 결과 JSON 추출
            string json = "";
            using (var doc = JsonDocument.Parse(raw))
            {
                if (doc.RootElement.TryGetProperty("result", out var r1) &&
                    r1.TryGetProperty("result", out var r2) &&
                    r2.TryGetProperty("value", out var v))
                    json = v.GetString() ?? "";
            }

            if (string.IsNullOrWhiteSpace(json) || json == "[]")
                return (0, "장비 목록을 찾지 못했습니다.\n장비 드롭다운을 먼저 열어놓은 상태에서 시도하세요.");

            var items = new List<(string 장비명, string 코드값)>();
            using (var doc = JsonDocument.Parse(json))
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    string n = item.GetProperty("name").GetString() ?? "";
                    string c = item.GetProperty("code").GetString() ?? "";
                    if (!string.IsNullOrWhiteSpace(n)) items.Add((n, c));
                }

            onStatus($"{items.Count}개 파싱 완료 — DB 저장 중...");
            int saved = MeasurerService.SaveEquipments(items);
            Log($"[ScrapeEquipments] {saved}/{items.Count}개 저장");
            return (saved, null);
        }
        catch (Exception ex)
        {
            Log($"[ScrapeEquipments] 오류: {ex.Message}");
            return (0, ex.Message);
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    //  공용 헬퍼
    // ──────────────────────────────────────────────────────────────────────

    private static string? FindWsUrl(string cdpJson)
    {
        using var doc = JsonDocument.Parse(cdpJson);
        // 1순위: 측정인 탭 — 한글 URL 또는 퓨니코드(xn--lu5b7kx8m.kr), DevTools 창 제외
        foreach (var item in doc.RootElement.EnumerateArray())
        {
            if (!item.TryGetProperty("type", out var t) || t.GetString() != "page") continue;
            if (!item.TryGetProperty("webSocketDebuggerUrl", out var ws)) continue;
            var url = item.TryGetProperty("url", out var u) ? (u.GetString() ?? "") : "";
            if (url.StartsWith("devtools://") || url.StartsWith("chrome-extension://")) continue;
            if (url.Contains("측정인") || url.Contains("xn--lu5b7kx8m"))
                return ws.GetString();
        }
        // 2순위: 첫 번째 실제 page 탭 (devtools:// 제외)
        foreach (var item in doc.RootElement.EnumerateArray())
        {
            if (!item.TryGetProperty("type", out var t) || t.GetString() != "page") continue;
            if (!item.TryGetProperty("webSocketDebuggerUrl", out var ws)) continue;
            var url = item.TryGetProperty("url", out var u) ? (u.GetString() ?? "") : "";
            if (url.StartsWith("devtools://") || url.StartsWith("chrome-extension://")) continue;
            return ws.GetString();
        }
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
        int total  = ba.Count + bb.Count;
        return total == 0 ? 0.0 : (double)common * 2 / total;
    }

    private static IEnumerable<string> Bigrams(string s)
    {
        for (int i = 0; i < s.Length - 1; i++) yield return s.Substring(i, 2);
    }
}
