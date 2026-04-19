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

    public record PageMatchInfo(
        bool   ok,
        string message,
        string rawTitle,
        string pageDate,
        string pageCompany,
        string pageSample,
        double companySim,
        double sampleSim,
        bool   dateOk);

    /// <summary>측정인 페이지 #edit_title 을 분석해 구조화된 정보를 반환.
    /// pageSample(parts[0]) 은 "{견적번호} {시료명}" 형식이므로 앱측도 동일 조합으로 비교.</summary>
    public static async Task<PageMatchInfo> GetPageMatchInfoAsync(
        string 약칭, string 시료명, string 채취일자, string 견적번호 = "")
    {
        try
        {
            string h2 = await GetEditTitleAsync();
            if (string.IsNullOrEmpty(h2))
                return new PageMatchInfo(false, "측정인 탭을 찾을 수 없습니다. 브라우저에서 측정인.kr을 먼저 열어주세요.",
                    "", "", "", "", 0, 0, false);

            var parts = h2.Split(new[] { " / " }, StringSplitOptions.None);
            string pageDate    = parts.Length > 2 ? parts[2].Trim() : "";
            string pageCompany = parts.Length > 3 ? parts[3].Trim() : "";
            string pageSample  = parts.Length > 0 ? parts[0].Trim() : "";

            string appCombined = string.IsNullOrWhiteSpace(견적번호)
                ? 시료명
                : $"{견적번호} {시료명}".Trim();

            double companySim = Similarity(약칭, pageCompany);
            double sampleSim  = Similarity(appCombined, pageSample);
            bool   dateOk     = !string.IsNullOrEmpty(채취일자) && pageDate.StartsWith(채취일자);

            return new PageMatchInfo(true, "", h2, pageDate, pageCompany, pageSample, companySim, sampleSim, dateOk);
        }
        catch (Exception ex)
        {
            return new PageMatchInfo(false, $"확인 오류: {ex.Message}", "", "", "", "", 0, 0, false);
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


            // ── Phase 0-pre: 우클릭 컨텍스트 메뉴 차단 (사이트 정책 + 브라우저 기본) ──
            await JsEval(@"(function(){
                try {
                    if(window.__etaCtxMenuBlocked) return;
                    window.__etaCtxMenuBlocked = true;
                    var block = function(e){ e.stopPropagation(); e.preventDefault(); return false; };
                    ['contextmenu','auxclick'].forEach(function(t){
                        window.addEventListener(t, block, true);
                        document.addEventListener(t, block, true);
                    });
                    document.oncontextmenu = block;
                } catch(e){}
            })()");

            // ── Phase 0: 그리드 높이 부풀려 전체 행 DOM 렌더 (VBA 5000% 대응) ──
            //   RealGrid 가상 스크롤이라 Phase3 동시그룹 DOM 클릭이 15행만 잡히는 문제 해결
            string inflateJson = await JsEval(@"(function(){
                try {
                    var grid = document.getElementById('gridAnalySampAnzeDataAirItemList1');
                    if(!grid) return JSON.stringify({ok:false, why:'no-grid'});
                    var root = grid.querySelector('.rg-root') || grid;
                    window.__etaGridOrigHeight   = root.style.height || '';
                    window.__etaGridOrigOverflow = root.style.overflow || '';
                    var gv = RealGrid.getGridInstance('gridAnalySampAnzeDataAirItemList1');
                    var rc = gv ? gv.getDataSource().getRowCount() : 100;
                    root.style.height   = (rc * 32 + 200) + 'px';
                    root.style.overflow = 'visible';
                    if(gv && typeof gv.refresh === 'function') gv.refresh();
                    return JSON.stringify({ok:true, rowCount:rc});
                } catch(e){ return JSON.stringify({ok:false, why:e.message}); }
            })()");
            Log($"[Phase0] inflate: {inflateJson}");
            await Task.Delay(250);

            // ── Phase 0-dates: 폼 상단 시료접수일/분석시작일/분석종료일 직접 입력 ──
            //   smpl_rcpt_dt = 시료채취일, anze_start_dt = 시료채취일과 동일, anze_end_dt = 분석항목별 최대 종료일
            string rcptDt = "", startDt = "", endDt = "";
            if (rows.Count > 0)
            {
                rcptDt = (rows[0].채취일자 ?? "").Trim();
                if (DateTime.TryParse(rcptDt, out var rd)) rcptDt = rd.ToString("yyyy-MM-dd");

                DateTime? maxEnd = null;
                foreach (var r in rows)
                {
                    if (DateTime.TryParse(r.종료일, out var e) && (maxEnd == null || e > maxEnd)) maxEnd = e;
                }
                startDt = rcptDt;  // 분석시작일 = 시료채취일
                endDt   = maxEnd?.ToString("yyyy-MM-dd") ?? rcptDt;
            }
            Log($"[Phase0-dates] rcpt={rcptDt} start={startDt} end={endDt}");
            string datesJson = await JsEval($@"(function(){{
                function setInput(id, val){{
                    var el = document.getElementById(id);
                    if(!el) return 'miss:'+id;
                    el.value = val;
                    if(el.hasAttribute('data-value')) el.setAttribute('data-value', val);
                    ['input','change','blur'].forEach(function(t){{
                        try{{ el.dispatchEvent(new Event(t, {{bubbles:true}})); }}catch(e){{}}
                    }});
                    return 'ok:'+id+'='+el.value;
                }}
                var out = [];
                out.push(setInput('smpl_rcpt_dt',  '{rcptDt}'));
                out.push(setInput('anze_start_dt', '{startDt}'));
                out.push(setInput('anze_end_dt',   '{endDt}'));
                return out.join(' | ');
            }})()");
            Log($"[Phase0-dates] result: {datesJson}");

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
                    legal    = row.법적기준 ?? "",
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

                    // 현재 dp에 존재하는 필드명 세트 (대소문자 구분 없이 소문자 정규화)
                    var fieldSet = {{}};
                    try {{
                        var fs = dp.getFields ? dp.getFields() : [];
                        for(var fi = 0; fi < fs.length; fi++){{
                            var fn = fs[fi] && (fs[fi].fieldName || fs[fi].name);
                            if(fn) fieldSet[String(fn).toLowerCase()] = fn;
                        }}
                    }} catch(e){{}}
                    if(Object.keys(fieldSet).length === 0){{
                        try {{
                            var sample = dp.getJsonRow(0) || {{}};
                            for(var kk in sample) fieldSet[String(kk).toLowerCase()] = kk;
                        }} catch(e){{}}
                    }}
                    function setIf(i, fname, val){{
                        var lk = String(fname).toLowerCase();
                        if(!(lk in fieldSet)) return 'skip';
                        var actual = fieldSet[lk] || fname;
                        try {{ dp.setValue(i, actual, val); return 'ok'; }}
                        catch(e){{
                            try {{ dp.setValue(i, fname, val); return 'ok2'; }}
                            catch(e2){{ return 'err:'+e2.message; }}
                        }}
                    }}

                    function norm(s){{ return (s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,''); }}

                    // anze_mthd / anze_login_id 텍스트→코드 변환 맵
                    function buildLookup(colName, lbName){{
                        var map = {{}};
                        var col = gv.columnByName(colName);
                        if(col && col.values && col.labels){{
                            for(var mi = 0; mi < col.labels.length; mi++){{
                                if(col.labels[mi]) map[norm(col.labels[mi])] = col.values[mi];
                            }}
                        }}
                        // lookup 객체 계층 (RealGrid lookupData)
                        try {{
                            if(col && col.lookupData && col.lookupData.values && col.lookupData.labels){{
                                for(var mi2 = 0; mi2 < col.lookupData.labels.length; mi2++){{
                                    if(col.lookupData.labels[mi2])
                                        map[norm(col.lookupData.labels[mi2])] = col.lookupData.values[mi2];
                                }}
                            }}
                        }} catch(e){{}}
                        return map;
                    }}
                    var mthdMap  = buildLookup('anze_mthd',     'anze_mthd_lb');
                    var loginMap = buildLookup('anze_login_id', 'anze_login_id_lb');

                    // dp 기존 행 데이터에서 lb→code 수확
                    var rc = dp.getRowCount();
                    for(var ri = 0; ri < rc; ri++){{
                        var dr = dp.getJsonRow(ri);
                        if(dr.anze_mthd && dr.anze_mthd_lb)
                            mthdMap[norm(dr.anze_mthd_lb)] = dr.anze_mthd;
                        if(dr.anze_login_id && dr.anze_login_id_lb)
                            loginMap[norm(dr.anze_login_id_lb)] = dr.anze_login_id;
                    }}

                    var inputs = JSON.parse('{inputJson}');
                    var inputMap = {{}};
                    inputs.forEach(function(r){{ inputMap[norm(r.analyte)] = r; }});

                    var rowCount = dp.getRowCount();
                    var filled = [], failed = [], mgrFails = [];
                    var mgrDiag = {{ empty:0, set:0, readback:0, sample:[] }};

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
                        if(mgrDiag.sample.length < 5)
                            mgrDiag.sample.push({{it: row.anze_item_nm, mgr: match.mgr || ''}});

                        try {{
                            if(match.legal && String(match.legal).trim() !== '')
                                setIf(i, 'allow_val', String(match.legal).trim());
                            setIf(i, 'anze_val',      match.val);
                            setIf(i, 'anze_start_dt', match.startDt);
                            setIf(i, 'anze_start_tm', match.startTm);
                            setIf(i, 'anze_end_dt',   match.endDt);
                            setIf(i, 'anze_end_tm',   match.endTm);
                            setIf(i, 'smpl_vol',      '100');
                            if(match.mthd){{
                                var mthdCode = mthdMap[norm(match.mthd)];
                                if(mthdCode) {{
                                    setIf(i, 'anze_mthd',    mthdCode);
                                    setIf(i, 'anze_mthd_lb', match.mthd);
                                }}
                            }}
                            if(match.equip) setIf(i, 'anze_equip_no', match.equip);
                            if(match.mgr){{
                                var mgrKey = norm(match.mgr);
                                var loginCode = loginMap[mgrKey];
                                if(!loginCode){{
                                    var keys = Object.keys(loginMap);
                                    for(var lk = 0; lk < keys.length; lk++){{
                                        if(keys[lk].indexOf(mgrKey) >= 0 || mgrKey.indexOf(keys[lk]) >= 0){{
                                            loginCode = loginMap[keys[lk]]; break;
                                        }}
                                    }}
                                }}
                                if(!loginCode && mgrKey.length >= 2){{
                                    var pref = mgrKey.substr(0, 2);
                                    var keys2 = Object.keys(loginMap);
                                    for(var lk2 = 0; lk2 < keys2.length; lk2++){{
                                        if(keys2[lk2].indexOf(pref) >= 0){{
                                            loginCode = loginMap[keys2[lk2]]; break;
                                        }}
                                    }}
                                }}
                                if(loginCode){{
                                    setIf(i, 'anze_login_id',    loginCode);
                                    setIf(i, 'anze_login_id_lb', match.mgr);
                                    mgrDiag.set++;
                                    // 직후 read-back 검증
                                    var back = dp.getJsonRow(i);
                                    if(back.anze_login_id == loginCode) mgrDiag.readback++;
                                }} else {{
                                    mgrFails.push(match.mgr);
                                }}
                            }} else {{
                                mgrDiag.empty++;
                            }}
                            filled.push(row.anze_item_nm);
                        }} catch(e) {{
                            failed.push(row.anze_item_nm + ':' + e.message);
                        }}
                    }}

                    return JSON.stringify({{filled: filled.length, items: filled, failed: failed,
                        mthdMapSize: Object.keys(mthdMap).length,
                        loginMapSize: Object.keys(loginMap).length,
                        loginKeys: Object.keys(loginMap).slice(0,20),
                        mgrFails: mgrFails.slice(0,10),
                        mgrDiag: mgrDiag,
                        fieldSetSize: Object.keys(fieldSet).length,
                        fieldSample: Object.keys(fieldSet).slice(0,30)}});
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

            string groupApplied = "";

            // ── Phase 8: 분석자 컬럼 Phase2 직후 값 검증 ─────────────────
            try
            {
                string verify = await JsEval(@"(function(){
                    try {
                        var gv = RealGrid.getGridInstance('gridAnalySampAnzeDataAirItemList1');
                        if(!gv) return JSON.stringify({err:'no-grid'});
                        var dp = gv.getDataSource();
                        var rc = dp.getRowCount();
                        var populated = 0, empty = 0, samples = [];
                        for(var i = 0; i < rc; i++){
                            var r = dp.getJsonRow(i);
                            if(r.anze_login_id && String(r.anze_login_id).length > 0) populated++;
                            else empty++;
                            if(samples.length < 5)
                                samples.push({it: r.anze_item_nm, id: r.anze_login_id || '', lb: r.anze_login_id_lb || ''});
                        }
                        return JSON.stringify({populated:populated, empty:empty, samples:samples});
                    } catch(e){ return JSON.stringify({err:e.message}); }
                })()");
                Log($"[Phase8-verify-pre] {verify}");
            }
            catch { }

            // ── Phase 3: 동시분석그룹 지정 — 맨 마지막 단계로 이동 ──────────
            //   Ctrl+Click 이 Phase2 setValue 한 분석자/결과값을 교란할 수 있어
            //   데이터 입력이 모두 끝난 뒤에 수행 (VBA CommandButton21 포팅)
            try
            {
                var concGroups = MeasurerService.GetAllConcurrentGroups();
                if (concGroups.Count > 0)
                {
                    onStatus("동시분석그룹 지정 중...");
                    var groupMap = concGroups
                        .GroupBy(kv => kv.Value)
                        .Where(g => g.Key > 0)
                        .OrderBy(g => g.Key)
                        .ToDictionary(g => g.Key.ToString(), g => g.Select(x => x.Key).ToList());

                    string groupsJson = JsonSerializer.Serialize(groupMap)
                        .Replace("\\", "\\\\").Replace("'", "\\'");

                    // ① 그룹별 DOM 행 인덱스만 수집 (좌표는 매 클릭 직전 재계산)
                    string groupIdxJs = $@"(function(){{
                        try {{
                            function norm(s){{ return (s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,''); }}
                            var trs = Array.from(document.querySelectorAll(""tr[role='row'].rg-data-row""));
                            var groups = JSON.parse('{groupsJson}');
                            var grpKeys = Object.keys(groups).map(Number).sort(function(a,b){{return a-b;}});
                            var out = {{}}, skipped = [];
                            for (var gi = 0; gi < grpKeys.length; gi++) {{
                                var G = grpKeys[gi];
                                var items = groups[String(G)].map(norm);
                                var itemSet = new Set(items);
                                var idx = [];
                                for (var i = 0; i < trs.length; i++) {{
                                    var td4 = trs[i].querySelector(""td:nth-child(4) div.rg-renderer"");
                                    if (!td4) continue;
                                    var nm = norm((td4.innerText || td4.textContent || '').trim());
                                    if (!itemSet.has(nm)) continue;
                                    idx.push(i);
                                }}
                                if (idx.length < 2) {{ skipped.push('G'+G+':'+idx.length+'건'); continue; }}
                                out[String(G)] = idx;
                            }}
                            return JSON.stringify({{groups: out, skipped: skipped, domRows: trs.length}});
                        }} catch (ex) {{ return JSON.stringify({{error: ex.message}}); }}
                    }})()";

                    // ② 특정 DOM 행 인덱스를 스크롤 후 현재 뷰포트 기준 좌표 반환
                    async Task<(int x, int y, bool ok)> ScrollAndGetCoordAsync(int domIdx)
                    {
                        string js = $@"(function(){{
                            try {{
                                var trs = document.querySelectorAll(""tr[role='row'].rg-data-row"");
                                var tr = trs[{domIdx}];
                                if (!tr) return JSON.stringify({{ok:false}});
                                var td4 = tr.querySelector(""td:nth-child(4) div.rg-renderer"");
                                if (!td4) return JSON.stringify({{ok:false}});
                                td4.scrollIntoView({{block:'center', inline:'nearest'}});
                                var r = td4.getBoundingClientRect();
                                return JSON.stringify({{ok:true, x:Math.round(r.left+r.width/2), y:Math.round(r.top+r.height/2)}});
                            }} catch(e){{ return JSON.stringify({{ok:false, err:e.message}}); }}
                        }})()";
                        using var d = JsonDocument.Parse(await JsEval(js));
                        var r = d.RootElement;
                        if (!r.GetProperty("ok").GetBoolean()) return (0, 0, false);
                        return (r.GetProperty("x").GetInt32(), r.GetProperty("y").GetInt32(), true);
                    }

                    async Task CdpClickAsync(int x, int y, bool ctrl)
                    {
                        int modifiers = ctrl ? 2 : 0; // 1=Alt, 2=Ctrl, 4=Meta, 8=Shift
                        using var _d1 = await CdpSend("Input.dispatchMouseEvent", new
                        {
                            type = "mousePressed", x, y,
                            button = "left", clickCount = 1, modifiers,
                        });
                        await Task.Delay(30);
                        using var _d2 = await CdpSend("Input.dispatchMouseEvent", new
                        {
                            type = "mouseReleased", x, y,
                            button = "left", clickCount = 1, modifiers,
                        });
                    }

                    var applied = new List<string>();
                    var skipped = new List<string>();
                    using (var idxDoc = JsonDocument.Parse(await JsEval(groupIdxJs)))
                    {
                        var root = idxDoc.RootElement;
                        if (root.TryGetProperty("error", out var errEl))
                            throw new Exception(errEl.GetString() ?? "그룹 행 수집 오류");
                        if (root.TryGetProperty("skipped", out var skEl))
                            foreach (var s in skEl.EnumerateArray())
                                skipped.Add(s.GetString() ?? "");

                        if (root.TryGetProperty("groups", out var gsEl))
                        {
                            var groupList = gsEl.EnumerateObject().ToList();
                            int totalGroups = groupList.Count;
                            int doneGroups = 0;
                            foreach (var gprop in groupList)
                            {
                                string grpNum = gprop.Name;
                                var domIdxs = gprop.Value.EnumerateArray().Select(e => e.GetInt32()).ToList();
                                if (domIdxs.Count < 2) continue;

                                onStatus($"동시분석그룹 G{grpNum}: {doneGroups + 1}/{totalGroups}");
                                int hits = 0;
                                for (int k = 0; k < domIdxs.Count; k++)
                                {
                                    var (x, y, ok) = await ScrollAndGetCoordAsync(domIdxs[k]);
                                    if (!ok) continue;
                                    await Task.Delay(60);  // 스크롤 안정 대기
                                    bool useCtrl = k > 0;
                                    await CdpClickAsync(x, y, ctrl: useCtrl);
                                    await Task.Delay(useCtrl ? 100 : 140);
                                    hits++;
                                }

                                if (hits < 2) { skipped.Add($"G{grpNum}:클릭{hits}건"); doneGroups++; continue; }

                                // 버튼 클릭
                                await JsEval("(function(){ var b=document.getElementById('set-conc_group_no-1'); if(b){ b.scrollIntoView({block:'center'}); b.click(); return 'ok'; } return 'no-btn'; })()");
                                await Task.Delay(600);
                                applied.Add($"G{grpNum}:{hits}건");
                                doneGroups++;
                            }
                        }
                    }
                    Log($"[Phase3] applied=[{string.Join(",", applied)}] skipped=[{string.Join(",", skipped)}]");
                    if (applied.Count > 0) groupApplied = string.Join(",", applied);
                }
            }
            catch (Exception ex3)
            {
                Log($"Phase3 예외: {ex3.Message}");
                groupApplied = "예외:" + ex3.Message;
            }

            // ── Phase 9: 그리드 높이 원복 ───────────────────────────────────
            try
            {
                await JsEval(@"(function(){
                    try {
                        var grid = document.getElementById('gridAnalySampAnzeDataAirItemList1');
                        if(!grid) return 'no-grid';
                        var root = grid.querySelector('.rg-root') || grid;
                        root.style.height   = window.__etaGridOrigHeight   || '';
                        root.style.overflow = window.__etaGridOrigOverflow || '';
                        var gv = RealGrid.getGridInstance('gridAnalySampAnzeDataAirItemList1');
                        if(gv && typeof gv.refresh === 'function') gv.refresh();
                        return 'ok';
                    } catch(e){ return e.message; }
                })()");
            }
            catch { }

            string summary = $"전체 {total}건 중 입력완료 {filled}건";
            if (commitFailed.Count > 0) summary += $" / 실패 {commitFailed.Count}건: {string.Join(",", commitFailed.Take(3))}";
            if (!string.IsNullOrEmpty(groupApplied)) summary += $" / 동시분석그룹: {groupApplied}";

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
    //  현장측정장비 스크랩 — 의뢰추가 모달의 select#edit_meas_equip_no 옵션 전체
    //  (의뢰추가 모달이 열려있고 측정장비 드롭다운이 DOM에 존재해야 함)
    // ──────────────────────────────────────────────────────────────────────
    public static async Task<(int saved, string? error)> ScrapeFieldMeasEquipmentsAsync(
        Action<string> onStatus)
    {
        const int port = 9222;
        const string script = @"(function(){
            function norm(s){ return (s||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim(); }
            var out = [];
            var seen = {};
            var prefixes = ['edit_meas_equip_no', 'add_meas_equip_no'];
            var usedId = null;

            for (var p=0;p<prefixes.length;p++){
                var pid = prefixes[p];
                // 1차: native <select>
                var sel = document.getElementById(pid);
                if (sel) {
                    var opts = sel.querySelectorAll('option');
                    for (var i=0;i<opts.length;i++){
                        var v = (opts[i].value || '').trim();
                        var t = norm(opts[i].textContent);
                        if (!v || !t || seen[v]) continue;
                        seen[v] = true;
                        out.push({name: t, code: v});
                    }
                }
                // 2차: Select2 results UL
                var ul = document.getElementById('select2-' + pid + '-results');
                if (ul) {
                    var lis = ul.querySelectorAll('li[id^=""select2-' + pid + '-result-""]');
                    for (var j=0;j<lis.length;j++){
                        var idAttr = lis[j].id || '';
                        var re = new RegExp('^select2-' + pid + '-result-[^-]+-(.+)$');
                        var m = idAttr.match(re);
                        if (!m) continue;
                        var code = m[1];
                        var name = norm(lis[j].textContent);
                        if (!code || !name || seen[code]) continue;
                        seen[code] = true;
                        out.push({name: name, code: code});
                    }
                }
                // 3차: jQuery Select2 data
                if (window.$ && window.$('#' + pid).length && window.$('#' + pid).data('select2')) {
                    try {
                        var data = window.$('#' + pid).select2('data') || [];
                        for (var k=0;k<data.length;k++){
                            var c = (data[k].id || '').toString();
                            var nm = norm(data[k].text || '');
                            if (!c || !nm || seen[c]) continue;
                            seen[c] = true;
                            out.push({name: nm, code: c});
                        }
                    } catch(e){}
                }
                if (out.length > 0) { usedId = pid; break; }
            }
            return JSON.stringify({items: out, usedId: usedId});
        })()";

        try
        {
            onStatus("브라우저 연결 중...");
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
            string cdpJson = await http.GetStringAsync($"http://localhost:{port}/json");
            string? wsUrl = FindWsUrl(cdpJson);
            if (wsUrl == null)
                return (0, "측정인 탭을 찾지 못했습니다. 브라우저에서 측정인.kr을 열어주세요.");

            onStatus("edit_meas_equip_no 읽는 중...");
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

            string json = "";
            using (var doc = JsonDocument.Parse(raw))
            {
                if (doc.RootElement.TryGetProperty("result", out var r1) &&
                    r1.TryGetProperty("result", out var r2) &&
                    r2.TryGetProperty("value", out var v))
                    json = v.GetString() ?? "";
            }

            if (string.IsNullOrWhiteSpace(json))
                return (0, "측정장비 드롭다운을 찾지 못했습니다.\n의뢰추가/수정 모달을 연 상태에서 시도하세요.");

            var items = new List<(string 장비명, string 코드값)>();
            string usedId = "";
            using (var doc = JsonDocument.Parse(json))
            {
                var root = doc.RootElement;
                if (root.TryGetProperty("usedId", out var uid) && uid.ValueKind == JsonValueKind.String)
                    usedId = uid.GetString() ?? "";
                if (root.TryGetProperty("items", out var arr))
                {
                    foreach (var item in arr.EnumerateArray())
                    {
                        string n = item.GetProperty("name").GetString() ?? "";
                        string c = item.GetProperty("code").GetString() ?? "";
                        if (!string.IsNullOrWhiteSpace(n) && !string.IsNullOrWhiteSpace(c))
                            items.Add((n, c));
                    }
                }
            }
            Log($"[ScrapeFieldMeasEquipments] usedId='{usedId}' parsed={items.Count}");
            if (items.Count == 0)
                return (0, "측정장비 option을 찾지 못했습니다.\n의뢰추가 모달에서 측정장비 드롭다운을 한번 열었다 닫은 뒤 시도하세요.");

            onStatus($"{items.Count}개 파싱 완료 — DB 저장 중...");
            int saved = MeasurerService.SaveFieldMeasEquipments(items);
            Log($"[ScrapeFieldMeasEquipments] {saved}/{items.Count}개 저장 (from {usedId})");
            return (saved, null);
        }
        catch (Exception ex)
        {
            Log($"[ScrapeFieldMeasEquipments] 오류: {ex.Message}");
            return (0, ex.Message);
        }
    }

    // ──────────────────────────────────────────────────────────────────────
    //  측정인계약 + 측정인처리시설 스크랩
    //  측정인.kr 의뢰추가 모달이 열려있는 상태에서 호출해야 함
    //  (add_meas_cont_no / cmb_emis_cmpy_plc_no 드롭다운이 DOM에 존재해야 함)
    // ──────────────────────────────────────────────────────────────────────
    public static async Task<(int contracts, int places, int facilities, string? error)> ScrapeContractsAsync(
        Action<string> onStatus)
    {
        const int port = 9222;

        try
        {
            onStatus("브라우저 연결 중...");
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
            string cdpJson = await http.GetStringAsync($"http://localhost:{port}/json");
            string? wsUrl = FindWsUrl(cdpJson);
            if (wsUrl == null)
                return (0, 0, 0, "측정인 탭을 찾지 못했습니다. 브라우저에서 측정인.kr을 열어주세요.");

            using var socket = new ClientWebSocket();
            using var cts    = new CancellationTokenSource(TimeSpan.FromMinutes(5));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            int msgId = 0;

            // 1. 계약 드롭다운 옵션 수집
            onStatus("계약 목록 읽는 중...");
            string contractScript = @"(function(){
                var sel = document.getElementById('add_meas_cont_no');
                if (!sel) return '[]';
                var opts = [];
                for (var i = 0; i < sel.options.length; i++) {
                    var o = sel.options[i];
                    if (!o.value) continue;
                    opts.push({ value: o.value, text: (o.text||'').replace(/\s+/g,' ').trim() });
                }
                return JSON.stringify(opts);
            })()";
            string contractJson = await EvalAsync(socket, contractScript, ++msgId, cts.Token);
            if (string.IsNullOrWhiteSpace(contractJson) || contractJson == "[]")
                return (0, 0, 0, "계약 목록을 찾지 못했습니다.\n측정인.kr에서 의뢰추가 모달을 먼저 열어주세요.");

            var contracts = new List<MeasurerService.MeasurerContract>();
            var contractValues = new List<string>();
            using (var doc = JsonDocument.Parse(contractJson))
            {
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    string v = item.GetProperty("value").GetString() ?? "";
                    string t = item.GetProperty("text").GetString() ?? "";
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    // 측정인 포맷: "계약번호 / 업체명 / 대표사업장 / 기간"
                    var parts = t.Split(new[] { " / " }, StringSplitOptions.None);
                    string co = parts.Length > 1 ? parts[1].Trim() : "";
                    string pl = parts.Length > 2 ? parts[2].Trim() : "";
                    string pd = parts.Length > 3 ? parts[3].Trim() : "";
                    contracts.Add(new MeasurerService.MeasurerContract(v.Trim(), co, pl, pd, t));
                    contractValues.Add(v.Trim());
                }
            }
            int savedContracts = MeasurerService.SaveMeasurerContracts(contracts);
            Log($"[ScrapeContracts] 계약 {savedContracts}/{contracts.Count}개 저장");

            // 2. 각 계약별 처리시설 목록 수집
            int savedPlaces = 0;
            int savedFacilities = 0;
            for (int i = 0; i < contractValues.Count; i++)
            {
                string cNo = contractValues[i];
                onStatus($"처리시설 읽는 중 {i + 1}/{contractValues.Count} — {cNo}");

                // 1단계: 빈값으로 리셋 (같은/다른 값 재선택도 AJAX 강제 발화)
                string resetScript = @"(function(){
                    if (!window.$) return 'NO_JQ';
                    var $el = window.$('#add_meas_cont_no');
                    if (!$el.length) return 'NO_SEL';
                    $el.val('').trigger('change');
                    var $pl = window.$('#cmb_emis_cmpy_plc_no');
                    if ($pl.length) { $pl.empty().trigger('change'); }
                    return 'OK';
                })()";
                await EvalAsync(socket, resetScript, ++msgId, cts.Token);
                await Task.Delay(40, cts.Token);

                // 2단계: 타겟 계약번호 선택 (jQuery.val().trigger('change') 정공법)
                string cvEsc = cNo.Replace("'", "\\'").Replace("\"", "\\\"");
                string selectScript = $@"(function(){{
                    if (!window.$) return 'NO_JQ';
                    var $el = window.$('#add_meas_cont_no');
                    if (!$el.length) return 'NO_SEL';
                    var opt = $el.find('option[value=""{cvEsc}""]');
                    if (!opt.length) return 'NOT_FOUND';
                    $el.val('{cvEsc}').trigger('change');
                    $el.trigger({{ type:'select2:select', params:{{ data:{{ id:'{cvEsc}', text: opt.text() }} }} }});
                    return 'OK';
                }})()";
                await EvalAsync(socket, selectScript, ++msgId, cts.Token);

                // 3단계: 측정대상 사업장 드롭다운 로딩 대기 (최대 6초, 150ms)
                // — valued>=1 이면 수락 (한 계약에 사업장 1곳인 경우도 포함)
                string placesJson = "[]";
                for (int w = 0; w < 6000; w += 150)
                {
                    await Task.Delay(150, cts.Token);
                    string readScript = @"(function(){
                        var sel = document.getElementById('cmb_emis_cmpy_plc_no');
                        if (!sel) return JSON.stringify({ valued:0, opts:[] });
                        var opts = [];
                        for (var i = 0; i < sel.options.length; i++) {
                            var o = sel.options[i];
                            if (!o.value) continue;
                            opts.push({ value: o.value, text: (o.text||'').replace(/\s+/g,' ').trim() });
                        }
                        return JSON.stringify({ valued:opts.length, opts:opts });
                    })()";
                    string r = await EvalAsync(socket, readScript, ++msgId, cts.Token);
                    if (string.IsNullOrWhiteSpace(r)) continue;
                    using var d = JsonDocument.Parse(r);
                    if (!d.RootElement.TryGetProperty("valued", out var vEl)) continue;
                    if (vEl.GetInt32() < 1) continue;
                    if (d.RootElement.TryGetProperty("opts", out var optsEl))
                        placesJson = optsEl.GetRawText();
                    break;
                }

                if (string.IsNullOrWhiteSpace(placesJson) || placesJson == "[]")
                {
                    Log($"[ScrapeContracts] 계약 {cNo} — 처리시설 없음");
                    continue;
                }

                var places = new List<MeasurerService.MeasurerPlace>();
                using (var doc = JsonDocument.Parse(placesJson))
                {
                    foreach (var item in doc.RootElement.EnumerateArray())
                    {
                        string v = item.GetProperty("value").GetString() ?? "";
                        string t = item.GetProperty("text").GetString() ?? "";
                        if (string.IsNullOrWhiteSpace(t)) continue;
                        places.Add(new MeasurerService.MeasurerPlace(cNo, t, v));
                    }
                }
                savedPlaces += MeasurerService.SaveMeasurerPlaces(cNo, places);
                Log($"[ScrapeContracts] 계약 {cNo} — 처리시설 {places.Count}건 저장");

                // 4단계: 각 처리시설별로 측정시설(edit_emis_fac_no) 스크랩
                for (int p = 0; p < places.Count; p++)
                {
                    var place = places[p];
                    if (string.IsNullOrWhiteSpace(place.처리시설코드)) continue;
                    onStatus($"측정시설 읽는 중 {i + 1}/{contractValues.Count} 계약 · {p + 1}/{places.Count} 처리시설 — {place.처리시설명}");

                    // 처리시설 선택 (빈값 리셋 → 타겟)
                    string pcEsc = place.처리시설코드.Replace("'", "\\'").Replace("\"", "\\\"");
                    string plcResetScript = @"(function(){
                        if (!window.$) return 'NO_JQ';
                        var $el = window.$('#cmb_emis_cmpy_plc_no');
                        if (!$el.length) return 'NO_SEL';
                        $el.val('').trigger('change');
                        var $fac = window.$('#edit_emis_fac_no');
                        if ($fac.length) { $fac.empty().trigger('change'); }
                        return 'OK';
                    })()";
                    await EvalAsync(socket, plcResetScript, ++msgId, cts.Token);
                    await Task.Delay(40, cts.Token);

                    string plcSelScript = $@"(function(){{
                        if (!window.$) return 'NO_JQ';
                        var $el = window.$('#cmb_emis_cmpy_plc_no');
                        if (!$el.length) return 'NO_SEL';
                        var opt = $el.find('option[value=""{pcEsc}""]');
                        if (!opt.length) return 'NOT_FOUND';
                        $el.val('{pcEsc}').trigger('change');
                        $el.trigger({{ type:'select2:select', params:{{ data:{{ id:'{pcEsc}', text: opt.text() }} }} }});
                        return 'OK';
                    }})()";
                    await EvalAsync(socket, plcSelScript, ++msgId, cts.Token);

                    // edit_emis_fac_no 로딩 대기 (최대 5초, 150ms)
                    string facJson = "[]";
                    for (int w = 0; w < 5000; w += 150)
                    {
                        await Task.Delay(150, cts.Token);
                        string readFac = @"(function(){
                            var sel = document.getElementById('edit_emis_fac_no');
                            if (!sel) return JSON.stringify({ valued:0, opts:[] });
                            var opts = [];
                            for (var i = 0; i < sel.options.length; i++) {
                                var o = sel.options[i];
                                if (!o.value) continue;
                                opts.push({ value: o.value, text: (o.text||'').replace(/\s+/g,' ').trim() });
                            }
                            return JSON.stringify({ valued:opts.length, opts:opts });
                        })()";
                        string r2 = await EvalAsync(socket, readFac, ++msgId, cts.Token);
                        if (string.IsNullOrWhiteSpace(r2)) continue;
                        using var d = JsonDocument.Parse(r2);
                        if (!d.RootElement.TryGetProperty("valued", out var vEl)) continue;
                        if (vEl.GetInt32() < 1) continue;
                        if (d.RootElement.TryGetProperty("opts", out var optsEl))
                            facJson = optsEl.GetRawText();
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(facJson) || facJson == "[]")
                    {
                        Log($"[ScrapeContracts] 계약 {cNo} · 처리시설 {place.처리시설명} — 측정시설 없음");
                        continue;
                    }

                    var facilities = new List<MeasurerService.MeasurerFacility>();
                    using (var doc = JsonDocument.Parse(facJson))
                    {
                        foreach (var item in doc.RootElement.EnumerateArray())
                        {
                            string v = item.GetProperty("value").GetString() ?? "";
                            string t = item.GetProperty("text").GetString() ?? "";
                            if (string.IsNullOrWhiteSpace(t)) continue;
                            facilities.Add(new MeasurerService.MeasurerFacility(
                                cNo, place.처리시설코드, place.처리시설명, t, v));
                        }
                    }
                    savedFacilities += MeasurerService.SaveMeasurerFacilities(cNo, place.처리시설코드, facilities);
                    Log($"[ScrapeContracts] 계약 {cNo} · 처리시설 {place.처리시설명} — 측정시설 {facilities.Count}건 저장");
                }
            }

            onStatus($"계약 {savedContracts}건, 처리시설 {savedPlaces}건, 측정시설 {savedFacilities}건 저장 완료");
            return (savedContracts, savedPlaces, savedFacilities, null);
        }
        catch (Exception ex)
        {
            Log($"[ScrapeContracts] 오류: {ex.Message}");
            return (0, 0, 0, ex.Message);
        }
    }

    // 요청·응답을 id로 매칭해 실제 값만 반환
    private static async Task<string> EvalAsync(ClientWebSocket socket, string script, int id, CancellationToken ct)
    {
        var cmd = JsonSerializer.Serialize(new
        {
            id,
            method = "Runtime.evaluate",
            @params = new { expression = script, returnByValue = true }
        });
        await socket.SendAsync(Encoding.UTF8.GetBytes(cmd), WebSocketMessageType.Text, true, ct);

        var buf = new byte[262144];
        var sb  = new StringBuilder();
        while (true)
        {
            sb.Clear();
            WebSocketReceiveResult recv;
            do
            {
                recv = await socket.ReceiveAsync(buf, ct);
                sb.Append(Encoding.UTF8.GetString(buf, 0, recv.Count));
            } while (!recv.EndOfMessage);

            string raw = sb.ToString();
            using var doc = JsonDocument.Parse(raw);
            // id 일치하는 결과만 수용 (CDP는 이벤트도 함께 오기 때문)
            if (!doc.RootElement.TryGetProperty("id", out var idEl)) continue;
            if (idEl.GetInt32() != id) continue;
            if (!doc.RootElement.TryGetProperty("result", out var r1)) return "";
            if (!r1.TryGetProperty("result", out var r2)) return "";
            if (r2.TryGetProperty("value", out var v))
                return v.ValueKind == JsonValueKind.String ? (v.GetString() ?? "") : v.ToString();
            return "";
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
