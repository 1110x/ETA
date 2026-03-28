using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.WebSockets;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ETA.Views;

public partial class MeasurerLoginWindow : Window
{
    private const string LoginUrl  = "https://측정인.kr/login.go";
    private const string TargetUrl = "https://측정인.kr/ms/field_water.do";
    private const int    CdpPort   = 9222;

    private static readonly string LogPath =
        Path.Combine(AppContext.BaseDirectory, "측정인.log");

    private ClientWebSocket? _ws;
    private CancellationTokenSource _pollCts = new();

    // ── 브라우저 옵션 ─────────────────────────────────────────────────────────
    private record BrowserOption(string Label, string? ExePath);

    private List<BrowserOption> _browsers = new();

    public MeasurerLoginWindow()
    {
        InitializeComponent();
        Loaded  += OnLoaded;
        Closing += OnWindowClosing;
    }

    // ── 사용 가능한 브라우저 목록 구성 ───────────────────────────────────────
    private List<BrowserOption> BuildBrowserList()
    {
        var list = new List<BrowserOption>();

        if (OperatingSystem.IsMacOS())
        {
            // Mac Edge
            const string macEdge = "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge";
            if (File.Exists(macEdge))
                list.Add(new BrowserOption("Microsoft Edge (Mac)", macEdge));

            // Mac Chrome
            const string macChrome = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
            if (File.Exists(macChrome))
                list.Add(new BrowserOption("Google Chrome (Mac)", macChrome));

            // Safari (CDP 미지원 — 수동 로그인)
            list.Add(new BrowserOption("Safari (수동 로그인)", null));
        }
        else
        {
            // Windows Edge
            foreach (var p in new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Microsoft", "Edge", "Application", "msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "Microsoft", "Edge", "Application", "msedge.exe"),
            })
            {
                if (File.Exists(p)) { list.Add(new BrowserOption("Microsoft Edge", p)); break; }
            }

            // Windows Chrome
            foreach (var p in new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                    "Google", "Chrome", "Application", "chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "Google", "Chrome", "Application", "chrome.exe"),
            })
            {
                if (File.Exists(p)) { list.Add(new BrowserOption("Google Chrome", p)); break; }
            }
        }

        return list;
    }

    private void OnWindowClosing(object? sender, WindowClosingEventArgs e)
    {
        _pollCts.Cancel();
        _ws?.Dispose();
    }

    private void OnLoaded(object? sender, RoutedEventArgs e)
    {
        _browsers = BuildBrowserList();
        cmbBrowser.ItemsSource  = _browsers.Select(b => b.Label).ToList();
        cmbBrowser.SelectedIndex = 0;

        if (_browsers.Count == 0)
            SetStatus("사용 가능한 브라우저를 찾을 수 없습니다.", "#ee4444");

        LoadCredentials();
    }

    // ── 자격증명 로드 ─────────────────────────────────────────────────────────
    private void LoadCredentials()
    {
        try
        {
            string empId = MainPage.CurrentEmployeeId;
            if (string.IsNullOrEmpty(empId)) return;

            var (id, pw) = AgentService.GetMeasurerCredentials(empId);
            if (!string.IsNullOrEmpty(id))
            {
                txbUserId.Text   = id;
                txbPassword.Text = pw;
                SetStatus("저장된 자격증명을 불러왔습니다. 바로 로그인 실행 가능합니다.", "#88cc88");
            }
        }
        catch { }
    }

    // ── 자격증명 저장 ─────────────────────────────────────────────────────────
    private static void SaveCredentials(string userId, string password)
    {
        try
        {
            string empId = MainPage.CurrentEmployeeId;
            if (!string.IsNullOrEmpty(empId))
                AgentService.SaveMeasurerCredentials(empId, userId, password);
        }
        catch { }
    }

    // ── 로그인 실행 ───────────────────────────────────────────────────────────
    private async void BtnLogin_Click(object? sender, RoutedEventArgs e)
    {
        int idx = cmbBrowser.SelectedIndex;
        if (idx < 0 || idx >= _browsers.Count)
        {
            SetStatus("브라우저를 선택하세요.", "#ee4444");
            return;
        }

        var browser = _browsers[idx];

        // Safari: CDP 미지원 — 수동으로 브라우저 열기
        if (browser.ExePath == null)
        {
            SetStatus("Safari로 측정인.kr을 엽니다. 직접 로그인 후 사용하세요.", "#aaaacc");
            try { Process.Start(new ProcessStartInfo("open", $"-a Safari \"{LoginUrl}\"") { UseShellExecute = false }); }
            catch (Exception ex) { SetStatus($"Safari 열기 실패: {ex.Message}", "#ee4444"); }
            return;
        }

        string userId   = txbUserId.Text?.Trim()   ?? "";
        string password = txbPassword.Text?.Trim() ?? "";

        if (string.IsNullOrEmpty(userId))   { SetStatus("아이디를 입력하세요.",   "#ee4444"); txbUserId.Focus();   return; }
        if (string.IsNullOrEmpty(password)) { SetStatus("비밀번호를 입력하세요.", "#ee4444"); txbPassword.Focus(); return; }

        btnLogin.IsEnabled = false;
        string? errorMsg = null;

        await Task.Run(async () =>
        {
            try
            {
                await RunCdpLoginAsync(userId, password, browser.ExePath);
                SaveCredentials(userId, password);
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
            }
        });

        if (errorMsg != null)
        {
            SetStatus($"오류: {errorMsg}", "#ee4444");
            btnLogin.IsEnabled = true;
        }
        else
        {
            // 창을 닫지 않고 대기 — 모달 버튼 클릭 감지 루프 시작
            SetStatus("로그인 완료 — 측정인.kr 모달에서 'ETA 계약 DB 업데이트' 또는 'ETA 분석DB 업데이트' 버튼을 누르세요.", "#88cc88");
            _pollCts = new CancellationTokenSource();
            _ = Task.Run(() => PollForSyncRequestAsync(_pollCts.Token));
        }
    }

    // ── CDP 기반 자동 로그인 ──────────────────────────────────────────────────
    private async Task RunCdpLoginAsync(string userId, string password, string browserExePath)
    {
        // ── 1. 브라우저 실행 ─────────────────────────────────────────────────
        Post("브라우저 실행 중...", "#aaaaaa");

        string sessionDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ETA", "BrowserSession");
        Directory.CreateDirectory(sessionDir);

        Process.Start(new ProcessStartInfo
        {
            FileName        = browserExePath,
            Arguments       = $"--remote-debugging-port={CdpPort} " +
                              $"--user-data-dir=\"{sessionDir}\" " +
                              $"\"{LoginUrl}\"",
            UseShellExecute = false,
        });

        // ── 3. Edge CDP 응답 대기 ────────────────────────────────────────────
        Post("Edge 응답 대기 중...", "#aaaaaa");
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
        string json = await PollCdpJsonAsync(http);

        // ── 4. WebSocket 디버거 URL 파싱 ────────────────────────────────────
        string? wsUrl = null;
        using (var doc = JsonDocument.Parse(json))
        {
            foreach (var item in doc.RootElement.EnumerateArray())
            {
                if (item.TryGetProperty("type", out var t) && t.GetString() == "page")
                {
                    wsUrl = item.GetProperty("webSocketDebuggerUrl").GetString();
                    break;
                }
            }
        }
        if (wsUrl == null) throw new Exception("CDP 페이지를 찾을 수 없습니다.");

        // ── 5. WebSocket 연결 (필드에 보관 — 창이 닫힐 때까지 유지) ─────────
        _ws = new ClientWebSocket();
        await _ws.ConnectAsync(new Uri(wsUrl), CancellationToken.None);

        // ── 6. 로그인 폼 대기 ────────────────────────────────────────────────
        Post("로그인 폼 로딩 대기 중...", "#aaaaaa");
        await PollElementExistsAsync("user_email", timeoutMs: 15000);
        await DismissPopupsAsync();

        // ── 7. 폼 입력 및 로그인 클릭 ───────────────────────────────────────
        Post("아이디 · 비밀번호 입력 중...", "#aaaaaa");

        string safeId = userId  .Replace("\\", "\\\\").Replace("'", "\\'");
        string safePw = password.Replace("\\", "\\\\").Replace("'", "\\'");

        string fillScript = $@"(function(){{
            var eId  = document.getElementById('user_email');
            var ePw  = document.getElementById('login_pwd_confirm');
            var eBtn = document.getElementById('login');
            if (!eId) return 'NO_FORM';
            eId.value  = '{safeId}';
            ePw.value  = '{safePw}';
            eBtn.click();
            return 'OK';
        }})()";

        string result = await Evaluate(fillScript);
        if (result.Contains("NO_FORM"))
            throw new Exception("로그인 폼을 찾을 수 없습니다.\nURL이 올바른지 확인하세요: " + LoginUrl);

        // ── 8. 로그인 완료 확인 ──────────────────────────────────────────────
        Post("로그인 완료 확인 중...", "#aaaaaa");
        await PollUrlNotContainsAsync("login", timeoutMs: 15000);

        // ── 9. 수질분야 페이지 이동 ──────────────────────────────────────────
        Post("수질분야 페이지 이동 중...", "#aaaaaa");
        await Evaluate($"window.location.href = '{TargetUrl}';");

        // ── 10. 페이지 로딩 대기 ────────────────────────────────────────────
        Post("페이지 로딩 확인 중...", "#aaaaaa");
        await PollPageReadyAsync(urlMustContain: "field_water", timeoutMs: 15000);

        // ── 11. 모달에 'ETA DB 업데이트' 버튼 자동 주입 ─────────────────────
        await InjectSyncButtonObserverAsync();
    }

    // ── 모달 버튼 주입 (MutationObserver) ────────────────────────────────────
    // 모달이 열릴 때마다 'ETA 계약 DB 업데이트' / 'ETA 분석DB 업데이트' 버튼을 자동으로 삽입한다.
    private async Task InjectSyncButtonObserverAsync()
    {
        const string script = @"(function() {
            if (window.__etaObserver) return 'ALREADY';
            window.__etaSyncRequested    = false;
            window.__etaAnalysisRequested = false;

            function tryInject() {
                var contSel = document.getElementById('add_meas_cont_no');
                if (!contSel || !contSel.offsetParent) return;   // 보이지 않으면 skip
                if (document.getElementById('__eta_sync_btn__')) return; // 이미 있음

                // ── 계약 DB 업데이트 버튼 (파란색) ──────────────────────────
                var btnSync = document.createElement('button');
                btnSync.id   = '__eta_sync_btn__';
                btnSync.type = 'button';
                btnSync.textContent = 'ETA 계약 DB 업데이트';
                btnSync.style.cssText =
                    'background:#1a6fc4;color:#fff;border:none;padding:5px 14px;' +
                    'border-radius:4px;cursor:pointer;font-size:13px;' +
                    'margin-top:8px;display:inline-block;margin-right:6px;';
                btnSync.onclick = function() {
                    if (window.__etaSyncRequested) return;
                    window.__etaSyncRequested = true;
                    btnSync.textContent = '수집 중...';
                    btnSync.disabled = true;
                };

                // ── 분석DB 업데이트 버튼 (녹색) ─────────────────────────────
                var btnAnalysis = document.createElement('button');
                btnAnalysis.id   = '__eta_analysis_btn__';
                btnAnalysis.type = 'button';
                btnAnalysis.textContent = 'ETA 분석DB 업데이트';
                btnAnalysis.style.cssText =
                    'background:#1a6f2a;color:#fff;border:none;padding:5px 14px;' +
                    'border-radius:4px;cursor:pointer;font-size:13px;' +
                    'margin-top:8px;display:inline-block;';
                btnAnalysis.onclick = function() {
                    if (window.__etaAnalysisRequested) return;
                    window.__etaAnalysisRequested = true;
                    btnAnalysis.textContent = '수집 중...';
                    btnAnalysis.disabled = true;
                };

                var parent = contSel.closest('.modal-body, .modal-content, .form-group')
                          || contSel.parentElement;
                if (parent) {
                    var wrap = document.createElement('div');
                    wrap.style.marginTop = '8px';
                    wrap.appendChild(btnSync);
                    wrap.appendChild(btnAnalysis);
                    parent.appendChild(wrap);
                }
            }

            window.__etaObserver = new MutationObserver(tryInject);
            window.__etaObserver.observe(document.body, {
                childList: true, subtree: true,
                attributes: true, attributeFilter: ['style', 'class']
            });
            tryInject();
            return 'OBSERVING';
        })()";

        await Evaluate(script);
        Post("준비 완료 — 측정인.kr 모달을 열고 버튼을 누르세요.", "#88cc88");
    }

    // ── 버튼 클릭 감지 폴링 루프 ─────────────────────────────────────────────
    private async Task PollForSyncRequestAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested && _ws?.State == WebSocketState.Open)
        {
            try
            {
                await Task.Delay(600, ct);

                // ── 계약 DB 업데이트 플래그 ──────────────────────────────────
                string flag = ExtractCdpStringValue(await Evaluate("String(window.__etaSyncRequested)"));
                if (flag == "true")
                {
                    await Evaluate("window.__etaSyncRequested = false;");
                    Post("ETA DB 동기화 중...", "#aaaaaa");
                    await ScrapeSamplingPointsAsync();
                    await Evaluate(@"(function(){
                        var b = document.getElementById('__eta_sync_btn__');
                        if(b){ b.textContent='ETA 계약 DB 업데이트 (완료)'; b.disabled=false; }
                    })()");
                    continue;
                }

                // ── 분석DB 업데이트 플래그 ───────────────────────────────────
                string flag2 = ExtractCdpStringValue(await Evaluate("String(window.__etaAnalysisRequested)"));
                if (flag2 == "true")
                {
                    await Evaluate("window.__etaAnalysisRequested = false;");
                    Post("분석항목 코드 수집 중...", "#aaaaaa");
                    await ScrapeAnalysisItemsAsync();
                    await Evaluate(@"(function(){
                        var b = document.getElementById('__eta_analysis_btn__');
                        if(b){ b.textContent='ETA 분석DB 업데이트 (완료)'; b.disabled=false; }
                    })()");
                }
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex) { Debug.WriteLine($"[Poll] {ex.Message}"); }
        }
    }

    // ── 채취지점 스크래핑: 계약 → 현장(cmb_emis_cmpy_plc_no) → 채취지점(add_emis_fac_no) ──
    private async Task ScrapeSamplingPointsAsync()
    {
        Log("=== ScrapeSamplingPointsAsync 시작 ===");

        // ── 1단계: 계약 목록 ─────────────────────────────────────────────────
        const string getContractsScript = @"(function() {
            var sel = document.getElementById('add_meas_cont_no');
            if (!sel) return '[]';
            var opts = [];
            for (var i = 0; i < sel.options.length; i++) {
                var o = sel.options[i];
                if (!o.value) continue;
                opts.push({ value: o.value, text: o.text.trim() });
            }
            return JSON.stringify(opts);
        })()";

        string contractsJson = ExtractCdpStringValue(await Evaluate(getContractsScript));
        Log($"contracts = {contractsJson}");
        if (string.IsNullOrEmpty(contractsJson) || contractsJson == "[]")
        {
            Post("계약 목록을 찾을 수 없습니다", "#ffaa44");
            return;
        }

        // 옵션 형식: "계약번호 / 업체명 / 사업장명 / 계약기간"
        var contracts = new List<(string Value, string 계약번호, string 계약기간)>();
        using (var doc = JsonDocument.Parse(contractsJson))
            foreach (var item in doc.RootElement.EnumerateArray())
            {
                string val  = item.GetProperty("value").GetString() ?? "";
                string text = item.GetProperty("text").GetString()  ?? "";
                if (string.IsNullOrEmpty(val)) continue;
                var parts = text.Split(" / ");
                string 계약번호 = parts.Length > 0 ? parts[0].Trim() : "";
                string 계약기간 = parts.Length > 3 ? parts[3].Trim() : "";
                contracts.Add((val, 계약번호, 계약기간));
            }

        MeasurerService.EnsureMeasurerTable();
        MeasurerService.ClearAll();   // 재스크래핑 전 초기화
        int totalSaved = 0;
        Dispatcher.UIThread.Post(() => SetProgress(0, contracts.Count));

        // ── 2단계: 계약별 반복 ───────────────────────────────────────────────
        for (int ci = 0; ci < contracts.Count; ci++)
        {
            var (contVal, 계약번호, 계약기간) = contracts[ci];
            Post($"[{ci + 1}/{contracts.Count}] 계약 선택 중...", "#aaaaaa");
            Dispatcher.UIThread.Post(() => SetProgress(ci + 1, contracts.Count));
            Log($"--- 계약 [{ci+1}/{contracts.Count}] {계약번호} ({계약기간}) ---");

            // 계약 선택 (select2)
            string selectContScript = $@"(function(){{
                var $el = window.$ && window.$('#add_meas_cont_no');
                if ($el && $el.length) {{
                    var optText = '';
                    $el.find('option[value=""{contVal}""]').each(function(){{ optText = this.text; return false; }});
                    $el.val('{contVal}');
                    $el.trigger({{ type: 'select2:select', params: {{ data: {{ id: '{contVal}', text: optText }} }} }});
                    $el.trigger('change');
                    return 'select2';
                }}
                var sel = document.getElementById('add_meas_cont_no');
                if (!sel) return 'NO_SEL';
                sel.value = '{contVal}';
                sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                return 'native';
            }})()";
            Log($"계약 trigger: {ExtractCdpStringValue(await Evaluate(selectContScript))}");

            // ── 3단계: 현장 목록 대기 ────────────────────────────────────────
            const string getPlcScript = @"(function() {
                var sel = document.getElementById('cmb_emis_cmpy_plc_no');
                if (!sel || sel.options.length === 0) return '[]';
                var opts = [];
                for (var i = 0; i < sel.options.length; i++) {
                    var o = sel.options[i];
                    if (!o.value) continue;
                    var 현장명 = o.text.split(' / ')[0].trim();
                    opts.push({ value: o.value, 현장명: 현장명 });
                }
                return JSON.stringify(opts);
            })()";

            string plcJson = "[]";
            for (int w = 0; w < 5000; w += 400)
            {
                await Task.Delay(400);
                plcJson = ExtractCdpStringValue(await Evaluate(getPlcScript));
                if (!string.IsNullOrEmpty(plcJson) && plcJson != "[]") break;
            }
            Log($"현장 목록: {plcJson}");

            if (string.IsNullOrEmpty(plcJson) || plcJson == "[]") { Log("현장 없음 — 스킵"); continue; }

            var places = new List<(string Value, string 현장명)>();
            using (var plcDoc = JsonDocument.Parse(plcJson))
                foreach (var item in plcDoc.RootElement.EnumerateArray())
                {
                    string val  = item.GetProperty("value").GetString()  ?? "";
                    string name = item.GetProperty("현장명").GetString() ?? "";
                    if (!string.IsNullOrEmpty(val) && !string.IsNullOrEmpty(name))
                        places.Add((val, name));
                }

            // ── 4단계: 현장별 채취지점 수집 ──────────────────────────────────
            foreach (var (plcVal, 현장명) in places)
            {
                Post($"  └ {현장명}", "#aaaaaa");
                Log($"  현장: {현장명} (value={plcVal})");

                // 현장 선택 (일반 select — native change)
                string selectPlcScript = $@"(function(){{
                    var sel = document.getElementById('cmb_emis_cmpy_plc_no');
                    if (!sel) return 'NO_SEL';
                    sel.value = '{plcVal}';
                    sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                    return 'native';
                }})()";
                Log($"  현장 trigger: {ExtractCdpStringValue(await Evaluate(selectPlcScript))}");

                // add_emis_fac_no 로딩 대기 (최대 5초)
                const string getFacScript = @"(function() {
                    var el = document.getElementById('add_emis_fac_no');
                    if (!el) return 'NO_EL';
                    var pts = [];
                    for (var j = 0; j < el.options.length; j++) {
                        var t = el.options[j].text.trim();
                        if (t && t !== '선택' && el.options[j].value) pts.push(t);
                    }
                    return JSON.stringify(pts);
                })()";

                string pntJson = "[]";
                for (int w = 0; w < 5000; w += 400)
                {
                    await Task.Delay(400);
                    string raw = ExtractCdpStringValue(await Evaluate(getFacScript));
                    if (raw == "NO_EL") { Log("  add_emis_fac_no 없음"); break; }
                    pntJson = raw;
                    if (!string.IsNullOrEmpty(pntJson) && pntJson != "[]") break;
                }
                Log($"  pntJson = {pntJson}");

                var points = new List<string>();
                if (!string.IsNullOrEmpty(pntJson) && pntJson != "[]")
                    using (var pntDoc = JsonDocument.Parse(pntJson))
                        foreach (var p in pntDoc.RootElement.EnumerateArray())
                        {
                            string? pt = p.GetString();
                            if (!string.IsNullOrWhiteSpace(pt)) points.Add(pt);
                        }

                if (points.Count > 0)
                {
                    MeasurerService.SaveSamplingPoints(현장명, points, 계약번호, 계약기간);
                    totalSaved += points.Count;
                    Log($"  저장: {points.Count}개 → {string.Join(", ", points)}");
                }
                else Log("  채취지점 0개");
            }
        }

        Log($"=== 완료: {contracts.Count}개 계약, 채취지점 {totalSaved}개 저장 ===");
        Dispatcher.UIThread.Post(() => SetProgress(contracts.Count, contracts.Count));
        Post($"완료 — 채취지점 {totalSaved}개 저장", "#88cc88");

        // 인력 고유번호도 함께 동기화
        await ScrapeEmployeeIdsAsync();
    }

    // ── 인력 고유번호 스크래핑 ────────────────────────────────────────────────
    /// <summary>
    /// 측정인.kr add_emp_id 드롭다운에서 인력 목록을 수집하여
    /// ETA Agent 테이블의 측정인고유번호를 업데이트한다.
    /// (채취지점 스크래핑 후 자동 호출)
    /// </summary>
    private async Task ScrapeEmployeeIdsAsync()
    {
        Log("=== 인력 고유번호 스크래핑 ===");
        Post("인력 고유번호 수집 중...", "#aaaaaa");

        // add_emp_id 옵션 전체 JSON 추출
        // 옵션 텍스트 형식: "이름 / 소속 / ..." (VBA와 동일)
        const string script = @"(function() {
            var sel = document.getElementById('add_emp_id');
            if (!sel) return '[]';
            var opts = [];
            for (var i = 0; i < sel.options.length; i++) {
                var o = sel.options[i];
                if (!o.value) continue;
                var parts = o.text.trim().split(' / ');
                opts.push({ value: o.value, name: parts[0].trim() });
            }
            return JSON.stringify(opts);
        })()";

        string json = ExtractCdpStringValue(await Evaluate(script));
        Log($"인력 raw: {(json.Length > 200 ? json[..200] + "..." : json)}");

        if (string.IsNullOrEmpty(json) || json == "[]")
        {
            Log("인력 고유번호: add_emp_id 요소 없음 또는 옵션 없음");
            return;
        }

        var pairs = new List<(string Name, string Id)>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            foreach (var item in doc.RootElement.EnumerateArray())
            {
                string name = item.GetProperty("name").GetString() ?? "";
                string id   = item.GetProperty("value").GetString() ?? "";
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(id))
                    pairs.Add((name, id));
            }
        }
        catch (Exception ex) { Log($"인력 JSON 파싱 오류: {ex.Message}"); return; }

        Log($"인력 옵션 {pairs.Count}개 수집: {string.Join(", ", pairs.Take(5).Select(p => p.Name))}");

        int updated = ETA.Services.AgentService.UpdateMeasurerEmployeeIds(pairs);
        Log($"인력 고유번호 업데이트: {updated}명 / {pairs.Count}개 매칭");
        Post($"인력 {updated}명 고유번호 업데이트 완료", "#88cc88");
    }

    // ── 분석항목 코드 스크래핑 (add_meas_item) ──────────────────────────────
    /// <summary>
    /// 측정인.kr add_meas_item 드롭다운에서 분석항목 목록(항목명, 코드값, select2id)을
    /// 수집하여 ETA 측정인_분석항목 테이블에 저장한다.
    /// </summary>
    private async Task ScrapeAnalysisItemsAsync()
    {
        Log("=== 분석항목 코드 스크래핑 시작 ===");
        Post("분석항목 코드 수집 중...", "#aaaaaa");

        const string script = @"(function() {
            var sel = document.getElementById('add_meas_item');
            if (!sel) return '[]';
            var opts = [];
            for (var i = 0; i < sel.options.length; i++) {
                var o = sel.options[i];
                if (!o.value) continue;
                var parts = o.text.split('|');
                var field = '', category = '', name = '';
                if (parts.length >= 3) {
                    field    = parts[0].trim();
                    category = parts[1].trim();
                    name     = parts[2].trim();
                } else if (parts.length === 2) {
                    field = parts[0].trim();
                    name  = parts[1].trim();
                } else {
                    name = o.text.trim();
                }
                opts.push({
                    field:    field,
                    category: category,
                    name:     name,
                    code:     o.value,
                    select2id: o.getAttribute('data-select2-id') || ''
                });
            }
            return JSON.stringify(opts);
        })()";

        string json = ExtractCdpStringValue(await Evaluate(script));
        Log($"분석항목 raw: {(json.Length > 200 ? json[..200] + "..." : json)}");

        if (string.IsNullOrEmpty(json) || json == "[]")
        {
            Log("분석항목: add_meas_item 요소 없음 또는 옵션 없음 — 모달이 열린 상태인지 확인하세요.");
            Post("분석항목을 찾을 수 없습니다 (모달을 열고 다시 시도하세요)", "#ffaa44");
            return;
        }

        var items = new List<(string 분야, string 항목구분, string 항목명, string 코드값, string Select2Id)>();
        try
        {
            using var doc = JsonDocument.Parse(json);
            foreach (var item in doc.RootElement.EnumerateArray())
            {
                string field    = item.GetProperty("field").GetString()    ?? "";
                string category = item.GetProperty("category").GetString() ?? "";
                string name     = item.GetProperty("name").GetString()     ?? "";
                string code     = item.GetProperty("code").GetString()     ?? "";
                string s2id     = item.GetProperty("select2id").GetString() ?? "";
                if (!string.IsNullOrEmpty(code))
                    items.Add((field, category, name, code, s2id));
            }
        }
        catch (Exception ex) { Log($"분석항목 JSON 파싱 오류: {ex.Message}"); return; }

        Log($"분석항목 {items.Count}개 수집: {string.Join(", ", items.Take(5).Select(i => $"{i.분야}|{i.항목구분}|{i.항목명}"))}");

        int saved = MeasurerService.SaveAnalysisItems(items);
        Log($"분석항목 저장 완료: {saved}개 반영 / {items.Count}개");
        Post($"분석항목 {items.Count}개 저장 완료", "#88cc88");
    }

    // ── CDP 인스턴스 헬퍼 (_ws 사용) ─────────────────────────────────────────
    private Task<string> Evaluate(string script) => EvaluateAsync(_ws!, script);

    private async Task PollElementExistsAsync(string id, int timeoutMs = 10000)
    {
        for (int elapsed = 0; elapsed < timeoutMs; elapsed += 300)
        {
            await Task.Delay(300);
            string r = await Evaluate($"!!document.getElementById('{id}')");
            if (r.Contains("true")) return;
        }
        throw new Exception($"요소 '#{id}'를 찾을 수 없습니다 ({timeoutMs / 1000}초 초과).");
    }

    private async Task PollUrlNotContainsAsync(string avoidStr, int timeoutMs = 10000)
    {
        for (int elapsed = 0, iter = 0; elapsed < timeoutMs; elapsed += 300, iter++)
        {
            await Task.Delay(300);
            if (iter % 5 == 0) await DismissPopupsAsync();
            string url = await Evaluate("window.location.href");
            if (!url.Contains(avoidStr)) return;
        }
        throw new Exception("로그인 실패: 아이디 또는 비밀번호를 확인하세요.");
    }

    private async Task PollPageReadyAsync(string urlMustContain, int timeoutMs = 10000)
    {
        for (int elapsed = 0, iter = 0; elapsed < timeoutMs; elapsed += 300, iter++)
        {
            await Task.Delay(300);
            if (iter % 5 == 0) await DismissPopupsAsync();
            string state = await Evaluate("document.readyState");
            if (!state.Contains("complete")) continue;
            string url = await Evaluate("window.location.href");
            if (url.Contains(urlMustContain)) return;
        }
    }

    private async Task DismissPopupsAsync()
    {
        const string script = @"(function() {
            window.alert   = function() {};
            window.confirm = function() { return true; };
            window.prompt  = function(msg, def) { return def || ''; };

            var selectors = [
                '[class*=""close""]', '[class*=""Close""]',
                '[id*=""close""]',    '[id*=""Close""]',
                '.layer-close',  '.popup-close',  '.modal-close',
                '.btn-close',    '.btn_close',    '.pop_close',
                '.layer_close',  '.ui-dialog-titlebar-close',
                '[title=""닫기""]', '[alt=""닫기""]'
            ].join(',');
            try {
                document.querySelectorAll(selectors).forEach(function(el) {
                    if (el.offsetParent !== null) el.click();
                });
            } catch(e) {}

            var closeTexts = ['닫기', '확인', '×', '✕', 'X', 'Close', 'OK'];
            var btns = document.querySelectorAll('button, a, input[type=""button""], input[type=""submit""]');
            btns.forEach(function(el) {
                if (el.offsetParent === null) return;
                var t = el.textContent.trim();
                if (closeTexts.indexOf(t) !== -1) el.click();
            });
            return 'ok';
        })()";

        try { await Evaluate(script); } catch { }
    }

    // ── CDP 응답에서 실제 문자열 값 추출 ────────────────────────────────────
    private static string ExtractCdpStringValue(string cdpResponse)
    {
        try
        {
            using var doc = JsonDocument.Parse(cdpResponse);
            return doc.RootElement
                .GetProperty("result")
                .GetProperty("result")
                .GetProperty("value")
                .GetString() ?? "";
        }
        catch { return ""; }
    }

    // ── CDP Runtime.evaluate (static) ────────────────────────────────────────
    private static int _cdpId;

    private static async Task<string> EvaluateAsync(ClientWebSocket ws, string script)
    {
        int id  = Interlocked.Increment(ref _cdpId);
        var cmd = JsonSerializer.Serialize(new
        {
            id,
            method  = "Runtime.evaluate",
            @params = new { expression = script, returnByValue = true }
        });

        var sendBuf = Encoding.UTF8.GetBytes(cmd);
        await ws.SendAsync(sendBuf, WebSocketMessageType.Text, true, CancellationToken.None);

        var recvBuf = new byte[16384];
        var recv    = await ws.ReceiveAsync(recvBuf, CancellationToken.None);
        return Encoding.UTF8.GetString(recvBuf, 0, recv.Count);
    }

    // ── CDP JSON 폴링 ────────────────────────────────────────────────────────
    private static async Task<string> PollCdpJsonAsync(HttpClient http, int timeoutMs = 12000)
    {
        Exception? last = null;
        for (int elapsed = 0; elapsed < timeoutMs; elapsed += 400)
        {
            await Task.Delay(400);
            try   { return await http.GetStringAsync($"http://localhost:{CdpPort}/json"); }
            catch (Exception ex) { last = ex; }
        }
        throw new Exception(
            $"Edge CDP 응답 없음 (포트 {CdpPort})\n원인: {last?.Message}\n" +
            "기존에 열린 Edge 창을 모두 닫고 다시 시도하세요.");
    }

    private void BtnClose_Click(object? sender, RoutedEventArgs e) => Close();

    private void Post(string text, string hexColor) =>
        Dispatcher.UIThread.Post(() => SetStatus(text, hexColor));

    // ── 파일 로그 ─────────────────────────────────────────────────────────────
    private static void Log(string msg)
    {
        try
        {
            File.AppendAllText(LogPath,
                $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
        }
        catch { }
    }

    private void SetStatus(string text, string hexColor)
    {
        txbStatus.Text       = text;
        txbStatus.Foreground = new SolidColorBrush(Color.Parse(hexColor));
    }

    private void SetProgress(int current, int total)
    {
        pgbScrape.IsVisible = total > 0;
        pgbScrape.Value     = total > 0 ? (double)current / total * 100 : 0;
    }
}
