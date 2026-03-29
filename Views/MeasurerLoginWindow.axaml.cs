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

    // Personal SET.log 저장 경로
    private static readonly string PersonalSetPath =
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Documents", "ETA", "Data", "Personal SET.log");

    private static readonly string LogPath =
        Path.Combine(AppContext.BaseDirectory, "측정인.log");

    private ClientWebSocket? _ws;
    private CancellationTokenSource _pollCts = new();

    /// <summary>로그인 성공 시 발생 — 외부에서 창 닫힌 후 데이터 주입 재시도용</summary>
    public bool LoginSucceeded { get; private set; }

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
        // 로그인 성공 시에는 폴링/WebSocket 유지 (백그라운드에서 계속 동작)
        if (LoginSucceeded) return;
        _pollCts.Cancel();
        _ws?.Dispose();
    }

    private void OnLoaded(object? sender, RoutedEventArgs e)
    {
        LogSessionStart();
        _browsers = BuildBrowserList();
        cmbBrowser.ItemsSource  = _browsers.Select(b => b.Label).ToList();
        cmbBrowser.SelectedIndex = 0;

        if (_browsers.Count == 0)
            SetStatus("사용 가능한 브라우저를 찾을 수 없습니다.", "#ee4444");

        LoadCredentials();
    }

    // ── 자격증명 로드 (Personal SET.log 우선, AgentService fallback) ─────────────────
    private void LoadCredentials()
    {
        try
        {
            string empId = MainPage.CurrentEmployeeId;

            // 1순위: Personal SET.log 파일
            if (File.Exists(PersonalSetPath))
            {
                var entries = LoadPersonalSet();
                var entry = entries.FirstOrDefault(e =>
                    string.IsNullOrEmpty(empId) || e.EmpId == empId);
                if (entry != null && !string.IsNullOrEmpty(entry.UserId))
                {
                    txbUserId.Text   = entry.UserId;
                    txbPassword.Text = entry.Password;
                    SetStatus("저장된 자격증명을 불러왔습니다. 바로 로그인 실행 가능합니다.", "#88cc88");
                    return;
                }
            }

            // 2순위: AgentService (DB)
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

    // ── Personal SET.log 모델 ──────────────────────────────────────────────────
    private record PersonalSetEntry(string EmpId, string UserId, string Password);

    private static List<PersonalSetEntry> LoadPersonalSet()
    {
        try
        {
            string json = File.ReadAllText(PersonalSetPath);
            using var doc = JsonDocument.Parse(json);
            var list = new List<PersonalSetEntry>();
            foreach (var el in doc.RootElement.EnumerateArray())
            {
                string empId = el.TryGetProperty("empId",    out var e1) ? e1.GetString() ?? "" : "";
                string uid   = el.TryGetProperty("userId",   out var e2) ? e2.GetString() ?? "" : "";
                string pw    = el.TryGetProperty("password", out var e3) ? e3.GetString() ?? "" : "";
                list.Add(new PersonalSetEntry(empId, uid, pw));
            }
            return list;
        }
        catch { return new List<PersonalSetEntry>(); }
    }

    // ── 자격증명 저장 (Personal SET.log + AgentService 동시) ─────────────────────
    private static void SaveCredentials(string userId, string password)
    {
        try
        {
            string empId = MainPage.CurrentEmployeeId ?? "";

            // Personal SET.log 저장
            var entries = File.Exists(PersonalSetPath)
                ? LoadPersonalSet()
                : new List<PersonalSetEntry>();

            bool found = false;
            for (int i = 0; i < entries.Count; i++)
            {
                if (entries[i].EmpId == empId)
                {
                    entries[i] = new PersonalSetEntry(empId, userId, password);
                    found = true;
                    break;
                }
            }
            if (!found)
                entries.Add(new PersonalSetEntry(empId, userId, password));

            Directory.CreateDirectory(Path.GetDirectoryName(PersonalSetPath)!);
            var jsonArr = entries.Select(e => new
            {
                empId    = e.EmpId,
                userId   = e.UserId,
                password = e.Password
            });
            File.WriteAllText(PersonalSetPath,
                JsonSerializer.Serialize(jsonArr,
                    new JsonSerializerOptions { WriteIndented = true }));

            // AgentService(DB) 백업
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
            // 로그인 성공 — 창을 닫지 않고 DB 업데이트 대기 상태로 전환
            LoginSucceeded = true;
            Log($"=== 로그인 완료, 폴링 시작 (UserId={txbUserId.Text?.Trim()}) ===");

            // 로그인 폼 비활성화, DB 업데이트 패널 표시
            btnLogin.IsEnabled   = false;
            cmbBrowser.IsEnabled = false;
            txbUserId.IsEnabled  = false;
            txbPassword.IsEnabled = false;
            if (pnlDbUpdate != null) pnlDbUpdate.IsVisible = true;
            SetStatus("✅ 로그인 완료. 브라우저에서 DB 업데이트 버튼을 클릭하세요.", "#88ffaa");

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

    // 단일 버튼 주입 (MutationObserver)
    // 계획 작성 모달이 열릴 때마다 'ETA DB 전체 업데이트' 단일 버튼을 자동으로 삽입한다.
    private async Task InjectSyncButtonObserverAsync()
    {
        const string script = @"(function() {
            if (window.__etaObserver) return 'ALREADY';
            window.__etaSyncRequested        = false;
            window.__etaAnalysisRequested    = false;
            window.__etaFullUpdateRequested  = false;
            window.__etaFieldPlanRequested   = false;
            window.__etaAutoOpenFieldPlanRequested = false;

            function tryInjectModal() {
                var contSel = document.getElementById('add_meas_cont_no');
                if (!contSel || !contSel.offsetParent) return;
                if (document.getElementById('__eta_full_update_btn__')) return;

                var btn = document.createElement('button');
                btn.id   = '__eta_full_update_btn__';
                btn.type = 'button';
                btn.textContent = 'ETA DB 전체 업데이트';
                btn.style.cssText =
                    'background:#7a3aaa;color:#fff;border:none;padding:6px 16px;' +
                    'border-radius:4px;cursor:pointer;font-size:13px;font-weight:bold;' +
                    'margin-top:8px;display:inline-block;';
                btn.onclick = function() {
                    if (window.__etaFullUpdateRequested) return;
                    window.__etaFullUpdateRequested = true;
                    btn.textContent = '수집 중...';
                    btn.disabled = true;
                };

                var parent = contSel.closest('.modal-body, .modal-content, .form-group')
                          || contSel.parentElement;
                if (parent) {
                    var wrap = document.createElement('div');
                    wrap.id = '__eta_btn_wrap__';
                    wrap.style.marginTop = '8px';
                    wrap.appendChild(btn);
                    parent.appendChild(wrap);
                }
            }

            function hookFieldPlanBtn() {
                var planBtn = document.getElementById('addFieldPlan');
                if (!planBtn || !planBtn.offsetParent) return;
                if (planBtn.dataset.etaHooked === '1') return;

                planBtn.dataset.etaHooked = '1';
                planBtn.addEventListener('click', function() {
                    if (!window.__etaRequestData || !window.__etaRequestData.length) {
                        alert('의룰데이터 없음. 측정인 전송을 먼저 누르세요.');
                        return;
                    }
                    if (window.__etaFieldPlanRequested) return;
                    window.__etaFieldPlanRequested = true;
                }, true);
            }

            function tryAutoOpenFieldPlan() {
                if (!window.__etaAutoOpenFieldPlanRequested) return;
                var planBtn = document.getElementById('addFieldPlan');
                if (!planBtn || !planBtn.offsetParent) return;
                window.__etaAutoOpenFieldPlanRequested = false;
                planBtn.click();
            }

            function tryInject() {
                tryInjectModal();
                hookFieldPlanBtn();
                tryAutoOpenFieldPlan();
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
                // 전체 DB 업데이트 플래그 (통합 버튼)
                string flagFull = ExtractCdpStringValue(await Evaluate("String(window.__etaFullUpdateRequested)"));
                if (flagFull == "true")
                {
                    await Evaluate("window.__etaFullUpdateRequested = false;");
                    Post("전체 DB 업데이트 시작...", "#bb88ff");
                    try { await ScrapeFullDbAsync(); }
                    catch (Exception ex2)
                    {
                        Log($"[전체DB] {ex2}");
                        Post($"전체 DB 업데이트 실패: {ex2.Message}", "#ff8844");
                    }
                    await Evaluate(@"(function(){
                        var b = document.getElementById('__eta_full_update_btn__');
                        if(b){ b.textContent='ETA DB 전체 업데이트'; b.disabled=false; }
                    })()");
                    continue;
                }

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
                    continue;
                }

                // ── 의뢰계획 일괄작성 플래그 ─────────────────────────────────
                string flag3 = ExtractCdpStringValue(await Evaluate("String(window.__etaFieldPlanRequested)"));
                if (flag3 == "true")
                {
                    await Evaluate("window.__etaFieldPlanRequested = false;");
                    Post("의뢰계획 일괄작성 중...", "#aa88ff");

                    try
                    {
                        await FillFieldPlanAsync();
                    }
                    catch (Exception ex2)
                    {
                        Log($"[의뢰계획] {ex2}");
                        Post($"의뢰계획 작성 실패: {ex2.Message}", "#ff8844");
                    }

                    continue;
                }
            }
            catch (OperationCanceledException) { break; }
            catch (Exception ex) { Debug.WriteLine($"[Poll] {ex.Message}"); }
        }
    }

    private async Task FillFieldPlanAsync()
    {
        int requestCount = await GetRequestDataCountAsync();
        if (requestCount <= 0)
            throw new Exception("의뢰데이터가 없습니다. 먼저 측정인 전송을 실행하세요.");

        await Task.Delay(900);

        int rowCount = await GetFieldPlanRowCountAsync();
        for (int attempt = 0; rowCount <= 0 && attempt < 12; attempt++)
        {
            await Task.Delay(300);
            rowCount = await GetFieldPlanRowCountAsync();
        }

        if (rowCount <= 0)
            throw new Exception("계획 작성 모달의 입력 Row를 찾지 못했습니다.");

        string fillResultJson = ExtractCdpStringValue(await Evaluate(@"(function() {
            var data = Array.isArray(window.__etaRequestData) ? window.__etaRequestData : [];
            function isVisible(el) {
                return !!el && el.offsetParent !== null && getComputedStyle(el).visibility !== 'hidden';
            }
            function normalize(value) {
                return (value || '').toString().trim();
            }
            function toDateText(value) {
                var text = normalize(value);
                if (!text) return '';
                text = text.replace(/\./g, '-').replace(/\//g, '-');
                return text.length >= 10 ? text.slice(0, 10) : text;
            }
            function trigger(el) {
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
            }
            function setValue(el, value) {
                if (!el || !isVisible(el) || el.disabled || el.readOnly) return false;
                var text = normalize(value);
                if (!text) return false;
                el.focus();
                el.value = text;
                trigger(el);
                return true;
            }
            function setSelect(el, preferred) {
                if (!el || el.disabled) return false;
                var options = Array.from(el.options || []);
                if (!options.length) return false;

                var wants = (preferred || []).map(normalize).filter(Boolean);
                var match = -1;

                for (var wi = 0; wi < wants.length && match < 0; wi++) {
                    var want = wants[wi].toLowerCase();
                    match = options.findIndex(function(opt) {
                        var text = normalize(opt.text).toLowerCase();
                        var value = normalize(opt.value).toLowerCase();
                        return text.indexOf(want) >= 0 || value === want;
                    });
                }

                if (match < 0) {
                    match = options.findIndex(function(opt) {
                        var text = normalize(opt.text);
                        return normalize(opt.value) && text !== '선택' && text.toLowerCase() !== 'select';
                    });
                }

                if (match < 0) return false;
                el.selectedIndex = match;
                trigger(el);
                return true;
            }
            function setSelect2Value(el, option) {
                if (!el || !option) return false;
                if (window.$) {
                    var $el = window.$(el);
                    $el.val(option.value);
                    $el.trigger({
                        type: 'select2:select',
                        params: { data: { id: option.value, text: option.text } }
                    });
                    $el.trigger('change');
                    return true;
                }
                el.value = option.value;
                trigger(el);
                return true;
            }
            function findContractOption(options, item) {
                var wantNo = normalize(item.contractNo).toLowerCase();
                var wantCompany = normalize(item.contractCompany || item.company).toLowerCase();
                var wantPoint = normalize(item.contractPoint || item.workSite).toLowerCase();
                var wantLabel = normalize(item.contractLabel).toLowerCase();

                var scored = options.map(function(opt, idx) {
                    var text = normalize(opt.text).toLowerCase();
                    var value = normalize(opt.value).toLowerCase();
                    var score = 0;
                    if (wantNo && (value === wantNo || text.indexOf(wantNo) >= 0)) score += 500;
                    if (wantCompany && text.indexOf(wantCompany) >= 0) score += 220;
                    if (wantPoint && text.indexOf(wantPoint) >= 0) score += 180;
                    if (wantLabel && text === wantLabel) score += 320;
                    if (score === 0 && wantCompany && wantPoint && text.indexOf(wantCompany) >= 0 && text.indexOf(wantPoint) >= 0) score += 260;
                    return { idx: idx, option: opt, score: score };
                }).sort(function(a, b) { return b.score - a.score; });

                return scored.length && scored[0].score > 0 ? scored[0].option : null;
            }
            function selectContractForRow(row, item, rowIndex, totalRows) {
                // 행 안에 계약 select가 있으면 사용
                var scopedSelects = Array.from(row.querySelectorAll('select')).filter(function(el) {
                    var key = (normalize(el.name) + ' ' + normalize(el.id)).toLowerCase();
                    return key.indexOf('cont') >= 0 || key.indexOf('contract') >= 0 || key.indexOf('meas_cont_no') >= 0;
                });

                // 전역 add_meas_cont_no (모달 전체에서 하나) — 첫 행에서만 설정
                if (!scopedSelects.length && rowIndex === 0) {
                    var globalSelect = document.getElementById('add_meas_cont_no');
                    if (globalSelect && globalSelect.tagName === 'SELECT') scopedSelects.push(globalSelect);
                }

                for (var si = 0; si < scopedSelects.length; si++) {
                    var select = scopedSelects[si];
                    var options = Array.from(select.options || []).filter(function(opt) {
                        return normalize(opt.value) && normalize(opt.text) && normalize(opt.text) !== '선택';
                    });
                    if (!options.length) continue;

                    var chosen = findContractOption(options, item);
                    if (!chosen) continue;
                    if (setSelect2Value(select, chosen)) return true;
                }

                return false;
            }
            function findRows() {
                var tables = Array.from(document.querySelectorAll('table')).filter(isVisible);
                var bestRows = [];
                tables.forEach(function(table) {
                    var rows = Array.from(table.querySelectorAll('tbody tr, tr')).filter(function(row) {
                        return row.querySelector('input, textarea, select');
                    });
                    if (rows.length > bestRows.length) bestRows = rows;
                });
                return bestRows;
            }
            function pickByKeyword(elements, keywords) {
                return elements.find(function(el) {
                    var key = (normalize(el.name) + ' ' + normalize(el.id) + ' ' + normalize(el.placeholder)).toLowerCase();
                    return keywords.some(function(word) { return key.indexOf(word) >= 0; });
                });
            }

            var rows = findRows();
            if (!rows.length)
                return JSON.stringify({ ok: false, error: '계획 입력 Row를 찾을 수 없습니다.' });

            var filled = 0;
            rows.forEach(function(row, index) {
                var item = data[index];
                if (!item) return;

                var textInputs = Array.from(row.querySelectorAll('input:not([type]), input[type=text], textarea')).filter(function(el) {
                    return isVisible(el) && !el.disabled && !el.readOnly;
                });
                var dateInputs = Array.from(row.querySelectorAll('input[type=date], input[id*=date], input[name*=date], input[id*=dt], input[name*=dt]')).filter(function(el) {
                    return isVisible(el) && !el.disabled && !el.readOnly;
                });
                var selects = Array.from(row.querySelectorAll('select')).filter(isVisible);

                var analyteText = Array.isArray(item.analytes) ? item.analytes.filter(Boolean).join(', ') : '';
                var planText = normalize(item.planText) || analyteText || normalize(item.sample);
                var summaryText = [normalize(item.accNo), normalize(item.sample), analyteText].filter(Boolean).join(' / ');
                var company = normalize(item.company);
                var manager = normalize(item.manager);
                var workSite = normalize(item.workSite);
                var note = normalize(item.note);
                var dateText = toDateText(item.sampleDate || item.date);

                selectContractForRow(row, item, index, rows.length);

                setValue(pickByKeyword(textInputs, ['acc', 'receipt', 'request', '접수']), item.accNo);
                setValue(pickByKeyword(textInputs, ['sample', '시료', 'title', 'name']), item.sample);
                setValue(pickByKeyword(textInputs, ['analyte', 'item', 'target', '항목']), analyteText || planText);
                setValue(pickByKeyword(textInputs, ['memo', 'remark', '비고', 'content', '내용', 'plan', '계획']), note || summaryText || planText);
                setValue(pickByKeyword(textInputs, ['company', 'corp', '업체']), company);
                setValue(pickByKeyword(textInputs, ['manager', '담당']), manager);
                setValue(pickByKeyword(textInputs, ['site', 'place', '사업장', '현장']), workSite);

                var candidates = [
                    normalize(item.accNo),
                    normalize(item.sample),
                    analyteText || planText,
                    note || summaryText || planText,
                    company,
                    manager,
                    workSite
                ].filter(Boolean);

                textInputs.forEach(function(input) {
                    if (normalize(input.value)) return;
                    var nextValue = candidates.shift();
                    if (nextValue) setValue(input, nextValue);
                });

                if (dateText) {
                    var dateInput = dateInputs.find(function(input) { return !normalize(input.value); }) || dateInputs[0];
                    if (dateInput) setValue(dateInput, dateText);
                }

                selects.forEach(function(select) {
                    var current = normalize(select.value);
                    if (current) return;
                    setSelect(select, [item.abbr, item.sample, analyteText, item.company]);
                });

                filled++;
            });

            return JSON.stringify({ ok: true, rows: rows.length, filled: filled, requested: data.length });
        })()"));

        using (var doc = JsonDocument.Parse(fillResultJson))
        {
            bool ok = doc.RootElement.TryGetProperty("ok", out var okProp) && okProp.GetBoolean();
            if (!ok)
            {
                string error = doc.RootElement.TryGetProperty("error", out var errProp)
                    ? errProp.GetString() ?? "알 수 없는 오류"
                    : "알 수 없는 오류";
                throw new Exception(error);
            }
        }

        await PollElementExistsAsync("insertFieldPlanBtn", 8000);
        await Task.Delay(200);
        await Evaluate(@"(function(){
            var b = document.getElementById('insertFieldPlanBtn');
            if (!b || b.offsetParent === null) return 'NO_BTN';
            b.click();
            return 'OK';
        })()");

        Post($"의뢰계획 {requestCount}건 작성 후 저장 완료", "#88cc88");
        Log($"의뢰계획 자동작성 완료: {requestCount}건");
    }

    private async Task<int> GetRequestDataCountAsync()
    {
        string raw = ExtractCdpStringValue(await Evaluate("String((window.__etaRequestData || []).length)"));
        return int.TryParse(raw, out var count) ? count : 0;
    }

    private async Task<int> GetFieldPlanRowCountAsync()
    {
        string raw = ExtractCdpStringValue(await Evaluate(@"(function() {
            function isVisible(el) {
                return !!el && el.offsetParent !== null && getComputedStyle(el).visibility !== 'hidden';
            }
            var best = 0;
            Array.from(document.querySelectorAll('table')).filter(isVisible).forEach(function(table) {
                var rows = Array.from(table.querySelectorAll('tbody tr, tr')).filter(function(row) {
                    return row.querySelector('input, textarea, select');
                }).length;
                if (rows > best) best = rows;
            });
            return String(best);
        })()"));

        return int.TryParse(raw, out var count) ? count : 0;
    }

    // 전체 DB 수집: 더미 계획서 생성 → 저장 → 행 더블클릭
    private async Task ScrapeFullDbAsync()
    {
        Post("전체 DB 업데이트 — 더미 계획서 작성 중...", "#bb88ff");

        // 1. 계획 작성 모달 열기
        Log("[전체DB] 계획 작성 모달 열기");
        await Evaluate("(function(){ var b=document.getElementById('addFieldPlan'); if(b) b.click(); })()");
        await PollElementExistsAsync("add_meas_cont_no", 10000);
        await Task.Delay(600);

        // 2. 첫번째 계약 선택
        Post("더미: 계약 선택 중...", "#bb88ff");
        string selContract = ExtractCdpStringValue(await Evaluate(@"(function(){
            var sel = document.getElementById('add_meas_cont_no');
            if (!sel || sel.options.length < 2) return 'NO_OPT';
            var opt = sel.options[1];
            if (window.$) {
                window.$('#add_meas_cont_no').val(opt.value)
                    .trigger({type:'select2:select', params:{data:{id:opt.value, text:opt.text}}})
                    .trigger('change');
                return 'select2:' + opt.value;
            }
            sel.selectedIndex = 1;
            sel.dispatchEvent(new Event('change', {bubbles:true}));
            return 'native:' + opt.value;
        })()"));
        Log($"[전체DB] 계약 선택: {selContract}");

        // 3. 현장 선택
        await Task.Delay(1500);
        string selPlc = ExtractCdpStringValue(await Evaluate(@"(function(){
            var sel = document.getElementById('cmb_emis_cmpy_plc_no');
            if (!sel || sel.options.length < 2) return 'NO_OPT';
            sel.selectedIndex = 1;
            sel.dispatchEvent(new Event('change', {bubbles:true}));
            return 'native:' + sel.options[1].value;
        })()"));
        Log($"[전체DB] 현장 선택: {selPlc}");

        // 4. 채취지점 선택 (add_emis_fac_no)
        await Task.Delay(1500);
        string selFac = ExtractCdpStringValue(await Evaluate(@"(function(){
            var sel = document.getElementById('add_emis_fac_no');
            if (!sel || sel.options.length < 2) return 'NO_OPT';
            sel.selectedIndex = 1;
            sel.dispatchEvent(new Event('change', {bubbles:true}));
            return 'native:' + sel.options[1].value;
        })()"));
        Log($"[전체DB] 채취지점 선택: {selFac}");
        await Task.Delay(500);

        // 5. 측정용도 = 자가측정용
        await Evaluate(@"(function(){
            var selects = Array.from(document.querySelectorAll('select'));
            for (var s of selects) {
                var sid = (s.id + ' ' + s.name).toLowerCase();
                if (sid.indexOf('prp') < 0 && sid.indexOf('purpose') < 0 && sid.indexOf('meas_div') < 0) continue;
                for (var i = 0; i < s.options.length; i++) {
                    if (s.options[i].text.indexOf('자가측정') >= 0) {
                        s.selectedIndex = i;
                        s.dispatchEvent(new Event('change', {bubbles:true}));
                        break;
                    }
                }
            }
        })");

        // 6. 분석항목 첫번째 선택
        await Task.Delay(400);
        string selItem = ExtractCdpStringValue(await Evaluate(@"(function(){
            var sel = document.getElementById('add_meas_item');
            if (!sel || sel.options.length < 2) return 'NO_OPT';
            var opt = sel.options[1];
            if (window.$) {
                window.$('#add_meas_item').val(opt.value)
                    .trigger({type:'select2:select', params:{data:{id:opt.value, text:opt.text}}})
                    .trigger('change');
                return 'select2:' + opt.value;
            }
            sel.selectedIndex = 1;
            sel.dispatchEvent(new Event('change', {bubbles:true}));
            return 'native:' + opt.value;
        })()"));
        Log($"[전체DB] 분석항목 선택: {selItem}");
        await Task.Delay(400);

        // 7. meas_no_yn 체크박스 클릭 -> add_meas_no 활성화
        await Evaluate(@"(function(){
            var cb = document.getElementById('meas_no_yn');
            if (cb && !cb.checked) {
                cb.click();
                cb.dispatchEvent(new Event('change', {bubbles:true}));
            }
        })");
        await Task.Delay(400);

        // 8. add_meas_no = "ETA DB 업데이트" 입력 후 Enter
        await Evaluate(@"(function(){
            var inp = document.getElementById('add_meas_no');
            if (!inp) return;
            inp.removeAttribute('disabled');
            inp.focus();
            inp.value = 'ETA DB 업데이트';
            inp.dispatchEvent(new Event('input',  {bubbles:true}));
            inp.dispatchEvent(new Event('change', {bubbles:true}));
            inp.dispatchEvent(new KeyboardEvent('keydown',  {key:'Enter', code:'Enter', keyCode:13, bubbles:true}));
            inp.dispatchEvent(new KeyboardEvent('keypress', {key:'Enter', code:'Enter', keyCode:13, bubbles:true}));
            inp.dispatchEvent(new KeyboardEvent('keyup',    {key:'Enter', code:'Enter', keyCode:13, bubbles:true}));
        })");
        Log("[전체DB] add_meas_no 입력 Enter 발생");
        await Task.Delay(400);

        // 9. 저장 버튼 클릭
        Post("더미 계획서 저장 중...", "#bb88ff");
        await PollElementExistsAsync("insertFieldPlanBtn", 5000);
        await Evaluate(@"(function(){
            var btn = document.getElementById('insertFieldPlanBtn');
            if (btn) btn.click();
        })");
        Log("[전체DB] insertFieldPlanBtn 클릭");

        // 10. 모달 닫힌 대기 — 모달이 완전히 닫혀야만 다음 진행
        Log("[전체DB] 모달 닫힘 대기 시작");
        bool modalClosed = false;
        for (int w = 0; w < 15000; w += 300)
        {
            await Task.Delay(300);
            string vis = ExtractCdpStringValue(await Evaluate(
                "String(!!(document.getElementById('add_meas_cont_no')?.offsetParent))"));
            if (vis == "false") 
            { 
                modalClosed = true;
                break;
            }
        }
        
        if (!modalClosed)
        {
            Log("[전체DB] 경고: 모달이 닫히지 않음. 강제 닫기 시도");
            await Evaluate(@"(function(){
                var modals = document.querySelectorAll('.modal, [role=""dialog""], .overlay');
                modals.forEach(function(m) { if (m) m.style.display = 'none'; });
            })");
        }
        
        // 모달 닫힌 후 추가 대기 (테이블 렌더링)
        await Task.Delay(2500);
        Log("[전체DB] 모달 닫힘 완료 - 테이블 렌더링 대기");
        
        // 테이블 스크롤 최상단으로 이동 (첫 행 보이게)
        await Evaluate(@"(function(){
            var tbl = document.querySelector('table, [role=""grid""]');
            if (!tbl) return;
            tbl.scrollTop = 0;
            var firstRow = tbl.querySelector('tr[role=""row""], tr.rg-data-row');
            if (firstRow) firstRow.scrollIntoView({behavior:'instant', block:'start'});
        })");
        await Task.Delay(800);

        // 11. 테이블에서 "ETA DB 업데이트" 행 더블클릭 (최대 6초 반복 탐색)
        // 체크박스 셀이 아닌 데이터 셀을 더블클릭해서 체크박스 토글 방지
        Post("테이블에서 ETA DB 업데이트 행 탐색 중...", "#bb88ff");
        string dblResult = "NOT_FOUND";
        for (int retry = 0; retry < 12; retry++) // 최대 6초 (500ms * 12회)
        {
            dblResult = ExtractCdpStringValue(await Evaluate(@"(function(){
                var rows = Array.from(document.querySelectorAll('tr[role=""row""], tr.rg-data-row'));
                var debugCount = rows.length;
                
                console.log('[ETA] Table rows:', debugCount);
                if (debugCount === 0) return 'EMPTY_TABLE';
                
                // ETA DB 업데이트 행 찾기
                for (var i = 0; i < rows.length; i++) {
                    var row = rows[i];
                    var t = (row.innerText || row.textContent || '').trim();
                    if (t && t.indexOf('ETA DB') >= 0) {
                        console.log('[ETA] Found at row', i);
                        
                        // 행 내에서 체크박스가 아닌 데이터 셀 찾기
                        var cells = Array.from(row.querySelectorAll('td, th, .rg-renderer, [role=""gridcell""]'));
                        var dataCell = null;
                        
                        for (var cell of cells) {
                            // input[type=checkbox]는 스킵
                            if (cell.querySelector('input[type=""checkbox""]')) continue;
                            // 텍스트가 있는 첫 번째 셀 사용
                            var ct = (cell.innerText || cell.textContent || '').trim();
                            if (ct && ct.length > 0) {
                                dataCell = cell;
                                break;
                            }
                        }
                        
                        // 데이터 셀이 없으면 행 전체 사용
                        if (!dataCell) dataCell = row;
                        
                        // 데이터 셀에 더블클릭
                        dataCell.focus();
                        dataCell.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, view:window}));
                        dataCell.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, view:window}));
                        dataCell.dispatchEvent(new MouseEvent('click', {bubbles:true, view:window}));
                        setTimeout(function() {
                            dataCell.dispatchEvent(new MouseEvent('dblclick', {bubbles:true, view:window}));
                        }, 150);
                        return 'OK_CELL:' + t.slice(0, 80);
                    }
                }
                
                return 'NOT_FOUND';\n            })()"""));
            
            if (!dblResult.StartsWith("NOT_FOUND")) 
            {
                Log($"[전체DB] 행 탐색 성공 (시도 {retry + 1}/12): {dblResult}");
                await Task.Delay(300); // 더블클릭 이벤트 처리 대기
                break;
            }
            
            if (retry < 11)
            {
                Log($"[전체DB] 행 탐색 시도 {retry + 1}/12");
                await Task.Delay(500);
            }
        }
        
        Log($"[전체DB] 행 더블클릭 최종: {dblResult}");

        if (dblResult.StartsWith("NOT_FOUND"))
            Post($"'ETA DB 업데이트' 행을 찾지 못했습니다. ({dblResult})", "#ff6644");
        else
            Post($"더블클릭 완료: {dblResult}", "#88cc88");
    }

    // ── 채취지점 스크래핑: 계약 → 현장(cmb_emis_cmpy_plc_no) → 채취지점(add_emis_fac_no) ──
    private async Task ScrapeSamplingPointsAsync()
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        Log($"=== 계약 DB 스크래핑 시작 [{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ===");

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
            Log("[계약DB] 계약 목록 없음 — 모달이 열린 상태인지 확인하세요.");
            Post("계약 목록을 찾을 수 없습니다 (모달을 먼저 여세요)", "#ffaa44");
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
        // ClearAll() 호출 제거 — UPSERT로 신규/변경분만 반영, 기존 데이터 유지
        int totalSaved = 0;
        Dispatcher.UIThread.Post(() => SetProgress(0, contracts.Count));
        await WebProgress_InjectAsync($"ETA 계약 DB 업데이트 (쳙 {contracts.Count}건)", contracts.Count, "#1a6fc4");

        // ── 2단계: 계약별 반복 ───────────────────────────────────────────
        for (int ci = 0; ci < contracts.Count; ci++)
        {
            var (contVal, 계약번호, 계약기간) = contracts[ci];
            Post($"[{ci + 1}/{contracts.Count}] 계약 선택 중...", "#aaaaaa");
            Dispatcher.UIThread.Post(() => SetProgress(ci + 1, contracts.Count));
            await WebProgress_UpdateAsync(ci + 1, contracts.Count, $"[{ci + 1}/{contracts.Count}] {{계약번호}}");
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

        sw.Stop();
        Log($"=== 계약 DB 스크래핑 완료: {contracts.Count}개 계약, 채취지점 {totalSaved}개 저장 (소요 {sw.Elapsed.TotalSeconds:F1}초) ===");
        Dispatcher.UIThread.Post(() => SetProgress(contracts.Count, contracts.Count));
        await WebProgress_UpdateAsync(contracts.Count, contracts.Count, $"완료 — 채취지점 {totalSaved}개 저장");
        Post($"완료 — 채취지점 {totalSaved}개 저장", "#88cc88");

        // 인력 고유번호도 함께 동기화
        await ScrapeEmployeeIdsAsync();
        await WebProgress_RemoveAsync();
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
        var sw = System.Diagnostics.Stopwatch.StartNew();
        Log($"=== 분석항목 코드 스크래핑 시작 [{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ===");
        Post("분석항목 코드 수집 중...", "#aaaaaa");        await WebProgress_InjectAsync("ETA 분석DB 업데이트 중...", 100, "#1a6f2a");
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
            Log("[분석DB] add_meas_item 요소 없음 또는 옵션 없음 — 모달이 열린 상태인지 확인하세요.");
            await WebProgress_RemoveAsync();
            Post("분석항목을 찾을 수 없습니다 (모달을 열고 다시 시도하세요)", "#ffaa44");
            return;
        }
        Log($"[분석DB] raw JSON 길이: {json.Length}자");

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
        catch (Exception ex) { Log($"[분석DB] JSON 파싱 오류: {ex.Message}\n{ex.StackTrace}"); throw; }

        Log($"[분석DB] {items.Count}개 수집. 샘플: {string.Join(", ", items.Take(5).Select(i => $"{i.분야}|{i.항목구분}|{i.항목명}"))}");
        Post($"분석항목 {items.Count}개 파싱 완료 — DB 저장 중...", "#aaaaaa");
        await WebProgress_UpdateAsync(50, 100, $"항목 {items.Count}개 파싱 완료 — DB 저장 중...");

        int saved = MeasurerService.SaveAnalysisItems(items);
        sw.Stop();
        Log($"[분석DB] 저장 완료: {saved}개 반영 / {items.Count}개 (소요 {sw.Elapsed.TotalSeconds:F1}초)");
        await WebProgress_UpdateAsync(100, 100, $"분석항목 {items.Count}개 저장 완료");
        Post($"분석항목 {items.Count}개 저장 완료", "#88cc88");
        await WebProgress_RemoveAsync();
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
                '[class*=`close`]', '[class*=`Close`]',
                '[id*=`close`]',    '[id*=`Close`]',
                '.layer-close',  '.popup-close',  '.modal-close',
                '.btn-close',    '.btn_close',    '.pop_close',
                '.layer_close',  '.ui-dialog-titlebar-close',
                '[title=`닫기`]', '[alt=`닫기`]'
            ].join(',');
            try {
                document.querySelectorAll(selectors).forEach(function(el) {
                    if (el.offsetParent !== null) el.click();
                });
            } catch(e) {}

            var closeTexts = ['닫기', '확인', '×', '✕', 'X', 'Close', 'OK'];
            var btns = document.querySelectorAll('button, a, input[type=`button`], input[type=`submit`]');
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
                $"[{DateTime.Now:HH:mm:ss.fff}] {msg}{Environment.NewLine}");
            Debug.WriteLine($"[측정인] {msg}");
        }
        catch { }
    }

    // 세션 시작 시 로그 헤더 기록
    private static void LogSessionStart()
    {
        try
        {
            File.AppendAllText(LogPath,
                $"{Environment.NewLine}{'=',60}{Environment.NewLine}" +
                $"  세션 시작: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                $"  로그 경로: {LogPath}{Environment.NewLine}" +
                $"{'=',60}{Environment.NewLine}");
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

    // ── 웹페이지 JS 프로그레스바 주입 헬퍼 ─────────────────────────────────────
    private async Task WebProgress_InjectAsync(string title, int total, string color = "#1a6fc4")
    {
        string safeSrc = title.Replace("\\", "\\\\").Replace("'", "\\'");
        int    tot     = total;
        string js = $@"(function(){{
            var old = document.getElementById('__eta_pb__');
            if (old) old.remove();
            var d   = document.createElement('div');
            d.id    = '__eta_pb__';
            d.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:999999;'
                + 'background:rgba(18,18,36,0.97);color:#ddd;font-size:12px;'
                + 'padding:7px 16px;display:flex;align-items:center;gap:10px;'
                + 'box-shadow:0 2px 8px rgba(0,0,0,0.6);font-family:sans-serif;';
            var lbl = document.createElement('span');
            lbl.id = '__eta_pb_lbl__';
            lbl.textContent = '{safeSrc}';
            lbl.style.cssText = 'min-width:280px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';
            var wrap = document.createElement('div');
            wrap.style.cssText = 'flex:1;background:#444;border-radius:4px;height:8px;';
            var bar = document.createElement('div');
            bar.id = '__eta_pb_bar__';
            bar.style.cssText = 'height:8px;background:{color};border-radius:4px;width:0%;transition:width 0.3s;';
            wrap.appendChild(bar);
            var cnt = document.createElement('span');
            cnt.id = '__eta_pb_cnt__';
            cnt.textContent = '0 / {tot}';
            cnt.style.cssText = 'min-width:60px;text-align:right;color:#aaa;';
            d.appendChild(lbl); d.appendChild(wrap); d.appendChild(cnt);
            document.body.prepend(d);
            return 'OK';
        }})()";
        try {{ await Evaluate(js); }} catch {{ }}
    }

    private async Task WebProgress_UpdateAsync(int current, int total, string msg)
    {
        int    pct     = total > 0 ? (int)(current * 100.0 / total) : 100;
        string safeMsg = msg.Replace("\\", "\\\\").Replace("'", "\\'");
        string js = $@"(function(){{
            var l = document.getElementById('__eta_pb_lbl__');
            var b = document.getElementById('__eta_pb_bar__');
            var c = document.getElementById('__eta_pb_cnt__');
            if (l) l.textContent = '{safeMsg}';
            if (b) b.style.width = '{pct}%';
            if (c) c.textContent = '{current} / {total}';
        }})()";
        try {{ await Evaluate(js); }} catch {{ }}
    }

    private async Task WebProgress_RemoveAsync()
    {
        const string js = "(function(){{ var e=document.getElementById('__eta_pb__'); if(e) e.remove(); }})()";
        try { await Evaluate(js); } catch { }
    }
}
