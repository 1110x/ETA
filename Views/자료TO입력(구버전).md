using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
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

namespace ETA.Views;

/// <summary>
/// 자료TO측정인 — 분석 결과를 측정인 사이트 입력 형식으로 정리해서 표시하고 직접 입력하는 창.
/// 시험성적서 메뉴 BT7("자료TO측정인") 클릭 시 열림.
/// </summary>
public class DataToMeasurerWindow : Window
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");
    private static readonly object LogLock = new();
    private static readonly string InputLogPath = ResolveInputLogPath();

    private static string ResolveInputLogPath()
    {
        try
        {
            string cwdData = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            if (Directory.Exists(cwdData))
                return Path.Combine(cwdData, "측정in.log");
        }
        catch { }

        return Path.Combine(AppContext.BaseDirectory, "측정in.log");
    }

    private static void AppendInputLog(string msg)
    {
        try
        {
            lock (LogLock)
            {
                var dir = Path.GetDirectoryName(InputLogPath);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(
                    InputLogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {msg}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch { }
    }

    private static string Shorten(string s, int max = 120)
    {
        if (string.IsNullOrEmpty(s)) return s;
        s = s.Replace("\r", " ").Replace("\n", " ");
        return s.Length <= max ? s : s[..max] + "...";
    }

    // ── 분석기기 드롭다운 선택지 (측정인.kr 기준) ─────────────────────────────
    private static readonly string[] InstrumentOptions =
    {
        "ICP-OES", "ICP-MS", "원자흡광광도계(AAS)", "자외선분광광도계(UV)",
        "이온크로마토그래프(IC)", "GC", "GC-MSD", "HPLC",
        "TOC분석기", "수질자동측정기", "pH미터", "DO미터",
        "탁도계", "전도도계", "부유물측정기", "COD분석기",
        "BOD분석기", "암모니아이온선택전극", "기타"
    };

    // ── UI 컨트롤 ────────────────────────────────────────────────────────────
    private readonly DataGrid  _grid     = new();
    private readonly TextBlock _statusTb = new();
    private readonly TextBlock _matchTb  = new();  // 측정인 페이지 매칭 결과 표시
    private readonly Button    _inputBtn = new();
    private readonly Button    _closeBtn = new();

    // ── 데이터 ───────────────────────────────────────────────────────────────
    private List<MeasurerInputRow> _rows = new();
    private readonly ETA.Models.SampleRequest? _sourceSample; // 선택 노드로 필터링

    public DataToMeasurerWindow() : this(null) {{ }}

    public DataToMeasurerWindow(ETA.Models.SampleRequest? selectedSample)
    {{
        _sourceSample = selectedSample;
        Title  = selectedSample == null
            ? "자료 TO 측정인 — 분석 결과 입력 준비"
            : $"자료 TO 측정인 — {selectedSample.약칭} / {selectedSample.시료명}";
        Width  = 1280;
        Height = 760;
        CanResize = true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(Color.Parse("#1e1e2e"));

        BuildUi();
        Loaded += OnLoaded;
    }}

    // ════════════════════════════════════════════════════════════════════════
    //  UI 구성
    // ════════════════════════════════════════════════════════════════════════

    private void BuildUi()
    {
        // ── 제목 ──────────────────────────────────────────────────────────
        var title = new TextBlock
        {
            Text       = "자료 TO 측정인",
            FontFamily = Font,
            FontSize   = 16,
            FontWeight = FontWeight.Bold,
            Foreground = Brushes.WhiteSmoke,
            Margin     = new Thickness(12, 10, 0, 4),
        };

        var subtitle = new TextBlock
        {
            Text       = _sourceSample != null
                ? $"선택된 시료: {_sourceSample.약칭} / {_sourceSample.시료명} (채취일: {_sourceSample.채취일자})"
                : "데이터공유를 통해 시험성적서의 분석 결과를 측정인 사이트에 입력할 수 있도록 정리합니다. 항목을 확인 후 [측정인에 입력] 버튼을 클릭하세요.",
            FontFamily = Font,
            FontSize   = 11,
            Foreground = new SolidColorBrush(Color.Parse("#8888bb")),
            Margin     = new Thickness(12, 0, 0, 6),
            TextWrapping = TextWrapping.Wrap,
        };

        // ── 상태 표시 ──────────────────────────────────────────────────────
        _statusTb.FontFamily = Font;
        _statusTb.FontSize   = 11;
        _statusTb.Foreground = new SolidColorBrush(Color.Parse("#88cc88"));
        _statusTb.Margin     = new Thickness(12, 0, 0, 6);
        _statusTb.Text       = "데이터 로드 중...";

        // ── DataGrid ───────────────────────────────────────────────────────
        _grid.IsReadOnly           = false;
        _grid.CanUserResizeColumns = true;
        _grid.Background           = new SolidColorBrush(Color.Parse("#1e1e2e"));
        _grid.Foreground           = Brushes.WhiteSmoke;
        _grid.GridLinesVisibility  = DataGridGridLinesVisibility.All;
        _grid.HorizontalGridLinesBrush = new SolidColorBrush(Color.Parse("#333355"));
        _grid.VerticalGridLinesBrush   = new SolidColorBrush(Color.Parse("#333355"));
        _grid.FontFamily           = Font;
        _grid.FontSize             = 12;
        _grid.Margin               = new Thickness(8, 0, 8, 8);

        BuildGridColumns();

        // ── 버튼 열 ────────────────────────────────────────────────────────
        // ── 측정인 페이지 확인 영역 ─────────────────────────────────────
        _matchTb.FontFamily  = Font;
        _matchTb.FontSize    = 11;
        _matchTb.Foreground  = new SolidColorBrush(Color.Parse("#8888bb"));
        _matchTb.Margin      = new Thickness(12, 2, 12, 2);
        _matchTb.Text        = "▶ [측정값 입력 진행] 버튼: 현재 브라우저의 측정인 페이지와 로드된 시료를 확인한 뒤 입력을 진행합니다.";
        _matchTb.TextWrapping = TextWrapping.Wrap;

        StyleBtn(_inputBtn, "측정값 입력 진행", "#264026");
        _inputBtn.Width = 150;
        StyleBtn(_closeBtn, "닫기",         "#3a2a2a");
        _inputBtn.Click += async (_, _) => await OnInputClickAsync();
        _closeBtn.Click += (_, _) => Close();

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 8,
            Margin      = new Thickness(8, 0, 8, 10),
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        btnRow.Children.Add(_inputBtn);
        btnRow.Children.Add(_closeBtn);

        // ── 레이아웃 조립 ──────────────────────────────────────────────────
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        root.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        root.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        root.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        root.RowDefinitions.Add(new RowDefinition(new GridLength(1, GridUnitType.Star)));
        root.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        Grid.SetRow(title,     0);
        Grid.SetRow(subtitle,  1);
        Grid.SetRow(_matchTb,  2);
        Grid.SetRow(_statusTb, 3);
        Grid.SetRow(_grid,     4);
        Grid.SetRow(btnRow,    5);

        root.Children.Add(title);
        root.Children.Add(subtitle);
        root.Children.Add(_matchTb);
        root.Children.Add(_statusTb);
        root.Children.Add(_grid);
        root.Children.Add(btnRow);

        Content = root;
    }

    private void BuildGridColumns()
    {
        _grid.Columns.Clear();

        AddCol("분석항목",         nameof(MeasurerInputRow.분석항목),    120, true);
        AddCol("법적기준",         nameof(MeasurerInputRow.법적기준),     80,  false);
        AddCol("분석결과",         nameof(MeasurerInputRow.분석결과),     80,  false);
        AddCol("측정분석방법",     nameof(MeasurerInputRow.측정분석방법),  120, false);
        AddCol("분석장비",         nameof(MeasurerInputRow.분석장비),     120, false);
        AddCol("분석자",           nameof(MeasurerInputRow.분석자),       80,  false);
        AddCol("분석일(시작)",     nameof(MeasurerInputRow.분석일시작),    90,  false);
        AddCol("시작시간",         nameof(MeasurerInputRow.시작시간),     70,  false);
        AddCol("분석일(종료)",     nameof(MeasurerInputRow.분석일종료),    90,  false);
        AddCol("종료시간",         nameof(MeasurerInputRow.종료시간),     70,  false);
        AddCol("시료용량",         nameof(MeasurerInputRow.시료용량),     70,  false);
        AddCol("비고",             nameof(MeasurerInputRow.비고),         100, false);

    }

    private void AddCol(string header, string path, double width, bool readOnly)
    {
        _grid.Columns.Add(new DataGridTextColumn
        {
            Header   = header,
            Binding  = new Binding(path),
            Width    = new DataGridLength(width),
            IsReadOnly = readOnly,
        });
    }

    private static void StyleBtn(Button btn, string label, string bg)
    {
        btn.Content         = label;
        btn.Width           = 120;
        btn.Height          = 32;
        btn.FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");
        btn.FontSize        = 12;
        btn.Background      = new SolidColorBrush(Color.Parse(bg));
        btn.Foreground      = Brushes.WhiteSmoke;
        btn.BorderThickness = new Thickness(0);
        btn.Padding         = new Thickness(0);
        btn.Cursor          = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand);
    }

    // ════════════════════════════════════════════════════════════════════════
    //  데이터 로드
    // ════════════════════════════════════════════════════════════════════════

    private void OnLoaded(object? sender, RoutedEventArgs e)
    {
        try
        {
            LoadData();
        }
        catch (Exception ex)
        {
            SetStatus($"로드 오류: {ex.Message}", "#ee4444");
        }
    }

    private void LoadData()
    {
        _rows.Clear();

        Debug.WriteLine($"[DataToMeasurer] LoadData 시작, _sourceSample={(_sourceSample == null ? "null" : _sourceSample.시료명)}");

        // 현재 로그인 사용자
        string analyst = CurrentUserManager.Instance.CurrentUserId;

        // 분석정보 메타 (방법, 기기)
        List<ETA.Models.AnalysisItem> allItems;
        try   { allItems = AnalysisService.GetAllItems(); }
        catch (Exception ex) { allItems = new(); Debug.WriteLine($"[DataToMeasurer] AnalysisService 오류: {ex.Message}"); }
        var meta = allItems.ToDictionary(x => x.Analyte, x => x, StringComparer.OrdinalIgnoreCase);
        Debug.WriteLine($"[DataToMeasurer] 메타 {meta.Count}개 로드");

        // 방류기준 (방류기준표)
        // TODO: 방류기준표 서비스 연동 — 현재 빈값으로 처리
        var legalStd = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 선택 시료만 수집 (_sourceSample)
        if (_sourceSample == null)
        {
            SetStatus("⚠️ 시험성적서 Show1에서 시료 노드를 먼저 선택하세요.", "#ffaa44");
            return;
        }

        Debug.WriteLine($"[DataToMeasurer] 분석결과 항목 수: {_sourceSample.분석결과.Count}");

        // 분장표준처리 — 채취일자 기준 담당자 조회
        var managers = AnalysisRequestService.GetManagersByDate(_sourceSample.채취일자);
        var stdDays  = AnalysisRequestService.GetStandardDaysInfo();
        Debug.WriteLine($"[DataToMeasurer] managers {managers.Count}개, stdDays {stdDays.Count}개");

        // 분석 결과가 있는 항목만 행 생성
        foreach (var kv in _sourceSample.분석결과)
        {
            var analyte = kv.Key;
            var result  = kv.Value;
            if (string.IsNullOrWhiteSpace(result)) continue;

            meta.TryGetValue(analyte, out var m);

            _rows.Add(new MeasurerInputRow
            {
                약칭         = _sourceSample.약칭,
                시료명       = _sourceSample.시료명,
                채취일자     = _sourceSample.채취일자,
                분석항목     = analyte,
                법적기준     = legalStd.TryGetValue(analyte, out var std) ? std : "",
                분석결과     = result,
                분석기기     = m?.instrument ?? "",
                측정분석방법 = m?.Method     ?? "",
                분석장비     = m?.instrument ?? "",
                분석자       = managers.TryGetValue(analyte, out var mgr) && !string.IsNullOrEmpty(mgr) ? mgr : analyst,
                분석일시작   = _sourceSample.채취일자,
                시작시간     = string.IsNullOrWhiteSpace(_sourceSample.채취시간) ? "09:00" : _sourceSample.채취시간,
                분석일종료   = string.IsNullOrWhiteSpace(_sourceSample.분석종료일)
                    ? (stdDays.TryGetValue(analyte, out var si) && si.days > 0
                        ? AddBusinessDays(_sourceSample.채취일자, si.days)
                        : _sourceSample.채취일자)
                    : _sourceSample.분석종료일,
                종료시간     = "18:00",
                시료용량     = "",
                비고         = "",
            });
        }

        _grid.ItemsSource = _rows;
        SetStatus($"입력 항목 수: {_rows.Count}건 로드 완료 — {_sourceSample.약칭} / {_sourceSample.시료명}", "#88cc88");
        Debug.WriteLine($"[DataToMeasurer] {_rows.Count}건 로드");
    }

    // ════════════════════════════════════════════════════════════════════════
    //  입력 버튼
    // ════════════════════════════════════════════════════════════════════════

    private async Task OnInputClickAsync()
    {
        AppendInputLog("측정값 입력 진행 시작");

        if (_rows.Count == 0)
        {
            SetStatus("입력할 항목이 없습니다.", "#ffaa44");
            AppendInputLog("입력할 항목 없음");
            return;
        }

        SetStatus("측정인 페이지 확인 후 입력 진행 중...", "#88cc88");
        _inputBtn.IsEnabled = false;
        try
        {
            // 로그인 창을 띄우지 않고 현재 열려 있는 측정인 페이지 기준으로 바로 진행
            await CheckMeasurerPageAsync();

            var result = await TryAutoInputResultsAsync();
            if (result.ok)
            {
                SetStatus($"측정값 입력 완료: {result.message}", "#88cc88");
                AppendInputLog($"입력 성공: {result.message}");
            }
            else
            {
                SetStatus($"측정값 입력 실패: {Shorten(result.message)} (상세: 측정in.log)", "#ee6666");
                AppendInputLog($"입력 실패: {result.message}");
            }
        }
        catch (Exception ex)
        {
            SetStatus($"측정값 입력 실패: {Shorten(ex.Message)} (상세: 측정in.log)", "#ee6666");
            AppendInputLog($"예외: {ex}");
        }
        finally
        {
            _inputBtn.IsEnabled = true;
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    //  측정인 페이지 확인 (CDP)
    // ════════════════════════════════════════════════════════════════════════

    private async Task CheckMeasurerPageAsync()
    {
        _matchTb.Foreground = new SolidColorBrush(Color.Parse("#aaaacc"));
        _matchTb.Text = "페이지 확인 중...";
        try
        {
            string h2Text = await GetEditTitleFromCdpAsync();
            if (string.IsNullOrEmpty(h2Text))
            {
                _matchTb.Text = "⚠️ 측정인 탭을 찾을 수 없습니다. 브라우저에서 측정인.kr을 먼저 열어주세요.";
                _matchTb.Foreground = new SolidColorBrush(Color.Parse("#ffaa44"));
                return;
            }

            // h2 형식: "20260329-130000 / 수질 / 2026-03-29 / 업체명 / 시료명"
            var parts = h2Text.Split(new[] { " / " }, StringSplitOptions.None);
            string pageNo     = parts.Length > 0 ? parts[0].Trim() : "";
            string pageDate   = parts.Length > 2 ? parts[2].Trim() : "";
            string pageCompany= parts.Length > 3 ? parts[3].Trim() : "";
            string pageSample = parts.Length > 4 ? parts[4].Trim() : "";

            string myCompany = _sourceSample?.약칭   ?? "";
            string mySample  = _sourceSample?.시료명 ?? "";
            string myDate    = _sourceSample?.채취일자 ?? "";

            double companySim = Similarity(myCompany, pageCompany);
            double sampleSim  = Similarity(mySample,  pageSample);
            bool   dateMatch  = !string.IsNullOrEmpty(myDate) && pageDate.StartsWith(myDate);

            string dateIcon    = dateMatch  ? "✅" : "⚠️";
            string companyIcon = companySim > 0.5 ? "✅" : "⚠️";
            string sampleIcon  = sampleSim  > 0.4 ? "✅" : "⚠️";

            _matchTb.Text =
                $"[페이지] {h2Text}\n"
                + $"  {dateIcon} 날짜: {pageDate} | "
                + $"{companyIcon} 업체: {pageCompany} ↔ {myCompany} ({companySim * 100:F0}%) | "
                + $"{sampleIcon} 시료: {pageSample} ↔ {mySample} ({sampleSim * 100:F0}%)";

            bool allGood = dateMatch && companySim > 0.5 && sampleSim > 0.4;
            _matchTb.Foreground = new SolidColorBrush(allGood
                ? Color.Parse("#66cc88")
                : Color.Parse("#ffaa44"));

            Debug.WriteLine($"[DataToMeasurer] h2={h2Text}, 업체유사={companySim*100:F0}%, 시료유사={sampleSim*100:F0}%");
        }
        catch (Exception ex)
        {
            _matchTb.Text = $"⚠️ 확인 오류: {ex.Message}";
            _matchTb.Foreground = new SolidColorBrush(Color.Parse("#ee4444"));
        }
    }

    // CDP 연결하여 현재 브라우저의 #edit_title h2 텍스트를 읽는다
    private static async Task<string> GetEditTitleFromCdpAsync()
    {
        const int port = 9222;
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
            string json = await http.GetStringAsync($"http://localhost:{port}/json");

            string? wsUrl = null;
            using (var doc = JsonDocument.Parse(json))
            {
                // 1순위: 측정인 탭
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    if (item.TryGetProperty("type", out var t) && t.GetString() == "page" &&
                        item.TryGetProperty("url",  out var u) && (u.GetString() ?? "").Contains("측정인"))
                    {
                        wsUrl = item.GetProperty("webSocketDebuggerUrl").GetString();
                        break;
                    }
                }
                // 2순위: 첫 번째 page 탭
                if (wsUrl == null)
                    foreach (var item in doc.RootElement.EnumerateArray())
                        if (item.TryGetProperty("type", out var t) && t.GetString() == "page" &&
                            item.TryGetProperty("webSocketDebuggerUrl", out var ws))
                        { wsUrl = ws.GetString(); break; }
            }
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
            var buf = Encoding.UTF8.GetBytes(cmd);
            await socket.SendAsync(buf, WebSocketMessageType.Text, true, cts.Token);

            var recvBuf = new byte[8192];
            var recv = await socket.ReceiveAsync(recvBuf, cts.Token);
            string resp = Encoding.UTF8.GetString(recvBuf, 0, recv.Count);

            using var respDoc = JsonDocument.Parse(resp);
            return respDoc.RootElement
                .GetProperty("result").GetProperty("result").GetProperty("value")
                .GetString() ?? "";
        }
        catch { return ""; }
    }

    // 현재 측정인 탭에 분석결과 값을 자동 입력 (분석항목명 기준 매칭)
    private async Task<(bool ok, string message)> TryAutoInputResultsAsync()
    {
        const int port = 9222;
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
            string json = await http.GetStringAsync($"http://localhost:{port}/json");

            string? wsUrl = null;
            using (var doc = JsonDocument.Parse(json))
            {
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    if (item.TryGetProperty("type", out var tabType) && tabType.GetString() == "page" &&
                        item.TryGetProperty("url", out var u) && (u.GetString() ?? "").Contains("측정인"))
                    {
                        wsUrl = item.GetProperty("webSocketDebuggerUrl").GetString();
                        break;
                    }
                }

                if (wsUrl == null)
                {
                    foreach (var item in doc.RootElement.EnumerateArray())
                    {
                        if (item.TryGetProperty("type", out var tabType2) && tabType2.GetString() == "page" &&
                            item.TryGetProperty("webSocketDebuggerUrl", out var ws))
                        {
                            wsUrl = ws.GetString();
                            break;
                        }
                    }
                }
            }

            if (wsUrl == null)
                return (false, "측정인 탭을 찾지 못했습니다.");

            var payload = _rows.Select(r => new
            {
                analyte = r.분석항목,
                result = r.분석결과,
            }).ToList();

            string dataJson = JsonSerializer.Serialize(payload);
            string dataJs = dataJson
                .Replace("\\", "\\\\")
                .Replace("'", "\\'");

            // ── WebSocket + CDP 헬퍼 준비 ──
            using var socket = new ClientWebSocket();
            int timeoutSeconds = Math.Clamp(30 + (_rows.Count * 3), 90, 600);
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
            await socket.ConnectAsync(new Uri(wsUrl), cts.Token);

            int cmdId = 0;

            // CDP 명령 전송 + 해당 id 응답 대기
            async Task<JsonDocument> CdpSendAsync(string method, object parms)
            {
                var id = ++cmdId;
                var cmdObj = new Dictionary<string, object> { ["id"] = id, ["method"] = method, ["params"] = parms };
                var bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(cmdObj));
                await socket.SendAsync(bytes, WebSocketMessageType.Text, true, cts.Token);
                for (int i = 0; i < 200; i++)
                {
                    using var ms = new MemoryStream();
                    WebSocketReceiveResult recv;
                    var seg = new byte[65536];
                    do
                    {
                        recv = await socket.ReceiveAsync(new ArraySegment<byte>(seg), cts.Token);
                        if (recv.MessageType == WebSocketMessageType.Close)
                            throw new InvalidOperationException("CDP WebSocket closed");
                        ms.Write(seg, 0, recv.Count);
                    } while (!recv.EndOfMessage);
                    var doc = JsonDocument.Parse(ms.ToArray());
                    if (doc.RootElement.TryGetProperty("id", out var idProp) && idProp.GetInt32() == id)
                        return doc;
                    doc.Dispose();
                }
                throw new TimeoutException($"CDP 응답 대기 시간 초과 id={id}");
            }

            // JS 실행 헬퍼
            async Task<string> JsEvalAsync(string expression, bool isAsync = false)
            {
                using var doc = await CdpSendAsync("Runtime.evaluate",
                    new { expression, returnByValue = true, awaitPromise = isAsync });
                var res = doc.RootElement.GetProperty("result");
                if (res.TryGetProperty("exceptionDetails", out var ex))
                    throw new Exception($"JS 오류: {ex}");
                var evalRes = res.GetProperty("result");
                if (evalRes.TryGetProperty("value", out var v))
                    return v.ValueKind == JsonValueKind.String ? (v.GetString() ?? "") : v.GetRawText();
                return "";
            }

            // CDP 단일 클릭
            async Task CdpClickAsync(double x, double y)
            {
                var left = "left";
                await CdpSendAsync("Input.dispatchMouseEvent",
                    new { type = "mousePressed", x, y, button = left, clickCount = 1 });
                await CdpSendAsync("Input.dispatchMouseEvent",
                    new { type = "mouseReleased", x, y, button = left, clickCount = 1 });
            }

            // ── Phase 0: 높이 확장 → 가상스크롤 비활성화, 전체 행 DOM 렌더링 ──
            string expandJs = @"(function(){
                var rootGrid = document.getElementById('gridAnalySampAnzeDataAirItemList1');
                if(!rootGrid) return 'noGrid';
                var saved = {};
                // 루트 그리드 높이
                saved.rootH = rootGrid.style.height;
                rootGrid.style.height = '5000%';
                // 내부 정확한 그리드 높이
                var inner = rootGrid.querySelector('.rg-exact');
                if(inner){ saved.innerH = inner.style.height; inner.style.height = '4200px'; }
                // 부모 div overflow
                var parent = rootGrid.closest('div[style*=""overflow""]') || rootGrid.parentElement;
                if(parent){ saved.parentOF = parent.style.overflow; parent.style.overflow = 'visible'; }
                // 뷰포트 높이 늘리기
                var vp = rootGrid.querySelector('.rg-viewport, .rg-body');
                if(vp){ saved.vpH = vp.style.height; saved.vpOF = vp.style.overflow; vp.style.height = '99999px'; vp.style.overflow = 'visible'; }
                window.__etaHeightSaved = saved;
                return 'ok';
            })()";
            await JsEvalAsync(expandJs);
            await Task.Delay(300, cts.Token); // DOM 리플로우 대기

            // ── Phase 1: JS로 행 매칭만 수행 (좌표는 수집 안 함) ──
            string findScript = $@"(function(){{
                function norm(s) {{
                    return (s || '').toString().toLowerCase()
                        .replace(/\s+/g, '')
                        .replace(/[()\[\]\-_,.:/%]/g, '')
                        .replace(/[^0-9a-z가-힣]/g, '');
                }}
                var rows = JSON.parse('{dataJs}');
                var tbody = document.evaluate(
                    ""//*[@id='gridAnalySampAnzeDataAirItemList1']/div/div[1]/div/div[1]/div[1]/div[4]/table/tbody"",
                    document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null
                ).singleNodeValue;
                var trs = [];
                if (tbody) trs = Array.from(tbody.children).filter(function(n){{ return n && n.tagName === 'TR' && n.querySelector('td:nth-child(4)') && n.querySelector('td:nth-child(6)'); }});
                if (!trs.length) {{
                    var scope = document.getElementById('gridAnalySampAnzeDataAirItemList1') || document;
                    trs = Array.from(scope.querySelectorAll('table tbody tr')).filter(function(tr){{ return !!(tr.querySelector('td:nth-child(4)') && tr.querySelector('td:nth-child(6)')); }});
                }}
                var rowByKey = {{}}, rowKeyList = [];
                for (var i = 0; i < rows.length; i++) {{
                    var rk = norm(rows[i].analyte);
                    if (!rk || rk === norm('현장측정 및 시료채취')) continue;
                    if (!(rk in rowByKey)) {{ rowByKey[rk] = rows[i]; rowKeyList.push(rk); }}
                }}
                var matched = [], usedKeys = {{}}, unmatched = [];
                for (var r = 0; r < trs.length; r++) {{
                    var tr = trs[r];
                    var td4 = tr.querySelector('td:nth-child(4) div') || tr.querySelector('td:nth-child(4)');
                    var label = (td4.innerText || td4.textContent || '').trim();
                    var dk = norm(label);
                    if (!dk) continue;
                    var item = rowByKey[dk], mk = dk;
                    if (!item) {{
                        var fk = rowKeyList.find(function(k){{ return k.indexOf(dk)>=0 || dk.indexOf(k)>=0; }});
                        if (fk) {{ item = rowByKey[fk]; mk = fk; }}
                    }}
                    if (!item || usedKeys[mk]) continue;
                    usedKeys[mk] = true;
                    var v = (item.result || '').toString().trim();
                    matched.push({{ a: item.analyte, v: v, ri: r }});
                }}
                for (var i = 0; i < rowKeyList.length; i++) {{
                    if (!usedKeys[rowKeyList[i]]) unmatched.push(rowByKey[rowKeyList[i]].analyte);
                }}
                return JSON.stringify({{ matched: matched, total: rows.length, domRows: trs.length, matchedCount: matched.length, unmatched: unmatched.slice(0,5) }});
            }})()";

            // ── Phase 1 실행 ──
            string phase1Json = await JsEvalAsync(findScript);

            int total = 0, domRows = 0, matchedCount = 0;
            int filled = 0;
            var commitFailed = new List<string>();
            var valuePreview = new List<string>();
            string unmatchedPreview = "";
            int unmatchedCount = 0;

            // 매칭 목록 파싱
            var matchedItems = new List<(string analyte, string value, int rowIndex)>();
            using (var p1Doc = JsonDocument.Parse(phase1Json))
            {
                var p1 = p1Doc.RootElement;
                total = p1.GetProperty("total").GetInt32();
                domRows = p1.GetProperty("domRows").GetInt32();
                matchedCount = p1.GetProperty("matchedCount").GetInt32();

                if (p1.TryGetProperty("unmatched", out var umArr) && umArr.ValueKind == JsonValueKind.Array)
                {
                    var umList = umArr.EnumerateArray().Select(x => x.GetString() ?? "").Where(x => x.Length > 0).ToList();
                    unmatchedCount = umList.Count;
                    unmatchedPreview = string.Join(", ", umList.Take(5));
                }

                foreach (var m in p1.GetProperty("matched").EnumerateArray())
                {
                    string a = m.GetProperty("a").GetString() ?? "";
                    string v = m.GetProperty("v").GetString() ?? "";
                    int ri = m.GetProperty("ri").GetInt32();
                    matchedItems.Add((a, v, ri));
                }
            }

            // ── Phase 2: 각 행(td6 분석결과 → td16 비고 "X") 순차 입력 ──
            double commitX = 0, commitY = 0;
            foreach (var (analyte, value, rowIndex) in matchedItems)
            {
                string safeAnalyte = analyte.Replace("\\", "\\\\").Replace("'", "\\'");

                // td 좌표 조회 헬퍼
                async Task<(double x, double y, double cx, double cy)?> GetCoordsForTdAsync(int tdIdx, string? subSel = null)
                {
                    string subSelJs = subSel != null ? $"'{subSel}'" : "null";
                    string js = $@"(function(){{
                        function norm(s){{ return (s||'').toString().toLowerCase().replace(/\s+/g,'').replace(/[()\[\]\-_,.:/%]/g,'').replace(/[^0-9a-z가-힣]/g,''); }}
                        var target = norm('{safeAnalyte}');
                        var tbody = document.evaluate(
                            ""//*[@id='gridAnalySampAnzeDataAirItemList1']/div/div[1]/div/div[1]/div[1]/div[4]/table/tbody"",
                            document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null
                        ).singleNodeValue;
                        if(!tbody) return JSON.stringify({{error:'noTbody'}});
                        var trs = Array.from(tbody.children).filter(function(n){{ return n && n.tagName==='TR' && n.querySelector('td:nth-child(4)'); }});
                        function getCoords(tr){{
                            var td = tr.querySelector('td:nth-child({tdIdx})');
                            var td3 = tr.querySelector('td:nth-child(3)');
                            if(!td) return null;
                            var sub = {subSelJs};
                            var clickEl = sub ? (td.querySelector(sub) || td) : td;
                            clickEl.scrollIntoView({{block:'center',behavior:'instant'}});
                            var r = clickEl.getBoundingClientRect();
                            var r3 = td3 ? td3.getBoundingClientRect() : r;
                            return {{x:r.x+r.width/2, y:r.y+r.height/2, cx:r3.x+r3.width/2, cy:r3.y+r3.height/2}};
                        }}
                        for(var i=0; i<trs.length; i++){{
                            var td4 = trs[i].querySelector('td:nth-child(4) div') || trs[i].querySelector('td:nth-child(4)');
                            var label = norm((td4.innerText||td4.textContent||'').trim());
                            if(label===target){{ var c=getCoords(trs[i]); if(c) return JSON.stringify(c); }}
                        }}
                        for(var i=0; i<trs.length; i++){{
                            var td4 = trs[i].querySelector('td:nth-child(4) div') || trs[i].querySelector('td:nth-child(4)');
                            var label = norm((td4.innerText||td4.textContent||'').trim());
                            if(label.indexOf(target)>=0 || target.indexOf(label)>=0){{ var c=getCoords(trs[i]); if(c) return JSON.stringify(c); }}
                        }}
                        return JSON.stringify({{error:'notFound'}});
                    }})()";
                    try
                    {
                        string coordJson = await JsEvalAsync(js);
                        using var cDoc = JsonDocument.Parse(coordJson);
                        if (cDoc.RootElement.TryGetProperty("error", out _)) return null;
                        return (cDoc.RootElement.GetProperty("x").GetDouble(),
                                cDoc.RootElement.GetProperty("y").GetDouble(),
                                cDoc.RootElement.GetProperty("cx").GetDouble(),
                                cDoc.RootElement.GetProperty("cy").GetDouble());
                    }
                    catch { return null; }
                }

                // 셀 클릭 + 입력
                async Task<bool> InputToTdAsync(int tdIdx, string inputText, string? subSel = null)
                {
                    var coords = await GetCoordsForTdAsync(tdIdx, subSel);
                    if (coords == null) return false;
                    var (cx, cy, _, _) = coords.Value;
                    await CdpClickAsync(cx, cy);
                    await Task.Delay(30, cts.Token);
                    await CdpSendAsync("Input.insertText", new { text = inputText });
                    return true;
                }

                // 날짜 셀: span.rg-button-calendar 클릭 → 입력 → Enter로 커밋
                async Task<bool> InputDateTdAsync(int tdIdx, string dateText)
                {
                    var coords = await GetCoordsForTdAsync(tdIdx, "span.rg-button-calendar");
                    if (coords == null) coords = await GetCoordsForTdAsync(tdIdx);
                    if (coords == null) return false;
                    var (cx, cy, _, _) = coords.Value;
                    await CdpClickAsync(cx, cy);
                    await Task.Delay(30, cts.Token);
                    await CdpSendAsync("Input.insertText", new { text = dateText });
                    // Enter로 값 커밋 (안 하면 포커스 이탈 시 사라짐)
                    await CdpSendAsync("Input.dispatchKeyEvent", new { type = "keyDown", key = "Enter", code = "Enter", windowsVirtualKeyCode = 13 });
                    await CdpSendAsync("Input.dispatchKeyEvent", new { type = "keyUp",   key = "Enter", code = "Enter", windowsVirtualKeyCode = 13 });
                    return true;
                }

                // 드롭다운 인풋 셀: 클릭 → 타이핑(필터) → 두 번째 옵션 선택
                async Task<bool> InputDropdownTdAsync(int tdIdx, string inputText)
                {
                    var coords = await GetCoordsForTdAsync(tdIdx);
                    if (coords == null) return false;
                    var (cx, cy, _, _) = coords.Value;
                    await CdpClickAsync(cx, cy);
                    await Task.Delay(80, cts.Token);
                    await CdpSendAsync("Input.insertText", new { text = inputText });
                    await Task.Delay(80, cts.Token);
                    // 필터 후 남은 목록의 두 번째 항목(첫 번째는 blank) 클릭
                    string optJs = @"(function(){
                        var opts = document.querySelectorAll(""[role='option']"");
                        if(opts.length < 2) return JSON.stringify({error:'lessThan2'});
                        var opt = opts[1];
                        opt.scrollIntoView({block:'nearest'});
                        var r = opt.getBoundingClientRect();
                        return JSON.stringify({x: r.x+r.width/2, y: r.y+r.height/2});
                    })()";
                    try
                    {
                        string optJson = await JsEvalAsync(optJs);
                        using var optDoc = JsonDocument.Parse(optJson);
                        if (optDoc.RootElement.TryGetProperty("error", out _)) return false;
                        await CdpClickAsync(
                            optDoc.RootElement.GetProperty("x").GetDouble(),
                            optDoc.RootElement.GetProperty("y").GetDouble());
                        return true;
                    }
                    catch { return false; }
                }

                // _rows에서 해당 행 찾기
                var row = _rows.FirstOrDefault(r => r.분석항목 == analyte);

                // ── td[5] 법적기준 "1" 입력 ──
                await InputToTdAsync(5, "1");

                // ── td[6] 분석결과 입력 ──
                if (!await InputToTdAsync(6, value))
                { commitFailed.Add(analyte + "(td6)"); continue; }

                // ── td[8] 측정분석방법 (컬럼4) ──
                if (row != null && !string.IsNullOrWhiteSpace(row.측정분석방법))
                    await InputDropdownTdAsync(8, row.측정분석방법);

                // ── td[9] 분석장비 (컬럼5) ──
                if (row != null && !string.IsNullOrWhiteSpace(row.분석장비))
                    await InputDropdownTdAsync(9, row.분석장비);

                // ── td[10] 분석자 (컬럼6) ──
                if (row != null && !string.IsNullOrWhiteSpace(row.분석자))
                    await InputDropdownTdAsync(10, row.분석자);

                // ── td[11] 시작일 (날짜필드 → div[2] 클릭) ──
                if (row != null && !string.IsNullOrWhiteSpace(row.분석일시작))
                    await InputDateTdAsync(11, row.분석일시작);

                // ── td[12] 시작시간 "09:00" ──
                await InputToTdAsync(12, "09:00");

                // ── td[13] 종료일 (날짜필드 → div[2] 클릭) ──
                if (row != null && !string.IsNullOrWhiteSpace(row.분석일종료))
                    await InputDateTdAsync(13, row.분석일종료);

                // ── td[14] 종료시간 "18:00" ──
                await InputToTdAsync(14, "18:00");

                // ── td[15] 시료용량 "100" ──
                await InputToTdAsync(15, "100");

                filled++;
                if (valuePreview.Count < 5)
                    valuePreview.Add($"[{analyte}→'{value}']");
            }

            // ── Phase 3: 높이 원복 ──
            string restoreJs = @"(function(){
                var rootGrid = document.getElementById('gridAnalySampAnzeDataAirItemList1');
                if(!rootGrid || !window.__etaHeightSaved) return 'noSaved';
                var s = window.__etaHeightSaved;
                rootGrid.style.height = s.rootH || '';
                var inner = rootGrid.querySelector('.rg-exact');
                if(inner && s.innerH !== undefined) inner.style.height = s.innerH;
                var parent = rootGrid.closest('div[style*=""overflow""]') || rootGrid.parentElement;
                if(parent && s.parentOF !== undefined) parent.style.overflow = s.parentOF;
                var vp = rootGrid.querySelector('.rg-viewport, .rg-body');
                if(vp){ if(s.vpH !== undefined) vp.style.height = s.vpH; if(s.vpOF !== undefined) vp.style.overflow = s.vpOF; }
                delete window.__etaHeightSaved;
                // 그리드 리사이즈 이벤트 → 스크롤바 복원
                window.dispatchEvent(new Event('resize'));
                if(rootGrid.rgGrid && rootGrid.rgGrid.refresh) rootGrid.rgGrid.refresh();
                return 'ok';
            })()";
            try { await JsEvalAsync(restoreJs); } catch { }

            string valuePreviewStr = string.Join(", ", valuePreview);
            string commitFailedPreview = string.Join(", ", commitFailed.Take(5));

            _matchTb.Text = $"자동입력 결과: 전체 {total} / DOM행 {domRows} / 행매칭 {matchedCount} / 입력완료 {filled}";
            _matchTb.Foreground = new SolidColorBrush(Color.Parse(filled > 0 ? "#66cc88" : "#ffaa44"));

            return (filled > 0, $"전체 {total}, DOM행 {domRows}, 매칭 {matchedCount}, 입력(실반영) {filled}, 미매칭 {unmatchedCount}, 미매칭샘플 [{unmatchedPreview}], 커밋실패 {commitFailed.Count}, 커밋실패샘플 [{commitFailedPreview}], 값프리뷰 [{valuePreviewStr}]");
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    // Dice coefficient (바이그램 기반 유사율 0.0~1.0)
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
        for (int i = 0; i < s.Length - 1; i++)
            yield return s.Substring(i, 2);
    }

    private static string AddBusinessDays(string startDateStr, int days)
    {
        if (!DateTime.TryParse(startDateStr, out var date) || days <= 0) return startDateStr;
        int added = 0;
        while (added < days)
        {
            date = date.AddDays(1);
            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                added++;
        }
        return date.ToString("yyyy-MM-dd");
    }

    private void SetStatus(string msg, string hex)
    {
        _statusTb.Text       = msg;
        _statusTb.Foreground = new SolidColorBrush(Color.Parse(hex));
    }
}

// ════════════════════════════════════════════════════════════════════════════
//  데이터 행 모델
// ════════════════════════════════════════════════════════════════════════════

/// <summary>측정인 입력 준비 행. DataGrid에 직접 바인딩.</summary>
public class MeasurerInputRow
{
    public string 약칭         { get; set; } = "";
    public string 시료명       { get; set; } = "";
    public string 채취일자     { get; set; } = "";
    public string 분석항목     { get; set; } = "";
    public string 법적기준     { get; set; } = "";
    public string 분석결과     { get; set; } = "";
    public string 분석기기     { get; set; } = "";
    public string 측정분석방법 { get; set; } = "";
    public string 분석장비     { get; set; } = "";
    public string 분석자       { get; set; } = "";
    public string 분석일시작   { get; set; } = "";
    public string 시작시간     { get; set; } = "";
    public string 분석일종료   { get; set; } = "";
    public string 종료시간     { get; set; } = "";
    public string 시료용량     { get; set; } = "";
    public string 비고         { get; set; } = "";
}
