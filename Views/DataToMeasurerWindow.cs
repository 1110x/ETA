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
    private readonly Button    _checkBtn = new();  // 측정인 페이지 확인 버튼
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
        _matchTb.Text        = "▶ [측정인 페이지 확인] 버튼: 현재 브라우저에서 측정인 열람 중인 페이지와 로드된 시료를 비교합니다.";
        _matchTb.TextWrapping = TextWrapping.Wrap;

        StyleBtn(_checkBtn, "측정인 페이지 확인", "#1e3a5f");
        _checkBtn.Width = 140;
        StyleBtn(_inputBtn, "측정인에 입력", "#264026");
        StyleBtn(_closeBtn, "닫기",         "#3a2a2a");
        _checkBtn.Click += async (_, _) => await CheckMeasurerPageAsync();
        _inputBtn.Click += OnInputClick;
        _closeBtn.Click += (_, _) => Close();

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 8,
            Margin      = new Thickness(8, 0, 8, 10),
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        btnRow.Children.Add(_checkBtn);
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

    private void OnInputClick(object? sender, RoutedEventArgs e)
    {
        // TODO: 측정인.kr CDP 자동 입력 연동
        // 현재: 측정인 로그인 창을 열어 수동 입력을 안내
        SetStatus("측정인 로그인 창을 여세요. CDP 자동 입력은 다음 버전에서 지원 예정입니다.", "#ffaa44");

        var loginWin = new MeasurerLoginWindow();
        loginWin.Show(this);
    }

    // ════════════════════════════════════════════════════════════════════════
    //  측정인 페이지 확인 (CDP)
    // ════════════════════════════════════════════════════════════════════════

    private async Task CheckMeasurerPageAsync()
    {
        _matchTb.Foreground = new SolidColorBrush(Color.Parse("#aaaacc"));
        _matchTb.Text = "페이지 확인 중...";
        _checkBtn.IsEnabled = false;
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
        finally { _checkBtn.IsEnabled = true; }
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
