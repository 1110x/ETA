using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;

namespace ETA.Views.Pages;

public partial class TestReportPage : UserControl
{
    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    /// ActivePageContent2 에 넣을 결과 리스트 컨트롤
    public event Action<Control?>? ResultListChanged;
    /// ActivePageContent3 에 넣을 수정 폼 컨트롤
    public event Action<Control?>? EditPanelChanged;

    // ── 상태 ────────────────────────────────────────────────────────────────
    private SampleRequest?                    _selectedSample;
    private List<AnalysisResultRow>           _resultRows  = new();
    private Dictionary<string, AnalysisItem>  _meta        = new();

    // 리스트 패널

    private readonly TextBlock  _listTitle  = new()
    {
        FontSize   = 12,
        FontFamily = Font,
        Foreground = new SolidColorBrush(Color.Parse("#8888bb")),
        Margin     = new Thickness(10, 6),
    };

    // 정렬 모드 / 방류기준 토글
    private bool _sortByDate   = false;
    // 선택된 시료 (일괄 출력용)
    private readonly System.Collections.Generic.HashSet<SampleRequest> _checkedSamples = new();
    private bool _showStandard = false;
    private ToggleSwitch?   _stdToggle;
    private ContentControl? _stdToggleContainer;

    // 선택된 결과 행
    private AnalysisResultRow? _selectedRow;
    private Border?            _selectedRowBorder;

    public TestReportPage()
    {
        InitializeComponent();
        _meta = TestReportService.GetAnalyteMeta();
    }

    // =========================================================================
    // 데이터 로드 (트리)
    // =========================================================================
    public void LoadData()
    {
        Log("LoadData() 시작");
        ReportTreeView.Items.Clear();
        _selectedSample = null;
        ResultListChanged?.Invoke(BuildListControl());
        EditPanelChanged?.Invoke(null);

        try
        {
            var companies = TestReportService.GetCompanyList();
            Log($"업체 {companies.Count}개");
            if (companies.Count == 0)
            {
                ReportTreeView.Items.Add(new TreeViewItem {
                    Header = new TextBlock { Text = "⚠️ 데이터 없음", FontSize = 12, FontFamily = Font,
                                            Foreground = new SolidColorBrush(Color.Parse("#cc8800")) }
                });
                return;
            }
            if (_sortByDate) BuildTreeByDate(companies);
            else             BuildTreeByCompany(companies);
        }
        catch (Exception ex) { Log("★ 크래시 ★ " + ex.Message); }
    }

    private void BuildTreeByCompany(List<string> companies)
    {
        foreach (var company in companies)
        {
            var node = new TreeViewItem { Tag = company, IsExpanded = false,
                Header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8,
                    Children = {
                        new TextBlock { Text = "🏭", FontSize = 16, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = company, FontSize = 13, FontFamily = Font,
                                        Foreground = Brushes.WhiteSmoke, VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            node.Items.Add(new TreeViewItem { Header = "로딩 중..." });
            node.PropertyChanged += (s, e) =>
            {
                if (e.Property == TreeViewItem.IsExpandedProperty && node.IsExpanded
                    && node.Items.Count == 1 && node.Items[0] is TreeViewItem d
                    && d.Header?.ToString() == "로딩 중...")
                    LoadSampleNodes(node, node.Tag?.ToString() ?? "");
            };
            ReportTreeView.Items.Add(node);
        }
    }

    private void BuildTreeByDate(List<string> companies)
    {
        var all = companies.SelectMany(co => TestReportService.GetSamplesByCompany(co))
                           .OrderByDescending(s => s.채취일자).ToList();
        foreach (var yg in all.GroupBy(s => s.채취일자.Length >= 4 ? s.채취일자[..4] : s.채취일자)
                               .OrderByDescending(g => g.Key))
        {
            var yn = new TreeViewItem { IsExpanded = false,
                Header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6,
                    Children = {
                        new TextBlock { Text = "📆", FontSize = 15, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = $"{yg.Key}년  ({yg.Count()}건)", FontSize = 13, FontFamily = Font,
                                        Foreground = Brushes.WhiteSmoke, VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            foreach (var mg in yg.GroupBy(s => s.채취일자.Length >= 7 ? s.채취일자[..7] : s.채취일자)
                                  .OrderByDescending(g => g.Key))
            {
                var mn = new TreeViewItem { IsExpanded = false,
                    Header = new TextBlock { Text = $"  📅 {mg.Key}  ({mg.Count()}건)",
                        FontSize = 12, FontFamily = Font, Foreground = new SolidColorBrush(Color.Parse("#aaaacc")) }
                };
                foreach (var s2 in mg.OrderByDescending(x => x.채취일자))
                    mn.Items.Add(MakeSampleNode(s2, showCompany: true));
                yn.Items.Add(mn);
            }
            ReportTreeView.Items.Add(yn);
        }
    }

    private TreeViewItem MakeSampleNode(SampleRequest sample, bool showCompany = false)
    {
        var sub = showCompany
            ? $"{sample.약칭}  {sample.채취일자}  {sample.분석결과.Count}항목"
            : $"{sample.채취일자}  {sample.분석결과.Count}항목";

        var chk = new CheckBox
        {
            IsChecked         = _checkedSamples.Contains(sample),
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Thickness(0, 0, 4, 0),
        };
        chk.IsCheckedChanged += (_, _) =>
        {
            if (chk.IsChecked == true) _checkedSamples.Add(sample);
            else                       _checkedSamples.Remove(sample);
        };

        return new TreeViewItem { Tag = sample,
            Header = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 4,
                Margin = new Thickness(4, 0, 0, 0),
                Children =
                {
                    chk,
                    new TextBlock { Text = "🧪", FontSize = 13, VerticalAlignment = VerticalAlignment.Center },
                    new StackPanel { Orientation = Orientation.Vertical, Spacing = 1,
                        Children =
                        {
                            new TextBlock { Text = sample.시료명, FontSize = 12, FontFamily = Font,
                                           Foreground = Brushes.WhiteSmoke },
                            new TextBlock { Text = sub, FontSize = 10, FontFamily = Font,
                                           Foreground = new SolidColorBrush(Color.Parse("#777799")) }
                        }
                    }
                }
            }
        };
    }


    // 업체 노드 하위에 시료 노드 로드
    private void LoadSampleNodes(TreeViewItem parent, string company)
    {
        parent.Items.Clear();

        var samples = TestReportService.GetSamplesByCompany(company);

        // 년-월별 그룹
        var byYearMonth = samples
            .GroupBy(s => s.채취일자.Length >= 7 ? s.채취일자[..7] : s.채취일자)
            .OrderByDescending(g => g.Key);

        foreach (var group in byYearMonth)
        {
            var monthNode = new TreeViewItem
            {
                Header = new TextBlock
                {
                    Text       = $"  📅 {group.Key}  ({group.Count()}건)",
                    FontSize   = 12,
                    FontFamily = Font,
                    Foreground = new SolidColorBrush(Color.Parse("#aaaacc")),
                }
            };

            foreach (var sample in group.OrderByDescending(s => s.채취일자))
                monthNode.Items.Add(MakeSampleNode(sample));

            parent.Items.Add(monthNode);
        }

        Log($"[{company}] {samples.Count}건 시료 로드");
    }

    // =========================================================================
    // 트리 선택 이벤트
    // =========================================================================
    public void SortToggle_Changed(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        _sortByDate = tglSortMode.IsChecked == true;
        LoadData();
    }

    public void ReportTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        SampleRequest? sample = null;

        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is SampleRequest s)
                sample = s;
        }

        if (sample == null) return;

        _selectedSample    = sample;
        _selectedRow       = null;
        _selectedRowBorder = null;

        // 결과 리스트 빌드
        BuildResultRows(sample);
        ResultListChanged?.Invoke(BuildListControl());
        EditPanelChanged?.Invoke(null);

        Log($"선택: {sample.약칭} / {sample.시료명} ({_resultRows.Count}항목)");
    }

    // =========================================================================
    // 결과 행 빌드 (NULL 컬럼 스킵)
    // =========================================================================
    private void BuildResultRows(SampleRequest sample)
    {
        _resultRows.Clear();

        foreach (var kv in sample.분석결과)
        {
            _meta.TryGetValue(kv.Key, out var meta);

            _resultRows.Add(new AnalysisResultRow
            {
                항목명         = kv.Key,
                결과값         = kv.Value,
                Original결과값 = kv.Value,
                DB컬럼명       = kv.Key,
                단위           = meta?.unit       ?? "",
                분석방법       = meta?.Method     ?? "",
                분석장비       = meta?.instrument ?? "",
                ES             = meta?.ES         ?? "",
                Category       = meta?.Category   ?? "",
            });
        }
    }

    // =========================================================================
    // 리스트 컨트롤 빌드 (ActivePageContent2)
    // =========================================================================
    private Control BuildListControl()
    {
        // 매번 새 StackPanel 생성 (visual parent 중복 방지)
        var listPanel = new StackPanel { Spacing = 0 };
        RefreshListPanel(listPanel);

        var outer = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };

        var infoPanel = BuildSampleInfoHeader();
        Grid.SetRow(infoPanel, 0);

        // 방류기준 토글 (최초 1회 생성, ContentControl 래퍼로 visual parent 중복 방지)
        if (_stdToggle == null)
        {
            _stdToggle = new ToggleSwitch
            {
                IsChecked = _showStandard,
                OnContent = "📋 방류기준 숨기기",
                OffContent = "📋 방류기준 표시",
                FontFamily = Font, FontSize = 11,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(8, 2),
            };
            _stdToggle.IsCheckedChanged += (_, _) =>
            {
                _showStandard = _stdToggle.IsChecked == true;
                Log($"[방류기준 토글] _showStandard={_showStandard}, 시료={_selectedSample?.시료명 ?? "null"}, 방류허용기준='{_selectedSample?.방류허용기준 ?? "null"}'");
                if (_selectedRow != null) EditPanelChanged?.Invoke(BuildEditPanel(_selectedRow));
                ResultListChanged?.Invoke(BuildListControl());
            };
        }
        else { _stdToggle.IsChecked = _showStandard; }
        if (_stdToggle.Parent is ContentControl oldC) oldC.Content = null;
        _stdToggleContainer = new ContentControl { Content = _stdToggle };
        Grid.SetRow(_stdToggleContainer, 1);

        var header = BuildResultListHeader();
        Grid.SetRow(header, 2);

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Content                       = listPanel
        };
        Grid.SetRow(scroll, 3);

        outer.Children.Add(infoPanel);
        outer.Children.Add(_stdToggleContainer);
        outer.Children.Add(header);
        outer.Children.Add(scroll);
        return outer;
    }

    // 시료 정보 헤더 패널
    private Border BuildSampleInfoHeader()
    {
        if (_selectedSample == null)
        {
            return new Border
            {
                Background = new SolidColorBrush(Color.Parse("#1e2030")),
                Padding    = new Thickness(12, 8),
                Child      = new TextBlock
                {
                    Text       = "왼쪽 트리에서 시료를 선택하세요",
                    FontFamily = Font, FontSize = 12,
                    Foreground = new SolidColorBrush(Color.Parse("#555566")),
                }
            };
        }

        var s = _selectedSample;
        var fields = new (string L, string V)[]
        {
            ("채취일자", s.채취일자), ("채취시간", s.채취시간), ("시료명", s.시료명),
            ("입회자",   s.입회자),   ("채취자",   $"{s.시료채취자1} {s.시료채취자2}".Trim()),
            ("방류기준", s.방류허용기준), ("정도보증", s.정도보증),
            ("분석종료", s.분석종료일),
        };

        var wrap = new WrapPanel { Orientation = Orientation.Horizontal };
        foreach (var (l, v) in fields)
        {
            if (string.IsNullOrEmpty(v)) continue;
            wrap.Children.Add(new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing     = 4,
                Margin      = new Thickness(0, 0, 16, 4),
                Children    =
                {
                    new TextBlock { Text = l + ":", FontFamily = Font, FontSize = 11,
                                   Foreground = new SolidColorBrush(Color.Parse("#888899")) },
                    new TextBlock { Text = v,       FontFamily = Font, FontSize = 11,
                                   Foreground = Brushes.WhiteSmoke }
                }
            });
        }

        return new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1e2030")),
            Padding    = new Thickness(12, 8),
            Child      = wrap
        };
    }

    // 결과 리스트 헤더
    private Border BuildResultListHeader() => new Border
    {
        Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
        Padding    = new Thickness(8, 5),
        Child      = new Grid
        {
            ColumnDefinitions = ResultColDefs,
            Children =
            {
                HeaderCell("No",     0), HeaderCell("구분",    1),
                HeaderCell("항목명", 2), HeaderCell("결과값",  3),
                HeaderCell("단위",   4), HeaderCell("분석방법",5),
                HeaderCell("ES",     6), HeaderCell("비고",    7),
            }
        }
    };


    // 결과 행 1줄
    private void RefreshListPanel(StackPanel listPanel)
    {
        listPanel.Children.Clear();
        _selectedRow = null; _selectedRowBorder = null;
        if (_resultRows.Count == 0)
        {
            listPanel.Children.Add(new TextBlock { Text = "분석 결과 없음",
                FontFamily = Font, FontSize = 12, Margin = new Thickness(12, 20),
                Foreground = new SolidColorBrush(Color.Parse("#555555")),
                HorizontalAlignment = HorizontalAlignment.Center });
            return;
        }
        bool odd = false; int no = 1;
        foreach (var row in _resultRows) { listPanel.Children.Add(BuildResultRow(row, no++, odd)); odd = !odd; }
    }

    private Border BuildResultRow(AnalysisResultRow row, int no, bool odd)
    {
        var bgColor = odd ? Color.Parse("#2e2e3a") : Color.Parse("#252530");
        var bg      = new SolidColorBrush(bgColor);

        // 결과값 색: 숫자면 흰색, 비었으면 회색
        IBrush valColor = string.IsNullOrEmpty(row.결과값)
            ? new SolidColorBrush(Color.Parse("#555555"))
            : (IBrush)Brushes.WhiteSmoke;

        var rowBorder = new Border
        {
            Background = bg,
            Padding    = new Thickness(8, 4),
            Cursor     = new Cursor(StandardCursorType.Hand),
            Child      = new Grid
            {
                ColumnDefinitions = ResultColDefs,
                Children =
                {
                    DataCell($"{no}",       0, Brushes.Gray),
                    DataCell(row.Category,  1),
                    DataCell(row.항목명,    2, Brushes.WhiteSmoke, bold: true),
                    DataCell(row.결과값,    3, (IBrush)valColor, bold: true),
                    DataCell(row.단위,      4),
                    DataCell(row.분석방법,  5),
                    DataCell(row.ES,        6, new SolidColorBrush(Color.Parse("#8888cc"))),
                    DataCell(GetStdDisplay(row), 7, new SolidColorBrush(Color.Parse("#ffaa44"))),
                }
            }
        };

        // 선택 이벤트 (Tunnel 방식)
        rowBorder.AddHandler(
            InputElement.PointerPressedEvent,
            (object? s, PointerPressedEventArgs ev) => SelectResultRow(rowBorder, row, bg),
            RoutingStrategies.Tunnel);

        rowBorder.PointerEntered += (_, _) =>
        {
            if (rowBorder != _selectedRowBorder)
                rowBorder.Background = new SolidColorBrush(Color.Parse("#3a3a50"));
        };
        rowBorder.PointerExited += (_, _) =>
        {
            if (rowBorder != _selectedRowBorder)
                rowBorder.Background = bg;
        };

        return rowBorder;
    }

    // =========================================================================
    // 결과 행 선택 → 수정 폼 (ActivePageContent3)
    // =========================================================================
    private void SelectResultRow(Border border, AnalysisResultRow row, SolidColorBrush normalBg)
    {
        if (_selectedRowBorder != null && _selectedRowBorder.Tag is SolidColorBrush prev)
            _selectedRowBorder.Background = prev;

        border.Tag            = normalBg;
        border.Background     = new SolidColorBrush(Color.Parse("#4a4a70"));
        _selectedRowBorder    = border;
        _selectedRow          = row;

        EditPanelChanged?.Invoke(BuildEditPanel(row));
        Log($"행 선택: {row.항목명} = {row.결과값}");
    }

    // =========================================================================
    // 수정 패널 (ActivePageContent3)
    // =========================================================================
    private string GetStdDisplay(AnalysisResultRow row)
    {
        if (!_showStandard) return "";
        if (_selectedSample == null) { Log("[방류기준] _selectedSample null"); return ""; }
        var col = _selectedSample.방류허용기준;
        if (string.IsNullOrEmpty(col)) { Log($"[방류기준] 방류허용기준 비어있음 - 시료: {_selectedSample.시료명}"); return ""; }
        var result = TestReportService.GetStandardValue(row.항목명, col);
        Log($"[방류기준] {row.항목명} / 기준컬럼={col} → '{result}'");
        return result;
    }

    private Control BuildEditPanel(AnalysisResultRow row)
    {
        var standardVal = GetStdDisplay(row);
        var valueTb = new TextBox
        {
            Text = row.결과값, Width = 120, FontFamily = Font, FontSize = 12,
            Background = new SolidColorBrush(Color.Parse("#3a3a4a")), Foreground = Brushes.WhiteSmoke,
            BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Color.Parse("#5555aa")),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 4),
        };
        var statusTb = new TextBlock { FontFamily = Font, FontSize = 11,
            Foreground = new SolidColorBrush(Color.Parse("#888888")), VerticalAlignment = VerticalAlignment.Center };
        var saveBtn   = MakeBtn("💾 저장", "#2a5a2a");
        var cancelBtn = MakeBtn("✖ 취소", "#444444");

        saveBtn.Click += (_, _) =>
        {
            var cur = _selectedSample; if (cur == null) return;
            var newVal = valueTb.Text?.Trim() ?? "";
            bool ok = TestReportService.UpdateResult(cur.Id, row.DB컬럼명, newVal);
            if (ok)
            {
                row.결과값 = row.Original결과값 = newVal;
                cur.분석결과[row.DB컬럼명] = newVal;
                if (_selectedRow?.DB컬럼명 == row.DB컬럼명) { _selectedRow.결과값 = newVal; _selectedRow.Original결과값 = newVal; }
                statusTb.Text = "✅ 저장 완료";
                statusTb.Foreground = new SolidColorBrush(Color.Parse("#44aa44"));
                ResultListChanged?.Invoke(BuildListControl());
            }
            else { statusTb.Text = "❌ 저장 실패"; statusTb.Foreground = new SolidColorBrush(Color.Parse("#cc4444")); }
        };
        cancelBtn.Click += (_, _) => { valueTb.Text = row.Original결과값; statusTb.Text = ""; };

        return new Border
        {
            Padding = new Thickness(10),
            Child = new StackPanel { Spacing = 8, Children =
            {
                new TextBlock { Text = "✏️  결과값 수정", FontFamily = Font, FontSize = 13,
                               FontWeight = FontWeight.SemiBold, Foreground = Brushes.WhiteSmoke },
                new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#555")) },
                new WrapPanel { Orientation = Orientation.Horizontal, Children =
                {
                    InfoCell("항목명",   row.항목명),
                    InfoCell("구분",     row.Category),
                    InfoCell("단위",     row.단위),
                    InfoCell("분석방법", row.분석방법),
                    InfoCell("ES",       row.ES),
                    string.IsNullOrEmpty(standardVal) ? new StackPanel() : InfoCell("방류기준", standardVal, highlight: true),
                    new StackPanel { Orientation = Orientation.Vertical, Spacing = 3, Children =
                    {
                        new TextBlock { Text = "결과값", FontFamily = Font, FontSize = 10,
                                       Foreground = new SolidColorBrush(Color.Parse("#aaaaaa")) },
                        new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6,
                            VerticalAlignment = VerticalAlignment.Center,
                            Children = { valueTb, saveBtn, cancelBtn, statusTb } }
                    }}
                }},
            }}
        };
    }

    private static StackPanel InfoCell(string label, string value, bool highlight = false)
    {
        if (string.IsNullOrEmpty(value)) return new StackPanel();
        return new StackPanel { Orientation = Orientation.Vertical, Spacing = 2, Margin = new Thickness(0,0,16,0),
            Children = {
                new TextBlock { Text = label, FontFamily = Font, FontSize = 10,
                               Foreground = new SolidColorBrush(Color.Parse("#888888")) },
                new TextBlock { Text = value, FontFamily = Font, FontSize = 12,
                               Foreground = highlight
                                   ? new SolidColorBrush(Color.Parse("#ffaa44"))
                                   : (IBrush)Brushes.LightGray }
            }
        };
    }

    private static Button MakeBtn(string text, string bg) => new Button
    {
        Content = text, FontFamily = Font, FontSize = 12,
        Background = new SolidColorBrush(Color.Parse(bg)), Foreground = Brushes.WhiteSmoke,
        BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        Padding = new Thickness(12, 5), VerticalAlignment = VerticalAlignment.Center,
    };


    // =========================================================================
    // ── 일괄 엑셀 출력 ────────────────────────────────────────────────────────
    public void BatchPrintExcel() => BatchPrint(toPdf: false);

    // ── 일괄 PDF 출력 ─────────────────────────────────────────────────────────
    public void BatchPrintPdf() => BatchPrint(toPdf: true);

    private void BatchPrint(bool toPdf)
    {
        if (_checkedSamples.Count == 0)
        {
            ShowToast("선택된 시료가 없습니다.\n시료 옆 체크박스를 선택해 주세요.");
            return;
        }
        int ok = 0, fail = 0;
        foreach (var sample in _checkedSamples.ToList())
        {
            try
            {
                var rows = BuildResultRowsForSample(sample);
                TestReportPrintService.FillAndSave(
                    sample: sample, rows: rows, meta: _meta,
                    toPdf: toPdf, openAfter: false);
                ok++;
                Log($"[일괄] ✅ {sample.시료명}");
            }
            catch (Exception ex) { fail++; Log($"[일괄] ❌ {sample.시료명}: {ex.Message}"); }
        }
        ShowToast($"✅ 일괄 {(toPdf ? "PDF" : "엑셀")} 완료\n성공: {ok}개  실패: {fail}개");
        var dir = System.IO.Path.GetFullPath(System.IO.Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "Data", "Reports"));
        TestReportPrintService.OpenFile(dir);
    }

    // ── 전체 선택 / 해제 ─────────────────────────────────────────────────────
    public void SelectAllSamples(bool select)
    {
        void Walk(System.Collections.IEnumerable items)
        {
            foreach (TreeViewItem node in items.OfType<TreeViewItem>())
            {
                if (node.Tag is SampleRequest s)
                {
                    if (select) _checkedSamples.Add(s);
                    else        _checkedSamples.Remove(s);
                    if (node.Header is StackPanel sp)
                        foreach (var cb in sp.Children.OfType<CheckBox>())
                            cb.IsChecked = select;
                }
                Walk(node.Items);
            }
        }
        Walk(ReportTreeView.Items);
        Log($"[선택] 전체 {(select ? "선택" : "해제")}: {_checkedSamples.Count}개");
        ShowToast($"{(select ? "✅ 전체 선택" : "☐ 전체 해제")} ({_checkedSamples.Count}개)");
    }

    // ── 시료의 분석결과 행 빌드 ───────────────────────────────────────────────
    private List<AnalysisResultRow> BuildResultRowsForSample(SampleRequest sample)
    {
        var rows = new List<AnalysisResultRow>();
        foreach (var kv in sample.분석결과)
        {
            _meta.TryGetValue(kv.Key, out var item);
            rows.Add(new AnalysisResultRow
            {
                항목명         = kv.Key,
                결과값         = kv.Value,
                단위           = item?.unit       ?? "",
                분석방법       = item?.Method     ?? "",
                분석장비       = item?.instrument ?? "",
                ES             = item?.ES         ?? "",
                Category       = item?.Category   ?? "",
                Original결과값 = kv.Value,
                DB컬럼명       = kv.Key,
            });
        }
        return rows;
    }

    /// BT1 — 출력
    public void Print() => PrintExcel();

    // ── 엑셀 출력 ─────────────────────────────────────────────────────────────
    public void PrintExcel()
    {
        if (_selectedSample == null) { ShowToast("시료를 먼저 선택하세요."); return; }
        if (!TestReportPrintService.TemplateExists())
        {
            ShowToast($"템플릿 없음\nData\\Templates\\ 폴더에 템플릿 파일을 넣어주세요.");
            TestReportPrintService.OpenTemplateFolder();
            return;
        }
        try
        {
            var rows = BuildResultRowsForSample(_selectedSample);
            var path = TestReportPrintService.FillAndSave(
                sample: _selectedSample, rows: rows, meta: _meta,
                toPdf: false, openAfter: true);
            Log($"✅ 엑셀 저장: {path}");
            ShowToast($"✅ 저장 완료\n{System.IO.Path.GetFileName(path)}");
        }
        catch (Exception ex) { Log($"❌ 엑셀 오류: {ex.Message}"); ShowToast($"❌ 오류: {ex.Message}"); }
    }

    // ── PDF 저장 ──────────────────────────────────────────────────────────────
    public void PrintPdf()
    {
        if (_selectedSample == null) { ShowToast("시료를 먼저 선택하세요."); return; }
        if (!TestReportPrintService.TemplateExists())
        {
            ShowToast("템플릿 없음\nData\\Templates\\ 폴더에 템플릿 파일을 넣어주세요.");
            return;
        }
        try
        {
            var rows = BuildResultRowsForSample(_selectedSample);
            var path = TestReportPrintService.FillAndSave(
                sample: _selectedSample, rows: rows, meta: _meta,
                toPdf: true, openAfter: true);
            Log($"✅ PDF 저장: {path}");
            ShowToast($"✅ 저장 완료\n{System.IO.Path.GetFileName(path)}");
        }
        catch (Exception ex) { Log($"❌ PDF 오류: {ex.Message}"); ShowToast($"❌ 오류: {ex.Message}"); }
    }

    /// BT2 — CSV 저장
    public void SaveCsv()
    {
        if (_selectedSample == null || _resultRows.Count == 0)
        {
            Log("저장: 선택 없음"); return;
        }

        try
        {
            var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
            var dir  = Path.Combine(root, "Data", "Export");
            Directory.CreateDirectory(dir);
            var file = Path.Combine(dir,
                $"시험성적서_{_selectedSample.약칭}_{_selectedSample.채취일자}_{DateTime.Now:HHmmss}.csv");

            var lines = new List<string>
            {
                // 시료 정보
                $"채취일자,{_selectedSample.채취일자}",
                $"채취일자,{_selectedSample.채취일자}",
                $"시료명,{_selectedSample.시료명}",
                $"입회자,{_selectedSample.입회자}",
                $"채취자,{_selectedSample.시료채취자1} {_selectedSample.시료채취자2}",
                $"방류허용기준,{_selectedSample.방류허용기준}",
                $"정도보증,{_selectedSample.정도보증}",
                "",
                "No,구분,항목명,결과값,단위,분석방법,ES"
            };

            int no = 1;
            foreach (var r in _resultRows)
                lines.Add($"{no++},{r.Category},{r.항목명},{r.결과값},{r.단위},{r.분석방법},{r.ES}");

            File.WriteAllLines(file, lines, System.Text.Encoding.UTF8);

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = file, UseShellExecute = true
            });
            Log($"CSV 저장: {file}");
        }
        catch (Exception ex) { Log($"CSV 저장 오류: {ex.Message}"); }
    }

    /// BT3 — 삭제
    public async Task DeleteSampleAsync()
    {
        if (_selectedSample == null) { Log("삭제: 선택 없음"); return; }

        var confirmed = await ShowConfirmAsync(
            $"'{_selectedSample.시료명}' 시료를 삭제하시겠습니까?\n(분석결과 포함 전체 삭제)");
        if (!confirmed) return;

        bool ok = TestReportService.DeleteSample(_selectedSample.Id);
        Log(ok ? $"✅ 삭제: {_selectedSample.시료명}" : "❌ 삭제 실패");

        if (ok)
        {
            // 트리에서 해당 노드 제거
            RemoveSampleNodeFromTree(_selectedSample);
            _selectedSample = null;
            _resultRows.Clear();
            ResultListChanged?.Invoke(BuildListControl());
            EditPanelChanged?.Invoke(null);
        }
    }

    private void RemoveSampleNodeFromTree(SampleRequest sample)
    {
        foreach (TreeViewItem companyNode in ReportTreeView.Items.OfType<TreeViewItem>())
        {
            foreach (TreeViewItem monthNode in companyNode.Items.OfType<TreeViewItem>())
            {
                var toRemove = monthNode.Items.OfType<TreeViewItem>()
                                   .FirstOrDefault(n => n.Tag is SampleRequest s && s.Id == sample.Id);
                if (toRemove != null)
                {
                    monthNode.Items.Remove(toRemove);
                    if (monthNode.Items.Count == 0) companyNode.Items.Remove(monthNode);
                    return;
                }
            }
        }
    }

    // =========================================================================
    // UI 헬퍼
    // =========================================================================
    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    private static ColumnDefinitions ResultColDefs =>
        new ColumnDefinitions("36,70,160,90,60,140,80,*");

    private static TextBlock HeaderCell(string text, int col)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = 11, FontWeight = FontWeight.SemiBold,
            Foreground = new SolidColorBrush(Color.Parse("#aaaacc")),
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(tb, col); return tb;
    }

    private static TextBlock DataCell(string text, int col, IBrush? fg = null, bool bold = false)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = 12,
            FontWeight   = bold ? FontWeight.SemiBold : FontWeight.Normal,
            Foreground   = fg ?? new SolidColorBrush(Color.Parse("#cccccc")),
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(tb, col); return tb;
    }

    private static StackPanel InfoRow(string label, string value)
    {
        if (string.IsNullOrEmpty(value))
            return new StackPanel();

        return new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 8,
            Children    =
            {
                new TextBlock { Text = label + " :", Width = 70, FontFamily = Font, FontSize = 11,
                                Foreground = new SolidColorBrush(Color.Parse("#888888")),
                                VerticalAlignment = VerticalAlignment.Center },
                new TextBlock { Text = value, FontFamily = Font, FontSize = 11,
                                Foreground = Brushes.LightGray,
                                VerticalAlignment = VerticalAlignment.Center },
            }
        };
    }

    private void ShowToast(string msg)
    {
        Log($"[Toast] {msg}");
        // 간단한 팝업 (향후 Toast UI 교체 가능)
        var dlg = new Window
        {
            Title = "알림", Width = 320, Height = 120, CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.Parse("#2d2d2d")),
        };
        var btn = new Button { Content = "확인", Width = 80,
                               Background = new SolidColorBrush(Color.Parse("#444")),
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        btn.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 14,
            Children =
            {
                new TextBlock { Text = msg, FontFamily = Font, FontSize = 12,
                                Foreground = Brushes.WhiteSmoke,
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel { Orientation = Orientation.Horizontal,
                                 HorizontalAlignment = HorizontalAlignment.Right,
                                 Children = { btn } }
            }
        };
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) dlg.ShowDialog(owner);
        else dlg.Show();
    }

    private async Task<bool> ShowConfirmAsync(string msg)
    {
        var dlg = new Window
        {
            Title = "삭제 확인", Width = 340, Height = 160, CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.Parse("#2d2d2d")),
        };
        bool result = false;
        var yes = new Button { Content = "삭제", Width = 80,
                               Background = new SolidColorBrush(Color.Parse("#c0392b")),
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        var no  = new Button { Content = "취소", Width = 80,
                               Background = new SolidColorBrush(Color.Parse("#444")),
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        yes.Click += (_, _) => { result = true;  dlg.Close(); };
        no.Click  += (_, _) => { result = false; dlg.Close(); };
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 14,
            Children =
            {
                new TextBlock { Text = msg, FontFamily = Font, FontSize = 12,
                                Foreground = Brushes.WhiteSmoke,
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10,
                                 HorizontalAlignment = HorizontalAlignment.Right,
                                 Children = { yes, no } }
            }
        };
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();
        return result;
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [TestReport] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("TestReportDebug.log", line + Environment.NewLine); } catch { }
    }
}