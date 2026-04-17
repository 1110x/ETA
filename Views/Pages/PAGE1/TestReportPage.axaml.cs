using Avalonia;
using Avalonia.Animation;
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
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.PAGE1;

public partial class TestReportPage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    /// Show2 에 넣을 결과 리스트 컨트롤
    public event Action<Control?>? ResultListChanged;
    /// Show3 에 넣을 수정 폼 컨트롤
    public event Action<Control?>? EditPanelChanged;
    /// 현재 트리에서 선택된 시료 (DataToMeasurerWindow에서 사용)
    public SampleRequest? SelectedSample => _selectedSample;

    // ── 상태 ────────────────────────────────────────────────────────────────
    private SampleRequest?                    _selectedSample;
    private List<AnalysisResultRow>           _resultRows  = new();
    private Dictionary<string, AnalysisItem>  _meta        = new();

    // 리스트 패널

    private readonly TextBlock  _listTitle  = new()
    {
        FontSize   = AppTheme.FontMD,
        FontFamily = Font,
        Foreground = new SolidColorBrush(Color.Parse("#8888bb")),
        Margin     = new Thickness(10, 6),
    };

    // 정렬 모드 / 방류기준 토글
    private bool _sortByDate   = true;

    // 검색용 전체 시료 캐시 + 디바운스 타이머
    private List<SampleRequest> _allSamples = [];
    private readonly Avalonia.Threading.DispatcherTimer _searchTimer =
        new() { Interval = TimeSpan.FromMilliseconds(250) };
    private string _pendingSearchText = "";
    // 선택된 시료 (일괄 출력용)
    private readonly System.Collections.Generic.HashSet<SampleRequest> _checkedSamples = new();
    private bool _showImportPanel = false;
    private Control? _importActionPanel;
    private ToggleSwitch?   _stdToggle;
    private ContentControl? _stdToggleContainer;

    // Excel 불러오기 공유 상태
    private readonly List<(string A, string Q, string Y, string S, string Ex, string Nv, string Std, string Bg, string Fg)> _importRows = new();
    private readonly Dictionary<int, Dictionary<string, string>> _pendingByRow = new();
    private readonly List<(int RowId, string Analyte, string Existing, string NewVal)> _conflictRows = new();
    private List<string> _importFileNames = new();
    private string _importStatus      = "";
    private string _importStatusHex   = "#888888";
    private bool   _importHasPending  = false;
    private bool   _importHasConflict = false;

    // 첨부파일
    private StackPanel? _attachListPanel;
    private int         _attachSampleId = -1;

    // 자료TO측정인 패널
    private bool            _showMeasurerPanel = false;
    private List<MeasurerRow> _measurerRows    = new();

    /// <summary>결과제출-측정인 모드 — 시료 선택시 항상 측정인 형식으로 표시</summary>
    public bool IsMeasurerMode { get; set; }

    // 선택된 결과 행
    private AnalysisResultRow? _selectedRow;
    private Border?            _selectedRowBorder;

    // ── Show2 부분 애니메이션: 헤더 고정, 데이터 행만 전환 ──
    private Grid?                       _persistentGrid;
    private TransitioningContentControl? _listTCC;
    private Border?                      _infoHeaderBorder;
    private ProgressBar?                 _infoLoadProgress;
    private Dictionary<string, string>   _standardValueCache = new();
    private ToggleSwitch?                _dischargeStdToggle;
    private bool                         _showDischargeStd = true;

    public TestReportPage()
    {
        InitializeComponent();
        _meta = TestReportService.GetAnalyteMeta();
        _searchTimer.Tick += (_, _) =>
        {
            _searchTimer.Stop();
            ApplySearch(_pendingSearchText);
        };
    }

    // =========================================================================
    // 데이터 로드 (트리)
    // =========================================================================
    public void LoadData()
    {
        Log("LoadData() 시작");
        ReportTreeView.Items.Clear();
        _selectedSample    = null;
                _importRows.Clear();
        _pendingByRow.Clear();
        _conflictRows.Clear();
        _importFileNames.Clear();
        _importStatus      = "";
        _importStatusHex   = "#888888";
        _importHasPending  = false;
        _importHasConflict = false;
        _importActionPanel = null;
        _persistentGrid    = null; _listTCC = null; // 전체 재빌드
        ResultListChanged?.Invoke(BuildListControl());
        EditPanelChanged?.Invoke(null);

        try
        {
            var companies = TestReportService.GetCompanyList();
            Log($"업체 {companies.Count}개");
            if (companies.Count == 0)
            {
                ReportTreeView.Items.Add(new TreeViewItem {
                    Header = new TextBlock { Text = "⚠️ 데이터 없음", FontSize = AppTheme.FontMD, FontFamily = Font,
                                            Foreground = new SolidColorBrush(Color.Parse("#cc8800")) }
                });
                return;
            }
            // 검색용 전체 캐시 갱신
            _allSamples = companies
                .SelectMany(co => TestReportService.GetSamplesByCompany(co))
                .OrderByDescending(s => s.채취일자)
                .ToList();
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
                        new TextBlock { Text = "🏭", FontSize = AppTheme.FontXL, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = company, FontSize = AppTheme.FontLG, FontFamily = Font,
                                        Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center }
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
        var now          = DateTime.Today;
        var curMonthKey  = now.ToString("yyyy-MM");
        var prevMonthKey = now.AddMonths(-1).ToString("yyyy-MM");
        var curYearKey   = now.Year.ToString();
        var prevYearKey  = now.AddMonths(-1).Year.ToString();

        var all = companies.SelectMany(co => TestReportService.GetSamplesByCompany(co))
                           .OrderByDescending(s => s.채취일자).ToList();
        foreach (var yg in all.GroupBy(s => s.채취일자.Length >= 4 ? s.채취일자[..4] : s.채취일자)
                               .OrderByDescending(g => g.Key))
        {
            bool yearOpen = yg.Key == curYearKey || yg.Key == prevYearKey;
            var yn = new TreeViewItem { IsExpanded = yearOpen,
                Header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6,
                    Children = {
                        new TextBlock { Text = "📆", FontSize = AppTheme.FontXL, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = $"{yg.Key}년  ({yg.Count()}건)", FontSize = AppTheme.FontLG, FontFamily = Font,
                                        Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center }
                    }
                }
            };
            foreach (var mg in yg.GroupBy(s => s.채취일자.Length >= 7 ? s.채취일자[..7] : s.채취일자)
                                  .OrderByDescending(g => g.Key))
            {
                bool monthOpen = mg.Key == curMonthKey || mg.Key == prevMonthKey;
                var mn = new TreeViewItem { IsExpanded = monthOpen,
                    Header = new TextBlock { Text = $"  📅 {mg.Key}  ({mg.Count()}건)",
                        FontSize = AppTheme.FontMD, FontFamily = Font, Foreground = AppTheme.FgInfo }
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
        bool incomplete = sample.분석결과.Values.Any(v =>
            string.Equals(v, "O", StringComparison.OrdinalIgnoreCase));
        string iconColor = incomplete ? "#ee4444" : "#44cc44";

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

        string mmdd = "";
        if (sample.채취일자.Length >= 10)
            mmdd = sample.채취일자.Substring(5, 5).Replace("-", "/");
        else if (sample.채취일자.Length >= 7)
            mmdd = sample.채취일자.Substring(5);

        string icon = incomplete ? "🥑" : "🍬";
        string labelText = string.IsNullOrEmpty(mmdd)
            ? sample.시료명
            : $"{sample.시료명}  ({mmdd})";

        var (badgeBg, badgeFg) = BadgeColorHelper.GetBadgeColor(sample.약칭);

        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,Auto,Auto,*"), Margin = new Thickness(2, 2) };
        row.Children.Add(new ContentControl { Content = chk, [Grid.ColumnProperty] = 0 });
        row.Children.Add(new TextBlock
        {
            Text = icon, FontSize = AppTheme.FontMD,
            Foreground = new SolidColorBrush(Color.Parse(iconColor)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(2, 0, 4, 0),
            [Grid.ColumnProperty] = 1,
        });
        row.Children.Add(new Border
        {
            Background = Brush.Parse(badgeBg), CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 2,
            Child = new TextBlock { Text = sample.약칭, FontSize = AppTheme.FontXS, FontFamily = Font,
                                    Foreground = Brush.Parse(badgeFg) },
        });
        row.Children.Add(new TextBlock
        {
            Text = labelText, FontSize = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgPrimary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 3,
        });

        return new TreeViewItem { Tag = sample, Header = row };
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
                    FontSize   = AppTheme.FontMD,
                    FontFamily = Font,
                    Foreground = AppTheme.FgInfo,
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
    // ── 검색 ─────────────────────────────────────────────────────────────
    public void TxbSearch_TextChanged(object? sender, TextChangedEventArgs e)
    {
        _pendingSearchText = sender is TextBox tb ? tb.Text ?? "" : "";
        _searchTimer.Stop();
        _searchTimer.Start();
    }

    private void ApplySearch(string keyword)
    {
        if (string.IsNullOrWhiteSpace(keyword))
        {
            LoadData();
            return;
        }

        var kw = keyword.Trim();
        var matched = _allSamples
            .Where(s =>
                (s.시료명   ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                (s.약칭     ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase) ||
                (s.채취일자 ?? "").Contains(kw, StringComparison.OrdinalIgnoreCase))
            .ToList();

        ReportTreeView.Items.Clear();
        foreach (var s in matched)
            ReportTreeView.Items.Add(MakeSampleNode(s, showCompany: true));
    }

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

        // 방류기준표 캐시 로드 (각 행마다 DB 조회 대신 1번만)
        _standardValueCache = !string.IsNullOrEmpty(sample.방류허용기준)
            && sample.방류허용기준 != "기준없음" && sample.방류허용기준 != "해당없음"
            ? TestReportService.GetStandardValueMap(sample.방류허용기준)
            : new Dictionary<string, string>();

        // 결과 리스트 빌드
        BuildResultRows(sample);

        // 측정인 모드 — 항상 측정인 형식으로 표시
        if (IsMeasurerMode)
        {
            _showMeasurerPanel = true;
            _measurerRows = DataToMeasurerService.BuildRows(sample);
        }

        ResultListChanged?.Invoke(BuildListControl());

        // 첨부파일 목록 갱신
        RefreshAttachmentList(sample.Id);

        Log($"선택: {sample.약칭} / {sample.시료명} ({_resultRows.Count}항목)");
    }

    private void SetLoadProgress(double value)
    {
        if (_infoLoadProgress == null) return;
        _infoLoadProgress.Value = value;
        // 100% → 파랑, 미만 → 주황
        _infoLoadProgress.Foreground = value >= 100
            ? new SolidColorBrush(Color.Parse("#2196F3"))  // 파랑
            : new SolidColorBrush(Color.Parse("#FF9800")); // 주황
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
    // 리스트 컨트롤 빌드 (Show2)
    // =========================================================================
    private Control BuildListControl()
    {
        if (_showMeasurerPanel) { _persistentGrid = null; _listTCC = null; return BuildMeasurerListControl(); }

        if (_stdToggle == null)
        {
            _stdToggle = new ToggleSwitch
            {
                IsChecked  = _showImportPanel,
                OnContent  = "📎 Excel 불러오기",
                OffContent = "📋 시험성적서",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(8, 2),
            };
            _stdToggle.IsCheckedChanged += (_, _) =>
            {
                _showImportPanel = _stdToggle.IsChecked == true;
                if (_showImportPanel)
                {
                    if (_importActionPanel == null) _importActionPanel = BuildImportActionPanel();
                    EditPanelChanged?.Invoke(_importActionPanel);
                }
                else
                {
                    EditPanelChanged?.Invoke(null);
                }
                _persistentGrid = null; _listTCC = null; // 모드 변경 시 전체 재빌드
                ResultListChanged?.Invoke(BuildListControl());
            };
        }
        else { _stdToggle.IsChecked = _showImportPanel; }

        if (_showImportPanel)
        {
            _persistentGrid = null; _listTCC = null;

            if (_stdToggle.Parent is ContentControl oldC2) oldC2.Content = null;
            _stdToggleContainer = new ContentControl { Content = _stdToggle };

            // Content3 = 액션 패널 (파일 선택, 저장 등)
            if (_importActionPanel == null) _importActionPanel = BuildImportActionPanel();
            EditPanelChanged?.Invoke(_importActionPanel);

            // Content2 = 토글 + 결과 테이블
            var tablePanel = new StackPanel { Spacing = 1 };
            if (_importRows.Count > 0)
            {
                tablePanel.Children.Add(MakeImportTableRow(
                    "분석항목", "견적번호", "약칭", "시료명",
                    "이전결과", "새결과값", "방류기준",
                    true, "#2a2a3a", "#aaaacc"));
                foreach (var row in _importRows)
                    tablePanel.Children.Add(MakeImportTableRow(
                        row.A, row.Q, row.Y, row.S, row.Ex, row.Nv, row.Std,
                        false, row.Bg, row.Fg));
            }
            else
            {
                tablePanel.Children.Add(new TextBlock
                {
                    Text = "오른쪽에서 Excel 파일을 선택하면 결과가 여기에 표시됩니다.",
                    FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppTheme.FgDimmed,
                    Margin     = new Thickness(12, 20),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                });
            }
            var tableScroll = new ScrollViewer
            {
                Content = tablePanel,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            };
            var importOuter = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
            Grid.SetRow(_stdToggleContainer, 0);
            Grid.SetRow(tableScroll, 1);
            importOuter.Children.Add(_stdToggleContainer);
            importOuter.Children.Add(tableScroll);
            return importOuter;
        }

        // ── 결과 리스트 모드 ──
        // 영속 그리드 존재 시 헤더를 먼저 갱신해 _infoLoadProgress를 최신으로 만든 뒤 리스트 빌드
        if (_persistentGrid != null && _listTCC != null)
            UpdateSampleInfoHeader();

        var listPanel = new StackPanel { Spacing = 0 };
        RefreshListPanel(listPanel);

        var listScroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Content                       = listPanel
        };

        // ── 영속 그리드 존재 시: 헤더 고정, 데이터 행만 CrossFade 전환 ──
        if (_persistentGrid != null && _listTCC != null)
        {
            _listTCC.Content = listScroll;
            return _persistentGrid;
        }

        // ── 최초 빌드: 영속 구조 생성 ──
        // _stdToggle이 이전 Parent(ContentControl 또는 Panel)에 붙어있으면 제거
        if (_stdToggle.Parent is ContentControl oldC) oldC.Content = null;
        else if (_stdToggle.Parent is Panel oldP) oldP.Children.Remove(_stdToggle);
        var togglePanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 20,
            VerticalAlignment = VerticalAlignment.Center,
        };
        togglePanel.Children.Add(_stdToggle);
        togglePanel.Children.Add(new TextBlock
        {
            Text = "방류기준",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(12, 0, 4, 0),
        });
        togglePanel.Children.Add(CreateDischargeToggle());

        // ── 출력 타입 버튼 3개 ──────────────────────────────
        Button MkBtn(string text, string bg, string fg, EventHandler<RoutedEventArgs> onClick)
        {
            var b = new Button
            {
                Content = text,
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Background = new SolidColorBrush(Color.Parse(bg)),
                Foreground = new SolidColorBrush(Color.Parse(fg)),
                BorderThickness = new Thickness(0),
                Margin = new Thickness(0, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(10, 4),
                CornerRadius = new CornerRadius(4),
            };
            b.Click += onClick;
            return b;
        }

        var btnTR   = MkBtn("📄 시험성적서",   "#1a3a1a", "#7cd87c", OnPrintTestReport);
        var btnWMR  = MkBtn("📋 수질측정기록부", "#1a2a3a", "#7aaae8", OnPrintWaterRecord);
        var btnBoth = MkBtn("📄+📋 둘다",        "#2a2a1a", "#e8d87a", OnPrintBoth);
        btnTR.Margin   = new Thickness(20, 0, 0, 0);
        btnWMR.Margin  = new Thickness(4, 0, 0, 0);
        btnBoth.Margin = new Thickness(4, 0, 0, 0);
        togglePanel.Children.Add(btnTR);
        togglePanel.Children.Add(btnWMR);
        togglePanel.Children.Add(btnBoth);

        // GS 발송양식 버튼
        var btnGsExport = MkBtn("📤 GS 발송양식", "#2a1a3a", "#cc99ee", OnGsExportClick);
        btnGsExport.Margin = new Thickness(16, 0, 0, 0);
        togglePanel.Children.Add(btnGsExport);

        _stdToggleContainer = new ContentControl { Content = togglePanel };

        _infoHeaderBorder = BuildSampleInfoHeader();
        var resultHeader  = BuildResultListHeader();

        _listTCC = new TransitioningContentControl
        {
            PageTransition = new CrossFade(TimeSpan.FromMilliseconds(200)),
            Content        = listScroll
        };

        _persistentGrid = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };
        Grid.SetRow(_infoHeaderBorder,  0);
        Grid.SetRow(_stdToggleContainer, 1);
        Grid.SetRow(resultHeader,        2);
        Grid.SetRow(_listTCC,            3);

        _persistentGrid.Children.Add(_infoHeaderBorder);
        _persistentGrid.Children.Add(_stdToggleContainer);
        _persistentGrid.Children.Add(resultHeader);
        _persistentGrid.Children.Add(_listTCC);

        return _persistentGrid;
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
                    FontFamily = Font, FontSize = AppTheme.FontMD,
                    Foreground = AppTheme.FgDimmed,
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
                    new TextBlock { Text = l + ":", FontFamily = Font, FontSize = AppTheme.FontBase,
                                   Foreground = AppTheme.FgMuted },
                    new TextBlock { Text = v,       FontFamily = Font, FontSize = AppTheme.FontBase,
                                   Foreground = AppRes("AppFg") }
                }
            });
        }

        _infoLoadProgress = new ProgressBar
        {
            Minimum = 0, Maximum = 100, Value = 0,
            Height = 3, Margin = new Thickness(0, 6, 0, 0),
            IsIndeterminate = false,
            Foreground = new SolidColorBrush(Color.Parse("#FF9800")),  // 주황
        };

        var outerPanel = new StackPanel { Spacing = 0 };
        outerPanel.Children.Add(wrap);
        outerPanel.Children.Add(_infoLoadProgress);

        return new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1e2030")),
            Padding    = new Thickness(12, 8),
            Child      = outerPanel
        };
    }

    /// <summary>영속 그리드의 시료 정보 헤더를 현재 선택된 시료로 갱신 (in-place)</summary>
    private void UpdateSampleInfoHeader()
    {
        if (_infoHeaderBorder == null) return;

        if (_selectedSample == null)
        {
            _infoHeaderBorder.Child = new TextBlock
            {
                Text       = "왼쪽 트리에서 시료를 선택하세요",
                FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppTheme.FgDimmed,
            };
            return;
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
                    new TextBlock { Text = l + ":", FontFamily = Font, FontSize = AppTheme.FontBase,
                                   Foreground = AppTheme.FgMuted },
                    new TextBlock { Text = v,       FontFamily = Font, FontSize = AppTheme.FontBase,
                                   Foreground = AppRes("AppFg") }
                }
            });
        }

        _infoLoadProgress = new ProgressBar
        {
            Minimum = 0, Maximum = 100, Value = 0,
            Height = 3, Margin = new Thickness(0, 6, 0, 0),
            IsIndeterminate = false,
            Foreground = new SolidColorBrush(Color.Parse("#FF9800")),  // 주황
        };

        var outerPanel = new StackPanel { Spacing = 0 };
        outerPanel.Children.Add(wrap);
        outerPanel.Children.Add(_infoLoadProgress);

        _infoHeaderBorder.Child = outerPanel;
    }

    /// <summary>📄 시험성적서 출력</summary>
    private void OnPrintTestReport(object? sender, RoutedEventArgs e)
    {
        OpenPrintWindow();
    }

    /// <summary>📋 수질측정기록부 Excel 생성</summary>
    private async void OnPrintWaterRecord(object? sender, RoutedEventArgs e)
    {
        if (_selectedSample == null) { ShowToast("시료를 먼저 선택하세요."); return; }
        if (!WaterMeasurementRecordService.TemplatesExist())
        {
            ShowToast("수질측정기록부 템플릿 없음\nData/Templates/수질측정기록부1(2)_template.xlsx 필요");
            return;
        }
        try
        {
            var rows = BuildResultRowsForSample(_selectedSample);
            var path = await Task.Run(() =>
                WaterMeasurementRecordService.Generate(_selectedSample, rows, _meta));
            Log($"[수질측정기록부] 생성: {path}");
            ShowToast($"✅ 저장 완료: {System.IO.Path.GetFileName(path)}");
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                { FileName = path, UseShellExecute = true });
            }
            catch { }
        }
        catch (Exception ex)
        {
            Log($"[수질측정기록부] 오류: {ex.Message}");
            ShowToast($"❌ 오류: {ex.Message}");
        }
    }

    /// <summary>📄+📋 둘다 출력</summary>
    private async void OnPrintBoth(object? sender, RoutedEventArgs e)
    {
        OnPrintWaterRecord(sender, e);
        await Task.Delay(500);
        OpenPrintWindow();
    }

    /// <summary>GS 발송양식 버튼 클릭</summary>
    private async void OnGsExportClick(object? sender, RoutedEventArgs e)
    {
        if (_selectedSample == null)
        {
            await ShowInfoToast("시료를 먼저 선택하세요");
            return;
        }
        // 지에스칼텍스 업체 확인
        var company = _selectedSample.의뢰사업장 ?? "";
        if (!company.Contains("지에스") && !company.Contains("GS") && !company.Contains("칼텍스"))
        {
            await ShowInfoToast("지에스칼텍스 시료만 GS 발송양식을 생성할 수 있습니다.");
            return;
        }

        try
        {
            var path = await Task.Run(() =>
                GsReportService.GenerateReport(new List<SampleRequest> { _selectedSample }, _meta));
            Log($"[GS] 발송양식 생성 완료: {path}");
            await ShowInfoToast($"GS 발송양식 생성: {System.IO.Path.GetFileName(path)}");

            // 폴더에서 열기
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true,
                });
            }
            catch { }
        }
        catch (Exception ex)
        {
            Log($"[GS] 발송양식 생성 오류: {ex.Message}");
            await ShowInfoToast($"오류: {ex.Message}");
        }
    }

    private async Task ShowInfoToast(string msg)
    {
        ShowToast(msg);
        await Task.CompletedTask;
    }

    /// <summary>방류기준 표시 ON/OFF 토글 스위치 생성</summary>
    private ToggleSwitch CreateDischargeToggle()
    {
        _dischargeStdToggle = new ToggleSwitch
        {
            IsChecked     = _showDischargeStd,
            OnContent     = "",
            OffContent    = "",
            FontSize      = AppTheme.FontXS,
            Margin        = new Thickness(8, -6, 0, -6),
            VerticalAlignment = VerticalAlignment.Center,
        };
        _dischargeStdToggle.IsCheckedChanged += (_, _) =>
        {
            _showDischargeStd = _dischargeStdToggle.IsChecked == true;
            // 결과 리스트 재렌더링
            ResultListChanged?.Invoke(BuildListControl());
        };
        return _dischargeStdToggle;
    }

    // 결과 리스트 헤더
    private Border BuildResultListHeader() => new Border
    {
        Background = AppTheme.BorderSubtle,
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
    private async void RefreshListPanel(StackPanel listPanel)
    {
        listPanel.Children.Clear();
        _selectedRow = null; _selectedRowBorder = null;
        if (_infoLoadProgress != null) _infoLoadProgress.Value = 0;

        // 스냅샷: async 루프 중 다른 시료 선택으로 _resultRows 변경돼도 안전
        var snapshot = _resultRows.ToList();
        var sampleAtStart = _selectedSample;

        if (snapshot.Count == 0)
        {
            listPanel.Children.Add(new TextBlock { Text = "분석 결과 없음",
                FontFamily = Font, FontSize = AppTheme.FontMD, Margin = new Thickness(12, 20),
                Foreground = AppTheme.FgDimmed,
                HorizontalAlignment = HorizontalAlignment.Center });
            SetLoadProgress(100);
            return;
        }
        bool odd = false; int no = 1;
        int total = snapshot.Count;
        foreach (var row in snapshot)
        {
            // 시료가 바뀌면 이 루프는 취소 (새로운 RefreshListPanel이 실행 중)
            if (_selectedSample != sampleAtStart) return;

            listPanel.Children.Add(BuildResultRow(row, no, odd));
            odd = !odd;
            SetLoadProgress((double)no / total * 100);
            no++;
            // UI 스레드에 렌더링 기회 양보 (한 줄씩 표시되게)
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(
                () => { }, Avalonia.Threading.DispatcherPriority.Background);
        }
        SetLoadProgress(100);
    }

    private Border BuildResultRow(AnalysisResultRow row, int no, bool odd)
    {
        var bgColor = odd ? Color.Parse("#2e2e3a") : Color.Parse("#252530");
        var bg      = new SolidColorBrush(bgColor);

        // 결과값 색: 기준 초과 시 빨간색, 아니면 흰색, 비었으면 회색
        var stdStr  = GetStdDisplay(row);
        IBrush valColor = ExceedsStandard(row.결과값, stdStr)
            ? AppTheme.FgDanger
            : string.IsNullOrEmpty(row.결과값)
                ? AppTheme.BorderDefault
                : (IBrush)AppRes("AppFg");

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
                    DataCell(row.항목명,    2, AppRes("AppFg"), bold: true),
                    DataCell(row.결과값,    3, (IBrush)valColor, bold: true),
                    DataCell(row.단위,      4),
                    DataCell(row.분석방법,  5),
                    DataCell(row.ES,        6, new SolidColorBrush(Color.Parse("#8888cc"))),
                    DataCell(stdStr,        7, AppTheme.FgWarn),
                }
            }
        };

        // 우클릭 복사 메뉴
        var copyItem = new MenuItem { Header = "📋 행 복사" };
        copyItem.Click += async (_, _) =>
        {
            var clip = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clip != null)
                await clip.SetTextAsync($"{row.항목명}\t{row.결과값}\t{row.단위}\t{row.분석방법}");
        };
        var copyValItem = new MenuItem { Header = "📋 결과값 복사" };
        copyValItem.Click += async (_, _) =>
        {
            var clip = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clip != null)
                await clip.SetTextAsync(row.결과값 ?? "");
        };
        rowBorder.ContextMenu = new ContextMenu { Items = { copyItem, copyValItem } };

        // 선택 이벤트 (Tunnel 방식)
        rowBorder.AddHandler(
            InputElement.PointerPressedEvent,
            (object? s, PointerPressedEventArgs ev) => SelectResultRow(rowBorder, row, bg),
            RoutingStrategies.Tunnel);

        rowBorder.PointerEntered += (_, _) =>
        {
            if (rowBorder != _selectedRowBorder)
                rowBorder.Background = AppTheme.BorderMuted;
        };
        rowBorder.PointerExited += (_, _) =>
        {
            if (rowBorder != _selectedRowBorder)
                rowBorder.Background = bg;
        };

        return rowBorder;
    }

    // =========================================================================
    // 결과 행 선택 → 수정 폼 (Show3)
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
    // Excel 결과 불러오기 패널 (Show3 — 항상 활성)
    // =========================================================================
    // =========================================================================
    // Excel 불러오기 액션 패널 (Show3)
    // =========================================================================
    private Control BuildImportActionPanel()
    {
        var statusTb = new TextBlock
        {
            Text         = _importStatus,
            Foreground   = new SolidColorBrush(Color.Parse(_importStatusHex)),
            FontFamily   = Font, FontSize = AppTheme.FontBase,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin       = new Thickness(0, 4, 0, 0),
        };

        var fileListPanel = new StackPanel { Spacing = 3 };
        foreach (var fn in _importFileNames)
            fileListPanel.Children.Add(new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 5,
                Children =
                {
                    new TextBlock { Text = "📅", FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center },
                    new TextBlock { Text = fn, FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppTheme.FgInfo,
                        VerticalAlignment = VerticalAlignment.Center },
                }
            });

        var conflictPanel = new StackPanel { Spacing = 6, IsVisible = false };

        var saveBtn = new Button
        {
            Content         = "💾 결과 일괄 저장",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = AppTheme.BorderActive,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(12, 5),
            IsEnabled       = _importHasPending,
            IsVisible       = _importHasPending,
        };

        if (_importHasConflict && _conflictRows.Count > 0)
            BuildConflictUI(conflictPanel, saveBtn);

        saveBtn.Click += (_, _) =>
        {
            if (_pendingByRow.Count == 0) return;
            int totalSaved = 0, totalFailed = 0;
            foreach (var kv in _pendingByRow)
            {
                try   { totalSaved += TestReportService.BulkUpdateResults(kv.Key, kv.Value); }
                catch { totalFailed++; }
            }
            _importStatus    = $"✅ {totalSaved}개 항목 저장 완료" + (totalFailed > 0 ? $" ({totalFailed}행 실패)" : "");
            _importStatusHex = "#44aa44";
            statusTb.Text       = _importStatus;
            statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
            _importHasPending   = false;
            _importHasConflict  = false;
            _pendingByRow.Clear();
            _conflictRows.Clear();
            saveBtn.IsEnabled = false;
            saveBtn.IsVisible = false;
            LoadData();
        };

        var fileBtn = new Button
        {
            Content         = "📂  Excel 결과 파일 선택",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = AppTheme.BgActiveBlue,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(12, 6),
        };

        var uploadBtn = new Button
        {
            Content         = "📎  파일 첨부",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = AppTheme.BorderSeparator,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(12, 6),
        };

        uploadBtn.Click += async (_, _) =>
        {
            if (_selectedSample == null) { ShowToast("시료를 먼저 선택하세요."); return; }
            var dlg = new Avalonia.Platform.Storage.FilePickerOpenOptions
            {
                Title         = "첨부파일 선택",
                AllowMultiple = true,
                FileTypeFilter = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("모든 파일") { Patterns = new[] { "*.*" } },
                    new Avalonia.Platform.Storage.FilePickerFileType("PDF")       { Patterns = new[] { "*.pdf" } },
                    new Avalonia.Platform.Storage.FilePickerFileType("이미지")    { Patterns = new[] { "*.jpg", "*.jpeg", "*.png", "*.bmp" } },
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel")     { Patterns = new[] { "*.xlsx", "*.xls" } },
                },
            };
            var top = TopLevel.GetTopLevel(this);
            if (top == null) return;
            var files = await top.StorageProvider.OpenFilePickerAsync(dlg);
            if (files.Count == 0) return;

            int sampleId = _selectedSample.Id;
            var destDir  = System.IO.Path.Combine(DbPathHelper.AttachmentDirectory, sampleId.ToString());
            System.IO.Directory.CreateDirectory(destDir);

            int ok = 0;
            foreach (var file in files)
            {
                var src  = file.Path.LocalPath;
                var orig = System.IO.Path.GetFileName(src);
                var dest = System.IO.Path.Combine(destDir, $"{Guid.NewGuid():N}_{orig}");
                try
                {
                    System.IO.File.Copy(src, dest, overwrite: true);
                    AttachmentService.AddAttachment(sampleId, orig, dest);
                    ok++;
                }
                catch (Exception ex) { Log($"[첨부] 복사 실패: {ex.Message}"); }
            }
            ShowToast($"✅ {ok}개 파일 첨부 완료");
            RefreshAttachmentList(sampleId);
        };

        fileBtn.Click += async (_, _) =>
        {
            var dlg = new Avalonia.Platform.Storage.FilePickerOpenOptions
            {
                Title         = "분석기록부 Excel 파일 선택",
                AllowMultiple = true,
                FileTypeFilter = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel 파일")
                    {
                        Patterns = new[] { "*.xlsx" }
                    }
                },
            };
            var top = TopLevel.GetTopLevel(this);
            if (top == null) return;
            var files = await top.StorageProvider.OpenFilePickerAsync(dlg);
            if (files.Count == 0) return;

            _importFileNames = files.Select(f => System.IO.Path.GetFileName(f.Path.LocalPath)).ToList();
            fileListPanel.Children.Clear();
            foreach (var fn in _importFileNames)
                fileListPanel.Children.Add(new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 5,
                    Children =
                    {
                        new TextBlock { Text = "📅", FontSize = AppTheme.FontLG, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = fn, FontFamily = Font, FontSize = AppTheme.FontBase,
                            Foreground = AppTheme.FgInfo,
                            VerticalAlignment = VerticalAlignment.Center },
                    }
                });

            _importStatus    = "파일 읽는 중...";
            _importStatusHex = "#888888";
            statusTb.Text       = _importStatus;
            statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
            _importRows.Clear();
            _pendingByRow.Clear();
            _conflictRows.Clear();
            _importHasPending   = false;
            _importHasConflict  = false;
            saveBtn.IsEnabled   = false;
            saveBtn.IsVisible   = false;
            conflictPanel.IsVisible = false;
            conflictPanel.Children.Clear();

            var rows = new List<ExcelResultRow>();
            foreach (var f in files)
            {
                try
                {
                    var partial = await Task.Run(() => AnalysisRecordService.ReadResultsFromFile(f.Path.LocalPath));
                    rows.AddRange(partial);
                }
                catch (Exception ex)
                {
                    _importStatus    = $"❌ 파일 읽기 실패: {ex.Message}";
                    _importStatusHex = "#cc4444";
                    statusTb.Text       = _importStatus;
                    statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
                    return;
                }
            }

            if (rows.Count == 0)
            {
                _importStatus    = $"⚠ 데이터 없음 ({files.Count}개 파일)";
                _importStatusHex = "#cc8800";
                statusTb.Text       = _importStatus;
                statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
                ResultListChanged?.Invoke(BuildListControl());
                return;
            }

            int matchedCnt = 0, notFoundCnt = 0;
            bool odd = false;
            foreach (var r in rows)
            {
                if (string.IsNullOrEmpty(r.결과값)) continue;
                int? rowId = await Task.Run(() =>
                    TestReportService.FindRowId(r.견적번호, r.약칭, r.시료명));

                bool   found       = rowId.HasValue;
                string existingVal = "";
                bool   hasConflict = false;

                if (found)
                {
                    existingVal = await Task.Run(() =>
                        TestReportService.GetAnalyteValue(rowId!.Value, r.AnalyteName)) ?? "";
                    hasConflict = !string.IsNullOrEmpty(existingVal)
                        && !string.Equals(existingVal, "O", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(existingVal, r.결과값, StringComparison.Ordinal);
                    if (!_pendingByRow.ContainsKey(rowId!.Value))
                        _pendingByRow[rowId!.Value] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    _pendingByRow[rowId!.Value][r.AnalyteName] = r.결과값;
                    if (hasConflict)
                    {
                        _conflictRows.Add((rowId!.Value, r.AnalyteName, existingVal, r.결과값));
                        _importHasConflict = true;
                    }
                    matchedCnt++;
                }
                else notFoundCnt++;

                string bg  = hasConflict ? "#3a2a10" : (odd ? "#1e1e2e" : "#252535");
                string fgR = found ? (hasConflict ? "#ffaa44" : "#88ee88") : "#cc6644";
                odd = !odd;
                _importRows.Add((r.AnalyteName, r.견적번호, r.약칭, r.시료명,
                    existingVal, r.결과값, r.방류기준, bg, fgR));
            }

            int withValue = rows.Count(x => !string.IsNullOrEmpty(x.결과값));
            if (matchedCnt > 0)
            {
                string conflictNote = _conflictRows.Count > 0 ? $"  (⚠ {_conflictRows.Count}개 충돌)" : "";
                _importStatus    = $"✅ {matchedCnt}개 매칭{conflictNote}"
                    + (notFoundCnt > 0 ? $", {notFoundCnt}개 미매칭" : "")
                    + $"  ← {files.Count}개 파일";
                _importStatusHex  = _conflictRows.Count > 0 ? "#ffaa44" : "#44aa44";
                _importHasPending = true;
                saveBtn.IsEnabled = true;
                saveBtn.IsVisible = true;
                if (_importHasConflict)
                {
                    conflictPanel.Children.Clear();
                    BuildConflictUI(conflictPanel, saveBtn);
                    conflictPanel.IsVisible = true;
                }
            }
            else
            {
                _importStatus    = $"⚠ 매칭된 시료 없음 ({withValue}행 읽음)";
                _importStatusHex = "#cc8800";
            }
            statusTb.Text       = _importStatus;
            statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
            ResultListChanged?.Invoke(BuildListControl());
        };

        // 첨부파일 패널
        _attachListPanel = new StackPanel { Spacing = 4 };
        var attachDeleteBtn = new Button
        {
            Content         = "🗑  선택 삭제",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = new SolidColorBrush(Color.Parse("#5a1a1a")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };
        attachDeleteBtn.Click += (_, _) =>
        {
            if (_attachListPanel == null || _selectedSample == null) return;
            var toDelete = new List<(CheckBox cb, SampleAttachment att)>();
            foreach (var child in _attachListPanel.Children)
            {
                if (child is Border row && row.Child is Grid g)
                {
                    var cb = g.Children.OfType<CheckBox>().FirstOrDefault();
                    if (cb?.IsChecked == true && cb.Tag is SampleAttachment att)
                        toDelete.Add((cb, att));
                }
            }
            if (toDelete.Count == 0) { ShowToast("삭제할 파일을 선택하세요."); return; }
            foreach (var (_, att) in toDelete)
                AttachmentService.DeleteAttachment(att.Id, att.저장경로);
            ShowToast($"🗑 {toDelete.Count}개 삭제 완료");
            RefreshAttachmentList(_selectedSample.Id);
        };

        if (_selectedSample != null)
            RefreshAttachmentList(_selectedSample.Id);

        return new Border
        {
            Padding = new Thickness(10),
            Child   = new StackPanel
            {
                Spacing  = 8,
                Children =
                {
                    new TextBlock
                    {
                        Text       = "📊  Excel 결과 불러오기",
                        FontFamily = Font, FontSize = AppTheme.FontLG,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = AppRes("AppFg"),
                    },
                    new Border { Height = 1, Background = AppTheme.BorderMuted },
                    // Excel 선택 + 파일 첨부 버튼 나란히
                    new StackPanel
                    {
                        Orientation = Orientation.Horizontal, Spacing = 6,
                        Children    = { fileBtn, uploadBtn },
                    },
                    fileListPanel,
                    statusTb,
                    conflictPanel,
                    saveBtn,
                    // ── 첨부파일 섹션 ──────────────────────────────────────
                    new Border { Height = 1, Background = AppTheme.BorderSubtle,
                                 Margin = new Thickness(0, 4, 0, 0) },
                    new TextBlock
                    {
                        Text       = "📁  첨부파일",
                        FontFamily = Font, FontSize = AppTheme.FontMD,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = AppTheme.FgInfo,
                    },
                    _attachListPanel,
                    attachDeleteBtn,
                }
            }
        };
    }

    // =========================================================================
    // 첨부파일 목록 갱신
    // =========================================================================
    private void RefreshAttachmentList(int sampleId)
    {
        _attachSampleId = sampleId;
        if (_attachListPanel == null) return;

        _attachListPanel.Children.Clear();
        var attachments = AttachmentService.GetAttachments(sampleId);

        if (attachments.Count == 0)
        {
            _attachListPanel.Children.Add(new TextBlock
            {
                Text       = "첨부된 파일이 없습니다.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Thickness(2, 2),
            });
            return;
        }

        foreach (var att in attachments)
        {
            var cb = new CheckBox
            {
                Tag             = att,
                VerticalAlignment = VerticalAlignment.Center,
            };

            var nameTb = new TextBlock
            {
                Text              = att.원본파일명,
                FontFamily        = Font, FontSize = AppTheme.FontBase,
                Foreground        = AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
                MaxWidth          = 200,
            };

            var dateTb = new TextBlock
            {
                Text              = att.등록일시.Length >= 10 ? att.등록일시[..10] : att.등록일시,
                FontFamily        = Font, FontSize = AppTheme.FontSM,
                Foreground        = AppTheme.FgDimmed,
                VerticalAlignment = VerticalAlignment.Center,
            };

            var openBtn = new Button
            {
                Content         = "↗",
                FontSize        = AppTheme.FontBase, Padding = new Thickness(4, 1),
                Background      = new SolidColorBrush(Color.Parse("#1a2a3a")),
                Foreground      = AppRes("AppFg"),
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(3),
                VerticalAlignment = VerticalAlignment.Center,
                [ToolTip.TipProperty] = "파일 열기",
            };
            openBtn.Click += (_, _) =>
            {
                if (!System.IO.File.Exists(att.저장경로))
                { ShowToast("❌ 파일을 찾을 수 없습니다."); return; }
                try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    { FileName = att.저장경로, UseShellExecute = true }); }
                catch (Exception ex) { ShowToast($"❌ 열기 실패: {ex.Message}"); }
            };

            var row = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto,Auto"),
                Margin            = new Thickness(0, 1),
            };
            row.Children.Add(cb);
            row.Children.Add(nameTb);    nameTb.SetValue(Grid.ColumnProperty, 1);
            row.Children.Add(dateTb);    dateTb.SetValue(Grid.ColumnProperty, 2);
            row.Children.Add(openBtn);   openBtn.SetValue(Grid.ColumnProperty, 3);

            _attachListPanel.Children.Add(new Border
            {
                Background    = AppTheme.BgPrimary,
                CornerRadius  = new CornerRadius(4),
                Padding       = new Thickness(4, 2),
                Child         = row,
            });
        }
    }

    private void BuildConflictUI(StackPanel conflictPanel, Button saveBtn)
    {
        conflictPanel.Children.Add(new TextBlock
        {
            Text         = $"⚠  {_conflictRows.Count}개 항목이 기존 결과와 다릅니다.",
            FontFamily   = Font, FontSize = AppTheme.FontBase,
            Foreground   = AppTheme.FgWarn,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
        });
        var btnRow       = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var skipBtn      = new Button
        {
            Content         = "🚫 충돌 항목 건너버림",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = new SolidColorBrush(Color.Parse("#4a2a10")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };
        var overwriteBtn = new Button
        {
            Content         = "✏ 덮어쓰기",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = AppTheme.BgActiveBlue,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };
        skipBtn.Click += (_, _) =>
        {
            foreach (var c in _conflictRows)
                if (_pendingByRow.TryGetValue(c.RowId, out var d)) d.Remove(c.Analyte);
            _importHasConflict      = false;
            conflictPanel.IsVisible = false;
            saveBtn.Content         = "💾 결과 저장 (충돌 제외)";
        };
        overwriteBtn.Click += (_, _) =>
        {
            _importHasConflict      = false;
            conflictPanel.IsVisible = false;
            saveBtn.Content         = "💾 결과 일괄 저장";
        };
        btnRow.Children.Add(skipBtn);
        btnRow.Children.Add(overwriteBtn);
        conflictPanel.Children.Add(btnRow);
    }

    private static Border MakeImportTableRow(
        string c0, string c1, string c2, string c3, string c4, string c5, string c6,
        bool isHeader, string bgHex, string fgValueHex)
    {
        Color fgH  = Color.Parse(fgValueHex);
        Color fg0  = isHeader ? fgH : Color.Parse("#aaddff");
        Color fg1  = isHeader ? fgH : Color.Parse("#888899");
        Color fg4  = isHeader ? fgH : Color.Parse("#ffaa44");  // 이전결과
        Color fg5  = Color.Parse(fgValueHex);                  // 새결과값
        Color fg6  = isHeader ? fgH : Color.Parse("#ffaa44");  // 방류기준
        return new Border
        {
            Background = new SolidColorBrush(Color.Parse(bgHex)),
            Padding    = new Thickness(6, isHeader ? 4 : 3),
            Margin     = new Thickness(0, 0, 0, isHeader ? 2 : 0),
            Child = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("90,80,55,70,55,55,*"),
                Children =
                {
                    MakeImportCell(c0, 0, fg0),
                    MakeImportCell(c1, 1, fg1),
                    MakeImportCell(c2, 2, fg1),
                    MakeImportCell(c3, 3, fg1),
                    MakeImportCell(c4, 4, fg4),
                    MakeImportCell(c5, 5, fg5),
                    MakeImportCell(c6, 6, fg6),
                }
            }
        };
    }

    private static TextBlock MakeImportCell(string text, int col, Color fg) =>
        new TextBlock
        {
            Text = text ?? "",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = new SolidColorBrush(fg),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            [Grid.ColumnProperty] = col,
        };

    private static TextBlock MakePreviewHeader(string text, int col)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontSM,
            Foreground = AppTheme.FgInfo,
            [Grid.ColumnProperty] = col,
        };
        return tb;
    }

    private static TextBlock MakePreviewCell(string text, int col, IBrush fg)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = fg,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            [Grid.ColumnProperty] = col,
        };
        return tb;
    }

    // =========================================================================
    // 수정 패널 (Show3)
    // =========================================================================
    private string GetStdDisplay(AnalysisResultRow row)
    {
        if (!_showDischargeStd) return "";
        if (_selectedSample == null) return "";
        var col = _selectedSample.방류허용기준;
        if (string.IsNullOrEmpty(col) || col == "기준없음" || col == "해당없음")
            return "";
        // 캐시에서 조회 (각 행마다 DB 조회 없음)
        return _standardValueCache.TryGetValue(row.항목명.Trim(), out var v) ? v : "";
    }

    /// <summary>
    /// 결과값이 기준을 초과하면 true.<br/>
    /// - 기준값이 "A~B" 범위 형식이면: 결과값 &lt; A 이거나 결과값 &gt; B 일 때 true (범위 이탈)<br/>
    /// - 그 외: 결과값 &gt;= 기준값 일 때 true
    /// </summary>
    private static bool ExceedsStandard(string? 결과값, string? 기준값)
    {
        if (string.IsNullOrWhiteSpace(결과값) || string.IsNullOrWhiteSpace(기준값))
            return false;

        // 숫자 추출 헬퍼
        static decimal? Parse(string s)
        {
            var clean = System.Text.RegularExpressions.Regex
                .Match(s.Replace(",", "").Replace("₩", "").Trim(),
                       @"[\d]+(?:\.[\d]+)?")
                .Value;
            return decimal.TryParse(clean,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : null;
        }

        var r = Parse(결과값);
        if (!r.HasValue) return false;

        // "A~B" 범위 형식 (수소이온농도 등)
        var tildeIdx = 기준값.IndexOf('~');
        if (tildeIdx > 0)
        {
            var lo = Parse(기준값[..tildeIdx]);
            var hi = Parse(기준값[(tildeIdx + 1)..]);
            if (!lo.HasValue || !hi.HasValue) return false;
            return r.Value < lo.Value || r.Value > hi.Value;
        }

        // 단순 상한 비교
        var s = Parse(기준값);
        return s.HasValue && r.Value >= s.Value;
    }

    private Control BuildEditPanel(AnalysisResultRow row)
    {
        var standardVal = GetStdDisplay(row);
        var valueTb = new TextBox
        {
            Text = row.결과값, Width = 120, FontFamily = Font, FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator, Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1), BorderBrush = AppTheme.BorderAccent,
            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 4),
        };
        var statusTb = new TextBlock { FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgMuted, VerticalAlignment = VerticalAlignment.Center };
        var saveBtn   = MakeBtn("💾 저장", "#2a5a2a");
        var cancelBtn = MakeBtn("✖ 취소", "#e00000");

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
            else { statusTb.Text = "❌ 저장 실패"; statusTb.Foreground = AppTheme.FgDanger; }
        };
        cancelBtn.Click += (_, _) => { valueTb.Text = row.Original결과값; statusTb.Text = ""; };

        return new Border
        {
            Padding = new Thickness(10),
            Child = new StackPanel { Spacing = 8, Children =
            {
                new TextBlock { Text = "✏️  결과값 수정", FontFamily = Font, FontSize = AppTheme.FontLG,
                               FontWeight = FontWeight.SemiBold, Foreground = AppRes("AppFg") },
                new Border { Height = 1, Background = AppTheme.BorderDefault },
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
                        new TextBlock { Text = "결과값", FontFamily = Font, FontSize = AppTheme.FontSM,
                                       Foreground = AppTheme.FgMuted },
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
                new TextBlock { Text = label, FontFamily = Font, FontSize = AppTheme.FontSM,
                               Foreground = AppTheme.FgMuted },
                new TextBlock { Text = value, FontFamily = Font, FontSize = AppTheme.FontMD,
                               Foreground = highlight
                                   ? AppTheme.FgWarn
                                   : (IBrush)AppRes("FgMuted") }
            }
        };
    }

    private static Button MakeBtn(string text, string bg) => new Button
    {
        Content = text, FontFamily = Font, FontSize = AppTheme.FontMD,
        Background = new SolidColorBrush(Color.Parse(bg)), Foreground = AppRes("AppFg"),
        BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        Padding = new Thickness(12, 5), VerticalAlignment = VerticalAlignment.Center,
    };


    // =========================================================================
    // 자료TO측정인 패널 (Show2 인라인)
    // =========================================================================

    /// <summary>BT8 클릭 시 MainPage 에서 호출 — Show2를 측정인 패널로 전환</summary>
    public void ShowMeasurerPanel()
    {
        if (_selectedSample == null)
        {
            ShowToast("시료를 먼저 선택하세요.");
            return;
        }
        _showMeasurerPanel = true;
        _measurerRows = DataToMeasurerService.BuildRows(_selectedSample);
        ResultListChanged?.Invoke(BuildListControl());
        EditPanelChanged?.Invoke(BuildMeasurerActionPanel());
    }

    private Control BuildMeasurerListControl()
    {
        var topRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
        };

        // IsMeasurerMode 에서는 뒤로가기 불필요 (항상 측정인 형식)
        if (!IsMeasurerMode)
        {
            var backBtn = new Button
            {
                Content         = "◀ 결과 목록으로",
                FontFamily      = Font, FontSize = AppTheme.FontBase,
                Background      = AppTheme.BorderSubtle,
                Foreground      = AppRes("AppFg"),
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
                Padding         = new Thickness(10, 4),
                Margin          = new Thickness(8, 6),
            };
            backBtn.Click += (_, _) =>
            {
                _showMeasurerPanel = false;
                ResultListChanged?.Invoke(BuildListControl());
                EditPanelChanged?.Invoke(null);
            };
            topRow.Children.Add(backBtn);
        }

        var titleTb = new TextBlock
        {
            Text       = IsMeasurerMode
                ? $"측정인  —  {_selectedSample?.약칭} / {_selectedSample?.시료명}"
                : $"자료TO측정인  —  {_selectedSample?.약칭} / {_selectedSample?.시료명}",
            FontFamily = Font, FontSize = AppTheme.FontMD,
            Foreground = new SolidColorBrush(Color.Parse("#aaaaee")),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(IsMeasurerMode ? 10 : 4, 6, 0, 0),
        };
        topRow.Children.Add(titleTb);

        // ── 컬럼 정의 ──────────────────────────────────────────────────────
        const string MeasurerColStr = "*,80,80,*,*,80,80,80";

        // ── 헤더 행 ────────────────────────────────────────────────────────
        var headerBorder = new Border
        {
            Background = AppTheme.BorderSubtle,
            Padding    = new Thickness(8, 5),
            Child = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(MeasurerColStr),
                Children =
                {
                    MeasurerHeaderCell("분석항목",   0),
                    MeasurerHeaderCell("법적기준",   1),
                    MeasurerHeaderCell("결과값",     2),
                    MeasurerHeaderCell("측정방법",   3),
                    MeasurerHeaderCell("장비",       4),
                    MeasurerHeaderCell("담당자",     5),
                    MeasurerHeaderCell("시작일",     6),
                    MeasurerHeaderCell("종료일",     7),
                }
            }
        };

        // ── 데이터 행 ──────────────────────────────────────────────────────
        var listPanel = new StackPanel { Spacing = 0 };
        if (_measurerRows.Count == 0)
        {
            listPanel.Children.Add(new TextBlock
            {
                Text = "분석결과가 없습니다. 결과값을 먼저 입력하세요.",
                FontFamily = Font, FontSize = AppTheme.FontMD,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Thickness(12, 20),
                HorizontalAlignment = HorizontalAlignment.Center,
            });
        }
        else
        {
            bool odd = false;
            foreach (var row in _measurerRows)
            {
                var bg = new SolidColorBrush(odd ? Color.Parse("#2e2e3a") : Color.Parse("#252530"));
                listPanel.Children.Add(new Border
                {
                    Background = bg,
                    Padding    = new Thickness(8, 4),
                    Child = new Grid
                    {
                        ColumnDefinitions = new ColumnDefinitions(MeasurerColStr),
                        Children =
                        {
                            MeasurerDataCell(row.항목명,   0, AppRes("AppFg"), bold: true),
                            MeasurerDataCell(row.법적기준, 1),
                            MeasurerDataCell(row.결과값, 2, AppTheme.FgSuccess, bold: true),
                            MeasurerDataCell(row.측정방법, 3),
                            MeasurerDataCell(row.장비명,   4),
                            MeasurerDataCell(row.담당자,   5),
                            MeasurerDataCell(row.시작일 + " " + row.시작시간, 6),
                            MeasurerDataCell(row.종료일,   7),
                        }
                    }
                });
                odd = !odd;
            }
        }

        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = listPanel,
        };

        var outer = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,*") };
        Grid.SetRow(topRow,       0);
        Grid.SetRow(headerBorder, 1);
        Grid.SetRow(scroll,       2);
        outer.Children.Add(topRow);
        outer.Children.Add(headerBorder);
        outer.Children.Add(scroll);
        return outer;
    }

    private Control BuildMeasurerActionPanel()
    {
        var matchTb = new TextBlock
        {
            FontFamily   = Font, FontSize = AppTheme.FontBase,
            Foreground   = new SolidColorBrush(Color.Parse("#8888aa")),
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin       = new Thickness(0, 4, 0, 0),
            Text         = "브라우저에서 측정인.kr 입력 화면을 먼저 열어 주세요.",
        };

        var statusTb = new TextBlock
        {
            FontFamily   = Font, FontSize = AppTheme.FontBase,
            Foreground   = AppTheme.FgSuccess,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin       = new Thickness(0, 4, 0, 0),
            IsVisible    = false,
        };

        var checkBtn = new Button
        {
            Content         = "페이지 확인",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Background      = AppTheme.BgActiveBlue,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(12, 5),
        };

        var inputBtn = new Button
        {
            Content         = "측정값 입력 진행",
            FontFamily      = Font, FontSize = AppTheme.FontMD,
            Background      = new SolidColorBrush(Color.Parse("#264026")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(16, 6),
        };

        var infoTb = new TextBlock
        {
            Text         = $"항목 {_measurerRows.Count}건 준비",
            FontFamily   = Font, FontSize = AppTheme.FontSM,
            Foreground   = AppTheme.FgDimmed,
            Margin       = new Thickness(0, 8, 0, 0),
        };

        checkBtn.Click += async (_, _) =>
        {
            if (_selectedSample == null) return;
            checkBtn.IsEnabled = false;
            matchTb.Text       = "페이지 확인 중...";
            matchTb.Foreground = AppTheme.FgInfo;
            var (text, hex) = await MeasurerCdpService.CheckPageMatchAsync(
                _selectedSample.약칭, _selectedSample.시료명, _selectedSample.채취일자);
            matchTb.Text       = text;
            matchTb.Foreground = new SolidColorBrush(Color.Parse(hex));
            checkBtn.IsEnabled = true;
        };

        inputBtn.Click += async (_, _) =>
        {
            if (_measurerRows.Count == 0)
            {
                statusTb.Text     = "입력할 항목이 없습니다.";
                statusTb.IsVisible = true;
                return;
            }
            inputBtn.IsEnabled  = false;
            checkBtn.IsEnabled  = false;
            statusTb.IsVisible  = true;
            statusTb.Foreground = AppTheme.FgSuccess;
            try
            {
                var (ok, msg) = await MeasurerCdpService.AutoInputAsync(
                    _measurerRows,
                    text => Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                    {
                        statusTb.Text = text;
                    }));
                statusTb.Text       = (ok ? "✅ " : "⚠ ") + msg;
                statusTb.Foreground = new SolidColorBrush(Color.Parse(ok ? "#44cc44" : "#ffaa44"));
            }
            catch (Exception ex)
            {
                statusTb.Text       = "❌ " + ex.Message;
                statusTb.Foreground = AppTheme.FgDanger;
            }
            finally { inputBtn.IsEnabled = true; checkBtn.IsEnabled = true; }
        };

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 8,
            Children    = { checkBtn, inputBtn },
        };

        return new Border
        {
            Padding = new Thickness(10),
            Child   = new StackPanel { Spacing = 6, Children = { btnRow, matchTb, statusTb, infoTb } }
        };
    }

    private static TextBlock MeasurerHeaderCell(string text, int col)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(tb, col);
        return tb;
    }

    private static TextBlock MeasurerDataCell(string text, int col, IBrush? fg = null, bool bold = false)
    {
        var tb = new TextBlock
        {
            Text         = text ?? "",
            FontFamily   = Font, FontSize = AppTheme.FontBase,
            FontWeight   = bold ? FontWeight.SemiBold : FontWeight.Normal,
            Foreground   = fg ?? AppTheme.FgSecondary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(tb, col);
        return tb;
    }

    // =========================================================================
    // ── 일괄 엑셀 — 시험성적서통합.xlsm 자료 시트에 선택 시료 입력 ────────────
    public async void BatchPrintExcel()
    {
        if (_checkedSamples.Count == 0)
        {
            ShowToast("선택된 시료가 없습니다.\n시료 옆 체크박스를 선택해 주세요.");
            return;
        }

        var progressDlg = ShowToast($"⏳ 통합 엑셀 생성 중… ({_checkedSamples.Count}건)");

        try
        {
            var samples = _checkedSamples.ToList();
            var path = await System.Threading.Tasks.Task.Run(
                () => TestReportPrintService.ExportToIntegratedTemplate(samples, _meta));

            progressDlg.Close();
            ShowToast($"✅ 통합 엑셀 저장 완료\n{System.IO.Path.GetFileName(path)}");
            TestReportPrintService.OpenFile(path);
        }
        catch (Exception ex)
        {
            progressDlg.Close();
            ShowToast($"❌ 오류: {ex.Message}");
            Log($"[일괄통합] 오류: {ex.Message}");
        }
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

    /// BT1 — 시험성적서 엑셀 생성 (템플릿 기반)
    public void Print() => OpenPrintWindow();

    public void OpenPrintWindow()
    {
        if (_selectedSample == null) { ShowToast("시료를 먼저 선택하세요."); return; }
        if (!TestReportPrintService.TemplateExists())
        {
            ShowToast("템플릿 없음\nData/Templates/ 폴더에 템플릿 파일을 넣어주세요.");
            TestReportPrintService.OpenTemplateFolder();
            return;
        }
        try
        {
            var rows = BuildResultRowsForSample(_selectedSample);
            var path = TestReportPrintService.FillAndSave(
                sample: _selectedSample, rows: rows, meta: _meta,
                toPdf: false, openAfter: true);
            Log($"✅ 시험성적서 엑셀 저장: {path}");
            ShowToast($"✅ 저장 완료\n{System.IO.Path.GetFileName(path)}");
        }
        catch (Exception ex) { Log($"❌ 엑셀 오류: {ex.Message}"); ShowToast($"❌ 오류: {ex.Message}"); }
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

    /// BT3 — 삭제 (체크된 노드 전체, 없으면 선택 노드 단건)
    public async Task DeleteSampleAsync()
    {
        var targets = _checkedSamples.Count > 0
            ? _checkedSamples.ToList()
            : (_selectedSample != null ? new List<SampleRequest> { _selectedSample } : new List<SampleRequest>());

        if (targets.Count == 0) { Log("삭제: 선택 없음"); return; }

        string msg = targets.Count == 1
            ? $"'{targets[0].시료명}' 시료를 삭제하시겠습니까?\n(분석결과 포함 전체 삭제)"
            : $"체크된 {targets.Count}건을 모두 삭제하시겠습니까?\n(분석결과 포함 전체 삭제)";

        var confirmed = await ShowConfirmAsync(msg);
        if (!confirmed) return;

        int successCount = 0;
        foreach (var sample in targets)
        {
            bool ok = TestReportService.DeleteSample(sample.Id);
            Log(ok ? $"✅ 삭제: {sample.시료명}" : $"❌ 삭제 실패: {sample.시료명}");
            if (ok) successCount++;
        }

        _checkedSamples.Clear();
        _selectedSample    = null;
        _selectedRow       = null;
        _selectedRowBorder = null;
        _resultRows.Clear();

        Log($"삭제 완료 {successCount}/{targets.Count}건 — 트리 새로고침");
        LoadData();
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
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    private static ColumnDefinitions ResultColDefs =>
        new ColumnDefinitions("36,130,200,90,60,300,100,*");

    private static TextBlock HeaderCell(string text, int col)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(tb, col); return tb;
    }

    private static SelectableTextBlock DataCell(string text, int col, IBrush? fg = null, bool bold = false)
    {
        var tb = new SelectableTextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontMD,
            FontWeight   = bold ? FontWeight.SemiBold : FontWeight.Normal,
            Foreground   = fg ?? AppTheme.FgSecondary,
            VerticalAlignment = VerticalAlignment.Center,
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
                new TextBlock { Text = label + " :", Width = 70, FontFamily = Font, FontSize = AppTheme.FontBase,
                                Foreground = AppTheme.FgMuted,
                                VerticalAlignment = VerticalAlignment.Center },
                new TextBlock { Text = value, FontFamily = Font, FontSize = AppTheme.FontBase,
                                Foreground = AppRes("FgMuted"),
                                VerticalAlignment = VerticalAlignment.Center },
            }
        };
    }

    private Window ShowToast(string msg)
    {
        Log($"[Toast] {msg}");
        // 간단한 팝업 (향후 Toast UI 교체 가능)
        var dlg = new Window
        {
            Title = "알림", Width = 320, Height = 120, CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgSecondary,
        };
        var btn = new Button { Content = "확인", Width = 80,
                               Background = AppTheme.BorderMuted,
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        btn.Click += (_, _) => dlg.Close();
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 14,
            Children =
            {
                new TextBlock { Text = msg, FontFamily = Font, FontSize = AppTheme.FontMD,
                                Foreground = AppRes("AppFg"),
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel { Orientation = Orientation.Horizontal,
                                 HorizontalAlignment = HorizontalAlignment.Right,
                                 Children = { btn } }
            }
        };
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) dlg.ShowDialog(owner);
        else dlg.Show();
        return dlg;
    }

    private async Task<bool> ShowConfirmAsync(string msg)
    {
        var dlg = new Window
        {
            Title = "삭제 확인", Width = 340, Height = 160, CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgSecondary,
        };
        bool result = false;
        var yes = new Button { Content = "삭제", Width = 80,
                               Background = AppTheme.FgDanger,
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        var no  = new Button { Content = "취소", Width = 80,
                               Background = AppTheme.BorderMuted,
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        yes.Click += (_, _) => { result = true;  dlg.Close(); };
        no.Click  += (_, _) => { result = false; dlg.Close(); };
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 14,
            Children =
            {
                new TextBlock { Text = msg, FontFamily = Font, FontSize = AppTheme.FontMD,
                                Foreground = AppRes("AppFg"),
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
        if (App.EnableLogging)
        {
            try { File.AppendAllText("Logs/TestReportDebug.log", line + Environment.NewLine); } catch { }
        }
    }
}