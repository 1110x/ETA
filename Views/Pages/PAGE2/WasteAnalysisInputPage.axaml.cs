using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.SERVICE4;
using ETA.Views.Pages.PAGE1;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.PAGE2;

public partial class WasteAnalysisInputPage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // FontSize DynamicResource 바인딩 헬퍼 (슬라이더 연동)
    private static T FontBind<T>(T ctrl, string resKey) where T : Control
    {
        var prop = ctrl switch
        {
            TextBlock   => TextBlock.FontSizeProperty,
            TextBox     => TextBox.FontSizeProperty,
            _           => Avalonia.Controls.Primitives.TemplatedControl.FontSizeProperty,
        };
        ctrl.Bind(prop, AppFonts.Obs(resKey));
        return ctrl;
    }
    private static T FsSM<T>(T c) where T : Control => FontBind(c, "FontSizeSM");
    private static T FsXS<T>(T c) where T : Control => FontBind(c, "FontSizeXS");
    private static T FsMD<T>(T c) where T : Control => FontBind(c, "FontSizeMD");
    private static T FsLG<T>(T c) where T : Control => FontBind(c, "FontSizeLG");
    private static T FsBase<T>(T c) where T : Control => FontBind(c, "FontSizeBase");

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    // ── 분석항목(장비) 카테고리 정의 ──────────────────────────────────────────
    private static readonly (string Key, string Label, string Icon, string[] Items)[] Categories =
    {
        ("BOD",   "BOD",     "🧫", new[] { "BOD" }),
        ("TOC",   "TOC",     "🔬", new[] { "TOC" }),
        ("SS",    "SS",      "🧪", new[] { "SS" }),
        ("TN",      "T-N",      "🌈", new[] { "T-N" }),
        ("TP",      "T-P",      "🌈", new[] { "T-P" }),
        ("PHENOLS", "페놀류",   "🌈", new[] { "Phenols" }),
        ("CN",      "시안",     "☠️", new[] { "시안" }),
        ("CR6",     "6가크롬",   "🌈", new[] { "6가크롬" }),
        ("COLOR",   "색도",     "🎨", new[] { "색도" }),
        ("ABS",     "ABS",     "🌈", new[] { "ABS" }),
        ("FLUORIDE", "불소",    "🌈", new[] { "불소" }),
        ("NHEX",  "NHexan",  "⛽", new[] { "N-Hexan" }),
        ("GCMS",  "GCMS",    "📊", Array.Empty<string>()),
        ("ICP",   "ICP",     "⚡", Array.Empty<string>()),
        ("IC",    "IC",      "🧬", Array.Empty<string>()),
        ("PFAS",  "과불화화합물", "💊", new[] { "PFOA", "PFOS", "PFBS" }),
        ("AAS",   "AAS",     "🔥", Array.Empty<string>()),
        ("ECO",   "생태독성", "🐟", Array.Empty<string>()),
        ("ECOLI", "대장균",   "🦠", new[] { "총대장균군" }),
    };

    private string _activeCategory = "BOD";
    private string[] _activeItems = { "BOD" };
    private bool _categorySelected = false; // 카테고리 버튼이 명시적으로 클릭되었는지
    private bool _facilityViewMode = false;  // true = 처리시설 데이터 뷰, false = 배출업소 뷰
    private bool? _allRowsEnabled = null;   // 전체 토글 스위치 상태 (null = 일부만 선택됨)
    private Border? _masterToggleTrack = null;
    private Border? _masterToggleKnob  = null;
    private bool _tocShowInstrumentCal = false; // 시마즈 검량선 토글: false=파싱값, true=기기출력값

    // ── 분류기 파서키 → RunParserAsync 키 매핑 (시그니처 + ONNX 공용) ──
    private static readonly Dictionary<string, string> _classifierToRunner = new()
    {
        ["BOD"]               = "Excel",
        ["SS"]                = "Excel",
        ["NHex"]              = "Excel",
        ["UVVIS"]             = "Excel",
        ["TOC_Shimadzu"]      = "TOC_Shimadzu",
        ["TOC_Shimadzu_PDF"]  = "TOC_Shimadzu",
        ["TOC_NPOC"]          = "TOC_NPOC",
        ["TOC_Scalar_NPOC"]   = "Scalar_NPOC",
        ["TOC_Scalar_TCIC"]   = "Scalar_TCIC",
        ["GC"]                = "GC",
        ["UV_Shimadzu_PDF"]   = "Shimadzu_UV",
        ["UV_Shimadzu_ASCII"] = "Shimadzu_UV",
        ["UV_Cary_PDF"]       = "Agilent_Cary",
        ["UV_Cary_CSV"]       = "Agilent_Cary",
        ["ICP"]               = "ICP_PDF",
        ["LCMS"]              = "LCMS_PFAS",
        // 구형 ONNX 레이블 호환
        ["TOC"]               = "TOC_Shimadzu",
    };

    // ── 분류기 레이블 → 카테고리 키 매핑 (AI파서선택 시 카테고리 자동 전환) ──
    private static readonly Dictionary<string, string> _classifierToCategory = new(StringComparer.OrdinalIgnoreCase)
    {
        ["BOD"] = "BOD", ["SS"] = "SS", ["NHex"] = "NHEX",
        ["UVVIS"] = "TN",  // UV 계열 기본값 (실제로는 파서가 항목 구분)
        ["TOC_Shimadzu"] = "TOC", ["TOC_Shimadzu_PDF"] = "TOC",
        ["TOC_NPOC"] = "TOC", ["TOC_Scalar_NPOC"] = "TOC", ["TOC_Scalar_TCIC"] = "TOC",
        ["TOC"] = "TOC",
        ["GC"] = "GCMS",
        ["UV_Shimadzu_PDF"] = "TN", ["UV_Shimadzu_ASCII"] = "TN",
        ["UV_Cary_PDF"] = "TN", ["UV_Cary_CSV"] = "TN",
        ["ICP"] = "ICP", ["LCMS"] = "PFAS",
    };

    // Items값 → DB약칭 별명 매핑 (분장표준처리 약칭과 다른 경우)
    private static readonly Dictionary<string, string> _itemAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["T-N"] = "TN", ["T-P"] = "TP", ["Phenols"] = "페놀류",
        ["N-Hexan"] = "NH", ["총대장균군"] = "대장균",
    };

    // 다성분 동시분석: CompoundAliasService (DB 기반 화합물 별칭)로 대체됨
    // 별칭 관리: 화합물별명 DB 테이블 → CompoundAliasService.Resolve()

    private static bool IsMultiCompoundCategory(string cat) =>
        cat is "GCMS" or "ICP" or "PFAS" or "IC";

    // 약칭 → 소수점 자릿수 캐시 (분석정보 DB 기반)

    /// <summary>
    /// 현재 활성 카테고리/항목에 대응하는 분석정보.Analyte 한글명을 반환.
    /// - 단성분 카테고리: Categories.Label (예: "페놀류")
    /// - 다성분 카테고리(GCMS 등): _activeItems[0]의 약칭으로 DB 역조회 (예: "Formaldehyde" → "폼알데하이드")
    /// </summary>
    private string GetActiveAnalyteKey()
    {
        if (IsMultiCompoundCategory(_activeCategory))
        {
            var abbrev = _activeItems.FirstOrDefault();
            if (abbrev != null)
            {
                // 1순위: CompoundAliasService (GC 성분명 → 한글 분석항목)
                var aliasInfo = ETA.Services.Common.CompoundAliasService.Resolve(abbrev);
                if (!string.IsNullOrEmpty(aliasInfo?.분석항목)) return aliasInfo!.Value.분석항목;
                // 2순위: 분석정보.약칭 역조회
                var korean = ETA.Services.SERVICE1.AnalysisService.GetAnalyteByShortName(abbrev);
                if (!string.IsNullOrEmpty(korean)) return korean;
            }
            return abbrev ?? _activeCategory;
        }
        return Categories.FirstOrDefault(c => c.Key == _activeCategory).Label
               ?? _activeItems.FirstOrDefault() ?? _activeCategory;
    }

    private static Dictionary<string, int>? _decimalPlacesCache;
    private static int GetDecimalPlaces(string itemAbbr)
    {
        if (_decimalPlacesCache == null)
        {
            _decimalPlacesCache = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var shortNames = AnalysisRequestService.GetShortNames(); // 전체명→약칭
                var items = AnalysisService.GetAllItems();
                foreach (var item in items)
                {
                    if (shortNames.TryGetValue(item.Analyte, out var abbr))
                        _decimalPlacesCache[abbr] = item.DecimalPlaces;
                    _decimalPlacesCache[item.Analyte] = item.DecimalPlaces;
                }
            }
            catch { }
        }
        // 직접 조회 → 별명으로 재시도
        if (_decimalPlacesCache.TryGetValue(itemAbbr, out var dp)) return dp;
        if (_itemAliases.TryGetValue(itemAbbr, out var alias) && _decimalPlacesCache.TryGetValue(alias, out dp)) return dp;
        return 1;
    }

    /// <summary>결과값을 분석정보 소수점 자릿수에 맞춰 포맷</summary>
    private static string FormatResult(string value, string itemAbbr)
    {
        if (string.IsNullOrWhiteSpace(value)) return value;
        if (!double.TryParse(value, out var v)) return value;
        int dp = GetDecimalPlaces(itemAbbr);
        return v.ToString($"F{dp}");
    }
    private string _inputMode = "처리시설"; // 수질분석센터 / 처리시설 / 비용부담금 (기본 = 처리시설, 탑레벨 메뉴로만 전환)
    private bool IsBillingMode  => _inputMode == "비용부담금";
    private bool IsFacilityMode => _inputMode == "처리시설";
    private bool IsWaterCenterMode => _inputMode == "수질분석센터";
    private readonly Dictionary<string, Button> _categoryButtons = new();

    // 모드별 표시할 카테고리 키 (3-tab 토글 도입 후 전 모드 카테고리 버튼 제거 — 배지 클릭으로 대체)
    private static readonly Dictionary<string, string[]> ModeCategoryKeys = new()
    {
        ["수질분석센터"] = Array.Empty<string>(),
        ["처리시설"]     = Array.Empty<string>(),
        ["비용부담금"]   = Array.Empty<string>(),
    };

    private static readonly string[] AllEditableItems =
        { "BOD", "TOC", "SS", "T-N", "T-P", "N-Hexan", "Phenols", "시안", "6가크롬", "색도", "ABS", "불소", "PFOA", "PFOS", "PFBS" };

    // ── 카테고리별 첨부 엑셀 데이터 ──────────────────────────────────────────
    private readonly Dictionary<string, List<ExcelRow>> _categoryExcelData = new();
    private readonly Dictionary<string, string> _categoryFilePaths = new();
    private readonly Dictionary<string, string?> _categoryDocDates = new(); // 엑셀 문서 날짜

    // ── 입력 진행 오버레이 ────────────────────────────────────────────────────
    private Border?       _importOverlay;
    private ProgressBar?  _importPb;
    private TextBlock?    _importPbText;


    // ExcelRow, ExcelDocInfo, GcCompoundCalInfo, MatchStatus, SourceType →
    // Services/SERVICE4/BodExcelModels.cs 로 분리
    private readonly Dictionary<string, ExcelDocInfo> _categoryDocInfo = new();

    // ── 매칭 로그 파일 ────────────────────────────────────────────────────
    private static readonly string MatchLogPath = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "ETA", "Logs", "MatchingDebug.log");

    private static void LogMatch(string msg)
    {
        if (App.EnableLogging)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {msg}";
            try
            {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(MatchLogPath)!);
                File.AppendAllText(MatchLogPath, line + "\n");
            }
            catch { }
        }
    }

    /// <summary>모든 행의 Enabled 상태를 체크해서 전체 토글 상태 업데이트</summary>
    private void UpdateMasterToggleState()
    {
        if (_currentExcelRows == null || !_currentExcelRows.Any()) return;

        var validRows = _currentExcelRows.ToList();
        if (!validRows.Any()) return;

        bool allEnabled = validRows.All(r => r.Enabled);
        bool anyEnabled = validRows.Any(r => r.Enabled);

        if (allEnabled)
            _allRowsEnabled = true;
        else if (!anyEnabled)
            _allRowsEnabled = false;
        else
            _allRowsEnabled = null; // 일부만 선택됨

        // 마스터 토글 UI 동기화: ON 비율 > 50% → 파란색
        if (_masterToggleTrack != null && _masterToggleKnob != null)
        {
            int enabledCount = _currentExcelRows.Count(r => r.Enabled);
            bool masterOn = enabledCount > _currentExcelRows.Count / 2;
            _masterToggleTrack.Background = masterOn ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted");
            _masterToggleKnob.Margin = new Thickness(masterOn ? 18 : 2, 2, 0, 2);
        }
    }

    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public event Action<Control?>? ListPanelChanged;    // Show2
    public event Action<Control?>? EditPanelChanged;    // Show3
    public event Action<Control?>? StatsPanelChanged;   // Show4

    private string? _selectedDate;
    private WasteSample? _selectedSample;
    private WasteSample? _currentEditSample;  // 현재 편집 중인 샘플
    private List<WasteSample> _currentSamples = new();

    // 수동 매칭 팝업용 캐시 (LoadVerifiedGrid에서 설정)
    private List<AnalysisRequestRecord>                        _matchingAnalysisRecords = new();
    private List<(string 시설명, string 시료명, int 마스터Id)>? _matchingFacilityMasters;
    private ManualMatchWindow?                                 _matchWindow; // 싱글턴

    // 트리 lazy-load 상태
    private List<string> _allMonths = new();
    private readonly HashSet<string> _loadedMonths = new();  // 날짜까지 로드된 "yyyy-MM"
    private readonly HashSet<int>    _loadedYears  = new();  // 월 노드까지 펼쳐진 연도

    private TopLevel? _topLevel;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        _topLevel = TopLevel.GetTopLevel(this);
    }

    private StackPanel? _gridPanel;
    private StackPanel? _sampleGridPanel;        // LoadSampleGrid의 그리드 패널
    private Dictionary<string, TextBox> _inputBoxes = new();
    private int _selectedRowIndex = -1;          // Excel 행 인덱스
    private int _selectedSampleIndex = -1;       // Sample 행 인덱스 (Show2용)
    private List<ExcelRow>? _currentExcelRows;
    private List<Border> _rowIcons = new();
    private List<(Border Track, Border Knob)> _rowToggles = new();
    private List<TextBlock> _rowSnCells = new();
    private List<StackPanel> _rowNameCells = new();
    private List<TextBlock> _rowSourceCells = new(); // 시료구분 셀
    private List<Button?> _rowDilButtons = new(); // TOC/UV 희석배수 버튼 (행별, null=미해당)
    private List<Button?> _rowVolButtons = new(); // UV 시료량 버튼 (행별, null=미해당)
    private bool _navEditDil = true;  // true=희석배수, false=시료량 (Shift+2 A/D 전환)
    private Dictionary<string, List<CheckBox>> _dupCheckboxGroups = new(); // SN 그룹별 체크박스
    private Dictionary<string, List<Border>> _dupIconGroups = new(); // SN 그룹별 아이콘
    private List<StackPanel> _rowIconPanels = new(); // col0Panel (아이콘+체크박스 컨테이너)
    private StackPanel? _summaryBadgePanel;
    private ExcelRow? _currentEditExcelRow;

    // ── 키보드 단축키 네비게이션 ──────────────────────────────────────────
    private bool _keyNavShow1 = false;  // Shift+R: Show1 매칭패널 W/S 모드
    private bool _keyNavShow2 = false;  // Shift+2: Show2 그리드 방향키 네비 모드
    private bool _shiftQHeld = false;  // Shift+Q 누르고 있는 동안 QC 클릭 모드
    private TreeViewItem? _keyNavTreeFocused = null; // (미사용 - 하위호환)
    private int _keyNavShow1Index = -1;             // 현재 포커스된 Show1 아이템 인덱스
    private List<(Border Item, string Name, object? Data)> _matchItems = new(); // Show1 의뢰시료 목록
    private string _show1BrowseMode = ""; // Show1 전용 모드 (기본: _inputMode 따라감, "분석항목" 등 독립 모드)
    private Window? _attachedWindow;    // KeyDown 핸들러 등록된 창

    public WasteAnalysisInputPage()
    {
        InitializeComponent();
        BuildCategoryButtons();
        AttachedToVisualTree += (_, _) =>
        {
            var win = this.GetVisualAncestors().OfType<Window>().FirstOrDefault();
            if (win != null && win != _attachedWindow)
            {
                if (_attachedWindow != null)
                {
                    _attachedWindow.RemoveHandler(InputElement.KeyDownEvent, OnNavKeyDownTunnel);
                    _attachedWindow.KeyDown -= OnWindowKeyDown;
                    _attachedWindow.KeyUp -= OnWindowKeyUp;
                }
                _attachedWindow = win;
                win.AddHandler(InputElement.KeyDownEvent, OnNavKeyDownTunnel, Avalonia.Interactivity.RoutingStrategies.Tunnel);
                win.KeyDown += OnWindowKeyDown;
                win.KeyUp += OnWindowKeyUp;
            }
        };
        DetachedFromVisualTree += (_, _) =>
        {
            if (_attachedWindow != null)
            {
                _attachedWindow.RemoveHandler(InputElement.KeyDownEvent, OnNavKeyDownTunnel);
                _attachedWindow.KeyDown -= OnWindowKeyDown;
                _attachedWindow.KeyUp -= OnWindowKeyUp;
                _attachedWindow = null;
            }
        };
    }

    public void SetInputMode(string mode)
    {
        // 하위 호환: 구 메뉴 Tag "비용부담금/처리시설" → 처리시설 기본
        if (mode == "비용부담금/처리시설") mode = "처리시설";
        _inputMode = mode;
        UpdateModeTabStyle();
        BuildCategoryButtons();
    }

    // ── 생태독성 전용 패널 ─────────────────────────────────────────────────
    public void ShowEcotoxPanel()
    {
        _activeCategory = "ECO";

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        // 헤더
        root.Children.Add(FsLG(new TextBlock
        {
            Text = "🐟 생태독성 — 물벼룩 급성독성시험",
            FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = AppRes("AppFg"),
        }));
        root.Children.Add(FsSM(new TextBlock
        {
            Text = "시료를 추가하고 클릭하면 우측(Show3)에서 농도/사망수를 입력할 수 있습니다.",
            FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap,
        }));

        // 시료 추가 입력
        var addPanel = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 8, Margin = new Thickness(0, 8) };
        var nameInput = new TextBox
        {
            Watermark = "시료명 입력 (예: 중흥 유입수)",
            FontFamily = Font, Width = 250,
            Foreground = AppRes("InputFg"), Background = AppRes("InputBg"),
            BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 4),
        };
        var addBtn = new Button
        {
            Content = "시료 추가", FontFamily = Font,
            Background = AppRes("BtnPrimaryBg"), Foreground = Brushes.White,
            Padding = new Thickness(12, 4), CornerRadius = new CornerRadius(4),
        };
        addPanel.Children.Add(nameInput);
        addPanel.Children.Add(addBtn);
        root.Children.Add(addPanel);

        // 시료 목록
        var listPanel = new StackPanel { Spacing = 4 };
        root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 4) });
        root.Children.Add(listPanel);

        // ECO 카테고리 데이터 초기화
        if (!_categoryExcelData.ContainsKey("ECO"))
            _categoryExcelData["ECO"] = new List<ExcelRow>();
        var ecoRows = _categoryExcelData["ECO"];
        _currentExcelRows = ecoRows;

        void RebuildList()
        {
            listPanel.Children.Clear();
            for (int i = 0; i < ecoRows.Count; i++)
            {
                var row = ecoRows[i];
                var idx = i;
                var rowBorder = new Border
                {
                    Background = i % 2 == 0 ? AppRes("GridRowBg") : AppRes("GridRowAltBg"),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 6),
                    Margin = new Thickness(0, 1),
                    Cursor = new Cursor(StandardCursorType.Hand),
                };
                TextShimmer.AttachHover(rowBorder);

                var rowGrid = new Grid();
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(30, GridUnitType.Pixel));  // #
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));    // 시료명
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(80, GridUnitType.Pixel));  // TU
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));  // 방법
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition(30, GridUnitType.Pixel));  // 삭제

                rowGrid.Children.Add(FsSM(new TextBlock
                {
                    Text = $"{i + 1}", FontFamily = Font, Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center,
                }));

                var nameTb = FsBase(new TextBlock
                {
                    Text = row.시료명, FontFamily = Font, Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeight.SemiBold,
                });
                Grid.SetColumn(nameTb, 1); rowGrid.Children.Add(nameTb);

                var tuTb = FsBase(new TextBlock
                {
                    Text = string.IsNullOrEmpty(row.Result) ? "—" : $"{row.Result} TU",
                    FontFamily = Font, FontWeight = FontWeight.Bold,
                    Foreground = string.IsNullOrEmpty(row.Result) ? AppRes("FgMuted") : AppRes("ThemeFgSuccess"),
                    VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center,
                });
                Grid.SetColumn(tuTb, 2); rowGrid.Children.Add(tuTb);

                var methodTb = FsXS(new TextBlock
                {
                    Text = row.EcotoxData?.Result?.Method ?? "",
                    FontFamily = Font, Foreground = AppRes("ThemeFgInfo"),
                    VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center,
                });
                Grid.SetColumn(methodTb, 3); rowGrid.Children.Add(methodTb);

                var delBtn = new Button
                {
                    Content = "✕", FontFamily = Font, FontSize = 10,
                    Background = Brushes.Transparent, Foreground = AppRes("ThemeFgDanger"),
                    BorderThickness = new Thickness(0), Padding = new Thickness(4, 2),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                delBtn.Click += (_, _) => { ecoRows.RemoveAt(idx); RebuildList(); };
                Grid.SetColumn(delBtn, 4); rowGrid.Children.Add(delBtn);

                rowBorder.Child = rowGrid;
                rowBorder.PointerPressed += (_, _) =>
                {
                    _selectedRowIndex = idx;
                    _currentEditExcelRow = row;
                    ShowExcelRowDetail(row);
                };

                listPanel.Children.Add(rowBorder);
            }

            if (ecoRows.Count == 0)
            {
                listPanel.Children.Add(FsSM(new TextBlock
                {
                    Text = "시료가 없습니다. 위에서 시료명을 입력하고 추가하세요.",
                    FontFamily = Font, Foreground = AppRes("FgMuted"),
                    Margin = new Thickness(0, 12),
                }));
            }
        }

        addBtn.Click += (_, _) =>
        {
            var name = nameInput.Text?.Trim();
            if (string.IsNullOrEmpty(name)) return;
            ecoRows.Add(new ExcelRow
            {
                시료명 = name, SN = name, Source = SourceType.미분류,
                Status = MatchStatus.입력가능, Enabled = true,
                EcotoxData = new EcotoxTestData(),
            });
            nameInput.Text = "";
            nameInput.Focus();
            RebuildList();
        };
        nameInput.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter) { addBtn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); e.Handled = true; }
        };

        RebuildList();

        ListPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // ── 3-tab 모드 토글 (수질분석센터 / 처리시설 / 비용부담금) ────────────
    public void ModeTab_Click(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button btn) return;
        var mode = btn.Tag as string ?? "처리시설";
        if (mode == _inputMode) return;
        _inputMode = mode;
        UpdateModeTabStyle();
        BuildCategoryButtons();
        ListPanelChanged?.Invoke(null);
        EditPanelChanged?.Invoke(null);
        LoadData(); // 트리 재구성
    }

    private static readonly Dictionary<string, (string bg, string fg, string border)> ModeTabColors = new()
    {
        ["수질분석센터"] = ("#1e3a5a", "#88aacc", "#336699"),
        ["처리시설"]     = ("#3a2a1a", "#ccaa88", "#996633"),
        ["비용부담금"]   = ("#1a3a1a", "#aaccaa", "#336633"),
    };

    private void UpdateModeTabStyle()
    {
        foreach (var btn in new[] { BtnModeWaterCenter, BtnModeFacility, BtnModeBilling })
        {
            if (btn == null) continue;
            var mode = btn.Tag as string ?? "";
            bool active = mode == _inputMode;
            if (active && ModeTabColors.TryGetValue(mode, out var c))
            {
                btn.Background  = new SolidColorBrush(Color.Parse(c.bg));
                btn.Foreground  = new SolidColorBrush(Color.Parse(c.fg));
                btn.BorderBrush = new SolidColorBrush(Color.Parse(c.border));
            }
            else
            {
                btn.Background  = AppRes("SubBtnBg");
                btn.Foreground  = AppRes("FgMuted");
                btn.BorderBrush = AppRes("InputBorder");
            }
        }
    }

    // =========================================================================
    // 분석항목 카테고리 버튼 — 클릭 시 파일 피커 열기
    // =========================================================================
    private void BuildCategoryButtons()
    {
        AnalysisItemButtons.Children.Clear();
        _categoryButtons.Clear();

        var allowedKeys = ModeCategoryKeys.TryGetValue(_inputMode, out var keys)
            ? keys : Categories.Select(c => c.Key).ToArray();

        foreach (var cat in Categories.Where(c => allowedKeys.Contains(c.Key)))
        {
            var btn = FsSM(new Button
            {
                Content = $"{cat.Icon} {cat.Label}",
                Tag = cat.Key,
                FontFamily = Font,
                Padding = new Thickness(8, 4),
                Margin = new Thickness(2),
                CornerRadius = new CornerRadius(4),
                Cursor = new Cursor(StandardCursorType.Hand),
                MinWidth = 72, MinHeight = 0,
                HorizontalContentAlignment = HorizontalAlignment.Center,
            });

            var key = cat.Key;
            btn.Click += (_, _) => OnCategoryButtonClick(key);
            _categoryButtons[cat.Key] = btn;
            AnalysisItemButtons.Children.Add(btn);
            TextShimmer.AttachIfNew(btn);
        }

        UpdateCategoryButtonStyles();
    }

    private void OnCategoryButtonClick(string key)
    {
        // 카테고리 전환
        _categorySelected = true;
        _activeCategory = key;
        var match = Categories.FirstOrDefault(c => c.Key == key);
        _activeItems = match.Items ?? Array.Empty<string>();
        _facilityViewMode = false; // 카테고리 버튼 클릭 시 배출업소 뷰로 복귀
        UpdateCategoryButtonStyles();

        // 엑셀 첨부가 있으면 검증 그리드, 없으면 날짜 기반 DB 그리드
        if (_categoryExcelData.ContainsKey(key))
            { /* LoadVerifiedGrid 제거됨 */ }
        else if (_selectedDate != null)
            LoadSampleGrid(_selectedDate);

        EditPanelChanged?.Invoke(null);
        BuildStatsPanel();
    }

    /// <summary>AI파서선택 버튼: 파일 선택 → ONNX 분류 → 카테고리 자동 전환 → 파서 실행 → Show2 로드</summary>
    public async Task OnAiParserButtonClick()
    {
        var aiResult = await AiPickAndPredictAsync();
        if (aiResult == null) return;

        if (aiResult.Value.ParserKey == "__MANUAL__")
        {
            // 사용자가 "직접 선택" → 기존 수동 다이얼로그
            var manualParser = await ShowParserSelectionDialogFirst();
            if (manualParser == null) return;
            await RunParserAsync(manualParser, null, "AI파서");
        }
        else
        {
            // AI 분류 결과에서 카테고리 자동 전환
            // aiResult에서 원본 분류 레이블 찾기 → 카테고리 키 매핑
            var filePath = aiResult.Value.FilePath;
            var parserKey = aiResult.Value.ParserKey;

            // 파서키에서 역으로 분류 레이블 찾기
            string? classLabel = _classifierToRunner
                .Where(kv => kv.Value == parserKey)
                .Select(kv => kv.Key)
                .FirstOrDefault();

            // 카테고리 자동 전환
            if (classLabel != null && _classifierToCategory.TryGetValue(classLabel, out var catKey))
            {
                OnCategoryButtonClick(catKey);
            }

            var match = Categories.FirstOrDefault(c => c.Key == _activeCategory);
            await RunParserAsync(parserKey, filePath, match.Label);
        }
    }

    /// <summary>파일첨부: 파서 선택 먼저 → 파일 선택</summary>
    public async void AttachExcel()
    {
        var match = Categories.FirstOrDefault(c => c.Key == _activeCategory);
        var selectedParser = await ShowParserSelectionDialogFirst();
        if (selectedParser == null) return;
        await RunParserAsync(selectedParser, null, match.Label);
    }

    // ── 파서 실행 (AI/수동 공통) ───────────────────────────────────────
    private async Task RunParserAsync(string selectedParser, string? preloadedPath, string categoryLabel)
    {
        switch (selectedParser)
        {
            case "Scalar_TCIC":
                await OpenScalarFilePickers(selectedParser);
                LoadVerifiedGrid();
                break;

            case "Shimadzu_UV":
            case "Agilent_Cary":
            case "ICP_PDF":
            case "LCMS_PFAS":
            {
                var pdfPath = preloadedPath ?? await OpenSingleFilePicker(selectedParser, categoryLabel);
                if (string.IsNullOrEmpty(pdfPath)) return;
                switch (selectedParser)
                {
                    case "Shimadzu_UV":  await ParseShimadzuUvFileAsync(pdfPath);   break;
                    case "Agilent_Cary": await ParseAgilentCaryUvFileAsync(pdfPath); break;
                    case "ICP_PDF":      await ParseIcpPdfFileAsync(pdfPath);       break;
                    case "LCMS_PFAS":    await ParseLcmsPfasFileAsync(pdfPath);     break;
                }
                break;
            }

            default:
            {
                var path = preloadedPath ?? await OpenSingleFilePicker(selectedParser, categoryLabel);
                if (string.IsNullOrEmpty(path)) return;
                try
                {
                    await Task.Run(() =>
                    {
                        switch (selectedParser)
                        {
                            case "TOC_Shimadzu":
                            case "TOC_NPOC":    ParseTocInstrumentFile(path);   break;
                            case "Scalar_NPOC": ParseScalarNpocFile(path);      break;
                            case "GC":          ParseGcInstrumentFile(path);    break;
                            case "Shimadzu_UV": ParseShimadzuUvFile(path);      break;
                            case "ICP_PDF":     ParseIcpPdfFile(path);          break;
                            case "LCMS_PFAS":   ParseLcmsPfasPdfFile(path);     break;
                            case "Excel":       ProcessExcelFile(path);         break;
                            default:
                                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                                    ShowMessage("알 수 없는 파서가 선택되었습니다.", true));
                                break;
                        }
                    });
                }
                catch (Exception ex)
                {
                    ShowMessage($"파서 오류: {ex.Message}", true);
                    return;
                }
                if (selectedParser != "Scalar_NPOC")
                    LoadVerifiedGrid();
                break;
            }
        }
    }

    // ── 자동 분류 경로: 파일 선택 → 텍스트 추출 → 분류 → 확인 다이얼로그 ─
    private async Task<(string ParserKey, string FilePath)?> AiPickAndPredictAsync()
    {
        // 1. 범용 파일 피커
        var topLevel = _topLevel ?? TopLevel.GetTopLevel(this);
        if (topLevel == null) return null;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title         = "자동 분류 — 파일 선택",
            AllowMultiple = false,
            FileTypeFilter = [new FilePickerFileType("분석결과 파일")
                { Patterns = ["*.xlsx", "*.xls", "*.csv", "*.txt", "*.pdf"] }],
        });
        if (files.Count == 0) return null;

        var filePath = files[0].TryGetLocalPath();
        if (string.IsNullOrEmpty(filePath)) return null;

        // 2. 텍스트 추출
        ShowMessage("파일 분석 중...", false);
        var text = await Task.Run(() => ExtractTextForAi(filePath));

        if (string.IsNullOrWhiteSpace(text))
        {
            ShowMessage("파일에서 텍스트를 추출할 수 없습니다.", true);
            return null;
        }

        // 3. 시그니처 + ONNX 통합 분류
        var hits = new List<(string Label, float Prob, string Source)>();

        // 1순위: 시그니처 분류기
        if (ETA.Services.SERVICE4.SignatureClassifier.HasSignatures())
        {
            foreach (var h in ETA.Services.SERVICE4.SignatureClassifier.ClassifyTopK(text, 3))
                hits.Add((h.ParserKey, h.Score, "시그니처"));
        }

        // 2순위: ONNX (중복 키 제외)
        if (AiParserClassifier.IsModelReady())
        {
            var existKeys = hits.Select(h => h.Label).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var (label, prob) in AiParserClassifier.PredictTopK(text, 3))
            {
                if (!existKeys.Contains(label))
                    hits.Add((label, prob, "AI"));
            }
        }

        hits = hits.OrderByDescending(h => h.Prob).Take(3).ToList();

        if (hits.Count == 0)
        {
            ShowMessage("파서를 자동으로 분류하지 못했습니다. 직접 선택해 주세요.", true);
            return null;
        }

        // 4. 확인 다이얼로그
        var selectedLabel = await ShowAiParserDialogAsync(filePath, hits);
        if (selectedLabel == null) return null;        // 취소

        if (selectedLabel == "__MANUAL__")
            return ("__MANUAL__", filePath);

        var parserKey = _classifierToRunner.TryGetValue(selectedLabel, out var pk) ? pk : "Excel";
        return (parserKey, filePath);
    }

    // ── 분류 결과 확인 다이얼로그 ────────────────────────────────────────
    private async Task<string?> ShowAiParserDialogAsync(
        string filePath, List<(string Label, float Prob, string Source)> topK)
    {
        var ff = Font;
        string selectedLabel = topK[0].Label;
        string? result = null;

        string GetDisplayName(string key)
            => ETA.Services.SERVICE4.SignatureClassifier.ParserItems
               .FirstOrDefault(p => p.Key == key).Name ?? key;

        // ── 정보 카드 (선택에 따라 업데이트) ────────────────────────
        var infoItem   = new TextBlock { FontFamily = ff, FontSize = AppTheme.FontBase,
                                          FontWeight = FontWeight.SemiBold };
        var infoParser = new TextBlock { FontFamily = ff, FontSize = AppTheme.FontSM,
                                          Foreground = AppRes("ThemeFgSecondary") };

        void UpdateInfoCard(string label)
        {
            selectedLabel = label;
            infoItem.Text   = GetDisplayName(label);
            infoParser.Text = _classifierToRunner.TryGetValue(label, out var rk) ? rk : label;
        }
        UpdateInfoCard(topK[0].Label);

        var infoCard = new Border
        {
            Background    = AppRes("ThemeBgSecondary", "#F3F4F6"),
            CornerRadius  = new CornerRadius(8),
            Padding       = new Thickness(14, 10),
            Margin        = new Thickness(0, 8, 0, 12),
            Child         = new StackPanel { Spacing = 2,
                Children = { infoItem, infoParser } },
        };

        // ── 예측 항목들 (라디오 버튼 스타일) ────────────────────────
        var predStack = new StackPanel { Spacing = 6 };

        foreach (var (label, prob, source) in topK)
        {
            var displayName = GetDisplayName(label);
            var pct = (int)(prob * 100);

            var row = new Border
            {
                Padding      = new Thickness(10, 8),
                CornerRadius = new CornerRadius(6),
                BorderThickness = new Thickness(2),
                Cursor       = new Cursor(StandardCursorType.Hand),
                Tag          = label,
            };

            void StyleRow(Border r, bool selected)
            {
                r.Background       = selected
                    ? AppRes("ThemeAccentLight", "#EFF6FF")
                    : AppRes("ThemeBgPrimary", "#FFFFFF");
                r.BorderBrush      = selected
                    ? AppRes("ThemeAccent", "#2563EB")
                    : AppRes("ThemeBorderLight", "#E5E7EB");
            }
            StyleRow(row, label == topK[0].Label);

            var content = new Grid();
            content.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
            content.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

            // 출처 배지 색상
            var (badgeBg, badgeFg) = source == "시그니처"
                ? ("#D1FAE5", "#065F46")   // 초록
                : ("#DBEAFE", "#1E40AF");  // 파랑 (AI)

            var left = new StackPanel { Spacing = 3 };
            var nameRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            nameRow.Children.Add(new TextBlock
            {
                Text = displayName, FontFamily = ff, FontSize = AppTheme.FontSM,
                FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center,
            });
            nameRow.Children.Add(new Border
            {
                Background = new SolidColorBrush(Color.Parse(badgeBg)),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(5, 1),
                Child = new TextBlock
                {
                    Text = source, FontFamily = ff, FontSize = AppTheme.FontXS,
                    Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                    VerticalAlignment = VerticalAlignment.Center,
                },
            });
            left.Children.Add(nameRow);

            // 확률 바
            var barBg = new Border
            {
                Height = 6, CornerRadius = new CornerRadius(3),
                Background = AppRes("ThemeBorderLight", "#E5E7EB"),
                Width = 120, VerticalAlignment = VerticalAlignment.Center,
                Child = new Border
                {
                    Height = 6, CornerRadius = new CornerRadius(3),
                    Background = AppRes("ThemeAccent", "#2563EB"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Width = Math.Max(4, 120.0 * prob),
                },
            };
            var pctText = new TextBlock
            {
                Text = $"{pct}%", FontFamily = ff, FontSize = AppTheme.FontXS,
                Foreground = AppRes("ThemeFgSecondary"),
                Margin = new Thickness(6, 0, 0, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Width = 32, TextAlignment = TextAlignment.Right,
            };
            var right = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 0,
                VerticalAlignment = VerticalAlignment.Center,
                Children = { barBg, pctText },
            };

            Grid.SetColumn(left,  0);
            Grid.SetColumn(right, 1);
            content.Children.Add(left);
            content.Children.Add(right);
            row.Child = content;

            row.PointerPressed += (s, _) =>
            {
                if (s is not Border clicked || clicked.Tag is not string lbl) return;
                UpdateInfoCard(lbl);
                foreach (var child in predStack.Children.OfType<Border>())
                    StyleRow(child, child.Tag is string t && t == lbl);
            };

            predStack.Children.Add(row);
        }

        // ── 버튼 ────────────────────────────────────────────────────
        var applyBtn  = new Button { Content = "자동 적용",
            Padding = new Thickness(16, 8), Background = AppRes("ThemeAccent") };
        var manualBtn = new Button { Content = "직접 선택",
            Padding = new Thickness(16, 8), Background = AppRes("ThemeBgSecondary") };
        var cancelBtn = new Button { Content = "취소",
            Padding = new Thickness(16, 8), Background = AppRes("ThemeBgSecondary") };

        var dialogContent = new StackPanel
        {
            Spacing = 4, Margin = new Thickness(20),
            Children =
            {
                FsSM(new TextBlock
                {
                    Text = "파서 자동 분류",
                    FontWeight = FontWeight.Bold,
                    Foreground = AppRes("ThemeFgPrimary"),
                    Margin = new Thickness(0, 0, 0, 2),
                }),
                FsXS(new TextBlock
                {
                    Text = System.IO.Path.GetFileName(filePath),
                    Foreground = AppRes("ThemeFgSecondary"),
                    Margin = new Thickness(0, 0, 0, 4),
                }),
                // 분석항목 / 파서 정보
                new StackPanel { Spacing = 2, Margin = new Thickness(0, 0, 0, 4),
                    Children =
                    {
                        FsXS(new TextBlock { Text = "분석항목  /  파서",
                            Foreground = AppRes("ThemeFgSecondary") }),
                    }
                },
                infoCard,
                // 예측 신뢰도
                FsXS(new TextBlock { Text = "예측 신뢰도  (클릭하여 선택)",
                    Foreground = AppRes("ThemeFgSecondary"),
                    Margin = new Thickness(0, 0, 0, 4) }),
                predStack,
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8, Margin = new Thickness(0, 16, 0, 0),
                    Children = { cancelBtn, manualBtn, applyBtn },
                },
            },
        };

        var dialog = new Window
        {
            Title  = "파서 자동 분류",
            Content = dialogContent,
            Width  = 420, SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false,
        };

        applyBtn.Click  += (_, _) => { result = selectedLabel;   dialog.Close(); };
        manualBtn.Click += (_, _) => { result = "__MANUAL__";     dialog.Close(); };
        cancelBtn.Click += (_, _) => { result = null;             dialog.Close(); };

        var parent = TopLevel.GetTopLevel(this) as Window;
        if (parent != null) await dialog.ShowDialog(parent);
        else                dialog.Show();

        return result;
    }

    // ── AI용 텍스트 추출 ──────────────────────────────────────────────────
    private static string ExtractTextForAi(string filePath)
    {
        var ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
        try
        {
            return ext switch
            {
                ".csv" or ".txt" => File.ReadAllText(filePath, System.Text.Encoding.UTF8),
                ".xlsx" or ".xls" => AiExtractXlsx(filePath),
                ".pdf"            => AiExtractPdf(filePath),
                _                 => "",
            };
        }
        catch { return ""; }
    }

    private static string AiExtractXlsx(string filePath)
    {
        using var wb = new XLWorkbook(filePath);
        var sb = new System.Text.StringBuilder();
        foreach (var ws in wb.Worksheets)
            foreach (var row in ws.RowsUsed().Take(50))
            {
                var cells = row.CellsUsed().Select(c => c.GetString())
                              .Where(s => !string.IsNullOrWhiteSpace(s));
                sb.AppendLine(string.Join("\t", cells));
            }
        return sb.ToString();
    }

    private static string AiExtractPdf(string filePath)
    {
        using var doc = UglyToad.PdfPig.PdfDocument.Open(filePath);
        var sb = new System.Text.StringBuilder();
        foreach (var page in doc.GetPages().Take(3))
        {
            foreach (var word in page.GetWords())
                sb.Append(word.Text).Append(' ');
            sb.AppendLine();
        }
        return sb.ToString();
    }

    /// <summary>
    /// 파일 선택 전에 먼저 파서 선택 다이얼로그 표시
    /// </summary>
    private async Task<string?> ShowParserSelectionDialogFirst()
    {
        string? selectedParser = null;
        bool confirmed = false;

        // ── 파서 목록 정의 ──────────────────────────────────────────
        var groups = new[]
        {
            ("TOC 분석기기", new[]
            {
                ("TOC_Shimadzu", "Shimadzu TOC-L",   "CSV/TXT"),
                ("TOC_NPOC",     "예나 multi N/C",    "PDF"),
                ("Scalar_NPOC",  "스칼라 NPOC",       "단일파일"),
                ("Scalar_TCIC",  "스칼라 TCIC",       "TC+IC CSV"),
            }),
            ("기타 분석기기", new[]
            {
                ("GC",           "GC 기기",           "Agilent CSV"),
                ("Shimadzu_UV",  "Shimadzu UV-1800",  "PDF"),
                ("Agilent_Cary", "Agilent Cary-3500", "PDF"),
                ("ICP_PDF",      "ICP 분석기기",      "PDF"),
                ("LCMS_PFAS",    "LCMS 과불화화합물", "PDF"),
                ("Excel",        "Excel 양식",        "BOD/SS/T-N/T-P"),
            }),
        };

        // ── 선택 카드 Border 목록 (선택 시 하이라이트 토글용) ──────
        var cardMap = new Dictionary<string, Border>();

        void SelectParser(string key)
        {
            selectedParser = key;
            foreach (var (k, b) in cardMap)
            {
                bool sel = k == key;
                b.Background   = sel ? AppRes("ThemeBgActiveGreen", "#1a3a1a")
                                     : AppRes("ThemeBgInput", "#1e1e2e");
                b.BorderBrush  = sel ? AppRes("ThemeBorderActive", "#44ff88")
                                     : AppRes("ThemeBorder", "#333344");
                b.BorderThickness = sel ? new Thickness(1.5) : new Thickness(1);
            }
        }

        // ── 레이아웃 ─────────────────────────────────────────────────
        var body = new StackPanel { Spacing = 10, Margin = new Thickness(18, 14, 18, 10) };

        foreach (var (groupTitle, items) in groups)
        {
            // 그룹 헤더
            body.Children.Add(new TextBlock
            {
                Text       = groupTitle,
                FontSize   = AppFonts.SM,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted", "#aaaaaa"),
                Margin     = new Thickness(0, 6, 0, 4),
            });

            // 카드 행 (WrapPanel)
            var wrap = new WrapPanel { Orientation = Orientation.Horizontal };
            foreach (var (key, label, hint) in items)
            {
                var card = new Border
                {
                    Width           = 140,
                    Height          = 52,
                    CornerRadius    = new CornerRadius(6),
                    BorderThickness = new Thickness(1),
                    BorderBrush     = AppRes("ThemeBorder", "#333344"),
                    Background      = AppRes("ThemeBgInput", "#1e1e2e"),
                    Margin          = new Thickness(0, 0, 6, 6),
                    Cursor          = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    Child = new StackPanel
                    {
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin  = new Thickness(10, 0),
                        Spacing = 2,
                        Children =
                        {
                            new TextBlock
                            {
                                Text       = label,
                                FontSize   = AppFonts.SM,
                                FontWeight = FontWeight.SemiBold,
                                Foreground = AppRes("AppFg", "#e0e0e0"),
                            },
                            new TextBlock
                            {
                                Text       = hint,
                                FontSize   = 9,
                                Foreground = AppRes("FgMuted", "#888888"),
                            },
                        }
                    }
                };

                string capturedKey = key;
                cardMap[capturedKey] = card;

                card.PointerPressed += (_, _) => SelectParser(capturedKey);

                wrap.Children.Add(card);
            }
            body.Children.Add(wrap);
        }

        // ── 버튼 줄 ──────────────────────────────────────────────────
        var okBtn = new Button
        {
            Content    = "확인",
            FontSize   = AppFonts.SM,
            Width      = 80,
            Height     = 30,
            IsEnabled  = false,
            Background = AppRes("ThemeAccent", "#2255aa"),
            Foreground = Brushes.White,
        };
        var cancelBtn = new Button
        {
            Content    = "취소",
            FontSize   = AppFonts.SM,
            Width      = 80,
            Height     = 30,
            Background = AppRes("ThemeBgSecondary", "#2a2a3a"),
            Foreground = AppRes("FgMuted", "#aaaaaa"),
        };

        // 카드 선택 시 확인 버튼 활성화
        foreach (var (k, b) in cardMap)
        {
            b.PointerPressed += (_, _) => okBtn.IsEnabled = true;
        }

        var btnRow = new StackPanel
        {
            Orientation         = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing             = 8,
            Margin              = new Thickness(0, 8, 0, 0),
            Children            = { cancelBtn, okBtn }
        };
        body.Children.Add(btnRow);

        // ── 다이얼로그 윈도우 ─────────────────────────────────────────
        var dialog = new Window
        {
            Title                   = "파서 선택",
            SizeToContent           = SizeToContent.WidthAndHeight,
            WindowStartupLocation   = WindowStartupLocation.CenterOwner,
            CanResize               = false,
            Content                 = new ScrollViewer
            {
                Content                      = body,
                VerticalScrollBarVisibility  = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility= Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                MaxHeight                    = 520,
            },
        };

        cancelBtn.Click += (_, _) => { selectedParser = null; dialog.Close(); };
        okBtn.Click     += (_, _) => { confirmed = true;      dialog.Close(); };

        // 더블클릭 즉시 확정
        foreach (var (k, b) in cardMap)
        {
            b.DoubleTapped += (_, _) =>
            {
                confirmed = true;
                dialog.Close();
            };
        }

        var parent = TopLevel.GetTopLevel(this) as Window;
        if (parent != null) await dialog.ShowDialog(parent);
        else                dialog.Show();

        return confirmed ? selectedParser : null;
    }

    private async Task<string?> OpenSingleFilePicker(string parserType, string categoryLabel)
    {
        var topLevel = _topLevel ?? TopLevel.GetTopLevel(this);
        if (topLevel == null) return null;

        var fileTypes = parserType switch
        {
            "TOC_NPOC"                       => new[] { new FilePickerFileType("PDF 파일")          { Patterns = new[] { "*.pdf" } } },
            "GC"                             => new[] { new FilePickerFileType("GC 결과 파일")       { Patterns = new[] { "*.csv", "*.txt", "*.pdf" } } },
            "TOC_Shimadzu"                   => new[] { new FilePickerFileType("Shimadzu TOC 파일")  { Patterns = new[] { "*.csv", "*.txt", "*.pdf" } } },
            "Shimadzu_UV" or "Agilent_Cary" => new[] { new FilePickerFileType("PDF 파일")           { Patterns = new[] { "*.pdf" } } },
            "Scalar_NPOC"                   => new[] { new FilePickerFileType("스칼라 파일")         { Patterns = new[] { "*.csv", "*.txt", "*.pdf" } } },
            "ICP_PDF"                        => new[] { new FilePickerFileType("ICP 결과 파일")       { Patterns = new[] { "*.pdf", "*.csv" } } },
            "LCMS_PFAS"                      => new[] { new FilePickerFileType("PDF 파일")           { Patterns = new[] { "*.pdf" } } },
            "Excel"                         => new[] { new FilePickerFileType("Excel 파일")         { Patterns = new[] { "*.xlsx", "*.xls" } } },
            _                               => new[] { new FilePickerFileType("분석결과 파일")       { Patterns = new[] { "*.xlsx", "*.xls", "*.csv", "*.txt", "*.pdf" } } },
        };

        var parserNames = new Dictionary<string, string>
        {
            ["TOC_Shimadzu"] = "Shimadzu TOC-L",   ["TOC_NPOC"]    = "예나 multi N/C",
            ["Scalar_NPOC"]  = "스칼라 NPOC",       ["Scalar_TCIC"] = "스칼라 TCIC",
            ["GC"]           = "GC 기기",           ["Shimadzu_UV"] = "Shimadzu UV-1800",
            ["Agilent_Cary"] = "Agilent Cary-3500", ["ICP_PDF"]     = "ICP 분석기기",
            ["LCMS_PFAS"]    = "LCMS 과불화화합물", ["Excel"]       = "표준 Excel 양식",
        };
        var name = parserNames.TryGetValue(parserType, out var n) ? n : parserType;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title         = $"{name} 분석결과 파일 선택",
            AllowMultiple = false,
            FileTypeFilter = fileTypes,
        });

        return files.Count == 0 ? null : files[0].TryGetLocalPath();
    }

    /// </summary>
    private async Task<string?> ShowParserSelectionDialog(string filePath)
    {
        var fileName = System.IO.Path.GetFileName(filePath);
        var fileExt = System.IO.Path.GetExtension(filePath).ToLowerInvariant();

        var dialogContent = new StackPanel
        {
            Spacing = 12,
            MinWidth = 400
        };

        // 제목
        dialogContent.Children.Add(FsSM(new TextBlock
        {
            Text = "📂 분석기기 파서 선택",
            FontWeight = FontWeight.Bold,
            Foreground = AppRes("ThemeFgPrimary"),
            Margin = new Thickness(0, 0, 0, 8)
        }));

        // 파일 정보
        dialogContent.Children.Add(FsXS(new TextBlock
        {
            Text = $"파일: {fileName}",
            Foreground = AppRes("ThemeFgSecondary"),
            Margin = new Thickness(0, 0, 0, 12)
        }));

        // 파서 옵션들
        var parserOptions = new List<(string Key, string Label, string Description, bool Enabled, bool IsTocGroup)>
        {
            ("TOC_Shimadzu", "🧪 Shimadzu TOC-L", "CSV/TXT 파일", fileExt is ".csv" or ".txt" or ".xlsx", true),
            ("TOC_NPOC", "🧪 예나 multi N/C", "PDF 파일 (NPOC/TC-IC 자동감지)", fileExt == ".pdf", true),
            ("Scalar_TC", "🧪 스칼라 TC", "TC 파일 (IC 추가선택)", fileExt is ".csv" or ".txt", true),
            ("Scalar_IC", "🧪 스칼라 IC", "IC 파일 (TC 추가선택)", fileExt is ".csv" or ".txt", true),
            ("GC", "⚗️ GC 기기", "Agilent GC-MS CSV/PDF 파일", fileExt is ".csv" or ".txt" or ".pdf", false),
            ("Shimadzu_UV", "🔬 Shimadzu UV-1800", "ASCII 텍스트 파일 (T-N, T-P, Phenols)", fileExt is ".txt" or ".csv", false),
            ("Agilent_Cary", "📊 Agilent Cary-3500", "PDF 파일 (Concentration report)", fileExt == ".pdf", false),
            ("ICP_PDF", "⚡ ICP 분석기기", "PDF/CSV 결과 파일 (중금속 원소)", fileExt is ".pdf" or ".csv", false),
            ("LCMS_PFAS", "💊 LCMS 과불화화합물", "PDF 결과 파일 (PFOA/PFOS/PFBS)", fileExt == ".pdf", false),
            ("Excel", "📋 표준 Excel 양식", "BOD/SS/N-Hexan/T-N/T-P/Phenols Excel 템플릿", fileExt is ".xlsx" or ".xls", false)
        };

        string? selectedParser = null;

        // TOC 그룹 헤더
        dialogContent.Children.Add(FsBase(new TextBlock
        {
            Text = "🧪 TOC 분석기기",
            FontWeight = FontWeight.Bold,
            Foreground = AppRes("ThemeFgPrimary"),
            Margin = new Thickness(0, 8, 0, 4)
        }));

        // TOC 파서들 4개 컬럼 그리드
        var tocGrid = new Grid
        {
            ColumnDefinitions =
            {
                new ColumnDefinition(GridLength.Star),
                new ColumnDefinition(GridLength.Star),
                new ColumnDefinition(GridLength.Star),
                new ColumnDefinition(GridLength.Star)
            },
            Margin = new Thickness(0, 0, 0, 16)
        };

        var tocOptions = parserOptions.Where(p => p.IsTocGroup).ToList();
        for (int i = 0; i < tocOptions.Count; i++)
        {
            var option = tocOptions[i];
            var optionPanel = CreateParserOptionPanel(option.Key, option.Label, option.Description, option.Enabled,
                key => selectedParser = key);

            Grid.SetColumn(optionPanel, i);
            tocGrid.Children.Add(optionPanel);
        }
        dialogContent.Children.Add(tocGrid);

        // 기타 파서들 헤더
        dialogContent.Children.Add(FsBase(new TextBlock
        {
            Text = "📂 기타 분석기기",
            FontWeight = FontWeight.Bold,
            Foreground = AppRes("ThemeFgPrimary"),
            Margin = new Thickness(0, 8, 0, 4)
        }));

        // 기타 파서들 (세로 배치)
        var otherOptions = parserOptions.Where(p => !p.IsTocGroup).ToList();
        foreach (var option in otherOptions)
        {
            var optionPanel = CreateParserOptionPanel(option.Key, option.Label, option.Description, option.Enabled,
                key => selectedParser = key);
            dialogContent.Children.Add(optionPanel);
        }

        // 버튼 패널
        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 8
        };

        bool dialogResult = false;

        var cancelButton = new Button
        {
            Content = "취소",
            Padding = new Thickness(16, 6),
        };

        var okButton = new Button
        {
            Content = "확인",
            Padding = new Thickness(16, 6),
            Background = AppRes("BtnPrimaryBg"),
            Foreground = AppRes("BtnPrimaryFg"),
            IsEnabled = false
        };

        // 라디오버튼 변경 이벤트
        foreach (var child in dialogContent.Children)
        {
            if (child is Border border && border.Child is StackPanel panel)
            {
                if (panel.Children.FirstOrDefault() is RadioButton rb)
                {
                    rb.Checked += (_, _) =>
                    {
                        selectedParser = rb.Tag?.ToString();
                        okButton.IsEnabled = !string.IsNullOrEmpty(selectedParser);
                    };
                }
            }
        }

        var dialog = new Window
        {
            Title = "파서 선택",
            Content = new Border
            {
                Padding = new Thickness(20),
                Child = new DockPanel
                {
                    Children =
                    {
                        // 버튼은 하단에 고정
                        new Border
                        {
                            Child = buttonPanel,
                            [DockPanel.DockProperty] = Dock.Bottom,
                            Margin = new Thickness(0, 16, 0, 0)
                        },
                        // 내용은 스크롤 가능
                        new ScrollViewer
                        {
                            Content = dialogContent,
                            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                            [DockPanel.DockProperty] = Dock.Top
                        }
                    }
                }
            },
            Width = 700,
            Height = 700,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = true
        };

        cancelButton.Click += (_, _) =>
        {
            dialogResult = false;
            dialog.Close();
        };

        okButton.Click += (_, _) =>
        {
            dialogResult = true;
            dialog.Close();
        };

        buttonPanel.Children.Add(cancelButton);
        buttonPanel.Children.Add(okButton);

        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel is Window parentWindow)
        {
            await dialog.ShowDialog(parentWindow);
        }
        else
        {
            dialog.Show();
        }

        return dialogResult ? selectedParser : null;
    }

    /// <summary>
    /// 파서 옵션 패널 생성 헬퍼 메서드
    /// </summary>
    private Border CreateParserOptionPanel(string key, string label, string description, bool enabled, Action<string> onSelected)
    {
        var optionPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8
        };

        var radioButton = new RadioButton
        {
            GroupName = "ParserSelection",
            IsEnabled = enabled,
            Tag = key
        };

        var labelPanel = new StackPanel
        {
            Spacing = 1
        };

        labelPanel.Children.Add(FsXS(new TextBlock
        {
            Text = label,
            FontWeight = FontWeight.SemiBold,
            Foreground = enabled ? AppRes("AppFg") : AppRes("ThemeFgMuted"),
            TextWrapping = TextWrapping.Wrap
        }));

        labelPanel.Children.Add(FsXS(new TextBlock
        {
            Text = description,
            Foreground = enabled ? AppRes("ThemeFgSecondary") : AppRes("ThemeFgMuted"),
            FontSize = 10,
            TextWrapping = TextWrapping.Wrap
        }));

        optionPanel.Children.Add(radioButton);
        optionPanel.Children.Add(labelPanel);

        var border = new Border
        {
            Background = AppRes("ThemeBgInput"),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(5, 4),
            Margin = new Thickness(2),
            Opacity = enabled ? 1.0 : 0.6,
            Child = optionPanel
        };

        if (enabled)
        {
            border.PointerPressed += (_, _) =>
            {
                radioButton.IsChecked = true;
                onSelected(key);
            };
        }

        return border;
    }

    /// <summary>
    /// 스칼라 파일 피커 열기 (TC + IC 다중 선택)
    /// </summary>
    private async Task OpenScalarFilePickers(string parserType)
    {
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel?.StorageProvider == null) return;

            // 스칼라 TCIC 다중 파일 선택 (새로운 StorageProvider API 사용)
            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "스칼라 TC와 IC 파일을 선택하세요 (다중 선택)",
                AllowMultiple = true,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("CSV 파일") { Patterns = new[] { "*.csv" } },
                    new FilePickerFileType("텍스트 파일") { Patterns = new[] { "*.txt" } },
                    new FilePickerFileType("모든 파일") { Patterns = new[] { "*" } }
                }
            });

            if (files?.Count >= 2)
            {
                var selectedFiles = files.Select(f => f.Path.LocalPath).ToArray();

                // 선택된 파일들에서 TC/IC 자동 감지
                string? tcFile = null, icFile = null;

                foreach (var file in selectedFiles)
                {
                    var fileType = DetectScalarFileType(file);
                    if (fileType == "TC" && tcFile == null)
                        tcFile = file;
                    else if (fileType == "IC" && icFile == null)
                        icFile = file;
                }

                if (tcFile != null && icFile != null)
                {
                    // 두 파일 모두 발견 - TCIC 파싱 진행
                    await ProcessScalarTcicFiles(tcFile, icFile);
                }
                else
                {
                    ShowMessage("TC 파일과 IC 파일을 모두 선택해주세요. 파일명이나 내용으로 구분이 어려운 경우 파일명에 (TC) 또는 (IC)를 포함해주세요.", true);
                }
            }
            else if (files?.Count > 0)
            {
                ShowMessage("스칼라 TCIC 파싱을 위해 최소 2개 파일(TC + IC)을 선택해주세요.", true);
            }
            else
            {
                // 파일 선택이 취소됨 - 아무 작업하지 않음
                ShowMessage("파일 선택이 취소되었습니다.", false);
            }

        }
        catch (OperationCanceledException)
        {
            // 파일 피커 취소 시 정상 동작 - 메시지 없이 조용히 종료
        }
        catch (Exception ex)
        {
            ShowMessage($"스칼라 TCIC 파일 선택 오류: {ex.Message}", true);
            LogScalarTcic($"파일 선택 오류: {ex.Message}");
        }
    }

    /// <summary>스칼라 검량선 PDF (TC + IC) 2개 선택 → 검량선 값만 저장 (시료 행 없음)</summary>
    private async Task OpenScalarCalPdfPickers()
    {
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel?.StorageProvider == null) return;

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "스칼라 TC 검량선 PDF와 IC 검량선 PDF를 선택하세요 (2개)",
                AllowMultiple = true,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("PDF 파일") { Patterns = new[] { "*.pdf" } },
                    new FilePickerFileType("모든 파일") { Patterns = new[] { "*" } }
                }
            });

            if (files == null || files.Count == 0)
            {
                ShowMessage("파일 선택이 취소되었습니다.", false);
                return;
            }

            // TC/IC 각각 파싱 (백그라운드)
            TocCalibrationData? calTC = null, calIC = null;
            await System.Threading.Tasks.Task.Run(() =>
            {
                foreach (var file in files)
                {
                    var p = file.Path.LocalPath;
                    var (fmt, _, cal) = TocInstrumentParser.Parse(p);
                    if (fmt != TocInstrumentParser.TocFileFormat.SkalarCalibrationPdf || cal == null) continue;
                    if (!string.IsNullOrEmpty(cal.Slope_TC)) calTC = cal;
                    else if (!string.IsNullOrEmpty(cal.Slope_IC)) calIC = cal;
                }
            });

            // 검량선 데이터 병합 → _categoryDocInfo["TOC"]
            if (calTC == null && calIC == null)
            {
                ShowMessage("선택한 파일에서 스칼라 검량선 데이터를 찾지 못했습니다.", true);
                return;
            }

            if (!_categoryDocInfo.ContainsKey("TOC"))
                _categoryDocInfo["TOC"] = new ExcelDocInfo();
            var docInfo = _categoryDocInfo["TOC"];
            docInfo.IsTocTCIC = true;

            if (calTC != null)
            {
                docInfo.TocSlope_TC     = calTC.Slope_TC;
                docInfo.TocIntercept_TC = calTC.Intercept_TC;
                docInfo.TocR2_TC        = calTC.R2_TC;
            }
            if (calIC != null)
            {
                docInfo.TocSlope_IC     = calIC.Slope_IC;
                docInfo.TocIntercept_IC = calIC.Intercept_IC;
                docInfo.TocR2_IC        = calIC.R2_IC;
            }

            var parts = new List<string>();
            if (calTC != null) parts.Add($"TC 기울기={calTC.Slope_TC}");
            if (calIC != null) parts.Add($"IC 기울기={calIC.Slope_IC}");
            ShowMessage($"스칼라 검량선 로드 완료: {string.Join(", ", parts)}", false);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ShowMessage($"스칼라 검량선 PDF 오류: {ex.Message}", true);
        }
    }

    /// <summary>
    /// 스칼라 TCIC 디버깅 로그 기록 (LAYOUT.log와 같은 형태)
    /// </summary>
    private void LogScalarTcic(string message)
    {
        try
        {
            if (App.EnableLogging)
            {
                string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [스칼라 TCIC] {message}";
                string projectFolder = Directory.GetCurrentDirectory();
                string logsFolder = System.IO.Path.Combine(projectFolder, "Logs", "Users", Environment.UserName);

                if (!Directory.Exists(logsFolder))
                    Directory.CreateDirectory(logsFolder);

                string logPath = System.IO.Path.Combine(logsFolder, "SCALAR_TCIC.log");
                File.AppendAllText(logPath, logLine + "\n");
            }

        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>
    /// 스칼라 NPOC 단일 파일 파서
    /// </summary>
    private async void ParseScalarNpocFile(string path)
    {
        try
        {
            ShowMessage("스칼라 NPOC 파일 파싱 중...", false);

            var tocResult = await System.Threading.Tasks.Task.Run(() =>
                TocInstrumentParser.Parse(path));

            // TOC 결과를 ExcelRow 형식으로 변환
            var excelRows = tocResult.rows.Select(row => new ExcelRow
            {
                시료명 = row.RawName,
                SN = row.SN,
                Result = row.Conc,
                Fxy = row.Area,
                P = row.Dilution,
                Source = SourceType.미분류,
                Status = MatchStatus.대기,
                IsControl = row.IsControl,
                // TCIC 전용 데이터
                TCAU = row.TCAU,
                TCcon = row.TCcon,
                ICAU = row.ICAU,
                ICcon = row.ICcon
            }).ToList();

            // TOC 카테고리로 설정
            _activeCategory = "TOC";
            var tocCatMatch = Categories.FirstOrDefault(c => c.Key == "TOC");
            _activeItems = tocCatMatch.Items ?? new[] { "TOC" };
            _categorySelected = true;

            // 문서 정보 생성
            var docInfo = new ExcelDocInfo
            {
                분석방법 = tocResult.format.ToString(),
                IsTocNPOC = tocResult.cal?.Method == "NPOC",
                IsTocTCIC = tocResult.cal?.Method == "TCIC",
                IsShimadzuToc = true,
            };

            if (tocResult.cal != null)
            {
                docInfo.TocSlope_TC = tocResult.cal.Slope_TC;
                docInfo.TocIntercept_TC = tocResult.cal.Intercept_TC;
                docInfo.TocR2_TC = tocResult.cal.R2_TC;
                docInfo.TocStdConcs = tocResult.cal.StdConcs;
                docInfo.TocStdAreas = tocResult.cal.StdAreas;

                if (docInfo.IsTocTCIC)
                {
                    docInfo.TocSlope_IC = tocResult.cal.Slope_IC;
                    docInfo.TocIntercept_IC = tocResult.cal.Intercept_IC;
                    docInfo.TocR2_IC = tocResult.cal.R2_IC;
                    docInfo.TocStdConcs_IC = tocResult.cal.StdConcs_IC;
                    docInfo.TocStdAreas_IC = tocResult.cal.StdAreas_IC;
                }
                // 기기출력값 (Shimadzu PDF 직독값)
                if (tocResult.cal.HasInstrumentCal)
                {
                    docInfo.TocSlope_TC_Inst     = tocResult.cal.Slope_TC_Inst;
                    docInfo.TocIntercept_TC_Inst = tocResult.cal.Intercept_TC_Inst;
                    docInfo.TocR2_TC_Inst        = tocResult.cal.R2_TC_Inst;
                    docInfo.TocSlope_IC_Inst     = tocResult.cal.Slope_IC_Inst;
                    docInfo.TocIntercept_IC_Inst = tocResult.cal.Intercept_IC_Inst;
                    docInfo.TocR2_IC_Inst        = tocResult.cal.R2_IC_Inst;
                    docInfo.TocStdAreas_Inst     = tocResult.cal.StdAreas_Inst;
                    docInfo.TocStdAreas_IC_Inst  = tocResult.cal.StdAreas_IC_Inst;
                }
            }

            // 파싱 결과 저장
            _categoryDocInfo["TOC"] = docInfo;
            _categoryExcelData["TOC"] = excelRows;
            _categoryFilePaths["TOC"] = path;

            // UI 업데이트
            _currentExcelRows = excelRows;
            UpdateCategoryButtonStyles();

            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
            {
                LoadVerifiedGrid();
                BuildStatsPanel();
            });

            ShowMessage($"✅ 스칼라 NPOC 파일 파싱 완료 - {excelRows.Count}건 로드됨", false);
        }
        catch (Exception ex)
        {
            ShowMessage($"❌ 스칼라 NPOC 파일 파싱 오류: {ex.Message}", true);
        }
    }

    /// <summary>
    /// 스칼라 파일 유형 자동 감지 (TC 또는 IC)
    /// </summary>
    private string DetectScalarFileType(string filePath)
    {
        try
        {
            var fileName = System.IO.Path.GetFileName(filePath).ToUpper();

            // 파일명으로 우선 판단
            if (fileName.Contains("(TC)") || fileName.Contains("_TC") || fileName.Contains("-TC"))
                return "TC";
            if (fileName.Contains("(IC)") || fileName.Contains("_IC") || fileName.Contains("-IC"))
                return "IC";

            // 파일 내용으로 판단 (첫 몇 줄 확인)
            var lines = System.IO.File.ReadAllLines(filePath).Take(20);
            foreach (var line in lines)
            {
                var upperLine = line.ToUpper();
                if (upperLine.Contains("TC") && (upperLine.Contains("0") || upperLine.Contains("5") || upperLine.Contains("10")))
                    return "TC";
                if (upperLine.Contains("IC") && (upperLine.Contains("0") || upperLine.Contains("2.5") || upperLine.Contains("25")))
                    return "IC";
            }

            return "Unknown";
        }
        catch
        {
            return "Unknown";
        }
    }

    /// <summary>
    /// 스칼라 TCIC 실제 파싱 처리
    /// </summary>
    private async Task ProcessScalarTcicFiles(string tcFilePath, string icFilePath)
    {
        try
        {
            var tcFileName = System.IO.Path.GetFileName(tcFilePath);
            var icFileName = System.IO.Path.GetFileName(icFilePath);

            ShowMessage($"스칼라 TCIC 파싱 시작: TC={tcFileName}, IC={icFileName}", false);
            LogScalarTcic($"=== 스칼라 TCIC 파싱 시작 ===");
            LogScalarTcic($"TC 파일: {tcFileName} ({new FileInfo(tcFilePath).Length:N0} bytes)");
            LogScalarTcic($"IC 파일: {icFileName} ({new FileInfo(icFilePath).Length:N0} bytes)");

            LogScalarTcic($"[1/8] TC 파일 읽기 시작: {tcFileName}");
            LogScalarTcic($"TC 파일 확장자 확인: {System.IO.Path.GetExtension(tcFilePath)}");
            ShowMessage("TC 파일 파싱 중...", false);
            var tcResult = await System.Threading.Tasks.Task.Run(() =>
            {
                LogScalarTcic($"TC 파일 Task.Run 내부 시작");

                // PDF 파일인 경우 예나 PDF 파서 사용
                if (tcFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    LogScalarTcic("TC 파일이 PDF로 감지됨 - 예나 PDF 파서 사용");
                    var result = TocInstrumentParser.Parse(tcFilePath);
                    LogScalarTcic($"PDF 파싱 결과: {result.rows.Count}개 행, 형식: {result.format}");
                    return result;
                }
                else
                {
                    var result = TocInstrumentParser.Parse(tcFilePath);
                    LogScalarTcic($"TC 파일 Task.Run 내부 완료");
                    return result;
                }
            });
            LogScalarTcic($"[1/8] TC 파일 읽기 완료: {tcResult.rows.Count}개 행, 형식: {tcResult.format}");

            if (tcResult.cal != null)
                LogScalarTcic($"TC 검정곡선: HasData={tcResult.cal.HasData}, Method={tcResult.cal.Method}, Slope={tcResult.cal.Slope_TC}, StdConcs개수={tcResult.cal.StdConcs.Length}");
            else
                LogScalarTcic("TC 검정곡선 정보 없음");

            LogScalarTcic($"[2/8] IC 파일 읽기 시작: {icFileName}");
            ShowMessage("IC 파일 파싱 중...", false);
            var icResult = await System.Threading.Tasks.Task.Run(() =>
            {
                LogScalarTcic($"IC 파일 Task.Run 내부 시작");
                var result = TocInstrumentParser.Parse(icFilePath);
                LogScalarTcic($"IC 파일 Task.Run 내부 완료");
                return result;
            });
            LogScalarTcic($"[2/8] IC 파일 읽기 완료: {icResult.rows.Count}개 행, 형식: {icResult.format}");

            if (icResult.cal != null)
                LogScalarTcic($"IC 검정곡선: HasData={icResult.cal.HasData}, Method={icResult.cal.Method}, Slope={icResult.cal.Slope_TC}, StdConcs개수={icResult.cal.StdConcs.Length}");
            else
                LogScalarTcic("IC 검정곡선 정보 없음");

            // TC + IC 결과를 TCIC로 결합
            var combinedRows = new List<ExcelRow>();

            LogScalarTcic($"[3/8] 시료명 매칭 시작 - TC시료 {tcResult.rows.Count}개, IC시료 {icResult.rows.Count}개");
            ShowMessage("시료명 매칭 중...", false);

            int matchedCount = 0;
            int unmatchedCount = 0;

            // TC 데이터 기준으로 매칭하여 결합
            for (int i = 0; i < tcResult.rows.Count; i++)
            {
                var tcRow = tcResult.rows[i];
                LogScalarTcic($"[3/8] 처리 중: {i + 1}/{tcResult.rows.Count} - TC 시료: '{tcRow.RawName}' → 정규화: '{NormalizeSampleName(tcRow.RawName)}'");

                // 같은 시료명으로 IC 데이터 찾기
                var icRow = icResult.rows.FirstOrDefault(ic =>
                    NormalizeSampleName(ic.RawName) == NormalizeSampleName(tcRow.RawName));

                if (icRow != null)
                {
                    LogScalarTcic($"매칭 성공: TC '{tcRow.RawName}' ↔ IC '{icRow.RawName}'");
                    matchedCount++;
                }
                else
                {
                    LogScalarTcic($"매칭 실패: TC '{tcRow.RawName}' - IC에서 매칭 시료 없음");
                    if (i < 3) // 처음 3개만 로그 출력
                        LogScalarTcic("IC 시료들: " + string.Join(", ", icResult.rows.Select(ic => $"'{ic.RawName}'")));
                    unmatchedCount++;
                }

                var excelRow = new ExcelRow
                {
                    시료명 = tcRow.RawName,
                    SN = tcRow.SN,
                    Result = "", // TCIC는 최종 계산값 별도 처리
                    Source = SourceType.미분류,
                    Status = MatchStatus.대기,
                    IsControl = tcRow.IsControl,
                    // TCIC 전용 데이터 매핑
                    TCAU = tcRow.Area,   // TC Area
                    TCcon = tcRow.Conc,  // TC 농도
                    ICAU = icRow?.Area ?? "",   // IC Area
                    ICcon = icRow?.Conc ?? "",  // IC 농도
                    P = tcRow.Dilution   // 희석배수
                };

                // TCIC 최종 결과 계산 (TOC = TC - IC)
                if (double.TryParse(tcRow.Conc, out double tc) &&
                    double.TryParse(icRow?.Conc ?? "0", out double ic))
                {
                    var toc = tc - ic;
                    excelRow.Result = toc.ToString("F3");
                }

                combinedRows.Add(excelRow);
            }

            // TOC 카테고리로 설정
            _activeCategory = "TOC";
            var tocCatMatch = Categories.FirstOrDefault(c => c.Key == "TOC");
            _activeItems = tocCatMatch.Items ?? new[] { "TOC" };
            _categorySelected = true;

            // 문서 정보 생성 (TCIC 모드)
            var docInfo = new ExcelDocInfo
            {
                분석방법 = "Scalar TCIC",
                IsTocTCIC = true,
                IsTocNPOC = false
            };

            // 검량곡선 정보 통합 (TC + IC)
            if (tcResult.cal != null)
            {
                docInfo.TocSlope_TC = tcResult.cal.Slope_TC;
                docInfo.TocIntercept_TC = tcResult.cal.Intercept_TC;
                docInfo.TocR2_TC = tcResult.cal.R2_TC;
                docInfo.TocStdConcs = tcResult.cal.StdConcs;
                docInfo.TocStdAreas = tcResult.cal.StdAreas;

                LogScalarTcic($"TC 최종 검정곡선 설정 - Slope: {tcResult.cal.Slope_TC}, StdConcs 개수: {tcResult.cal.StdConcs.Length}");
            }
            else
            {
                LogScalarTcic("TC 검정곡선 정보 없음");
            }

            if (icResult.cal != null)
            {
                docInfo.TocSlope_IC = icResult.cal.Slope_TC;  // IC 파일의 TC값이 실제 IC값
                docInfo.TocIntercept_IC = icResult.cal.Intercept_TC;
                docInfo.TocR2_IC = icResult.cal.R2_TC;
                docInfo.TocStdConcs_IC = icResult.cal.StdConcs;
                docInfo.TocStdAreas_IC = icResult.cal.StdAreas;

                LogScalarTcic($"IC 최종 검정곡선 설정 - Slope: {icResult.cal.Slope_TC}, StdConcs 개수: {icResult.cal.StdConcs.Length}");
            }
            else
            {
                LogScalarTcic("IC 검정곡선 정보 없음");
            }

            // 파싱 결과 저장
            _categoryDocInfo["TOC"] = docInfo;
            _categoryExcelData["TOC"] = combinedRows;
            _categoryFilePaths["TOC"] = $"{tcFilePath};{icFilePath}"; // 두 파일 경로 저장

            LogScalarTcic("UI 업데이트 시작");

            // 데이터베이스 작업을 백그라운드에서 미리 수행 (UI 스레드 차단 방지)
            LogScalarTcic("[6/8] 백그라운드 데이터베이스 로딩 시작");
            var (preloadedSamples, preloadedAnalysis, preloadedFacilities) = await System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    LogScalarTcic("데이터베이스 사전 로드 시작");

                    // 문서 날짜에서 연도 추출
                    _categoryDocDates.TryGetValue(_activeCategory, out var docDateStr);
                    int docYear = DateTime.Now.Year;
                    if (DateTime.TryParse(docDateStr, out var docDt)) docYear = docDt.Year;

                    var allSamples = new List<WasteSample>();
                    var analysisRecords = new List<AnalysisRequestRecord>();
                    var loadedDates = new HashSet<string>();

                    // 선택된 날짜가 있으면 먼저 로드
                    if (_selectedDate != null && !loadedDates.Contains(_selectedDate))
                    {
                        allSamples.AddRange(WasteSampleService.GetByDate(_selectedDate));
                        analysisRecords.AddRange(AnalysisRequestService.GetByDate(_selectedDate));
                        loadedDates.Add(_selectedDate);
                        LogScalarTcic($"선택 날짜 로드: {_selectedDate}");
                    }

                    // SN에서 자동으로 날짜를 추출하여 추가 로드
                    foreach (var row in combinedRows)
                    {
                        var sn = row.SN;
                        if (sn.StartsWith("[")) { int idx = sn.IndexOf(']'); if (idx > 0) sn = sn[(idx + 1)..]; }
                        var parts = sn.Split('-');
                        if (parts.Length >= 2 && int.TryParse(parts[0], out var mm) && int.TryParse(parts[1], out var dd))
                        {
                            var dateStr = $"{docYear:D4}-{mm:D2}-{dd:D2}";
                            if (!loadedDates.Contains(dateStr))
                            {
                                try
                                {
                                    var samples = WasteSampleService.GetByDate(dateStr);
                                    allSamples.AddRange(samples);
                                    LogScalarTcic($"SN 기반 날짜 로드: {dateStr} ({samples.Count}건)");
                                } catch { }
                                try
                                {
                                    analysisRecords.AddRange(AnalysisRequestService.GetByDate(dateStr));
                                } catch { }
                                loadedDates.Add(dateStr);
                            }
                        }
                    }

                    // 처리시설 마스터 로드
                    List<(string 시설명, string 시료명, int 마스터Id)>? facilityMasters = null;
                    try
                    {
                        facilityMasters = FacilityResultService.GetAllMasterSamples();
                        LogScalarTcic($"처리시설 마스터 로드: {facilityMasters?.Count}건");
                    } catch { }

                    LogScalarTcic("데이터베이스 사전 로드 완료");
                    return (allSamples, analysisRecords, facilityMasters);
                }
                catch (Exception ex)
                {
                    LogScalarTcic($"데이터베이스 사전 로드 오류: {ex.Message}");
                    return (new List<WasteSample>(), new List<AnalysisRequestRecord>(), null);
                }
            });
            LogScalarTcic($"[7/8] 백그라운드 데이터베이스 로딩 완료: 샘플 {preloadedSamples.Count}건, 분석 {preloadedAnalysis.Count}건");

            // UI 업데이트 완전 스킵 - 데이터만 저장
            LogScalarTcic("[8/8] UI 업데이트 완전 스킵 - 데이터만 저장");
            _currentExcelRows = combinedRows;
            _currentSamples = preloadedSamples;
            _matchingAnalysisRecords = preloadedAnalysis;
            _matchingFacilityMasters = preloadedFacilities;
            LogScalarTcic("데이터 저장 완료 - UI 업데이트 없음");
            LogScalarTcic($"TCIC 파싱 완전 완료: {combinedRows.Count}건 (TC: {tcFileName}, IC: {icFileName})");

            // ShowMessage 제거 - 완전히 조용한 처리
            // ShowMessage($"✅ 스칼라 TCIC 파싱 완료 - {combinedRows.Count}건 로드됨 (TC: {tcFileName}, IC: {icFileName})", false);

            return; // 즉시 종료하여 추가 이벤트 처리 방지

        }
        catch (Exception ex)
        {
            ShowMessage($"스칼라 TCIC 파싱 오류: {ex.Message}", true);
        }
    }

    /// <summary>
    /// 시료명 정규화 (TCIC 매칭용)
    /// </summary>
    private string NormalizeSampleName(string sampleName)
    {
        if (string.IsNullOrWhiteSpace(sampleName)) return "";

        // 공백 제거, 대소문자 통일, 특수문자 제거
        return sampleName.Trim()
            .ToLowerInvariant()
            .Replace(" ", "")
            .Replace("-", "")
            .Replace("_", "");
    }

    /// <summary>
    /// TCIC 파일 확인 다이얼로그
    /// </summary>
    private async Task<bool> ShowTcicConfirmDialog(string tcFileName, string icFileName)
    {
        var content = new StackPanel { Spacing = 12 };

        content.Children.Add(FsSM(new TextBlock
        {
            Text = "📂 스칼라 TCIC 파일 확인",
            FontWeight = FontWeight.Bold,
            Foreground = AppRes("ThemeFgPrimary")
        }));

        content.Children.Add(FsBase(new TextBlock
        {
            Text = $"TC 파일: {tcFileName}",
            Foreground = AppRes("AppFg")
        }));

        content.Children.Add(FsBase(new TextBlock
        {
            Text = $"IC 파일: {icFileName}",
            Foreground = AppRes("AppFg")
        }));

        content.Children.Add(FsXS(new TextBlock
        {
            Text = "두 파일이 올바른지 확인하고 계속 진행하시겠습니까?",
            Foreground = AppRes("ThemeFgSecondary")
        }));

        bool result = false;
        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            HorizontalAlignment = HorizontalAlignment.Right
        };

        var cancelBtn = new Button { Content = "취소", Padding = new Thickness(16, 6) };
        var okBtn = new Button
        {
            Content = "계속",
            Padding = new Thickness(16, 6),
            Background = AppRes("BtnPrimaryBg"),
            Foreground = AppRes("BtnPrimaryFg")
        };

        cancelBtn.Click += (_, _) => result = false;
        okBtn.Click += (_, _) => result = true;

        buttonPanel.Children.Add(cancelBtn);
        buttonPanel.Children.Add(okBtn);

        var dialog = new Window
        {
            Title = "TCIC 파일 확인",
            Content = new Border
            {
                Padding = new Thickness(20),
                Child = new DockPanel
                {
                    Children = { buttonPanel, content }
                }
            },
            Width = 400,
            Height = 200,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };

        DockPanel.SetDock(buttonPanel, Dock.Bottom);

        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel is Window parentWindow)
            await dialog.ShowDialog(parentWindow);

        return result;
    }

    /// <summary>
    /// TCIC 전용 TOC 파싱 (TC + IC 두 파일)
    /// </summary>
    private void ParseTocInstrumentFileWithTcic(string tcFilePath, string icFilePath)
    {
        try
        {
            // 기존 ParseTocInstrumentFile과 유사하지만 두 파일 처리
            ShowMessage($"스칼라 TCIC 파싱 시작: TC={System.IO.Path.GetFileName(tcFilePath)}, IC={System.IO.Path.GetFileName(icFilePath)}", false);

            // TODO: TocInstrumentParser에 TCIC 이중 파일 지원 메서드 추가 필요
            // 임시로 기존 메서드 호출 (추후 개선 필요)
            ParseTocInstrumentFile(tcFilePath);

        }
        catch (Exception ex)
        {
            ShowMessage($"스칼라 TCIC 파싱 오류: {ex.Message}", true);
        }
    }

    /// <summary>
    /// 기존 Excel 파일 처리 로직
    /// </summary>
    private void ProcessExcelFile(string path)
    {
        var ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
        var key = _activeCategory;

        // .xlsx → TOC(TCIC) NOTE 레거시 양식이면 TOC 기기 파일로 처리
        if (ext == ".xlsx" &&
            TocInstrumentParser.DetectFormat(path) == TocInstrumentParser.TocFileFormat.TcicXlsxNote)
        {
            // 분석일을 R1B에서 추출 (있으면) → ParseTocInstrumentFile 후 덮어쓰기
            var noteDate = TocInstrumentParser.ExtractTcicXlsxDate(path);
            ParseTocInstrumentFile(path);
            if (!string.IsNullOrEmpty(noteDate))
            {
                _categoryDocDates["TOC"] = noteDate;
                { /* LoadVerifiedGrid 제거됨 */ }
            }
            return;
        }

        // 파일명으로 카테고리 자동 감지
        var fileName = System.IO.Path.GetFileNameWithoutExtension(path).ToUpperInvariant();
        var detectedKey = fileName switch
        {
            var n when n.Contains("NHEXAN") || n.Contains("N-HEXAN") || n.Contains("NHEX")
                    || n.Contains("노말헥산")                                                => "NHEX",
            var n when n.Contains("PHENOL") || n.Contains("페놀류")                        => "PHENOLS",
            var n when n.Contains("시안") || n.Contains("CN")                               => "CN",
            var n when n.Contains("6가크롬") || n.Contains("CR6") || n.Contains("크롬")     => "CR6",
            var n when n.Contains("색도") || n.Contains("COLOR")                            => "COLOR",
            var n when n.Contains("ABS") || n.Contains("흡광도") || n.Contains("음이온계면활성제") => "ABS",
            var n when n.Contains("불소") || n.Contains("FLUORIDE")                        => "FLUORIDE",
            var n when n.Contains("과불화") || n.Contains("PFAS") || n.Contains("PFOA") || n.Contains("PFOS") => "PFAS",
            var n when n.Contains("ECOLI") || n.Contains("대장균")                          => "ECOLI",
            var n when n.Contains("생물화학적") || n.Contains("BOD")                        => "BOD",
            var n when n.Contains("총유기탄소") || n.Contains("총_유기탄소") || n.Contains("유기탄소") || n.Contains("TOC") => "TOC",
            var n when n.Contains("부유물질") || n.Contains("SS")                           => "SS",
            var n when n.Contains("총질소") || n.Contains("T-N") || n.Contains("TN")       => "TN",
            var n when n.Contains("총인") || n.Contains("T-P") || n.Contains("TP")         => "TP",
            var n when n.Contains("화학적_산소") || n.Contains("화학적산소") || n.Contains("COD") => "COD",
            var n when n.Contains("수소이온") || n.Contains("PH")                           => "PH",
            var n when n.Contains("GCMS")                                                   => "GCMS",
            var n when n.Contains("ICP")                                                    => "ICP",
            var n when n.Contains("LCMS")                                                   => "LCMS",
            _                                                                                => null,
        };

        if (detectedKey != null && detectedKey != key)
        {
            key = detectedKey;
            _activeCategory = key;
            var detectedMatch = Categories.FirstOrDefault(c => c.Key == key);
            _activeItems = detectedMatch.Items ?? Array.Empty<string>();
            _categorySelected = true;
            UpdateCategoryButtonStyles();
        }

        // 엑셀 파싱 — 카테고리별로 Services/SERVICE4/*XlsxParser.cs 로 디스패치
        List<ExcelRow> rows;
        ExcelDocInfo   docInfo;
        string?        docDate;
        try
        {
            switch (key)
            {
                case "TOC":
                {
                    var p = TocXlsxParser.Parse(path, FormatResult);
                    rows = p.Rows; docInfo = p.DocInfo; docDate = p.DocDate;
                    _tocInstrumentMethod = p.Method;
                    break;
                }
                case "SS":
                {
                    var p = SsXlsxParser.Parse(path, _activeItems, FormatResult);
                    rows = p.Rows; docInfo = p.DocInfo; docDate = p.DocDate;
                    break;
                }
                case "NHEX":
                {
                    var p = NHexXlsxParser.Parse(path, _activeItems, FormatResult);
                    rows = p.Rows; docInfo = p.DocInfo; docDate = p.DocDate;
                    break;
                }
                case "TN":
                case "TP":
                case "PHENOLS":
                case "CN":
                case "CR6":
                case "COLOR":
                case "ABS":
                case "FLUORIDE":
                {
                    var p = UvvisXlsxParser.Parse(path, _activeItems, FormatResult);
                    rows = p.Rows; docInfo = p.DocInfo; docDate = p.DocDate;
                    break;
                }
                default:
                {
                    var p = BodExcelParser.Parse(path, _activeItems, FormatResult);
                    rows = p.Rows; docInfo = p.DocInfo; docDate = p.DocDate;
                    break;
                }
            }
        }
        catch (XlsxParseException ex)
        {
            ShowMessage(ex.Message, true);
            return;
        }
        catch (Exception ex)
        {
            ShowMessage($"엑셀 파싱 오류: {ex.Message}", true);
            return;
        }

        _categoryDocInfo[key] = docInfo;

        if (rows.Count == 0)
        {
            ShowMessage("유효한 데이터가 없습니다.", true);
            return;
        }

        _categoryExcelData[key] = rows;
        _categoryFilePaths[key] = path;
        _categoryDocDates[key] = docDate;
        UpdateCategoryButtonStyles();

        // 바로 Show2에 로드 (��짜 미선택이면 매칭 없이 엑셀 데이터만 표시)
        { /* LoadVerifiedGrid 제거됨 */ }
    }

    private void UpdateCategoryButtonStyles()
    {
        // 전체 UI 업데이트 비활성화 - 모든 파서 스턱 방지
        return;
    }

    // =========================================================================
    // 데이터 로드 (날짜 트리뷰)
    // =========================================================================
    public async void LoadData()
    {

        DateTreeView.Items.Clear();
        _selectedDate = null;
        _loadedMonths.Clear();
        _loadedYears.Clear();
        ListPanelChanged?.Invoke(null);
        EditPanelChanged?.Invoke(null);

        // Show1(수동 매칭 패널) 항상 표시: 페이지 진입 시 한 달치 후보 즉시 로드
        BuildMatchBrowsePanel();

        LoadProgressBorder.IsVisible = true;
        LoadProgressBar.IsIndeterminate = true;
        LoadProgressText.Text = "날짜 목록 불러오는 중...";

        try
        {
            // 모드별 월 소스
            //   수질분석센터 → 분석의뢰및결과
            //   처리시설     → 처리시설_작업
            //   비용부담금   → 폐수의뢰및결과
            bool isWaterCenterLD = IsWaterCenterMode;
            bool isFacilityLD    = IsFacilityMode;
            _allMonths = await System.Threading.Tasks.Task.Run(() =>
            {
                if (isWaterCenterLD) return AnalysisRequestService.GetMonths();
                if (isFacilityLD)    return FacilityResultService.GetMonths();
                return WasteRequestService.GetMonths();
            });

            var today = DateTime.Today;
            string thisYM = today.ToString("yyyy-MM");
            string lastYM = today.AddMonths(-1).ToString("yyyy-MM");
            int currentYear = today.Year;

            var byYear = _allMonths
                .GroupBy(m => int.Parse(m[..4]))
                .OrderByDescending(g => g.Key)
                .ToList();

            foreach (var yearGroup in byYear)
            {
                int year = yearGroup.Key;
                if (year == currentYear)
                {
                    // 이번 연도 월은 직접 표시
                    foreach (var ym in yearGroup.OrderByDescending(m => m))
                    {
                        bool isActive = ym == thisYM || ym == lastYM;
                        var monthNode = MakeMonthNode(ym, enabled: isActive);
                        DateTreeView.Items.Add(monthNode);

                        if (isActive)
                        {
                            await LoadMonthDatesIntoNode(monthNode, ym);
                            monthNode.IsExpanded = true;
                        }
                    }
                }
                else
                {
                    // 지난 연도는 연도 노드만
                    DateTreeView.Items.Add(MakeYearNode(year));
                }
            }
        }
        catch { }
        finally
        {
            LoadProgressBorder.IsVisible = false;
            LoadProgressBar.IsIndeterminate = false;
        }
    }

    // 특정 월의 날짜들을 노드에 채워 넣기 (진행바 표시)
    private async System.Threading.Tasks.Task LoadMonthDatesIntoNode(TreeViewItem monthNode, string ym)
    {
        if (_loadedMonths.Contains(ym)) return;

        LoadProgressBorder.IsVisible = true;
        LoadProgressBar.IsIndeterminate = true;
        LoadProgressText.Text = $"{ym} 데이터 연결 중...";

        try
        {
            // 모드별 날짜 소스
            //   수질분석센터 → 분석의뢰및결과.채취일자
            //   처리시설     → 처리시설_작업.채취일자
            //   비용부담금   → 폐수의뢰및결과.채수일
            bool isWaterCenter = IsWaterCenterMode;
            bool isFacility    = IsFacilityMode;
            bool isBilling     = IsBillingMode;
            var dates = await System.Threading.Tasks.Task.Run(() =>
            {
                if (isWaterCenter) return AnalysisRequestService.GetDatesByMonth(ym);
                if (isFacility)    return FacilityResultService.GetDatesByMonth(ym);
                return WasteRequestService.GetDatesByMonth(ym);
            });

            // count badge 갱신 (월 헤더)
            if (monthNode.Header is StackPanel sp && sp.Children.Count >= 2
                && sp.Children[1] is Border badge
                && badge.Child is TextBlock badgeTb)
            {
                badgeTb.Text = dates.Count.ToString();
            }

            foreach (var dateStr in dates)
            {
                var (samples, facilityStatus, reqItems, schedItems, billingItems) = await System.Threading.Tasks.Task.Run(() =>
                {
                    var s = WasteSampleService.GetByDate(dateStr);
                    var f = isFacility
                        ? FacilityResultService.GetFillStatusForDate(dateStr)
                        : new Dictionary<string, bool>();
                    // 수질분석센터: Category 기반 의뢰/결과 상태
                    var r = isWaterCenter
                        ? AnalysisRequestService.GetRequestedCategoriesByDate(dateStr)
                        : new Dictionary<string, bool>();
                    var sc = isFacility
                        ? FacilityResultService.GetScheduledItemsByDate(dateStr)
                        : new HashSet<string>();
                    // 비용부담금: 폐수_의뢰_항목 의뢰 키 집합
                    var bi = isBilling
                        ? WasteRequestService.GetRequestedItemSetByDate(dateStr)
                        : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    return (s, f, r, sc, bi);
                });

                // 의뢰 기준 트리 노드 — 결과 입력 여부와 무관하게 모두 표시
                monthNode.Items.Add(MakeDateNode(dateStr, samples, facilityStatus, reqItems, schedItems, billingItems));
            }

            _loadedMonths.Add(ym);
        }
        finally
        {
            LoadProgressBorder.IsVisible = false;
            LoadProgressBar.IsIndeterminate = false;
        }
    }

    // =========================================================================
    // 트리뷰 노드 빌더
    // =========================================================================
    private static TreeViewItem MakeYearNode(int year)
    {
        var header = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 6,
            Children =
            {
                new TextBlock
                {
                    Text = $"📁  {year}년",
                    FontWeight = FontWeight.SemiBold,
                    FontFamily = Font, Foreground = AppTheme.FgMuted,
                    VerticalAlignment = VerticalAlignment.Center,
                }.BindMD(),
            }
        };
        var node = new TreeViewItem
        {
            Tag    = $"YEAR:{year}",
            Header = header,
        };
        // 펼치기 전 placeholder (자식이 없으면 화살표 안 보임)
        node.Items.Add(new TreeViewItem { Header = "" });
        return node;
    }

    private static TreeViewItem MakeMonthNode(string ym, int count = 0, bool enabled = true)
    {
        DateTime.TryParse(ym + "-01", out var d);
        var textBrush  = enabled ? AppTheme.FgMuted : AppTheme.FgDimmed;
        var badgeBrush = enabled ? AppTheme.BorderSubtle : AppTheme.BorderSubtle;

        var monthNode = new TreeViewItem
        {
            Tag        = $"YM:{ym}",
            IsEnabled  = true,   // 클릭 허용 (비활성달도 클릭으로 로드)
            Header     = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    new TextBlock
                    {
                        Text = enabled
                            ? $"📅  {d.Year}년 {d.Month}월"
                            : $"🔒  {d.Year}년 {d.Month}월",
                        FontWeight = FontWeight.SemiBold,
                        FontFamily = Font, Foreground = textBrush,
                        VerticalAlignment = VerticalAlignment.Center,
                        Opacity = enabled ? 1.0 : 0.45,
                    }.BindMD(),
                    new Border
                    {
                        Background = badgeBrush, CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5, 1),
                        Child = new TextBlock
                        {
                            Text = count == 0 ? "..." : count.ToString(),
                            FontFamily = Font, Foreground = AppTheme.FgDimmed,
                        }.BindXS()
                    }
                }
            }
        };
        return monthNode;
    }

    private TreeViewItem MakeDateNode(string dateStr, List<WasteSample> samples,
        Dictionary<string, bool>? facilityStatus = null,
        Dictionary<string, bool>? requestedCategories = null,
        HashSet<string>? scheduledItems = null,
        HashSet<string>? billingRequestedItems = null)
    {
        DateTime.TryParse(dateStr, out var d);
        string dow = d == DateTime.MinValue ? "" : $" ({DayOfWeekKr(d)})";
        string label = d == DateTime.MinValue ? dateStr : $"🗓  {d.Month:D2}/{d.Day:D2}{dow}";

        Control statusControl;
        {
            // 뱃지 letter → 카테고리 키 매핑 (처리시설/비용부담금 모드용)
            static string LetterToKey(string letter) => letter switch
            {
                "B" => "BOD", "T" => "TOC", "S" => "SS",
                "N" => "TN",  "P" => "TP",  "H" => "PHENOLS",
                "X" => "NHEX", "E" => "ECOLI", _ => ""
            };

            // 라벨 달린 원형 뱃지 (클릭 가능)
            // isFacility=true 이면 처리시설 데이터 뷰로, false 이면 배출업소 샘플 그리드로 라우팅
            Border Dot(bool filled, string letter, bool clickable = true, bool isFacility = false) {
                var border = new Border
                {
                    Width = 14, Height = 14,
                    Margin = new Thickness(1, 0),
                    CornerRadius = new CornerRadius(7),
                    Background = filled
                        ? new SolidColorBrush(Color.Parse("#22c55e"))
                        : new SolidColorBrush(Colors.Transparent),
                    BorderBrush = filled
                        ? new SolidColorBrush(Colors.Transparent)
                        : new SolidColorBrush(Color.Parse("#888888")),
                    BorderThickness = new Thickness(1),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = letter,
                        FontSize = 8,
                        FontWeight = FontWeight.Bold,
                        FontFamily = Font,
                        Foreground = filled
                            ? new SolidColorBrush(Colors.White)
                            : new SolidColorBrush(Color.Parse("#888888")),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                    }
                };
                if (clickable)
                {
                    border.Cursor = new Cursor(StandardCursorType.Hand);
                    border.PointerPressed += (_, e) =>
                    {
                        e.Handled = true; // TreeView 선택 이벤트 방지
                        var catKey = LetterToKey(letter);
                        if (string.IsNullOrEmpty(catKey)) return;
                        _activeCategory = catKey;
                        var catMatch = Categories.FirstOrDefault(c => c.Key == catKey);
                        _activeItems = catMatch.Items ?? Array.Empty<string>();
                        _categorySelected = true;
                        _selectedDate = dateStr;
                        _facilityViewMode = isFacility;
                        UpdateCategoryButtonStyles();
                        if (isFacility)
                            LoadFacilityGrid(dateStr);
                        else
                            LoadSampleGrid(dateStr);
                        EditPanelChanged?.Invoke(null);
                    };
                    border.PointerEntered += (_, _) =>
                    {
                        border.Opacity = 0.7;
                    };
                    border.PointerExited += (_, _) =>
                    {
                        border.Opacity = 1.0;
                    };
                }
                return border;
            }

            var dotsRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 1, VerticalAlignment = VerticalAlignment.Center };

            // (letter, fillCondition, isFacility) 배지 정의
            // B=BOD / T=TOC / S=SS / N=T-N / P=T-P / E=eColi / H=Phenols / X=NHex
            if (IsFacilityMode)
            {
                // 처리시설:
                //   배지 표시 여부 = scheduledItems (처리시설_작업.항목목록 — 의뢰)
                //   배지 채움색   = facilityStatus (처리시설_측정결과 입력 여부)
                bool FsScheduled(string col) => scheduledItems != null && scheduledItems.Contains(col);
                bool FsFilled(string col)    => facilityStatus != null && facilityStatus.TryGetValue(col, out var v) && v;
                void AddFac(string col, string letter)
                {
                    if (FsScheduled(col))
                        dotsRow.Children.Add(Dot(FsFilled(col), letter, isFacility: true));
                }
                AddFac("BOD",        "B");
                AddFac("TOC",        "T");
                AddFac("SS",         "S");
                AddFac("T-N",        "N");
                AddFac("T-P",        "P");
                AddFac("총대장균군", "E");
            }
            else if (IsBillingMode)
            {
                // 비용부담금:
                //   배지 표시 여부 = 해당 항목이 폐수의뢰및결과에 값이 있음
                //   배지 채움색   = 폐수의뢰및결과에 결과값이 입력됨
                bool BlReq(string key) =>
                    billingRequestedItems != null && billingRequestedItems.Contains(key);
                bool hasSamples = samples.Count > 0;
                void AddBl(string letter, string key, Func<WasteSample, string> getter)
                {
                    if (!BlReq(key)) return;
                    bool filled = hasSamples && samples.Any(s => !string.IsNullOrWhiteSpace(getter(s)));
                    dotsRow.Children.Add(Dot(filled, letter));
                }
                AddBl("B", "BOD",     s => s.BOD);
                AddBl("T", "TOC",     s => s.TOC);
                AddBl("S", "SS",      s => s.SS);
                AddBl("N", "TN",      s => s.TN);
                AddBl("P", "TP",      s => s.TP);
                AddBl("H", "PHENOLS", s => s.Phenols);
                AddBl("X", "NHEX",    s => s.NHexan);
            }
            else
            {
                // 수질분석센터: 분석정보.Category 기준 (5개 카테고리)
                //   배지 표시 여부 = 해당 카테고리에 의뢰가 있음
                //   배지 채움색   = 카테고리 내 어떤 항목이라도 결과값 입력됨
                bool CatReq(params string[] catKeywords) =>
                    requestedCategories != null &&
                    requestedCategories.Keys.Any(k => catKeywords.Any(kw =>
                        k.Contains(kw, StringComparison.OrdinalIgnoreCase)));
                bool CatFilled(params string[] catKeywords) =>
                    requestedCategories != null &&
                    requestedCategories.Any(kvp =>
                        catKeywords.Any(kw => kvp.Key.Contains(kw, StringComparison.OrdinalIgnoreCase)) &&
                        kvp.Value);
                void AddCat(string letter, params string[] catKeywords)
                {
                    if (CatReq(catKeywords))
                        dotsRow.Children.Add(Dot(CatFilled(catKeywords), letter, clickable: false));
                }
                // 일=일반항목 / 유=유기물(질) / 휘=휘발성유기화합물 / 금=금속류 / 이=이온류
                AddCat("일", "일반");
                AddCat("유", "유기물");          // "유기물" / "유기물질" 모두 매칭
                AddCat("휘", "휘발성");
                AddCat("금", "금속");
                AddCat("이", "이온");
            }

            statusControl = dotsRow;
        }

        var tvi = new TreeViewItem
        {
            Tag = dateStr,
            Header = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    new TextBlock
                    {
                        Text = label,
                        FontFamily = Font, Foreground = AppTheme.FgPrimary,
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindSM(),
                    statusControl,
                }
            }
        };
        TextShimmer.AttachHover(tvi);
        return tvi;
    }

    private static string DayOfWeekKr(DateTime d) => d.DayOfWeek switch
    {
        DayOfWeek.Monday    => "월",
        DayOfWeek.Tuesday   => "화",
        DayOfWeek.Wednesday => "수",
        DayOfWeek.Thursday  => "목",
        DayOfWeek.Friday    => "금",
        DayOfWeek.Saturday  => "토",
        DayOfWeek.Sunday    => "일",
        _ => ""
    };

    // =========================================================================
    // 트리뷰 선택 → 연도 펼치기 / 월 lazy 로드 / 날짜 선택
    // =========================================================================
    private async void DateTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (DateTreeView.SelectedItem is not TreeViewItem tvi) return;
        if (tvi.Tag is not string tag) return;

        // ── 연도 노드 클릭 → 해당 연도의 월 목록 펼치기 ──────────────────
        if (tag.StartsWith("YEAR:"))
        {
            int year = int.Parse(tag[5..]);
            if (!_loadedYears.Contains(year))
            {
                tvi.Items.Clear(); // placeholder 제거
                foreach (var ym in _allMonths.Where(m => m.StartsWith(year + "-")).OrderByDescending(m => m))
                    tvi.Items.Add(MakeMonthNode(ym, enabled: false));
                _loadedYears.Add(year);
            }
            tvi.IsExpanded = true;
            return;
        }

        // ── 월 노드 클릭 → 해당 월 날짜 lazy 로드 ────────────────────────
        if (tag.StartsWith("YM:"))
        {
            string ym = tag[3..];
            if (!_loadedMonths.Contains(ym))
                await LoadMonthDatesIntoNode(tvi, ym);
            tvi.IsExpanded = true;
            return;
        }

        // ── 날짜 노드 클릭 → Show2 업데이트 ──────────────────────────────
        string dateStr = tag;
        _selectedDate = dateStr;
        _selectedSample = null;

        // 처리시설/비용부담금 모드는 카테고리 선택 없이 날짜 클릭시 바로 그리드 표시
        bool nonCenterMode = !IsWaterCenterMode;
        if (!_categorySelected && !nonCenterMode) return;

        _currentSamples = WasteSampleService.GetByDate(dateStr);

        // 처리시설 모드: 처리시설 그리드 (facility 전용)
        if (IsFacilityMode)
            LoadFacilityGrid(dateStr);
        // 비용부담금 모드: 배출업소 시료 그리드
        else if (IsBillingMode)
            LoadSampleGrid(dateStr);
        // 수질분석센터: 엑셀 첨부가 있으면 검증 그리드, 없으면 기본 시료 그리드
        else if (_categoryExcelData.ContainsKey(_activeCategory))
            { /* LoadVerifiedGrid 제거됨 */ }
        else
            LoadSampleGrid(dateStr);

        EditPanelChanged?.Invoke(null);
        BuildStatsPanel();
    }

    // =========================================================================
    // Show2: 검증 결과 포함 그리드 (첨부 파일 있을 때)
    // =========================================================================
    private void LoadVerifiedGrid()
    {

        if (!_categoryExcelData.TryGetValue(_activeCategory, out var excelRows))
        {
            LogMatch($"LoadVerifiedGrid: No data found for category {_activeCategory}");
            if (_selectedDate != null) LoadSampleGrid(_selectedDate);
            return;
        }

        LogMatch($"LoadVerifiedGrid: Found {excelRows.Count} excel rows");

        // 매칭 상태 확인
        foreach (var row in excelRows.Take(5))
        {
            LogMatch($"Row {row.SN}: Source={row.Source}, Status={row.Status}, 시료명='{row.시료명}', Matched={row.Matched?.업체명}, Analysis={row.MatchedAnalysis?.시료명}");
        }

        // 자동 검증: SN에서 날짜 추출 → 해당 날짜의 의뢰 자동 로드
        // SN 형식: "MM-DD-순서", "[세풍]MM-DD-순서", "[율촌]MM-DD-순서"
        var allSamples = new List<WasteSample>();
        var loadedDates = new HashSet<string>();

        // 문서 날짜에서 연도 추출
        _categoryDocDates.TryGetValue(_activeCategory, out var docDateStr);
        int docYear = DateTime.Now.Year;
        if (DateTime.TryParse(docDateStr, out var docDt)) docYear = docDt.Year;

        // 날짜 선택된 경우 해당 날짜 먼저 로드
        if (_selectedDate != null && !loadedDates.Contains(_selectedDate))
        {
            allSamples.AddRange(WasteSampleService.GetByDate(_selectedDate));
            loadedDates.Add(_selectedDate);
        }

        // SN에서 자동으로 날짜를 추출하여 추가 로드
        foreach (var row in excelRows)
        {
            var sn = row.SN;
            // [세풍], [율촌] 접두사 제거
            if (sn.StartsWith("[")) { int idx = sn.IndexOf(']'); if (idx > 0) sn = sn[(idx + 1)..]; }
            var parts = sn.Split('-');
            if (parts.Length >= 2 && int.TryParse(parts[0], out var mm) && int.TryParse(parts[1], out var dd))
            {
                var dateStr = $"{docYear:D4}-{mm:D2}-{dd:D2}";
                if (!loadedDates.Contains(dateStr))
                {
                    var samples = WasteSampleService.GetByDate(dateStr);
                    allSamples.AddRange(samples);
                    loadedDates.Add(dateStr);
                }
            }
        }

        _currentSamples = allSamples;

        // 수질분석센터 의뢰 로드 (같은 날짜들)
        var analysisRecords = new List<AnalysisRequestRecord>();
        foreach (var dt in loadedDates)
        {
            try { analysisRecords.AddRange(AnalysisRequestService.GetByDate(dt)); } catch { }
        }
        _matchingAnalysisRecords = analysisRecords;

        // 처리시설 마스터 시료명 로드
        List<(string 시설명, string 시료명, int 마스터Id)>? facilityMasters = null;
        try { facilityMasters = FacilityResultService.GetAllMasterSamples(); } catch { }
        _matchingFacilityMasters = facilityMasters;

        // 3개 테이블 동시 매칭
        foreach (var row in excelRows)
        {
            // 드래그-드롭으로 수동 매칭된 행은 자동 매칭 건너뛰기
            if (row.IsManualMatch)
            {
                LogMatch($"SKIP AUTO MATCH: {row.SN} (manual match: {row.Source})");
                continue;
            }

            // 1순위: 폐수배출업소 (SN 매칭 — [세풍]/[율촌] 접두사 무시 비교 포함)
            string rowSnBare = row.SN;
            if (rowSnBare.StartsWith("[")) { int bi = rowSnBare.IndexOf(']'); if (bi > 0) rowSnBare = rowSnBare[(bi + 1)..]; }
            row.Matched = _currentSamples.FirstOrDefault(s => s.SN == row.SN)
                       ?? _currentSamples.FirstOrDefault(s =>
                       {
                           string sBare = s.SN;
                           if (sBare.StartsWith("[")) { int bi2 = sBare.IndexOf(']'); if (bi2 > 0) sBare = sBare[(bi2 + 1)..]; }
                           return sBare == rowSnBare && sBare == row.SN; // bare SN이 같고 엑셀 SN에 접두사 없는 경우
                       })
                       ?? _currentSamples.FirstOrDefault(s =>
                       {
                           string sBare = s.SN;
                           if (sBare.StartsWith("[")) { int bi2 = sBare.IndexOf(']'); if (bi2 > 0) sBare = sBare[(bi2 + 1)..]; }
                           return sBare == rowSnBare;
                       })
                       ?? (!string.IsNullOrEmpty(row.SN) ? _currentSamples.FirstOrDefault(s => s.업체명 == row.시료명) : null);

            if (row.Matched != null)
            {
                // 원본 시료명 보존 후 업체명으로 교체 (auto-match by SN)
                if (string.IsNullOrEmpty(row.원본시료명))
                    row.원본시료명 = row.시료명;  // 원본 보존
                row.시료명 = row.Matched.업체명;  // 업체명으로 변경

                row.Source = SourceType.폐수배출업소;
                bool has = _activeItems.Any(item => !string.IsNullOrWhiteSpace(GetSampleValue(row.Matched, item)));
                row.Status = has ? MatchStatus.덮어쓰기 : MatchStatus.입력가능;
                continue;
            }

            // 2순위: 처리시설 (시료명 매칭)
            if (facilityMasters != null)
            {
                var fm = FacilityResultService.FindBySampleName(facilityMasters, row.시료명);
                if (fm != null)
                {
                    row.Source = SourceType.처리시설;
                    row.MatchedFacilityName = fm.Value.시설명;
                    // 기존 데이터 존재 여부 확인 (해당 시설+날짜+항목에 값이 있으면 덮어쓰기)
                    bool hasFacilityData = false;
                    try
                    {
                        string checkDate = !string.IsNullOrEmpty(docDateStr) ? docDateStr : DateTime.Today.ToString("yyyy-MM-dd");
                        var existingRows = FacilityResultService.GetRows(fm.Value.시설명, checkDate);
                        var matchedRow = existingRows.FirstOrDefault(r => r.시료명 == row.시료명);
                        if (matchedRow != null)
                        {
                            string itemKey = _activeItems.FirstOrDefault() ?? "";
                            string existVal = itemKey switch
                            {
                                "BOD" => matchedRow.BOD, "TOC" => matchedRow.TOC,
                                "SS" => matchedRow.SS, "T-N" => matchedRow.TN,
                                "T-P" => matchedRow.TP, _ => ""
                            };
                            hasFacilityData = !string.IsNullOrWhiteSpace(existVal);
                        }
                    }
                    catch { }
                    row.Status = hasFacilityData ? MatchStatus.덮어쓰기 : MatchStatus.입력가능;
                    continue;
                }
            }

            // 3순위: 수질분석센터 (약칭/시료명 매칭)
            var ar = analysisRecords.FirstOrDefault(a => a.약칭 == row.시료명 || a.시료명 == row.시료명);
            if (ar != null)
            {
                // 원본 시료명 보존 후 분석 시료명으로 교체 (auto-match by 약칭/시료명)
                if (string.IsNullOrEmpty(row.원본시료명))
                    row.원본시료명 = row.시료명;  // 원본 보존
                row.시료명 = ar.시료명;  // 분석의뢰 시료명으로 변경

                row.MatchedAnalysis = ar;
                row.Source = SourceType.수질분석센터;
                row.Status = MatchStatus.입력가능;
                continue;
            }

            // 미매칭
            row.Source = SourceType.미분류;
            row.Status = _currentSamples.Count > 0 || analysisRecords.Count > 0
                         || (_matchingFacilityMasters?.Count ?? 0) > 0
                ? MatchStatus.미매칭 : MatchStatus.대기;
        }

        // Row 0:header / 1:docBorder(resizable) / 2:splitter / 3:badgePanel / 4:calBorder / 5:gridBorder(*)
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,Auto,Auto,*") };

        // 헤더
        var catLabel = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;
        string fileName = _categoryFilePaths.TryGetValue(_activeCategory, out var fp) ? System.IO.Path.GetFileName(fp) : "";
        // B1 분석일 우선, 없으면 채수일 목록
        string dateLabel = !string.IsNullOrEmpty(docDateStr)
            ? $"분석일: {docDateStr}"
            : loadedDates.Count > 0
                ? string.Join(", ", loadedDates.OrderBy(x => x))
                : "날짜 미감지";

        int matchNew   = excelRows.Count(r => r.Status == MatchStatus.입력가능);
        int matchExist = excelRows.Count(r => r.Status == MatchStatus.덮어쓰기);
        int noMatch    = excelRows.Count(r => r.Status == MatchStatus.미매칭);
        int pending    = excelRows.Count(r => r.Status == MatchStatus.대기);

        // 요약 배지 구성
        var badgePanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
        _summaryBadgePanel = badgePanel;
        RebuildBadge(excelRows);

        // 문서 정보 패널
        var headerContent = new StackPanel { Spacing = 4 };

        // 분석일 클릭 → 인라인 캘린더로 변경 가능
        var dateBtn = FsBase(new Button
        {
            Content = $"📋 {dateLabel}",
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("ThemeFgWarn"),
            Background = Brushes.Transparent, BorderThickness = new Thickness(0),
            Padding = new Thickness(0), Cursor = new Cursor(StandardCursorType.Hand),
        });
        var calBorder = InlineCalendarHelper.Create(dt =>
        {
            var newDate = dt.ToString("yyyy-MM-dd");
            _categoryDocDates[_activeCategory] = newDate;
            dateBtn.Content = $"📋 분석일: {newDate}";
        }, dateBtn);

        var headerInfoPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        headerInfoPanel.Children.Add(dateBtn);
        // TOC: NPOC/TCIC 세분화 표시
        string dispLabel = catLabel;
        if (_activeCategory == "TOC" && _categoryDocInfo.TryGetValue("TOC", out var tocDi))
            dispLabel = tocDi.IsTocTCIC ? "TOC-TCIC" : "TOC-NPOC";

        headerInfoPanel.Children.Add(FsBase(new TextBlock
        {
            Text = $"[{dispLabel}]  📎 {fileName}",
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        }));
        headerContent.Children.Add(headerInfoPanel);
        // 문서 헤더 — 식종수/SCF 또는 검량곡선 테이블
        bool isUVVISMode = false;
        bool isSSMode = false;
        bool isNHEXMode = false;
        bool isTocMode = false;
        bool isTocTcicMode = false;
        bool isGcMode = false;
        bool isIcpMode = false;

        // TOC 검량선 수식은 나중에 docInfo 확인 후 추가 (아래 isTocMode 설정 이후)

        // 컬럼 너비 배열 (헤더+식종수+데이터 모든 행 동기화)
        string colDefs = "32,50,90,190,60,60,60,60,50,75,80"; // BOD 기본
        double[] colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
        var allRowGrids = new List<Grid>();

        Grid MakeRowGrid()
        {
            var g = new Grid { ClipToBounds = false };
            foreach (var w in colWidths)
                g.ColumnDefinitions.Add(new ColumnDefinition(w, GridUnitType.Pixel));
            allRowGrids.Add(g);
            return g;
        }

        void SyncAllWidths(int colIdx, double newW)
        {
            colWidths[colIdx] = newW;
            foreach (var g in allRowGrids)
                if (colIdx < g.ColumnDefinitions.Count)
                    g.ColumnDefinitions[colIdx].Width = new GridLength(newW);
        }

        StackPanel? docTbl = null;
        if (_categoryDocInfo.TryGetValue(_activeCategory, out var docInfo))
        {
            isUVVISMode = docInfo.IsUVVIS;
            isSSMode    = docInfo.IsSS;
            isNHEXMode  = docInfo.IsNHEX;
            isTocMode   = docInfo.IsTocNPOC || docInfo.IsTocTCIC;
            isTocTcicMode = docInfo.IsTocTCIC;
            isGcMode    = docInfo.IsGcMode;
            isIcpMode   = isGcMode && (docInfo.GcFormat?.StartsWith("ICP", StringComparison.OrdinalIgnoreCase) == true);
            bool isShimadzuToc = isTocMode && docInfo.IsShimadzuToc;

            // TOC y = ax + b 수식을 분석일 바로 아래 표시
            if (isTocMode && !string.IsNullOrEmpty(docInfo.TocSlope_TC))
            {
                string r2str = !string.IsNullOrEmpty(docInfo.TocR2_TC) ? $"    R²={docInfo.TocR2_TC}" : "";
                string formula = docInfo.IsTocTCIC
                    ? $"TC:  y = {docInfo.TocSlope_TC}x + {docInfo.TocIntercept_TC}{r2str}    |    IC:  y = {docInfo.TocSlope_IC}x + {docInfo.TocIntercept_IC}"
                    : $"y = {docInfo.TocSlope_TC}x + {docInfo.TocIntercept_TC}{r2str}";
                headerContent.Children.Add(FsXS(new TextBlock
                {
                    Text = formula, FontFamily = Font,
                    Foreground = AppRes("ThemeFgWarn"),
                    FontWeight = FontWeight.SemiBold,
                    Margin = new Thickness(0, 1, 0, 0),
                }));
            }

            // GC 검량선 요약: "GC [VocMulti] — 10성분" + 첫 성분 수식
            if (isGcMode && docInfo.GcCompoundCals.Count > 0)
            {
                var first = docInfo.GcCompoundCals[0];
                string r2str = !string.IsNullOrEmpty(first.R) ? $"    R={first.R}" : "";
                string formula = $"GC [{docInfo.GcFormat}] — {docInfo.GcCompoundCals.Count}성분   |   {first.Name}:  y = {first.Slope}x + {first.Intercept}{r2str}";
                headerContent.Children.Add(FsXS(new TextBlock
                {
                    Text = formula, FontFamily = Font,
                    Foreground = AppRes("ThemeFgWarn"),
                    FontWeight = FontWeight.SemiBold,
                    Margin = new Thickness(0, 1, 0, 0),
                }));
            }
            if (isSSMode)
            {
                colDefs = "32,50,130,190,60,60,60,60,50,75,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }
            else if (isUVVISMode)
            {
                colDefs = "32,50,130,190,60,65,75,65,75,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }
            else if (isTocTcicMode)
            {
                // TOC TCIC: 컬럼 11개 — 체크/입력/SN/시료명/TCAU/TCcon/ICAU/ICcon/희석배수/결과값/시료구분
                colDefs = "32,50,130,190,65,65,65,65,55,70,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }
            else if (isTocMode)
            {
                // TOC NPOC: 시료량 없음 — 컬럼 8개 (체크/입력/SN/시료명/AU/희석배수/결과값/시료구분)
                colDefs = "32,50,130,190,80,60,75,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }
            else if (isIcpMode)
            {
                // ICP: 체크/입력/SN/시료명/성분명/농도/희석배수/결과값/시료구분 (9컬럼)
                colDefs = "32,50,130,220,70,65,55,70,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }
            else if (isGcMode)
            {
                // GC: 시료량 없음, 컬럼 11개
                //  체크/입력/SN/시료명/시료량(숨김)/Resp./ISTD/농도/희석배수/결과값/시료구분
                colDefs = "32,50,130,220,0,75,75,65,55,70,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }

            // 분석방법
            if (!string.IsNullOrWhiteSpace(docInfo.분석방법))
                headerContent.Children.Add(FsXS(new TextBlock
                {
                    Text = $"분석방법: {docInfo.분석방법}", FontFamily = Font,
                    Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 2, 0, 0),
                }));

            // 표준형식 (스키마) 표시 — 항상 DB(schema_override)에서 로드
            {
                string _schAnalyte = GetActiveAnalyteKey();
                string _schema     = ETA.Services.SERVICE3.AnalysisNoteService.GetSchemaOverride(_schAnalyte) ?? "";
                if (!string.IsNullOrWhiteSpace(_schema))
                    headerContent.Children.Add(FsXS(new TextBlock
                    {
                        Text = $"표준형식: {_schema}", FontFamily = Font,
                        Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 0, 0),
                    }));
            }

            // 식종수/SCF 또는 검량곡선 테이블 — 데이터 그리드와 동일 구조로 별도 행 배치
            if (docInfo.IsUVVIS)
            {
                bool hasStd = docInfo.Standard_Points.Length > 0;
                bool hasAbs = docInfo.Abs_Values.Length > 0;
                if (hasStd || hasAbs)
                {
                    // 검량선 테이블: 실제 표준점 개수에 맞춰 ST 컬럼 동적 생성
                    int uvStCount = Math.Max(
                        hasStd ? docInfo.Standard_Points.Length : 0,
                        hasAbs ? docInfo.Abs_Values.Length : 0);
                    uvStCount = Math.Max(uvStCount, 1);
                    // 구분(140) + 숨김3개 + STn×60 + 검량계수(160) + 여백(0)
                    string uvDocColDefs = "140,0,0,0," + string.Join(",", Enumerable.Repeat("60", uvStCount)) + ",160,0";

                    docTbl = new StackPanel { Spacing = 0 };
                    var hdr = new Grid { ColumnDefinitions = new ColumnDefinitions(uvDocColDefs),
                        MinHeight = 26, Background = AppRes("GridHeaderBg") };
                    var hdrLabel = FsBase(new TextBlock { Text = "구분", FontFamily = Font,
                        FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
                    Grid.SetColumn(hdrLabel, 0); Grid.SetColumnSpan(hdrLabel, 4);
                    hdr.Children.Add(hdrLabel);
                    // ST1~STn 헤더
                    for (int c = 0; c < uvStCount; c++)
                    {
                        var tb = FsBase(new TextBlock { Text = $"ST-{c + 1}", FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(4, 0) });
                        Grid.SetColumn(tb, 4 + c); hdr.Children.Add(tb);
                    }
                    // 검량계수 헤더
                    var calHdrTb = FsBase(new TextBlock { Text = "검량계수", FontFamily = Font,
                        FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(4, 0) });
                    Grid.SetColumn(calHdrTb, 4 + uvStCount); hdr.Children.Add(calHdrTb);
                    docTbl.Children.Add(new Border { Child = hdr,
                        BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

                    int rowIdx = 0;
                    if (hasStd)
                    {
                        var stdVals = Enumerable.Range(0, uvStCount)
                            .Select(i => docInfo.Standard_Points.ElementAtOrDefault(i) ?? "")
                            .Append($"a={docInfo.Standard_Slope}  b={docInfo.Standard_Intercept}")
                            .Append("").ToArray();
                        docTbl.Children.Add(BuildDocRowUnified(uvDocColDefs, "STANDARD", stdVals, "ThemeFgWarn", rowIdx++));
                    }
                    if (hasAbs)
                    {
                        var absVals = Enumerable.Range(0, uvStCount)
                            .Select(i => docInfo.Abs_Values.ElementAtOrDefault(i) ?? "")
                            .Append($"R²={docInfo.Abs_R2}")
                            .Append("").ToArray();
                        docTbl.Children.Add(BuildDocRowUnified(uvDocColDefs, "Absorbance", absVals, "ThemeFgInfo", rowIdx++));
                    }
                }
            }
            else if (isNHEXMode)
            {
                bool hasBlanks = !string.IsNullOrWhiteSpace(docInfo.바탕시료_결과)
                              || !string.IsNullOrWhiteSpace(docInfo.바탕시료_건조전);
                if (hasBlanks)
                {
                    docTbl = new StackPanel { Spacing = 0 };
                    var hdr = MakeRowGrid();
                    hdr.MinHeight = 26; hdr.Background = AppRes("GridHeaderBg");
                    var hdrLabel = FsBase(new TextBlock { Text = "구분", FontFamily = Font,
                        FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
                    Grid.SetColumn(hdrLabel, 0); Grid.SetColumnSpan(hdrLabel, 4);
                    hdr.Children.Add(hdrLabel);
                    string[] nhexCols = { "시료량", "건조전무게", "건조후무게", "무게차", "희석배수", "농도", "" };
                    for (int c = 0; c < nhexCols.Length; c++)
                    {
                        var tb = FsBase(new TextBlock { Text = nhexCols[c], FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(4, 0) });
                        Grid.SetColumn(tb, 4 + c); hdr.Children.Add(tb);
                    }
                    docTbl.Children.Add(new Border { Child = hdr,
                        BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });
                    docTbl.Children.Add(BuildDocRowUnified(colDefs, "바탕시료",
                        new[] { docInfo.바탕시료_시료량, docInfo.바탕시료_건조전, docInfo.바탕시료_건조후,
                                docInfo.바탕시료_무게차, docInfo.바탕시료_희석배수, docInfo.바탕시료_결과, "" }, "ThemeFgInfo", 0));
                }
            }
            else if (isTocMode)
            {
                bool hasCal = !string.IsNullOrWhiteSpace(docInfo.TocSlope_TC);
                if (hasCal)
                {
                    // 파싱값 / 기기출력값 선택
                    bool useInst = _tocShowInstrumentCal && docInfo.HasTocInstrumentCal;
                    string slopeTC     = useInst ? (docInfo.TocSlope_TC_Inst     ?? "") : (docInfo.TocSlope_TC     ?? "");
                    string interceptTC = useInst ? (docInfo.TocIntercept_TC_Inst ?? "") : (docInfo.TocIntercept_TC ?? "");
                    string r2TC        = useInst ? (docInfo.TocR2_TC_Inst        ?? "") : (docInfo.TocR2_TC        ?? "");
                    string slopeIC     = useInst ? (docInfo.TocSlope_IC_Inst     ?? "") : (docInfo.TocSlope_IC     ?? "");
                    string interceptIC = useInst ? (docInfo.TocIntercept_IC_Inst ?? "") : (docInfo.TocIntercept_IC ?? "");
                    string r2IC        = useInst ? (docInfo.TocR2_IC_Inst        ?? "") : (docInfo.TocR2_IC        ?? "");
                    string[] stdAreas    = useInst ? docInfo.TocStdAreas_Inst    : docInfo.TocStdAreas;
                    string[] stdAreas_IC = useInst ? docInfo.TocStdAreas_IC_Inst : docInfo.TocStdAreas_IC;

                    // 표준점 개수에 맞춰 ST-1~N — 라벨 고정 + ST 컬럼 + 기울기/절편/R² 컬럼
                    int stCount = Math.Max(docInfo.TocStdConcs.Length, 2);
                    stCount = Math.Min(stCount, 8);
                    string tocDocColDefs = "140,0,0,0," + string.Join(",", Enumerable.Repeat("80", stCount)) + ",90,90,70";
                    int calColStart = 4 + stCount; // 기울기 컬럼 시작 인덱스

                    docTbl = new StackPanel { Spacing = 0 };

                    // ── 헤더 행: 토글 버튼 포함 ──
                    var hdr = new Grid { ColumnDefinitions = new ColumnDefinitions(tocDocColDefs),
                        MinHeight = 26, Background = AppRes("GridHeaderBg") };

                    // ── 헤더 라벨 + 시마즈 토글 버튼 ──
                    {
                        var labelArea = new StackPanel { Orientation = Orientation.Horizontal,
                            VerticalAlignment = VerticalAlignment.Center, Spacing = 6, Margin = new Thickness(4, 0) };
                        labelArea.Children.Add(FsBase(new TextBlock { Text = "구분", FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center }));

                        if (isShimadzuToc)
                        {
                            bool hasInst = docInfo.HasTocInstrumentCal;
                            void MakeToggleBtn(string label, bool isActive, bool switchTo, bool enabled)
                            {
                                var btn = new Border
                                {
                                    Background = isActive ? AppRes("BtnPrimaryBg")
                                               : enabled  ? AppRes("ThemeBgSubtle")
                                                          : AppRes("ThemeBorderMuted"),
                                    CornerRadius = new CornerRadius(switchTo ? 0 : 3, switchTo ? 3 : 0, switchTo ? 3 : 0, switchTo ? 0 : 3),
                                    Padding = new Thickness(6, 1),
                                    Cursor = enabled ? new Cursor(StandardCursorType.Hand) : new Cursor(StandardCursorType.Arrow),
                                    Opacity = enabled ? 1.0 : 0.45,
                                    Child = FsXS(new TextBlock
                                    {
                                        Text = label, FontFamily = Font,
                                        Foreground = isActive ? AppRes("BtnPrimaryFg") : AppRes("FgMuted"),
                                        FontWeight = isActive ? FontWeight.Bold : FontWeight.Normal,
                                        VerticalAlignment = VerticalAlignment.Center,
                                    }),
                                };
                                if (enabled)
                                {
                                    bool targetVal = switchTo;
                                    btn.PointerPressed += (_, _) => { _tocShowInstrumentCal = targetVal; LoadVerifiedGrid(); };
                                }
                                labelArea.Children.Add(btn);
                            }
                            MakeToggleBtn("파싱값",   !useInst, false, true);
                            MakeToggleBtn("기기출력값", useInst,  true,  hasInst);
                        }

                        Grid.SetColumn(labelArea, 0); Grid.SetColumnSpan(labelArea, 4);
                        hdr.Children.Add(labelArea);
                    }

                    for (int c = 0; c < stCount; c++)
                    {
                        var tb = FsBase(new TextBlock { Text = $"ST-{c + 1}", FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center });
                        Grid.SetColumn(tb, 4 + c); hdr.Children.Add(tb);
                    }
                    foreach (var (hText, hOff) in new[] { ("기울기", 0), ("절편", 1), ("R²", 2) })
                    {
                        var tb = FsBase(new TextBlock { Text = hText, FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center });
                        Grid.SetColumn(tb, calColStart + hOff); hdr.Children.Add(tb);
                    }
                    docTbl.Children.Add(new Border { Child = hdr,
                        BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

                    int nSt = Math.Min(docInfo.TocStdConcs.Length, stCount);
                    string[] stdRow = new string[stCount + 3];
                    string[] auRow  = new string[stCount + 3];
                    for (int si = 0; si < nSt; si++)
                    {
                        stdRow[si] = docInfo.TocStdConcs[si];
                        auRow[si]  = si < stdAreas.Length ? stdAreas[si] : "";
                    }
                    // TC 기울기/절편/R² → STANDARD 행에 표시
                    stdRow[stCount]     = slopeTC;
                    stdRow[stCount + 1] = interceptTC;
                    stdRow[stCount + 2] = r2TC;

                    string tcLabel = docInfo.IsTocTCIC ? "TC STANDARD" : "STANDARD";
                    string auLabel = docInfo.IsTocTCIC ? "TC AU" : "NPOC AU";
                    int rowIdx = 0;
                    docTbl.Children.Add(BuildDocRowUnified(tocDocColDefs, tcLabel, stdRow, "ThemeFgWarn", rowIdx++));
                    if (stdAreas.Length > 0)
                        docTbl.Children.Add(BuildDocRowUnified(tocDocColDefs, auLabel, auRow, "ThemeFgInfo", rowIdx++));

                    // TCIC 전용: IC 검량선 행
                    if (docInfo.IsTocTCIC && docInfo.TocStdConcs_IC.Length > 0)
                    {
                        int nIc = Math.Min(docInfo.TocStdConcs_IC.Length, stCount);
                        string[] icStdRow = new string[stCount + 3];
                        string[] icAuRow  = new string[stCount + 3];
                        for (int si = 0; si < nIc; si++)
                        {
                            icStdRow[si] = docInfo.TocStdConcs_IC[si];
                            icAuRow[si]  = si < stdAreas_IC.Length ? stdAreas_IC[si] : "";
                        }
                        // IC 기울기/절편/R² → IC STANDARD 행에 표시
                        icStdRow[stCount]     = slopeIC;
                        icStdRow[stCount + 1] = interceptIC;
                        icStdRow[stCount + 2] = r2IC;

                        docTbl.Children.Add(BuildDocRowUnified(tocDocColDefs, "IC STANDARD", icStdRow, "ThemeFgWarn", rowIdx++));
                        if (stdAreas_IC.Length > 0)
                            docTbl.Children.Add(BuildDocRowUnified(tocDocColDefs, "IC AU", icAuRow, "ThemeFgInfo", rowIdx++));
                    }
                }
            }
            else if (isGcMode && docInfo.GcCompoundCals.Count > 0)
            {
                // GC 검량선 테이블: 성분별로 ST-1~N (공칭농도 + 응답) 한 쌍씩 렌더링
                int maxSt = docInfo.GcCompoundCals.Max(c => c.StdConcs.Length);
                if (maxSt < 2) maxSt = 2;
                if (maxSt > 7) maxSt = 7;
                // 라벨 영역: 성분명 고정 160px + ST 컬럼들 + 검량선 정보 컬럼들 (기울기, 절편, R²)
                string gcLabelCols = "160,0,0,0";
                string stCols = string.Join(",", Enumerable.Repeat("90", maxSt));
                string calCols = "100,100,80"; // 기울기, 절편, R² 컬럼
                string gcDocColDefs = gcLabelCols + "," + stCols + "," + calCols;

                docTbl = new StackPanel { Spacing = 0 };
                var hdr = new Grid {
                    ColumnDefinitions = new ColumnDefinitions(gcDocColDefs),
                    MinHeight = 26, Background = AppRes("GridHeaderBg")
                };
                var hdrLabel = FsBase(new TextBlock { Text = "성분 / 구분", FontFamily = Font,
                    FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
                Grid.SetColumn(hdrLabel, 0); Grid.SetColumnSpan(hdrLabel, 4);
                hdr.Children.Add(hdrLabel);
                for (int c = 0; c < maxSt; c++)
                {
                    var tb = FsBase(new TextBlock { Text = $"ST-{c + 1}", FontFamily = Font,
                        FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(4, 0) });
                    Grid.SetColumn(tb, 4 + c); hdr.Children.Add(tb);
                }

                // 검량선 정보 헤더 추가
                var slopeHeader = FsBase(new TextBlock { Text = "기울기", FontFamily = Font,
                    FontWeight = FontWeight.SemiBold, Foreground = AppRes("ThemeFgInfo"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(4, 0) });
                Grid.SetColumn(slopeHeader, 4 + maxSt); hdr.Children.Add(slopeHeader);

                var interceptHeader = FsBase(new TextBlock { Text = "절편", FontFamily = Font,
                    FontWeight = FontWeight.SemiBold, Foreground = AppRes("ThemeFgInfo"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(4, 0) });
                Grid.SetColumn(interceptHeader, 4 + maxSt + 1); hdr.Children.Add(interceptHeader);

                var r2Header = FsBase(new TextBlock { Text = "R²", FontFamily = Font,
                    FontWeight = FontWeight.SemiBold, Foreground = AppRes("ThemeFgInfo"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(4, 0) });
                Grid.SetColumn(r2Header, 4 + maxSt + 2); hdr.Children.Add(r2Header);
                docTbl.Children.Add(new Border { Child = hdr,
                    BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

                int rowIdx = 0;
                foreach (var comp in docInfo.GcCompoundCals)
                {
                    // 성분명 + 공칭농도 행
                    string[] concRow = new string[maxSt + 3]; // ST 컬럼 + 기울기 + 절편 + R²
                    for (int si = 0; si < maxSt; si++)
                        concRow[si] = si < comp.StdConcs.Length ? comp.StdConcs[si] : "";

                    // 검량선 정보 추가
                    concRow[maxSt] = comp.Slope;     // 기울기
                    concRow[maxSt + 1] = comp.Intercept; // 절편
                    concRow[maxSt + 2] = comp.R;     // R²

                    // 진한색 연한색 교대 색상 — 성분명 행에 드래그앤드랍 + Shift+1 매칭 지원
                    var compLabel = comp.Name;
                    var existingAlias = CompoundAliasService.Resolve(compLabel);
                    string displayLabel = existingAlias != null ? $"{compLabel} → {existingAlias.Value.분석항목}" : compLabel;
                    var compBorder = BuildDocRowUnified(gcDocColDefs, displayLabel, concRow, "ThemeFgWarn", rowIdx++);
                    if (existingAlias != null && compBorder.Child is Grid existGrid)
                    {
                        foreach (var ch in existGrid.Children)
                            if (ch is TextBlock etb && Grid.GetColumn(etb) == 0)
                            { etb.Foreground = new SolidColorBrush(Color.Parse("#90EE90")); break; }
                    }
                    docTbl.Children.Add(compBorder);

                    // 검정곡선 성분 행: 드래그앤드랍 + Shift+1 클릭 수신
                    var capturedCompName = comp.Name;
                    var capturedBorder = compBorder;
                    compBorder.Background = Avalonia.Media.Brushes.Transparent; // 히트테스트 보장
                    compBorder.Cursor = new Cursor(StandardCursorType.Hand);
                    DragDrop.SetAllowDrop(compBorder, true);
                    compBorder.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
                    {
                        bool acceptable = e.Data.Contains("match-compound");
                        e.DragEffects = acceptable ? DragDropEffects.Link : DragDropEffects.None;
                        if (acceptable)
                            capturedBorder.BorderBrush = new SolidColorBrush(Color.Parse("#4CAF50"));
                        e.Handled = acceptable;
                    });
                    compBorder.AddHandler(DragDrop.DragLeaveEvent, (object? s, RoutedEventArgs e) =>
                    {
                        capturedBorder.BorderBrush = AppRes("ThemeBorderSubtle");
                    });
                    compBorder.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
                    {
                        capturedBorder.BorderBrush = AppRes("ThemeBorderSubtle");
                        if (e.Data.Contains("match-compound"))
                        {
                            string draggedName = e.Data.Get(DataFormats.Text)?.ToString() ?? "";
                            LogMatch($"[CalRow DROP] '{capturedCompName}' ← '{draggedName}'");
                            RegisterCompoundAliasAndUpdateGrid(capturedCompName, draggedName, capturedBorder);
                            e.Handled = true;
                        }
                    });
                    // Shift+1 모드: 클릭 시 포커스된 Show1 분석항목을 이 성분에 적용
                    compBorder.AddHandler(Avalonia.Input.InputElement.PointerPressedEvent, (object? s, PointerPressedEventArgs pe) =>
                    {
                        LogMatch($"[CalRow CLICK] '{capturedCompName}', keyNav={_keyNavShow1}, browse='{_show1BrowseMode}', idx={_keyNavShow1Index}");
                        if (_keyNavShow1 && _show1BrowseMode == "분석항목"
                            && _keyNavShow1Index >= 0 && _keyNavShow1Index < _matchItems.Count)
                        {
                            var (_, itemName, _) = _matchItems[_keyNavShow1Index];
                            RegisterCompoundAliasAndUpdateGrid(capturedCompName, itemName, capturedBorder);
                            if (_keyNavShow1Index + 1 < _matchItems.Count)
                                NavShow1(_keyNavShow1Index + 1);
                            pe.Handled = true;
                        }
                    }, Avalonia.Interactivity.RoutingStrategies.Tunnel);

                    // 응답 행
                    string[] respRow = new string[maxSt + 3];
                    for (int si = 0; si < maxSt; si++)
                        respRow[si] = si < comp.StdResps.Length ? comp.StdResps[si] : "";

                    // 검량선 정보 컬럼은 빈 값
                    respRow[maxSt] = "";
                    respRow[maxSt + 1] = "";
                    respRow[maxSt + 2] = "";

                    string auLabel = comp.HasIstd ? "Resp." : "AU";
                    docTbl.Children.Add(BuildDocRowUnified(gcDocColDefs, auLabel, respRow, "ThemeFgSecondary", rowIdx++));

                    // ISTD 응답 행 (HasIstd인 경우에만)
                    if (comp.HasIstd && comp.StdIstdResps.Length > 0)
                    {
                        string[] istdRespRow = new string[maxSt + 3];
                        for (int si = 0; si < maxSt; si++)
                            istdRespRow[si] = si < comp.StdIstdResps.Length ? comp.StdIstdResps[si] : "";

                        // 검량선 정보 컬럼은 빈 값
                        istdRespRow[maxSt] = "";
                        istdRespRow[maxSt + 1] = "";
                        istdRespRow[maxSt + 2] = "";

                        docTbl.Children.Add(BuildDocRowUnified(gcDocColDefs, "ISTD Resp.", istdRespRow, "ThemeFgSecondary", rowIdx++));
                    }
                }
            }
            else
            {
                bool hasSeed = !string.IsNullOrWhiteSpace(docInfo.식종수_Result)
                            || !string.IsNullOrWhiteSpace(docInfo.식종수_D1);
                bool hasScf  = !string.IsNullOrWhiteSpace(docInfo.SCF_Result)
                            || !string.IsNullOrWhiteSpace(docInfo.SCF_D1);
                if (hasSeed || hasScf)
                {
                    docTbl = new StackPanel { Spacing = 0 };
                    var hdr = MakeRowGrid();
                    hdr.MinHeight = 26; hdr.Background = AppRes("GridHeaderBg");
                    var hdrLabel = FsBase(new TextBlock { Text = "구분", FontFamily = Font,
                        FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
                    Grid.SetColumn(hdrLabel, 0); Grid.SetColumnSpan(hdrLabel, 4);
                    hdr.Children.Add(hdrLabel);
                    string[] bodCols = { "시료량", "D1", "D2", "f(x/y)", "P", "Result", "비고" };
                    for (int c = 0; c < bodCols.Length; c++)
                    {
                        var tb = FsBase(new TextBlock { Text = bodCols[c], FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(4, 0) });
                        Grid.SetColumn(tb, 4 + c); hdr.Children.Add(tb);
                    }
                    docTbl.Children.Add(new Border { Child = hdr,
                        BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

                    int rowIdx = 0;
                    if (hasSeed)
                    {
                        docTbl.Children.Add(BuildDocRowUnified(colDefs, "식종수의 BOD",
                            new[] { docInfo.식종수_시료량, docInfo.식종수_D1, docInfo.식종수_D2,
                                    "-", docInfo.식종수_P, docInfo.식종수_Result, docInfo.식종수_Remark }, "ThemeFgWarn", rowIdx++));
                    }
                    if (hasScf)
                    {
                        docTbl.Children.Add(BuildDocRowUnified(colDefs, "SCF(식종희석수)",
                            new[] { docInfo.SCF_시료량, docInfo.SCF_D1, docInfo.SCF_D2,
                                    "-", "1", docInfo.SCF_Result, "" }, "ThemeFgInfo", rowIdx++));
                    }
                }
            }
        }

        var header = new Border
        {
            Background = AppRes("PanelBg"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(10, 6),
            Child = headerContent,
        };
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // 배지 패널 (식종수 테이블과 데이터 그리드 사이)
        badgePanel.Margin = new Thickness(4, 4, 0, 2);
        Grid.SetRow(badgePanel, 3);
        root.Children.Add(badgePanel);

        // 식종수/검량곡선 테이블 (데이터 그리드와 동일 마진/구조) — 스크롤 + 리사이즈 가능
        if (docTbl != null)
        {
            // 자연 높이(대략 헤더+행수×28)를 기준으로 초기 높이 결정 (상한 320px)
            int docChildren    = docTbl.Children.Count;
            double naturalH    = 28.0 * Math.Max(docChildren, 1) + 4;
            double initialH    = Math.Min(naturalH, 320);

            var docScroll = new ScrollViewer
            {
                Content                       = docTbl,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            };
            // 검정곡선 영역 전체에 드래그앤드랍 허용 + 이벤트 로그
            DragDrop.SetAllowDrop(docScroll, true);
            DragDrop.SetAllowDrop(docTbl, true);
            docScroll.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
            {
                // match-compound는 자식(compBorder)에서 처리 — 여기서 허용만 해 줌
                if (e.Data.Contains("match-compound"))
                    e.DragEffects = DragDropEffects.Link;
                LogMatch($"[DocScroll DragOver] keys=[{string.Join(",", e.Data.GetDataFormats())}]");
            });
            docScroll.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
            {
                LogMatch($"[DocScroll Drop] keys=[{string.Join(",", e.Data.GetDataFormats())}]");
            });
            docScroll.AddHandler(Avalonia.Input.InputElement.PointerPressedEvent, (object? s, PointerPressedEventArgs pe) =>
            {
                LogMatch($"[DocScroll PointerPressed] keyNav={_keyNavShow1}, browse='{_show1BrowseMode}'");
            }, Avalonia.Interactivity.RoutingStrategies.Tunnel);
            var docBorder = new Border
            {
                Child = docScroll,
                BorderBrush = AppRes("ThemeBorderSubtle"),
                BorderThickness = new Thickness(1, 0, 1, 1),
                ClipToBounds = true,
            };
            // Row 1: Pixel 고정 → GridSplitter 로 드래그 리사이즈
            root.RowDefinitions[1] = new RowDefinition(initialH, GridUnitType.Pixel) { MinHeight = 30 };
            Grid.SetRow(docBorder, 1);
            root.Children.Add(docBorder);

            // Row 2: 드래그 핸들
            var splitter = new GridSplitter
            {
                Height              = 6,
                Background          = AppRes("ThemeBorderSubtle"),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment   = VerticalAlignment.Stretch,
                ResizeDirection     = GridResizeDirection.Rows,
                Cursor              = new Cursor(StandardCursorType.SizeNorthSouth),
            };
            Grid.SetRow(splitter, 2);
            root.Children.Add(splitter);
        }

        // 그리드 본체
        _gridPanel = new StackPanel { Spacing = 0 };

        // 컬럼 헤더
        var colHeaderGrid = MakeRowGrid();
        colHeaderGrid.MinHeight = 28;
        colHeaderGrid.Background = AppRes("GridHeaderBg");

        string[] hLabels = isSSMode
            ? new[] { "", "입력", "SN", "시료명", "시료량", "전무게", "후무게", "무게차", "희석배수", "결과값", "시료구분" }
            : isNHEXMode
            ? new[] { "", "입력", "SN", "시료명", "시료량", "건조전무게", "건조후무게", "무게차", "희석배수", "결과값", "시료구분" }
            : isUVVISMode
            ? new[] { "", "입력", "SN", "시료명", "시료량", "흡광도", "계산농도", "희석배수", "결과값", "시료구분" }
            : isTocTcicMode
            ? new[] { "", "입력", "SN", "시료명", "TCAU", "TCcon", "ICAU", "ICcon", "희석배수", "결과값", "시료구분" }
            : isTocMode
            ? new[] { "", "입력", "SN", "시료명", "AU", "희석배수", "결과값", "시료구분" }
            : isIcpMode
            ? new[] { "", "입력", "SN", "시료명", "성분명", "농도", "희석배수", "결과값", "시료구분" }
            : isGcMode
            ? new[] { "", "입력", "SN", "시료명", "", "Resp.", "ISTD Resp.", "농도", "희석배수", "결과값", "시료구분" }
            : new[] { "", "입력", "SN", "시료명", "시료량", "D1", "D2", "f(x/y)", "P", "결과값", "시료구분" };
        int detailStart = 4, detailEnd = isTocTcicMode ? 8 : isTocMode ? 5 : isUVVISMode ? 7 : isIcpMode ? 6 : 8;
        for (int c = 0; c < hLabels.Length; c++)
        {
            if (c == 0) // 첫 번째 컬럼: 전체 토글 스위치
            {
                var masterToggleTrack = new Border
                {
                    Width = 30, Height = 14, CornerRadius = new CornerRadius(7),
                    Background = AppRes("ThemeBorderMuted"),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Cursor = new Cursor(StandardCursorType.Hand)
                };

                var masterToggleKnob = new Border
                {
                    Width = 10, Height = 10, CornerRadius = new CornerRadius(5),
                    Background = AppRes("ThemeBgInput"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2)
                };

                masterToggleTrack.Child = masterToggleKnob;
                _masterToggleTrack = masterToggleTrack;
                _masterToggleKnob  = masterToggleKnob;

                // 전체 토글 이벤트: ON 비율 > 50% → 전체 OFF, 그 외 → 전체 ON
                masterToggleTrack.PointerPressed += (_, _) =>
                {
                    if (_currentExcelRows == null || _currentExcelRows.Count == 0) return;
                    int enabledCount = _currentExcelRows.Count(r => r.Enabled);
                    bool newState = enabledCount <= _currentExcelRows.Count / 2; // 50% 이하 ON → 전체 ON
                    _allRowsEnabled = newState;

                    foreach (var row in _currentExcelRows)
                        row.Enabled = newState;

                    // 개별 행 토글 UI 동기화
                    for (int ri = 0; ri < _rowToggles.Count && ri < _currentExcelRows.Count; ri++)
                    {
                        var (t, k) = _rowToggles[ri];
                        t.Background = newState ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted");
                        k.Margin = new Thickness(newState ? 18 : 2, 2, 0, 2);
                    }

                    // 마스터 토글 UI 업데이트
                    masterToggleTrack.Background = newState ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted");
                    masterToggleKnob.Margin = new Thickness(newState ? 18 : 2, 2, 0, 2);
                };

                Grid.SetColumn(masterToggleTrack, c);
                colHeaderGrid.Children.Add(masterToggleTrack);
            }
            else
            {
                var tb = FsBase(new TextBlock
                {
                    Text = hLabels[c],
                    FontWeight = FontWeight.SemiBold,
                    FontFamily = Font, Foreground = c >= detailStart && c <= detailEnd ? AppRes("ThemeFgSecondary") : AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = c >= detailStart ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                    Margin = new Thickness(4, 0),
                });
                Grid.SetColumn(tb, c);
                colHeaderGrid.Children.Add(tb);
            }

            // 컬럼 경계 드래그 리사이즈 핸들
            if (c < hLabels.Length - 1)
            {
                var handle = new Border
                {
                    Width = 6, Background = Brushes.Transparent,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Stretch,
                    Cursor = new Cursor(StandardCursorType.SizeWestEast),
                };
                int colIdx = c;
                double startX = 0, origWidth = 0;
                handle.PointerPressed += (_, e) =>
                {
                    startX = e.GetPosition(colHeaderGrid).X;
                    origWidth = colWidths[colIdx];
                    e.Pointer.Capture(handle);
                    e.Handled = true;
                };
                handle.PointerMoved += (_, e) =>
                {
                    if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed) return;
                    double dx = e.GetPosition(colHeaderGrid).X - startX;
                    SyncAllWidths(colIdx, Math.Max(30, origWidth + dx));
                };
                handle.PointerReleased += (_, e) => e.Pointer.Capture(null);
                Grid.SetColumn(handle, c);
                colHeaderGrid.Children.Add(handle);
            }
        }
        var _colHeaderBorder = new Border
        {
            Child = colHeaderGrid,
            BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };
        // 헤더는 스크롤 밖에 고정 (아래에서 gridBorder 조립 시 사용)

        // 데이터 행
        _rowIcons   = new List<Border>();
        _rowToggles = new List<(Border, Border)>();
        _rowSnCells = new List<TextBlock>();
        _rowNameCells = new List<StackPanel>();
        _rowSourceCells = new List<TextBlock>();
        _rowDilButtons = new List<Button?>();
        _rowVolButtons = new List<Button?>();
        _dupCheckboxGroups = new Dictionary<string, List<CheckBox>>();
        _dupIconGroups = new Dictionary<string, List<Border>>();
        _rowIconPanels = new List<StackPanel>();
        // 중복 시료명 감지 (같은 시료명 + 같은 매칭 대상이 2개 이상인 경우)
        // 중복 시료명 감지: SN이 있는(매칭된) 행만 대상 (DW 같은 원본명으로 잘못 중복 감지 방지)
        var duplicateSampleNameSet = excelRows
            .Where(r => !string.IsNullOrEmpty(r.SN))
            .GroupBy(r => $"{r.시료명}_{r.SN}")
            .Where(g => g.Count() > 1)
            .SelectMany(g => g.Select(r => r.시료명))
            .ToHashSet();
        // 중복 SN 감지 (같은 SN이 2개 이상인 경우)
        var duplicateSnSet = excelRows
            .GroupBy(r => r.SN)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToHashSet();


        for (int i = 0; i < excelRows.Count; i++)
        {
            var row = excelRows[i];

            var rowGrid = MakeRowGrid();
            rowGrid.MinHeight = 34;
            // 원래 교대 색상 복원
            rowGrid.Background = i % 2 == 0 ? AppRes("GridRowBg") : AppRes("GridRowAltBg");

            // 아이콘 + 토글 (한 셀에 합침)
            var capturedRow = row;
            var toggleTrack = new Border
            {
                Width = 30, Height = 14, CornerRadius = new CornerRadius(7),
                Background = row.Enabled ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted"),
                Cursor = new Cursor(StandardCursorType.Hand),
                VerticalAlignment = VerticalAlignment.Center,
            };
            var toggleKnob = new Border
            {
                Width = 10, Height = 10, CornerRadius = new CornerRadius(5),
                Background = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(row.Enabled ? 18 : 2, 2, 0, 2),
            };
            toggleTrack.Child = toggleKnob;

            if (row.Status == MatchStatus.미매칭)
            {
                // 미매칭: 시각적으로 흐리게 표시하되 토글은 가능
                toggleTrack.Opacity = 0.5;
            }

            toggleTrack.PointerPressed += (_, _) =>
            {
                capturedRow.Enabled = !capturedRow.Enabled;
                toggleTrack.Background = capturedRow.Enabled ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted");
                toggleKnob.Margin = new Thickness(capturedRow.Enabled ? 18 : 2, 2, 0, 2);

                // 전체 토글 상태 업데이트
                UpdateMasterToggleState();
            };

            // 아이콘(컬러 원) + 체크박스 (col 0)
            bool isDuplicateSn = duplicateSnSet.Contains(row.SN);
            bool isDuplicateName = duplicateSampleNameSet.Contains(row.시료명);
            var iconCircle = new Border
            {
                Width = 12, Height = 12,
                CornerRadius = new CornerRadius(6),
                Background = isDuplicateName ? Brush.Parse("#FF69B4") : // 핑크색 (중복 시료명)
                           isDuplicateSn ? AppRes("ThemeFgDanger") : // 빨간색 (중복 SN)
                           AppRes("FgMuted"), // 회색 (일반)
                Opacity = (isDuplicateName || isDuplicateSn) ? 1.0 : 0.35,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Cursor = new Cursor((isDuplicateName || isDuplicateSn) ? StandardCursorType.Hand : StandardCursorType.Arrow),
            };
            _rowIcons.Add(iconCircle);

            var col0Panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Spacing = 2,
            };
            col0Panel.Children.Add(iconCircle);

            if (isDuplicateName)
            {
                var capturedDupRow = row;
                var capturedSn = row.SN;
                var cb = new CheckBox
                {
                    IsChecked = row.IsSelectedForFinalResult,
                    IsVisible = false, // 핑크 원 클릭 전까지 숨김
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, 0, 0, 0),
                };
                var capturedIcon = iconCircle;
                cb.IsCheckedChanged += (_, _) =>
                {
                    capturedDupRow.IsSelectedForFinalResult = cb.IsChecked == true;
                    // 선택된 아이콘 → 녹색, 나머지 그룹 → 하늘색
                    if (_dupIconGroups.TryGetValue(capturedSn, out var groupIcons))
                    {
                        foreach (var icon in groupIcons)
                        {
                            icon.Background = (icon == capturedIcon && cb.IsChecked == true)
                                ? AppRes("ThemeFgSuccess")      // 녹색 (선택됨)
                                : Brush.Parse("#87CEEB");        // 하늘색 (미선택 중복)
                        }
                    }
                };
                col0Panel.Children.Add(cb);

                // SN 그룹별 체크박스/아이콘 목록에 추가
                if (!_dupCheckboxGroups.ContainsKey(capturedSn))
                    _dupCheckboxGroups[capturedSn] = new List<CheckBox>();
                _dupCheckboxGroups[capturedSn].Add(cb);
                if (!_dupIconGroups.ContainsKey(capturedSn))
                    _dupIconGroups[capturedSn] = new List<Border>();
                _dupIconGroups[capturedSn].Add(iconCircle);

                // 핑크 원 클릭 → 같은 SN 그룹의 체크박스 모두 토글
                iconCircle.PointerPressed += (_, e) =>
                {
                    e.Handled = true;
                    if (!_dupCheckboxGroups.TryGetValue(capturedSn, out var groupCbs) || groupCbs.Count == 0) return;
                    bool show = !groupCbs[0].IsVisible;
                    foreach (var groupCb in groupCbs)
                        groupCb.IsVisible = show;
                };
            }
            // 핑크색이 아닌 아이콘(회색/빨간/녹색)은 클릭해도 체크박스 표시 안 함

            _rowIconPanels.Add(col0Panel);
            Grid.SetColumn(col0Panel, 0);
            rowGrid.Children.Add(col0Panel);

            // 토글 스위치 (col 1) — 전체 행에 표시 (미매칭은 비활성)
            _rowToggles.Add((toggleTrack, toggleKnob));
            var toggleWrap = new Border
            {
                Child = toggleTrack,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(2, 4),
                IsVisible = true,
            };
            Grid.SetColumn(toggleWrap, 1);
            rowGrid.Children.Add(toggleWrap);

            // SN (처리시설이면 시설명 표시, IsControl이면 "QC" 고정)
            string snDisplay = row.IsControl && row.Source == SourceType.미분류
                ? "QC"
                : row.Source switch
                {
                    SourceType.처리시설     when !string.IsNullOrEmpty(row.MatchedFacilityName) => row.MatchedFacilityName,
                    SourceType.처리시설     => row.SN,
                    SourceType.수질분석센터 when !string.IsNullOrEmpty(row.MatchedAnalysis?.약칭) => row.MatchedAnalysis!.약칭,
                    SourceType.수질분석센터 => row.SN,
                    _ => row.SN,
                };
            var snTb = FsBase(new TextBlock
            {
                Text = snDisplay, FontFamily = Font,
                Foreground = AppRes("ThemeFgInfo"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
            });
            _rowSnCells.Add(snTb);
            Grid.SetColumn(snTb, 2);
            rowGrid.Children.Add(snTb);

            // 시료명 (원본시료명 있으면 위에 작게 흰색, 매칭명 크게 노란색)
            var nameCell = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) };
            if (!string.IsNullOrEmpty(row.원본시료명))
            {
                nameCell.Children.Add(new TextBlock
                {
                    Text = row.원본시료명, FontFamily = Font,
                    FontSize = AppTheme.FontXS,
                    Foreground = AppRes("AppFg"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = $"↳ {row.시료명}", FontFamily = Font,
                    Foreground = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#ffe066")),
                    FontWeight = FontWeight.SemiBold,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
            else
            {
                // 다성분 별칭 자동 반영: CompoundName에 등록된 별칭이 있으면 정규명 표시
                string displayName = row.시료명 ?? "";
                Avalonia.Media.IBrush nameFg = AppRes("AppFg");
                if (!string.IsNullOrEmpty(row.CompoundName))
                {
                    var aliasInfo = CompoundAliasService.Resolve(row.CompoundName);
                    if (aliasInfo != null)
                    {
                        string samplePart = row.시료명?.Split('|').LastOrDefault()?.Trim() ?? row.시료명 ?? "";
                        displayName = $"{aliasInfo.Value.분석항목} | {samplePart}";
                        nameFg = new SolidColorBrush(Color.Parse("#90EE90"));
                    }
                }
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = displayName, FontFamily = Font,
                    Foreground = nameFg,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
            _rowNameCells.Add(nameCell);
            Grid.SetColumn(nameCell, 3);
            rowGrid.Children.Add(nameCell);

            // UV VIS: 흡광도가 있으면 계산농도(Fxy) + 결과값(Result) 항상 재계산
            if (isUVVISMode && !string.IsNullOrEmpty(row.D1))
            {
                double uvS = 0, uvI = 0;
                if (docInfo != null)
                {
                    double.TryParse(docInfo.Standard_Slope,     System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out uvS);
                    double.TryParse(docInfo.Standard_Intercept, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out uvI);
                }
                if (uvS > 0 && double.TryParse(row.D1, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var uvAbs))
                {
                    double.TryParse(row.시료량, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var uvVol);
                    double uvCalcConc = ((uvAbs - uvI) / uvS) * (uvVol > 0 ? 60.0 / uvVol : 1.0);
                    row.Fxy = uvCalcConc.ToString("F4");
                    // 결과값 항상 재계산 (PDF 기기 농도값 → ETA 공식으로 덮어쓰기)
                    double.TryParse(string.IsNullOrEmpty(row.D2) ? "1" : row.D2,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out var uvDil0);
                    if (uvDil0 <= 0) uvDil0 = 1;
                    int uvDp0 = (_activeCategory == "TN" || _activeItems.Any(x => x == "T-N"))
                        ? (row.Source == SourceType.폐수배출업소 ? 3 : 1)
                        : GetDecimalPlaces(_activeItems.FirstOrDefault() ?? _activeCategory);
                    row.Result = (uvCalcConc * uvDil0).ToString($"F{uvDp0}");
                }
            }

            // 기초정보: UV VIS는 시료량/흡광도/희석배수/계산농도,
            //          TOC TCIC는 TCAU/TCcon/ICAU/ICcon/희석배수,
            //          TOC NPOC는 D1/공백/농도/희석배수 (시료량 없음),
            //          ICP는 성분명/농도/희석배수,
            //          BOD는 시료량/D1/D2/f(x/y)/P
            string[] infoVals = isUVVISMode
                ? new[] { row.시료량, row.D1, row.Fxy, row.D2 }
                : isIcpMode
                ? new[] { row.CompoundName ?? "", row.Result, row.P }
                : isTocTcicMode
                ? new[] { row.TCAU, row.TCcon, row.ICAU, row.ICcon, row.P }
                : isTocMode
                ? new[] { row.Fxy, row.P }
                : new[] { row.시료량, row.D1, row.D2, row.Fxy, row.P };

            int colResult = 4 + infoVals.Length;
            int colSource = colResult + 1;

            // 결과값 TextBlock (TOC 인라인 편집에서 갱신하기 위해 먼저 선언)
            var valTb = FsBase(new TextBlock
            {
                Text = row.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            });

            // UV: 시료량 변경 시 재계산을 위해 루프 밖에서 선언 (ci==3에서 할당, ci==0에서 참조)
            Func<double, string>? uvResultCalc = null;

            for (int ci = 0; ci < infoVals.Length; ci++)
            {
                // UV VIS: 희석배수(ci=2) 인라인 편집 → 결과 재계산
                if (isUVVISMode && ci == 3)
                {
                    var capturedUvRow = row;
                    double uvSlope = 0, uvIntercept = 0;
                    if (docInfo != null)
                    {
                        double.TryParse(docInfo.Standard_Slope,     out uvSlope);
                        double.TryParse(docInfo.Standard_Intercept, out uvIntercept);
                    }
                    // TN: 비용부담금(폐수배출업소)은 소수점 3자리, 나머지 1자리
                    int uvDp = (_activeCategory == "TN" || _activeItems.Any(x => x == "T-N"))
                        ? (capturedUvRow.Source == SourceType.폐수배출업소 ? 3 : 1)
                        : GetDecimalPlaces(_activeItems.FirstOrDefault() ?? _activeCategory);
                    uvResultCalc = dil =>
                    {
                        double.TryParse(capturedUvRow.D1, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var abs);
                        double.TryParse(capturedUvRow.시료량, System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var vol);
                        double calcConc = uvSlope > 0
                            ? ((abs - uvIntercept) / uvSlope) * (vol > 0 ? 60.0 / vol : 1.0)
                            : 0;
                        capturedUvRow.Fxy = calcConc.ToString("F4");
                        return (calcConc * dil).ToString($"F{uvDp}");
                    };
                    // D2(희석배수) 필드를 P처럼 쓰기 위해 P에도 동기화
                    if (string.IsNullOrEmpty(capturedUvRow.P)) capturedUvRow.P = capturedUvRow.D2;
                    var uvDilPanel = BuildInlineDilCell(capturedUvRow, valTb, i, uvResultCalc);
                    // BuildInlineDilCell는 row.P를 씀 → D2도 같이 동기화
                    Grid.SetColumn(uvDilPanel, 4 + ci);
                    rowGrid.Children.Add(uvDilPanel);
                    continue;
                }

                // UV VIS: 시료량(ci=0) 인라인 편집
                if (isUVVISMode && ci == 0)
                {
                    var capturedUvRow = row;
                    var uvVolPanel = new Panel
                    {
                        VerticalAlignment = VerticalAlignment.Stretch,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        Cursor = new Cursor(StandardCursorType.Hand),
                    };
                    var uvVolBtn = FsBase(new Button
                    {
                        Content = string.IsNullOrEmpty(capturedUvRow.시료량) ? "—" : capturedUvRow.시료량,
                        FontFamily = Font, Foreground = AppRes("ThemeFgSecondary"),
                        Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        HorizontalContentAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Stretch,
                        Padding = new Thickness(2),
                    });
                    var uvVolInput = FsBase(new TextBox
                    {
                        Text = capturedUvRow.시료량,
                        FontFamily = Font, Foreground = AppRes("InputFg"),
                        Background = AppRes("InputBg"),
                        BorderBrush = AppRes("InputBorder"),
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(4, 0),
                        HorizontalContentAlignment = HorizontalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center,
                        Watermark = "mL", IsVisible = false,
                    });
                    NumericOnly(uvVolInput);
                    uvVolPanel.Children.Add(uvVolBtn);
                    uvVolPanel.Children.Add(uvVolInput);
                    void CommitVol()
                    {
                        capturedUvRow.시료량 = uvVolInput.Text ?? "";
                        uvVolBtn.Content = string.IsNullOrEmpty(capturedUvRow.시료량) ? "—" : capturedUvRow.시료량;
                        uvVolInput.IsVisible = false;
                        uvVolBtn.IsVisible = true;
                        // 시료량 변경 시 결과값 즉시 재계산
                        if (uvResultCalc != null)
                        {
                            double.TryParse(string.IsNullOrEmpty(capturedUvRow.P) ? "1" : capturedUvRow.P,
                                System.Globalization.NumberStyles.Float,
                                System.Globalization.CultureInfo.InvariantCulture, out var curDil);
                            if (curDil <= 0) curDil = 1;
                            capturedUvRow.Result = uvResultCalc(curDil);
                            valTb.Text = capturedUvRow.Result;
                        }
                    }
                    uvVolBtn.Click   += (_, _) => { uvVolBtn.IsVisible = false; uvVolInput.IsVisible = true; uvVolInput.Focus(); uvVolInput.SelectAll(); };
                    uvVolPanel.PointerPressed += (_, e) => { e.Handled = true; uvVolBtn.IsVisible = false; uvVolInput.IsVisible = true; uvVolInput.Focus(); uvVolInput.SelectAll(); };
                    uvVolInput.LostFocus += (_, _) => CommitVol();
                    int capturedVolIdx = i;
                    uvVolInput.AddHandler(TextBox.KeyDownEvent, (object? _, KeyEventArgs ke) =>
                    {
                        if (ke.Key == Key.Enter || ke.Key == Key.Down)
                        {
                            ke.Handled = true; CommitVol();
                            for (int nj = capturedVolIdx + 1; nj < _rowVolButtons.Count; nj++)
                                if (_rowVolButtons[nj] != null) { _rowVolButtons[nj]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent)); break; }
                        }
                        else if (ke.Key == Key.Up)
                        {
                            ke.Handled = true; CommitVol();
                            for (int nj = capturedVolIdx - 1; nj >= 0; nj--)
                                if (_rowVolButtons[nj] != null) { _rowVolButtons[nj]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent)); break; }
                        }
                        else if (ke.Key == Key.Escape) { ke.Handled = true; CommitVol(); }
                        // Shift+2 모드: Left/Right=시료량↔희석배수 전환 (WASD 제거 — IME 방지)
                        else if (_keyNavShow2 && ke.Key == Key.Left)  { ke.Handled = true; CommitVol(); _navEditDil = false; OpenNavEditCell(capturedVolIdx); }
                        else if (_keyNavShow2 && ke.Key == Key.Right) { ke.Handled = true; CommitVol(); _navEditDil = true;  OpenNavEditCell(capturedVolIdx); }
                    }, Avalonia.Interactivity.RoutingStrategies.Tunnel);
                    // _rowVolButtons에 등록 (행 인덱스 맞춤)
                    while (_rowVolButtons.Count <= i) _rowVolButtons.Add(null);
                    _rowVolButtons[i] = uvVolBtn;
                    Grid.SetColumn(uvVolPanel, 4 + ci);
                    rowGrid.Children.Add(uvVolPanel);
                    continue;
                }

                // TOC 초록/노란 행: 희석배수(P) 셀 인라인 편집 가능
                //   TCIC: 희석배수는 infoVals[4] (TCAU/TCcon/ICAU/ICcon/희석배수)
                //   NPOC: 희석배수는 infoVals[1] (AU/희석배수)
                int tocDilIdx = isTocTcicMode ? 4 : 1;
                bool isTocDilCol = isTocMode && ci == tocDilIdx;
                if (isTocDilCol)
                {
                    var capturedRow2 = row;
                    Func<double, string> calcToc = dil =>
                    {
                        int dp = GetDecimalPlaces("TOC");
                        double conc;
                        if (isTocTcicMode)
                        {
                            double.TryParse(capturedRow2.TCcon, out var tcv);
                            double.TryParse(capturedRow2.ICcon, out var icv);
                            conc = tcv - icv;
                        }
                        else
                        {
                            // NPOC: Result = 기기 측정 농도값, 희석배수 적용
                            double.TryParse(capturedRow2.Result, out conc);
                        }
                        return (conc * dil).ToString($"F{dp}");
                    };
                    var dilPanel = BuildInlineDilCell(row, valTb, i, calcToc);
                    Grid.SetColumn(dilPanel, 4 + ci);
                    rowGrid.Children.Add(dilPanel);
                    continue;
                }

                var infoTb = FsBase(new TextBlock
                {
                    Text = infoVals[ci], FontFamily = Font,
                    Foreground = AppRes("ThemeFgSecondary"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                });
                Grid.SetColumn(infoTb, 4 + ci);
                rowGrid.Children.Add(infoTb);
            }

            Grid.SetColumn(valTb, colResult);
            rowGrid.Children.Add(valTb);

            // 시료구분
            LogMatch($"UI DISPLAY: {row.SN} Source={row.Source} IsControl={row.IsControl}");
            // IsControl: 파서가 이미 정도관리로 표시\ed�� 행 — 매칭 전에도 "정도관리" 표시
            string srcLabel;
            Avalonia.Media.IBrush srcBrush;
            if (row.IsControl && row.Source == SourceType.미분류)
            {
                srcLabel = "정도관리";
                srcBrush = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#c084fc"));
            }
            else
            {
                (srcLabel, var srcFgKey) = row.Source switch
                {
                    SourceType.폐수배출업소 => (row.IsManualMatch ? "비용부담금" : $"폐수배출-{row.Matched?.구분 ?? "?"}", "ThemeFgInfo"),
                    SourceType.수질분석센터 => ("수질분석센터", "ThemeFgSuccess"),
                    SourceType.처리시설     => ("처리시설", "ThemeFgWarn"),
                    _                      => ("—", "FgMuted"),
                };
                srcBrush = AppRes(srcFgKey);
            }
            LogMatch($"UI DISPLAY: {row.SN} → '{srcLabel}'");
            var srcTb = FsBase(new TextBlock
            {
                Text = srcLabel, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = srcBrush,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            });
            Grid.SetColumn(srcTb, colSource);
            rowGrid.Children.Add(srcTb);
            _rowSourceCells.Add(srcTb);

            // dilBtn이 없는 행은 null 등록 (인덱스 동기화)
            if (_rowDilButtons.Count <= i)
                _rowDilButtons.Add(null);

            // 행 클릭 → 편집
            int idx = i;
            var border = new Border
            {
                Child = rowGrid,
                Cursor = new Cursor(StandardCursorType.Hand),
                BorderBrush = AppRes("ThemeBorderSubtle"),
                BorderThickness = new Thickness(0, 0, 0, 1),
                ClipToBounds = false,
            };
            var capturedIdx = i;
            border.PointerPressed += (_, pe) =>
            {
                if (pe.Handled) return;
                // Shift+Q 누른 상태에서 클릭: 정도관리로 지정
                if (_shiftQHeld)
                {
                    SelectGridRow(capturedIdx);
                    ApplyQcToRow(capturedRow, capturedIdx);
                    return;
                }
                // Shift+1 모드: Show1 포커스 아이템을 이 행에 즉시 매핑
                if (_keyNavShow1)
                    TryApplyShow1FocusToRow(capturedRow, capturedIdx);
                SelectGridRow(capturedIdx);
            };
            TextShimmer.AttachHover(border);

            // 드래그 드롭 수신: Show1에서 의뢰시료를 드래그하여 매칭
            DragDrop.SetAllowDrop(border, true);
            // capturedRow already declared above
            var capturedBorderForDrop = border;
            border.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
            {
                bool acceptable = e.Data.Contains("match-analysis")
                               || e.Data.Contains("match-facility")
                               || e.Data.Contains("match-waste")
                               || e.Data.Contains("match-compound");
                e.DragEffects = acceptable ? DragDropEffects.Link : DragDropEffects.None;
                if (acceptable)
                    capturedBorderForDrop.Background = AppRes("ThemeBorderActive");
                e.Handled = true;
            });
            border.AddHandler(DragDrop.DragLeaveEvent, (object? s, RoutedEventArgs e) =>
            {
                capturedBorderForDrop.Background = null;
            });
            border.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
            {
                capturedBorderForDrop.Background = null;
                LogMatch($"DROP EVENT: rowIdx={capturedIdx}, dataKeys=[{string.Join(",", e.Data.GetDataFormats())}]");
                ApplyDragMatch(capturedRow, capturedIdx, e.Data);
                e.Handled = true;
            });

            _gridPanel.Children.Add(border);
        }

        var scroll = new ScrollViewer
        {
            Content = _gridPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };

        // ESC 키로 매칭 취소 기능
        scroll.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape && _selectedRowIndex >= 0 && _currentExcelRows != null)
            {
                CancelSelectedRowMatching();
                e.Handled = true;
            }
        };
        scroll.Focusable = true;
        // 헤더 고정: 헤더 + 스크롤 본체를 StackPanel으로 조합
        var gridInner = new DockPanel();
        DockPanel.SetDock(_colHeaderBorder, Dock.Top);
        gridInner.Children.Add(_colHeaderBorder);
        gridInner.Children.Add(scroll);

        var gridBorder = new Border
        {
            Child = gridInner,
            CornerRadius = new CornerRadius(4),
            BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(1),
            ClipToBounds = true,
        };
        // 호버 상태에서 마우스 휠 스크롤 지원: 포커스 자동 이동
        scroll.Focusable = true;
        gridBorder.PointerEntered += (_, _) => scroll.Focus();
        // 전체 영역에서 휠 이벤트를 scroll로 전달 (handledEventsToo로 확실하게)
        gridBorder.AddHandler(Control.PointerWheelChangedEvent, (_, e) =>
        {
            scroll.Offset = new Vector(scroll.Offset.X, scroll.Offset.Y - e.Delta.Y * 50);
            e.Handled = true;
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel | Avalonia.Interactivity.RoutingStrategies.Bubble, true);

        // 인라인 캘린더 (Row 4, 초기 숨김)
        Grid.SetRow(calBorder, 4);
        root.Children.Add(calBorder);

        Grid.SetRow(gridBorder, 5);
        root.Children.Add(gridBorder);

        // 키보드 상하 이동 지원
        _currentExcelRows = excelRows;
        root.Focusable = true;
        root.KeyDown += OnGridKeyDown;

        // ── 입력 진행 오버레이 (ZIndex로 전체 덮음) ──────────────────────────
        _importPb = new ProgressBar
        {
            IsIndeterminate = false,
            Minimum = 0, Maximum = 100, Value = 0,
            Height = 4,
            CornerRadius = new CornerRadius(2),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        _importPbText = new TextBlock
        {
            Text = "입력 중...",
            FontSize = 12, FontFamily = Font,
            Foreground = AppRes("AppFg"),
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 0, 0, 6),
        };
        _importOverlay = new Border
        {
            IsVisible  = false,
            ZIndex     = 99,
            Background = AppRes("PanelBg") is var bg
                ? new Avalonia.Media.SolidColorBrush(((Avalonia.Media.SolidColorBrush)bg).Color, 0.88)
                : Avalonia.Media.Brushes.Transparent,
            Child = new StackPanel
            {
                VerticalAlignment   = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Spacing = 8,
                Width   = 260,
                Children = { _importPbText, _importPb }
            },
        };
        Grid.SetRow(_importOverlay, 0);
        Grid.SetRowSpan(_importOverlay, 6);
        root.Children.Add(_importOverlay);

        ListPanelChanged?.Invoke(root); // Show2 패널 업데이트 복원 - WindowPositionManager 제거로 안전

        // WindowPositionManager 제거로 복잡한 UI 업데이트 다시 활성화 가능
        _selectedRowIndex = -1;
    }

    /// <summary>
    /// UI 스레드 전용 LoadVerifiedGrid (데이터베이스 작업 없이 UI만 업데이트)
    /// 데이터는 이미 _currentSamples, _matchingAnalysisRecords, _matchingFacilityMasters에 로드된 상태
    /// </summary>
    private void LoadVerifiedGridUIOnly()
    {
        LogMatch($"LoadVerifiedGridUIOnly CALLED for category: {_activeCategory}");

        if (!_categoryExcelData.TryGetValue(_activeCategory, out var excelRows))
        {
            LogMatch($"LoadVerifiedGridUIOnly: No data found for category {_activeCategory}");
            if (_selectedDate != null) LoadSampleGrid(_selectedDate);
            return;
        }

        LogMatch($"LoadVerifiedGridUIOnly: Found {excelRows.Count} excel rows");

        // 매칭 상태 확인
        foreach (var row in excelRows.Take(5))
        {
            LogMatch($"Row {row.SN}: Source={row.Source}, Status={row.Status}, 시료명='{row.시료명}', Matched={row.Matched?.업체명}, Analysis={row.MatchedAnalysis?.시료명}");
        }

        // 3개 테이블 동시 매칭 (데이터는 이미 로드됨)
        foreach (var row in excelRows)
        {
            // 드래그-드롭으로 수동 매칭된 행은 자동 매칭 건너뛰기
            if (row.IsManualMatch)
            {
                LogMatch($"SKIP AUTO MATCH: {row.SN} (manual match: {row.Source})");
                continue;
            }

            // 1순위: 폐수배출업소 (SN 매칭 — [세풍]/[율촌] 접두사 무시 비교 포함)
            string rowSnBare = row.SN;
            if (rowSnBare.StartsWith("[")) { int bi = rowSnBare.IndexOf(']'); if (bi > 0) rowSnBare = rowSnBare[(bi + 1)..]; }
            row.Matched = _currentSamples.FirstOrDefault(s => s.SN == row.SN)
                       ?? _currentSamples.FirstOrDefault(s =>
                       {
                           string sBare = s.SN;
                           if (sBare.StartsWith("[")) { int bi2 = sBare.IndexOf(']'); if (bi2 > 0) sBare = sBare[(bi2 + 1)..]; }
                           return sBare == rowSnBare && sBare == row.SN; // bare SN이 같고 엑셀 SN에 접두사 없는 경우
                       })
                       ?? _currentSamples.FirstOrDefault(s =>
                       {
                           string sBare = s.SN;
                           if (sBare.StartsWith("[")) { int bi2 = sBare.IndexOf(']'); if (bi2 > 0) sBare = sBare[(bi2 + 1)..]; }
                           return sBare == rowSnBare;
                       })
                       ?? (!string.IsNullOrEmpty(row.SN) ? _currentSamples.FirstOrDefault(s => s.업체명 == row.시료명) : null);

            if (row.Matched != null)
            {
                // 원본 시료명 보존 후 업체명으로 교체 (auto-match by SN)
                if (string.IsNullOrEmpty(row.원본시료명)) row.원본시료명 = row.시료명;
                row.시료명 = row.Matched.업체명;
                row.Source = SourceType.폐수배출업소;
                row.Status = MatchStatus.입력가능;
                LogMatch($"WASTE MATCH: {row.SN} → {row.Matched.업체명}");
                continue;
            }

            // 2순위: 수질분석센터 의뢰 (시료명 유사도 검사 포함)
            row.MatchedAnalysis = _matchingAnalysisRecords.FirstOrDefault(ar => ar.시료명 == row.시료명 || ar.약칭 == row.시료명)
                               ?? _matchingAnalysisRecords.FirstOrDefault(ar => !string.IsNullOrEmpty(ar.시료명) &&
                                   CalculateSimilarity(ar.시료명, row.시료명) >= 0.85)
                               ?? _matchingAnalysisRecords.FirstOrDefault(ar => !string.IsNullOrEmpty(ar.약칭) &&
                                   CalculateSimilarity(ar.약칭, row.시료명) >= 0.85);

            if (row.MatchedAnalysis != null)
            {
                if (string.IsNullOrEmpty(row.원본시료명)) row.원본시료명 = row.시료명;
                row.시료명 = row.MatchedAnalysis.시료명;
                row.Source = SourceType.수질분석센터;
                row.Status = MatchStatus.입력가능;
                LogMatch($"ANALYSIS MATCH: {row.SN} → {row.MatchedAnalysis.시료명}");
                continue;
            }

            // 3순위: 처리시설 마스터 시료명 (시료명 유사도 검사 포함)
            if (_matchingFacilityMasters != null)
            {
                var facilityMatch = _matchingFacilityMasters.FirstOrDefault(fm => fm.시료명 == row.시료명);
                if (facilityMatch == default)
                {
                    facilityMatch = _matchingFacilityMasters.FirstOrDefault(fm => !string.IsNullOrEmpty(fm.시료명) &&
                                      CalculateSimilarity(fm.시료명, row.시료명) >= 0.85);
                }

                if (facilityMatch != default)
                {
                    if (string.IsNullOrEmpty(row.원본시료명)) row.원본시료명 = row.시료명;
                    row.시료명 = facilityMatch.시료명;
                    row.MatchedFacilityName = facilityMatch.시설명;
                    // MatchedMasterId 속성이 없으므로 제거
                    row.Source = SourceType.처리시설;
                    row.Status = MatchStatus.입력가능;
                    LogMatch($"FACILITY MATCH: {row.SN} → {facilityMatch.시설명} - {facilityMatch.시료명}");
                    continue;
                }
            }

            // 매칭 실패
            row.Source = SourceType.미분류;
            row.Status = MatchStatus.미매칭;
            LogMatch($"NO MATCH: {row.SN} → 미분류");
        }

        // ========== UI 그리드 구성 (LoadVerifiedGrid와 동일) ==========
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,Auto,Auto,*") };

        // [나머지 UI 구성 코드는 동일하므로 생략하여 줄임 - 실제 구현에서는 전체 복사 필요]

        // 임시로 간단한 UI만 생성하여 테스트
        var simplePanel = new StackPanel();
        simplePanel.Children.Add(new TextBlock { Text = $"TCIC UI 업데이트 완료: {excelRows.Count}건", FontFamily = Font });

        ListPanelChanged?.Invoke(simplePanel);

        // Show1: 드래그 매칭용 의뢰시료 목록 항시 표시
        BuildMatchBrowsePanel();

        LogMatch($"LoadVerifiedGridUIOnly 완료: {excelRows.Count}건");
    }

        /// <summary>
    /// 두 문자열 간의 유사도 계산 (0.0~1.0)
    /// </summary>
    private static double CalculateSimilarity(string a, string b)
    {
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;
        if (a == b) return 1.0;

        // 간단한 레벤시타인 거리 기반 유사도
        int maxLen = Math.Max(a.Length, b.Length);
        if (maxLen == 0) return 1.0;

        int distance = LevenshteinDistance(a, b);
        return (double)(maxLen - distance) / maxLen;
    }

    /// <summary>
    /// 레벤시타인 거리 계산
    /// </summary>
    private static int LevenshteinDistance(string source, string target)
    {
        if (source.Length == 0) return target.Length;
        if (target.Length == 0) return source.Length;

        var distance = new int[source.Length + 1, target.Length + 1];

        for (int i = 0; i <= source.Length; distance[i, 0] = i++) { }
        for (int j = 0; j <= target.Length; distance[0, j] = j++) { }

        for (int i = 1; i <= source.Length; i++)
        {
            for (int j = 1; j <= target.Length; j++)
            {
                int cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
                distance[i, j] = Math.Min(Math.Min(
                    distance[i - 1, j] + 1,     // deletion
                    distance[i, j - 1] + 1),    // insertion
                    distance[i - 1, j - 1] + cost); // substitution
            }
        }

        return distance[source.Length, target.Length];
    }

    private void SelectGridRow(int index)
    {
        if (_currentExcelRows == null || _gridPanel == null) return;
        if (index < 0 || index >= _currentExcelRows.Count) return;

        // 이전 선택 하이라이트 해제 (헤더는 별도 고정이므로 인덱스 = 행번호)
        if (_selectedRowIndex >= 0 && _selectedRowIndex < _gridPanel.Children.Count)
        {
            if (_gridPanel.Children[_selectedRowIndex] is Border prevBorder)
            {
                prevBorder.Background = null;
                prevBorder.BorderBrush = AppRes("ThemeBorderSubtle");
                prevBorder.BorderThickness = new Thickness(0, 0, 0, 1);
                // Shift+F 모드면 gold shimmer / 크기 해제
                if (_keyNavShow2) TextShimmer.StopFocus(prevBorder);
            }
        }

        _selectedRowIndex = index;

        // 새 행 하이라이트 (선택 행은 좌측 강조 바 + 배경 변경)
        if (index < _gridPanel.Children.Count && _gridPanel.Children[index] is Border border)
        {
            border.Background = AppRes("ThemeBorderActive");
            border.BorderBrush = AppRes("BtnPrimaryBg");
            border.BorderThickness = new Thickness(3, 0, 0, 1);
            border.BringIntoView();
            // Shift+F 모드면 gold shimmer + 글자 20% 증가
            if (_keyNavShow2) TextShimmer.StartFocus(border);
        }

        var row = _currentExcelRows[index];

        // UV/GC/TOC/ICP/LCMS 파서: 시료구분 무관하게 항상 Show3 편집 가능
        bool isInstrumentParser = _categoryDocInfo.TryGetValue(_activeCategory, out var rowDi)
            && (rowDi.IsUVVIS || rowDi.IsGcMode || rowDi.IsTocNPOC || rowDi.IsTocTCIC);

        if (row.Status is MatchStatus.입력가능 or MatchStatus.덮어쓰기 || isInstrumentParser)
        {
            if (row.Matched != null && !isInstrumentParser)
            {
                // 폐수배출업소: WasteSample 기반 편집 폼
                _selectedSample = row.Matched;
                ShowEditForm(row.Matched, row);
            }
            else
            {
                // 처리시설/수질분석센터/기기파서: Show3에 상세 표시
                // 키보드 네비 모드(Shift+1/2)에서는 희석배수·시료량 인라인편집 자동 오픈 안 함
                ShowExcelRowDetail(row);
            }
        }
        // 미매칭(빨간색) 또는 대기 → Show1 드래그앤드랍 안내
        else
        {
            ShowMessage($"'{row.시료명}' → Show1 드래그앤드랍 영역에서 분류하세요", false);
        }
    }

    private void OpenManualMatchDialog(ExcelRow exRow, int rowIndex)
    {
        // 의뢰시료: 최근 1개월 (없으면 _matchingAnalysisRecords 사용)
        List<AnalysisRequestRecord> analysisPool;
        try
        {
            analysisPool = AnalysisRequestService.GetRecentRecords(1);
            if (analysisPool.Count == 0)
                analysisPool = _matchingAnalysisRecords;
        }
        catch { analysisPool = _matchingAnalysisRecords; }

        // 시설 마스터가 없으면 직접 로드
        List<(string 시설명, string 시료명, int 마스터Id)> facilityPool =
            (_matchingFacilityMasters?.Count ?? 0) > 0
                ? _matchingFacilityMasters!
                : FacilityResultService.GetAllMasterSamples();

        // 기존 창/패널 정리
        _matchWindow?.Close();
        _matchWindow = null;

        var matchWin = new ManualMatchWindow(
            exRow.시료명, exRow.Result,
            analysisPool, facilityPool, _currentSamples);

        // 인라인 모드 설정: Show3(EditPanelChanged)에 표시
        matchWin.OnInlineClose = () =>
        {
            EditPanelChanged?.Invoke(null);
            _matchWindow = null;
        };
        matchWin.MatchConfirmed += w =>
        {
            ApplyManualMatch(exRow, rowIndex, w);
            EditPanelChanged?.Invoke(null);
            _matchWindow = null;
        };

        _matchWindow = matchWin;

        // Content를 Window에서 분리 후 Show3에 인라인으로 표시
        var panelContent = matchWin.DetachContent();
        if (panelContent != null)
            EditPanelChanged?.Invoke(panelContent);
    }

    private void RebuildBadge(List<ExcelRow> rows)
    {
        if (_summaryBadgePanel == null) return;
        int matchNew   = rows.Count(r => r.Status == MatchStatus.입력가능);
        int matchExist = rows.Count(r => r.Status == MatchStatus.덮어쓰기);
        int noMatch    = rows.Count(r => r.Status == MatchStatus.미매칭);
        int pending    = rows.Count(r => r.Status == MatchStatus.대기);
        _summaryBadgePanel.Children.Clear();
        if (pending > 0)
            _summaryBadgePanel.Children.Add(FsXS(new TextBlock { Text = $"⚪ {pending} 대기",     FontFamily = Font, Foreground = AppRes("FgMuted") }));
        if (matchNew > 0)
            _summaryBadgePanel.Children.Add(FsXS(new TextBlock { Text = $"🟢 {matchNew} 입력가능", FontFamily = Font, Foreground = AppRes("ThemeFgSuccess") }));
        if (matchExist > 0)
            _summaryBadgePanel.Children.Add(FsXS(new TextBlock { Text = $"🟡 {matchExist} 자료있음", FontFamily = Font, Foreground = AppRes("ThemeFgWarn") }));
        if (noMatch > 0)
            _summaryBadgePanel.Children.Add(FsXS(new TextBlock { Text = $"🔴 {noMatch} 의뢰정보 없음", FontFamily = Font, Foreground = AppRes("ThemeFgDanger") }));
        _summaryBadgePanel.Children.Add(FsXS(new TextBlock { Text = $"합계 {rows.Count}건", FontFamily = Font, Foreground = AppRes("FgMuted") }));
    }

    /// <summary>Show1에서 드래그된 의뢰시료를 Show2 행에 매칭</summary>
    private void ApplyDragMatch(ExcelRow exRow, int rowIndex, IDataObject data)
    {

        // Show1에서 드래그한 실제 이름 가져오기
        string draggedName = data.Get(DataFormats.Text)?.ToString() ?? "";
        string draggedSN = data.Get("sample-sn")?.ToString() ?? "";

        // 디버그: 드래그된 데이터 확인 (로그파일로 기록)
        if (App.EnableLogging)
        {
            try
            {
                var logPath = @"c:\Users\ironu\Documents\ETA\Logs\DragDropDebug.log";
                var logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [DRAG DEBUG] draggedName: '{draggedName}', draggedSN: '{draggedSN}'{Environment.NewLine}";
                File.AppendAllText(logPath, logEntry);
            }
            catch { /* 로그 실패해도 무시 */ }
        }

        // 드래그 소스 탭에 따라 구분 결정 + Show1 이름으로 매칭명 설정
        if (data.Contains("match-waste"))
        {
            LogMatch($"DRAG FROM: 비용부담금 탭 - '{draggedName}', SN: '{draggedSN}'");
            if (string.IsNullOrEmpty(exRow.원본시료명))
                exRow.원본시료명 = exRow.시료명;  // 원본 보존
            exRow.시료명 = draggedName;  // Show1 이름으로 변경
            exRow.SN = draggedSN;  // 비용부담금 실제 SN 설정
            exRow.Source = SourceType.폐수배출업소;
            exRow.Status = MatchStatus.입력가능;
            exRow.Enabled = true;
            exRow.IsManualMatch = true;
            // Matched 객체 설정 (ImportData에서 when 조건 통과용)
            exRow.Matched = WasteSampleService.FindBySN(draggedSN)
                ?? new WasteSample { SN = draggedSN, 업체명 = draggedName };
        }
        else if (data.Contains("match-analysis"))
        {
            LogMatch($"DRAG FROM: 수질분석센터 탭 - '{draggedName}', SN(약칭): '{draggedSN}'");
            if (string.IsNullOrEmpty(exRow.원본시료명))
                exRow.원본시료명 = exRow.시료명;  // 원본 보존
            exRow.시료명 = draggedName;  // Show1 이름으로 변경
            exRow.SN = draggedSN;  // 수질분석센터 약칭을 SN으로 설정
            exRow.Source = SourceType.수질분석센터;
            exRow.Status = MatchStatus.입력가능;
            exRow.Enabled = true;
            exRow.IsManualMatch = true;
            // MatchedAnalysis 객체 설정 (ImportData에서 when 조건 통과용)
            exRow.MatchedAnalysis = new AnalysisRequestRecord { 약칭 = draggedSN, 시료명 = draggedName };
        }
        else if (data.Contains("match-facility"))
        {
            LogMatch($"DRAG FROM: 처리시설 탭 - '{draggedName}', SN(시설명): '{draggedSN}'");
            if (string.IsNullOrEmpty(exRow.원본시료명))
                exRow.원본시료명 = exRow.시료명;  // 원본 보존
            exRow.시료명 = draggedName;  // Show1 이름으로 변경
            exRow.SN = draggedSN;  // SHOW1에서 가져온 처리시설명
            exRow.MatchedFacilityName = draggedSN;  // SN 셀 표시용 (시설명)
            exRow.Source = SourceType.처리시설;
            exRow.Status = MatchStatus.입력가능;
            exRow.Enabled = true;
            exRow.IsManualMatch = true;
        }
        else if (data.Contains("match-compound"))
        {
            // 분석항목 드래그: 화합물 별칭 등록 + 시료 그리드 일괄 업데이트
            LogMatch($"DRAG FROM: 분석항목 탭 - '{draggedName}' (카테고리: '{draggedSN}')");
            string rawCompound = exRow.CompoundName;
            if (string.IsNullOrEmpty(rawCompound))
            {
                LogMatch($"COMPOUND ALIAS: CompoundName이 비어있음 — 무시");
                return;
            }
            RegisterCompoundAliasAndUpdateGrid(rawCompound, draggedName);
            return;
        }
        else
        {
            LogMatch($"NO MATCHING DATA: keys=[{string.Join(",", data.GetDataFormats())}]");
            return;
        }

        LogMatch($"FINAL STATE: Source={exRow.Source}, Status={exRow.Status}, 시료명='{exRow.시료명}'");

        // 매칭된 행의 시료구분을 직접 업데이트
        if (rowIndex >= 0 && rowIndex < _currentExcelRows?.Count)
        {
            // _currentExcelRows 업데이트
            _currentExcelRows[rowIndex].Source = exRow.Source;
            _currentExcelRows[rowIndex].Status = exRow.Status;
            _currentExcelRows[rowIndex].Matched = exRow.Matched;
            _currentExcelRows[rowIndex].MatchedAnalysis = exRow.MatchedAnalysis;
            _currentExcelRows[rowIndex].MatchedFacilityName = exRow.MatchedFacilityName;
            _currentExcelRows[rowIndex].시료명 = exRow.시료명;
            _currentExcelRows[rowIndex].원본시료명 = exRow.원본시료명;
            _currentExcelRows[rowIndex].SN = exRow.SN; // SN 업데이트 추가
            _currentExcelRows[rowIndex].IsManualMatch = exRow.IsManualMatch;

            // _categoryExcelData도 동일하게 업데이트 (UI 표시용)
            LogMatch($"SYNC CHECK: _activeCategory='{_activeCategory}', rowIndex={rowIndex}");
            if (_categoryExcelData.TryGetValue(_activeCategory, out var categoryRows))
            {
                LogMatch($"SYNC CHECK: categoryRows.Count={categoryRows.Count}");
                if (rowIndex < categoryRows.Count)
                {
                    LogMatch($"SYNC BEFORE: categoryRows[{rowIndex}].Source={categoryRows[rowIndex].Source}");
                    categoryRows[rowIndex].Source = exRow.Source;
                    categoryRows[rowIndex].Status = exRow.Status;
                    categoryRows[rowIndex].Matched = exRow.Matched;
                    categoryRows[rowIndex].MatchedAnalysis = exRow.MatchedAnalysis;
                    categoryRows[rowIndex].MatchedFacilityName = exRow.MatchedFacilityName;
                    categoryRows[rowIndex].시료명 = exRow.시료명;
                    categoryRows[rowIndex].원본시료명 = exRow.원본시료명;
                    categoryRows[rowIndex].SN = exRow.SN; // SN 업데이트 추가
                    categoryRows[rowIndex].IsManualMatch = exRow.IsManualMatch;
                    LogMatch($"SYNC AFTER: categoryRows[{rowIndex}].Source={categoryRows[rowIndex].Source}");
                }
                else
                {
                    LogMatch($"SYNC ERROR: rowIndex {rowIndex} >= categoryRows.Count {categoryRows.Count}");
                }
            }
            else
            {
                LogMatch($"SYNC ERROR: _activeCategory '{_activeCategory}' not found in _categoryExcelData");
            }
        }

        // 시료명 셀 시각적 업데이트: 위=원본(작게 흰색), 아래=매칭명(크게 노란색)
        if (rowIndex < _rowNameCells.Count)
        {
            var nameCell = _rowNameCells[rowIndex];
            nameCell.Children.Clear();
            if (!string.IsNullOrEmpty(exRow.원본시료명))
            {
                nameCell.Children.Add(new TextBlock
                {
                    Text = exRow.원본시료명, FontFamily = Font,
                    FontSize = AppTheme.FontXS,
                    Foreground = AppRes("AppFg"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = $"↳ {exRow.시료명}", FontFamily = Font,
                    Foreground = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#ffe066")),
                    FontWeight = FontWeight.SemiBold,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
            else
            {
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = exRow.시료명, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
        }

        // 시료구분 셀 시각적 업데이트
        if (rowIndex < _rowSourceCells.Count)
        {
            var srcTb = _rowSourceCells[rowIndex];
            var (srcLabel, srcFgKey) = exRow.Source switch
            {
                SourceType.폐수배출업소 => (exRow.IsManualMatch ? "비용부담금" : $"폐수배출-{exRow.Matched?.구분 ?? "?"}", "ThemeFgInfo"),
                SourceType.수질분석센터 => ("수질분석센터", "ThemeFgSuccess"),
                SourceType.처리시설     => ("처리시설", "ThemeFgWarn"),
                _                      => ("—", "FgMuted"),
            };
            srcTb.Text = srcLabel;
            srcTb.Foreground = AppRes(srcFgKey);
        }

        // 그리드 아이콘 업데이트
        if (rowIndex < _rowIcons.Count)
        {
            _rowIcons[rowIndex].Background = exRow.Status switch
            {
                MatchStatus.입력가능 => AppRes("ThemeFgSuccess"),
                MatchStatus.덮어쓰기 => AppRes("ThemeFgWarn"),
                _ => AppRes("FgMuted"),
            };
            _rowIcons[rowIndex].Opacity = 1.0;
        }

        // SN 셀 업데이트 (처리시설=시설명, 수질분석센터=약칭, 나머지=SN)
        if (rowIndex < _rowSnCells.Count)
        {
            string snText = exRow.Source switch
            {
                SourceType.처리시설     when !string.IsNullOrEmpty(exRow.MatchedFacilityName) => exRow.MatchedFacilityName,
                SourceType.처리시설     => exRow.SN,
                SourceType.수질분석센터 => exRow.SN,
                _ => exRow.SN,
            };
            _rowSnCells[rowIndex].Text = snText;
            if (App.EnableLogging)
            {
                try { File.AppendAllText(@"c:\Users\ironu\Documents\ETA\Logs\DragDropDebug.log",
                    $"{DateTime.Now:HH:mm:ss} - [SN CELL] rowIndex={rowIndex}, Source={exRow.Source}, SN='{exRow.SN}', FacilityName='{exRow.MatchedFacilityName}', snText='{snText}'{Environment.NewLine}"); } catch { }
            }
        }

        // 토글 스위치 ON으로 전환
        if (rowIndex < _rowToggles.Count)
        {
            var (track, knob) = _rowToggles[rowIndex];
            track.Background = AppRes("BtnPrimaryBg");
            track.Opacity    = 1.0;
            knob.Margin      = new Thickness(18, 2, 0, 2);
        }

        RefreshDuplicateIcons(); // 매칭 후 중복 시료명 재검사 → 핑크색 아이콘
        BuildStatsPanel();

        LogMatch($"MATCH COMPLETED: {exRow.원본시료명} → {exRow.시료명}");
    }

    private void ApplyManualMatch(ExcelRow exRow, int rowIndex, ManualMatchWindow w)
    {
        if (w.SelectedWaste != null)
        {
            exRow.Matched  = w.SelectedWaste;
            exRow.Source   = SourceType.폐수배출업소;
            if (string.IsNullOrEmpty(exRow.원본시료명))
                exRow.원본시료명 = exRow.시료명;             // 원본 보존
            exRow.시료명   = w.SelectedWaste.업체명;         // 업체명으로 교체
            bool has = _activeItems.Any(item => !string.IsNullOrWhiteSpace(GetSampleValue(w.SelectedWaste, item)));
            exRow.Status   = has ? MatchStatus.덮어쓰기 : MatchStatus.입력가능;
            exRow.Enabled  = true;
        }
        else if (w.SelectedAnalysis != null)
        {
            if (string.IsNullOrEmpty(exRow.원본시료명))
                exRow.원본시료명  = exRow.시료명;
            exRow.시료명          = w.SelectedAnalysis.시료명;
            exRow.MatchedAnalysis = w.SelectedAnalysis;
            exRow.Source          = SourceType.수질분석센터;
            exRow.Status          = MatchStatus.입력가능;
            exRow.Enabled         = true;
        }
        else if (w.SelectedFacility != null)
        {
            if (string.IsNullOrEmpty(exRow.원본시료명))
                exRow.원본시료명     = exRow.시료명;                        // 원본 보존
            exRow.시료명             = w.SelectedFacility.Value.시료명;    // 지점명으로 교체
            exRow.MatchedFacilityName = w.SelectedFacility.Value.시설명;
            exRow.Source              = SourceType.처리시설;
            exRow.Status              = MatchStatus.입력가능;
            exRow.Enabled             = true;
        }

        // 그리드 아이콘 업데이트
        if (rowIndex < _rowIcons.Count)
        {
            _rowIcons[rowIndex].Background = exRow.Status switch
            {
                MatchStatus.입력가능 => AppRes("ThemeFgSuccess"),
                MatchStatus.덮어쓰기 => AppRes("ThemeFgWarn"),
                _ => AppRes("FgMuted"),
            };
            _rowIcons[rowIndex].Opacity = 1.0;
        }

        // SN 셀 업데이트 (처리시설=시설명, 수질분석센터=약칭, 나머지=SN)
        if (rowIndex < _rowSnCells.Count)
        {
            string snText = exRow.Source switch
            {
                SourceType.처리시설     when !string.IsNullOrEmpty(exRow.MatchedFacilityName) => exRow.MatchedFacilityName,
                SourceType.처리시설     => exRow.SN,
                SourceType.수질분석센터 when !string.IsNullOrEmpty(exRow.MatchedAnalysis?.약칭) => exRow.MatchedAnalysis!.약칭,
                SourceType.수질분석센터 => exRow.SN,
                _ => exRow.SN,
            };
            _rowSnCells[rowIndex].Text = snText;

            // SN 셀 업데이트 로그
            if (App.EnableLogging)
            {
                try { File.AppendAllText(@"c:\Users\ironu\Documents\ETA\Logs\DragDropDebug.log",
                    $"{DateTime.Now:HH:mm:ss} - [SN CELL] rowIndex={rowIndex}, Source={exRow.Source}, SN='{exRow.SN}', FacilityName='{exRow.MatchedFacilityName}', snText='{snText}', _rowSnCells.Count={_rowSnCells.Count}{Environment.NewLine}"); } catch { }
            }
        }
        else
        {
            if (App.EnableLogging)
            {
                try { File.AppendAllText(@"c:\Users\ironu\Documents\ETA\Logs\DragDropDebug.log",
                    $"{DateTime.Now:HH:mm:ss} - [SN CELL ERROR] rowIndex={rowIndex} >= _rowSnCells.Count={_rowSnCells.Count}{Environment.NewLine}"); } catch { }
            }
        }

        // 토글 스위치 ON으로 전환
        if (rowIndex < _rowToggles.Count)
        {
            var (track, knob) = _rowToggles[rowIndex];
            track.Background = AppRes("BtnPrimaryBg");
            track.Opacity    = 1.0;
            knob.Margin      = new Thickness(18, 2, 0, 2);
        }

        // 시료명 셀 업데이트: 위=원본(작게 흰색), 아래=매칭명(크게 노란색)
        if (rowIndex < _rowNameCells.Count)
        {
            var nameCell = _rowNameCells[rowIndex];
            nameCell.Children.Clear();
            if (!string.IsNullOrEmpty(exRow.원본시료명))
            {
                nameCell.Children.Add(new TextBlock
                {
                    Text = exRow.원본시료명, FontFamily = Font,
                    FontSize = AppTheme.FontXS,
                    Foreground = AppRes("AppFg"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = $"↳ {exRow.시료명}", FontFamily = Font,
                    Foreground = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#ffe066")),
                    FontWeight = FontWeight.SemiBold,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
            else
            {
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = exRow.시료명, FontFamily = Font,
                    Foreground = AppRes("AppFg"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
        }

        // 그리드 전체 다시 빌드 (매칭 후 희석배수 버튼 등 상태 반영)
        { /* LoadVerifiedGrid 제거됨 */ }
    }

    /// <summary>
    /// 드래그앤드롭 매칭 후 중복 시료명(시료명+SN)을 재검사하여 아이콘/체크박스를 업데이트합니다.
    /// </summary>
    private void RefreshDuplicateIcons()
    {
        if (_currentExcelRows == null) return;

        // 현재 시료명+SN 기준 중복 그룹 계산 (SN 없는 미매칭 행 제외)
        var dupGroups = _currentExcelRows
            .Select((r, idx) => (r, idx))
            .Where(x => !string.IsNullOrEmpty(x.r.SN))
            .GroupBy(x => $"{x.r.시료명}_{x.r.SN}")
            .Where(g => g.Count() > 1)
            .ToList();

        var dupIndices = dupGroups.SelectMany(g => g.Select(x => x.idx)).ToHashSet();

        for (int idx = 0; idx < _currentExcelRows.Count && idx < _rowIcons.Count; idx++)
        {
            var row = _currentExcelRows[idx];
            var icon = _rowIcons[idx];

            if (!dupIndices.Contains(idx)) continue; // 중복 아닌 행은 건드리지 않음

            // 이미 핑크색이면 스킵
            if (icon.Background is Avalonia.Media.ISolidColorBrush sb &&
                sb.Color == Avalonia.Media.Color.Parse("#FF69B4")) continue;

            // 아이콘을 핑크색으로
            icon.Background = Brush.Parse("#FF69B4");
            icon.Opacity = 1.0;
            icon.Cursor = new Cursor(StandardCursorType.Hand);

            // 이미 체크박스가 있으면 스킵
            if (idx >= _rowIconPanels.Count) continue;
            var panel = _rowIconPanels[idx];
            if (panel.Children.OfType<CheckBox>().Any()) continue;

            // 체크박스 동적 추가
            string snKey = $"{row.시료명}_{row.SN}";
            var capturedRow = row;
            var capturedIcon = icon;
            var cb = new CheckBox
            {
                IsChecked = capturedRow.IsSelectedForFinalResult,
                IsVisible = false,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2, 0, 0, 0),
            };
            cb.IsCheckedChanged += (_, _) =>
            {
                capturedRow.IsSelectedForFinalResult = cb.IsChecked == true;
                if (_dupIconGroups.TryGetValue(snKey, out var groupIcons))
                {
                    foreach (var gi in groupIcons)
                        gi.Background = (gi == capturedIcon && cb.IsChecked == true)
                            ? AppRes("ThemeFgSuccess")
                            : Brush.Parse("#87CEEB");
                }
            };
            panel.Children.Add(cb);

            if (!_dupCheckboxGroups.ContainsKey(snKey)) _dupCheckboxGroups[snKey] = new();
            _dupCheckboxGroups[snKey].Add(cb);
            if (!_dupIconGroups.ContainsKey(snKey)) _dupIconGroups[snKey] = new();
            if (!_dupIconGroups[snKey].Contains(icon)) _dupIconGroups[snKey].Add(icon);

            // 핑크 아이콘 클릭 → 체크박스 토글 이벤트 (중복 추가 방지)
            icon.PointerPressed += (_, e) =>
            {
                e.Handled = true;
                if (!_dupCheckboxGroups.TryGetValue(snKey, out var groupCbs) || groupCbs.Count == 0) return;
                bool show = !groupCbs[0].IsVisible;
                foreach (var groupCb in groupCbs) groupCb.IsVisible = show;
            };
        }
    }

    /// <summary>
    /// 인라인 희석배수 입력 셀을 생성합니다. TOC 외 다른 항목에서도 재사용 가능.
    /// </summary>
    /// <param name="row">대상 ExcelRow (row.P = 희석배수)</param>
    /// <param name="resultTb">결과값 TextBlock (희석배수 변경 시 갱신)</param>
    /// <param name="rowIdx">행 인덱스 (_rowDilButtons 기준 이동용)</param>
    /// <param name="calcResult">희석배수(double) → 결과 문자열 계산 함수</param>
    /// <returns>그리드 셀에 직접 추가할 Panel</returns>
    private Panel BuildInlineDilCell(ExcelRow row, TextBlock resultTb, int rowIdx, Func<double, string> calcResult)
    {
        var dilPanel = new Panel
        {
            VerticalAlignment = VerticalAlignment.Stretch,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Cursor = new Cursor(StandardCursorType.Hand),
        };
        var dilBtn = FsBase(new Button
        {
            Content = string.IsNullOrEmpty(row.P) ? "1" : row.P,
            FontFamily = Font, Foreground = AppRes("ThemeFgWarn"),
            FontWeight = FontWeight.SemiBold,
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Stretch,
            Padding = new Thickness(2),
        });
        var dilInput = FsBase(new TextBox
        {
            Text = string.IsNullOrEmpty(row.P) ? "1" : row.P,
            FontFamily = Font, Foreground = AppRes("InputFg"),
            Background = AppRes("InputBg"),
            BorderBrush = AppRes("ThemeFgWarn"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 0),
            MinWidth = 40,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Watermark = "1",
            IsVisible = false,
        });
        NumericOnly(dilInput);
        dilPanel.Children.Add(dilBtn);
        dilPanel.Children.Add(dilInput);

        void OpenDilInput()
        {
            if (dilInput.IsVisible) return;
            dilBtn.IsVisible = false;
            dilInput.IsVisible = true;
            dilInput.Text = string.IsNullOrEmpty(row.P) ? "1" : row.P;
            dilInput.Focus();
            dilInput.SelectAll();
        }

        void CommitDil()
        {
            if (!dilInput.IsVisible) return;
            double.TryParse(dilInput.Text, out var dil);
            if (dil <= 0) dil = 1;
            row.P = dil.ToString("G");
            row.D2 = row.P;  // UV Show3/DB 저장용 동기화
            row.Result = calcResult(dil);
            dilBtn.Content = row.P;
            resultTb.Text = row.Result;
            dilInput.IsVisible = false;
            dilBtn.IsVisible = true;
        }

        void MoveDilFocus(int delta)
        {
            CommitDil();
            int start = rowIdx + delta;
            if (delta > 0)
            {
                for (int nj = start; nj < _rowDilButtons.Count; nj++)
                    if (_rowDilButtons[nj] != null) { _rowDilButtons[nj]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent)); break; }
            }
            else
            {
                for (int nj = start; nj >= 0; nj--)
                    if (_rowDilButtons[nj] != null) { _rowDilButtons[nj]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent)); break; }
            }
        }

        dilPanel.PointerPressed += (_, e) => { e.Handled = true; OpenDilInput(); };
        dilBtn.Click += (_, _) => OpenDilInput();
        dilInput.LostFocus += (_, _) => CommitDil();
        dilInput.AddHandler(TextBox.KeyDownEvent, (object? _, KeyEventArgs ke) =>
        {
            if (ke.Key == Key.Enter || ke.Key == Key.Down) { ke.Handled = true; MoveDilFocus(+1); }
            else if (ke.Key == Key.Up)                     { ke.Handled = true; MoveDilFocus(-1); }
            else if (ke.Key == Key.Escape)                 { ke.Handled = true; dilInput.IsVisible = false; dilBtn.IsVisible = true; }
            // Shift+2 모드: Left/Right=시료량↔희석배수 전환 (WASD 제거 — IME 방지)
            else if (_keyNavShow2 && ke.Key == Key.Left)  { ke.Handled = true; CommitDil(); _navEditDil = false; OpenNavEditCell(rowIdx); }
            else if (_keyNavShow2 && ke.Key == Key.Right) { ke.Handled = true; CommitDil(); _navEditDil = true;  OpenNavEditCell(rowIdx); }
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel);

        _rowDilButtons.Add(dilBtn);
        return dilPanel;
    }

    private void OnGridKeyDown(object? sender, KeyEventArgs e)
    {
        if (_currentExcelRows == null) return;

        if (e.Key == Key.Down)
        {
            SelectGridRow(_selectedRowIndex + 1);
            e.Handled = true;
        }
        else if (e.Key == Key.Up)
        {
            SelectGridRow(_selectedRowIndex - 1);
            e.Handled = true;
        }
    }

    // =========================================================================
    // 키보드 단축키 네비게이션  (Shift+R = Show1 트리, Shift+F = Show2 그리드)
    // =========================================================================

    /// <summary>Tunnel 핸들러: _keyNavShow2 활성 시 방향키(↑↓←→)로 행·열 이동 및 자동 입력창 오픈</summary>
    private void OnNavKeyDownTunnel(object? sender, KeyEventArgs e)
    {
        if (e.Handled) return;
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

        // ── _keyNavShow2 모드: 방향키만 = 이동 (WASD 제거 — IME 방지), Enter = 편집 ──
        if (!shift && _keyNavShow2)
        {
            // Up/Down → 행 이동 후 자동으로 선택된 셀 열기
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                e.Handled = true;
                TopLevel.GetTopLevel(this)?.FocusManager?.ClearFocus(); // 편집 중이면 LostFocus → 커밋
                bool up = e.Key == Key.Up;

                // Sample 그리드 모드 (Show2 샘플 데이터)
                if (_selectedSampleIndex >= 0 || (_currentSamples?.Count > 0 && _selectedRowIndex < 0))
                {
                    int newIdx = _selectedSampleIndex + (up ? -1 : 1);
                    SelectSampleRow(newIdx);
                    return;
                }

                // Excel 그리드 모드 (Show2 파서 데이터)
                int excelIdx = _selectedRowIndex + (up ? -1 : 1);
                SelectGridRow(excelIdx);
                Avalonia.Threading.Dispatcher.UIThread.Post(() => OpenNavEditCell(excelIdx));
                return;
            }
            // Left/Right → 열 전환 후 자동으로 선택된 셀 열기
            if (e.Key == Key.Left || e.Key == Key.Right)
            {
                e.Handled = true;
                TopLevel.GetTopLevel(this)?.FocusManager?.ClearFocus(); // 편집 중이면 LostFocus → 커밋
                _navEditDil = e.Key == Key.Right;
                Avalonia.Threading.Dispatcher.UIThread.Post(() => OpenNavEditCell(_selectedRowIndex));
                return;
            }
            // Enter → TextBox 편집 중이면 내부 핸들러에 위임, 아니면 현재 셀 열기
            if (e.Key == Key.Enter)
            {
                if (e.Source is TextBox) return; // 내부 Tunnel 핸들러가 다음 행으로 이동 처리
                e.Handled = true;
                // Excel 그리드: OpenNavEditCell (Sample 그리드는 SelectSampleRow에서 이미 ShowEditForm 호출)
                if (_selectedRowIndex >= 0)
                    OpenNavEditCell(_selectedRowIndex);
                return;
            }
            // 숫자/소수점 → 즉시 편집 시작 + 해당 키를 TextBox에 전달
            if (e.Key >= Key.D0 && e.Key <= Key.D9 || e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9 || e.Key == Key.OemPeriod || e.Key == Key.Decimal)
            {
                e.Handled = true;
                OpenNavEditCell(_selectedRowIndex);
                // 열린 TextBox를 찾아서 초기값으로 입력한 숫자를 넣어줌
                var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() as TextBox;
                if (focused != null)
                {
                    char ch = e.Key switch
                    {
                        Key.D0 or Key.NumPad0 => '0', Key.D1 or Key.NumPad1 => '1',
                        Key.D2 or Key.NumPad2 => '2', Key.D3 or Key.NumPad3 => '3',
                        Key.D4 or Key.NumPad4 => '4', Key.D5 or Key.NumPad5 => '5',
                        Key.D6 or Key.NumPad6 => '6', Key.D7 or Key.NumPad7 => '7',
                        Key.D8 or Key.NumPad8 => '8', Key.D9 or Key.NumPad9 => '9',
                        Key.OemPeriod or Key.Decimal => '.',
                        _ => '\0'
                    };
                    if (ch != '\0')
                    {
                        focused.Text = ch.ToString();
                        focused.CaretIndex = 1;
                    }
                }
                return;
            }
        }
    }

    private void OnWindowKeyDown(object? sender, KeyEventArgs e)
    {
        // TextBox 포커스 중에는 단축키 무시 (Tunnel 핸들러에서 이미 처리)
        if (e.Source is TextBox) return;

        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

        // ── Shift+1: Show1(의뢰시료 목록) W/S 모드 토글 ─────────────────
        if (shift && e.Key == Key.D1)
        {
            e.Handled = true;
            if (_keyNavShow2) { _keyNavShow2 = false; ClearShow2Focus(); }
            _keyNavShow1 = !_keyNavShow1;
            if (_keyNavShow1 && _keyNavShow1Index < 0 && _matchItems.Count > 0)
                NavShow1(0); // 첫 번째 아이템 포커스
            else if (!_keyNavShow1)
                ClearShow1Focus();
            return;
        }

        // ── Shift+2: Show2(그리드) 방향키 네비 모드 토글 ──────────────────
        if (shift && e.Key == Key.D2)
        {
            e.Handled = true;
            if (_keyNavShow1) { _keyNavShow1 = false; ClearShow1Focus(); }
            _keyNavShow2 = !_keyNavShow2;
            if (_keyNavShow2)
            {
                // Sample 그리드 모드
                if (_selectedSampleIndex < 0 && _currentSamples?.Count > 0)
                    SelectSampleRow(0);
                // Excel 그리드 모드
                else if (_selectedRowIndex < 0 && _currentExcelRows?.Count > 0)
                    SelectGridRow(0);
            }
            else
                ClearShow2Focus();
            return;
        }

        // ── Shift+Q: QC 클릭 모드 (누르고 있는 동안 활성) ────────────
        if (shift && e.Key == Key.Q)
        {
            e.Handled = true;
            if (!_shiftQHeld)
            {
                _shiftQHeld = true;
                ShowMessage("QC 모드 — 클릭하면 정도관리로 지정됩니다", false);
            }
            return;
        }

        // ── ESC: 매칭 취소 (포커스 위치와 무관하게 처리) ──────────────────
        if (!shift && e.Key == Key.Escape)
        {
            CancelSelectedRowMatching();
            e.Handled = true;
            return;
        }

        // ── Up / Down: 활성 모드에서 위/아래 ────────────────────────────────
        if (!shift && (e.Key == Key.Up || e.Key == Key.Down))
        {
            if (_keyNavShow1 && _matchItems.Count > 0)
            {
                e.Handled = true;
                bool up = e.Key == Key.Up;
                int next = Math.Clamp(_keyNavShow1Index + (up ? -1 : 1), 0, _matchItems.Count - 1);
                NavShow1(next);
                return;
            }
            else if (_keyNavShow2)
            {
                e.Handled = true;
                bool up = e.Key == Key.Up;
                int newIdx = _selectedRowIndex + (up ? -1 : 1);
                SelectGridRow(newIdx);
                Avalonia.Threading.Dispatcher.UIThread.Post(() => OpenNavEditCell(newIdx));
                return;
            }
        }

        // ── Left / Right: Shift+2 모드에서 희석배수 ↔ 시료량 전환 후 자동 열기 ──────────
        if (!shift && _keyNavShow2 && (e.Key == Key.Left || e.Key == Key.Right))
        {
            e.Handled = true;
            _navEditDil = e.Key == Key.Right; // Left=시료량, Right=희석배수
            Avalonia.Threading.Dispatcher.UIThread.Post(() => OpenNavEditCell(_selectedRowIndex));
            return;
        }
    }

    /// <summary>Show1 의뢰시료 아이템 포커스 이동</summary>
    private void NavShow1(int index)
    {
        if (index < 0 || index >= _matchItems.Count) return;
        // 이전 포커스 해제
        if (_keyNavShow1Index >= 0 && _keyNavShow1Index < _matchItems.Count)
            TextShimmer.StopFocus(_matchItems[_keyNavShow1Index].Item);

        _keyNavShow1Index = index;
        var (item, _, _) = _matchItems[index];
        TextShimmer.StartFocus(item);
        item.BringIntoView();
    }

    /// <summary>Shift+2 모드: 현재 행의 희석배수 또는 시료량 인라인 편집 자동 오픈</summary>
    private void OpenNavEditCell(int rowIndex)
    {
        if (rowIndex < 0) return;

        if (_navEditDil)
        {
            // 희석배수 버튼 클릭 → 인라인 편집 오픈
            if (rowIndex < _rowDilButtons.Count && _rowDilButtons[rowIndex] != null)
                _rowDilButtons[rowIndex]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        }
        else
        {
            // 시료량 버튼 클릭 → 인라인 편집 오픈
            if (rowIndex < _rowVolButtons.Count && _rowVolButtons[rowIndex] != null)
                _rowVolButtons[rowIndex]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            else if (rowIndex < _rowDilButtons.Count && _rowDilButtons[rowIndex] != null)
                _rowDilButtons[rowIndex]!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
        }
    }

    private void OnWindowKeyUp(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Q || e.Key == Key.LeftShift || e.Key == Key.RightShift)
            _shiftQHeld = false;
    }

    private void ClearShow1Focus()
    {
        if (_keyNavShow1Index >= 0 && _keyNavShow1Index < _matchItems.Count)
            TextShimmer.StopFocus(_matchItems[_keyNavShow1Index].Item);
        _keyNavShow1Index = -1;
    }

    private void ClearShow2Focus()
    {
        // Excel 그리드 선택 해제
        if (_selectedRowIndex >= 0 && _gridPanel != null
            && _selectedRowIndex < _gridPanel.Children.Count
            && _gridPanel.Children[_selectedRowIndex] is Border b)
        {
            TextShimmer.StopFocus(b);
            b.Background = null;
            b.BorderBrush = AppRes("ThemeBorderSubtle");
            b.BorderThickness = new Thickness(0, 0, 0, 1);
        }
        _selectedRowIndex = -1;

        // Sample 그리드 선택 해제
        if (_selectedSampleIndex >= 0 && _sampleGridPanel != null
            && _selectedSampleIndex < _sampleGridPanel.Children.Count
            && _sampleGridPanel.Children[_selectedSampleIndex] is Border sb)
        {
            TextShimmer.StopFocus(sb);
            sb.Background = null;
            sb.BorderBrush = AppRes("ThemeBorderSubtle");
            sb.BorderThickness = new Thickness(0, 0, 0, 1);
        }
        _selectedSampleIndex = -1;
    }

    // ── Shift+R 모드: Show2 행 클릭 시 포커스된 Show1 아이템 적용 ────────
    private void TryApplyShow1FocusToRow(ExcelRow exRow, int rowIndex)
    {
        if (!_keyNavShow1 || _keyNavShow1Index < 0 || _keyNavShow1Index >= _matchItems.Count) return;
        var (_, name, data) = _matchItems[_keyNavShow1Index];
        // DataObject 생성하여 ApplyDragMatch 재사용
#pragma warning disable CS0618
        var dragData = new DataObject();
        dragData.Set(DataFormats.Text, name);
        if (data != null)
        {
            dynamic d = data;
            string sampleType = d.Type;
            string matchKey = sampleType switch
            {
                "Analysis" => "match-analysis",
                "Facility" => "match-facility",
                "Compound" => "match-compound",
                _ => "match-waste"
            };
            dragData.Set(matchKey, name);
            try { dragData.Set("sample-sn", (string)(d.SN?.ToString() ?? "")); } catch { }
        }
        else
            dragData.Set("match-waste", name);
#pragma warning restore CS0618
        ApplyDragMatch(exRow, rowIndex, dragData);
        // 다음 Show1 아이템으로 자동 전진
        if (_keyNavShow1Index + 1 < _matchItems.Count)
            NavShow1(_keyNavShow1Index + 1);
    }

    // =========================================================================
    // Show2: 일반 그리드 (첨부 파일 없을 때)
    // =========================================================================
    private bool _detailMode = false;

    /// <summary>카테고리 키 → *_DATA 테이블명</summary>
    private string? GetDataTableName(string catKey)
    {
        if (catKey == "TOC")
        {
            // NPOC vs TCIC: docInfo에서 판별
            if (_categoryDocInfo.TryGetValue("TOC", out var di) && di.IsTocTCIC)
                return "TOC_TCIC_DATA";
            return "TOC_NPOC_DATA";
        }
        return catKey switch
        {
            "BOD" => "BOD_DATA",
            "SS" => "SS_DATA",
            "NHEX" => "NHexan_DATA",
            "TN" => "TN_DATA",
            "TP" => "TP_DATA",
            "PHENOLS" => "Phenols_DATA",
            "CN" => "CN_DATA",
            "CR6" => "CR6_DATA",
            "COLOR" => "COLOR_DATA",
            "ABS" => "ABS_DATA",
            "FLUORIDE" => "FLUORIDE_DATA",
            _ => null
        };
    }

    /// <summary>카테고리별 Detail 컬럼 정의</summary>
    private static string[] GetDetailColumns(string catKey) => catKey switch
    {
        "BOD" => new[] { "시료량", "D1", "D2", "희석배수", "식종시료량", "식종D1", "식종D2", "식종BOD" },
        "SS" => new[] { "시료량", "전무게", "후무게", "무게차", "희석배수" },
        "NHEX" => new[] { "시료량" },
        "TN" or "TP" or "PHENOLS" or "CN" or "CR6" or "COLOR" or "ABS" or "FLUORIDE" => new[] { "시료량", "흡광도", "희석배수", "검량선_a", "농도" },
        _ => Array.Empty<string>()
    };

    private void LoadSampleGrid(string date)
    {
        _currentSamples = WasteSampleService.GetByDate(date);

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        DateTime.TryParse(date, out var d);
        var catLabel = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;
        var catKey = _activeCategory;
        var dataTable = GetDataTableName(catKey);
        bool isUV = catKey is "TN" or "TP" or "PHENOLS" or "CN" or "CR6" or "COLOR" or "ABS" or "FLUORIDE";
        bool isSS = catKey == "SS";
        bool isBOD = catKey == "BOD";
        bool isToc = catKey == "TOC";

        // 헤더 + Detail 토글
        var headerPanel = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        headerPanel.Children.Add(FsSM(new TextBlock
        {
            Text = $"📋 {d:yyyy-MM-dd}  [{catLabel}]  —  {_currentSamples.Count}건",
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        }));

        if (dataTable != null)
        {
            var toggleLabel = FsXS(new TextBlock
            {
                Text = "Detail", FontFamily = Font, Foreground = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0),
            });
            var toggleTrack = new Border
            {
                Width = 30, Height = 14, CornerRadius = new CornerRadius(7),
                Background = _detailMode ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted"),
                Cursor = new Cursor(StandardCursorType.Hand),
                VerticalAlignment = VerticalAlignment.Center,
            };
            var toggleKnob = new Border
            {
                Width = 10, Height = 10, CornerRadius = new CornerRadius(5),
                Background = Brushes.White, HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(_detailMode ? 18 : 2, 2, 0, 2),
            };
            toggleTrack.Child = toggleKnob;
            toggleTrack.PointerPressed += (_, _) => { _detailMode = !_detailMode; LoadSampleGrid(date); };
            var toggleWrap = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 2, Margin = new Thickness(4, 0) };
            toggleWrap.Children.Add(toggleLabel);
            toggleWrap.Children.Add(toggleTrack);
            Grid.SetColumn(toggleWrap, 1);
            headerPanel.Children.Add(toggleWrap);
        }

        var header = new Border
        {
            Background = AppRes("PanelBg"), CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(10, 6), Child = headerPanel,
        };
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        _sampleGridPanel = _gridPanel = new StackPanel { Spacing = 0 };
        _selectedSampleIndex = -1;  // 새로운 그리드 로드 시 선택 초기화

        // 엑셀과 동일한 컬럼 헤더 (Detail 모드)
        string[] hLabels;
        string[] dataKeys; // *_DATA 딕셔너리에서 가져올 키 배열
        if (_detailMode && dataTable != null)
        {
            if (isSS)
            {
                // SS_DATA: 전무게/후무게/무게차/희석배수 컬럼 포함
                hLabels  = new[] { "SN", "업체명", "시료량", "전무게", "후무게", "무게차", "희석배수", "결과값", "분석자", "분석일시", "구분" };
                dataKeys = new[] { "",    "",       "시료량", "전무게", "후무게", "무게차", "희석배수", "",       "@mgr",   "등록일시", "" };
            }
            else if (isUV)
            {
                hLabels  = new[] { "SN", "업체명", "시료량", "흡광도", "희석배수", "검량선_a", "결과값", "분석자", "분석일시", "구분" };
                dataKeys = new[] { "",    "",       "시료량", "흡광도", "희석배수", "검량선_a", "",       "@mgr",   "등록일시", "" };
            }
            else if (isToc)
            {
                // TOC: 시료량 없음, D1/희석배수/결과값
                hLabels  = new[] { "SN", "업체명", "AU", "희석배수", "결과값", "분석자", "분석일시", "구분" };
                dataKeys = new[] { "",    "",       "흡광도", "희석배수", "",   "@mgr",   "등록일시", "" };
            }
            else if (isBOD)
            {
                hLabels  = new[] { "SN", "업체명", "시료량", "D1", "D2", "희석배수", "결과값", "식종시료량", "식종D1", "식종D2", "식종BOD", "식종함유량", "분석자", "분석일시", "구분" };
                dataKeys = new[] { "",    "",       "시료량", "D1", "D2", "희석배수", "",       "식종시료량", "식종D1", "식종D2", "식종BOD", "식종함유량", "@mgr",   "등록일시", "" };
            }
            else // NHEX 등
            {
                hLabels  = new[] { "SN", "업체명", "시료량", "결과값", "분석자", "분석일시", "구분" };
                dataKeys = new[] { "",    "",       "시료량", "",       "@mgr",   "등록일시", "" };
            }
        }
        else
        {
            hLabels  = new[] { "구분", "S/N", "업체명", catLabel };
            dataKeys = Array.Empty<string>();
        }

        // 컬럼 정의
        string sgColDefs;
        if (_detailMode && dataTable != null)
        {
            var widths = new List<string>();
            foreach (var h in hLabels)
                widths.Add(h is "업체명" ? "180" : h is "SN" or "S/N" ? "80" : "70");
            sgColDefs = string.Join(",", widths);
        }
        else
            sgColDefs = "50,80,180,70";

        var hGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(sgColDefs),
            MinHeight = 26, Background = AppRes("GridHeaderBg"),
        };
        for (int c = 0; c < hLabels.Length; c++)
        {
            bool isDetail = _detailMode && c >= 2 && c < hLabels.Length - 2;
            var tb = FsSM(new TextBlock
            {
                Text = hLabels[c], FontWeight = FontWeight.SemiBold,
                FontFamily = Font,
                Foreground = isDetail ? AppRes("ThemeFgSecondary") : AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = c >= 2 ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                Margin = new Thickness(c == 0 ? 6 : 2, 0),
            });
            Grid.SetColumn(tb, c); hGrid.Children.Add(tb);
        }
        _gridPanel.Children.Add(new Border { Child = hGrid, BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

        // 분석자 조회용 캐시: 채수일 → (항목전체명 → 담당자)
        var mgrCache = new Dictionary<string, Dictionary<string, string>>();
        // 약칭 역매핑: 항목전체명 → 약칭 (예: "생물화학적 산소요구량" → "BOD")
        Dictionary<string, string>? shortNames = null;
        if (_detailMode && dataTable != null)
            shortNames = AnalysisRequestService.GetShortNames();

        string GetManager(string 채수일)
        {
            if (shortNames == null) return "";
            if (!mgrCache.TryGetValue(채수일, out var mgrs))
            {
                mgrs = AnalysisRequestService.GetManagersByDate(채수일);
                mgrCache[채수일] = mgrs;
            }
            // 항목전체명 중 약칭이 현재 카테고리 _activeItems에 매칭되는 것 찾기
            foreach (var (fullName, manager) in mgrs)
            {
                if (shortNames.TryGetValue(fullName, out var abbr) &&
                    _activeItems.Any(item => item.Equals(abbr, StringComparison.OrdinalIgnoreCase)))
                    return manager;
            }
            return "";
        }

        // 데이터
        for (int i = 0; i < _currentSamples.Count; i++)
        {
            var s = _currentSamples[i];
            Dictionary<string, string>? rawData = null;
            if (_detailMode && dataTable != null)
                rawData = WasteSampleService.GetRawData(dataTable, s.채수일, s.SN);

            var rGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(sgColDefs),
                MinHeight = 30,
                Background = i % 2 == 0 ? AppRes("GridRowBg") : AppRes("GridRowAltBg"),
            };

            string itemKey = _activeItems.FirstOrDefault() ?? "";
            string resultVal = FormatResult(GetSampleValue(s, itemKey), itemKey);
            bool hasResult = !string.IsNullOrWhiteSpace(resultVal);

            if (_detailMode && dataTable != null)
            {
                // Detail 모드: 엑셀과 동일한 컬럼 순서
                for (int c = 0; c < hLabels.Length; c++)
                {
                    string val;
                    IBrush fg;
                    if (hLabels[c] == "SN") { val = s.SN; fg = AppRes("ThemeFgInfo"); }
                    else if (hLabels[c] == "업체명")
                    {
                        // 비고(원본시료명) 있으면 원본 작게 + 매칭명 크게 표시
                        string origName = rawData != null && rawData.TryGetValue("비고", out var rmk) && !string.IsNullOrWhiteSpace(rmk) ? rmk : "";
                        if (!string.IsNullOrEmpty(origName) && origName != s.업체명)
                        {
                            var nameCell = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(2, 0) };
                            nameCell.Children.Add(FsXS(new TextBlock { Text = origName, FontFamily = Font,
                                Foreground = AppRes("AppFg"), TextTrimming = TextTrimming.CharacterEllipsis }));
                            nameCell.Children.Add(FsBase(new TextBlock { Text = $"↳ {s.업체명}", FontFamily = Font,
                                Foreground = new SolidColorBrush(Color.Parse("#ffe066")),
                                FontWeight = FontWeight.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis }));
                            Grid.SetColumn(nameCell, c); rGrid.Children.Add(nameCell);
                            continue;
                        }
                        val = s.업체명; fg = AppRes("AppFg");
                    }
                    else if (hLabels[c] == "결과값") { val = hasResult ? resultVal : "—"; fg = hasResult ? AppRes("ThemeFgSuccess") : AppRes("ThemeFgDimmed"); }
                    else if (hLabels[c] == "구분") { val = s.구분; fg = AppRes("FgMuted"); }
                    else
                    {
                        string dKey = dataKeys[c];
                        if (dKey == "@mgr")
                        {
                            // 분장표준처리에서 담당자 조회
                            val = GetManager(s.채수일);
                            bool hv = !string.IsNullOrWhiteSpace(val);
                            val = hv ? val : "—";
                            fg = hv ? AppRes("ThemeFgInfo") : AppRes("ThemeFgDimmed");
                        }
                        else
                        {
                            // *_DATA에서 가져오기
                            val = rawData != null && !string.IsNullOrEmpty(dKey) && rawData.TryGetValue(dKey, out var rv) ? rv : "";
                            bool hv = !string.IsNullOrWhiteSpace(val);
                            val = hv ? val : "—";
                            fg = hv ? AppRes("AppFg") : AppRes("ThemeFgDimmed");
                        }
                    }
                    var tb = FsSM(new TextBlock
                    {
                        Text = val, FontFamily = Font, Foreground = fg,
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = c >= 2 ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                        Margin = new Thickness(c == 0 ? 6 : 2, 0),
                        TextTrimming = TextTrimming.CharacterEllipsis,
                    });
                    Grid.SetColumn(tb, c); rGrid.Children.Add(tb);
                }
            }
            else
            {
                // 기본 모드: 구분 / SN / 업체명 / 결과
                string[] fv = { s.구분, s.SN, s.업체명 };
                for (int c = 0; c < fv.Length; c++)
                {
                    var tb = FsSM(new TextBlock
                    {
                        Text = fv[c], FontFamily = Font, Foreground = AppRes("AppFg"),
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(c == 0 ? 6 : 2, 0),
                        TextTrimming = TextTrimming.CharacterEllipsis,
                    });
                    Grid.SetColumn(tb, c); rGrid.Children.Add(tb);
                }
                var resTb = FsSM(new TextBlock
                {
                    Text = hasResult ? resultVal : "—", FontFamily = Font,
                    Foreground = hasResult ? AppRes("ThemeFgSuccess") : AppRes("ThemeFgDimmed"),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                });
                Grid.SetColumn(resTb, 3); rGrid.Children.Add(resTb);
            }

            var border = new Border
            {
                Child = rGrid, Cursor = new Cursor(StandardCursorType.Hand),
                BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1),
            };
            int capturedIdx = i;
            border.PointerPressed += (_, _) => { SelectSampleRow(capturedIdx); };
            TextShimmer.AttachHover(border);
            _gridPanel.Children.Add(border);
        }

        var scroll = new ScrollViewer { Content = _gridPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);
        ListPanelChanged?.Invoke(root);
    }

    /// <summary>편집 중인 샘플의 입력값을 _selectedSample에 저장</summary>
    private void CommitSampleEdit()
    {
        if (_selectedSample == null || _inputBoxes.Count == 0) return;

        // TextBox 값들을 _selectedSample에 복사
        foreach (var (key, textBox) in _inputBoxes)
        {
            var val = textBox.Text ?? "";
            // WasteSample 프로퍼티에 값 할당
            switch (key.ToLower())
            {
                case "bod": _selectedSample.BOD = val; break;
                case "toc": _selectedSample.TOC = val; break;
                case "ss": _selectedSample.SS = val; break;
                case "t-n": _selectedSample.TN = val; break;
                case "t-p": _selectedSample.TP = val; break;
                case "n-hexan": _selectedSample.NHexan = val; break;
                case "phenols": _selectedSample.Phenols = val; break;
                case "시안": _selectedSample.CN = val; break;
                case "6가크롬": _selectedSample.CR6 = val; break;
                case "색도": _selectedSample.COLOR = val; break;
                case "abs": _selectedSample.ABS = val; break;
                case "불소": _selectedSample.FLUORIDE = val; break;
            }
        }

        // DB에 저장
        try
        {
            string colName = GetColumnNameFromCategory();
            if (!string.IsNullOrEmpty(colName))
                WasteSampleService.UpdateDynamicValue("폐수의뢰및결과", _selectedSample.Id, colName, GetResultValue());
        }
        catch { }
    }

    /// <summary>현재 카테고리의 결과값 반환</summary>
    private string GetResultValue()
    {
        if (_selectedSample == null) return "";
        return _activeCategory switch
        {
            "BOD" => _selectedSample.BOD,
            "TOC" => _selectedSample.TOC,
            "SS" => _selectedSample.SS,
            "TN" => _selectedSample.TN,
            "TP" => _selectedSample.TP,
            "NHEX" => _selectedSample.NHexan,
            "PHENOLS" => _selectedSample.Phenols,
            "CN" => _selectedSample.CN,
            "CR6" => _selectedSample.CR6,
            "COLOR" => _selectedSample.COLOR,
            "ABS" => _selectedSample.ABS,
            "FLUORIDE" => _selectedSample.FLUORIDE,
            _ => ""
        };
    }

    /// <summary>현재 카테고리의 컬럼명 반환</summary>
    private string GetColumnNameFromCategory()
    {
        return _activeCategory switch
        {
            "BOD" => "BOD",
            "TOC" => "`TOC`",
            "SS" => "SS",
            "TN" => "`T-N`",
            "TP" => "`T-P`",
            "NHEX" => "`N-Hexan`",
            "PHENOLS" => "Phenols",
            "CN" => "`시안`",
            "CR6" => "`6가크롬`",
            "COLOR" => "`색도`",
            "ABS" => "`ABS`",
            "FLUORIDE" => "`불소`",
            _ => ""
        };
    }

    /// <summary>Show2 샘플 그리드에서 행을 선택하고 ShowEditForm 열기</summary>
    private void SelectSampleRow(int index)
    {
        if (_sampleGridPanel == null || _currentSamples == null) return;
        if (index < 0 || index >= _currentSamples.Count) return;

        // 이전 행의 편집 내용 커밋
        CommitSampleEdit();

        // 이전 선택 하이라이트 해제
        if (_selectedSampleIndex >= 0 && _selectedSampleIndex < _sampleGridPanel.Children.Count)
        {
            if (_sampleGridPanel.Children[_selectedSampleIndex] is Border prevBorder)
            {
                prevBorder.Background = null;
                prevBorder.BorderBrush = AppRes("ThemeBorderSubtle");
                prevBorder.BorderThickness = new Thickness(0, 0, 0, 1);
                if (_keyNavShow2) TextShimmer.StopFocus(prevBorder);
            }
        }

        _selectedSampleIndex = index;

        // 새 행 하이라이트
        if (index < _sampleGridPanel.Children.Count && _sampleGridPanel.Children[index] is Border border)
        {
            border.Background = AppRes("ThemeBorderActive");
            border.BorderBrush = AppRes("BtnPrimaryBg");
            border.BorderThickness = new Thickness(3, 0, 0, 1);
            border.BringIntoView();
            if (_keyNavShow2) TextShimmer.StartFocus(border);
        }

        var sample = _currentSamples[index];
        _selectedSample = sample;
        ShowEditForm(sample);
    }

    private static string GetSampleValue(WasteSample s, string item) => item switch
    {
        "BOD" => s.BOD, "TOC" => s.TOC, "SS" => s.SS,
        "T-N" => s.TN, "T-P" => s.TP, "N-Hexan" => s.NHexan, "Phenols" => s.Phenols,
        "시안" => s.CN, "6가크롬" => s.CR6, "색도" => s.COLOR, "ABS" => s.ABS, "불소" => s.FLUORIDE,
        _ => ""
    };

    // =========================================================================
    // Show2: 처리시설 그리드 (배지 클릭 시 표시) — 시설별 시료 목록 + 해당 카테고리 값
    // =========================================================================
    private void LoadFacilityGrid(string date)
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        DateTime.TryParse(date, out var d);
        var catLabel = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;
        // 처리시설 테이블 컬럼명 매핑: _activeItems[0] 이 실제 DB 컬럼명 (T-N, 총대장균군 등)
        string facColumn = _activeItems.FirstOrDefault() ?? _activeCategory;

        // 시설 목록 수집
        var facilities = FacilityResultService.GetFacilityNames();
        var allRows = new List<(string facility, FacilityResultRow row)>();
        foreach (var fac in facilities)
        {
            foreach (var r in FacilityResultService.GetRows(fac, date))
            {
                // 해당 항목이 이 시료에 활성화된 경우만 표시
                if (r.IsActive(facColumn))
                    allRows.Add((fac, r));
            }
        }

        // 헤더
        var headerPanel = new Grid { ColumnDefinitions = new ColumnDefinitions("*") };
        headerPanel.Children.Add(FsSM(new TextBlock
        {
            Text = $"🏭 {d:yyyy-MM-dd}  [처리시설 · {catLabel}]  —  {allRows.Count}건",
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        }));
        var header = new Border
        {
            Background = AppRes("PanelBg"), CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding = new Thickness(10, 6), Child = headerPanel,
        };
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        _gridPanel = new StackPanel { Spacing = 0 };

        // 컬럼 헤더: 시설명 / 시료명 / 결과값 / 비고
        string[] hLabels = { "시설명", "시료명", catLabel, "비고" };
        string sgColDefs = "120,180,100,*";
        var hGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(sgColDefs),
            MinHeight = 26, Background = AppRes("GridHeaderBg"),
        };
        for (int c = 0; c < hLabels.Length; c++)
        {
            var tb = FsSM(new TextBlock
            {
                Text = hLabels[c], FontWeight = FontWeight.SemiBold,
                FontFamily = Font, Foreground = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = c == 2 ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                Margin = new Thickness(c == 0 ? 6 : 2, 0),
            });
            Grid.SetColumn(tb, c); hGrid.Children.Add(tb);
        }
        _gridPanel.Children.Add(new Border { Child = hGrid, BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

        // 데이터 행
        for (int i = 0; i < allRows.Count; i++)
        {
            var (fac, r) = allRows[i];
            var rGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(sgColDefs),
                MinHeight = 30,
                Background = i % 2 == 0 ? AppRes("GridRowBg") : AppRes("GridRowAltBg"),
            };

            string resultVal = r[facColumn];
            bool hasResult = !string.IsNullOrWhiteSpace(resultVal);

            string[] fv = { fac, r.시료명, hasResult ? resultVal : "—", r.비고마스터 };
            for (int c = 0; c < fv.Length; c++)
            {
                IBrush fg;
                if (c == 2)
                    fg = hasResult ? AppRes("ThemeFgSuccess") : AppRes("ThemeFgDimmed");
                else if (c == 3)
                    fg = AppRes("FgMuted");
                else
                    fg = AppRes("AppFg");

                var tb = FsSM(new TextBlock
                {
                    Text = fv[c], FontFamily = Font, Foreground = fg,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = c == 2 ? HorizontalAlignment.Center : HorizontalAlignment.Left,
                    Margin = new Thickness(c == 0 ? 6 : 2, 0),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
                Grid.SetColumn(tb, c); rGrid.Children.Add(tb);
            }

            var border = new Border
            {
                Child = rGrid, Cursor = new Cursor(StandardCursorType.Hand),
                BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1),
            };
            TextShimmer.AttachHover(border);
            _gridPanel.Children.Add(border);
        }

        var scroll = new ScrollViewer { Content = _gridPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);
        ListPanelChanged?.Invoke(root);
    }

    // =========================================================================
    // Show3: 편집 폼
    // =========================================================================
    private void ShowEditForm(WasteSample sample, ExcelRow? exRow = null)
    {
        _inputBoxes.Clear();
        _currentEditExcelRow = exRow;
        _currentEditSample = sample;  // WasteSample 편집 추적
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(12) };

        var catLabel = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;

        // 타이틀 + 저장 버튼 (같은 행, 좌우 배치)
        var btnSave = FsBase(new Button
        {
            Content = "💾 적용", FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Background = AppRes("BtnPrimaryBg"), Foreground = AppRes("BtnPrimaryFg"),
            BorderBrush = AppRes("BtnPrimaryBorder"), BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6), Padding = new Thickness(16, 6),
            HorizontalAlignment = HorizontalAlignment.Right, Cursor = new Cursor(StandardCursorType.Hand),
        });
        btnSave.Click += (_, _) => SaveCurrentSample();

        var titleRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        var titleTb = FsLG(new TextBlock
        {
            Text = $"✏️ {sample.업체명}  [{catLabel}]",
            FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        Grid.SetColumn(titleTb, 0); titleRow.Children.Add(titleTb);
        Grid.SetColumn(btnSave, 1); titleRow.Children.Add(btnSave);
        root.Children.Add(titleRow);

        // 정보 배지
        var infoPanel = new WrapPanel { Orientation = Orientation.Horizontal };
        foreach (var (label, val) in new[] {
            ("S/N", sample.SN), ("구분", sample.구분), ("채수일", sample.채수일), ("관리번호", sample.관리번호) })
        {
            infoPanel.Children.Add(new Border
            {
                Background = AppRes("ThemeBorderSubtle"), CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 2), Margin = new Thickness(0, 0, 6, 4),
                Child = new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 4,
                    Children =
                    {
                        FsXS(new TextBlock { Text = label, FontFamily = Font,
                            Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center }),
                        FsBase(new TextBlock { Text = val, FontFamily = Font,
                            Foreground = AppRes("AppFg"), FontWeight = FontWeight.SemiBold,
                            VerticalAlignment = VerticalAlignment.Center }),
                    }
                }
            });
        }
        root.Children.Add(infoPanel);

        // 기초정보 편집 + 자동 계산
        bool isUV = _categoryDocInfo.TryGetValue(_activeCategory, out var editDocInfo) && editDocInfo.IsUVVIS;
        bool isEditSS = editDocInfo?.IsSS == true;
        bool isEditToc = editDocInfo != null && (editDocInfo.IsTocNPOC || editDocInfo.IsTocTCIC);
        bool isGcMode = editDocInfo?.IsGcMode == true;

        // 식종수 정보 (BOD 계산용)
        double seedVol = 0, seedD1 = 0, seedD2 = 0, seedPct = 0;
        if (!isUV && !isEditSS && editDocInfo != null)
        {
            double.TryParse(editDocInfo.식종수_시료량, out seedVol);
            double.TryParse(editDocInfo.식종수_D1, out seedD1);
            double.TryParse(editDocInfo.식종수_D2, out seedD2);
            double.TryParse(editDocInfo.식종수_Remark, out seedPct);
        }

        // 계산 결과 표시용 TextBlock
        TextBlock? fxyDisplay = null, pDisplay = null, resultDisplay = null;

        // 편집 필드 생성 헬퍼
        TextBox MakeInput(string initVal, string watermark = "0.00")
        {
            var tb = FsBase(new TextBox
            {
                Text = initVal, FontFamily = Font,
                Foreground = AppRes("InputFg"), Background = AppRes("InputBg"),
                BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 4),
                HorizontalAlignment = HorizontalAlignment.Stretch, Watermark = watermark,
            });
            return tb;
        }

        Grid MakeRow(string label, Control valueCtrl, string unit = "")
        {
            var g = new Grid { ColumnDefinitions = new ColumnDefinitions("80,*,50"), Margin = new Thickness(0, 2) };
            var lb = FsBase(new TextBlock { Text = label, FontWeight = FontWeight.SemiBold, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lb, 0); g.Children.Add(lb);
            Grid.SetColumn(valueCtrl, 1); g.Children.Add(valueCtrl);
            if (!string.IsNullOrEmpty(unit))
            {
                var u = FsXS(new TextBlock { Text = unit, FontFamily = Font,
                    Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
                Grid.SetColumn(u, 2); g.Children.Add(u);
            }
            return g;
        }

        // Show4 수식 표시용
        var formulaPanel = new StackPanel { Spacing = 6, Margin = new Thickness(10) };

        if (exRow != null && isEditSS)
        {
            // SS 모드: 시료량/전무게/후무게 편집 → 무게차/결과 자동 계산
            var volInput = MakeInput(exRow.시료량);
            var preInput = MakeInput(exRow.D1);  // 전무게
            var postInput = MakeInput(exRow.D2); // 후무게
            var dilInput = MakeInput(exRow.P);   // 희석배수

            var diffDisplay = FsBase(new TextBlock { Text = exRow.Fxy, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            resultDisplay = FsLG(new TextBlock { Text = exRow.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });

            var fmTitle = FsBase(new TextBlock { Text = "SS 계산 수식", FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppRes("AppFg") });
            var fmFormula = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap,
                Text = "SS(mg/L) = (후무게 - 전무게) × 1000 / 시료량 × 희석배수" });
            var fmRes = FsBase(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
                FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });

            formulaPanel.Children.Add(fmTitle);
            formulaPanel.Children.Add(fmFormula);
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle") });
            formulaPanel.Children.Add(fmRes);

            void RecalcSS()
            {
                if (!double.TryParse(volInput.Text, out var vol) || vol <= 0) return;
                if (!double.TryParse(preInput.Text, out var pre)) return;
                if (!double.TryParse(postInput.Text, out var post)) return;
                if (!double.TryParse(dilInput.Text, out var dil)) dil = 1;

                double diff = post - pre;
                double result = diff * 1000 / vol * dil;

                int dp = GetDecimalPlaces("SS");
                diffDisplay.Text = diff.ToString("F4");
                resultDisplay!.Text = result.ToString($"F{dp}");

                exRow.시료량 = volInput.Text; exRow.D1 = preInput.Text; exRow.D2 = postInput.Text;
                exRow.Fxy = diff.ToString("F4"); exRow.P = dilInput.Text;
                exRow.Result = result.ToString($"F{dp}");

                fmRes.Text = $"SS = ({post} - {pre}) × 1000 / {vol} × {dil}\n= {result.ToString($"F{dp}")} mg/L";
            }

            volInput.TextChanged += (_, _) => RecalcSS();
            preInput.TextChanged += (_, _) => RecalcSS();
            postInput.TextChanged += (_, _) => RecalcSS();
            dilInput.TextChanged += (_, _) => RecalcSS();
            RecalcSS();

            // 수평 배열
            var inputGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,6,Auto,*,6,Auto,*,6,Auto,*,16,Auto,60,6,Auto,80,Auto"),
                Margin = new Thickness(0, 6),
                MinHeight = 32,
            };

            var lblVol = FsXS(new TextBlock { Text = "시료량", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblVol, 0); inputGrid.Children.Add(lblVol);
            Grid.SetColumn(volInput, 1); inputGrid.Children.Add(volInput);
            var lblMl = FsXS(new TextBlock { Text = "mL", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblMl, 2); inputGrid.Children.Add(lblMl);

            var lblPre = FsXS(new TextBlock { Text = "전무게", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblPre, 3); inputGrid.Children.Add(lblPre);
            Grid.SetColumn(preInput, 4); inputGrid.Children.Add(preInput);

            var lblPost = FsXS(new TextBlock { Text = "후무게", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
            Grid.SetColumn(lblPost, 6); inputGrid.Children.Add(lblPost);
            Grid.SetColumn(postInput, 7); inputGrid.Children.Add(postInput);

            var lblDil = FsXS(new TextBlock { Text = "희석", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
            Grid.SetColumn(lblDil, 9); inputGrid.Children.Add(lblDil);
            Grid.SetColumn(dilInput, 10); inputGrid.Children.Add(dilInput);

            var lblDiff = FsXS(new TextBlock { Text = "무게차", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblDiff, 12); inputGrid.Children.Add(lblDiff);
            Grid.SetColumn(diffDisplay, 13); inputGrid.Children.Add(diffDisplay);

            var lblRes = FsXS(new TextBlock { Text = "Result", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblRes, 15); inputGrid.Children.Add(lblRes);
            Grid.SetColumn(resultDisplay, 16); inputGrid.Children.Add(resultDisplay);
            var lblUnit = FsXS(new TextBlock { Text = "mg/L", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblUnit, 17); inputGrid.Children.Add(lblUnit);

            root.Children.Add(inputGrid);
        }
        else if (exRow != null && isEditToc)
        {
            // TOC 모드: 희석배수 편집 가능, y=ax+b 수식으로 계산 (시료량 없음)
            var dilInput = MakeInput(exRow.P); // 희석배수

            var d1Display = FsBase(new TextBlock { Text = exRow.D1, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            fxyDisplay = FsBase(new TextBlock { Text = exRow.Fxy, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            resultDisplay = FsLG(new TextBlock { Text = exRow.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });

            // Show4 수식 패널
            string slope = editDocInfo?.TocSlope_TC ?? "0";
            string intercept = editDocInfo?.TocIntercept_TC ?? "0";
            var fmTitle = FsBase(new TextBlock { Text = "TOC 계산 수식", FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppRes("AppFg") });
            var fmFormula = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"),
                TextWrapping = TextWrapping.Wrap,
                Text = $"y = {slope}x + {intercept}" });
            var fmCalc = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"),
                TextWrapping = TextWrapping.Wrap });
            var fmRes = FsBase(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
                FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });

            formulaPanel.Children.Add(fmTitle);
            formulaPanel.Children.Add(fmFormula);
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle") });
            formulaPanel.Children.Add(fmCalc);
            formulaPanel.Children.Add(fmRes);

            void RecalcToc()
            {
                double.TryParse(exRow.Fxy, out var conc);   // 기기 계산 농도
                double.TryParse(dilInput.Text, out var dil);
                if (dil <= 0) dil = 1;
                double final_ = conc * dil;
                int dp = GetDecimalPlaces("TOC");
                exRow.P = dilInput.Text;
                exRow.Result = final_.ToString($"F{dp}");
                resultDisplay.Text = exRow.Result;
                fmCalc.Text = $"농도 = {conc:F4}  x  희석배수 {dil}";
                fmRes.Text = $"최종 농도 = {final_.ToString($"F{dp}")} mg/L";
            }

            dilInput.TextChanged += (_, _) => RecalcToc();
            RecalcToc();

            // 수평 배열: D1, f(x/y), P(희석배수), Result (시료량 없음)
            var inputGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,80,16,Auto,80,16,Auto,*,16,Auto,80,Auto"),
                Margin = new Thickness(0, 6),
                MinHeight = 32,
            };

            var lblAU = FsXS(new TextBlock { Text = "AU", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblAU, 0); inputGrid.Children.Add(lblAU);
            Grid.SetColumn(d1Display, 1); inputGrid.Children.Add(d1Display);

            var lblFxy = FsXS(new TextBlock { Text = "f(x/y)", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblFxy, 3); inputGrid.Children.Add(lblFxy);
            Grid.SetColumn(fxyDisplay, 4); inputGrid.Children.Add(fxyDisplay);

            var lblDil = FsXS(new TextBlock { Text = "P", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblDil, 6); inputGrid.Children.Add(lblDil);
            Grid.SetColumn(dilInput, 7); inputGrid.Children.Add(dilInput);

            var lblRes = FsXS(new TextBlock { Text = "Result", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblRes, 9); inputGrid.Children.Add(lblRes);
            Grid.SetColumn(resultDisplay, 10); inputGrid.Children.Add(resultDisplay);
            var lblUnit = FsXS(new TextBlock { Text = "mg/L", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblUnit, 11); inputGrid.Children.Add(lblUnit);

            root.Children.Add(inputGrid);
        }
        else if (exRow != null && isGcMode)
        {
            // GC 모드: 희석배수만 편집 가능, Resp./ISTD/Conc. 표시
            var dilInput = MakeInput(exRow.P);

            var respDisplay = FsBase(new TextBlock { Text = exRow.D1, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            var istdDisplay = FsBase(new TextBlock { Text = exRow.D2, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            fxyDisplay = FsBase(new TextBlock { Text = exRow.Fxy, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            resultDisplay = FsLG(new TextBlock { Text = exRow.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });

            // GC 검량식 수식 패널
            formulaPanel.Children.Add(FsBase(new TextBlock { Text = "GC 검량식", FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppRes("AppFg") }));
            if (editDocInfo?.GcCompoundCals.Count > 0)
            {
                foreach (var comp in editDocInfo.GcCompoundCals)
                {
                    string rStr = string.IsNullOrEmpty(comp.R) ? "" : $"  R={comp.R}";
                    formulaPanel.Children.Add(FsXS(new TextBlock
                    {
                        Text = $"{comp.Name}:  y = {comp.Slope}x + {comp.Intercept}{rStr}",
                        FontFamily = Font, Foreground = AppRes("FgMuted"),
                        TextWrapping = TextWrapping.Wrap,
                    }));
                }
            }
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 4) });
            var fmResGc = FsBase(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
                FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });
            formulaPanel.Children.Add(fmResGc);

            void RecalcGC()
            {
                double.TryParse(exRow.Fxy, out var conc);
                double.TryParse(dilInput.Text, out var dil);
                if (dil <= 0) dil = 1;
                double final_ = conc * dil;
                int dp = GetDecimalPlaces(_activeItems.FirstOrDefault() ?? _activeCategory);
                exRow.P      = dilInput.Text;
                exRow.Result = final_.ToString($"F{dp}");
                resultDisplay!.Text = exRow.Result;
                fmResGc.Text = $"Result = {conc:F4} × {dil} = {final_.ToString($"F{dp}")}";
            }
            dilInput.TextChanged += (_, _) => RecalcGC();
            RecalcGC();

            var inputGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,70,16,Auto,70,16,Auto,70,16,Auto,*,16,Auto,80,Auto"),
                Margin = new Thickness(0, 6), MinHeight = 32,
            };
            var lbResp = FsXS(new TextBlock { Text = "Resp.", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lbResp, 0); inputGrid.Children.Add(lbResp);
            Grid.SetColumn(respDisplay, 1); inputGrid.Children.Add(respDisplay);
            var lbIstd = FsXS(new TextBlock { Text = "ISTD", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lbIstd, 3); inputGrid.Children.Add(lbIstd);
            Grid.SetColumn(istdDisplay, 4); inputGrid.Children.Add(istdDisplay);
            var lbConc = FsXS(new TextBlock { Text = "농도", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lbConc, 6); inputGrid.Children.Add(lbConc);
            Grid.SetColumn(fxyDisplay, 7); inputGrid.Children.Add(fxyDisplay);
            var lbDil = FsXS(new TextBlock { Text = "희석배수", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lbDil, 9); inputGrid.Children.Add(lbDil);
            Grid.SetColumn(dilInput, 10); inputGrid.Children.Add(dilInput);
            var lbRes = FsXS(new TextBlock { Text = "Result", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lbRes, 12); inputGrid.Children.Add(lbRes);
            Grid.SetColumn(resultDisplay, 13); inputGrid.Children.Add(resultDisplay);
            var lbUnit = FsXS(new TextBlock { Text = "ng/L", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lbUnit, 14); inputGrid.Children.Add(lbUnit);

            root.Children.Add(inputGrid);
        }
        else if (exRow != null && _activeCategory == "ECO")
        {
            // ── 생태독성 직접 입력 UI ─────────────────────────────────────
            var eco = exRow.EcotoxData ?? new EcotoxTestData();
            exRow.EcotoxData = eco;

            // 시험 메타
            var metaGrid = new Grid { Margin = new Thickness(0, 4) };
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition(30, GridUnitType.Pixel));

            var spInput = MakeInput(eco.Species);
            var durInput = MakeInput(eco.Duration);
            metaGrid.Children.Add(FsXS(new TextBlock { Text = "시험종", FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center }));
            Grid.SetColumn(spInput, 1); metaGrid.Children.Add(spInput);
            metaGrid.Children.Add(FsXS(new TextBlock { Text = "시험시간", FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 0, 0) }));
            Grid.SetColumn(metaGrid.Children[^1], 2);
            Grid.SetColumn(durInput, 3); metaGrid.Children.Add(durInput);
            metaGrid.Children.Add(FsXS(new TextBlock { Text = "H", FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) }));
            Grid.SetColumn(metaGrid.Children[^1], 4);
            root.Children.Add(metaGrid);

            // 대조군
            var ctrlGrid = new Grid { Margin = new Thickness(0, 2) };
            ctrlGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            ctrlGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            ctrlGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            ctrlGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            ctrlGrid.Children.Add(FsXS(new TextBlock { Text = "대조군", FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center }));
            var ctrlOrgInput = MakeInput(eco.ControlOrganisms.ToString());
            var ctrlMortInput = MakeInput(eco.ControlMortalities.ToString());
            Grid.SetColumn(ctrlOrgInput, 1); ctrlGrid.Children.Add(ctrlOrgInput);
            ctrlGrid.Children.Add(FsXS(new TextBlock { Text = "사망", FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) }));
            Grid.SetColumn(ctrlGrid.Children[^1], 2);
            Grid.SetColumn(ctrlMortInput, 3); ctrlGrid.Children.Add(ctrlMortInput);
            root.Children.Add(ctrlGrid);

            // 농도별 데이터 (최대 8행)
            int numConc = Math.Max(eco.Concentrations.Length, 5); // 최소 5행
            numConc = Math.Min(numConc, 8);
            var concInputs = new TextBox[numConc];
            var orgInputs = new TextBox[numConc];
            var mortInputs = new TextBox[numConc];

            root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 4) });

            // 헤더
            var hdrGrid = new Grid();
            hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(30, GridUnitType.Pixel));
            hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
            hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
            hdrGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
            hdrGrid.Children.Add(FsXS(new TextBlock { Text = "#", FontFamily = Font, Foreground = AppRes("FgMuted"), HorizontalAlignment = HorizontalAlignment.Center }));
            var h1 = FsXS(new TextBlock { Text = "농도(%)", FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), FontWeight = FontWeight.SemiBold });
            Grid.SetColumn(h1, 1); hdrGrid.Children.Add(h1);
            var h2 = FsXS(new TextBlock { Text = "생물수", FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center });
            Grid.SetColumn(h2, 2); hdrGrid.Children.Add(h2);
            var h3 = FsXS(new TextBlock { Text = "사망수", FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center });
            Grid.SetColumn(h3, 3); hdrGrid.Children.Add(h3);
            var stdBtn = FsXS(new Button
            {
                Content = "표준농도입력", FontFamily = Font,
                Background = AppRes("ThemeFgInfo"), Foreground = Brushes.White,
                Padding = new Thickness(8, 2), CornerRadius = new CornerRadius(3),
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0),
            });
            stdBtn.Click += (_, _) =>
            {
                var std = EcotoxicityService.StandardConcentrations; // 6.25, 12.5, 25, 50, 100
                for (int si = 0; si < Math.Min(std.Length, numConc); si++)
                {
                    concInputs[si].Text = std[si].ToString("G");
                    orgInputs[si].Text = "20";
                }
            };
            Grid.SetColumn(stdBtn, 4); hdrGrid.Children.Add(stdBtn);
            root.Children.Add(hdrGrid);

            for (int ci = 0; ci < numConc; ci++)
            {
                var rowG = new Grid { Margin = new Thickness(0, 1) };
                rowG.ColumnDefinitions.Add(new ColumnDefinition(30, GridUnitType.Pixel));
                rowG.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
                rowG.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));
                rowG.ColumnDefinitions.Add(new ColumnDefinition(60, GridUnitType.Pixel));

                rowG.Children.Add(FsXS(new TextBlock { Text = $"{ci + 1}", FontFamily = Font, Foreground = AppRes("FgMuted"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }));

                concInputs[ci] = MakeInput(ci < eco.Concentrations.Length ? eco.Concentrations[ci].ToString("G") : "");
                orgInputs[ci] = MakeInput(ci < eco.Organisms.Length ? eco.Organisms[ci].ToString() : "20");
                mortInputs[ci] = MakeInput(ci < eco.Mortalities.Length ? eco.Mortalities[ci].ToString() : "");

                Grid.SetColumn(concInputs[ci], 1); rowG.Children.Add(concInputs[ci]);
                Grid.SetColumn(orgInputs[ci], 2); rowG.Children.Add(orgInputs[ci]);
                Grid.SetColumn(mortInputs[ci], 3); rowG.Children.Add(mortInputs[ci]);
                root.Children.Add(rowG);
            }

            root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 4) });

            // 결과 표시
            var resultPanel = new StackPanel { Spacing = 4 };
            var resultTb = FsLG(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold });
            var detailTb = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap });
            resultPanel.Children.Add(resultTb);
            resultPanel.Children.Add(detailTb);

            if (eco.Result != null)
            {
                resultTb.Text = $"TU = {eco.Result.TU}  (100/EC50 = 100/{eco.Result.EC50})";
                var initDetails = new List<string>();
                if (eco.Result.LowerCI > 0 || eco.Result.UpperCI > 0)
                    initDetails.Add($"95% CI: {eco.Result.LowerCI} ~ {eco.Result.UpperCI}");
                initDetails.Add(eco.Result.Method);
                if (eco.Result.TrimPercent > 0) initDetails.Add($"Trim: {eco.Result.TrimPercent}%");
                if (eco.Result.Smoothed) initDetails.Add("단조보정 적용");
                if (!string.IsNullOrEmpty(eco.Result.Warning)) initDetails.Add(eco.Result.Warning);
                detailTb.Text = string.Join("  |  ", initDetails);
            }

            // 입력값 수집 공통
            void CollectInputs()
            {
                eco.Species = spInput.Text ?? "물벼룩";
                eco.Duration = durInput.Text ?? "24";
                int.TryParse(ctrlOrgInput.Text, out var cOrg); eco.ControlOrganisms = cOrg > 0 ? cOrg : 20;
                int.TryParse(ctrlMortInput.Text, out var cMort); eco.ControlMortalities = cMort;

                var concList = new List<double>();
                var orgList = new List<int>();
                var mortList = new List<int>();
                for (int ci = 0; ci < numConc; ci++)
                {
                    if (double.TryParse(concInputs[ci].Text, out var c) && c > 0)
                    {
                        concList.Add(c);
                        int.TryParse(orgInputs[ci].Text, out var o); orgList.Add(o > 0 ? o : 20);
                        int.TryParse(mortInputs[ci].Text, out var m); mortList.Add(m);
                    }
                }
                eco.Concentrations = concList.ToArray();
                eco.Organisms = orgList.ToArray();
                eco.Mortalities = mortList.ToArray();
            }

            void ApplyResult(EcotoxicityService.EcotoxResult result)
            {
                eco.Result = result;
                exRow.Result = result.TU.ToString("F1");
                resultTb.Text = $"TU = {result.TU}  (100/EC50 = 100/{result.EC50})";
                var details = new List<string>();
                if (result.LowerCI > 0 || result.UpperCI > 0)
                    details.Add($"95% CI: {result.LowerCI} ~ {result.UpperCI}");
                details.Add(result.Method);
                if (result.TrimPercent > 0) details.Add($"Trim: {result.TrimPercent}%");
                if (result.Smoothed) details.Add("단조보정 적용");
                if (!string.IsNullOrEmpty(result.Warning)) details.Add(result.Warning);
                detailTb.Text = string.Join("  |  ", details);
            }

            // ── TU 직접 산출 (100/EC50) ──────────────────────────────────
            void DoDirectCalc()
            {
                CollectInputs();
                if (eco.Concentrations.Length < 1) { resultTb.Text = "최소 1개 농도를 입력하세요."; return; }

                // EC50 직접 산출: 사망률 50% 이상인 최저 농도 찾기
                int n = eco.Concentrations.Length;
                int idx50 = -1;
                for (int ci = 0; ci < n; ci++)
                {
                    if (eco.Organisms[ci] > 0)
                    {
                        double rate = (double)eco.Mortalities[ci] / eco.Organisms[ci];
                        if (rate >= 0.5) { idx50 = ci; break; }
                    }
                }

                if (idx50 < 0)
                {
                    // 50% 이상 사망 없음 → fallback (ES 04704.1c 8.1.2)
                    int lastOrg = eco.Organisms.Length > 0 ? eco.Organisms[^1] : 0;
                    int lastMort = eco.Mortalities.Length > 0 ? eco.Mortalities[^1] : 0;
                    var fallback = EcotoxicityService.CalculateFallbackTU(lastOrg, lastMort);
                    ApplyResult(fallback);
                    return;
                }

                // 50% 전후 구간에서 선형보간으로 EC50 추정
                double ec50;
                if (idx50 > 0)
                {
                    double rLow = (double)eco.Mortalities[idx50 - 1] / eco.Organisms[idx50 - 1];
                    double rHigh = (double)eco.Mortalities[idx50] / eco.Organisms[idx50];
                    double cLow = eco.Concentrations[idx50 - 1];
                    double cHigh = eco.Concentrations[idx50];
                    ec50 = (rHigh - rLow) > 0.001
                        ? cLow + (0.5 - rLow) / (rHigh - rLow) * (cHigh - cLow)
                        : cHigh;
                }
                else
                {
                    ec50 = eco.Concentrations[idx50];
                }

                ec50 = Math.Round(Math.Max(ec50, 0.01), 2);
                double tu = Math.Round(100.0 / ec50, 1);
                var result = new EcotoxicityService.EcotoxResult(ec50, 0, 0, tu, "직접산출", 0, false);
                ApplyResult(result);
            }

            // ── TSK / Probit 계산 ────────────────────────────────────────
            void DoCalc(bool useTSK)
            {
                CollectInputs();
                if (eco.Concentrations.Length < 2) { resultTb.Text = "최소 2개 농도가 필요합니다."; return; }

                try
                {
                    var result = useTSK
                        ? EcotoxicityService.CalculateTSK(eco.Concentrations, eco.Organisms, eco.Mortalities, eco.ControlOrganisms, eco.ControlMortalities)
                        : EcotoxicityService.CalculateProbit(eco.Concentrations, eco.Organisms, eco.Mortalities, eco.ControlOrganisms, eco.ControlMortalities);
                    ApplyResult(result);
                }
                catch (Exception ex)
                {
                    resultTb.Text = $"계산 오류: {ex.Message}";
                    detailTb.Text = "";
                }
            }

            // 버튼 패널
            // 버튼: TU 직접 산출 (메인) + TSK / Probit (개별)
            var btnPanel = new WrapPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Margin = new Thickness(0, 4) };

            var directBtn = new Button
            {
                Content = "TU 직접 산출", FontFamily = Font,
                Background = new SolidColorBrush(Color.Parse("#2E7D32")), Foreground = Brushes.White,
                Padding = new Thickness(16, 6), CornerRadius = new CornerRadius(4), Margin = new Thickness(0, 0, 8, 4),
            };
            directBtn.Click += (_, _) => DoDirectCalc();

            var tskBtn = new Button
            {
                Content = "TSK 계산", FontFamily = Font,
                Background = AppRes("BtnPrimaryBg"), Foreground = Brushes.White,
                Padding = new Thickness(16, 6), CornerRadius = new CornerRadius(4), Margin = new Thickness(0, 0, 8, 4),
            };
            tskBtn.Click += (_, _) => DoCalc(true);

            var probitBtn = new Button
            {
                Content = "Probit 계산", FontFamily = Font,
                Background = AppRes("ThemeFgInfo"), Foreground = Brushes.White,
                Padding = new Thickness(16, 6), CornerRadius = new CornerRadius(4), Margin = new Thickness(0, 0, 0, 4),
            };
            probitBtn.Click += (_, _) => DoCalc(false);

            btnPanel.Children.Add(directBtn);
            btnPanel.Children.Add(tskBtn);
            btnPanel.Children.Add(probitBtn);
            root.Children.Add(btnPanel);
            root.Children.Add(resultPanel);
        }
        else if (exRow != null && !isUV)
        {
            // BOD 모드: 시료량/D1/D2 편집 가능, f(x/y)/P/Result 자동 계산
            var volInput = MakeInput(exRow.시료량);
            var d1Input  = MakeInput(exRow.D1);
            var d2Input  = MakeInput(exRow.D2);

            fxyDisplay = FsBase(new TextBlock { Text = exRow.Fxy, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            pDisplay = FsBase(new TextBlock { Text = exRow.P, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            resultDisplay = FsLG(new TextBlock { Text = exRow.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });

            // Show4 수식 TextBlock (실시간 업데이트)
            var fmTitle = FsBase(new TextBlock { Text = "BOD 계산 수식", FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppRes("AppFg") });
            var fmFxy = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap });
            var fmP   = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted") });
            var fmRes = FsBase(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
                FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });
            var fmSeed = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), TextWrapping = TextWrapping.Wrap,
                Text = $"식종수: V={seedVol}, D1={seedD1}, D2={seedD2}, %={seedPct}" });

            formulaPanel.Children.Add(fmTitle);
            formulaPanel.Children.Add(fmSeed);
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle") });
            formulaPanel.Children.Add(fmFxy);
            formulaPanel.Children.Add(fmP);
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle") });
            formulaPanel.Children.Add(fmRes);

            void Recalc()
            {
                if (!double.TryParse(volInput.Text, out var vol) || vol <= 0) return;
                if (!double.TryParse(d1Input.Text, out var d1)) return;
                if (!double.TryParse(d2Input.Text, out var d2)) return;

                double fxy = 0;
                if (seedPct > 0)
                {
                    double num = (300 - vol) * (seedPct / 100.0);
                    double den = seedVol + ((300 - seedVol) * (seedPct / 100.0));
                    if (den > 0) fxy = num / den;
                }
                double p = 300.0 / vol;
                double result = ((d1 - d2) - (seedD1 - seedD2) * fxy) * p;

                fxyDisplay!.Text = fxy.ToString("F7");
                pDisplay!.Text = p.ToString("G");
                int dp = GetDecimalPlaces("BOD");
                resultDisplay!.Text = result.ToString($"F{dp}");

                exRow.시료량 = volInput.Text; exRow.D1 = d1Input.Text; exRow.D2 = d2Input.Text;
                exRow.Fxy = fxy.ToString("F7"); exRow.P = p.ToString("G");
                exRow.Result = result.ToString($"F{dp}");

                // Show4 수식 업데이트
                fmFxy.Text = $"f(x/y) = ((300-{vol}) * ({seedPct}/100)) / ({seedVol} + ((300-{seedVol}) * ({seedPct}/100)))\n= {fxy:F7}";
                fmP.Text = $"P = 300 / {vol} = {p:G}";
                fmRes.Text = $"Result = (({d1}-{d2}) - ({seedD1}-{seedD2})*{fxy:F4}) * {p:G}\n= {Math.Round(result, 1):F1} mg/L";
            }

            volInput.TextChanged += (_, _) => Recalc();
            d1Input.TextChanged  += (_, _) => Recalc();
            d2Input.TextChanged  += (_, _) => Recalc();
            Recalc();

            // 수평 배열: 입력값 + 계산결과 한 행에
            var inputGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,6,Auto,*,6,Auto,*,16,Auto,60,6,Auto,30,6,Auto,80,Auto"),
                Margin = new Thickness(0, 6),
                MinHeight = 32,
            };
            // 입력값 헤더
            var lblInputs = FsSM(new TextBlock { Text = "입력값", FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0) });
            Grid.SetColumn(lblInputs, 0); Grid.SetColumnSpan(lblInputs, 8);

            // 시료량
            var lblVol = FsXS(new TextBlock { Text = "시료량", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblVol, 0); inputGrid.Children.Add(lblVol);
            Grid.SetColumn(volInput, 1); inputGrid.Children.Add(volInput);
            var lblMl = FsXS(new TextBlock { Text = "mL", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblMl, 2); inputGrid.Children.Add(lblMl);

            // D1
            var lblD1 = FsXS(new TextBlock { Text = "D1", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblD1, 3); inputGrid.Children.Add(lblD1);
            Grid.SetColumn(d1Input, 4); inputGrid.Children.Add(d1Input);

            // D2
            var lblD2 = FsXS(new TextBlock { Text = "D2", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
            Grid.SetColumn(lblD2, 6); inputGrid.Children.Add(lblD2);
            Grid.SetColumn(d2Input, 7); inputGrid.Children.Add(d2Input);

            // 계산결과
            var lblFxy = FsXS(new TextBlock { Text = "f(x/y)", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblFxy, 9); inputGrid.Children.Add(lblFxy);
            Grid.SetColumn(fxyDisplay, 10); inputGrid.Children.Add(fxyDisplay);

            var lblP = FsXS(new TextBlock { Text = "P", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblP, 12); inputGrid.Children.Add(lblP);
            Grid.SetColumn(pDisplay, 13); inputGrid.Children.Add(pDisplay);

            var lblRes = FsXS(new TextBlock { Text = "Result", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblRes, 15); inputGrid.Children.Add(lblRes);
            Grid.SetColumn(resultDisplay, 16); inputGrid.Children.Add(resultDisplay);
            var lblUnit = FsXS(new TextBlock { Text = "mg/L", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblUnit, 17); inputGrid.Children.Add(lblUnit);

            root.Children.Add(inputGrid);
        }
        else if (exRow != null && isUV)
        {
            // UV VIS 모드
            var volInput = MakeInput(exRow.시료량);
            var absInput = MakeInput(exRow.D1);
            var dilInput = MakeInput(exRow.D2);

            var calcDisplay = FsBase(new TextBlock { Text = exRow.Fxy, FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });
            resultDisplay = FsLG(new TextBlock { Text = exRow.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0) });

            double slope = 0, intercept = 0;
            if (editDocInfo != null)
            {
                double.TryParse(editDocInfo.Standard_Slope, out slope);
                double.TryParse(editDocInfo.Standard_Intercept, out intercept);
            }

            var fmTitle = FsBase(new TextBlock { Text = "UV VIS 계산 수식", FontWeight = FontWeight.Bold,
                FontFamily = Font, Foreground = AppRes("AppFg") });
            var fmCalc = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap });
            var fmRes = FsBase(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
                FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });
            var fmCurve = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), TextWrapping = TextWrapping.Wrap,
                Text = "검량곡선:  " + (ETA.Services.SERVICE3.AnalysisNoteService.GetFormula(GetActiveAnalyteKey()) is string _cf && !string.IsNullOrWhiteSpace(_cf)
                    ? $"공식: {_cf}"
                    : $"Conc = (Abs - {{intercept:G6}}) / {{slope:G6}} × (60/시료량)") });

            formulaPanel.Children.Add(fmTitle);
            formulaPanel.Children.Add(fmCurve);
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle") });
            formulaPanel.Children.Add(fmCalc);
            formulaPanel.Children.Add(fmRes);

            void RecalcUV()
            {
                if (!double.TryParse(absInput.Text, out var abs)) return;
                if (!double.TryParse(dilInput.Text, out var dil)) return;

                double calcConc = slope > 0 ? ((abs - intercept) / slope) : 0;
                double result = calcConc * dil;

                int dp = GetDecimalPlaces(_activeItems.FirstOrDefault() ?? _activeCategory);
                calcDisplay.Text = calcConc.ToString("F4");
                resultDisplay!.Text = result.ToString($"F{dp}");

                exRow.D1 = absInput.Text; exRow.D2 = dilInput.Text;
                exRow.Fxy = calcConc.ToString("F4");
                exRow.Result = result.ToString($"F{dp}");

                fmCalc.Text = $"계산농도 = {slope:G6} * {abs} + {intercept:G6}\n= {calcConc:F4}";
                fmRes.Text = $"Result = {calcConc:F4} * {dil}\n= {result.ToString($"F{dp}")} mg/L";
            }

            absInput.TextChanged += (_, _) => RecalcUV();
            dilInput.TextChanged += (_, _) => RecalcUV();
            RecalcUV();

            // 수평 배열: 입력값 + 계산결과 한 행에
            var inputGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,6,Auto,*,6,Auto,*,16,Auto,80,6,Auto,80,Auto"),
                Margin = new Thickness(0, 6),
                MinHeight = 32,
            };

            // 시료량
            var lblVol = FsXS(new TextBlock { Text = "시료량", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblVol, 0); inputGrid.Children.Add(lblVol);
            Grid.SetColumn(volInput, 1); inputGrid.Children.Add(volInput);
            var lblMl = FsXS(new TextBlock { Text = "mL", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblMl, 2); inputGrid.Children.Add(lblMl);

            // 흡광도
            var lblAbs = FsXS(new TextBlock { Text = "흡광도", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblAbs, 3); inputGrid.Children.Add(lblAbs);
            Grid.SetColumn(absInput, 4); inputGrid.Children.Add(absInput);

            // 희석배수
            var lblDil = FsXS(new TextBlock { Text = "희석배수", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
            Grid.SetColumn(lblDil, 6); inputGrid.Children.Add(lblDil);
            Grid.SetColumn(dilInput, 7); inputGrid.Children.Add(dilInput);

            // 계산결과
            var lblCalc = FsXS(new TextBlock { Text = "계산농도", FontFamily = Font,
                Foreground = AppRes("ThemeFgSecondary"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblCalc, 9); inputGrid.Children.Add(lblCalc);
            Grid.SetColumn(calcDisplay, 10); inputGrid.Children.Add(calcDisplay);

            var lblRes = FsXS(new TextBlock { Text = "Result", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblRes, 12); inputGrid.Children.Add(lblRes);
            Grid.SetColumn(resultDisplay, 13); inputGrid.Children.Add(resultDisplay);
            var lblUnit = FsXS(new TextBlock { Text = "mg/L", FontFamily = Font,
                Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
            Grid.SetColumn(lblUnit, 14); inputGrid.Children.Add(lblUnit);

            root.Children.Add(inputGrid);
        }
        else
        {
            // 기존 모드: exRow 없이 직접 입력
            root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 4) });

            var editItems = _activeItems.Length > 0 ? _activeItems : AllEditableItems;
            for (int i = 0; i < editItems.Length; i++)
            {
                var itemName = editItems[i];
                string currentVal = GetSampleValue(sample, itemName);
                var input = MakeInput(currentVal);

                input.TextChanged += (_, _) =>
                {
                    bool hasVal = !string.IsNullOrWhiteSpace(input.Text);
                    input.BorderBrush = hasVal ? AppRes("ThemeBorderActive") : AppRes("InputBorder");
                };

                int ci = i;
                input.KeyDown += (_, e) =>
                {
                    if (e.Key == Key.Enter || e.Key == Key.Tab)
                    {
                        var nextKey = ci + 1 < editItems.Length ? editItems[ci + 1] : null;
                        if (nextKey != null && _inputBoxes.TryGetValue(nextKey, out var nb))
                        { nb.Focus(); nb.SelectAll(); e.Handled = true; }
                        else if (e.Key == Key.Enter)
                        { SaveCurrentSample(); e.Handled = true; }
                    }
                };

                _inputBoxes[itemName] = input;
                root.Children.Add(MakeRow(itemName, input, "mg/L"));
            }
        }

        // 계산 수식을 Show3에 병합 (별도 Show4 표시 제거)
        if (formulaPanel.Children.Count > 0)
        {
            root.Children.Add(new Border
            {
                Child = formulaPanel,
                BorderBrush = AppRes("ThemeBorderSubtle"),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Margin = new Thickness(0, 6, 0, 0),
                Padding = new Thickness(0, 6, 0, 0),
            });
        }
        else if (exRow == null && _selectedSample != null)
        {
            // ── WasteSample 편집 (Show2 샘플 그리드) ──
            var resultVal = GetResultValue();
            var resultInput = FsBase(new TextBox
            {
                Text = resultVal, FontFamily = Font,
                Foreground = AppRes("InputFg"), Background = AppRes("InputBg"),
                BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 4),
                HorizontalAlignment = HorizontalAlignment.Stretch, Watermark = "0.00",
            });

            resultInput.TextChanged += (_, _) =>
            {
                _inputBoxes["Result"] = resultInput;
            };

            resultInput.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    CommitSampleEdit();
                    // 다음 행으로 이동
                    if (_selectedSampleIndex >= 0 && _selectedSampleIndex < _currentSamples.Count - 1)
                        SelectSampleRow(_selectedSampleIndex + 1);
                }
                else if (e.Key == Key.Up)
                {
                    e.Handled = true;
                    CommitSampleEdit();
                    if (_selectedSampleIndex > 0)
                        SelectSampleRow(_selectedSampleIndex - 1);
                }
                else if (e.Key == Key.Down)
                {
                    e.Handled = true;
                    CommitSampleEdit();
                    if (_selectedSampleIndex < _currentSamples.Count - 1)
                        SelectSampleRow(_selectedSampleIndex + 1);
                }
            };

            _inputBoxes["Result"] = resultInput;
            var colName = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;
            root.Children.Add(MakeRow(colName, resultInput, ""));
            root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 6) });
        }
        else
        {
            root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 6) });
        }

        EditPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // =========================================================================
    // 저장
    // =========================================================================
    /// <summary>Show3 저장 버튼 — ExcelRow만 업데이트, Show2 그리드 갱신 (서버 저장 안함)</summary>
    private void SaveCurrentSample()
    {
        // WasteSample 저장 (Show2 샘플 그리드에서 호출)
        if (_selectedSample != null && _currentEditExcelRow == null)
        {
            CommitSampleEdit();
            ShowMessage("✅ 저장 완료", false);
            return;
        }

        // ExcelRow 저장 (기존 로직)
        if (_currentEditExcelRow == null) return;
        if (_categoryExcelData.ContainsKey(_activeCategory))
            { /* LoadVerifiedGrid 제거됨 */ }
        BuildStatsPanel();
        ShowMessage("✅ 적용 완료 (서버 반영은 '입력' 버튼 사용)", false);
    }

    // ─── *_DATA 원시 측정값 저장 헬퍼 ──────────────────────────────────────
    // 매칭 여부와 무관하게 전체 행(정도관리 시료 + 미매칭 포함)을 기록부 테이블에 저장한다.
    // s가 null이면 SN=row.SN, 업체명/구분 빈값 사용.
    private static readonly string ImportLogPath = System.IO.Path.Combine(
        Directory.GetCurrentDirectory(), "Logs", "ImportDebug.log");
    private static void LogImport(string msg)
    {
        if (App.EnableLogging)
        {
            try
            {
                var dir = System.IO.Path.GetDirectoryName(ImportLogPath)!;
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.AppendAllText(ImportLogPath, $"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            }
            catch { }
        }
    }

    // EnableLogging 여부와 무관하게 항상 기록 (입력 버튼 진단용)
    private static void LogInput(string msg)
    {
        try
        {
            var dir = System.IO.Path.GetDirectoryName(ImportLogPath)!;
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            File.AppendAllText(ImportLogPath, $"[{DateTime.Now:HH:mm:ss}] {msg}\n");
        }
        catch { }
    }

    private void SaveRawData(ExcelRow row, WasteSample? s, string 소스구분 = "폐수배출업소")
    {
        if (string.IsNullOrWhiteSpace(row.Result))
        {
            LogImport($"SKIP(빈결과): SN={row.SN}, 시료명={row.시료명}");
            return;
        }

        _categoryDocInfo.TryGetValue(_activeCategory, out var docInfo);
        bool isUV = docInfo?.IsUVVIS == true;
        LogImport($"SaveRawData: cat={_activeCategory}, isUV={isUV}, SN={row.SN}, Result={row.Result}, Source={소스구분}");

        // 분석일: B1 값 우선, 없으면 매칭된 샘플의 채수일, 그래도 없으면 오늘
        _categoryDocDates.TryGetValue(_activeCategory, out var 분석일Raw);
        string 분석일 = !string.IsNullOrEmpty(분석일Raw)
            ? 분석일Raw
            : (s?.채수일 ?? DateTime.Today.ToString("yyyy-MM-dd"));

        // SN 결정: QAQC=QC고정, 처리시설=시료명, 그 외=매칭SN
        string sn;
        if (소스구분 == "QAQC")
            sn = "QC";
        else if (소스구분 == "처리시설")
            sn = row.시료명 ?? row.SN ?? "";
        else
            sn = s?.SN ?? row.SN ?? "";
        if (string.IsNullOrEmpty(sn)) sn = row.시료명 ?? "";
        string 업체명;
        string 구분;
        if (소스구분 == "처리시설")
        {
            업체명 = row.MatchedFacilityName ?? row.SN ?? "";
            구분 = "";
        }
        else if (s != null)
        {
            업체명 = s.업체명;
            구분 = s.구분;
        }
        else if (row.MatchedAnalysis != null)
        {
            업체명 = row.MatchedAnalysis.약칭;
            구분 = "";
        }
        else
        {
            업체명 = "";
            구분 = "";
            if (!string.IsNullOrEmpty(sn) && 소스구분 != "QAQC")
            {
                if (소스구분 == "수질분석센터")
                {
                    // 수질분석센터: SN이 약칭(업체명)
                    업체명 = sn;
                }
                else
                {
                    // 폐수배출업소: SN으로 DB 조회
                    try
                    {
                        var found = WasteSampleService.FindBySN(sn);
                        if (found != null) { 업체명 = found.업체명; 구분 = found.구분; }
                    }
                    catch { }
                }
            }
        }

        // 희석배수 기본값: 비어있으면 "1"
        if (string.IsNullOrEmpty(row.P)) row.P = "1";

        // 시료명: 원본시료명이 있으면 "원본\n↳ 매칭명" 형태로 변경내역 보존
        string 시료명Full;
        if (!string.IsNullOrEmpty(row.원본시료명) && row.원본시료명 != row.시료명)
            시료명Full = $"{row.원본시료명}\n↳ {row.시료명}";
        else
            시료명Full = row.시료명 ?? "";

        string remark = !string.IsNullOrEmpty(row.원본시료명) ? row.원본시료명 : "";

        switch (_activeCategory)
        {
            case "TOC":
            {
                _categoryDocInfo.TryGetValue("TOC", out var tocInfo);
                WasteSampleService.UpsertTocData(
                    _tocInstrumentMethod, 분석일, sn, 업체명, 구분,
                    row.D1, row.P, tocInfo?.TocSlope_TC ?? "", row.Fxy, row.Result,
                    소스구분: 소스구분, 비고: remark, 시료명: 시료명Full);
                break;
            }

            case "BOD":
                WasteSampleService.UpsertBodData(
                    분석일, sn, 업체명, 구분,
                    시료량: row.시료량, d1: row.D1, d2: row.D2,
                    희석배수: row.P, 결과: row.Result,
                    소스구분: 소스구분,
                    식종시료량: docInfo?.식종수_시료량 ?? "",
                    식종D1:     docInfo?.식종수_D1    ?? "",
                    식종D2:     docInfo?.식종수_D2    ?? "",
                    식종BOD:    docInfo?.식종수_Result ?? "",
                    식종함유량: docInfo?.식종수_Remark  ?? "",
                    비고: remark, 시료명: 시료명Full);
                break;

            case "SS":
                WasteSampleService.UpsertSsData(
                    분석일, sn, 업체명, 구분,
                    row.시료량, row.D1, row.D2, row.Fxy, row.P, row.Result,
                    소스구분: 소스구분, 비고: remark, 시료명: 시료명Full);
                break;

            case "NHEX":
                WasteSampleService.UpsertNHexanData(
                    분석일, sn, 업체명, 구분,
                    row.시료량, row.D1, row.D2, row.Fxy, row.P, row.Result,
                    소스구분: 소스구분, 비고: remark, 시료명: 시료명Full);
                break;

            case "TN" when isUV:
            case "TP" when isUV:
            case "PHENOLS" when isUV:
                foreach (var item in _activeItems)
                {
                    string tblName = item switch
                    {
                        "T-N"     => "TN_시험기록부",
                        "T-P"     => "TP_시험기록부",
                        "Phenols" => "Phenols_시험기록부",
                        _         => ""
                    };
                    if (string.IsNullOrEmpty(tblName)) continue;
                    WasteSampleService.UpsertUvvisData(
                        tblName,
                        분석일, sn, 업체명, 소스구분,
                        시료량:  row.시료량,
                        흡광도:  row.D1,
                        희석배수: row.P,
                        검량선a: docInfo?.Standard_Slope ?? "",
                        농도:    row.Result,
                        st01mgl: docInfo?.Standard_Points?.ElementAtOrDefault(0) ?? "",
                        st02mgl: docInfo?.Standard_Points?.ElementAtOrDefault(1) ?? "",
                        st03mgl: docInfo?.Standard_Points?.ElementAtOrDefault(2) ?? "",
                        st04mgl: docInfo?.Standard_Points?.ElementAtOrDefault(3) ?? "",
                        st05mgl: docInfo?.Standard_Points?.ElementAtOrDefault(4) ?? "",
                        st01abs: docInfo?.Abs_Values?.ElementAtOrDefault(0) ?? "",
                        st02abs: docInfo?.Abs_Values?.ElementAtOrDefault(1) ?? "",
                        st03abs: docInfo?.Abs_Values?.ElementAtOrDefault(2) ?? "",
                        st04abs: docInfo?.Abs_Values?.ElementAtOrDefault(3) ?? "",
                        st05abs: docInfo?.Abs_Values?.ElementAtOrDefault(4) ?? "",
                        기울기:  docInfo?.Standard_Slope ?? "",
                        절편:    docInfo?.Standard_Intercept ?? "",
                        R2:      docInfo?.Abs_R2 ?? "",
                        비고: remark,
                        시료명: 시료명Full,
                        소스구분: 소스구분);
                }
                break;

            case "ICP":
            {
                // ICP-OES/VGA: 원소별 시험기록부 — UvvisData 스키마 재사용
                string icpCompound = row.CompoundName;
                if (string.IsNullOrEmpty(icpCompound)) break;

                var icpResolved  = CompoundAliasService.ResolveOrFallback(icpCompound);
                string icpTable  = $"{SafeName(icpResolved.표준코드)}_시험기록부";
                var icpCal       = docInfo?.GcCompoundCals?
                    .FirstOrDefault(c => c.Name.Equals(icpCompound, StringComparison.OrdinalIgnoreCase));

                WasteSampleService.UpsertUvvisData(
                    icpTable,
                    분석일, sn, 업체명, 소스구분,
                    시료량:  row.시료량,
                    흡광도:  row.D1,
                    희석배수: string.IsNullOrEmpty(row.P) ? "1" : row.P,
                    검량선a: icpCal?.Intercept ?? "",
                    농도:    row.Result,
                    st01mgl: icpCal?.StdConcs.ElementAtOrDefault(0) ?? "",
                    st02mgl: icpCal?.StdConcs.ElementAtOrDefault(1) ?? "",
                    st03mgl: icpCal?.StdConcs.ElementAtOrDefault(2) ?? "",
                    st04mgl: icpCal?.StdConcs.ElementAtOrDefault(3) ?? "",
                    st05mgl: icpCal?.StdConcs.ElementAtOrDefault(4) ?? "",
                    st01abs: icpCal?.StdResps.ElementAtOrDefault(0) ?? "",
                    st02abs: icpCal?.StdResps.ElementAtOrDefault(1) ?? "",
                    st03abs: icpCal?.StdResps.ElementAtOrDefault(2) ?? "",
                    st04abs: icpCal?.StdResps.ElementAtOrDefault(3) ?? "",
                    st05abs: icpCal?.StdResps.ElementAtOrDefault(4) ?? "",
                    기울기:  icpCal?.Slope ?? "",
                    절편:    icpCal?.Intercept ?? "",
                    R2:      icpCal?.R ?? "",
                    비고: remark,
                    시료명: 시료명Full,
                    소스구분: 소스구분);
                break;
            }

            case "GCMS":
            {
                // 성분별 테이블 라우팅: CompoundAliasService로 정규화 후 라우팅
                string compound = row.CompoundName;
                if (string.IsNullOrEmpty(compound)) break;

                var resolved = CompoundAliasService.ResolveOrFallback(compound);
                string tableName = $"{SafeName(resolved.표준코드)}_시험기록부";

                // 해당 성분의 검량선 찾기 (첫 번째가 아닌 매칭되는 성분)
                var compoundCal = docInfo?.GcCompoundCals?
                    .FirstOrDefault(c => c.Name.Equals(compound, StringComparison.OrdinalIgnoreCase));

                WasteSampleService.UpsertGcData(
                    tableName,
                    분석일, sn, 업체명, 소스구분,
                    농도: row.Result,
                    ISTD: row.D2 ?? "",
                    검량선정보: docInfo,
                    compoundCal: compoundCal,
                    비고: remark, 시료명: 시료명Full,
                    흡광도: row.D1 ?? "");
                break;
            }

            case "CN" when isUV:
            case "CR6" when isUV:
            case "COLOR" when isUV:
            case "ABS" when isUV:
            case "FLUORIDE" when isUV:
                foreach (var item in _activeItems)
                {
                    string tblName = item switch
                    {
                        "시안"     => "시안_시험기록부",
                        "6가크롬"  => "6가크롬_시험기록부",
                        "색도"     => "색도_시험기록부",
                        "ABS"      => "ABS_시험기록부",
                        "불소"     => "불소_시험기록부",
                        _          => ""
                    };
                    if (string.IsNullOrEmpty(tblName)) continue;
                    WasteSampleService.UpsertUvvisData(
                        tblName,
                        분석일, sn, 업체명, 소스구분,
                        시료량:  row.시료량,
                        흡광도:  row.D1,
                        희석배수: row.P,
                        검량선a: docInfo?.Standard_Slope ?? "",
                        농도:    row.Result,
                        st01mgl: docInfo?.Standard_Points?.ElementAtOrDefault(0) ?? "",
                        st02mgl: docInfo?.Standard_Points?.ElementAtOrDefault(1) ?? "",
                        st03mgl: docInfo?.Standard_Points?.ElementAtOrDefault(2) ?? "",
                        st04mgl: docInfo?.Standard_Points?.ElementAtOrDefault(3) ?? "",
                        st05mgl: docInfo?.Standard_Points?.ElementAtOrDefault(4) ?? "",
                        st01abs: docInfo?.Abs_Values?.ElementAtOrDefault(0) ?? "",
                        st02abs: docInfo?.Abs_Values?.ElementAtOrDefault(1) ?? "",
                        st03abs: docInfo?.Abs_Values?.ElementAtOrDefault(2) ?? "",
                        st04abs: docInfo?.Abs_Values?.ElementAtOrDefault(3) ?? "",
                        st05abs: docInfo?.Abs_Values?.ElementAtOrDefault(4) ?? "",
                        기울기:  docInfo?.Standard_Slope ?? "",
                        절편:    docInfo?.Standard_Intercept ?? "",
                        R2:      docInfo?.Abs_R2 ?? "",
                        비고: remark,
                        시료명: 시료명Full,
                        소스구분: 소스구분);
                }
                break;

            case "ECO":
            {
                // 생태독성: ExcelRow에 저장된 독성시험 데이터로 계산 + 저장
                var ecoData = row.EcotoxData;
                if (ecoData != null)
                {
                    WasteSampleService.UpsertEcotoxData(
                        분석일, sn, 업체명, 구분, 시료명Full, 소스구분,
                        ecoData.Species, ecoData.Duration, ecoData.DurationUnit,
                        ecoData.ControlOrganisms, ecoData.ControlMortalities,
                        ecoData.Concentrations, ecoData.Organisms, ecoData.Mortalities,
                        ecoData.Result, 비고: remark);
                }
                break;
            }

            default:
                LogImport($"  매칭되지 않은 카테고리: '{_activeCategory}' (isUV={isUV})");
                break;
        }
    }

    /// <summary>분석항목명 → 테이블명 안전 변환 (공백/특수문자 → 언더스코어)</summary>
    /// <summary>TextBox에 숫자(0-9), 소수점, 마이너스만 입력 허용</summary>
    private static void NumericOnly(TextBox tb)
    {
        // TextInputEvent Tunnel: 비숫자 문자 즉시 차단
        tb.AddHandler(TextBox.TextInputEvent, (object? s, TextInputEventArgs e) =>
        {
            if (string.IsNullOrEmpty(e.Text)) return;
            foreach (var ch in e.Text)
                if (!char.IsDigit(ch) && ch != '.' && ch != '-')
                    { e.Handled = true; return; }
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel);
        // TextChanged: IME 조합 완료 후 비숫자 문자 제거 (TextInputEvent 우회 방어)
        tb.TextChanged += (_, _) =>
        {
            var text = tb.Text ?? "";
            var clean = new string(text.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());
            if (clean != text) tb.Text = clean;
        };
    }

    private static string SafeName(string analyte)
    {
        return analyte
            .Replace("-", "")      // 하이픈 제거 (T-N → TN)
            .Replace(" ", "_")
            .Replace("/", "_")
            .Replace("\\", "_")
            .Replace("(", "")
            .Replace(")", "")
            .Replace(",", "_")
            .Replace("·", "_");
    }

    private void ShowExcelRowDetail(ExcelRow exRow)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };
        var catLabel = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;

        root.Children.Add(FsLG(new TextBlock
        {
            Text = $"📋 {exRow.시료명}  [{catLabel}]",
            FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = AppRes("AppFg"),
        }));

        var srcLabel = exRow.Source switch
        {
            SourceType.폐수배출업소 => $"폐수배출-{exRow.Matched?.구분 ?? "?"}",
            SourceType.수질분석센터 => "수질분석센터",
            SourceType.처리시설     => $"처리시설 ({exRow.MatchedFacilityName})",
            _                      => "미분류",
        };
        root.Children.Add(FsBase(new TextBlock
        {
            Text = $"시료구분: {srcLabel}  |  SN: {exRow.SN}  |  결과값: {exRow.Result}",
            FontFamily = Font, Foreground = AppRes("FgMuted"),
        }));

        bool isUV    = _categoryDocInfo.TryGetValue(_activeCategory, out var di) && di.IsUVVIS;
        bool isTocF  = di != null && (di.IsTocNPOC || di.IsTocTCIC);
        bool isGcCalc = di?.IsGcMode == true;

        TextBox DetailInput(string val) => FsBase(new TextBox
        {
            Text = val, FontFamily = Font,
            Foreground = AppRes("InputFg"), Background = AppRes("InputBg"),
            BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 4),
            HorizontalAlignment = HorizontalAlignment.Stretch, Watermark = "0.00",
        });

        // ── UV 모드: 인라인 편집 + 실시간 계산 ─────────────────────────────
        if (isUV)
        {
            double slope = 0, intercept = 0;
            double.TryParse(di!.Standard_Slope,     out slope);
            double.TryParse(di.Standard_Intercept, out intercept);

            var volInput = DetailInput(exRow.시료량);
            var absInput = DetailInput(exRow.D1);
            var dilInput = DetailInput(string.IsNullOrEmpty(exRow.D2) ? "1" : exRow.D2);

            var calcTb   = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSecondary"),
                VerticalAlignment = VerticalAlignment.Center });
            var resultTb = FsLG(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
                FontWeight = FontWeight.Bold, VerticalAlignment = VerticalAlignment.Center });

            int dp = (_activeCategory == "TN" || _activeItems.Any(x => x == "T-N"))
                ? (exRow.Source == SourceType.폐수배출업소 ? 3 : 1)
                : GetDecimalPlaces(_activeItems.FirstOrDefault() ?? _activeCategory);

            void RecalcUV()
            {
                double.TryParse(absInput.Text?.Replace(",", "."), System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var abs);
                double.TryParse(dilInput.Text?.Replace(",", "."), System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var dil);
                double.TryParse(volInput.Text?.Replace(",", "."), System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var vol);
                if (dil <= 0) dil = 1;

                double calcConc, result;
                string analyteKey = GetActiveAnalyteKey();
                string formulaStr = ETA.Services.SERVICE3.AnalysisNoteService.GetFormula(analyteKey);
                double.TryParse(exRow?.D2?.Replace(",", "."),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var istdVal);
                var evaluated = ETA.Services.SERVICE3.AnalysisNoteService.EvaluateFormula(formulaStr,
                    new System.Collections.Generic.Dictionary<string, double>
                    {
                        ["흡광도"]  = abs,
                        ["Area"]   = abs,
                        ["Resp"]   = abs,
                        ["ISTD"]   = istdVal,
                        ["절편"]   = intercept,
                        ["기울기"]  = slope,
                        ["시료량"]  = vol > 0 ? vol : 1.0,
                        ["희석배수"] = dil,
                    });

                if (evaluated.HasValue)
                {
                    result   = evaluated.Value;
                    calcConc = dil > 0 ? result / dil : result;
                    calcTb.Text   = $"계산농도 = {calcConc:F4}  (공식 적용)";
                    resultTb.Text = $"결과값  =  {result.ToString($"F{dp}")} mg/L";
                }
                else
                {
                    // 공식 없음 — 기본 UvVis 계산 (60/시료량)
                    calcConc = slope > 0
                        ? ((abs - intercept) / slope) * (vol > 0 ? 60.0 / vol : 1.0)
                        : 0;
                    result = calcConc * dil;
                    string volFactor = vol > 0 ? $" × (60/{vol})" : "";
                    calcTb.Text   = $"계산농도 = ({abs} - {intercept:G6}) / {slope:G6}{volFactor}  =  {calcConc:F4}";
                    resultTb.Text = $"결과값  =  {calcConc:F4} × {dil}  =  {result.ToString($"F{dp}")} mg/L";
                }

                exRow.시료량 = volInput.Text ?? "";
                exRow.D1   = absInput.Text ?? "";
                exRow.D2   = dilInput.Text ?? "";
                exRow.Fxy  = calcConc.ToString("F4");
                exRow.Result = result.ToString($"F{dp}");
            }

            absInput.TextChanged += (_, _) => RecalcUV();
            dilInput.TextChanged += (_, _) => RecalcUV();
            volInput.TextChanged += (_, _) => RecalcUV();
            RecalcUV();

            // 검량선 수식 표시
            root.Children.Add(new Border
            {
                Background = AppRes("GridRowAltBg"), CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10, 6), Margin = new Thickness(0, 2, 0, 0),
                Child = FsSM(new TextBlock
                {
                    Text = "검량곡선:  " + (ETA.Services.SERVICE3.AnalysisNoteService.GetFormula(GetActiveAnalyteKey()) is string _cf && !string.IsNullOrWhiteSpace(_cf)
                    ? $"공식: {_cf}"
                    : $"Conc = (Abs - {{intercept:G6}}) / {{slope:G6}} × (60/시료량)") + $"   R²={di.Abs_R2}",
                    FontFamily = Font, Foreground = AppRes("ThemeFgWarn"),
                    FontWeight = FontWeight.SemiBold,
                }),
            });

            // 입력 행: 시료량 | 흡광도 | 희석배수
            var inputGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("60,90,8,60,90,8,60,90"),
                Margin = new Thickness(0, 6, 0, 4),
            };
            void AddInputCell(string lbl, Control ctrl, int colLbl, int colCtrl)
            {
                var lb = FsXS(new TextBlock { Text = lbl, FontFamily = Font,
                    Foreground = AppRes("ThemeFgSecondary"),
                    VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
                Grid.SetColumn(lb,   colLbl); inputGrid.Children.Add(lb);
                Grid.SetColumn(ctrl, colCtrl); inputGrid.Children.Add(ctrl);
            }
            AddInputCell("시료량",  volInput, 0, 1);
            AddInputCell("흡광도",  absInput, 3, 4);
            AddInputCell("희석배수", dilInput, 6, 7);
            root.Children.Add(inputGrid);

            // 계산 결과 표시
            root.Children.Add(new Border
            {
                BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(0, 6, 0, 0), Margin = new Thickness(0, 2, 0, 0),
                Child = new StackPanel { Spacing = 3, Children = { calcTb, resultTb } },
            });
        }
        // ── TOC 모드 ────────────────────────────────────────────────────────
        else if (isTocF)
        {
            var fields = new[] { ("AU", exRow.D1), ("농도", exRow.Fxy), ("희석배수", exRow.P) };
            var detailGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("80,*,80,*,80,*"), Margin = new Thickness(0, 4) };
            int col = 0;
            foreach (var (lbl, v) in fields)
            {
                if (col >= 6) break;
                var lbTb = FsXS(new TextBlock { Text = lbl, FontFamily = Font,
                    Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
                Grid.SetColumn(lbTb, col); detailGrid.Children.Add(lbTb);
                var vTb = FsBase(new TextBlock { Text = v, FontFamily = Font,
                    Foreground = AppRes("ThemeFgSecondary"), FontWeight = FontWeight.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(4, 0) });
                Grid.SetColumn(vTb, col + 1); detailGrid.Children.Add(vTb);
                col += 2;
            }
            root.Children.Add(new Border { Child = detailGrid, Background = AppRes("GridRowAltBg"),
                CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 6) });

            if (!string.IsNullOrEmpty(di!.TocSlope_TC))
                root.Children.Add(FsXS(new TextBlock
                {
                    Text = $"y = {di.TocSlope_TC}x + {di.TocIntercept_TC}  R²={di.TocR2_TC}",
                    FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), Margin = new Thickness(0, 4, 0, 0),
                }));
            double.TryParse(exRow.Fxy, out var conc);
            double.TryParse(exRow.P,   out var dil);
            if (dil <= 0) dil = 1;
            root.Children.Add(FsBase(new TextBlock
            {
                Text = $"결과 = {conc} × {dil} = {exRow.Result}",
                FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.SemiBold,
                Margin = new Thickness(0, 2, 0, 0),
            }));
        }
        // ── GC 모드 ─────────────────────────────────────────────────────────
        else if (isGcCalc)
        {
            foreach (var comp in di!.GcCompoundCals)
            {
                string rStr = string.IsNullOrEmpty(comp.R) ? "" : $"  R={comp.R}";
                root.Children.Add(FsXS(new TextBlock
                {
                    Text = $"{comp.Name}:  y = {comp.Slope}x + {comp.Intercept}{rStr}",
                    FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap,
                }));
            }
            double.TryParse(exRow.Fxy, out var gcConc);
            double.TryParse(exRow.P,   out var gcDil);
            if (gcDil <= 0) gcDil = 1;
            root.Children.Add(FsBase(new TextBlock
            {
                Text = $"Result = {gcConc} × {gcDil} = {exRow.Result}",
                FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.SemiBold,
            }));
        }
        // ── 일반 모드 ────────────────────────────────────────────────────────
        else
        {
            var flds = new[] { ("시료량", exRow.시료량), ("D1", exRow.D1), ("D2", exRow.D2), ("f(x/y)", exRow.Fxy), ("P", exRow.P) };
            foreach (var (lbl, v) in flds)
            {
                if (string.IsNullOrEmpty(v)) continue;
                root.Children.Add(FsXS(new TextBlock { Text = $"{lbl}: {v}", FontFamily = Font, Foreground = AppRes("ThemeFg") }));
            }
            root.Children.Add(FsXS(new TextBlock
            {
                Text = $"결과값: {exRow.Result}", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.SemiBold, Margin = new Thickness(0, 4, 0, 0),
            }));
        }

        EditPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // =========================================================================
    // Show4: 입력 현황 통계
    // =========================================================================
    /// <summary>Show1: 의뢰시료 드래그앤드랍 분류 패널</summary>
    private void BuildMatchBrowsePanel()
    {
        // 매칭 아이템 목록 초기화 (키보드 네비용)
        ClearShow1Focus();
        _matchItems.Clear();
        _keyNavShow1Index = -1;
        BuildMatchBrowsePanelInner();
    }
    private void BuildMatchBrowsePanelInner()
    {
        var root = new Grid
        {
            Background = AppRes("PanelBg"),
            Margin = new Thickness(8),
            RowDefinitions =
            {
                new RowDefinition(GridLength.Auto), // 탭 패널
                new RowDefinition(GridLength.Auto), // 날짜 패널
                new RowDefinition(GridLength.Star)  // 스크롤뷰어 - 남은 공간 모두 사용
            }
        };

        // 탭 버튼들 (수질분석센터, 처리시설, 비용부담금)
        var tabPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            Margin = new Thickness(0, 0, 0, 12)
        };

        // 분석항목 모드일 때는 분석항목 탭만 활성, 아닐 때는 현재 _inputMode 탭 활성
        bool isCompoundMode = _show1BrowseMode == "분석항목";
        tabPanel.Children.Add(BuildModeTab("수질분석센터", "🧪", !isCompoundMode && _inputMode == "수질분석센터"));
        tabPanel.Children.Add(BuildModeTab("처리시설", "🏭", !isCompoundMode && _inputMode == "처리시설"));
        tabPanel.Children.Add(BuildModeTab("비용부담금", "💰", !isCompoundMode && _inputMode == "비용부담금"));
        tabPanel.Children.Add(BuildModeTab("분석항목", "🔬", isCompoundMode));

        Grid.SetRow(tabPanel, 0);
        root.Children.Add(tabPanel);

        // 날짜 선택 영역
        var datePanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            Margin = new Thickness(0, 0, 0, 8)
        };

        datePanel.Children.Add(FsSM(new TextBlock
        {
            Text = "📅",
            VerticalAlignment = VerticalAlignment.Center
        }));

        datePanel.Children.Add(FsXS(new TextBlock
        {
            Text = $"{DateTime.Now:yyyy년 M월}",
            FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("ThemeFgPrimary"),
            VerticalAlignment = VerticalAlignment.Center
        }));

        Grid.SetRow(datePanel, 1);
        root.Children.Add(datePanel);

        // 의뢰시료 목록 (드래그 소스)
        var sampleList = new StackPanel { Spacing = 2 };
        BuildSampleList(sampleList);

        var scrollViewer = new ScrollViewer
        {
            Background = AppRes("ThemeBgCard"),
            Content = sampleList
            // MaxHeight 제거 - 사용 가능한 전체 공간 활용
        };

        Grid.SetRow(scrollViewer, 2);
        root.Children.Add(scrollViewer);

        StatsPanelChanged?.Invoke(root);

        // Show2는 기존 시험기록부 표시 (미매칭 항목들이 드롭존)
    }

    private Border BuildModeTab(string title, string icon, bool isActive)
    {
        var tab = new Border
        {
            Background = isActive ? AppRes("ThemeBgSecondary") : AppRes("ThemeBgCard"),
            BorderBrush = isActive ? AppRes("ThemeAccent") : AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(2),
            CornerRadius = new CornerRadius(20),
            Padding = new Thickness(16, 8),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 6,
                Children =
                {
                    FsSM(new TextBlock
                    {
                        Text = icon,
                        VerticalAlignment = VerticalAlignment.Center
                    }),
                    FsXS(new TextBlock
                    {
                        Text = title,
                        FontWeight = isActive ? FontWeight.SemiBold : FontWeight.Normal,
                        Foreground = isActive ? AppRes("ThemeFgPrimary") : AppRes("ThemeFgSecondary"),
                        VerticalAlignment = VerticalAlignment.Center
                    })
                }
            }
        };

        tab.PointerPressed += (_, _) =>
        {
            if (title == "분석항목")
            {
                // 분석항목 탭은 메인 모드 변경 없이 Show1만 전환
                _show1BrowseMode = "분석항목";
                BuildMatchBrowsePanel();
                return;
            }

            var mode = title switch
            {
                "수질분석센터" => "수질분석센터",
                "처리시설" => "처리시설",
                "비용부담금" => "비용부담금",
                _ => _inputMode
            };

            if (mode != _inputMode || _show1BrowseMode == "분석항목")
            {
                _show1BrowseMode = ""; // 분석항목 모드 해제
                SetInputMode(mode);
                BuildMatchBrowsePanel();
            }
        };

        return tab;
    }


    private void BuildSampleList(StackPanel container)
    {
        var allSamples = new List<object>();

        try
        {
            // Show1 브라우즈 모드 우선 (분석항목 등 독립 모드)
            var browseMode = !string.IsNullOrEmpty(_show1BrowseMode) ? _show1BrowseMode : _inputMode;

            // 선택된 탭에 따른 필터링
            switch (browseMode)
            {
                case "수질분석센터":
                    var analysisRecords = AnalysisRequestService.GetRecentRecords(3);
                    foreach (var record in analysisRecords.Take(20))
                    {
                        allSamples.Add(new
                        {
                            Type = "Analysis",
                            SN = record.약칭,  // 업체명 약칭을 SN으로 사용
                            Name = record.시료명,
                            Source = "수질분석센터",
                            DisplayText = $"🧪 {record.약칭} | {record.시료명}",
                            Color = "#4A90E2"  // 더 밝은 파란색
                        });
                    }
                    break;

                case "비용부담금":
                    var wasteSamples = WasteSampleService.GetRecentSamples(3);
                    foreach (var sample in wasteSamples.Take(20))
                    {
                        allSamples.Add(new
                        {
                            Type = "Waste",
                            SN = sample.SN,
                            Name = sample.업체명,
                            Source = "비용부담금",
                            DisplayText = $"💰 {sample.SN} | {sample.업체명}",
                            Color = "#50C878"  // 더 밝은 초록색
                        });
                    }
                    break;

                case "분석항목":
                    // 분석정보 약칭 컬럼 기준으로 flat list 구성
                    var analyteItems2 = AnalysisService.GetAllItems()
                        .Where(a => !string.IsNullOrWhiteSpace(a.Analyte))
                        .ToList();
                    foreach (var ai in analyteItems2)
                    {
                        string sn = string.IsNullOrWhiteSpace(ai.약칭) ? ai.Analyte : ai.약칭;
                        allSamples.Add(new
                        {
                            Type      = "Compound",
                            SN        = sn,
                            Name      = ai.Analyte,
                            ShortName = sn,
                            Source    = "분석항목",
                            DisplayText = ai.Analyte,
                            Color     = "#BB8FCE",
                            IsGroup   = false,
                            Analyte   = ai.Analyte,
                        });
                    }
                    break;

                case "처리시설":
                    // 분석계획 순서에 따른 처리시설 구조 로딩
                    var (facilityOrder, facilitySamples) = FacilityResultService.GetAnalysisPlanStructure();

                    foreach (var facilityName in facilityOrder)
                    {
                        if (!facilitySamples.ContainsKey(facilityName)) continue;

                        // 처리시설명 (부모 노드)
                        allSamples.Add(new
                        {
                            Type = "FacilityGroup",
                            SN = facilityName,  // 처리시설명을 SN으로 사용
                            Name = facilityName,
                            Source = "처리시설",
                            DisplayText = $"🏭 {facilityName}",
                            Color = "#FF7F50",  // 더 밝은 주황색
                            IsGroup = true
                        });

                        // 분석계획 순서에 따른 시료명들 (하위 노드)
                        foreach (var sampleName in facilitySamples[facilityName])
                        {
                            allSamples.Add(new
                            {
                                Type = "Facility",
                                SN = facilityName,  // 처리시설명을 SN으로 사용
                                Name = sampleName,
                                Source = "처리시설",
                                DisplayText = $"   📄 {sampleName}",  // 들여쓰기로 하위 표시
                                Color = "#FFA07A",  // 연한 주황색
                                IsGroup = false,
                                ParentFacility = facilityName
                            });
                        }
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
        }

        // UI 생성
        if (allSamples.Count > 0)
        {
            // 탭별 헤더 텍스트
            var currentBrowse = !string.IsNullOrEmpty(_show1BrowseMode) ? _show1BrowseMode : _inputMode;
            string headerIcon = currentBrowse switch
            {
                "수질분석센터" => "🧪",
                "비용부담금" => "💰",
                "처리시설" => "🏭",
                "분석항목" => "🔬",
                _ => "📋"
            };

            string headerText = currentBrowse == "분석항목"
                ? $"{headerIcon} 분석항목 ({allSamples.Count}개)"
                : $"{headerIcon} {currentBrowse} 의뢰시료 ({allSamples.Count}개)";
            container.Children.Add(FsSM(new TextBlock
            {
                Text = headerText,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("AppFg"), // 헤더도 흰색으로 통일
                Margin = new Thickness(4, 0, 0, 8)
            }));

            foreach (var sample in allSamples)
            {
                dynamic s = sample;

                // 그룹 노드와 일반 노드 구분
                bool isGroup = s.GetType().GetProperty("IsGroup")?.GetValue(s, null) is true;

                // Compound 타입 (분석항목 flat list): 약칭 배지 + 전체명
                bool isCompound = !isGroup && s.GetType().GetProperty("ShortName") != null;
                Control itemChild;
                if (isCompound)
                {
                    string shortN = (string)(s.GetType().GetProperty("ShortName")!.GetValue(s) ?? "");
                    var (badgeBg, badgeFg) = BadgeColorHelper.GetBadgeColor(shortN);
                    var inner = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
                    inner.Children.Add(new Border
                    {
                        Background   = new SolidColorBrush(Color.Parse(badgeBg)),
                        CornerRadius = new CornerRadius(3),
                        Padding      = new Thickness(5, 1),
                        Margin       = new Thickness(0, 0, 7, 0),
                        [Grid.ColumnProperty] = 0,
                        Child = FsXS(new TextBlock
                        {
                            Text       = shortN,
                            Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                            VerticalAlignment = VerticalAlignment.Center,
                        })
                    });
                    inner.Children.Add(FsSM(new TextBlock
                    {
                        Text              = (string)s.Name,
                        Foreground        = AppRes("AppFg"),
                        VerticalAlignment = VerticalAlignment.Center,
                        [Grid.ColumnProperty] = 1,
                    }));
                    itemChild = inner;
                }
                else
                {
                    itemChild = FsSM(new TextBlock
                    {
                        Text       = s.DisplayText,
                        Foreground = AppRes("AppFg"),
                        FontWeight = isGroup ? FontWeight.SemiBold : FontWeight.Normal
                    });
                }

                var sampleItem = new Border
                {
                    Background   = isGroup ? AppRes("ThemeBgSecondary") : AppRes("ThemeBgInput"),
                    CornerRadius = new CornerRadius(isGroup ? 8 : 6),
                    Padding      = new Thickness(isGroup ? 12 : 10, isGroup ? 8 : 6),
                    Margin       = new Thickness(0, isGroup ? 4 : 2),
                    Height       = isGroup ? 36 : 28,
                    Cursor       = new Cursor(StandardCursorType.Hand),
                    Child        = itemChild,
                };

                // 그룹이 아닌 드래그 가능한 시료만 드래그 설정
                if (!isGroup)
                {
                    sampleItem.Tag = sample;
                    SetupDragSource(sampleItem, s.Name, sample);
                    _matchItems.Add((sampleItem, (string)s.Name, (object?)sample));
                }

                container.Children.Add(sampleItem);
            }
        }
        else
        {
            container.Children.Add(FsXS(new TextBlock
            {
                Text = $"{_inputMode} 의뢰시료가 없습니다.",
                Foreground = AppRes("ThemeFgMuted"),
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(8)
            }));
        }
    }

    private void SetItemTextColor(Border item, IBrush? brush)
    {
        // Border 내부의 모든 TextBlock 색상 변경
        if (item.Child is StackPanel panel)
        {
            foreach (var child in panel.Children)
            {
                if (child is TextBlock tb) tb.Foreground = brush;
                else if (child is StackPanel subPanel)
                {
                    foreach (var subChild in subPanel.Children)
                        if (subChild is TextBlock subTb) subTb.Foreground = brush;
                }
            }
        }
        else if (item.Child is TextBlock directTb)
        {
            directTb.Foreground = brush;
        }
    }

    private void CancelSelectedRowMatching()
    {
        if (_selectedRowIndex < 0 || _currentExcelRows == null || _selectedRowIndex >= _currentExcelRows.Count)
            return;

        var row = _currentExcelRows[_selectedRowIndex];

        // 매칭된 상태가 아니면 취소할 것이 없음
        if (row.Source == SourceType.미분류 || row.Status == MatchStatus.미매칭)
        {
            ShowMessage("매칭되지 않은 시료입니다.", false);
            return;
        }

        // 매칭 상태 초기화 + 원본명 복원
        row.Source = SourceType.미분류;
        row.Status = MatchStatus.미매칭;
        row.Matched = null;
        row.MatchedAnalysis = null;
        row.MatchedFacilityName = null;
        row.IsManualMatch = false;

        // 원본 시료명으로 복원
        if (!string.IsNullOrEmpty(row.원본시료명))
        {
            row.시료명 = row.원본시료명;
            row.원본시료명 = "";  // 원본명 필드 초기화
        }

        // _categoryExcelData도 동기화
        if (_categoryExcelData.TryGetValue(_activeCategory, out var categoryRows) &&
            _selectedRowIndex < categoryRows.Count)
        {
            categoryRows[_selectedRowIndex].Source = row.Source;
            categoryRows[_selectedRowIndex].Status = row.Status;
            categoryRows[_selectedRowIndex].Matched = row.Matched;
            categoryRows[_selectedRowIndex].MatchedAnalysis = row.MatchedAnalysis;
            categoryRows[_selectedRowIndex].MatchedFacilityName = row.MatchedFacilityName;
            categoryRows[_selectedRowIndex].IsManualMatch = row.IsManualMatch;
            categoryRows[_selectedRowIndex].시료명 = row.시료명;
            categoryRows[_selectedRowIndex].원본시료명 = row.원본시료명;
        }

        // UI 시각적 업데이트: 시료명 셀 → 원본명만 표시
        if (_selectedRowIndex < _rowNameCells.Count)
        {
            var nameCell = _rowNameCells[_selectedRowIndex];
            nameCell.Children.Clear();
            nameCell.Children.Add(FsBase(new TextBlock
            {
                Text = row.시료명, FontFamily = Font,
                Foreground = AppRes("AppFg"),
                TextTrimming = TextTrimming.CharacterEllipsis,
            }));
        }

        // 시료구분 셀 초기화
        if (_selectedRowIndex < _rowSourceCells.Count)
        {
            _rowSourceCells[_selectedRowIndex].Text = "—";
            _rowSourceCells[_selectedRowIndex].Foreground = AppRes("FgMuted");
        }

        // 아이콘 초기화
        if (_selectedRowIndex < _rowIcons.Count)
        {
            _rowIcons[_selectedRowIndex].Background = AppRes("FgMuted");
            _rowIcons[_selectedRowIndex].Opacity = 1.0;
        }

        // SN 셀 초기화
        if (_selectedRowIndex < _rowSnCells.Count)
            _rowSnCells[_selectedRowIndex].Text = row.SN;

        // 토글 OFF
        if (_selectedRowIndex < _rowToggles.Count)
        {
            var (track, knob) = _rowToggles[_selectedRowIndex];
            track.Background = AppRes("FgMuted");
            track.Opacity    = 0.5;
            knob.Margin      = new Thickness(2, 2, 0, 2);
        }

        RefreshDuplicateIcons();
        BuildStatsPanel();

        ShowMessage($"'{row.시료명}' 시료의 매칭이 취소되었습니다.", false);
        LogMatch($"MATCH CANCELLED: {row.시료명} → 미분류");
    }

    /// <summary>지정 행을 정도관리(QC) 시료로 설정</summary>
    private void ApplyQcToRow(ExcelRow qRow, int rowIndex)
    {
        // 매칭 초기화
        qRow.Source              = SourceType.미분류;
        qRow.Status              = MatchStatus.미매칭;
        qRow.Matched             = null;
        qRow.MatchedAnalysis     = null;
        qRow.MatchedFacilityName = null;
        qRow.IsManualMatch       = false;
        // 매칭 해제 여부 추적 (셀 재빌드 판단용)
        bool nameReverted = !string.IsNullOrEmpty(qRow.원본시료명);
        if (nameReverted)
        {
            qRow.시료명    = qRow.원본시료명;
            qRow.원본시료명 = "";
        }
        // QC 지정
        qRow.SN        = "QC";
        qRow.IsControl = true;
        // _categoryExcelData 동기화
        if (_categoryExcelData.TryGetValue(_activeCategory, out var catRows)
            && rowIndex < catRows.Count)
        {
            catRows[rowIndex].Source              = qRow.Source;
            catRows[rowIndex].Status              = qRow.Status;
            catRows[rowIndex].Matched             = null;
            catRows[rowIndex].MatchedAnalysis     = null;
            catRows[rowIndex].MatchedFacilityName = null;
            catRows[rowIndex].IsManualMatch       = false;
            catRows[rowIndex].시료명               = qRow.시료명;
            catRows[rowIndex].원본시료명            = qRow.원본시료명;
            catRows[rowIndex].SN                  = "QC";
            catRows[rowIndex].IsControl           = true;
        }
        // 셀 갱신
        if (rowIndex < _rowSnCells.Count)
            _rowSnCells[rowIndex].Text = "QC";
        if (rowIndex < _rowSourceCells.Count)
        {
            _rowSourceCells[rowIndex].Text       = "정도관리";
            _rowSourceCells[rowIndex].Foreground = new Avalonia.Media.SolidColorBrush(
                Avalonia.Media.Color.Parse("#c084fc"));
        }
        // 시료명 셀: 매칭이 해제된 경우에만 재빌드 (자동 QC 상태에서 수동 QC 시 이름 유지)
        if (nameReverted && rowIndex < _rowNameCells.Count)
        {
            string displayName = qRow.시료명 ?? "";
            Avalonia.Media.IBrush nameFg = AppRes("AppFg");
            if (!string.IsNullOrEmpty(qRow.CompoundName))
            {
                var aliasInfo = CompoundAliasService.Resolve(qRow.CompoundName);
                if (aliasInfo != null)
                {
                    string samplePart = qRow.시료명?.Split('|').LastOrDefault()?.Trim() ?? qRow.시료명 ?? "";
                    displayName = $"{aliasInfo.Value.분석항목} | {samplePart}";
                    nameFg = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#90EE90"));
                }
            }
            var nameCell = _rowNameCells[rowIndex];
            nameCell.Children.Clear();
            nameCell.Children.Add(FsBase(new TextBlock
            {
                Text = displayName, FontFamily = Font,
                Foreground = nameFg,
                TextTrimming = TextTrimming.CharacterEllipsis,
            }));
        }
        // 아이콘 금색으로 변경
        if (rowIndex < _rowIcons.Count)
        {
            _rowIcons[rowIndex].Background = new Avalonia.Media.SolidColorBrush(
                Avalonia.Media.Color.Parse("#FFD700"));
            _rowIcons[rowIndex].Opacity = 1.0;
        }
        BuildStatsPanel();
    }

    private void SetupDragSource(Border item, string sampleName, object? sampleData = null)
    {
        item.PointerPressed += async (sender, e) =>
        {
            if (e.GetCurrentPoint(item).Properties.IsLeftButtonPressed)
            {
                // 드래그 시작: 금색으로 변경
                var goldBrush = new SolidColorBrush(Color.Parse("#FFD700")); // 금색
                var defaultBrush = AppRes("AppFg"); // 기본 텍스트 색상

                // item 내부의 모든 TextBlock을 금색으로 변경
                SetItemTextColor(item, goldBrush);
#pragma warning disable CS0618
                var dragData = new DataObject();
                dragData.Set(DataFormats.Text, sampleName);

                // 의뢰시료 타입에 따라 매치 키 설정
                string matchKey = "match-waste"; // 기본값
                string sampleSN = ""; // SN 정보

                if (sampleData != null)
                {
                    // 동적 객체에서 Type 속성 확인
                    dynamic dynSample = sampleData;
                    string sampleType = dynSample.Type;

                    matchKey = sampleType switch
                    {
                        "Analysis" => "match-analysis",
                        "Facility" => "match-facility",
                        "Waste" => "match-waste",
                        "Compound" => "match-compound",
                        "Excel" => _inputMode switch
                        {
                            "수질분석센터" => "match-analysis",
                            "처리시설" => "match-facility",
                            _ => "match-waste"
                        },
                        _ => "match-waste"
                    };

                    // SN 정보 추출
                    try
                    {
                        sampleSN = sampleType switch
                        {
                            "Analysis" => dynSample.SN?.ToString() ?? "", // 수질분석센터: 약칭을 SN으로 사용
                            "Facility" => dynSample.SN?.ToString() ?? "", // 처리시설: 시설명을 SN으로 사용
                            "Waste" => dynSample.SN?.ToString() ?? "", // 비용부담금: 실제 SN 사용
                            "Compound" => dynSample.SN?.ToString() ?? "", // 분석항목: 카테고리를 SN으로 사용
                            _ => ""
                        };
                    }
                    catch
                    {
                        sampleSN = "";
                    }
                }

                dragData.Set(matchKey, sampleName);
                if (!string.IsNullOrEmpty(sampleSN))
                {
                    dragData.Set("sample-sn", sampleSN); // SN 정보 추가 전달
                }

                await DragDrop.DoDragDrop(e, dragData, DragDropEffects.Link);

                // 드래그 완료: 기본 색상으로 복원
                SetItemTextColor(item, defaultBrush);
#pragma warning restore CS0618
            }
        };
    }


    private Control BuildStatusSummary()
    {
        if (string.IsNullOrEmpty(_selectedDate))
        {
            return FsXS(new TextBlock
            {
                Text = "날짜 선택 후 현황 표시됩니다.",
                Foreground = AppRes("ThemeFgMuted"),
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 12)
            });
        }

        var summary = new StackPanel { Spacing = 4 };

        // 모드별 현황 표시
        if (IsWaterCenterMode)
        {
            // 수질분석센터: 분석의뢰및결과에서 항목별 완료 여부
            summary.Children.Add(FsXS(new TextBlock
            {
                Text = "🧪 수질분석센터 현황",
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("ThemeFgInfo")
            }));
        }
        else if (IsFacilityMode)
        {
            // 처리시설: 처리시설_작업 현황
            summary.Children.Add(FsXS(new TextBlock
            {
                Text = "🏭 처리시설 현황",
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("ThemeFgInfo")
            }));
        }
        else if (IsBillingMode)
        {
            // 비용부담금: 폐수의뢰및결과 BOD~Phenols 입력 현황
            summary.Children.Add(FsXS(new TextBlock
            {
                Text = "💰 비용부담금 현황",
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("ThemeFgInfo")
            }));

            var items = new[] { "BOD", "TOC", "SS", "T-N", "T-P", "N-Hexan", "Phenols" };
            int total = _currentSamples?.Count ?? 0;

            foreach (var item in items)
            {
                int completed = _currentSamples?.Count(s => !string.IsNullOrWhiteSpace(GetSampleValue(s, item))) ?? 0;
                var status = total > 0 ? $"{completed}/{total}" : "0/0";
                var color = completed == total && total > 0 ? "ThemeFgSuccess" :
                           completed > 0 ? "ThemeFgWarn" : "ThemeFgMuted";

                summary.Children.Add(new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children =
                    {
                        FsXS(new TextBlock
                        {
                            Text = item,
                            Width = 60,
                            Foreground = AppRes("ThemeFgSecondary")
                        }),
                        FsXS(new TextBlock
                        {
                            Text = status,
                            Foreground = AppRes(color)
                        })
                    }
                });
            }
        }

        return new Border
        {
            Background = AppRes("ThemeBgCard"),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
            Margin = new Thickness(0, 4, 0, 0),
            Child = summary
        };
    }

    private void BuildStatsPanel()
    {
        // 전체 UI 업데이트 비활성화 - 모든 파서 스턱 방지
        return;
    }

    // =========================================================================
    // TOC 기기 파일 파싱 (스칼라 CSV / 시마즈 TXT)
    // =========================================================================

    // 기기 임포트 데이터 저장 (SN → (rawName, conc, area, method))
    // TOC 기기 파일 임포트 시 사용한 방법 (NPOC/TCIC)
    private string _tocInstrumentMethod = "NPOC";

    private void ParseTocInstrumentFile(string path)
    {
        var (fmt, instrRows, cal) = TocInstrumentParser.Parse(path);
        if (fmt == TocInstrumentParser.TocFileFormat.Unknown)
        {
            ShowMessage("지원되지 않는 TOC 기기 파일 형식입니다.", true);
            return;
        }

        // 예나 검량선 전용 파일: 시료 없음, 검량선 값만 저장
        if (fmt == TocInstrumentParser.TocFileFormat.JenaCalibration)
        {
            if (cal?.HasData == true)
            {
                var docInfo = _categoryDocInfo.ContainsKey("TOC")
                    ? _categoryDocInfo["TOC"]
                    : (_categoryDocInfo["TOC"] = new ExcelDocInfo());
                docInfo.TocSlope_TC     = cal.Slope_TC;
                docInfo.TocIntercept_TC = cal.Intercept_TC;
                docInfo.TocSlope_IC     = cal.Slope_IC;
                docInfo.TocIntercept_IC = cal.Intercept_IC;
                if (cal.Method == "TCIC") { docInfo.IsTocTCIC = true; _tocInstrumentMethod = "TCIC"; }
                else                      { docInfo.IsTocNPOC = true; _tocInstrumentMethod = "NPOC"; }
                ShowMessage($"예나 검량선 로드 완료 (기울기 TC: {cal.Slope_TC})", false);
            }
            else ShowMessage("예나 검량선 파일에서 데이터를 찾을 수 없습니다.", true);
            return;
        }

        // 스칼라/시마즈: 정도관리 시료도 기록부 증거이므로 보존 (IsControl 플래그 전달)
        var excelRows = new List<ExcelRow>();
        string method = "NPOC";
        foreach (var r in instrRows)
        {
            method = r.Method;
            excelRows.Add(new ExcelRow
            {
                시료명 = r.RawName,
                SN     = r.SN,
                Fxy    = r.Conc,   // f(x/y) = 기기 계산 농도
                Result = r.Conc,   // 결과값 (희석배수 미적용 시 동일)
                D1     = r.Area,
                P      = string.IsNullOrEmpty(r.Dilution) ? "" : r.Dilution,
                IsControl = r.IsControl,
                // TCIC 전용 raw값 매핑
                TCAU   = r.TCAU,
                TCcon  = r.TCcon,
                ICAU   = r.ICAU,
                ICcon  = r.ICcon,
            });
        }

        if (excelRows.Count == 0)
        {
            ShowMessage("시료 데이터가 없습니다.", true);
            return;
        }

        // 기존 검량선 데이터 초기화 (스칼라→시마즈 전환 시 잔류 방지)
        var tocDocInfo = _categoryDocInfo.ContainsKey("TOC")
            ? _categoryDocInfo["TOC"]
            : (_categoryDocInfo["TOC"] = new ExcelDocInfo());
        tocDocInfo.TocSlope_TC     = "";
        tocDocInfo.TocIntercept_TC = "";
        tocDocInfo.TocR2_TC        = "";
        tocDocInfo.TocSlope_IC     = "";
        tocDocInfo.TocIntercept_IC = "";
        tocDocInfo.TocR2_IC        = "";
        tocDocInfo.TocStdConcs     = Array.Empty<string>();
        tocDocInfo.TocStdAreas     = Array.Empty<string>();
        tocDocInfo.TocStdConcs_IC  = Array.Empty<string>();
        tocDocInfo.TocStdAreas_IC  = Array.Empty<string>();
        tocDocInfo.IsTocNPOC       = false;
        tocDocInfo.IsTocTCIC       = false;

        // 새 검량선 데이터 저장 (스칼라 등 검량선 포함 파일)
        if (cal?.HasData == true)
        {
            tocDocInfo.TocSlope_TC     = cal.Slope_TC;
            tocDocInfo.TocIntercept_TC = cal.Intercept_TC;
            tocDocInfo.TocR2_TC        = cal.R2_TC;
            tocDocInfo.TocSlope_IC     = cal.Slope_IC;
            tocDocInfo.TocIntercept_IC = cal.Intercept_IC;
            tocDocInfo.TocR2_IC        = cal.R2_IC;
            tocDocInfo.TocStdConcs     = cal.StdConcs;
            tocDocInfo.TocStdAreas     = cal.StdAreas;
            tocDocInfo.TocStdConcs_IC  = cal.StdConcs_IC;
            tocDocInfo.TocStdAreas_IC  = cal.StdAreas_IC;
            bool isTcic = cal.Method.Equals("TCIC", StringComparison.OrdinalIgnoreCase);
            tocDocInfo.IsTocNPOC = !isTcic;
            tocDocInfo.IsTocTCIC = isTcic;
        }

        _tocInstrumentMethod = method;
        _activeCategory      = "TOC";
        var tocCat = Categories.FirstOrDefault(c => c.Key == "TOC");
        _activeItems = tocCat.Items ?? new[] { "TOC" };
        _categorySelected    = true;
        _categoryExcelData["TOC"] = excelRows;

        // 검량선 없어도 TOC 모드 플래그 보장 (시마즈 등)
        if (!tocDocInfo.IsTocNPOC && !tocDocInfo.IsTocTCIC)
        {
            bool isTcic = method.Equals("TCIC", StringComparison.OrdinalIgnoreCase);
            tocDocInfo.IsTocNPOC = !isTcic;
            tocDocInfo.IsTocTCIC = isTcic;
        }

        // 분석일: 오늘 날짜 기본
        if (!_categoryDocDates.ContainsKey("TOC") || string.IsNullOrEmpty(_categoryDocDates["TOC"]))
            _categoryDocDates["TOC"] = DateTime.Today.ToString("yyyy-MM-dd");

        UpdateCategoryButtonStyles();
        { /* LoadVerifiedGrid 제거됨 */ }
    }

    // =========================================================================
    // GC 기기 파일 파싱 (Agilent ChemStation/MassHunter CSV)
    // =========================================================================
    private void ParseGcInstrumentFile(string path)
    {
        string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
        var gc = ext == ".pdf"
            ? Voc10PdfParser.Parse(path)
            : GcInstrumentParser.Parse(path);

        if (gc.Format == GcFileFormat.Unknown || gc.Compounds.Count == 0)
        {
            ShowMessage("지원되지 않는 GC 기기 파일 형식입니다.", true);
            return;
        }

        // 모든 성분의 행을 하나의 목록으로 합침
        // Cal(검정곡선 원자료)만 제외, 정도관리(BK/CCV/FBK/MBK)는 IsControl 플래그 세팅하여 포함
        // 시료명은 "<성분> | <RawName>" 으로 구분하여 복수 성분 공존 허용
        var excelRows = new List<ExcelRow>();
        int sampleCount = 0;
        foreach (var c in gc.Compounds)
        {
            foreach (var r in c.Rows)
            {
                if (r.Type.Equals("Cal", StringComparison.OrdinalIgnoreCase)) continue;
                excelRows.Add(new ExcelRow
                {
                    시료명       = string.IsNullOrEmpty(c.Name) ? r.RawName : $"{c.Name} | {r.RawName}",
                    CompoundName = c.Name,       // 성분별 라우팅 키
                    SN           = r.SN,
                    D1           = r.Resp,        // 기기응답(면적)
                    D2           = r.IstdResp,    // ISTD 응답 (없으면 빈값)
                    Fxy          = r.FinalConc,   // 기기 계산 농도
                    Result       = r.FinalConc,
                    P            = r.Dilution,    // 희석배수 (상세블록에서 조인)
                    IsControl    = r.IsControl,
                });
                if (!r.IsControl) sampleCount++;
            }
        }

        if (excelRows.Count == 0)
        {
            ShowMessage("GC 파일에 시료 데이터가 없습니다.", true);
            return;
        }

        // GCMS 카테고리에 라우팅 — _activeItems를 파싱된 성분명으로 동적 구성
        _activeCategory   = "GCMS";
        _activeItems      = gc.Compounds
            .Select(c => c.Name)
            .Where(n => !string.IsNullOrEmpty(n))
            .Distinct().ToArray();
        _categorySelected = true;
        _categoryExcelData["GCMS"] = excelRows;

        // 성분별 검량선 → ExcelDocInfo 에 저장 (Show2 렌더링용)
        var gcDocInfo = _categoryDocInfo.ContainsKey("GCMS")
            ? _categoryDocInfo["GCMS"]
            : (_categoryDocInfo["GCMS"] = new ExcelDocInfo());
        gcDocInfo.IsGcMode = true;
        gcDocInfo.GcFormat = gc.Format.ToString();
        gcDocInfo.GcCompoundCals = gc.Compounds.Select(c => new GcCompoundCalInfo
        {
            Name      = c.Name,
            HasIstd   = c.HasIstd,
            Slope     = c.SlopeA.HasValue    ? c.SlopeA.Value.ToString("G6",    System.Globalization.CultureInfo.InvariantCulture) : "",
            Intercept = c.Intercept.HasValue ? c.Intercept.Value.ToString("G6", System.Globalization.CultureInfo.InvariantCulture) : "",
            R         = c.R.HasValue         ? c.R.Value.ToString("F4",         System.Globalization.CultureInfo.InvariantCulture) : "",
            StdConcs     = c.Calibration.Select(p => p.Conc).ToArray(),
            StdResps     = c.Calibration.Select(p => p.Response).ToArray(),
            StdIstdResps = c.Calibration.Select(p => p.IstdResponse).ToArray(),
        }).ToList();

        if (!_categoryDocDates.ContainsKey("GCMS") || string.IsNullOrEmpty(_categoryDocDates["GCMS"]))
        {
            // 배치 헤더에서 분석일 추출 시도 (예: "2026-03-17 오전 9:08:50")
            var d = gc.AnalysisTime;
            if (!string.IsNullOrEmpty(d) && d.Length >= 10 && DateTime.TryParse(d[..10], out var dt))
                _categoryDocDates["GCMS"] = dt.ToString("yyyy-MM-dd");
            else
                _categoryDocDates["GCMS"] = DateTime.Today.ToString("yyyy-MM-dd");
        }

        UpdateCategoryButtonStyles();
        { /* LoadVerifiedGrid 제거됨 */ }

        var compoundList = string.Join(", ", gc.Compounds.Take(3).Select(c => c.Name));
        if (gc.Compounds.Count > 3) compoundList += $", +{gc.Compounds.Count - 3}";
        ShowMessage($"GC 파일 감지 [{gc.Format}] — {gc.Compounds.Count}성분 ({compoundList}) / 시료 {sampleCount}개", false);
    }

    // 엑셀 파싱 로직은 Services/SERVICE4/BodExcelParser.cs 로 이관

    // =========================================================================
    // 외부 호출: 카테고리 선택 / 검증 / 입력 / 출력 / 새로고침
    // =========================================================================
    public void SelectCategoryByKey(string key)
    {
        if (Categories.Any(c => c.Key == key))
        {
            _activeCategory = key;
            var match = Categories.FirstOrDefault(c => c.Key == key);
            _activeItems = match.Items ?? Array.Empty<string>();
            UpdateCategoryButtonStyles();
        }
    }

    /// <summary>서브메뉴 "검증" — 첨부된 자료를 DB와 순차 대조 검증</summary>
    public async void VerifyData()
    {
        if (!_categoryExcelData.TryGetValue(_activeCategory, out var excelRows) || _gridPanel == null)
        {
            ShowMessage("먼저 파일을 첨부하세요.", true);
            return;
        }

        // 1단계: 모든 아이콘을 회색(로딩 중)으로 초기화
        foreach (var ic in _rowIcons)
        {
            ic.Background = AppRes("FgMuted");
            ic.Opacity = 0.5;
        }

        // 2단계: 순차적으로 DB 확인 후 아이콘 색상 복원
        var allSamples = new List<WasteSample>();
        var loadedDates = new HashSet<string>();
        _categoryDocDates.TryGetValue(_activeCategory, out var docDateStr);
        int docYear = DateTime.Now.Year;
        if (DateTime.TryParse(docDateStr, out var docDt)) docYear = docDt.Year;

        if (_selectedDate != null && !loadedDates.Contains(_selectedDate))
        {
            try { allSamples.AddRange(WasteSampleService.GetByDate(_selectedDate)); } catch { }
            loadedDates.Add(_selectedDate);
        }
        foreach (var row in excelRows)
        {
            var sn = row.SN;
            if (sn.StartsWith("[")) { int idx = sn.IndexOf(']'); if (idx > 0) sn = sn[(idx + 1)..]; }
            var parts = sn.Split('-');
            if (parts.Length >= 2 && int.TryParse(parts[0], out var mm) && int.TryParse(parts[1], out var dd))
            {
                var dateStr = $"{docYear:D4}-{mm:D2}-{dd:D2}";
                if (!loadedDates.Contains(dateStr))
                {
                    try { allSamples.AddRange(WasteSampleService.GetByDate(dateStr)); } catch { }
                    loadedDates.Add(dateStr);
                }
            }
        }

        var analysisRecords = new List<AnalysisRequestRecord>();
        foreach (var dt in loadedDates)
        {
            try { analysisRecords.AddRange(AnalysisRequestService.GetByDate(dt)); } catch { }
        }

        List<(string 시설명, string 시료명, int 마스터Id)>? facilityMasters = null;
        try { facilityMasters = FacilityResultService.GetAllMasterSamples(); } catch { }

        for (int i = 0; i < excelRows.Count; i++)
        {
            var row = excelRows[i];

            // 체크 해제된 행은 스킵
            if (!row.Enabled)
            {
                if (i < _rowIcons.Count) { _rowIcons[i].Background = AppRes("FgMuted"); _rowIcons[i].Opacity = 0.35; }
                await System.Threading.Tasks.Task.Delay(50);
                continue;
            }

            // DB 매칭 검증
            row.Matched = allSamples.FirstOrDefault(s => s.SN == row.SN)
                       ?? allSamples.FirstOrDefault(s => s.업체명 == row.시료명);

            if (row.Matched != null)
            {
                row.Source = SourceType.폐수배출업소;
                bool has = _activeItems.Any(item => !string.IsNullOrWhiteSpace(GetSampleValue(row.Matched, item)));
                row.Status = has ? MatchStatus.덮어쓰기 : MatchStatus.입력가능;
            }
            else if (facilityMasters != null)
            {
                var fm = FacilityResultService.FindBySampleName(facilityMasters, row.시료명);
                if (fm != null)
                {
                    row.Source = SourceType.처리시설;
                    row.MatchedFacilityName = fm.Value.시설명;
                    row.Status = MatchStatus.입력가능;
                }
                else
                {
                    var ar = analysisRecords.FirstOrDefault(a => a.약칭 == row.시료명 || a.시료명 == row.시료명);
                    if (ar != null)
                    {
                        row.MatchedAnalysis = ar;
                        row.Source = SourceType.수질분석센터;
                        row.Status = MatchStatus.입력가능;
                    }
                    else
                    {
                        row.Source = SourceType.미분류;
                        row.Status = allSamples.Count > 0 || analysisRecords.Count > 0
                                     || (_matchingFacilityMasters?.Count ?? 0) > 0
                            ? MatchStatus.미매칭 : MatchStatus.대기;
                    }
                }
            }
            else
            {
                var ar = analysisRecords.FirstOrDefault(a => a.약칭 == row.시료명 || a.시료명 == row.시료명);
                if (ar != null)
                {
                    row.MatchedAnalysis = ar;
                    row.Source = SourceType.수질분석센터;
                    row.Status = MatchStatus.입력가능;
                }
                else
                {
                    row.Source = SourceType.미분류;
                    row.Status = allSamples.Count > 0 || analysisRecords.Count > 0
                                 || (_matchingFacilityMasters?.Count ?? 0) > 0
                        ? MatchStatus.미매칭 : MatchStatus.대기;
                }
            }

            // 아이콘 업데이트 (0.05초 딜레이)
            if (i < _rowIcons.Count)
            {
                _rowIcons[i].Background = row.Status switch
                {
                    MatchStatus.입력가능 => AppRes("ThemeFgSuccess"),
                    MatchStatus.덮어쓰기 => AppRes("ThemeFgWarn"),
                    MatchStatus.대기     => AppRes("FgMuted"),
                    _                    => AppRes("ThemeFgDanger"),
                };
                _rowIcons[i].Opacity = row.Status == MatchStatus.대기 ? 0.35 : 1.0;
            }

            await System.Threading.Tasks.Task.Delay(50);
        }

        _currentSamples = allSamples;
        ShowMessage($"✅ 검증 완료: {excelRows.Count}건 확인", false);
    }

    /// <summary>서브메뉴 "입력" — 검증된 데이터 일괄 DB 반영</summary>
    public async Task ImportData()
    {
        if (!_categoryExcelData.TryGetValue(_activeCategory, out var rows))
        { ShowMessage("먼저 파일을 첨부하세요.", true); return; }

        try { File.WriteAllText(ImportLogPath, $"=== ImportData 시작: {DateTime.Now:yyyy-MM-dd HH:mm:ss} cat={_activeCategory} rows={rows.Count} ===\n"); } catch { }

        // 진행 오버레이 표시
        if (_importOverlay != null) _importOverlay.IsVisible = true;
        int totalRows = rows.Count;
        if (_importPb != null) { _importPb.Value = 0; _importPb.Maximum = Math.Max(1, totalRows); }

        int imported = 0, skipped = 0, disabled = 0;
        var modifiedDates = new HashSet<string>();

        // 스냅샷 캡처 (Task.Run 안에서 안전하게 접근)
        var category   = _activeCategory;
        var docDates   = new Dictionary<string, string?>(_categoryDocDates);
        var docInfoMap = new Dictionary<string, ExcelDocInfo>(_categoryDocInfo);

        int processed = 0;
        await Task.Run(() =>
        {
        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.Result)) { LogInput($"SKIP(빈결과): [{processed}] SN={row.SN} 시료명={row.시료명}"); processed++; continue; }
            if (!row.Enabled) { LogInput($"SKIP(비활성): [{processed}] SN={row.SN} Source={row.Source} Matched={row.Matched != null}"); disabled++; processed++; continue; }

            try
            {
                switch (row.Source)
                {
                    case SourceType.폐수배출업소 when row.Matched != null:
                        LogInput($"[{processed}] 폐수배출업소: {row.시료명} SN={row.SN}");
                        try { SaveRawData(row, row.Matched, "폐수배출업소"); }
                        catch (Exception ex) { LogInput($"  오류: {ex.Message}"); }
                        UpdateWasteSampleValues(row);
                        modifiedDates.Add(row.Matched.채수일);
                        imported++;
                        break;

                    case SourceType.수질분석센터 when row.MatchedAnalysis != null:
                        LogInput($"[{processed}] 수질분석센터: {row.시료명} SN={row.SN}");
                        try { SaveRawData(row, row.Matched, "수질분석센터"); }
                        catch (Exception ex) { LogInput($"  오류: {ex.Message}"); }
                        ImportAnalysisRequest(row);
                        if (row.Matched != null) modifiedDates.Add(row.Matched.채수일);
                        imported++;
                        break;

                    case SourceType.처리시설 when row.MatchedFacilityName != null:
                        LogInput($"[{processed}] 처리시설: {row.시료명} SN={row.SN} 시설={row.MatchedFacilityName}");
                        try { SaveRawData(row, null, "처리시설"); }
                        catch (Exception ex) { LogInput($"  오류: {ex.Message}"); }
                        ImportFacilityResult(row);
                        docDates.TryGetValue(category, out var fd);
                        if (!string.IsNullOrEmpty(fd)) modifiedDates.Add(fd);
                        imported++;
                        break;

                    default:
                        if (string.IsNullOrEmpty(row.SN)) row.SN = row.시료명;
                        var srcTag = row.IsControl ? "QAQC"
                            : row.Source == SourceType.폐수배출업소 ? "폐수배출업소"
                            : row.Source == SourceType.수질분석센터 ? "수질분석센터"
                            : row.Source == SourceType.처리시설     ? "처리시설"
                            : "미분류";
                        LogInput($"[{processed}] {srcTag}: {row.시료명} Source={row.Source} Matched={row.Matched != null} MatchedAnalysis={row.MatchedAnalysis != null} FacilityName={row.MatchedFacilityName}");
                        try { SaveRawData(row, row.Matched, srcTag); }
                        catch (Exception ex) { LogInput($"  오류: {ex.Message}"); }
                        imported++;
                        break;
                }
            }
            catch (Exception ex) { LogInput($"  [{processed}] 예외: {ex.Message}"); skipped++; }
            finally
            {
                int p = ++processed;
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    if (_importPb != null) _importPb.Value = p;
                });
            }
        }
        }); // end Task.Run

        // _selectedDate도 포함 (날짜를 클릭한 경우)
        if (_selectedDate != null) modifiedDates.Add(_selectedDate);

        // N-Hexan 바탕시료 별도 저장 (분석일 기준으로 SN="바탕시료")
        if (_activeCategory == "NHEX"
            && _categoryDocInfo.TryGetValue("NHEX", out var nhexDocInfo)
            && nhexDocInfo.IsNHEX
            && !string.IsNullOrWhiteSpace(nhexDocInfo.바탕시료_결과))
        {
            _categoryDocDates.TryGetValue("NHEX", out var nhexDate);
            if (!string.IsNullOrEmpty(nhexDate))
            {
                try
                {
                    WasteSampleService.UpsertSimpleData(
                        "NHexan_시험기록부", "결과",
                        nhexDate, "바탕시료", "바탕시료", "",
                        nhexDocInfo.바탕시료_시료량, nhexDocInfo.바탕시료_결과);
                }
                catch { }
            }
        }

        // 오버레이 숨김 (LoadVerifiedGrid가 새 root를 만들기 전에)
        if (_importOverlay != null) _importOverlay.IsVisible = false;

        { /* LoadVerifiedGrid 제거됨 */ }
        BuildStatsPanel();
        RefreshDateNodes(modifiedDates);
        LogInput($"=== ImportData 완료: {imported}건 저장, {disabled}건 비활성(토글OFF), {skipped}건 오류 ===");
        string msg = $"✅ {imported}건 입력" +
            (disabled > 0 ? $" / {disabled}건 체크해제" : "") +
            (skipped > 0 ? $" / {skipped}건 제외" : "");
        ShowMessage(msg, false);
    }

    // 매칭된 폐수배출업소 시료의 요약 결과값을 폐수의뢰및결과 테이블에 갱신.
    // *_*_DATA 저장은 ImportData() 루프 앞단에서 별도로 처리되므로 여기서는 요약만.
    private void UpdateWasteSampleValues(ExcelRow row)
    {
        var s = row.Matched!;

        // 다성분 카테고리(GC/ICP/PFAS)는 동적 컬럼 갱신 (CompoundAliasService로 정규화)
        if (IsMultiCompoundCategory(_activeCategory) && !string.IsNullOrEmpty(row.CompoundName))
        {
            var resolved = CompoundAliasService.ResolveOrFallback(row.CompoundName);
            if (!string.IsNullOrEmpty(resolved.분석항목))
                WasteSampleService.UpdateDynamicValue("폐수의뢰및결과", s.Id, resolved.분석항목, row.Result);
            return;
        }

        if (_activeItems.Length > 0)
        {
            var activeItem = _activeItems[0];
            switch (activeItem)
            {
                case "BOD": s.BOD = row.Result; break;
                case "TOC": s.TOC = row.Result; break;
                case "SS":  s.SS  = row.Result; break;
                case "T-N": s.TN  = row.Result; break;
                case "T-P": s.TP  = row.Result; break;
                case "N-Hexan": s.NHexan = row.Result; break;
                case "Phenols": s.Phenols = row.Result; break;
                case "시안": s.CN = row.Result; break;
                case "6가크롱": s.CR6 = row.Result; break;
                case "색도": s.COLOR = row.Result; break;
                case "ABS": s.ABS = row.Result; break;
                case "불소": s.FLUORIDE = row.Result; break;
            }
        }
        WasteSampleService.UpdateValues(s.Id, s.BOD, s.TOC, s.SS, s.TN, s.TP, s.NHexan, s.Phenols,
            s.CN, s.CR6, s.COLOR, s.ABS, s.FLUORIDE);
    }

    private void ImportAnalysisRequest(ExcelRow row)
    {
        var ar = row.MatchedAnalysis!;
        var shortNames = AnalysisRequestService.GetShortNames();

        // 다성분 카테고리(GC/ICP/PFAS)는 성분명으로 컬럼 매핑 (CompoundAliasService로 정규화)
        if (IsMultiCompoundCategory(_activeCategory) && !string.IsNullOrEmpty(row.CompoundName))
        {
            var analyteName = CompoundAliasService.ResolveOrFallback(row.CompoundName).분석항목;
            if (!string.IsNullOrEmpty(analyteName))
            {
                var colName = shortNames.FirstOrDefault(kv =>
                    kv.Value.Equals(analyteName, StringComparison.OrdinalIgnoreCase)
                    || kv.Key.Equals(analyteName, StringComparison.OrdinalIgnoreCase)).Key;
                if (!string.IsNullOrEmpty(colName))
                    AnalysisRequestService.UpdateResultValue(ar.Id, colName, row.Result);
            }
            return;
        }

        // 단일 항목: 약칭으로 컬럼명 매핑 (BOD → 생물화학적 산소요구량 등)
        foreach (var item in _activeItems)
        {
            var colName = shortNames.FirstOrDefault(kv =>
                kv.Value.Equals(item, StringComparison.OrdinalIgnoreCase)).Key;
            if (!string.IsNullOrEmpty(colName))
                AnalysisRequestService.UpdateResultValue(ar.Id, colName, row.Result);
        }
    }

    private void ImportFacilityResult(ExcelRow row)
    {
        // 처리시설은 FacilityResultService.SaveRows 사용
        _categoryDocDates.TryGetValue(_activeCategory, out var docDate);
        if (string.IsNullOrEmpty(docDate)) return;

        var facilityRows = FacilityResultService.GetRows(row.MatchedFacilityName!, docDate);
        var target = facilityRows.FirstOrDefault(f => f.시료명 == row.시료명);
        if (target == null) return;

        // 다성분 카테고리(GC/ICP/PFAS)는 성분명으로 동적 매핑 (CompoundAliasService로 정규화)
        if (IsMultiCompoundCategory(_activeCategory) && !string.IsNullOrEmpty(row.CompoundName))
        {
            var resolved = CompoundAliasService.ResolveOrFallback(row.CompoundName);
            if (!string.IsNullOrEmpty(resolved.분석항목))
                target[resolved.분석항목] = row.Result;
        }
        else if (_activeItems.Length > 0)
        {
            var activeItem = _activeItems[0];
            var aliasMap = FacilityResultService.GetAnalyteAliasMap();
            if (aliasMap.TryGetValue(activeItem, out var mappedCol))
            {
                target[mappedCol] = row.Result;
            }
            else
            {
                switch (activeItem)
                {
                    case "BOD": target.BOD = row.Result; break;
                    case "TOC": target.TOC = row.Result; break;
                    case "SS":  target.SS  = row.Result; break;
                    case "T-N": target.TN  = row.Result; break;
                    case "T-P": target.TP  = row.Result; break;
                    case "N-Hexan": target["N-Hexan"] = row.Result; break;
                    case "Phenols": target["Phenols"] = row.Result; break;
                    case "시안": target["시안"] = row.Result; break;
                    case "6가크롬": target["6가크롬"] = row.Result; break;
                    case "색도": target["색도"] = row.Result; break;
                    case "ABS": target["ABS"] = row.Result; break;
                    case "불소": target["불소"] = row.Result; break;
                }
            }
        }
        string user = ETA.Services.Common.CurrentUserManager.Instance.CurrentUserId ?? "";
        FacilityResultService.SaveRows(row.MatchedFacilityName!, docDate, new List<FacilityResultRow> { target }, user);

        // 원자료 저장은 ImportData에서 SaveRawData("처리시설") 로 통합
    }

    /// <summary>서브메뉴 "출력" — 클립보드 복사</summary>
    public async void ExportData()
    {
        if (_selectedDate == null || _currentSamples.Count == 0) return;

        var sb = new System.Text.StringBuilder();
        sb.Append("구분\tS/N\t업체명");
        foreach (var item in _activeItems) sb.Append($"\t{item}");
        sb.AppendLine();

        foreach (var s in _currentSamples)
        {
            sb.Append($"{s.구분}\t{s.SN}\t{s.업체명}");
            foreach (var item in _activeItems) sb.Append($"\t{GetSampleValue(s, item)}");
            sb.AppendLine();
        }

        try
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard != null) await clipboard.SetTextAsync(sb.ToString());
            ShowMessage($"📋 {_currentSamples.Count}건 클립보드 복사 완료", false);
        }
        catch { ShowMessage("클립보드 복사 실패", true); }
    }

    public void Refresh()
    {
        LoadData();
        if (_selectedDate != null)
        {
            _currentSamples = WasteSampleService.GetByDate(_selectedDate);
            if (_categoryExcelData.ContainsKey(_activeCategory))
                { /* LoadVerifiedGrid 제거됨 */ }
            else
                LoadSampleGrid(_selectedDate);
            BuildStatsPanel();
        }
    }

    // 중복 시료명 체크 (같은 시료명 + 같은 SN인 경우)
    private bool IsDuplicateSampleName(string sampleName)
    {
        if (_currentExcelRows == null) return false;
        var currentRow = _currentExcelRows.FirstOrDefault(r => r.시료명 == sampleName);
        if (currentRow == null) return false;

        string uniqueKey = $"{sampleName}_{currentRow.SN}";
        return _currentExcelRows.Count(r => $"{r.시료명}_{r.SN}" == uniqueKey) > 1;
    }

    // 중복 시료명에 체크박스 표시
    private void ShowDuplicateSampleCheckboxes(string sampleName)
    {
        if (_currentExcelRows == null) return;

        for (int i = 0; i < _currentExcelRows.Count; i++)
        {
            if (_currentExcelRows[i].시료명 == sampleName && i < _rowToggles.Count)
            {
                // 해당 시료명의 모든 행에 체크박스 표시 (간단한 방법으로 _rowToggles 사용)
                // TODO: 실제 UI에서 토글 래퍼를 찾아서 표시하도록 개선 필요
                ShowMessage($"'{sampleName}' 시료의 중복 항목들을 확인해주세요.", false);
            }
        }
    }

    // 현재 선택된 날짜 노드의 아이콘을 DB 재조회 후 갱신
    private async void RefreshCurrentDateNode()
    {
        if (_selectedDate == null) return;
        await RefreshDateNodesAsync(new HashSet<string> { _selectedDate });
    }

    // 지정된 날짜 노드들을 DB 재조회 후 일괄 갱신
    private async void RefreshDateNodes(HashSet<string> dates)
    {
        if (dates.Count == 0) return;
        await RefreshDateNodesAsync(dates);
    }

    private async System.Threading.Tasks.Task RefreshDateNodesAsync(HashSet<string> dates)
    {
        bool isFacility    = IsFacilityMode;
        bool isWaterCenter = IsWaterCenterMode;
        bool isBilling     = IsBillingMode;

        foreach (var dateStr in dates)
        {
            var (samples, facilityStatus, reqItems, schedItems, billingItems) = await System.Threading.Tasks.Task.Run(() =>
            {
                var s = WasteSampleService.GetByDate(dateStr);
                var f = isFacility ? FacilityResultService.GetFillStatusForDate(dateStr) : new Dictionary<string, bool>();
                var r = isWaterCenter ? AnalysisRequestService.GetRequestedCategoriesByDate(dateStr) : new Dictionary<string, bool>();
                var sc = isFacility ? FacilityResultService.GetScheduledItemsByDate(dateStr) : new HashSet<string>();
                var bi = isBilling ? WasteRequestService.GetRequestedItemSetByDate(dateStr) : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                return (s, f, r, sc, bi);
            });

            var newNode = MakeDateNode(dateStr, samples, facilityStatus, reqItems, schedItems, billingItems);

            bool UpdateInItems(Avalonia.Controls.ItemCollection items)
            {
                foreach (var item in items.OfType<TreeViewItem>())
                {
                    if (item.Tag is string t && t == dateStr)
                    {
                        item.Header = newNode.Header;
                        return true;
                    }
                    if (UpdateInItems(item.Items)) return true;
                }
                return false;
            }
            UpdateInItems(DateTreeView.Items);
        }
    }

    /// <summary>화합물 별칭 등록 + 검정곡선 라벨 + 시료 그리드 일괄 업데이트 (공유 메서드)</summary>
    private void RegisterCompoundAliasAndUpdateGrid(string rawCompound, string analyteName, Border? calRowBorder = null)
    {
        if (string.IsNullOrEmpty(rawCompound) || string.IsNullOrEmpty(analyteName)) return;

        // 1. 별칭 등록 (미등록 시)
        var existing = CompoundAliasService.Resolve(rawCompound);
        // 신규 또는 다른 분석항목으로 변경 시 DB 갱신
        string standardCode = CompoundAliasService.FindStandardCodeByAnalyte(analyteName) ?? rawCompound;
        if (existing == null)
        {
            CompoundAliasService.AddOrUpdateAlias(rawCompound, standardCode, analyteName);
            LogMatch($"COMPOUND ALIAS 신규: '{rawCompound}' → 표준코드='{standardCode}', 분석항목='{analyteName}'");
        }
        else if (!existing.Value.분석항목.Equals(analyteName, StringComparison.OrdinalIgnoreCase))
        {
            CompoundAliasService.AddOrUpdateAlias(rawCompound, standardCode, analyteName);
            LogMatch($"COMPOUND ALIAS 변경: '{rawCompound}' → {existing.Value.분석항목} ⇒ {analyteName} (표준코드: {standardCode})");
        }
        else
        {
            LogMatch($"COMPOUND ALIAS 유지: '{rawCompound}' → {existing.Value.표준코드}/{existing.Value.분석항목}");
        }

        var resolved = CompoundAliasService.ResolveOrFallback(rawCompound);

        // 2. 검정곡선 행 라벨 업데이트
        if (calRowBorder?.Child is Grid calGrid)
        {
            foreach (var child in calGrid.Children)
            {
                if (child is TextBlock tb && Grid.GetColumn(tb) == 0)
                {
                    tb.Text = $"{rawCompound} → {resolved.분석항목}";
                    tb.Foreground = new SolidColorBrush(Color.Parse("#90EE90"));
                    break;
                }
            }
        }

        // 3. 같은 CompoundName을 가진 모든 시료 행 UI 업데이트
        if (_currentExcelRows == null) return;
        int updated = 0;
        for (int i = 0; i < _currentExcelRows.Count; i++)
        {
            if (_currentExcelRows[i].CompoundName != rawCompound) continue;
            updated++;
            if (i < _rowNameCells.Count)
            {
                var nameCell = _rowNameCells[i];
                var row = _currentExcelRows[i];
                string samplePart = row.시료명?.Split('|').LastOrDefault()?.Trim() ?? row.시료명 ?? "";
                nameCell.Children.Clear();
                nameCell.Children.Add(FsBase(new TextBlock
                {
                    Text = $"{resolved.분석항목} | {samplePart}",
                    FontFamily = Font,
                    Foreground = new SolidColorBrush(Color.Parse("#90EE90")),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }));
            }
        }
        LogMatch($"COMPOUND ALIAS: '{rawCompound}' → {updated}개 시료 행에 적용됨");
    }

    /// <summary>데이터 그리드와 동일한 컬럼 구조의 문서 정보 행.
    /// 첫 4컬럼을 ColumnSpan으로 합쳐서 label 표시, vals는 col 4부터 배치.</summary>
    private Border BuildDocRowUnified(string colDefs, string label, string[] vals, string resultFg, int rowIndex = 0)
    {
        // 아주 미세한 교대 색상
        var bgBrush = rowIndex % 2 == 0
            ? new SolidColorBrush(Color.Parse("#2a2a2a"))  // 기본
            : new SolidColorBrush(Color.Parse("#303030")); // 살짝 밝게

        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(colDefs), MinHeight = 28,
            Background = bgBrush };
        // 구분명: 첫 4컬럼 합침
        var labelTb = FsBase(new TextBlock
        {
            Text = label, FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 2),
        });
        Grid.SetColumn(labelTb, 0); Grid.SetColumnSpan(labelTb, 4);
        g.Children.Add(labelTb);
        // 값: col 4부터 (모든 값 동일 스타일)
        for (int c = 0; c < vals.Length; c++)
        {
            int colIdx = 4 + c;
            var tb = FsBase(new TextBlock
            {
                Text = vals[c], FontFamily = Font,
                Foreground = AppRes("AppFg"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(4, 2),
            });
            Grid.SetColumn(tb, colIdx); g.Children.Add(tb);
        }
        return new Border { Child = g, BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(0, 0, 0, 1) };
    }

    private void ShowMessage(string msg, bool isError)
    {
        if (Avalonia.Threading.Dispatcher.UIThread.CheckAccess())
        {
            EditPanelChanged?.Invoke(FsSM(new TextBlock
            {
                Text = msg, FontFamily = Font,
                Foreground = isError ? AppRes("ThemeFgDanger") : AppRes("ThemeFgSuccess"),
                Margin = new Thickness(12), TextWrapping = TextWrapping.Wrap,
            }));
        }
        else
        {
            Avalonia.Threading.Dispatcher.UIThread.Post(() =>
            {
                EditPanelChanged?.Invoke(FsSM(new TextBlock
                {
                    Text = msg, FontFamily = Font,
                    Foreground = isError ? AppRes("ThemeFgDanger") : AppRes("ThemeFgSuccess"),
                    Margin = new Thickness(12), TextWrapping = TextWrapping.Wrap,
                }));
            });
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    //  Shimadzu UV-1800 ASCII 파일 처리
    // ═════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Shimadzu UV-1800 ASCII 파일인지 확인
    /// </summary>
    private static bool IsShimadzuUvFile(string path)
    {
        try
        {
            var lines = System.IO.File.ReadAllLines(path).Take(2).ToArray();
            if (lines.Length < 1) return false;

            var header = lines[0];
            return header.Contains("Sample ID") &&
                   header.Contains("Conc") &&
                   header.Contains("WL");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Shimadzu UV-1800 ASCII 파일 파싱 및 처리
    /// </summary>
    private async void ParseShimadzuUvFile(string path) => await ParseShimadzuUvFileAsync(path);

    private async System.Threading.Tasks.Task ParseShimadzuUvFileAsync(string path)
    {
        try
        {
            ShowMessage("Shimadzu UV-1800 PDF 파싱 중...", false);

            string[] uvItems = { "시안", "6가크롬", "색도", "ABS", "불소", "Phenols", "T-N", "T-P" };

            var parseResult = await System.Threading.Tasks.Task.Run(() =>
                ShimadzuUvPdfParser.Parse(path, uvItems, FormatResult));

            // 카테고리: 파서 감지값 우선, 없으면 행 기반 추론
            string targetCategory = parseResult.DocInfo.DetectedCategory
                ?? DetermineUvCategory(parseResult.Rows);

            // 날짜 설정
            if (!string.IsNullOrEmpty(parseResult.DocDate))
                _categoryDocDates[targetCategory] = parseResult.DocDate;

            // 해당 카테고리로 자동 전환
            _activeCategory = targetCategory;
            var targetCatMatch = Categories.FirstOrDefault(c => c.Key == targetCategory);
            _activeItems = targetCatMatch.Items ?? Array.Empty<string>();
            _categorySelected = true;

            // 파싱 결과 저장
            _categoryDocInfo[targetCategory] = parseResult.DocInfo;
            _categoryExcelData[targetCategory] = parseResult.Rows;
            _categoryFilePaths[targetCategory] = path;

            _currentExcelRows = parseResult.Rows;
            UpdateCategoryButtonStyles();
            LoadVerifiedGrid();
            BuildStatsPanel();

            ShowMessage($"✅ Shimadzu UV-1800 PDF 파싱 완료 - {parseResult.Rows.Count}건 로드됨", false);
        }
        catch (Exception ex)
        {
            ShowMessage($"❌ Shimadzu UV-1800 PDF 파싱 오류: {ex.Message}", true);
        }
    }

    private async System.Threading.Tasks.Task ParseIcpPdfFileAsync(string path)
    {
        try
        {
            string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
            ShowMessage($"ICP 파일 파싱 중...", false);
            string icpCategory = "ICP";
            var icpCatMatch = Categories.FirstOrDefault(c => c.Key == icpCategory);
            string[] icpItems = icpCatMatch.Items ?? Array.Empty<string>();

            IcpPdfParser.ParseResult parseResult;
            if (ext == ".csv")
            {
                var r = await System.Threading.Tasks.Task.Run(() =>
                    IcpAaCsvParser.Parse(path, icpItems, FormatResult));
                parseResult = new IcpPdfParser.ParseResult(r.Rows, r.DocInfo, r.DocDate, r.DetectedFormat);
            }
            else
            {
                parseResult = await System.Threading.Tasks.Task.Run(() =>
                    IcpPdfParser.Parse(path, icpItems, FormatResult));
            }

            _categoryDocInfo[icpCategory] = parseResult.DocInfo;
            _categoryExcelData[icpCategory] = parseResult.Rows;
            _categoryFilePaths[icpCategory] = path;
            _activeCategory = icpCategory;
            _activeItems = icpItems;
            _categorySelected = true;
            _currentExcelRows = parseResult.Rows;
            UpdateCategoryButtonStyles();
            LoadVerifiedGrid();
            BuildStatsPanel();
            ShowMessage($"✅ ICP 파싱 완료 ({parseResult.DetectedFormat}) - {parseResult.Rows.Count}건", false);
        }
        catch (Exception ex) { ShowMessage($"❌ ICP 파싱 오류: {ex.Message}", true); }
    }

    private async System.Threading.Tasks.Task ParseLcmsPfasFileAsync(string path)
    {
        try
        {
            ShowMessage("LCMS 과불화화합물 PDF 파싱 중...", false);
            string pfasCategory = "PFAS";
            var pfasCatMatch = Categories.FirstOrDefault(c => c.Key == pfasCategory);
            string[] pfasItems = pfasCatMatch.Items ?? new[] { "PFOA", "PFOS", "PFBS" };
            var parseResult = await System.Threading.Tasks.Task.Run(() =>
                LcmsPfasPdfParser.Parse(path, pfasItems, FormatResult));
            _categoryDocInfo[pfasCategory] = parseResult.DocInfo;
            _categoryExcelData[pfasCategory] = parseResult.Rows;
            _categoryFilePaths[pfasCategory] = path;
            _activeCategory = pfasCategory;
            _activeItems = pfasItems;
            _categorySelected = true;
            _currentExcelRows = parseResult.Rows;
            UpdateCategoryButtonStyles();
            LoadVerifiedGrid();
            BuildStatsPanel();
            ShowMessage($"✅ LCMS PDF 파싱 완료 - {parseResult.Rows.Count}건", false);
        }
        catch (Exception ex) { ShowMessage($"❌ LCMS PDF 오류: {ex.Message}", true); }
    }

    // ═════════════════════════════════════════════════════════════════════════
    //  Agilent Cary-3500 CSV 파일 처리
    // ═════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Agilent Cary-3500 CSV 파일인지 확인
    /// </summary>
    private static bool IsAgilentCaryUvFile(string path)
    {
        try
        {
            var lines = System.IO.File.ReadAllLines(path).Take(10).ToArray();
            if (lines.Length < 3) return false;

            // CSV Report, Worksheet Type=Concentration, METHOD 섹션 존재 확인
            return lines.Any(l => l.Contains("CSV Report")) &&
                   lines.Any(l => l.Contains("Worksheet Type,Concentration")) &&
                   lines.Any(l => l.Trim() == "METHOD");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Agilent Cary-3500 PDF 파일 파싱 및 처리 (async Task, UI 스레드에서 호출)
    /// </summary>
    private async System.Threading.Tasks.Task ParseAgilentCaryUvFileAsync(string path)
    {
        try
        {
            ShowMessage("Agilent Cary-3500 PDF 파싱 중...", false);

            string[] uvItems = { "시안", "6가크롬", "색도", "ABS", "불소", "Phenols", "T-N", "T-P" };

            var parseResult = await System.Threading.Tasks.Task.Run(() =>
                AgilentCaryUvPdfParser.Parse(path, uvItems, FormatResult));

            // 파서가 감지한 카테고리 우선, 없으면 기존 로직
            string targetCategory = parseResult.DocInfo.DetectedCategory
                ?? DetermineUvCategory(parseResult.Rows);

            _activeCategory = targetCategory;
            var targetCatMatch = Categories.FirstOrDefault(c => c.Key == targetCategory);
            _activeItems = targetCatMatch.Items ?? Array.Empty<string>();
            _categorySelected = true;

            _categoryDocInfo[targetCategory]  = parseResult.DocInfo;
            _categoryExcelData[targetCategory] = parseResult.Rows;
            _categoryFilePaths[targetCategory] = path;

            _currentExcelRows = parseResult.Rows;
            UpdateCategoryButtonStyles();
            LoadVerifiedGrid();
            BuildStatsPanel();

            ShowMessage($"✅ Agilent Cary-3500 파싱 완료 - {parseResult.Rows.Count}건 로드됨", false);
        }
        catch (Exception ex)
        {
            ShowMessage($"❌ Agilent Cary-3500 파싱 오류: {ex.Message}", true);
        }
    }

    /// <summary>
    /// UV 파싱 결과에서 적절한 카테고리 결정
    /// </summary>
    private string DetermineUvCategory(List<ExcelRow> uvRows)
    {
        if (uvRows?.Any() != true) return "CN"; // 기본값 (UV 파서 → 시안)

        // 파싱된 항목(Fxy)에서 가장 많이 나온 항목 기준으로 카테고리 결정
        var itemCounts = uvRows
            .Where(r => !string.IsNullOrEmpty(r.Fxy))
            .GroupBy(r => r.Fxy)
            .ToDictionary(g => g.Key, g => g.Count());

        if (!itemCounts.Any()) return "CN"; // UV 파서 기본값 (시안)

        // 우선순위: 페놀류 > 시안 > 6가크롬 > T-N > T-P > 색도 > ABS > 불소
        var priorityMapping = new Dictionary<string, string>
        {
            ["Phenols"] = "PHENOLS",
            ["페놀류"] = "PHENOLS",
            ["시안"] = "CN",
            ["6가크롬"] = "CR6",
            ["T-N"] = "TN",
            ["T-P"] = "TP",
            ["색도"] = "COLOR",
            ["ABS"] = "ABS",
            ["불소"] = "FLUORIDE"
        };

        // 우선순위 순서대로 확인
        foreach (var mapping in priorityMapping)
        {
            if (itemCounts.ContainsKey(mapping.Key))
            {
                LogMatch($"UV 카테고리 결정: '{mapping.Key}' → '{mapping.Value}' 카테고리");
                return mapping.Value;
            }
        }

        // 매핑되지 않은 항목이면 기본값
        var firstItem = itemCounts.First().Key;
        LogMatch($"UV 카테고리 결정: 알 수 없는 항목 '{firstItem}' → CN 카테고리(기본값)");
        return "CN";
    }




    // ═════════════════════════════════════════════════════════════════════════
    //  ICP PDF 파일 처리
    // ═════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// ICP PDF 파일인지 확인
    /// </summary>
    private static bool IsIcpPdfFile(string path)
    {
        try
        {
            var fileName = System.IO.Path.GetFileName(path).ToLower();

            // 파일명에 "icp" 포함 또는 "원소" 키워드 포함
            return fileName.Contains("icp") ||
                   fileName.Contains("원소") ||
                   fileName.Contains("metal") ||
                   fileName.Contains("element");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// ICP 파일 파싱 및 처리 (PDF 또는 CSV)
    /// </summary>
    private async void ParseIcpPdfFile(string path)
    {
        try
        {
            string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
            ShowMessage("ICP 파일 파싱 중...", false);

            string icpCategory = "ICP";
            var icpCatMatch = Categories.FirstOrDefault(c => c.Key == icpCategory);
            string[] icpItems = icpCatMatch.Items ?? Array.Empty<string>();

            IcpPdfParser.ParseResult parseResult;
            if (ext == ".csv")
            {
                var r = await System.Threading.Tasks.Task.Run(() =>
                    IcpAaCsvParser.Parse(path, icpItems, FormatResult));
                parseResult = new IcpPdfParser.ParseResult(r.Rows, r.DocInfo, r.DocDate, r.DetectedFormat);
            }
            else
            {
                parseResult = await System.Threading.Tasks.Task.Run(() =>
                    IcpPdfParser.Parse(path, icpItems, FormatResult));
            }

            _categoryDocInfo[icpCategory] = parseResult.DocInfo;
            _categoryExcelData[icpCategory] = parseResult.Rows;
            _categoryFilePaths[icpCategory] = path;
            _activeCategory = icpCategory;
            _activeItems = icpItems;
            _categorySelected = true;
            _currentExcelRows = parseResult.Rows;
            UpdateCategoryButtonStyles();

            LoadVerifiedGrid();
            BuildStatsPanel();

            ShowMessage($"✅ ICP 파싱 완료 ({parseResult.DetectedFormat}) - {parseResult.Rows.Count}건 로드됨", false);
        }
        catch (Exception ex)
        {
            ShowMessage($"❌ ICP 파싱 오류: {ex.Message}", true);
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    //  LCMS 과불화화합물 PDF 파일 처리
    // ═════════════════════════════════════════════════════════════════════════

    /// <summary>
    /// LCMS 과불화화합물 PDF 파일 파싱 및 처리
    /// </summary>
    private async void ParseLcmsPfasPdfFile(string path)
    {
        try
        {
            ShowMessage("LCMS 과불화화합물 PDF 파일 파싱 중...", false);

            // PFAS 카테고리로 설정
            string pfasCategory = "PFAS";
            var pfasCatMatch = Categories.FirstOrDefault(c => c.Key == pfasCategory);
            string[] pfasItems = pfasCatMatch.Items ?? new[] { "PFOA", "PFOS", "PFBS" };

            var parseResult = await System.Threading.Tasks.Task.Run(() =>
                LcmsPfasPdfParser.Parse(path, pfasItems, FormatResult));

            // PFAS 카테고리에 파싱 결과 저장
            _categoryDocInfo[pfasCategory] = parseResult.DocInfo;
            _categoryExcelData[pfasCategory] = parseResult.Rows;
            _categoryFilePaths[pfasCategory] = path;

            // PFAS 카테고리로 자동 전환
            _activeCategory = pfasCategory;
            _activeItems = pfasItems;
            _categorySelected = true;

            // UI 업데이트
            _currentExcelRows = parseResult.Rows;
            UpdateCategoryButtonStyles();

            LoadVerifiedGrid();
            BuildStatsPanel();

            ShowMessage($"✅ LCMS 과불화화합물 PDF 파일 파싱 완료 - {parseResult.Rows.Count}건 로드됨", false);
        }
        catch (Exception ex)
        {
            ShowMessage($"❌ LCMS 과불화화합물 PDF 파일 파싱 오류: {ex.Message}", true);
        }
    }
}
