using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
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
        ("PHENOLS", "Phenols",  "🌈", new[] { "Phenols" }),
        ("NHEX",  "NHexan",  "⛽", new[] { "N-Hexan" }),
        ("GCMS",  "GCMS",    "📊", Array.Empty<string>()),
        ("ICP",   "ICP",     "⚡", Array.Empty<string>()),
        ("IC",    "IC",      "🧬", Array.Empty<string>()),
        ("LCMS",  "LCMS",    "💊", Array.Empty<string>()),
        ("CN",    "CN",      "☠️", Array.Empty<string>()),
        ("AAS",   "AAS",     "🔥", Array.Empty<string>()),
        ("COLOR", "COLOR",   "🎨", Array.Empty<string>()),
        ("ECO",   "생태독성", "🐟", Array.Empty<string>()),
        ("ECOLI", "대장균",   "🦠", new[] { "총대장균군" }),
    };

    private string _activeCategory = "BOD";
    private string[] _activeItems = { "BOD" };
    private bool _categorySelected = false; // 카테고리 버튼이 명시적으로 클릭되었는지

    // Items값 → DB약칭 별명 매핑 (분장표준처리 약칭과 다른 경우)
    private static readonly Dictionary<string, string> _itemAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["T-N"] = "TN", ["T-P"] = "TP", ["Phenols"] = "페놀류",
        ["N-Hexan"] = "NH", ["총대장균군"] = "대장균",
    };

    // 약칭 → 소수점 자릿수 캐시 (분석정보 DB 기반)
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
    private string _inputMode = "수질분석센터"; // 수질분석센터 / 비용부담금 / 처리시설
    private readonly Dictionary<string, Button> _categoryButtons = new();

    // 모드별 표시할 카테고리 키
    private static readonly Dictionary<string, string[]> ModeCategoryKeys = new()
    {
        ["수질분석센터"] = new[] { "BOD", "TOC", "SS", "TN", "TP", "PHENOLS", "NHEX", "GCMS", "ICP", "IC", "LCMS", "CN", "AAS", "COLOR", "ECO" },
        ["비용부담금/처리시설"] = new[] { "BOD", "TOC", "SS", "TN", "TP", "PHENOLS", "NHEX", "ECOLI" },
    };

    private static readonly string[] AllEditableItems =
        { "BOD", "TOC", "SS", "T-N", "T-P", "N-Hexan", "Phenols" };

    // ── 카테고리별 첨부 엑셀 데이터 ──────────────────────────────────────────
    private readonly Dictionary<string, List<ExcelRow>> _categoryExcelData = new();
    private readonly Dictionary<string, string> _categoryFilePaths = new();
    private readonly Dictionary<string, string?> _categoryDocDates = new(); // 엑셀 문서 날짜

    private class ExcelRow
    {
        public string 시료명 { get; set; } = "";
        public string SN { get; set; } = "";
        public string Result { get; set; } = "";
        public string 시료량 { get; set; } = "";
        public string D1 { get; set; } = "";
        public string D2 { get; set; } = "";
        public string Fxy { get; set; } = "";  // f(x/y) 식종액 함유율
        public string P { get; set; } = "";    // 희석배수
        public WasteSample? Matched { get; set; }               // 폐수배출업소 매칭
        public AnalysisRequestRecord? MatchedAnalysis { get; set; } // 수질분석센터 매칭
        public FacilityResultRow? MatchedFacility { get; set; }     // 처리시설 매칭
        public string? MatchedFacilityName { get; set; }            // 처리시설명
        public MatchStatus Status { get; set; }
        public SourceType Source { get; set; }
        public bool Enabled { get; set; } = true;
    }
    private enum MatchStatus { 입력가능, 덮어쓰기, 미매칭, 대기 }
    private enum SourceType { 미분류, 폐수배출업소, 수질분석센터, 처리시설 }

    // 엑셀 문서 헤더 정보 (행1~7)
    private class ExcelDocInfo
    {
        public string 문서번호 { get; set; } = "";
        public string 분석방법 { get; set; } = "";
        public string 결과표시 { get; set; } = "";
        public string 관련근거 { get; set; } = "";
        // 식종수의 BOD (행6)
        public string 식종수_시료량 { get; set; } = "";
        public string 식종수_D1 { get; set; } = "";
        public string 식종수_D2 { get; set; } = "";
        public string 식종수_P { get; set; } = "";
        public string 식종수_Result { get; set; } = "";
        public string 식종수_Remark { get; set; } = "";  // 식종수(%) 1.5
        // SCF (행7)
        public string SCF_시료량 { get; set; } = "";
        public string SCF_D1 { get; set; } = "";
        public string SCF_D2 { get; set; } = "";
        public string SCF_Result { get; set; } = "";
        // 검량곡선 (UV VIS 등)
        public bool IsUVVIS { get; set; }
        public bool IsSS { get; set; }
        public string[] Standard_Points { get; set; } = Array.Empty<string>(); // 표준용액 농도
        public string Standard_Slope { get; set; } = "";   // 기울기 a
        public string Standard_Intercept { get; set; } = ""; // 절편 b
        public string[] Abs_Values { get; set; } = Array.Empty<string>(); // 흡광도 측정값
        public string Abs_R2 { get; set; } = "";   // R²
    }
    private readonly Dictionary<string, ExcelDocInfo> _categoryDocInfo = new();

    // ── 외부 연결 ────────────────────────────────────────────────────────────
    public event Action<Control?>? ListPanelChanged;    // Show2
    public event Action<Control?>? EditPanelChanged;    // Show3
    public event Action<Control?>? StatsPanelChanged;   // Show4

    private string? _selectedDate;
    private WasteSample? _selectedSample;
    private List<WasteSample> _currentSamples = new();

    private TopLevel? _topLevel;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        _topLevel = TopLevel.GetTopLevel(this);
    }

    private StackPanel? _gridPanel;
    private Dictionary<string, TextBox> _inputBoxes = new();
    private int _selectedRowIndex = -1;
    private List<ExcelRow>? _currentExcelRows;
    private List<TextBlock> _rowIcons = new();
    private ExcelRow? _currentEditExcelRow;

    public WasteAnalysisInputPage()
    {
        InitializeComponent();
        BuildCategoryButtons();
    }

    public void SetInputMode(string mode)
    {
        _inputMode = mode;
        BuildCategoryButtons();
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

        // 파일첨부 버튼
        var attachBtn = FsSM(new Button
        {
            Content = "📂 파일첨부",
            FontFamily = Font,
            Padding = new Thickness(8, 4),
            Margin = new Thickness(6, 2, 2, 2),
            CornerRadius = new CornerRadius(4),
            Cursor = new Cursor(StandardCursorType.Hand),
            MinWidth = 80, MinHeight = 0,
            Background = AppRes("BtnInfoBg"),
            Foreground = AppRes("BtnInfoFg"),
            HorizontalContentAlignment = HorizontalAlignment.Center,
        });
        attachBtn.Click += (_, _) => AttachExcel();
        AnalysisItemButtons.Children.Add(attachBtn);
        TextShimmer.AttachIfNew(attachBtn);

        UpdateCategoryButtonStyles();
    }

    private void OnCategoryButtonClick(string key)
    {
        // 카테고리 전환
        _categorySelected = true;
        _activeCategory = key;
        var match = Categories.FirstOrDefault(c => c.Key == key);
        _activeItems = match.Items ?? Array.Empty<string>();
        UpdateCategoryButtonStyles();

        // 엑셀 첨부가 있으면 검증 그리드, 없으면 날짜 기반 DB 그리드
        if (_categoryExcelData.ContainsKey(key))
            LoadVerifiedGrid();
        else if (_selectedDate != null)
            LoadSampleGrid(_selectedDate);

        EditPanelChanged?.Invoke(null);
        BuildStatsPanel();
    }

    /// <summary>서브메뉴 "새로고침" 또는 별도 첨부 시 파일 피커 호출</summary>
    public async void AttachExcel()
    {
        var key = _activeCategory;
        var match = Categories.FirstOrDefault(c => c.Key == key);

        var topLevel = _topLevel ?? TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = $"📂 {match.Label} 분석결과 엑셀 첨부",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx", "*.xls" } },
            }
        });

        if (files.Count == 0) return;
        var path = files[0].TryGetLocalPath();
        if (string.IsNullOrEmpty(path)) return;

        // 엑셀 파싱
        var rows = new List<ExcelRow>();
        string? docDate = null;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);
            var ws = wb.Worksheets.First();

            // Row1 B열에서 문서 날짜 추출
            var dateCellVal = ws.Cell(1, 2).GetString().Trim();
            if (DateTime.TryParse(dateCellVal, out var parsedDocDate))
                docDate = parsedDocDate.ToString("yyyy-MM-dd");

            // 문서 헤더 정보 (행1~7)
            var docInfo = new ExcelDocInfo
            {
                문서번호 = ws.Cell(1, 2).GetString().Trim(),
                분석방법 = ws.Cell(2, 2).GetString().Trim(),
                결과표시 = ws.Cell(3, 2).GetString().Trim(),
                관련근거 = ws.Cell(4, 2).GetString().Trim(),
            };

            // 엑셀 내용 기반 형식 감지
            var row5A = ws.Cell(5, 1).GetString().Trim();
            var row7A = ws.Cell(7, 1).GetString().Trim();
            bool isUVVIS = row5A.Equals("STANDARD", StringComparison.OrdinalIgnoreCase);
            bool isSS = row7A.Contains("시료명");

            // 감지된 형식과 선택한 카테고리 불일치 검증
            var uvvisKeys = new HashSet<string> { "TN", "TP", "PHENOLS" };
            string detectedFormat = isSS ? "SS" : isUVVIS ? "UVVIS" : "BOD";
            // BOD 형식: BOD, TOC, NHEX 등 (SS, UVVIS 제외 나머지)
            bool formatMatch = detectedFormat switch
            {
                "SS" => key == "SS",
                "UVVIS" => uvvisKeys.Contains(key),
                _ => key != "SS" && !uvvisKeys.Contains(key), // BOD 형식 계열
            };
            if (!formatMatch)
            {
                ShowMessage($"선택한 카테고리({key})와 엑셀 형식({detectedFormat})이 일치하지 않습니다.", true);
                return;
            }

            if (isSS)
            {
                docInfo.IsSS = true;
                _categoryDocInfo[key] = docInfo;
                // SS: A=시료명, B=시료량, C=전무게, D=후무게, E=전후무게차, F=희석배수, G=결과, H=SN
                ParsePage(ws, rows, colName: 1, colResult: 7, colSN: 8, startRow: 8, itemAbbr: _activeItems.FirstOrDefault() ?? key);
            }
            else if (isUVVIS)
            {
                docInfo.IsUVVIS = true;
                // 행5: STANDARD 표준용액 농도 + 기울기(a) + 절편(b)
                var stdPts = new List<string>();
                for (int c = 2; c <= 6; c++)
                {
                    var v = ws.Cell(5, c).GetString().Trim();
                    if (double.TryParse(v, out var dv)) v = dv.ToString("G");
                    stdPts.Add(v);
                }
                docInfo.Standard_Points = stdPts.ToArray();
                var slope = ws.Cell(5, 7).GetString().Trim();
                if (double.TryParse(slope, out var sv)) slope = sv.ToString("G6");
                docInfo.Standard_Slope = slope;
                var intercept = ws.Cell(5, 8).GetString().Trim();
                if (double.TryParse(intercept, out var iv)) intercept = iv.ToString("G6");
                docInfo.Standard_Intercept = intercept;

                // 행6: abs 흡광도 + R²
                var absVals = new List<string>();
                for (int c = 2; c <= 6; c++)
                {
                    var v = ws.Cell(6, c).GetString().Trim();
                    if (double.TryParse(v, out var av)) v = av.ToString("G6");
                    absVals.Add(v);
                }
                docInfo.Abs_Values = absVals.ToArray();
                var r2 = ws.Cell(6, 7).GetString().Trim();
                if (double.TryParse(r2, out var r2v)) r2 = r2v.ToString("F5");
                docInfo.Abs_R2 = r2;

                _categoryDocInfo[key] = docInfo;
                // UV VIS: colResult=6 (BOD는 7), 페이지2는 colResult=14
                var uvItem = _activeItems.FirstOrDefault() ?? key;
                ParsePage(ws, rows, colName: 1, colResult: 6, colSN: 8, startRow: 8, itemAbbr: uvItem);
                ParsePage(ws, rows, colName: 9, colResult: 14, colSN: 16, startRow: 8, itemAbbr: uvItem);
            }
            else
            {
                // BOD 형식: 행6 식종수, 행7 SCF
                docInfo.식종수_시료량 = ws.Cell(6, 2).GetString().Trim();
                docInfo.식종수_D1 = ws.Cell(6, 3).GetString().Trim();
                docInfo.식종수_D2 = ws.Cell(6, 4).GetString().Trim();
                docInfo.식종수_P = ws.Cell(6, 6).GetString().Trim();
                var r6r = ws.Cell(6, 7).GetString().Trim();
                if (double.TryParse(r6r, out var r6v)) r6r = r6v.ToString("F1");
                docInfo.식종수_Result = r6r;
                docInfo.식종수_Remark = ws.Cell(6, 8).GetString().Trim();
                docInfo.SCF_시료량 = ws.Cell(7, 2).GetString().Trim();
                docInfo.SCF_D1 = ws.Cell(7, 3).GetString().Trim();
                docInfo.SCF_D2 = ws.Cell(7, 4).GetString().Trim();
                var r7r = ws.Cell(7, 7).GetString().Trim();
                if (double.TryParse(r7r, out var r7v)) r7r = r7v.ToString("F4");
                docInfo.SCF_Result = r7r;
                _categoryDocInfo[key] = docInfo;

                var bodItem = _activeItems.FirstOrDefault() ?? key;
                ParsePage(ws, rows, colName: 1, colResult: 7, colSN: 8, startRow: 8, itemAbbr: bodItem);
                ParsePage(ws, rows, colName: 9, colResult: 15, colSN: 16, startRow: 8, itemAbbr: bodItem);
            }
        }
        catch (Exception ex)
        {
            ShowMessage($"엑셀 파싱 오류: {ex.Message}", true);
            return;
        }

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
        LoadVerifiedGrid();
    }

    private void UpdateCategoryButtonStyles()
    {
        foreach (var (k, btn) in _categoryButtons)
        {
            bool active = k == _activeCategory;
            bool hasFile = _categoryExcelData.ContainsKey(k);

            var cat = Categories.FirstOrDefault(c => c.Key == k);
            string label = cat.Label ?? k;
            string icon = cat.Icon ?? "";

            if (hasFile)
                btn.Content = $"✅ {label}";
            else
                btn.Content = $"{icon} {label}";

            btn.Background = active ? AppRes("BtnPrimaryBg")
                : hasFile ? AppRes("BtnSuccessBg") : AppRes("BtnBg");
            btn.Foreground = active ? AppRes("BtnPrimaryFg")
                : hasFile ? AppRes("BtnSuccessFg") : AppRes("BtnFg");
            btn.BorderBrush = active ? AppRes("BtnPrimaryBorder")
                : hasFile ? AppRes("BtnSuccessBorder") : AppRes("BtnBorder");
            btn.BorderThickness = new Thickness(1);
        }
    }

    // =========================================================================
    // 데이터 로드 (날짜 트리뷰)
    // =========================================================================
    public async void LoadData()
    {
        DateTreeView.Items.Clear();
        _selectedDate = null;
        ListPanelChanged?.Invoke(null);
        EditPanelChanged?.Invoke(null);

        try
        {
            // DB 조회 전부 백그라운드에서 실행
            var byMonth = await System.Threading.Tasks.Task.Run(() =>
            {
                var dates = WasteSampleService.GetDates();
                return dates
                    .GroupBy(d => d[..7])
                    .OrderByDescending(g => g.Key)
                    .Select(g => (
                        Month: g.Key,
                        Dates: g.OrderByDescending(x => x)
                                .Select(d => (Date: d, Samples: WasteSampleService.GetByDate(d)))
                                .ToList()
                    ))
                    .ToList();
            });

            // UI 노드 생성은 메인 스레드에서
            foreach (var month in byMonth)
            {
                var monthNode = MakeMonthNode(month.Month, month.Dates.Count);
                foreach (var (d, samples) in month.Dates)
                    monthNode.Items.Add(MakeDateNode(d, samples));
                monthNode.IsExpanded = true;
                DateTreeView.Items.Add(monthNode);
            }
        }
        catch { }
    }

    // =========================================================================
    // 트리뷰 노드
    // =========================================================================
    private static TreeViewItem MakeMonthNode(string ym, int count)
    {
        DateTime.TryParse(ym + "-01", out var d);
        return new TreeViewItem
        {
            IsExpanded = true,
            Header = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    new TextBlock
                    {
                        Text = $"📅  {d.Year}년 {d.Month}월",
                        FontWeight = FontWeight.SemiBold,
                        FontFamily = Font, Foreground = AppTheme.FgMuted,
                        VerticalAlignment = VerticalAlignment.Center,
                    }.BindMD(),
                    new Border
                    {
                        Background = AppTheme.BorderSubtle, CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5,1),
                        Child = new TextBlock
                        {
                            Text = count.ToString(),
                            FontFamily = Font, Foreground = AppTheme.FgDimmed,
                        }.BindXS()
                    }
                }
            }
        };
    }

    private static TreeViewItem MakeDateNode(string dateStr, List<WasteSample> samples)
    {
        DateTime.TryParse(dateStr, out var d);
        string dow = d == DateTime.MinValue ? "" : $" ({DayOfWeekKr(d)})";
        string label = d == DateTime.MinValue ? dateStr : $"🗓  {d.Month}/{d.Day}{dow}";

        int total = samples.Count;
        int filled = samples.Count(s =>
            !string.IsNullOrWhiteSpace(s.BOD) || !string.IsNullOrWhiteSpace(s.TOC) ||
            !string.IsNullOrWhiteSpace(s.SS));

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
                    new Border
                    {
                        Background = filled == total && total > 0
                            ? AppTheme.BgActiveGreen : AppTheme.BorderSubtle,
                        CornerRadius = new CornerRadius(8),
                        Padding = new Thickness(5,1),
                        Child = new TextBlock
                        {
                            Text = $"{filled}/{total}",
                            FontFamily = Font,
                            Foreground = filled == total && total > 0
                                ? AppTheme.FgSuccess : AppTheme.FgDimmed,
                        }.BindXS()
                    }
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
    // 트리뷰 선택 → Show2
    // =========================================================================
    private void DateTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (DateTreeView.SelectedItem is not TreeViewItem tvi) return;
        if (tvi.Tag is not string dateStr) return;

        _selectedDate = dateStr;
        _selectedSample = null;

        // 카테고리 미선택 시 데이터 로드 안함
        if (!_categorySelected) return;

        _currentSamples = WasteSampleService.GetByDate(dateStr);

        // 첨부 파일이 있으면 검증 그리드, 없으면 일반 그리드
        if (_categoryExcelData.ContainsKey(_activeCategory))
            LoadVerifiedGrid();
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
            if (_selectedDate != null) LoadSampleGrid(_selectedDate);
            return;
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

        // 처리시설 마스터 시료명 로드
        List<(string 시설명, string 시료명, int 마스터Id)>? facilityMasters = null;
        try { facilityMasters = FacilityResultService.GetAllMasterSamples(); } catch { }

        // 3개 테이블 동시 매칭
        foreach (var row in excelRows)
        {
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
                       ?? _currentSamples.FirstOrDefault(s => s.업체명 == row.시료명);

            if (row.Matched != null)
            {
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
                    // FacilityResultRow 로드는 ImportData에서 처리
                    row.Status = MatchStatus.입력가능;
                    continue;
                }
            }

            // 3순위: 수질분석센터 (약칭/시료명 매칭)
            var ar = analysisRecords.FirstOrDefault(a => a.약칭 == row.시료명 || a.시료명 == row.시료명);
            if (ar != null)
            {
                row.MatchedAnalysis = ar;
                row.Source = SourceType.수질분석센터;
                row.Status = MatchStatus.입력가능;
                continue;
            }

            // 미매칭
            row.Source = SourceType.미분류;
            row.Status = _currentSamples.Count > 0 || analysisRecords.Count > 0
                ? MatchStatus.미매칭 : MatchStatus.대기;
        }

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };

        // 헤더
        var catLabel = Categories.FirstOrDefault(c => c.Key == _activeCategory).Label ?? _activeCategory;
        string fileName = _categoryFilePaths.TryGetValue(_activeCategory, out var fp) ? Path.GetFileName(fp) : "";
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
        if (pending > 0)
            badgePanel.Children.Add(FsXS(new TextBlock { Text = $"⚪ {pending} 대기", FontFamily = Font,
                Foreground = AppRes("FgMuted") }));
        if (matchNew > 0)
            badgePanel.Children.Add(FsXS(new TextBlock { Text = $"🟢 {matchNew} 입력가능", FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess") }));
        if (matchExist > 0)
            badgePanel.Children.Add(FsXS(new TextBlock { Text = $"🟡 {matchExist} 자료있음", FontFamily = Font,
                Foreground = AppRes("ThemeFgWarn") }));
        if (noMatch > 0)
            badgePanel.Children.Add(FsXS(new TextBlock { Text = $"🔴 {noMatch} 중복확인", FontFamily = Font,
                Foreground = AppRes("ThemeFgDanger") }));
        badgePanel.Children.Add(FsXS(new TextBlock { Text = $"합계 {excelRows.Count}건", FontFamily = Font,
            Foreground = AppRes("FgMuted") }));

        // 문서 정보 패널
        var headerContent = new StackPanel { Spacing = 4 };
        headerContent.Children.Add(FsBase(new TextBlock
        {
            Text = $"📋 {dateLabel}  [{catLabel}]  📎 {fileName}",
            FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
        }));

        // 문서 헤더 — 식종수/SCF 또는 검량곡선 테이블
        bool isUVVISMode = false;
        bool isSSMode = false;

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
            isSSMode = docInfo.IsSS;
            if (isSSMode)
            {
                colDefs = "32,50,90,190,60,60,60,60,50,75,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }
            else if (isUVVISMode)
            {
                colDefs = "32,50,90,190,60,65,65,75,75,80";
                colWidths = colDefs.Split(',').Select(double.Parse).ToArray();
            }

            // 분석방법
            if (!string.IsNullOrWhiteSpace(docInfo.분석방법))
                headerContent.Children.Add(FsXS(new TextBlock
                {
                    Text = $"분석방법: {docInfo.분석방법}", FontFamily = Font,
                    Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 2, 0, 0),
                }));

            // 식종수/SCF 또는 검량곡선 테이블 — 데이터 그리드와 동일 구조로 별도 행 배치
            if (docInfo.IsUVVIS)
            {
                bool hasStd = docInfo.Standard_Points.Length > 0;
                bool hasAbs = docInfo.Abs_Values.Length > 0;
                if (hasStd || hasAbs)
                {
                    docTbl = new StackPanel { Spacing = 0 };
                    var hdr = MakeRowGrid();
                    hdr.MinHeight = 26; hdr.Background = AppRes("GridHeaderBg");
                    var hdrLabel = FsBase(new TextBlock { Text = "구분", FontFamily = Font,
                        FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                        VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 0) });
                    Grid.SetColumn(hdrLabel, 0); Grid.SetColumnSpan(hdrLabel, 4);
                    hdr.Children.Add(hdrLabel);
                    string[] uvCols = { "시료량", "흡광도", "희석배수", "계산농도", "결과값", "시료구분" };
                    for (int c = 0; c < uvCols.Length; c++)
                    {
                        var tb = FsBase(new TextBlock { Text = uvCols[c], FontFamily = Font,
                            FontWeight = FontWeight.SemiBold, Foreground = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(4, 0) });
                        Grid.SetColumn(tb, 4 + c); hdr.Children.Add(tb);
                    }
                    docTbl.Children.Add(new Border { Child = hdr,
                        BorderBrush = AppRes("ThemeBorderSubtle"), BorderThickness = new Thickness(0,0,0,1) });

                    if (hasStd)
                        docTbl.Children.Add(BuildDocRowUnified(colDefs, "STANDARD",
                            new[] { docInfo.Standard_Points.ElementAtOrDefault(0) ?? "",
                                    docInfo.Standard_Points.ElementAtOrDefault(1) ?? "",
                                    docInfo.Standard_Points.ElementAtOrDefault(2) ?? "",
                                    $"a={docInfo.Standard_Slope}",
                                    $"b={docInfo.Standard_Intercept}", "" }, "ThemeFgWarn"));
                    if (hasAbs)
                        docTbl.Children.Add(BuildDocRowUnified(colDefs, "Absorbance",
                            new[] { docInfo.Abs_Values.ElementAtOrDefault(0) ?? "",
                                    docInfo.Abs_Values.ElementAtOrDefault(1) ?? "",
                                    docInfo.Abs_Values.ElementAtOrDefault(2) ?? "",
                                    docInfo.Abs_Values.ElementAtOrDefault(3) ?? "",
                                    $"R²={docInfo.Abs_R2}", "" }, "ThemeFgInfo"));
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

                    if (hasSeed)
                        docTbl.Children.Add(BuildDocRowUnified(colDefs, "식종수의 BOD",
                            new[] { docInfo.식종수_시료량, docInfo.식종수_D1, docInfo.식종수_D2,
                                    "-", docInfo.식종수_P, docInfo.식종수_Result, docInfo.식종수_Remark }, "ThemeFgWarn"));
                    if (hasScf)
                        docTbl.Children.Add(BuildDocRowUnified(colDefs, "SCF(식종희석수)",
                            new[] { docInfo.SCF_시료량, docInfo.SCF_D1, docInfo.SCF_D2,
                                    "-", "1", docInfo.SCF_Result, "" }, "ThemeFgInfo"));
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
        Grid.SetRow(badgePanel, 2);
        root.Children.Add(badgePanel);

        // 식종수/검량곡선 테이블 (데이터 그리드와 동일 마진/구조)
        if (docTbl != null)
        {
            var docBorder = new Border
            {
                Child = docTbl,
                BorderBrush = AppRes("ThemeBorderSubtle"),
                BorderThickness = new Thickness(1, 0, 1, 1),
                ClipToBounds = true,
            };
            Grid.SetRow(docBorder, 1);
            root.Children.Add(docBorder);
        }

        // 그리드 본체
        _gridPanel = new StackPanel { Spacing = 0 };

        // 컬럼 헤더
        var colHeaderGrid = MakeRowGrid();
        colHeaderGrid.MinHeight = 28;
        colHeaderGrid.Background = AppRes("GridHeaderBg");

        string[] hLabels = isSSMode
            ? new[] { "", "입력", "SN", "시료명", "시료량", "전무게", "후무게", "무게차", "희석배수", "결과값", "시료구분" }
            : isUVVISMode
            ? new[] { "", "입력", "SN", "시료명", "시료량", "흡광도", "희석배수", "계산농도", "결과값", "시료구분" }
            : new[] { "", "입력", "SN", "시료명", "시료량", "D1", "D2", "f(x/y)", "P", "결과값", "시료구분" };
        int detailStart = 4, detailEnd = isUVVISMode ? 7 : 8;
        for (int c = 0; c < hLabels.Length; c++)
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
        _gridPanel.Children.Add(new Border
        {
            Child = colHeaderGrid,
            BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        });

        // 데이터 행
        _rowIcons = new List<TextBlock>();
        for (int i = 0; i < excelRows.Count; i++)
        {
            var row = excelRows[i];
            var (icon, iconColor) = row.Status switch
            {
                MatchStatus.입력가능 => ("🟢", AppRes("ThemeFgSuccess")),
                MatchStatus.덮어쓰기 => ("🟡", AppRes("ThemeFgWarn")),
                MatchStatus.대기     => ("⚪", AppRes("FgMuted")),
                _                    => ("🔴", AppRes("ThemeFgDanger")),
            };

            var rowGrid = MakeRowGrid();
            rowGrid.MinHeight = 34;
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
                capturedRow.Enabled = false;
                toggleTrack.Background = AppRes("ThemeBorderMuted");
                toggleKnob.Margin = new Thickness(2, 2, 0, 2);
                toggleTrack.Opacity = 0.4;
            }
            else if (row.Status == MatchStatus.대기)
            {
                toggleTrack.Opacity = 0.6;
            }

            toggleTrack.PointerPressed += (_, _) =>
            {
                if (capturedRow.Status == MatchStatus.미매칭) return;
                capturedRow.Enabled = !capturedRow.Enabled;
                toggleTrack.Background = capturedRow.Enabled ? AppRes("BtnPrimaryBg") : AppRes("ThemeBorderMuted");
                toggleKnob.Margin = new Thickness(capturedRow.Enabled ? 18 : 2, 2, 0, 2);
            };

            // 아이콘 (col 0)
            var iconTb = new TextBlock { Text = icon, FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                ClipToBounds = false };
            Grid.SetColumn(iconTb, 0);
            rowGrid.Children.Add(iconTb);
            _rowIcons.Add(iconTb);

            // 토글 스위치 (col 1)
            var toggleWrap = new Border
            {
                Child = toggleTrack,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(2, 4),
            };
            Grid.SetColumn(toggleWrap, 1);
            rowGrid.Children.Add(toggleWrap);

            // SN
            var snTb = FsBase(new TextBlock
            {
                Text = row.SN, FontFamily = Font,
                Foreground = AppRes("ThemeFgInfo"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
            });
            Grid.SetColumn(snTb, 2);
            rowGrid.Children.Add(snTb);

            // 시료명
            var nameTb = FsBase(new TextBlock
            {
                Text = row.시료명, FontFamily = Font,
                Foreground = AppRes("AppFg"),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(4, 0),
            });
            Grid.SetColumn(nameTb, 3);
            rowGrid.Children.Add(nameTb);

            // 기초정보: UV VIS는 시료량/흡광도/희석배수/계산농도 (4열), BOD는 시료량/D1/D2/f(x/y)/P (5열)
            string[] infoVals = isUVVISMode
                ? new[] { row.시료량, row.D1, row.D2, row.Fxy }
                : new[] { row.시료량, row.D1, row.D2, row.Fxy, row.P };
            for (int ci = 0; ci < infoVals.Length; ci++)
            {
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

            int colResult = 4 + infoVals.Length;
            int colSource = colResult + 1;

            // 결과값
            var valTb = FsBase(new TextBlock
            {
                Text = row.Result, FontFamily = Font,
                Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            });
            Grid.SetColumn(valTb, colResult);
            rowGrid.Children.Add(valTb);

            // 시료구분
            var (srcLabel, srcFg) = row.Source switch
            {
                SourceType.폐수배출업소 => ($"폐수배출-{row.Matched?.구분 ?? "?"}", "ThemeFgInfo"),
                SourceType.수질분석센터 => ("수질분석", "ThemeFgSuccess"),
                SourceType.처리시설     => ("처리시설", "ThemeFgWarn"),
                _                      => ("—", "FgMuted"),
            };
            var srcTb = FsBase(new TextBlock
            {
                Text = srcLabel, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = AppRes(srcFg),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            });
            Grid.SetColumn(srcTb, colSource);
            rowGrid.Children.Add(srcTb);

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
            border.PointerPressed += (_, _) => SelectGridRow(capturedIdx);
            TextShimmer.AttachHover(border);
            _gridPanel.Children.Add(border);
        }

        var scroll = new ScrollViewer
        {
            Content = _gridPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        var gridBorder = new Border
        {
            Child = scroll,
            CornerRadius = new CornerRadius(4),
            BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(1),
            ClipToBounds = true,
        };
        Grid.SetRow(gridBorder, 3);
        root.Children.Add(gridBorder);

        // 키보드 상하 이동 지원
        _currentExcelRows = excelRows;
        root.Focusable = true;
        root.KeyDown += OnGridKeyDown;

        ListPanelChanged?.Invoke(root);

        // 포커스 설정 (키보드 이벤트 수신용)
        Avalonia.Threading.Dispatcher.UIThread.Post(() => root.Focus(), Avalonia.Threading.DispatcherPriority.Render);

        // 이전 선택 행 복원 (저장 후 리프레시 시), 없으면 첫 행
        int restoreIdx = _selectedRowIndex >= 0 && _selectedRowIndex < excelRows.Count
            ? _selectedRowIndex : 0;
        _selectedRowIndex = -1; // 강제 재선택
        SelectGridRow(restoreIdx);
    }

    private void SelectGridRow(int index)
    {
        if (_currentExcelRows == null || _gridPanel == null) return;
        if (index < 0 || index >= _currentExcelRows.Count) return;

        // 이전 선택 하이라이트 해제 (첫 child는 헤더)
        if (_selectedRowIndex >= 0 && _selectedRowIndex + 1 < _gridPanel.Children.Count)
        {
            if (_gridPanel.Children[_selectedRowIndex + 1] is Border prevBorder)
            {
                prevBorder.Background = null;
                prevBorder.BorderBrush = AppRes("ThemeBorderSubtle");
                prevBorder.BorderThickness = new Thickness(0, 0, 0, 1);
            }
        }

        _selectedRowIndex = index;

        // 새 행 하이라이트 (선택 행은 좌측 강조 바 + 배경 변경)
        if (index + 1 < _gridPanel.Children.Count && _gridPanel.Children[index + 1] is Border border)
        {
            border.Background = AppRes("ThemeBorderActive");
            border.BorderBrush = AppRes("BtnPrimaryBg");
            border.BorderThickness = new Thickness(3, 0, 0, 1);
            border.BringIntoView();
        }

        // Show3 표시
        var row = _currentExcelRows[index];
        if (row.Matched != null)
        {
            _selectedSample = row.Matched;
            ShowEditForm(row.Matched, row);
        }
        else
        {
            ShowExcelRowDetail(row);
        }
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
    // Show2: 일반 그리드 (첨부 파일 없을 때)
    // =========================================================================
    private bool _detailMode = false;

    /// <summary>카테고리 키 → *_DATA 테이블명</summary>
    private static string? GetDataTableName(string catKey) => catKey switch
    {
        "BOD" => "BOD_DATA",
        "SS" => "SS_DATA",
        "NHEX" => "NHexan_DATA",
        "TN" => "TN_DATA",
        "TP" => "TP_DATA",
        "PHENOLS" => "Phenols_DATA",
        _ => null
    };

    /// <summary>카테고리별 Detail 컬럼 정의</summary>
    private static string[] GetDetailColumns(string catKey) => catKey switch
    {
        "BOD" => new[] { "시료량", "D1", "D2", "희석배수", "식종시료량", "식종D1", "식종D2", "식종BOD" },
        "SS" => new[] { "시료량", "전무게", "후무게", "무게차", "희석배수" },
        "NHEX" => new[] { "시료량" },
        "TN" or "TP" or "PHENOLS" => new[] { "시료량", "흡광도", "희석배수", "검량선_a", "농도" },
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
        bool isUV = catKey is "TN" or "TP" or "PHENOLS";
        bool isSS = catKey == "SS";
        bool isBOD = catKey == "BOD";

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

        _gridPanel = new StackPanel { Spacing = 0 };

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
                    else if (hLabels[c] == "업체명") { val = s.업체명; fg = AppRes("AppFg"); }
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
            border.PointerPressed += (_, _) => { _selectedSample = s; ShowEditForm(s); };
            TextShimmer.AttachHover(border);
            _gridPanel.Children.Add(border);
        }

        var scroll = new ScrollViewer { Content = _gridPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);
        ListPanelChanged?.Invoke(root);
    }

    private static string GetSampleValue(WasteSample s, string item) => item switch
    {
        "BOD" => s.BOD, "TOC" => s.TOC, "SS" => s.SS,
        "T-N" => s.TN, "T-P" => s.TP, "N-Hexan" => s.NHexan, "Phenols" => s.Phenols,
        _ => ""
    };

    // =========================================================================
    // Show3: 편집 폼
    // =========================================================================
    private void ShowEditForm(WasteSample sample, ExcelRow? exRow = null)
    {
        _inputBoxes.Clear();
        _currentEditExcelRow = exRow;
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
                Text = $"검량곡선: y = {slope:G6}x + {intercept:G6}" });

            formulaPanel.Children.Add(fmTitle);
            formulaPanel.Children.Add(fmCurve);
            formulaPanel.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle") });
            formulaPanel.Children.Add(fmCalc);
            formulaPanel.Children.Add(fmRes);

            void RecalcUV()
            {
                if (!double.TryParse(absInput.Text, out var abs)) return;
                if (!double.TryParse(dilInput.Text, out var dil)) return;

                double calcConc = slope * abs + intercept;
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

        root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 6) });

        EditPanelChanged?.Invoke(new ScrollViewer { Content = root });

        // Show4에 계산 수식 표시
        if (formulaPanel.Children.Count > 0)
            StatsPanelChanged?.Invoke(new ScrollViewer { Content = formulaPanel });

    }

    // =========================================================================
    // 저장
    // =========================================================================
    /// <summary>Show3 저장 버튼 — ExcelRow만 업데이트, Show2 그리드 갱신 (서버 저장 안함)</summary>
    private void SaveCurrentSample()
    {
        if (_currentEditExcelRow == null) return;

        // Show2 그리드 갱신
        if (_categoryExcelData.ContainsKey(_activeCategory))
            LoadVerifiedGrid();
        BuildStatsPanel();
        ShowMessage("✅ 적용 완료 (서버 반영은 '입력' 버튼 사용)", false);
    }

    // ─── *_DATA 원시 측정값 저장 헬퍼 ──────────────────────────────────────
    private void SaveRawData(ExcelRow row, WasteSample s)
    {
        if (string.IsNullOrWhiteSpace(row.Result)) return;

        _categoryDocInfo.TryGetValue(_activeCategory, out var docInfo);
        bool isUV = docInfo?.IsUVVIS == true;

        // 분석일: B1 값 우선, 없으면 채수일
        _categoryDocDates.TryGetValue(_activeCategory, out var 분석일Raw);
        string 분석일 = !string.IsNullOrEmpty(분석일Raw) ? 분석일Raw : s.채수일;

        switch (_activeCategory)
        {
            case "BOD":
                WasteSampleService.UpsertBodData(
                    분석일, s.SN, s.업체명, s.구분,
                    시료량: row.시료량, d1: row.D1, d2: row.D2,
                    희석배수: row.P, 결과: row.Result,
                    식종시료량: docInfo?.식종수_시료량 ?? "",
                    식종D1:     docInfo?.식종수_D1    ?? "",
                    식종D2:     docInfo?.식종수_D2    ?? "",
                    식종BOD:    docInfo?.식종수_Result ?? "",
                    식종함유량: docInfo?.식종수_Remark  ?? "");
                break;

            case "SS":
                WasteSampleService.UpsertSsData(
                    분석일, s.SN, s.업체명, s.구분,
                    row.시료량, row.D1, row.D2, row.Fxy, row.P, row.Result);
                break;

            case "NHEX":
                WasteSampleService.UpsertSimpleData(
                    "NHexan_DATA", "결과",
                    분석일, s.SN, s.업체명, s.구분,
                    row.시료량, row.Result);
                break;

            case "TN" when isUV:
            case "TP" when isUV:
            case "PHENOLS" when isUV:
                // T-N, T-P, Phenols 모두 같은 엑셀 행에서 나오므로 활성 항목별로 저장
                foreach (var item in _activeItems)
                {
                    string tblName = item switch
                    {
                        "T-N"     => "TN_DATA",
                        "T-P"     => "TP_DATA",
                        "Phenols" => "Phenols_DATA",
                        _         => ""
                    };
                    if (string.IsNullOrEmpty(tblName)) continue;
                    WasteSampleService.UpsertUvvisData(
                        tblName,
                        분석일, s.SN, s.업체명, s.구분,
                        시료량:  row.시료량,
                        흡광도:  row.D1,
                        희석배수: row.D2,
                        검량선a: docInfo?.Standard_Slope ?? "",
                        농도:    row.Result);
                }
                break;
        }
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

        bool isUV = _categoryDocInfo.TryGetValue(_activeCategory, out var di) && di.IsUVVIS;
        var fields = isUV
            ? new[] { ("시료량", exRow.시료량), ("흡광도", exRow.D1), ("희석배수", exRow.D2), ("계산농도", exRow.Fxy) }
            : new[] { ("시료량", exRow.시료량), ("D1", exRow.D1), ("D2", exRow.D2), ("f(x/y)", exRow.Fxy), ("P", exRow.P) };

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
        root.Children.Add(new Border
        {
            Child = detailGrid, Background = AppRes("GridRowAltBg"),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 6),
        });

        EditPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // =========================================================================
    // Show4: 입력 현황 통계
    // =========================================================================
    private void BuildStatsPanel()
    {
        if (_selectedDate == null || _currentSamples.Count == 0)
        { StatsPanelChanged?.Invoke(null); return; }

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(10) };
        root.Children.Add(FsMD(new TextBlock
        {
            Text = "📊 입력 현황", FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = AppRes("AppFg"),
        }));

        int total = _currentSamples.Count;
        var itemStats = new (string Name, Func<WasteSample, string> Getter)[]
        {
            ("BOD", s => s.BOD), ("TOC", s => s.TOC), ("SS", s => s.SS),
            ("T-N", s => s.TN), ("T-P", s => s.TP),
            ("N-Hexan", s => s.NHexan), ("Phenols", s => s.Phenols),
        };

        foreach (var (name, getter) in itemStats)
        {
            int filled = _currentSamples.Count(s => !string.IsNullOrWhiteSpace(getter(s)));
            double pct = total > 0 ? (double)filled / total * 100 : 0;
            bool isActive = _activeItems.Contains(name);
            bool hasCatFile = _categoryExcelData.Keys.Any(k =>
                Categories.FirstOrDefault(c => c.Key == k).Items?.Contains(name) == true);

            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("22,60,*,50"), Margin = new Thickness(0, 2) };

            // 첨부 상태 아이콘
            var catIcon = FsXS(new TextBlock
            {
                Text = hasCatFile ? "📎" : "  ",
                VerticalAlignment = VerticalAlignment.Center,
            });
            Grid.SetColumn(catIcon, 0); row.Children.Add(catIcon);

            var label = FsXS(new TextBlock
            {
                Text = name,
                FontWeight = isActive ? FontWeight.Bold : FontWeight.SemiBold,
                FontFamily = Font,
                Foreground = isActive ? AppRes("ThemeFgInfo") : AppRes("ThemeFgSecondary"),
                VerticalAlignment = VerticalAlignment.Center,
            });
            Grid.SetColumn(label, 1); row.Children.Add(label);

            var barBg = new Border
            {
                Background = AppRes("ThemeBorderSubtle"), CornerRadius = new CornerRadius(3),
                Height = 10, HorizontalAlignment = HorizontalAlignment.Stretch,
            };
            var barFill = new Border
            {
                Background = pct >= 100 ? AppRes("ThemeFgSuccess") : pct >= 50 ? AppRes("ThemeFgWarn") : AppRes("ThemeFgDanger"),
                CornerRadius = new CornerRadius(3), Height = 10, Width = 0,
                HorizontalAlignment = HorizontalAlignment.Left,
            };
            var barGrid = new Grid { HorizontalAlignment = HorizontalAlignment.Stretch };
            barGrid.Children.Add(barBg); barGrid.Children.Add(barFill);
            Grid.SetColumn(barGrid, 2); row.Children.Add(barGrid);

            var pctText = FsXS(new TextBlock
            {
                Text = $"{filled}/{total}", FontFamily = Font,
                Foreground = AppRes("FgMuted"), HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center,
            });
            Grid.SetColumn(pctText, 3); row.Children.Add(pctText);
            root.Children.Add(row);

            barGrid.Loaded += (_, _) => { barFill.Width = barGrid.Bounds.Width * pct / 100; };
        }

        // 구분별 요약
        root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 6) });
        foreach (var g in _currentSamples.GroupBy(s => s.구분).OrderBy(g => g.Key))
        {
            int gTotal = g.Count();
            int gFilled = g.Count(s => !string.IsNullOrWhiteSpace(s.BOD) || !string.IsNullOrWhiteSpace(s.TOC));
            string icon = g.Key switch { "여수" => "🔵", "율촌" => "🟢", "세풍" => "🟡", _ => "⚪" };
            root.Children.Add(new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 6,
                Children =
                {
                    FsSM(new TextBlock { Text = $"{icon} {g.Key}", FontFamily = Font, Foreground = AppRes("AppFg") }),
                    FsSM(new TextBlock { Text = $"{gFilled}/{gTotal}", FontFamily = Font, Foreground = AppRes("FgMuted") }),
                }
            });
        }

        StatsPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // =========================================================================
    // 엑셀 파싱
    // =========================================================================
    private static void ParsePage(IXLWorksheet ws, List<ExcelRow> rows,
        int colName, int colResult, int colSN, int startRow, string itemAbbr = "BOD")
    {
        // 기초정보 컬럼: 시료량=colName+1, D1=colName+2, D2=colName+3, f(x/y)=colName+4, P=colName+5
        int colVol = colName + 1, colD1 = colName + 2, colD2 = colName + 3;
        int colFxy = colName + 4, colP = colName + 5;

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        for (int r = startRow; r <= lastRow; r++)
        {
            var nameCell   = ws.Cell(r, colName).GetString().Trim();
            var resultCell = ws.Cell(r, colResult).GetString().Trim();
            var snCell     = ws.Cell(r, colSN).GetString().Trim();

            if (string.IsNullOrEmpty(nameCell) && string.IsNullOrEmpty(snCell)) break;
            if (string.IsNullOrEmpty(nameCell)) continue;
            if (nameCell.Contains("식종") || nameCell.Contains("SCF") || nameCell.Contains("분석담당")) continue;

            if (double.TryParse(resultCell, out var val))
                resultCell = FormatResult(val.ToString(), itemAbbr);

            var exRow = new ExcelRow { 시료명 = nameCell, SN = snCell, Result = resultCell };
            exRow.시료량 = ws.Cell(r, colVol).GetString().Trim();
            exRow.D1   = ws.Cell(r, colD1).GetString().Trim();
            exRow.D2   = ws.Cell(r, colD2).GetString().Trim();
            exRow.Fxy  = ws.Cell(r, colFxy).GetString().Trim();
            exRow.P    = colP < colResult ? ws.Cell(r, colP).GetString().Trim() : "";
            rows.Add(exRow);
        }
    }

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

        // 1단계: 모든 아이콘을 검정(⚫)으로 초기화
        foreach (var ic in _rowIcons)
            ic.Text = "⚫";

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
                if (i < _rowIcons.Count) _rowIcons[i].Text = "⚪";
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
                        ? MatchStatus.미매칭 : MatchStatus.대기;
                }
            }

            // 아이콘 업데이트 (0.05초 딜레이)
            if (i < _rowIcons.Count)
            {
                var resultIcon = row.Status switch
                {
                    MatchStatus.입력가능 => "🟢",
                    MatchStatus.덮어쓰기 => "🟡",
                    MatchStatus.대기     => "⚪",
                    _                    => "🔴",
                };
                _rowIcons[i].Text = resultIcon;
            }

            await System.Threading.Tasks.Task.Delay(50);
        }

        _currentSamples = allSamples;
        ShowMessage($"✅ 검증 완료: {excelRows.Count}건 확인", false);
    }

    /// <summary>서브메뉴 "입력" — 검증된 데이터 일괄 DB 반영</summary>
    public void ImportData()
    {
        if (!_categoryExcelData.TryGetValue(_activeCategory, out var rows))
        { ShowMessage("먼저 파일을 첨부하세요.", true); return; }

        int imported = 0, skipped = 0, disabled = 0;
        foreach (var row in rows)
        {
            if (string.IsNullOrWhiteSpace(row.Result)) continue;
            if (!row.Enabled) { disabled++; continue; }

            try
            {
                switch (row.Source)
                {
                    case SourceType.폐수배출업소 when row.Matched != null:
                        ImportWasteSample(row);
                        imported++;
                        break;

                    case SourceType.수질분석센터 when row.MatchedAnalysis != null:
                        ImportAnalysisRequest(row);
                        imported++;
                        break;

                    case SourceType.처리시설 when row.MatchedFacilityName != null:
                        ImportFacilityResult(row);
                        imported++;
                        break;

                    default:
                        skipped++;
                        break;
                }
            }
            catch { skipped++; }
        }

        LoadVerifiedGrid();
        BuildStatsPanel();
        string msg = $"✅ {imported}건 입력" +
            (disabled > 0 ? $" / {disabled}건 체크해제" : "") +
            (skipped > 0 ? $" / {skipped}건 제외" : "");
        ShowMessage(msg, false);
    }

    private void ImportWasteSample(ExcelRow row)
    {
        var s = row.Matched!;
        if (_activeItems.Length > 0)
        {
            switch (_activeItems[0])
            {
                case "BOD": s.BOD = row.Result; break;
                case "TOC": s.TOC = row.Result; break;
                case "SS":  s.SS  = row.Result; break;
                case "T-N": s.TN  = row.Result; break;
                case "T-P": s.TP  = row.Result; break;
                case "N-Hexan": s.NHexan = row.Result; break;
                case "Phenols": s.Phenols = row.Result; break;
            }
        }
        WasteSampleService.UpdateValues(s.Id, s.BOD, s.TOC, s.SS, s.TN, s.TP, s.NHexan, s.Phenols);
        SaveRawData(row, s);
    }

    private void ImportAnalysisRequest(ExcelRow row)
    {
        var ar = row.MatchedAnalysis!;
        // 약칭으로 컬럼명 매핑 (BOD → 생물화학적 산소요구량 등)
        var shortNames = AnalysisRequestService.GetShortNames();
        foreach (var item in _activeItems)
        {
            // 약칭 → 전체 컬럼명 역매핑
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

        if (_activeItems.Length > 0)
        {
            switch (_activeItems[0])
            {
                case "BOD": target.BOD = row.Result; break;
                case "TOC": target.TOC = row.Result; break;
                case "SS":  target.SS  = row.Result; break;
                case "T-N": target.TN  = row.Result; break;
                case "T-P": target.TP  = row.Result; break;
            }
        }
        string user = ETA.Services.Common.CurrentUserManager.Instance.CurrentUserId ?? "";
        FacilityResultService.SaveRows(row.MatchedFacilityName!, docDate, new List<FacilityResultRow> { target }, user);
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
                LoadVerifiedGrid();
            else
                LoadSampleGrid(_selectedDate);
            BuildStatsPanel();
        }
    }

    /// <summary>데이터 그리드와 동일한 컬럼 구조의 문서 정보 행.
    /// 첫 4컬럼을 ColumnSpan으로 합쳐서 label 표시, vals는 col 4부터 배치.</summary>
    private Border BuildDocRowUnified(string colDefs, string label, string[] vals, string resultFg)
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(colDefs), MinHeight = 28,
            Background = AppRes("GridRowBg") };
        // 구분명: 첫 4컬럼 합침
        var labelTb = FsBase(new TextBlock
        {
            Text = label, FontFamily = Font, Foreground = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(6, 2),
        });
        Grid.SetColumn(labelTb, 0); Grid.SetColumnSpan(labelTb, 4);
        g.Children.Add(labelTb);
        // 값: col 4부터
        for (int c = 0; c < vals.Length; c++)
        {
            int colIdx = 4 + c;
            bool isResult = c == vals.Length - 2; // 결과값 컬럼 (뒤에서 두번째)
            var tb = FsBase(new TextBlock
            {
                Text = vals[c], FontFamily = Font,
                FontWeight = isResult ? FontWeight.Bold : FontWeight.Normal,
                Foreground = isResult ? AppRes(resultFg) : AppRes("AppFg"),
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
        EditPanelChanged?.Invoke(FsSM(new TextBlock
        {
            Text = msg, FontFamily = Font,
            Foreground = isError ? AppRes("ThemeFgDanger") : AppRes("ThemeFgSuccess"),
            Margin = new Thickness(12), TextWrapping = TextWrapping.Wrap,
        }));
    }
}
