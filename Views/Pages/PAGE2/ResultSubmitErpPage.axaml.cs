using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.PAGE2;

public partial class ResultSubmitErpPage : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly FontFamily FontM =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── Show2/Show3 연결 ─────────────────────────────────────────────────
    public Action<Control?>? Show2ContentRequest;
    public Action<Control?>? Show3ContentRequest;

    // ── 내부 상태 ─────────────────────────────────────────────────────────
    private TextBlock? _statusBadge;
    private StackPanel? _logPanel;
    private ScrollViewer? _logScroll;
    private TextBox?    _valuesBox;
    private TextBox?    _rowBox;
    private TextBlock?  _sampleLabel;

    private ErpUiAutoService.ExcelRow? _selectedRow;  // 현재 선택된 Excel 행
    private Border?                    _selectedBorder;

    private List<ErpUiAutoService.ExcelRow> _excelRows = [];
    private List<Border>                    _rowBorders = [];
    private Button?                        _btnRun;
    private DatePicker?                    _monthPicker;
    private RadioButton?                   _rbYeosu;
    private RadioButton?                   _rbYulchon;
    private RadioButton?                   _rbSepoong;
    private Button?                        _btnStop;
    private bool                           _stopRequested;
    private ProgressBar?                   _progressBar;
    private TextBlock?                     _progressText;
    private string?                        _loadedExcelPath;
    private HashSet<string>                _verifiedSns = [];   // 검증 완료된 S/N
    private List<TextBlock>                _checkIcons = [];    // Show2 행별 체크 아이콘

    public ResultSubmitErpPage()
    {
        InitializeComponent();
        var panel = this.FindControl<StackPanel>("ContentPanel");
        if (panel != null) BuildUi(panel);
    }

    // =========================================================================
    // Show2 — Excel 테이블 빌드 및 푸시
    // =========================================================================
    public void RefreshShow2()
    {
        Show2ContentRequest?.Invoke(WrapWithFilterBar(BuildExcelTableControl()));
        // 1행 자동 선택
        if (_excelRows.Count > 0 && _rowBorders.Count > 0)
            SelectExcelRow(_excelRows[0], _rowBorders[0]);
    }

    /// <summary>Show2 콘텐츠 위에 [여수][율촌][세풍] 라디오 바를 얹어 묶음.
    /// 라디오는 클래스 필드라 이전 부모에서 detach 후 새 부모에 attach.</summary>
    private Control WrapWithFilterBar(Control content)
    {
        // 이전 부모에서 분리
        DetachFromParent(_rbYeosu);
        DetachFromParent(_rbYulchon);
        DetachFromParent(_rbSepoong);

        var topBar = new Border
        {
            Background = Brush.Parse("#0e1820"),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(10, 6),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 16,
                VerticalAlignment = VerticalAlignment.Center,
                Children = { _rbYeosu!, _rbYulchon!, _rbSepoong! },
            },
        };
        var grid = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        Grid.SetRow(topBar, 0);
        grid.Children.Add(topBar);
        Grid.SetRow(content, 1);
        grid.Children.Add(content);
        return grid;
    }

    private static void DetachFromParent(Control? c)
    {
        if (c?.Parent is Panel p) p.Children.Remove(c);
    }

    private Control BuildExcelTableControl()
    {
        if (_loadedExcelPath == null)
        {
            _excelRows = [];
            _rowBorders = [];
            _checkIcons = [];
            return new TextBlock
            {
                Text = "📂  '분석결과 가져오기' 버튼으로 Excel 파일을 선택하세요.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = Brush.Parse("#88aabb"),
                Margin = new Thickness(16, 20),
            };
        }
        if (!System.IO.File.Exists(_loadedExcelPath))
        {
            _excelRows = []; _rowBorders = []; _checkIcons = [];
            _loadedExcelPath = null;
            return new TextBlock
            {
                Text = "⚠️  이전 Excel 파일을 찾을 수 없습니다 — 다시 가져와 주세요.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = Brush.Parse("#ee8866"),
                Margin = new Thickness(16, 20),
            };
        }

        _excelRows = ErpUiAutoService.LoadAllExcelData(_loadedExcelPath);
        _rowBorders = [];
        _checkIcons = [];

        // Excel 파일은 있으나 데이터 행이 0개
        if (_excelRows.Count == 0)
        {
            return new TextBlock
            {
                Text = "📭  Excel 파일에 데이터 행이 없습니다 (자료입력 시트 2행부터 비어있음).",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = Brush.Parse("#ddaa66"),
                Margin = new Thickness(16, 20),
            };
        }

        var sp = new StackPanel { Spacing = 0 };

        // 헤더
        sp.Children.Add(MakeTableHeader());

        // 데이터 행
        foreach (var row in _excelRows)
        {
            var rowBorder = MakeTableRow(row);
            _rowBorders.Add(rowBorder);
            sp.Children.Add(rowBorder);
        }

        if (_excelRows.Count == 0)
        {
            sp.Children.Add(new TextBlock
            {
                Text = "Excel 파일을 찾을 수 없습니다.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = Brush.Parse("#ee6666"),
                Margin = new Thickness(10, 8),
            });
        }

        return new Border
        {
            Background = Brush.Parse("#0d1520"),
            Child = new ScrollViewer
            {
                Content = sp,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            },
        };
    }

    private static Control MakeTableHeader()
    {
        return new Border
        {
            Background = AppTheme.BorderSubtle,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(4, 5),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 0,
                Children =
                {
                    MakeCell("",     24, "#778899", FontWeight.SemiBold),
                    MakeCell("번호", 36, "#778899", FontWeight.SemiBold),
                    MakeCell("S/N",  90, "#778899", FontWeight.SemiBold),
                    MakeCell("시료명", 220, "#778899", FontWeight.SemiBold),
                    MakeCell("BOD",  68, "#778899", FontWeight.SemiBold),
                    MakeCell("TOC",  68, "#778899", FontWeight.SemiBold),
                    MakeCell("SS",   68, "#778899", FontWeight.SemiBold),
                    MakeCell("NH",   68, "#778899", FontWeight.SemiBold),
                    MakeCell("PN",   68, "#778899", FontWeight.SemiBold),
                    MakeCell("TN",   68, "#778899", FontWeight.SemiBold),
                    MakeCell("TP",   68, "#778899", FontWeight.SemiBold),
                },
            },
        };
    }

    private Border MakeTableRow(ErpUiAutoService.ExcelRow row)
    {
        var checkIcon = new TextBlock
        {
            Text = _verifiedSns.Contains(row.SN) ? "✅" : "⬜",
            Width = 24,
            FontSize = AppTheme.FontBase,
            Foreground = _verifiedSns.Contains(row.SN) ? Brush.Parse("#66ee88") : Brush.Parse("#334455"),
            Margin = new Thickness(2, 0),
            TextAlignment = Avalonia.Media.TextAlignment.Center,
        };
        _checkIcons.Add(checkIcon);

        var b = new Border
        {
            Background = Brush.Parse(row.번호 % 2 == 0 ? "#0e1820" : "#0b1520"),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(4, 4),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 0,
                Children =
                {
                    checkIcon,
                    MakeCell(row.번호.ToString(), 36, "#6688aa"),
                    MakeCell(row.SN,              90, "#aaccff"),
                    MakeCell(row.시료명,          220, "#99aabb"),
                    MakeCell(row.Values.Length > 0 ? row.Values[0] : "", 68, "#88cc99"),
                    MakeCell(row.Values.Length > 1 ? row.Values[1] : "", 68, "#88cc99"),
                    MakeCell(row.Values.Length > 2 ? row.Values[2] : "", 68, "#88cc99"),
                    MakeCell(row.Values.Length > 3 ? row.Values[3] : "", 68, "#88cc99"),
                    MakeCell(row.Values.Length > 4 ? row.Values[4] : "", 68, "#88cc99"),
                    MakeCell(row.Values.Length > 5 ? row.Values[5] : "", 68, "#88cc99"),
                    MakeCell(row.Values.Length > 6 ? row.Values[6] : "", 68, "#88cc99"),
                },
            },
        };

        b.PointerPressed += (_, _) => SelectExcelRow(row, b);
        return b;
    }

    private static TextBlock MakeCell(string text, double width, string color,
        FontWeight weight = FontWeight.Normal)
    {
        return new TextBlock
        {
            Text = text,
            Width = width,
            FontFamily = Font,
            FontSize = AppTheme.FontBase,
            FontWeight = weight,
            Foreground = Brush.Parse(color),
            Margin = new Thickness(2, 0),
            TextTrimming = TextTrimming.CharacterEllipsis,
        };
    }

    private bool AdvanceToNextRow(int currentRow)
    {
        int nextIdx = currentRow; // 0-based index in _excelRows (row 1 = index 0)
        if (nextIdx >= _excelRows.Count) return false;

        var nextRow    = _excelRows[nextIdx];
        var nextBorder = nextIdx < _rowBorders.Count ? _rowBorders[nextIdx] : null;
        if (nextBorder == null) return false;

        SelectExcelRow(nextRow, nextBorder);
        return true;
    }

    private void SelectExcelRow(ErpUiAutoService.ExcelRow row, Border border)
    {
        // 이전 선택 해제
        if (_selectedBorder != null)
            _selectedBorder.Background = Brush.Parse(_selectedBorder.Tag is string tag ? tag : "#0e1820");

        // 선택 강조
        _selectedBorder = border;
        _selectedRow    = row;
        border.Background = AppTheme.BgActiveBlue;

        // 오른쪽 패널 채우기
        Dispatcher.UIThread.Post(() =>
        {
            if (_rowBox    != null) _rowBox.Text    = row.번호.ToString();
            if (_valuesBox != null) _valuesBox.Text = string.Join(",", row.Values);
            if (_sampleLabel != null)
                _sampleLabel.Text = $"[{row.번호}] {row.SN}  {row.시료명}";
        });
    }

    // =========================================================================
    private void BuildUi(StackPanel root)
    {
        // ── 안내 카드 ──────────────────────────────────────────────────────
        root.Children.Add(new Border
        {
            Background      = Brush.Parse("#1a2030"),
            BorderBrush     = AppTheme.BorderAccent,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(6),
            Padding         = new Thickness(14, 10),
            Child = new StackPanel { Spacing = 6, Children =
            {
                new TextBlock
                {
                    Text = "ERP iU — 채수분석 자동입력",
                    FontFamily = FontM, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
                    Foreground = AppTheme.FgInfo,
                },
                new TextBlock
                {
                    Text = "왼쪽 테이블에서 행 클릭(또는 1행 자동 선택) → 🚀 입력 실행 → 다음 행 자동 이동.",
                    FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = AppTheme.FgMuted,
                    TextWrapping = TextWrapping.Wrap,
                },
                new WrapPanel { Orientation = Orientation.Horizontal, ItemSpacing = 16, LineSpacing = 4, Children =
                {
                    MakeInfoItem("대상 프로세스", "neoweb.exe"),
                    MakeInfoItem("입력 순서", "BOD·TOC·SS·NH·PN·TN·TP"),
                }},
            }},
        });

        // ── 상태 배지 + 버튼 행 ──────────────────────────────────────────
        // ETA 는 항상 관리자 권한으로 실행되므로 "관리자로 재실행" 배너/버튼 제거
        _statusBadge = new TextBlock
        {
            Text = "대기",
            FontFamily = FontM, FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var btnLoad = new Button
        {
            Content = "📂  분석결과 가져오기",
            FontFamily = FontM, FontSize = AppTheme.FontMD,
            Padding = new Thickness(12, 7),
            Background = Brush.Parse("#1a3a2a"),
            Foreground = Brush.Parse("#66ddaa"),
            BorderBrush = Brush.Parse("#336644"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        btnLoad.Click += async (_, _) => await LoadExcelFileAsync(btnLoad);

        // DB 월별 불러오기 — 캘린더로 월 선택 (일자는 무시) + 지역 라디오
        var now = DateTime.Today;
        _monthPicker = new DatePicker
        {
            SelectedDate     = new DateTimeOffset(now.Year, now.Month, 1, 0, 0, 0, TimeSpan.Zero),
            DayVisible       = false,
            FontFamily       = FontM,
            FontSize         = AppTheme.FontMD,
            VerticalAlignment = VerticalAlignment.Center,
        };
        _rbYeosu   = new RadioButton { Content = "여수", GroupName = "ErpRegion", IsChecked = true,
            FontFamily = FontM, FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center };
        _rbYulchon = new RadioButton { Content = "율촌", GroupName = "ErpRegion",
            FontFamily = FontM, FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center };
        _rbSepoong = new RadioButton { Content = "세풍", GroupName = "ErpRegion",
            FontFamily = FontM, FontSize = AppTheme.FontMD, VerticalAlignment = VerticalAlignment.Center };

        // 캘린더 선택 시 자동 로드
        _monthPicker.SelectedDateChanged += (_, _) => TriggerDbLoad();
        // 라디오 변경 시 SENDER 기준으로 지역 결정 (그룹 IsChecked 동시 변경 race 회피)
        _rbYeosu.IsCheckedChanged   += (_, _) => { if (_rbYeosu.IsChecked   == true) TriggerDbLoad("여수"); };
        _rbYulchon.IsCheckedChanged += (_, _) => { if (_rbYulchon.IsChecked == true) TriggerDbLoad("율촌"); };
        _rbSepoong.IsCheckedChanged += (_, _) => { if (_rbSepoong.IsChecked == true) TriggerDbLoad("세풍"); };

        var btnProbe = new Button
        {
            Content = "🔍  프로브",
            FontFamily = FontM, FontSize = AppTheme.FontMD,
            Padding = new Thickness(12, 7),
            Background = Brush.Parse("#1a3a6a"),
            Foreground = AppTheme.FgInfo,
            BorderBrush = AppTheme.BorderInfo,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        btnProbe.Click += async (_, _) => await RunProbeAsync(btnProbe);

        var btnVerify = new Button
        {
            Content = "🔎  자료검증",
            FontFamily = FontM, FontSize = AppTheme.FontMD,
            Padding = new Thickness(12, 7),
            Background = Brush.Parse("#2a2a1a"),
            Foreground = Brush.Parse("#ddcc66"),
            BorderBrush = Brush.Parse("#665533"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        btnVerify.Click += async (_, _) => await VerifyDataAsync(btnVerify);

        // 좁은 Show1 폭 대응 — 모든 요소를 Stretch 로 깔고 WrapPanel 로 자동 줄바꿈
        btnLoad.HorizontalAlignment   = HorizontalAlignment.Stretch;
        btnVerify.HorizontalAlignment = HorizontalAlignment.Stretch;
        btnProbe.HorizontalAlignment  = HorizontalAlignment.Stretch;

        var rowLoaders = new WrapPanel { Orientation = Orientation.Horizontal, ItemSpacing = 6, LineSpacing = 4 };
        rowLoaders.Children.Add(btnLoad);

        // 캘린더 별도 행 (전체 폭)
        _monthPicker.HorizontalAlignment = HorizontalAlignment.Stretch;

        // 라디오는 Show2 상단으로 이동 (BuildExcelTableFromCurrentRows / BuildExcelTableControl 에서 합침)

        // 검증/프로브 행
        var rowActions = new WrapPanel { Orientation = Orientation.Horizontal, ItemSpacing = 6, LineSpacing = 4, Margin = new Thickness(0, 6, 0, 0) };
        rowActions.Children.Add(btnVerify);
        rowActions.Children.Add(btnProbe);

        root.Children.Add(new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing = 4,
            Children =
            {
                rowLoaders,
                _monthPicker,
                rowActions,
                _statusBadge,
            },
        });

        // ── 선택된 시료명 표시 ─────────────────────────────────────────────
        _sampleLabel = new TextBlock
        {
            Text = "로딩 중…",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = Brush.Parse("#4a6a8a"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        root.Children.Add(new Border
        {
            Background = Brush.Parse("#0e1820"),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(10, 6),
            Child = _sampleLabel,
        });

        // ── 입력 행 ───────────────────────────────────────────────────────
        _valuesBox = new TextBox
        {
            Text = "",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Width = 240,
            Padding = new Thickness(6, 5),
            Background = AppTheme.BgPrimary,
            Foreground = AppTheme.FgInfo,
            BorderBrush = AppTheme.BorderAccent,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Watermark = "BOD,TOC,SS,NH,PN,TN,TP",
        };

        _rowBox = new TextBox
        {
            Text = "1",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Width = 44,
            Padding = new Thickness(6, 5),
            Background = AppTheme.BgPrimary,
            Foreground = AppTheme.FgInfo,
            BorderBrush = AppTheme.BorderAccent,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            TextAlignment = Avalonia.Media.TextAlignment.Center,
        };

        _btnRun = new Button
        {
            Content = "🚀  입력 실행",
            IsEnabled = false,
            FontFamily = FontM, FontSize = AppTheme.FontMD,
            Padding = new Thickness(16, 8),
            Background = Brush.Parse("#2a1a3a"),
            Foreground = Brush.Parse("#cc88ff"),
            BorderBrush = AppTheme.BorderWarn,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        _btnRun.Click += async (_, _) => await RunWorkflowAsync(_btnRun!);

        _btnStop = new Button
        {
            Content = "⏹  중단",
            FontFamily = FontM, FontSize = AppTheme.FontMD,
            Padding = new Thickness(12, 7),
            Background = AppTheme.BgDanger,
            Foreground = AppTheme.FgDanger,
            BorderBrush = AppTheme.BorderDanger,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
            IsEnabled = false,
        };
        _btnStop.Click += (_, _) => _stopRequested = true;

        root.Children.Add(new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            Children =
            {
                new TextBlock
                {
                    Text = "행:", FontFamily = FontM, FontSize = AppTheme.FontBase,
                    Foreground = AppTheme.FgDimmed,
                    VerticalAlignment = VerticalAlignment.Center,
                },
                _rowBox,
                _valuesBox,
                _btnRun,
                _btnStop,
            },
        });

        // ── 진행률 ────────────────────────────────────────────────────────
        _progressBar = new ProgressBar
        {
            Minimum = 0, Maximum = 100, Value = 0,
            Height = 6, CornerRadius = new CornerRadius(3),
            Background = Brush.Parse("#1a2530"),
            Foreground = Brush.Parse("#66aaff"),
        };
        _progressText = new TextBlock
        {
            Text = "", FontFamily = Font, FontSize = AppTheme.FontSM,
            Foreground = Brush.Parse("#667788"),
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        root.Children.Add(new StackPanel
        {
            Spacing = 2, Margin = new Thickness(0, 4, 0, 0),
            Children = { _progressBar, _progressText },
        });

        // ── 로그 패널 ─────────────────────────────────────────────────────
        _logPanel = new StackPanel { Spacing = 2 };
        _logScroll = new ScrollViewer
        {
            Content = _logPanel,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Height = 280,
        };

        root.Children.Add(new Border
        {
            Background      = AppTheme.BgPrimary,
            BorderBrush     = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(5),
            Padding         = new Thickness(10, 8),
            Child           = _logScroll,
        });
    }

    // =========================================================================
    private async Task RunProbeAsync(Button btn)
    {
        if (_logPanel == null || _statusBadge == null) return;
        btn.IsEnabled = false;
        _logPanel.Children.Clear();
        SetStatus("실행 중…", "#eedd66");
        AddLog("──", "프로브 시작  " + DateTime.Now.ToString("HH:mm:ss"));

        try
        {
            var result = await Task.Run(() => ErpUiAutoService.Probe());
            foreach (var (icon, msg) in result.Lines) AddLog(icon, msg);
            SetStatus(result.CanAutomate ? "✅ 자동화 가능" : "⚠️ 확인 필요",
                      result.CanAutomate ? "#66ee88" : "#ee9944");
        }
        catch (Exception ex) { AddLog("❌", ex.Message); SetStatus("오류", "#ee6666"); }

        btn.IsEnabled = true;
        Dispatcher.UIThread.Post(() => _logScroll?.ScrollToEnd(), DispatcherPriority.Background);
    }

    // =========================================================================
    private async Task RunWorkflowAsync(Button btn)
    {
        if (_logPanel == null || _statusBadge == null || _selectedRow == null)
        {
            AddLog("⚠️", "왼쪽 테이블에서 행을 먼저 선택하세요.");
            return;
        }

        _stopRequested = false;
        btn.IsEnabled = false;
        if (_btnStop != null) _btnStop.IsEnabled = true;
        _logPanel.Children.Clear();
        SetStatus("실행 중…", "#eedd66");
        ErpUiAutoService.ResetCache();

        // 진행률 초기화
        int startIdx = _excelRows.IndexOf(_selectedRow);
        int totalRows = _excelRows.Count - startIdx;
        int doneCount = 0;
        UpdateProgress(0, totalRows);

        try
        {
            bool anyError = false;
            bool isFirst = true;

            while (_selectedRow != null && !_stopRequested)
            {
                var row    = _selectedRow;
                int rowIdx = row.번호;

                AddLog("──", $"[{doneCount+1}/{totalRows}]  row={rowIdx}  S/N={row.SN}");

                var lines = await Task.Run(() =>
                    ErpUiAutoService.RunWorkflow(row.Values, rowIdx, row.SN, navigateFromStart: isFirst));
                isFirst = false;
                foreach (var (icon, msg) in lines) AddLog(icon, msg);

                bool ok = lines.Exists(l => l.Icon == "✅");
                if (!ok) { anyError = true; break; }

                doneCount++;
                UpdateProgress(doneCount, totalRows);

                if (!AdvanceToNextRow(rowIdx)) break;
            }

            if (_stopRequested)
                SetStatus("⏹ 중단됨", "#ddaa44");
            else if (anyError)
                SetStatus("⚠️ 확인 필요", "#ee9944");
            else
                SetStatus("✅ 전체 완료", "#66ee88");
        }
        catch (Exception ex) { AddLog("❌", ex.Message); SetStatus("오류", "#ee6666"); }

        btn.IsEnabled = true;
        if (_btnStop != null) _btnStop.IsEnabled = false;
        Dispatcher.UIThread.Post(() => _logScroll?.ScrollToEnd(), DispatcherPriority.Background);
    }



    // =========================================================================
    // DB 월별 불러오기 — 비용부담금_결과 에서 해당 월 전체 의뢰일자×업체 행 로드
    // =========================================================================
    private void TriggerDbLoad(string? regionOverride = null)
    {
        if (_monthPicker == null) return;
        try
        {
            var picked = _monthPicker?.SelectedDate?.DateTime ?? DateTime.Today;
            int year  = picked.Year;
            int month = picked.Month;
            string region = regionOverride ?? (
                _rbYulchon?.IsChecked == true ? "율촌" :
                _rbSepoong?.IsChecked == true ? "세풍" :
                                                "여수");

            // 외부 Excel 경로 끊고 DB 직접 사용
            _loadedExcelPath = null;
            _excelRows = ErpUiAutoService.LoadFromDb(year, month, region);

            // 기존 RefreshShow2 는 _loadedExcelPath 의존이라, 직접 컨트롤 빌드 후 푸시
            Show2ContentRequest?.Invoke(WrapWithFilterBar(BuildExcelTableFromCurrentRows()));

            if (_excelRows.Count > 0 && _rowBorders.Count > 0)
                SelectExcelRow(_excelRows[0], _rowBorders[0]);

            AddLog("✅", $"DB 로드 {year:D4}-{month:D2} [{region}]: {_excelRows.Count}건");
            SetStatus($"📅 {year:D4}-{month:D2} {region}  {_excelRows.Count}건 로드", "#88ccff");
            if (_btnRun != null) _btnRun.IsEnabled = _excelRows.Count > 0;
        }
        catch (Exception ex)
        {
            AddLog("❌", $"DB 로드 실패: {ex.Message}");
            SetStatus("오류", "#ee6666");
        }
    }

    /// <summary>이미 로드된 _excelRows 로 테이블만 빌드 (DB 로드 후 사용)</summary>
    private Control BuildExcelTableFromCurrentRows()
    {
        _rowBorders = [];
        _checkIcons = [];

        if (_excelRows.Count == 0)
            return new TextBlock
            {
                Text = "📭  해당 월에 분석결과가 없습니다.",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = Brush.Parse("#ddaa66"),
                Margin = new Thickness(16, 20),
            };

        var sp = new StackPanel { Spacing = 0 };
        sp.Children.Add(MakeTableHeader());
        foreach (var row in _excelRows)
        {
            var bdr = MakeTableRow(row);
            sp.Children.Add(bdr);
            _rowBorders.Add(bdr);
        }
        return new Border
        {
            Background = Brush.Parse("#0d1520"),
            Child = new ScrollViewer
            {
                Content = sp,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            },
        };
    }

    // =========================================================================
    // 분석결과 가져오기 — Excel 파일 선택 → Show2에 로드
    // =========================================================================
    private async Task LoadExcelFileAsync(Button btn)
    {
        btn.IsEnabled = false;
        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null) { btn.IsEnabled = true; return; }

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "분석결과 Excel 파일 선택",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("Excel 파일") { Patterns = new[] { "*.xlsx", "*.xls" } },
                },
            });

            if (files.Count == 0) { btn.IsEnabled = true; return; }

            var path = files[0].TryGetLocalPath();
            if (string.IsNullOrEmpty(path)) { btn.IsEnabled = true; return; }

            _loadedExcelPath = path;
            AddLog("✅", $"파일 로드: {System.IO.Path.GetFileName(path)}");
            RefreshShow2();
            SetStatus($"📂 {_excelRows.Count}건 로드", "#66ddaa");
            if (_btnRun != null) _btnRun.IsEnabled = _excelRows.Count > 0;
        }
        catch (Exception ex)
        {
            AddLog("❌", $"파일 로드 실패: {ex.Message}");
            SetStatus("오류", "#ee6666");
        }
        btn.IsEnabled = true;
    }

    // =========================================================================
    // 자료검증 — ERP S/N과 Excel S/N 비교
    // =========================================================================
    private async Task VerifyDataAsync(Button btn)
    {
        if (_excelRows.Count == 0)
        {
            AddLog("⚠️", "먼저 분석결과를 가져오세요.");
            return;
        }
        btn.IsEnabled = false;
        _logPanel?.Children.Clear();
        SetStatus("검증 중…", "#eedd66");
        AddLog("──", $"자료검증 시작  {DateTime.Now:HH:mm:ss}  ({_excelRows.Count}건)");

        try
        {
            // ERP에서 현재 그리드의 S/N 목록 가져오기 (실패 시 Excel 자체 검증으로 진행)
            var erpSnList = await Task.Run(() => ErpUiAutoService.GetGridSnList());

            if (erpSnList.Count == 0)
            {
                AddLog("⚠️", "ERP 그리드 S/N 읽기 불가 — Excel 자체 검증으로 진행합니다.");
                // Excel S/N끼리만 중복 체크
                var excelSnAll = _excelRows.Select(r => r.SN).ToList();
                var duplicates = excelSnAll.GroupBy(s => s).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
                if (duplicates.Count > 0)
                    AddLog("⚠️", $"중복 S/N 발견: {string.Join(", ", duplicates)}");
                else
                    AddLog("✅", $"Excel S/N {excelSnAll.Count}건, 중복 없음");
                SetStatus("✅ Excel 검증 완료", "#66ee88");
                btn.IsEnabled = true;
                return;
            }

            AddLog("✅", $"ERP S/N {erpSnList.Count}건 읽기 완료");
            AddLog("──", "1:1 검증 시작…");

            var erpSnSet = new HashSet<string>(erpSnList);
            var excelSnSet = new HashSet<string>(_excelRows.Select(r => r.SN));
            int matchCount = 0, missingCount = 0, extraCount = 0;
            var missingSns = new List<string>();
            int totalItems = _excelRows.Count;

            // ── Show3 리뷰 패널 먼저 표시 (실시간 업데이트용 참조 확보) ──
            var liveRateText = new TextBlock
            {
                Text = "0.0%",
                FontFamily = FontM, FontSize = 32, FontWeight = FontWeight.Bold,
                Foreground = Brush.Parse("#ee5555"),
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            var liveBarFill = new Border
            {
                Background = new SolidColorBrush(Color.Parse("#ee5555")),
                CornerRadius = new CornerRadius(5),
                Height = 10,
                HorizontalAlignment = HorizontalAlignment.Left,
                Width = 0,
            };
            var liveBarContainer = new Grid
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Margin = new Thickness(0, 6, 0, 0),
                Children =
                {
                    new Border
                    {
                        Background = Brush.Parse("#1a2535"),
                        CornerRadius = new CornerRadius(5),
                        Height = 10,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                    },
                    liveBarFill,
                },
            };
            var liveCountText = new TextBlock
            {
                Text = $"0/{totalItems}",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = AppTheme.FgMuted,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            var liveDetailPanel = new StackPanel { Spacing = 6 };
            var liveMissingPanel = new StackPanel { Spacing = 4 };

            ShowLiveReviewPanel(erpSnList.Count, _excelRows.Count,
                liveRateText, liveBarFill, liveBarContainer, liveCountText,
                liveDetailPanel, liveMissingPanel);

            // ── 1:1 순차 검증 (Show1 + Show2 + Show3 동시 진행) ──
            int perItemMs = totalItems > 0 ? Math.Max(2500 / totalItems, 20) : 50;

            for (int i = 0; i < _excelRows.Count; i++)
            {
                if (_stopRequested) break;   // 중단 즉시 반응

                var row = _excelRows[i];
                bool ok = erpSnSet.Contains(row.SN);

                if (ok) matchCount++;
                else extraCount++;

                // Show1: 로그
                string icon = ok ? "✅" : "❌";
                string msg = ok
                    ? $"[{row.번호}] {row.SN}  {row.시료명} — 매칭"
                    : $"[{row.번호}] {row.SN}  {row.시료명} — ERP에 없음";
                AddLog(icon, msg);

                // Show2: 체크 아이콘
                int idx = i;
                int curMatch = matchCount;
                int curTotal = i + 1;
                // 매칭률 = 검증한 Excel 행 중 매칭된 비율 (사용자 직관에 맞춤)
                double curRate = curTotal > 0 ? (double)curMatch / curTotal * 100 : 0;

                Dispatcher.UIThread.Post(() =>
                {
                    // Show2 체크
                    if (idx < _checkIcons.Count)
                    {
                        _checkIcons[idx].Text = ok ? "✅" : "❌";
                        _checkIcons[idx].Foreground = Brush.Parse(ok ? "#66ee88" : "#ee4444");
                    }
                    _logScroll?.ScrollToEnd();

                    // Show3 실시간 업데이트
                    var phaseColor = LerpRateColor(curRate);
                    liveRateText.Text = $"{curRate:F1}%";
                    liveRateText.Foreground = new SolidColorBrush(phaseColor);
                    liveBarFill.Background = new SolidColorBrush(phaseColor);
                    liveCountText.Text = $"{curTotal}/{totalItems}";

                    double containerW = liveBarContainer.Bounds.Width;
                    if (containerW > 0)
                        liveBarFill.Width = containerW * (curRate / 100.0);
                });

                SetStatus($"검증 중… {i + 1}/{totalItems}", "#eedd66");
                await Task.Delay(perItemMs);
            }

            _verifiedSns = erpSnSet;

            // ERP S/N 중 Excel에 없는 것
            foreach (var erpSn in erpSnList)
            {
                if (!excelSnSet.Contains(erpSn))
                {
                    missingCount++;
                    missingSns.Add(erpSn);
                    AddLog("❌", $"ERP S/N '{erpSn}' → Excel에 없음");
                }
            }

            AddLog("──", $"매칭 {matchCount} / ERP누락 {missingCount} / Excel초과 {extraCount}");

            // Show3 최종 상태 반영
            Dispatcher.UIThread.Post(() =>
            {
                // 상세 항목
                liveDetailPanel.Children.Clear();
                liveDetailPanel.Children.Add(MakeReviewRow("매칭", matchCount, "#66ee88"));
                if (missingCount > 0)
                    liveDetailPanel.Children.Add(MakeReviewRow("ERP에만 존재 (Excel 누락)", missingCount, "#ee6666"));
                if (extraCount > 0)
                    liveDetailPanel.Children.Add(MakeReviewRow("Excel에만 존재 (ERP 누락)", extraCount, "#ddaa44"));

                // 누락 목록
                if (missingSns.Count > 0)
                {
                    liveMissingPanel.Children.Add(new TextBlock
                    {
                        Text = "누락된 의뢰 목록",
                        FontFamily = FontM, FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
                        Foreground = Brush.Parse("#ee6666"),
                        Margin = new Thickness(0, 0, 0, 4),
                    });
                    foreach (var sn in missingSns)
                    {
                        liveMissingPanel.Children.Add(new TextBlock
                        {
                            Text = $"  ERP {sn} 의 의뢰에 대한 분석결과 자료가 없습니다.",
                            FontFamily = Font, FontSize = AppTheme.FontBase,
                            Foreground = Brush.Parse("#cc6666"),
                            TextWrapping = TextWrapping.Wrap,
                        });
                    }
                }
            });

            if (missingCount == 0 && extraCount == 0)
            {
                SetStatus("✅ 검증 완료", "#66ee88");
                Dispatcher.UIThread.Post(() => { if (_btnRun != null) _btnRun.IsEnabled = true; });
            }
            else
            {
                SetStatus($"⚠️ 누락 {missingCount} / 초과 {extraCount}", "#ee9944");
                Dispatcher.UIThread.Post(() => { if (_btnRun != null) _btnRun.IsEnabled = false; });
            }
        }
        catch (Exception ex)
        {
            AddLog("❌", $"검증 오류: {ex.Message}");
            SetStatus("오류", "#ee6666");
        }

        btn.IsEnabled = true;
        Dispatcher.UIThread.Post(() => _logScroll?.ScrollToEnd(), DispatcherPriority.Background);
    }

    // =========================================================================
    // Show3 — 검증 리뷰 패널
    // =========================================================================
    // =========================================================================
    // Show3 — 실시간 리뷰 패널 (검증 시작과 동시에 표시)
    // =========================================================================
    private void ShowLiveReviewPanel(int erpCount, int excelCount,
        TextBlock rateText, Border barFill, Grid barContainer,
        TextBlock countText, StackPanel detailPanel, StackPanel missingPanel)
    {
        var panel = new StackPanel { Spacing = 12, Margin = new Thickness(16, 14) };

        // 제목
        panel.Children.Add(new TextBlock
        {
            Text = "🔎  자료검증 진행 중",
            FontFamily = FontM, FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#eedd66"),
        });

        // 요약 카드
        var summaryGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,*") };
        summaryGrid.Children.Add(MakeReviewCard("ERP 리스트", $"{erpCount}건", "#5588ff", 0));
        summaryGrid.Children.Add(MakeReviewCard("첨부 분석결과", $"{excelCount}건", "#66ddaa", 1));
        panel.Children.Add(summaryGrid);

        // 매칭률 + 프로그레스 바
        panel.Children.Add(new Border
        {
            Background = Brush.Parse("#0e1825"),
            BorderBrush = Brush.Parse("#334455"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(20, 16),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Child = new StackPanel { Spacing = 4, Children =
            {
                new StackPanel { HorizontalAlignment = HorizontalAlignment.Center, Spacing = 2, Children =
                {
                    rateText,
                    new TextBlock
                    {
                        Text = "매칭률",
                        FontFamily = Font, FontSize = AppTheme.FontBase,
                        Foreground = AppTheme.FgMuted,
                        HorizontalAlignment = HorizontalAlignment.Center,
                    },
                }},
                barContainer,
                countText,
            }},
        });

        // 상세 항목 (검증 완료 후 채워짐)
        panel.Children.Add(new Border
        {
            Background = Brush.Parse("#0e1825"),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(14, 10),
            Child = detailPanel,
        });

        // 누락 목록 (검증 완료 후 채워짐)
        panel.Children.Add(new Border
        {
            Background = Brush.Parse("#1a1015"),
            BorderBrush = Brush.Parse("#442222"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(12, 10),
            Child = missingPanel,
        });

        Show3ContentRequest?.Invoke(new Border
        {
            Background = Brush.Parse("#0d1520"),
            Child = new ScrollViewer
            {
                Content = panel,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            },
        });
    }

    private static Border MakeReviewCard(string label, string value, string color, int col)
    {
        var card = new Border
        {
            Background = Brush.Parse("#0e1825"),
            BorderBrush = Brush.Parse(color),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(14, 10),
            Margin = new Thickness(col == 0 ? 0 : 4, 0, col == 0 ? 4 : 0, 0),
            Child = new StackPanel { Spacing = 2, Children =
            {
                new TextBlock
                {
                    Text = label,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    FontSize = AppTheme.FontSM,
                    Foreground = AppTheme.FgDimmed,
                },
                new TextBlock
                {
                    Text = value,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    FontSize = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
                    Foreground = Brush.Parse(color),
                },
            }},
        };
        Grid.SetColumn(card, col);
        return card;
    }

    /// <summary>0~100% 매칭률에 따라 빨강→주황→노랑→초록→파랑 부드러운 보간</summary>
    private static Color LerpRateColor(double rate)
    {
        rate = Math.Clamp(rate, 0, 100);
        // 10단계 키프레임: (%, R, G, B)
        (double p, byte r, byte g, byte b)[] keys =
        [
            (0,   0xEE, 0x44, 0x44), // 빨강
            (12,  0xEE, 0x66, 0x33), // 진주황
            (25,  0xEE, 0x88, 0x33), // 주황
            (37,  0xDD, 0xAA, 0x33), // 황갈
            (50,  0xCC, 0xCC, 0x44), // 노랑
            (62,  0x88, 0xDD, 0x55), // 연두
            (75,  0x55, 0xDD, 0x77), // 초록
            (87,  0x55, 0xCC, 0x99), // 청록
            (95,  0x55, 0xBB, 0xDD), // 하늘
            (100, 0x55, 0x99, 0xFF), // 파랑
        ];

        // 구간 찾�� + 선형 보간
        for (int i = 0; i < keys.Length - 1; i++)
        {
            if (rate <= keys[i + 1].p)
            {
                double t = (rate - keys[i].p) / (keys[i + 1].p - keys[i].p);
                byte cr = (byte)(keys[i].r + (keys[i + 1].r - keys[i].r) * t);
                byte cg = (byte)(keys[i].g + (keys[i + 1].g - keys[i].g) * t);
                byte cb = (byte)(keys[i].b + (keys[i + 1].b - keys[i].b) * t);
                return Color.FromRgb(cr, cg, cb);
            }
        }
        return Color.FromRgb(keys[^1].r, keys[^1].g, keys[^1].b);
    }

    private static Control MakeReviewRow(string label, int count, string color)
    {
        return new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 8, Children =
            {
                new TextBlock
                {
                    Text = "●", FontSize = 8,
                    Foreground = Brush.Parse(color),
                    VerticalAlignment = VerticalAlignment.Center,
                },
                new TextBlock
                {
                    Text = label,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    FontSize = AppTheme.FontBase,
                    Foreground = AppTheme.FgMuted,
                },
                new TextBlock
                {
                    Text = $"{count}건",
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                    Foreground = Brush.Parse(color),
                },
            },
        };
    }

    private void AddLog(string icon, string msg)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_logPanel == null) return;
            var color = icon switch
            {
                "✅" => "#66cc88", "❌" => "#ee6666",
                "⚠️" => "#ddaa44", "🟢" => "#44ee88",
                "🔴" => "#ee4444", _ => "#556677",
            };
            _logPanel.Children.Add(new TextBlock
            {
                Text = $"{icon}  {msg}",
                FontFamily = Font, FontSize = AppTheme.FontBase,
                Foreground = Brush.Parse(color),
            });
        });
    }

    private void SetStatus(string text, string color) =>
        Dispatcher.UIThread.Post(() =>
        {
            if (_statusBadge == null) return;
            _statusBadge.Text = text;
            _statusBadge.Foreground = Brush.Parse(color);
        });

    private void UpdateProgress(int done, int total) =>
        Dispatcher.UIThread.Post(() =>
        {
            if (_progressBar == null) return;
            double pct = total > 0 ? (double)done / total * 100 : 0;
            _progressBar.Value = pct;
            if (_progressText != null)
                _progressText.Text = total > 0 ? $"{done}/{total}  ({pct:F0}%)" : "";
        });

    // =========================================================================
    private static Control MakeInfoItem(string label, string value) =>
        new StackPanel { Spacing = 2, Children =
        {
            new TextBlock { Text = label,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                FontSize = AppTheme.FontXS, Foreground = AppTheme.FgDimmed },
            new TextBlock { Text = value,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                FontSize = AppTheme.FontSM, Foreground = AppTheme.FgInfo },
        }};
}
