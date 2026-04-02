using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE2;

public partial class ResultSubmitErpPage : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
    private static readonly FontFamily FontM =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // ── Show2 연결 ────────────────────────────────────────────────────────
    public Action<Control?>? Show2ContentRequest;

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
    private Button?                        _btnStop;
    private bool                           _stopRequested;
    private ProgressBar?                   _progressBar;
    private TextBlock?                     _progressText;

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
        Show2ContentRequest?.Invoke(BuildExcelTableControl());
        // 1행 자동 선택
        if (_excelRows.Count > 0 && _rowBorders.Count > 0)
            SelectExcelRow(_excelRows[0], _rowBorders[0]);
    }

    private Control BuildExcelTableControl()
    {
        _excelRows  = ErpUiAutoService.LoadAllExcelData();
        _rowBorders = [];

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
                FontFamily = Font, FontSize = 11,
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
            Background = Brush.Parse("#1a2a3a"),
            BorderBrush = Brush.Parse("#2a3a4a"),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(4, 5),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 0,
                Children =
                {
                    MakeCell("번호", 36, "#778899", FontWeight.SemiBold),
                    MakeCell("S/N",  72, "#778899", FontWeight.SemiBold),
                    MakeCell("시료명", 200, "#778899", FontWeight.SemiBold),
                    MakeCell("BOD",  44, "#778899", FontWeight.SemiBold),
                    MakeCell("TOC",  44, "#778899", FontWeight.SemiBold),
                    MakeCell("SS",   44, "#778899", FontWeight.SemiBold),
                    MakeCell("NH",   44, "#778899", FontWeight.SemiBold),
                    MakeCell("PN",   44, "#778899", FontWeight.SemiBold),
                    MakeCell("TN",   44, "#778899", FontWeight.SemiBold),
                    MakeCell("TP",   44, "#778899", FontWeight.SemiBold),
                },
            },
        };
    }

    private Border MakeTableRow(ErpUiAutoService.ExcelRow row)
    {
        var b = new Border
        {
            Background = Brush.Parse(row.번호 % 2 == 0 ? "#0e1820" : "#0b1520"),
            BorderBrush = Brush.Parse("#1a2a3a"),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Padding = new Thickness(4, 4),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 0,
                Children =
                {
                    MakeCell(row.번호.ToString(), 36, "#6688aa"),
                    MakeCell(row.SN,              72, "#aaccff"),
                    MakeCell(row.시료명,          200, "#99aabb"),
                    MakeCell(row.Values.Length > 0 ? row.Values[0] : "", 44, "#88cc99"),
                    MakeCell(row.Values.Length > 1 ? row.Values[1] : "", 44, "#88cc99"),
                    MakeCell(row.Values.Length > 2 ? row.Values[2] : "", 44, "#88cc99"),
                    MakeCell(row.Values.Length > 3 ? row.Values[3] : "", 44, "#88cc99"),
                    MakeCell(row.Values.Length > 4 ? row.Values[4] : "", 44, "#88cc99"),
                    MakeCell(row.Values.Length > 5 ? row.Values[5] : "", 44, "#88cc99"),
                    MakeCell(row.Values.Length > 6 ? row.Values[6] : "", 44, "#88cc99"),
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
            FontSize = 11,
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
        border.Background = Brush.Parse("#1a3a5a");

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
            BorderBrush     = Brush.Parse("#334466"),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(6),
            Padding         = new Thickness(14, 10),
            Child = new StackPanel { Spacing = 6, Children =
            {
                new TextBlock
                {
                    Text = "ERP iU — 채수분석 자동입력",
                    FontFamily = FontM, FontSize = 13, FontWeight = FontWeight.SemiBold,
                    Foreground = Brush.Parse("#aaccff"),
                },
                new TextBlock
                {
                    Text = "왼쪽 테이블에서 행 클릭(또는 1행 자동 선택) → 🚀 입력 실행 → 다음 행 자동 이동.",
                    FontFamily = Font, FontSize = 11,
                    Foreground = Brush.Parse("#778899"),
                    TextWrapping = TextWrapping.Wrap,
                },
                new StackPanel { Orientation = Orientation.Horizontal, Spacing = 16, Children =
                {
                    MakeInfoItem("대상 프로세스", "neoweb.exe"),
                    MakeInfoItem("입력 순서", "BOD·TOC·SS·NH·PN·TN·TP"),
                }},
            }},
        });

        // ── 권한 경고 배너 ────────────────────────────────────────────────
        bool isAdmin = ErpUiAutoService.IsAdmin();
        if (!isAdmin)
        {
            var btnElevate = new Button
            {
                Content = "🛡️  관리자로 재실행",
                FontFamily = FontM, FontSize = 11,
                Padding = new Thickness(12, 6),
                Background = Brush.Parse("#3a1a00"),
                Foreground = Brush.Parse("#ffaa44"),
                BorderBrush = Brush.Parse("#aa6600"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
            };
            btnElevate.Click += (_, _) =>
            {
                if (ErpUiAutoService.RestartAsAdmin())
                    Environment.Exit(0);
            };
            root.Children.Add(new Border
            {
                Background = Brush.Parse("#1e1000"),
                BorderBrush = Brush.Parse("#aa6600"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(5),
                Padding = new Thickness(12, 8),
                Child = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12,
                    VerticalAlignment = VerticalAlignment.Center, Children =
                {
                    new TextBlock
                    {
                        Text = "⚠️  관리자 권한 필요 (neoweb.exe가 관리자로 실행 중).",
                        FontFamily = Font, FontSize = 11,
                        Foreground = Brush.Parse("#ddaa44"),
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    btnElevate,
                }},
            });
        }

        // ── 상태 배지 + 버튼 행 ──────────────────────────────────────────
        _statusBadge = new TextBlock
        {
            Text = isAdmin ? "대기 (관리자)" : "대기",
            FontFamily = FontM, FontSize = 11,
            Foreground = Brush.Parse("#778899"),
            VerticalAlignment = VerticalAlignment.Center,
        };

        var btnProbe = new Button
        {
            Content = "🔍  프로브",
            FontFamily = FontM, FontSize = 12,
            Padding = new Thickness(12, 7),
            Background = Brush.Parse("#1a3a6a"),
            Foreground = Brush.Parse("#88aaee"),
            BorderBrush = Brush.Parse("#336699"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        btnProbe.Click += async (_, _) => await RunProbeAsync(btnProbe);

        var btnTest = new Button
        {
            Content = "🧪  테스트",
            FontFamily = FontM, FontSize = 12,
            Padding = new Thickness(12, 7),
            Background = Brush.Parse("#1a3a1a"),
            Foreground = Brush.Parse("#88ee88"),
            BorderBrush = Brush.Parse("#336633"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        btnTest.Click += async (_, _) => await RunTestInputAsync(btnTest);

        root.Children.Add(new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 8,
            Children = { btnProbe, btnTest, _statusBadge },
        });

        // ── 선택된 시료명 표시 ─────────────────────────────────────────────
        _sampleLabel = new TextBlock
        {
            Text = "로딩 중…",
            FontFamily = Font, FontSize = 11,
            Foreground = Brush.Parse("#4a6a8a"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        root.Children.Add(new Border
        {
            Background = Brush.Parse("#0e1820"),
            BorderBrush = Brush.Parse("#223344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(10, 6),
            Child = _sampleLabel,
        });

        // ── 입력 행 ───────────────────────────────────────────────────────
        _valuesBox = new TextBox
        {
            Text = "",
            FontFamily = Font, FontSize = 11,
            Width = 240,
            Padding = new Thickness(6, 5),
            Background = Brush.Parse("#0e1520"),
            Foreground = Brush.Parse("#88aacc"),
            BorderBrush = Brush.Parse("#334466"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Watermark = "BOD,TOC,SS,NH,PN,TN,TP",
        };

        _rowBox = new TextBox
        {
            Text = "1",
            FontFamily = Font, FontSize = 11,
            Width = 44,
            Padding = new Thickness(6, 5),
            Background = Brush.Parse("#0e1520"),
            Foreground = Brush.Parse("#88aacc"),
            BorderBrush = Brush.Parse("#334466"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            TextAlignment = Avalonia.Media.TextAlignment.Center,
        };

        _btnRun = new Button
        {
            Content = "🚀  입력 실행",
            FontFamily = FontM, FontSize = 12,
            Padding = new Thickness(16, 8),
            Background = Brush.Parse("#2a1a3a"),
            Foreground = Brush.Parse("#cc88ff"),
            BorderBrush = Brush.Parse("#664488"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
        };
        _btnRun.Click += async (_, _) => await RunWorkflowAsync(_btnRun!);

        _btnStop = new Button
        {
            Content = "⏹  중단",
            FontFamily = FontM, FontSize = 12,
            Padding = new Thickness(12, 7),
            Background = Brush.Parse("#3a1a1a"),
            Foreground = Brush.Parse("#ee8888"),
            BorderBrush = Brush.Parse("#663333"),
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
                    Text = "행:", FontFamily = FontM, FontSize = 11,
                    Foreground = Brush.Parse("#556677"),
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
            Text = "", FontFamily = Font, FontSize = 10,
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
            Background      = Brush.Parse("#0e1520"),
            BorderBrush     = Brush.Parse("#223344"),
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

    private async Task RunTestInputAsync(Button btn)
    {
        if (_logPanel == null || _statusBadge == null) return;
        btn.IsEnabled = false;
        _logPanel.Children.Clear();
        SetStatus("입력 중…", "#eedd66");
        AddLog("──", "테스트 입력  " + DateTime.Now.ToString("HH:mm:ss"));

        try
        {
            var lines = await Task.Run(() => ErpUiAutoService.TestInput("12345"));
            foreach (var (icon, msg) in lines) AddLog(icon, msg);
            SetStatus("✅ 완료", "#66ee88");
            AddLog("──", "Logs/ERP.log 확인");
        }
        catch (Exception ex) { AddLog("❌", ex.Message); SetStatus("오류", "#ee6666"); }

        btn.IsEnabled = true;
        Dispatcher.UIThread.Post(() => _logScroll?.ScrollToEnd(), DispatcherPriority.Background);
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
                FontFamily = Font, FontSize = 11,
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
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                FontSize = 9, Foreground = Brush.Parse("#445566") },
            new TextBlock { Text = value,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                FontSize = 10, Foreground = Brush.Parse("#88aacc") },
        }};
}
