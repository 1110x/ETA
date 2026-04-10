using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE5;

public partial class ParserGeneratorPage : UserControl
{
    // ── 외부(MainPage) 연결 ────────────────────────────────────────────
    public event Action<Control?>? ListPanelChanged;   // Show2
    public event Action<Control?>? DetailPanelChanged; // Show3
    public event Action<Control?>? StatsPanelChanged;  // Show4

    private readonly List<ParseGenResult> _results = new();
    private ParseGenResult? _currentResult;
    private int _connectingFromCol = -1; // -1 = 연결 대기 없음

    // ── ETA 표준 필드 정의 ─────────────────────────────────────────────
    private static readonly (string Key, string Label)[] EtaFields =
    [
        ("SN",       "SN"),
        ("시료명",   "시료명"),
        ("흡광도",   "흡광도 / AU"),
        ("희석배수", "희석배수"),
        ("농도",     "농도"),
        ("결과",     "결과"),
        ("시료구분", "시료구분"),
        ("전무게",   "전무게"),
        ("후무게",   "후무게"),
        ("면적",     "면적"),
    ];

    // 연결선 색상 팔레트
    private static readonly string[] LineColors =
        ["#2563EB", "#7C3AED", "#0891B2", "#059669", "#D97706", "#DC2626", "#6366F1", "#BE185D", "#0F766E", "#B45309"];

    public ParserGeneratorPage() => InitializeComponent();

    // ── 파일 추가 + 분석 ─────────────────────────────────────────────
    public async Task UploadAndAnalyzeAsync()
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title          = "분석기 원본 파일 선택",
            AllowMultiple  = true,
            FileTypeFilter = [new FilePickerFileType("분석기 파일")
                { Patterns = ["*.csv", "*.txt", "*.xlsx", "*.pdf"] }],
        });
        if (files.Count == 0) return;

        foreach (var file in files)
        {
            try
            {
                await using var stream = await file.OpenReadAsync();
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                var ext  = System.IO.Path.GetExtension(file.Name).TrimStart('.').ToLower();
                var text = await Task.Run(() => ExtractText(ms.ToArray(), ext));
                var res  = await Task.Run(() => AnalyzeText(text, file.Name, ext));
                _results.Insert(0, res);
            }
            catch (Exception ex)
            {
                _results.Insert(0, new ParseGenResult { FileName = file.Name, Error = ex.Message });
            }
        }

        _connectingFromCol = -1;
        RefreshFileList();
        if (_results.Count > 0) ShowAnalysis(_results[0]);
    }

    // ── Show1 파일 목록 새로고침 ────────────────────────────────────────
    private void RefreshFileList()
    {
        FileListBox.Items.Clear();
        foreach (var r in _results)
            FileListBox.Items.Add(BuildFileListItem(r));
        if (FileListBox.Items.Count > 0) FileListBox.SelectedIndex = 0;
    }

    private static ListBoxItem BuildFileListItem(ParseGenResult r)
    {
        var ext      = (r.Format ?? "").ToUpper();
        var extColor = ext switch { "PDF" => "#DC2626", "XLSX" => "#16A34A", "CSV" or "TXT" => "#2563EB", _ => "#6B7280" };
        var ff       = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var panel    = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(2, 3) };

        panel.Children.Add(new Border
        {
            Width = 36, CornerRadius = new CornerRadius(3), Padding = new Thickness(3, 1),
            Background = new SolidColorBrush(Color.Parse(r.Error != null ? "#6B7280" : extColor)),
            Child = new TextBlock
            {
                Text = ext.Length > 0 ? ext : "?", FontSize = AppTheme.FontXS, FontFamily = ff,
                Foreground = Brushes.White, HorizontalAlignment = HorizontalAlignment.Center,
            },
        });
        panel.Children.Add(new TextBlock
        {
            Text = r.FileName, FontFamily = ff, FontSize = AppTheme.FontSM,
            VerticalAlignment = VerticalAlignment.Center, Width = 160,
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        if (r.FieldAssignments.Count > 0)
            panel.Children.Add(new Border
            {
                CornerRadius = new CornerRadius(8), Padding = new Thickness(5, 1),
                Background = new SolidColorBrush(Color.Parse("#2563EB")),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = $"연결 {r.FieldAssignments.Count}", FontFamily = ff,
                    FontSize = AppTheme.FontXS, Foreground = Brushes.White,
                },
            });
        return new ListBoxItem { Content = panel, Tag = r };
    }

    // ── Show2/Show3/Show4 패널 표시 ────────────────────────────────────
    private void ShowAnalysis(ParseGenResult r)
    {
        _currentResult = r;
        ListPanelChanged?.Invoke(BuildWireMapperPanel(r));
        DetailPanelChanged?.Invoke(BuildParserSettingsPanel(r));
        StatsPanelChanged?.Invoke(BuildCodePanel(r));
    }

    // ── Show2: 줄 잇기 파서 매핑 UI ─────────────────────────────────────
    private Control BuildWireMapperPanel(ParseGenResult r)
    {
        var ff   = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(10, 8) };

        root.Children.Add(new TextBlock
        {
            Text = r.FileName, FontFamily = ff, FontSize = AppTheme.FontSM,
            FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap,
        });

        if (r.Error != null)
        {
            root.Children.Add(ErrorCard(ff, r.Error));
            return new ScrollViewer { Content = root };
        }

        root.Children.Add(new TextBlock
        {
            Text = $"구분자: {r.Delimiter}  |  헤더: {r.HeaderLines}행  |  열: {r.ColumnCount}개  |  데이터: {r.DataLines:N0}행",
            FontFamily = ff, FontSize = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#9CA3AF")),
        });

        if (r.PreviewHeaders.Length == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "미리보기 데이터가 없습니다.",
                FontFamily = ff, FontSize = AppTheme.FontXS,
                Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
            });
            return new ScrollViewer { Content = root };
        }

        // 안내
        var guideText = _connectingFromCol >= 0
            ? $"← 열 {_connectingFromCol + 1} \"{r.PreviewHeaders[_connectingFromCol]}\" 선택됨  |  오른쪽 ETA 필드를 클릭하세요  (ESC = 취소)"
            : "왼쪽 열의 ● 을 클릭 → 오른쪽 ETA 필드를 클릭하여 연결  |  연결된 ● 을 다시 클릭하면 해제";
        root.Children.Add(new Border
        {
            Background = new SolidColorBrush(Color.Parse(_connectingFromCol >= 0 ? "#FEF3C7" : "#F1F5F9")),
            CornerRadius = new CornerRadius(6), Padding = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text = guideText, FontFamily = ff, FontSize = AppTheme.FontXS,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.Parse(_connectingFromCol >= 0 ? "#92400E" : "#475569")),
            },
        });

        // 매핑 패널 (3열: 파일컬럼 | 선 캔버스 | ETA필드)
        root.Children.Add(BuildMappingGrid(r, ff));

        return new ScrollViewer { Content = root };
    }

    private const int ItemH    = 54;  // 카드 높이
    private const int ItemGap  = 8;   // 카드 간격
    private const int ItemStep = 62;  // ItemH + ItemGap
    private const int CanvasW  = 110; // 중간 연결선 폭
    private const int LeftW    = 192; // 파일 컬럼 패널 폭
    private const int RightW   = 158; // ETA 필드 패널 폭

    private Control BuildMappingGrid(ParseGenResult r, FontFamily ff)
    {
        var headers   = r.PreviewHeaders;
        var leftCount = headers.Length;
        var rightCount= EtaFields.Length;
        int totalH    = Math.Max(leftCount, rightCount) * ItemStep;

        var lineCanvas = new Canvas { Width = CanvasW, Height = totalH };
        var leftCanvas = new Canvas { Width = LeftW,   Height = totalH };
        var rightCanvas= new Canvas { Width = RightW,  Height = totalH };

        // ── 연결선 그리기 ────────────────────────────────────────────
        for (int ci = 0; ci < r.FieldAssignments.Count; ci++)
        {
            var fa      = r.FieldAssignments[ci];
            int leftIdx = fa.ColumnIndex;
            int rightIdx= Array.FindIndex(EtaFields, f => f.Key == fa.Role);
            if (leftIdx < 0 || rightIdx < 0) continue;

            var color   = Color.Parse(LineColors[ci % LineColors.Length]);
            double startY = leftIdx  * ItemStep + ItemH / 2.0;
            double endY   = rightIdx * ItemStep + ItemH / 2.0;
            lineCanvas.Children.Add(MakeBezierPath(0, startY, CanvasW, endY, color));

            // 끝 점 원형 표시
            var dotL = new Ellipse { Width = 10, Height = 10, Fill = new SolidColorBrush(color) };
            Canvas.SetLeft(dotL, -5); Canvas.SetTop(dotL, startY - 5);
            lineCanvas.Children.Add(dotL);
            var dotR = new Ellipse { Width = 10, Height = 10, Fill = new SolidColorBrush(color) };
            Canvas.SetLeft(dotR, CanvasW - 5); Canvas.SetTop(dotR, endY - 5);
            lineCanvas.Children.Add(dotR);
        }

        // ── 파일 컬럼 카드 (왼쪽) ────────────────────────────────────
        for (int i = 0; i < leftCount; i++)
        {
            int colIdx   = i;
            var hdr      = headers[i];
            var fa       = r.FieldAssignments.FirstOrDefault(f => f.ColumnIndex == colIdx);
            int faIdx    = r.FieldAssignments.IndexOf(fa!);
            bool hasConn = fa != null;
            bool isPending = _connectingFromCol == colIdx;

            // 샘플 값 추출
            var sampleVals = r.PreviewRows.Take(4)
                .Select(row => row.Length > colIdx ? row[colIdx] : "")
                .Where(v => v.Length > 0).Take(3).ToArray();
            var sampleTxt = sampleVals.Length > 0 ? string.Join(",  ", sampleVals) : "—";

            var dotColor = isPending ? "#F59E0B"
                         : hasConn  ? LineColors[faIdx % LineColors.Length]
                         : "#94A3B8";

            var card = new Border
            {
                Width           = LeftW - 16,
                Height          = ItemH,
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(8, 6),
                Background      = new SolidColorBrush(Color.Parse(isPending ? "#FEF3C7" : hasConn ? "#F0FDF4" : "#F8FAFC")),
                BorderBrush     = new SolidColorBrush(Color.Parse(isPending ? "#F59E0B" : hasConn ? LineColors[faIdx % LineColors.Length] : "#E2E8F0")),
                BorderThickness = new Thickness(isPending || hasConn ? 2 : 1),
                Cursor          = new Cursor(StandardCursorType.Hand),
            };

            var innerSp = new StackPanel { Spacing = 2 };
            innerSp.Children.Add(new TextBlock
            {
                Text = hdr, FontFamily = ff, FontSize = AppTheme.FontXS,
                FontWeight = FontWeight.SemiBold,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Foreground = new SolidColorBrush(Color.Parse("#1E293B")),
            });
            innerSp.Children.Add(new TextBlock
            {
                Text = sampleTxt, FontFamily = ff, FontSize = 9,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Foreground = new SolidColorBrush(Color.Parse("#94A3B8")),
            });
            if (hasConn)
                innerSp.Children.Add(new TextBlock
                {
                    Text = $"→ {fa!.Role}", FontFamily = ff, FontSize = 9,
                    Foreground = new SolidColorBrush(Color.Parse(LineColors[faIdx % LineColors.Length])),
                });

            card.Child = innerSp;

            // 오른쪽 연결 점 (● )
            var dot = new Ellipse
            {
                Width = 14, Height = 14,
                Fill   = new SolidColorBrush(Color.Parse(dotColor)),
                Cursor = new Cursor(StandardCursorType.Hand),
            };
            dot.PointerPressed += (_, e) =>
            {
                e.Handled = true;
                if (isPending)
                {
                    _connectingFromCol = -1; // 취소
                }
                else if (hasConn)
                {
                    r.FieldAssignments.RemoveAll(f => f.ColumnIndex == colIdx);
                    _connectingFromCol = -1;
                }
                else
                {
                    _connectingFromCol = colIdx;
                }
                ShowAnalysis(r);
            };

            Canvas.SetLeft(card, 2); Canvas.SetTop(card, i * ItemStep);
            Canvas.SetLeft(dot, LeftW - 7); Canvas.SetTop(dot, i * ItemStep + ItemH / 2 - 7);
            leftCanvas.Children.Add(card);
            leftCanvas.Children.Add(dot);
        }

        // ── ETA 필드 슬롯 (오른쪽) ────────────────────────────────────
        for (int j = 0; j < EtaFields.Length; j++)
        {
            int fldIdx    = j;
            var (key, lbl) = EtaFields[j];
            var fa         = r.FieldAssignments.FirstOrDefault(f => f.Role == key);
            int faIdx      = r.FieldAssignments.IndexOf(fa!);
            bool hasConn   = fa != null;
            bool canConnect= _connectingFromCol >= 0;

            var dotColor   = hasConn ? LineColors[faIdx % LineColors.Length]
                           : canConnect ? "#94A3B8" : "#CBD5E1";

            var slot = new Border
            {
                Width           = RightW - 16,
                Height          = ItemH,
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 6),
                Background      = new SolidColorBrush(Color.Parse(hasConn ? "#EFF6FF" : canConnect ? "#F0FDF4" : "#1E293B")),
                BorderBrush     = new SolidColorBrush(Color.Parse(hasConn ? LineColors[faIdx % LineColors.Length] : canConnect ? "#86EFAC" : "#334155")),
                BorderThickness = new Thickness(hasConn ? 2 : 1),
                Cursor          = new Cursor(canConnect || hasConn ? StandardCursorType.Hand : StandardCursorType.Arrow),
            };

            var innerSp = new StackPanel { Spacing = 2 };
            innerSp.Children.Add(new TextBlock
            {
                Text = lbl, FontFamily = ff, FontSize = AppTheme.FontXS,
                FontWeight = FontWeight.SemiBold,
                Foreground = new SolidColorBrush(Color.Parse(hasConn ? "#1E40AF" : canConnect ? "#166534" : "#94A3B8")),
            });
            if (hasConn)
                innerSp.Children.Add(new TextBlock
                {
                    Text = $"← {fa!.ColumnHeader}", FontFamily = ff, FontSize = 9,
                    Foreground = new SolidColorBrush(Color.Parse(LineColors[faIdx % LineColors.Length])),
                });
            slot.Child = innerSp;

            slot.PointerPressed += (_, e) =>
            {
                e.Handled = true;
                if (_connectingFromCol >= 0)
                {
                    int fromCol  = _connectingFromCol;
                    string fromHdr = r.PreviewHeaders.Length > fromCol ? r.PreviewHeaders[fromCol] : $"Col{fromCol+1}";
                    r.FieldAssignments.RemoveAll(f => f.ColumnIndex == fromCol || f.Role == key);
                    r.FieldAssignments.Add(new FieldAssignment
                    {
                        ColumnIndex  = fromCol,
                        ColumnHeader = fromHdr,
                        Role         = key,
                    });
                    _connectingFromCol = -1;
                    r.GeneratedCode = GenerateParserCode(r, r.FileName);
                }
                else if (hasConn)
                {
                    r.FieldAssignments.RemoveAll(f => f.Role == key);
                    r.GeneratedCode = GenerateParserCode(r, r.FileName);
                }
                ShowAnalysis(r);
            };

            // 왼쪽 연결 점
            var dot = new Ellipse
            {
                Width = 14, Height = 14,
                Fill  = new SolidColorBrush(Color.Parse(dotColor)),
            };

            Canvas.SetLeft(slot, 14); Canvas.SetTop(slot, j * ItemStep);
            Canvas.SetLeft(dot, 0);   Canvas.SetTop(dot, j * ItemStep + ItemH / 2 - 7);
            rightCanvas.Children.Add(slot);
            rightCanvas.Children.Add(dot);
        }

        // ── 3열 Grid 조립 ─────────────────────────────────────────────
        var grid = new Grid { Height = totalH };
        grid.ColumnDefinitions.Add(new ColumnDefinition(LeftW,   GridUnitType.Pixel));
        grid.ColumnDefinitions.Add(new ColumnDefinition(CanvasW, GridUnitType.Pixel));
        grid.ColumnDefinitions.Add(new ColumnDefinition(RightW,  GridUnitType.Pixel));
        grid.RowDefinitions.Add(new RowDefinition(totalH, GridUnitType.Pixel));

        Grid.SetColumn(leftCanvas,  0); grid.Children.Add(leftCanvas);
        Grid.SetColumn(lineCanvas,  1); grid.Children.Add(lineCanvas);
        Grid.SetColumn(rightCanvas, 2); grid.Children.Add(rightCanvas);

        // 왼쪽/오른쪽 컬럼 헤더
        var headerGrid = new Grid();
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition(LeftW,   GridUnitType.Pixel));
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition(CanvasW, GridUnitType.Pixel));
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition(RightW,  GridUnitType.Pixel));
        Grid.SetColumn(new TextBlock { Text = "파일 컬럼 (자동 감지)", FontFamily = ff, FontSize = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#64748B")), FontWeight = FontWeight.SemiBold }, 0);
        var hdrTb1 = new TextBlock { Text = "파일 컬럼 (자동 감지)", FontFamily = ff, FontSize = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#64748B")), FontWeight = FontWeight.SemiBold };
        var hdrTb2 = new TextBlock { Text = "ETA 표준 필드", FontFamily = ff, FontSize = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#64748B")), FontWeight = FontWeight.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Right };
        Grid.SetColumn(hdrTb1, 0); Grid.SetColumn(hdrTb2, 2);
        headerGrid.Children.Add(hdrTb1); headerGrid.Children.Add(hdrTb2);

        var outer = new StackPanel { Spacing = 4 };
        outer.Children.Add(headerGrid);
        outer.Children.Add(new Border
        {
            CornerRadius = new CornerRadius(8),
            Background   = new SolidColorBrush(Color.Parse("#0F172A")),
            Padding      = new Thickness(0),
            Child        = grid,
        });

        return new ScrollViewer
        {
            Content = outer,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // ── Show3: 검정곡선 키워드 + 연결 요약 + 저장 ────────────────────
    private Control BuildParserSettingsPanel(ParseGenResult r)
    {
        var ff   = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(12, 12) };

        root.Children.Add(new TextBlock
        {
            Text = "파서 설정 및 저장", FontFamily = ff,
            FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
        });

        if (r.Error != null)
        {
            root.Children.Add(new TextBlock
            {
                Text = "오류가 있는 파일은 저장할 수 없습니다.",
                FontFamily = ff, FontSize = AppTheme.FontXS,
                Foreground = new SolidColorBrush(Color.Parse("#EF4444")),
            });
            return new ScrollViewer { Content = root };
        }

        // 검정곡선 키워드
        root.Children.Add(SectionTitle(ff, "검정곡선 행 키워드"));
        root.Children.Add(new TextBlock
        {
            Text = "이 문자열로 시작하는 행을 검정곡선으로 처리합니다.",
            FontFamily = ff, FontSize = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
        });

        var kwWrap = new WrapPanel { Orientation = Orientation.Horizontal };
        void RebuildKwChips()
        {
            kwWrap.Children.Clear();
            foreach (var kw in r.CalRowKeywords.ToList())
            {
                var kw2 = kw;
                var chip = new Border
                {
                    Margin = new Thickness(2), CornerRadius = new CornerRadius(10),
                    Padding = new Thickness(8, 3),
                    Background = new SolidColorBrush(Color.Parse("#92400E")),
                };
                var inner = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
                inner.Children.Add(new TextBlock
                {
                    Text = kw2, FontFamily = ff, FontSize = AppTheme.FontXS,
                    Foreground = new SolidColorBrush(Color.Parse("#FCD34D")),
                });
                var del = new TextBlock
                {
                    Text = "×", FontFamily = ff, FontSize = AppTheme.FontXS,
                    Foreground = Brushes.White, Cursor = new Cursor(StandardCursorType.Hand),
                };
                del.PointerPressed += (_, e) =>
                {
                    e.Handled = true;
                    r.CalRowKeywords.Remove(kw2);
                    r.GeneratedCode = GenerateParserCode(r, r.FileName);
                    RebuildKwChips();
                    StatsPanelChanged?.Invoke(BuildCodePanel(r));
                };
                inner.Children.Add(del);
                chip.Child = inner;
                kwWrap.Children.Add(chip);
            }
            var addBox = new TextBox
            {
                Watermark = "+ 키워드 추가", FontFamily = ff, FontSize = AppTheme.FontXS,
                Padding = new Thickness(6, 3), Width = 120, Margin = new Thickness(2),
            };
            addBox.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter && !string.IsNullOrWhiteSpace(addBox.Text))
                {
                    var kw = addBox.Text.Trim();
                    if (!r.CalRowKeywords.Contains(kw)) r.CalRowKeywords.Add(kw);
                    addBox.Text = "";
                    r.GeneratedCode = GenerateParserCode(r, r.FileName);
                    RebuildKwChips();
                    StatsPanelChanged?.Invoke(BuildCodePanel(r));
                }
            };
            kwWrap.Children.Add(addBox);
        }
        RebuildKwChips();
        root.Children.Add(kwWrap);

        // 연결 요약
        root.Children.Add(SectionTitle(ff, "컬럼 연결 요약"));
        if (r.FieldAssignments.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "(없음 — Show2에서 ● 를 클릭하여 연결하세요)",
                FontFamily = ff, FontSize = AppTheme.FontXS,
                Foreground = new SolidColorBrush(Color.Parse("#9CA3AF")),
            });
        }
        else
        {
            var sp = new StackPanel { Spacing = 4 };
            for (int i = 0; i < r.FieldAssignments.Count; i++)
            {
                var fa    = r.FieldAssignments[i];
                var color = LineColors[i % LineColors.Length];
                var row   = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                row.Children.Add(new Border
                {
                    CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 2),
                    Background   = new SolidColorBrush(Color.Parse(color)),
                    Child = new TextBlock
                    {
                        Text = fa.Role, FontFamily = ff, FontSize = AppTheme.FontXS,
                        Foreground = Brushes.White, Width = 60,
                    },
                });
                row.Children.Add(new TextBlock
                {
                    Text = $"← 열 {fa.ColumnIndex + 1}: {fa.ColumnHeader}",
                    FontFamily = ff, FontSize = AppTheme.FontXS,
                    Foreground = new SolidColorBrush(Color.Parse("#374151")),
                    VerticalAlignment = VerticalAlignment.Center,
                });
                sp.Children.Add(row);
            }
            root.Children.Add(new Border
            {
                CornerRadius = new CornerRadius(6), Padding = new Thickness(10, 8),
                Background   = new SolidColorBrush(Color.Parse("#F0F9FF")),
                Child        = sp,
            });
        }

        // 파서 타입 선택
        root.Children.Add(SectionTitle(ff, "파서 타입"));
        var parserItems = ETA.Services.SERVICE4.SignatureClassifier.ParserItems;
        var combo = new ComboBox
        {
            FontFamily = ff, FontSize = AppTheme.FontSM,
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        foreach (var (name, key) in parserItems)
            combo.Items.Add(new ComboBoxItem { Content = name, Tag = key });
        combo.SelectedIndex = 0;
        root.Children.Add(combo);

        var statusTb = new TextBlock
        {
            FontFamily = ff, FontSize = AppTheme.FontXS,
            Margin = new Thickness(0, 4, 0, 0), IsVisible = false,
        };
        root.Children.Add(statusTb);

        var saveBtn = new Button
        {
            Content = "시그니처 저장",
            FontFamily = ff, FontSize = AppTheme.FontSM,
            Padding = new Thickness(16, 8),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Background = new SolidColorBrush(Color.Parse("#2563EB")),
            Foreground = Brushes.White, Margin = new Thickness(0, 6, 0, 0),
        };
        saveBtn.Click += (_, _) =>
        {
            if (combo.SelectedItem is not ComboBoxItem item || item.Tag is not string parserKey)
            {
                statusTb.Text = "파서 타입을 선택해 주세요.";
                statusTb.Foreground = new SolidColorBrush(Color.Parse("#EF4444"));
                statusTb.IsVisible = true; return;
            }
            try
            {
                var parserName  = item.Content?.ToString() ?? parserKey;
                var nameColIdx  = r.FieldAssignments.FirstOrDefault(f => f.Role == "SN" || f.Role == "시료명")?.ColumnIndex ?? r.NameColumnIndex;
                var sig = new ETA.Services.SERVICE4.ParserSignature
                {
                    ParserKey       = parserKey,
                    ParserName      = parserName,
                    StdKeywords     = r.CalRowKeywords.Take(6).ToList(),
                    HeaderKeywords  = r.FieldAssignments.Count > 0
                        ? r.FieldAssignments.Select(f => f.ColumnHeader).Where(h => h.Length >= 2).ToList()
                        : r.DataColumns.Take(8).Select(c => c.Header).Where(h => h.Length >= 2).ToList(),
                    SampleKeywords  = new List<string>(),
                    Delimiter       = r.DelimChar.ToString(),
                    HeaderLines     = r.HeaderLines,
                    NameColumnIndex = nameColIdx,
                };
                ETA.Services.SERVICE4.SignatureClassifier.SaveSignature(sig);
                statusTb.Text = $"✅ 저장 완료 — {parserName}";
                statusTb.Foreground = new SolidColorBrush(Color.Parse("#059669"));
            }
            catch (Exception ex)
            {
                statusTb.Text = $"❌ 저장 실패: {ex.Message}";
                statusTb.Foreground = new SolidColorBrush(Color.Parse("#EF4444"));
            }
            statusTb.IsVisible = true;
        };
        root.Children.Add(saveBtn);
        return new ScrollViewer { Content = root };
    }

    // ── Show4: 생성된 C# 파서 코드 ─────────────────────────────────────
    private static Control BuildCodePanel(ParseGenResult r)
    {
        var ff   = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(10, 8) };

        root.Children.Add(new TextBlock
        {
            Text = "생성된 파서 코드", FontFamily = ff,
            FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
        });

        if (r.Error != null || string.IsNullOrEmpty(r.GeneratedCode))
        {
            root.Children.Add(new TextBlock
            {
                Text = "연결 완료 후 코드가 자동 생성됩니다.",
                FontFamily = ff, FontSize = AppTheme.FontXS,
                Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
            });
            return new ScrollViewer { Content = root };
        }

        root.Children.Add(new TextBlock
        {
            Text = r.GeneratedCode,
            FontFamily = new FontFamily("Consolas, Courier New"),
            FontSize = AppTheme.FontXS,
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.Parse("#1E293B")),
        });
        return new ScrollViewer { Content = root };
    }

    // ── Bezier 연결선 생성 ────────────────────────────────────────────
    private static Avalonia.Controls.Shapes.Path MakeBezierPath(double x1, double y1, double x2, double y2, Color color)
    {
        double cpX = (x1 + x2) / 2.0;
        var pathStr = string.Format(CultureInfo.InvariantCulture,
            "M {0} {1} C {2} {3} {4} {5} {6} {7}",
            x1, y1, cpX, y1, cpX, y2, x2, y2);
        return new Avalonia.Controls.Shapes.Path
        {
            Data = Geometry.Parse(pathStr),
            Stroke = new SolidColorBrush(color),
            StrokeThickness = 2.5,
        };
    }

    // ── 정적 UI 헬퍼 ─────────────────────────────────────────────────
    private static Border ErrorCard(FontFamily ff, string msg) =>
        new Border
        {
            CornerRadius = new CornerRadius(6), Padding = new Thickness(10, 8),
            Background   = new SolidColorBrush(Color.Parse("#FEE2E2")),
            Child        = new TextBlock { Text = msg, FontFamily = ff, FontSize = AppTheme.FontXS, TextWrapping = TextWrapping.Wrap },
        };

    private static TextBlock SectionTitle(FontFamily ff, string title) =>
        new TextBlock
        {
            Text = title, FontFamily = ff, FontSize = AppTheme.FontXS,
            FontWeight = FontWeight.SemiBold,
            Foreground = new SolidColorBrush(Color.Parse("#374151")),
            Margin = new Thickness(0, 4, 0, 2),
        };

    // ── 텍스트 추출 ───────────────────────────────────────────────────
    private static string ExtractText(byte[] data, string ext)
    {
        if (data.Length == 0) return "";
        try
        {
            return ext switch
            {
                "xlsx" => ExtractXlsxText(data),
                "pdf"  => ExtractPdfText(data),
                _      => DetectAndDecode(data),
            };
        }
        catch { return ""; }
    }

    private static string DetectAndDecode(byte[] data)
    {
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            return Encoding.UTF8.GetString(data, 3, data.Length - 3);
        try { return Encoding.UTF8.GetString(data); }
        catch { return Encoding.GetEncoding("EUC-KR").GetString(data); }
    }

    private static string ExtractXlsxText(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var wb = new XLWorkbook(ms);
        var sb = new StringBuilder();
        foreach (var ws in wb.Worksheets)
            foreach (var row in ws.RowsUsed().Take(200))
                sb.AppendLine(string.Join("\t", row.CellsUsed().Select(c => c.GetString())));
        return sb.ToString();
    }

    private static string ExtractPdfText(byte[] data)
    {
        using var ms  = new MemoryStream(data);
        using var doc = UglyToad.PdfPig.PdfDocument.Open(ms);
        var sb = new StringBuilder();
        foreach (var page in doc.GetPages().Take(5))
        {
            foreach (var word in page.GetWords())
                sb.Append(word.Text).Append(' ');
            sb.AppendLine();
        }
        return sb.ToString();
    }

    // ── 분석 ─────────────────────────────────────────────────────────
    private static ParseGenResult AnalyzeText(string text, string fileName, string ext)
    {
        var r = new ParseGenResult { FileName = fileName, Format = ext };

        if (string.IsNullOrWhiteSpace(text)) { r.Error = "텍스트를 추출할 수 없습니다."; return r; }

        var lines    = text.Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
        r.TotalLines = lines.Length;

        char delim   = DetectDelimiter(lines, ext);
        r.DelimChar  = delim;
        r.Delimiter  = delim == ',' ? "쉼표(,)" : delim == '\t' ? "탭(\\t)" : delim == ';' ? "세미콜론(;)" : "공백";

        int hdr      = DetectHeaderLines(lines, delim);
        r.HeaderLines = hdr;

        var dataLines = lines.Skip(hdr).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();
        r.DataLines   = dataLines.Length;

        string[]? headerRow = hdr > 0 ? SplitLine(lines[hdr - 1], delim) : null;
        r.ColumnCount = headerRow?.Length ?? (dataLines.Length > 0 ? SplitLine(dataLines[0], delim).Length : 0);

        r.StdKeywords     = DetectStdKeywords(dataLines, delim);
        var (nameCol, pattern) = DetectSamplePattern(dataLines, delim, r.StdKeywords, headerRow);
        r.NameColumnIndex = nameCol;
        r.SamplePattern   = pattern;
        r.DataColumns     = DetectDataColumns(dataLines, delim, headerRow, nameCol);

        // ── 미리보기 데이터 (줄 잇기 UI용) ──────────────────────────
        r.PreviewHeaders = headerRow
            ?? Enumerable.Range(1, r.ColumnCount).Select(i => $"Col{i}").ToArray();
        r.PreviewRows    = dataLines.Take(20).Select(l => SplitLine(l, delim)).ToArray();
        r.CalRowKeywords = r.StdKeywords.Take(3).Select(k => k.Keyword).ToList();

        r.GeneratedCode  = GenerateParserCode(r, fileName);
        return r;
    }

    // ── 구분자 탐지 ───────────────────────────────────────────────────
    private static char DetectDelimiter(string[] lines, string ext)
    {
        if (ext == "xlsx") return '\t';
        var sample = lines.Where(l => !string.IsNullOrWhiteSpace(l)).Take(10).ToArray();
        if (sample.Length == 0) return ',';
        var best = new[] { ',', '\t', ';' }
            .Select(c => (Delim: c, Avg: sample.Average(l => l.Count(x => x == c))))
            .OrderByDescending(x => x.Avg)
            .FirstOrDefault();
        return best.Avg > 0 ? best.Delim : ',';
    }

    private static string[] SplitLine(string line, char delim)
    {
        var result  = new List<string>();
        var current = new StringBuilder();
        bool inQ    = false;
        foreach (char c in line)
        {
            if (c == '"')               { inQ = !inQ; }
            else if (c == delim && !inQ){ result.Add(current.ToString().Trim('"', ' ')); current.Clear(); }
            else                        { current.Append(c); }
        }
        result.Add(current.ToString().Trim('"', ' '));
        return result.ToArray();
    }

    // ── 헤더 행 탐지 ─────────────────────────────────────────────────
    private static int DetectHeaderLines(string[] lines, char delim)
    {
        for (int i = 0; i < Math.Min(lines.Length, 15); i++)
        {
            var cols = SplitLine(lines[i], delim);
            int numC = cols.Count(c => double.TryParse(c,
                System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out _));
            if (numC > cols.Length / 2 && cols.Length > 2) return i;
        }
        return 1;
    }

    // ── 표준물질 키워드 탐지 ─────────────────────────────────────────
    private static readonly string[] StdPrefixes =
    [
        "STD", "Standard", "Cal", "CAL", "Calibration",
        "Blank", "BLANK", "BLK", "QC", "Check", "CCV", "ICV", "CCB", "ICB", "BQC",
        "MB", "MBK", "FBK", "Wash", "Spike", "MS", "MSD",
        "공시험", "표준", "표준액", "검량",
    ];

    private static List<StdKeywordInfo> DetectStdKeywords(string[] dataLines, char delim)
    {
        var found = new Dictionary<string, (int Count, string Example)>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in dataLines.Take(500))
        {
            var cols  = SplitLine(line, delim);
            if (cols.Length == 0) continue;
            var first = cols[0].Trim();
            if (string.IsNullOrWhiteSpace(first)) continue;
            foreach (var pfx in StdPrefixes)
            {
                if (!first.StartsWith(pfx, StringComparison.OrdinalIgnoreCase)) continue;
                var key = Regex.Match(first, @"^[A-Za-z가-힣_\-]+").Value;
                if (string.IsNullOrEmpty(key)) key = pfx;
                var ex = first.Length > 30 ? first[..30] : first;
                if (found.TryGetValue(key, out var old)) found[key] = (old.Count + 1, old.Example);
                else                                     found[key] = (1, ex);
                break;
            }
        }
        return found.OrderByDescending(kv => kv.Value.Count)
            .Select(kv => new StdKeywordInfo { Keyword = kv.Key, Count = kv.Value.Count, Example = kv.Value.Example })
            .ToList();
    }

    // ── 시료명 열 + 패턴 탐지 ────────────────────────────────────────
    private static (int ColIdx, string Pattern) DetectSamplePattern(
        string[] dataLines, char delim, List<StdKeywordInfo> stdKw, string[]? headerRow)
    {
        var stdSet = stdKw.Select(k => k.Keyword).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var sampleLines = dataLines
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Where(l => { var c = SplitLine(l, delim); return c.Length > 0 && !stdSet.Any(k => c[0].StartsWith(k, StringComparison.OrdinalIgnoreCase)); })
            .Take(50).ToArray();

        if (sampleLines.Length == 0) return (0, "시료 행을 찾지 못했습니다.");

        int nameCol    = 0;
        var sampleNames= sampleLines.Select(l => SplitLine(l, delim)).Where(c => c.Length > nameCol)
            .Select(c => c[nameCol].Trim()).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct().Take(10).ToArray();

        var sb = new StringBuilder();
        sb.AppendLine($"시료명 열: {nameCol + 1}번째 열");
        if (headerRow != null && headerRow.Length > nameCol)
            sb.AppendLine($"열 헤더:   \"{headerRow[nameCol]}\"");
        sb.AppendLine();

        bool hasCode  = sampleNames.Any(n => Regex.IsMatch(n, @"[A-Za-z]{1,4}-?\d+"));
        bool hasDate  = sampleNames.Any(n => Regex.IsMatch(n, @"\d{2,4}[-./]\d{1,2}[-./]\d{1,2}"));
        bool hasSeq   = sampleNames.Any(n => Regex.IsMatch(n, @"[-_]?\d+\.?[Dd]?$"));
        bool hasKorean= sampleNames.Any(n => n.Any(c => c >= 0xAC00 && c <= 0xD7A3));

        sb.AppendLine("탐지된 패턴:");
        if (hasCode)   sb.AppendLine("  • 코드형  (예: RE-07, WW-01)");
        if (hasDate)   sb.AppendLine("  • 날짜형  (예: 2024-01-15)");
        if (hasSeq)    sb.AppendLine("  • 연번형  (예: Sample-1.D)");
        if (hasKorean) sb.AppendLine("  • 한글 포함");
        if (!hasCode && !hasDate && !hasSeq && !hasKorean)
            sb.AppendLine("  • 패턴 불명확");

        sb.AppendLine(); sb.AppendLine("시료명 예시:");
        foreach (var name in sampleNames.Take(6)) sb.AppendLine($"  {name}");
        return (nameCol, sb.ToString().TrimEnd());
    }

    // ── 데이터 열 탐지 ────────────────────────────────────────────────
    private static List<DataColumnInfo> DetectDataColumns(
        string[] dataLines, char delim, string[]? headerRow, int nameColIdx)
    {
        var result = new List<DataColumnInfo>();
        if (dataLines.Length == 0) return result;
        var parsed  = dataLines.Take(30).Select(l => SplitLine(l, delim)).Where(c => c.Length > 1).ToArray();
        if (parsed.Length == 0) return result;
        int maxCols = parsed.Max(c => c.Length);
        for (int col = 0; col < maxCols; col++)
        {
            if (col == nameColIdx) continue;
            var vals    = parsed.Select(c => col < c.Length ? c[col].Trim() : "").ToArray();
            int numericC = vals.Count(v => double.TryParse(v, System.Globalization.NumberStyles.Any,
                CultureInfo.InvariantCulture, out _) && v.Length > 0);
            if (numericC >= parsed.Length * 0.4)
            {
                var header = headerRow != null && col < headerRow.Length ? headerRow[col] : $"Col{col + 1}";
                var sample = vals.FirstOrDefault(v => v.Length > 0) ?? "";
                result.Add(new DataColumnInfo { Index = col, Header = header, SampleValue = sample });
            }
        }
        return result;
    }

    // ── C# 파서 코드 생성 (FieldAssignments 반영) ─────────────────────
    private static string GenerateParserCode(ParseGenResult r, string fileName)
    {
        var baseName  = Regex.Replace(System.IO.Path.GetFileNameWithoutExtension(fileName), @"[^A-Za-z0-9]", "");
        var className = string.IsNullOrEmpty(baseName) ? "Custom"
                      : char.IsLetter(baseName[0]) ? baseName : "P" + baseName;

        bool hasMap      = r.FieldAssignments.Count > 0;
        var calList      = r.CalRowKeywords.Count > 0
            ? string.Join(", ", r.CalRowKeywords.Select(k => $"\"{k}\""))
            : "\"STD\", \"Blank\", \"QC\"";
        var delimStr     = r.DelimChar == '\t' ? "'\\t'" : $"'{r.DelimChar}'";
        int nameColIdx   = hasMap
            ? (r.FieldAssignments.FirstOrDefault(f => f.Role == "SN" || f.Role == "시료명")?.ColumnIndex ?? r.NameColumnIndex)
            : r.NameColumnIndex;

        var sb = new StringBuilder();
        sb.AppendLine($"// 자동 생성 파서 — {fileName}");
        sb.AppendLine($"// 생성일: {DateTime.Now:yyyy-MM-dd}");
        sb.AppendLine($"// 구분자: {r.Delimiter}  |  헤더: {r.HeaderLines}행  |  시료명 열: {nameColIdx + 1}번째");
        sb.AppendLine();
        sb.AppendLine("using System; using System.Collections.Generic;");
        sb.AppendLine("using System.IO; using System.Linq; using System.Text;");
        sb.AppendLine();
        sb.AppendLine("namespace ETA.Services.SERVICE4;");
        sb.AppendLine();
        sb.AppendLine($"public class {className}Row");
        sb.AppendLine("{");
        sb.AppendLine("    public string SN   { get; set; } = \"\";");
        sb.AppendLine("    public string Type { get; set; } = \"\"; // Sample/Cal/Blank/QC");
        if (hasMap)
        {
            foreach (var fa in r.FieldAssignments.Where(f => f.Role != "SN" && f.Role != "시료명"))
            {
                var prop = fa.Role;
                sb.AppendLine($"    public string {prop} {{ get; set; }} = \"\";  // 열 {fa.ColumnIndex + 1}: {fa.ColumnHeader}");
            }
        }
        else
        {
            foreach (var col in r.DataColumns.Take(4))
            {
                var prop = Regex.Replace(col.Header, @"[^A-Za-z0-9가-힣]", "");
                if (string.IsNullOrEmpty(prop)) prop = $"Col{col.Index + 1}";
                sb.AppendLine($"    public string {prop} {{ get; set; }} = \"\";  // 열 {col.Index + 1}");
            }
        }
        sb.AppendLine("}");
        sb.AppendLine();
        sb.AppendLine($"public class {className}InstrumentFile");
        sb.AppendLine("{");
        sb.AppendLine($"    public List<{className}Row> Rows {{ get; }} = new();");
        sb.AppendLine("}");
        sb.AppendLine();
        sb.AppendLine($"public static class {className}Parser");
        sb.AppendLine("{");
        sb.AppendLine($"    private static readonly HashSet<string> CalKeywords =");
        sb.AppendLine($"        new(StringComparer.OrdinalIgnoreCase) {{ {calList} }};");
        sb.AppendLine($"    private const int NameCol = {nameColIdx};");
        sb.AppendLine();
        sb.AppendLine($"    public static {className}InstrumentFile Parse(string filePath)");
        sb.AppendLine("    {");
        sb.AppendLine($"        var file  = new {className}InstrumentFile();");
        sb.AppendLine("        var lines = File.ReadAllLines(filePath, Encoding.UTF8);");
        sb.AppendLine($"        foreach (var line in lines.Skip({r.HeaderLines}))");
        sb.AppendLine("        {");
        sb.AppendLine("            if (string.IsNullOrWhiteSpace(line)) continue;");
        sb.AppendLine($"            var cols  = line.Split({delimStr});");
        sb.AppendLine("            if (cols.Length <= NameCol) continue;");
        sb.AppendLine("            var name  = cols[NameCol].Trim().Trim('\"');");
        sb.AppendLine("            bool isCal = CalKeywords.Any(k => name.StartsWith(k, StringComparison.OrdinalIgnoreCase));");
        sb.AppendLine($"            var row   = new {className}Row");
        sb.AppendLine("            {");
        sb.AppendLine("                SN   = name,");
        sb.AppendLine("                Type = isCal ? \"Cal\" : \"Sample\",");
        if (hasMap)
        {
            foreach (var fa in r.FieldAssignments.Where(f => f.Role != "SN" && f.Role != "시료명"))
                sb.AppendLine($"                {fa.Role} = cols.Length > {fa.ColumnIndex} ? cols[{fa.ColumnIndex}].Trim() : \"\",");
        }
        else if (r.DataColumns.Count > 0)
        {
            var col  = r.DataColumns[0];
            var prop = Regex.Replace(col.Header, @"[^A-Za-z0-9가-힣]", "");
            if (string.IsNullOrEmpty(prop)) prop = $"Col{col.Index + 1}";
            sb.AppendLine($"                {prop} = cols.Length > {col.Index} ? cols[{col.Index}].Trim() : \"\",");
        }
        sb.AppendLine("            };");
        sb.AppendLine("            file.Rows.Add(row);");
        sb.AppendLine("        }");
        sb.AppendLine("        return file;");
        sb.AppendLine("    }");
        sb.AppendLine("}");
        return sb.ToString();
    }

    // ── 이벤트 ────────────────────────────────────────────────────────
    private void FileListBox_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count > 0 && e.AddedItems[0] is ListBoxItem lbi && lbi.Tag is ParseGenResult r)
        {
            _connectingFromCol = -1;
            ShowAnalysis(r);
        }
    }
}

// ── 모델 ──────────────────────────────────────────────────────────────────
public class StdKeywordInfo
{
    public string Keyword { get; set; } = "";
    public int    Count   { get; set; }
    public string Example { get; set; } = "";
}

public class DataColumnInfo
{
    public int    Index       { get; set; }
    public string Header      { get; set; } = "";
    public string SampleValue { get; set; } = "";
}

public class FieldAssignment
{
    public int    ColumnIndex  { get; set; } = -1;
    public string ColumnHeader { get; set; } = "";
    public string Role         { get; set; } = ""; // ETA 표준 필드 키
}

public class ParseGenResult
{
    public string  FileName         { get; set; } = "";
    public string? Format           { get; set; }
    public string? Error            { get; set; }
    public int     TotalLines       { get; set; }
    public int     HeaderLines      { get; set; }
    public int     DataLines        { get; set; }
    public int     ColumnCount      { get; set; }
    public string  Delimiter        { get; set; } = "";
    public char    DelimChar        { get; set; } = ',';
    public List<StdKeywordInfo>  StdKeywords     { get; set; } = new();
    public string  SamplePattern    { get; set; } = "";
    public int     NameColumnIndex  { get; set; }
    public List<DataColumnInfo>  DataColumns     { get; set; } = new();
    public string  GeneratedCode    { get; set; } = "";
    // ── 줄 잇기 UI 용 ────────────────────────────────────────────────
    public string[]            PreviewHeaders   { get; set; } = Array.Empty<string>();
    public string[][]          PreviewRows      { get; set; } = Array.Empty<string[]>();
    public List<FieldAssignment> FieldAssignments { get; set; } = new();
    public List<string>        CalRowKeywords   { get; set; } = new();
}
