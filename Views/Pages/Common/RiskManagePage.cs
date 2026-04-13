using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading.Tasks;
using ETA.Views;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 리스크관리 — 시약/초자 관리
/// TreeControl  → Show1 (카테고리 버튼 + 검색 + 카드 목록)
/// UsageControl → Show2 (사용량 현황 프로그레스바)
/// FormControl  → Show3 (시약 정보 등록·수정 폼)
/// </summary>
public class RiskManagePage
{
    private static readonly FontFamily FontM = new("avares://ETA/Assets/Fonts#Pretendard");
    private static readonly FontFamily FontR = new("avares://ETA/Assets/Fonts#Pretendard");

    // 초성별 색상 (ㄱ~ㅎ 19개)
    private static readonly string[] ChosungColors =
    {
        "#b5455a", "#a03030", "#c06820", "#a08020", "#708020",
        "#3d8c3d", "#1e8c7a", "#2270b0", "#1a4d99", "#3a5599",
        "#5a3daa", "#7a3399", "#993377", "#bb2255", "#22774d",
        "#1a6655", "#7a4410", "#556677", "#334455",
    };

    // ── 4패널 ──────────────────────────────────────────────────────────────
    public Control TreeControl   { get; }   // Show1
    public Control UsageControl  { get; }   // Show2
    public Control FormControl   { get; }   // Show3
    public Control DetailControl { get; }   // Show4 — 분석항목 연결

    // ── Show1 내부 상태 ───────────────────────────────────────────────────
    private string _category = "시약";
    private readonly TextBox    _searchBox;
    private readonly StackPanel _listPanel = new() { Spacing = 3 };
    private readonly TextBlock  _listCount = new()
    {
        FontSize = AppTheme.FontSM, FontFamily = FontR,
        Foreground = AppTheme.FgDimmed,
        Margin = new Thickness(0, 4, 0, 4),
    };
    private Border? _selectedCard;

    // ── Show4 — 분석항목 연결 패널 ──────────────────────────────────────────
    private readonly StackPanel _analytePanel = new() { Spacing = 4 };

    // ── Show2 — 사용량 패널 ───────────────────────────────────────────────
    private readonly StackPanel _usagePanel = new() { Spacing = 4 };

    // ── Show3 — 시약 폼 필드 ──────────────────────────────────────────────
    private readonly StackPanel _formPanel = new() { Spacing = 8 };

    private readonly TextBox  _rItemNo   = MakeTxb("ITEM NO.");
    private readonly TextBox  _rName     = MakeTxb("국문명 (필수)");
    private readonly TextBox  _rEng      = MakeTxb("영문명");
    private readonly TextBox  _rFormula  = MakeTxb("화학식");
    private readonly TextBox  _rCas      = MakeTxb("CAS NO.");
    private readonly TextBox  _rSpec     = MakeTxb("규격 (예: 500mL, 1kg)");
    private readonly TextBox  _rUnit     = MakeTxb("단위 (예: mL, g, L)");
    private readonly ComboBox _rRisk     = MakeCombo(new[] { "일반", "주의", "위험" });
    private readonly TextBox  _rGhs      = MakeTxb("GHS 분류 (예: GHS02,GHS06)");
    private readonly TextBox  _rStorage  = MakeTxb("보관조건 (예: 냉장, 차광)");
    private readonly TextBox  _rMaker    = MakeTxb("제조사");
    private readonly TextBox  _rExpiry   = MakeTxb("만료일 (yyyy-MM-dd)");
    private readonly TextBox  _rNote     = MakeTxb("비고");
    private readonly ComboBox _rStatus   = MakeCombo(new[] { "정상", "주의", "폐기" });
    private readonly TextBox  _rStock    = MakeTxb("재고량");
    private readonly TextBox  _rOptUse   = MakeTxb("적정사용량");
    private readonly TextBox  _rMaxStock = MakeTxb("최대 적정보유량");
    private readonly TextBox  _rCurrUse  = MakeTxb("당월사용량");
    private readonly TextBox  _rPrevUse  = MakeTxb("전월사용량");
    private Reagent? _editingReagent;

    // ── 초자 폼 필드 ───────────────────────────────────────────────────────
    private readonly TextBox  _gName     = MakeTxb("품목명 (필수)");
    private readonly TextBox  _gUse      = MakeTxb("용도");
    private readonly TextBox  _gSpec     = MakeTxb("규격 (예: 100mL)");
    private readonly ComboBox _gMaterial = MakeCombo(new[] { "유리", "플라스틱", "기타" });
    private readonly TextBox  _gQty      = MakeTxb("수량");
    private readonly TextBox  _gPrice    = MakeTxb("단가");
    private readonly TextBox  _gNote     = MakeTxb("비고");
    private readonly ComboBox _gStatus   = MakeCombo(new[] { "정상", "파손", "폐기" });
    private Glassware? _editingGlass;

    public RiskManagePage()
    {
        _searchBox = MakeTxb("🔍  검색...");
        _searchBox.TextChanged += (_, _) => RenderList();

        TreeControl   = BuildListPanel();
        UsageControl  = BuildUsagePanel();
        FormControl   = BuildFormPanel();
        DetailControl = BuildAnalytePanel();

        ShowUsagePlaceholder();
        ShowAnalytePlaceholder();

        // 최초 실행 시 Excel → DB 자동 가져오기 후 목록 렌더
        Task.Run(() => ImportReagentsIfEmpty())
            .ContinueWith(_ =>
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    RenderList();
                    ShowReagentNewForm();
                }));
    }

    // =========================================================================
    // 공개 API
    // =========================================================================
    public void SelectCategory(string cat)
    {
        _category = cat;
        RenderList();
        ClearDetail();
    }

    public void Refresh() => RenderList();

    public void DeleteSelected()
    {
        if (_category == "시약" && _editingReagent != null)
        {
            ReagentService.Delete(_editingReagent.Id);
            _editingReagent = null;
            RenderList();
            ClearDetail();
        }
        else if (_category == "초자" && _editingGlass != null)
        {
            GlasswareService.Delete(_editingGlass.Id);
            _editingGlass = null;
            RenderList();
            ClearDetail();
        }
    }

    // =========================================================================
    // Show1 — 카드 목록 패널
    // =========================================================================
    private Control BuildListPanel()
    {
        var title = new TextBlock
        {
            Text = "시약관리", FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            FontFamily = FontM, Foreground = AppTheme.FgPrimary,
            Margin = new Thickness(0, 0, 0, 10),
        };

        var scroll = new ScrollViewer
        {
            Content = _listPanel,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };

        var root = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*"),
            Margin = new Thickness(10, 12, 10, 4),
        };
        Grid.SetRow(title,      0); root.Children.Add(title);
        Grid.SetRow(_searchBox, 1); root.Children.Add(_searchBox);
        Grid.SetRow(_listCount, 2); root.Children.Add(_listCount);
        Grid.SetRow(scroll,     3); root.Children.Add(scroll);

        return new Border { Child = root };
    }

    private void RenderList()
    {
        _listPanel.Children.Clear();
        _selectedCard = null;
        string query = _searchBox.Text?.Trim().ToLower() ?? "";

        if (_category == "시약") RenderReagentList(query);
        else                     RenderGlassList(query);
    }

    // ── 시약 목록 ────────────────────────────────────────────────────────────
    private void RenderReagentList(string query)
    {
        var all = ReagentService.GetAll();
        var filtered = string.IsNullOrEmpty(query)
            ? all
            : all.Where(r =>
                r.품목명.ToLower().Contains(query) ||
                r.영문명.ToLower().Contains(query) ||
                r.화학식.ToLower().Contains(query) ||
                r.CAS번호.ToLower().Contains(query) ||
                r.ITEM_NO.ToLower().Contains(query)).ToList();

        _listCount.Text = $"시약  {filtered.Count} / {all.Count}건";
        foreach (var item in filtered)
            _listPanel.Children.Add(MakeReagentCard(item));
    }

    private Border MakeReagentCard(Reagent item)
    {
        string color = GetChosungColor(item.품목명);

        string badgeText = string.IsNullOrEmpty(item.화학식)
            ? item.위험등급
            : TruncateFormula(item.화학식);

        var badge = new Border
        {
            Background      = new SolidColorBrush(Color.Parse(color)),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(5, 2),
            Margin          = new Thickness(0, 0, 8, 0),
            MinWidth        = 44,
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text          = badgeText,
                FontSize      = AppTheme.FontXS,
                FontFamily    = FontM,
                Foreground    = Brushes.White,
                TextAlignment = TextAlignment.Center,
                TextWrapping  = TextWrapping.NoWrap,
            },
        };

        var nameStack = new StackPanel { Spacing = 1 };
        nameStack.Children.Add(new TextBlock
        {
            Text = item.품목명, FontSize = AppTheme.FontMD, FontFamily = FontM,
            Foreground = AppTheme.FgPrimary,
            TextTrimming = TextTrimming.CharacterEllipsis,
        });

        var subs = new List<string>();
        if (!string.IsNullOrEmpty(item.규격))   subs.Add(item.규격);
        if (!string.IsNullOrEmpty(item.제조사)) subs.Add(item.제조사);
        if (subs.Count > 0)
            nameStack.Children.Add(new TextBlock
            {
                Text = string.Join("  ", subs), FontSize = AppTheme.FontXS, FontFamily = FontR,
                Foreground = AppTheme.FgDimmed,
                TextTrimming = TextTrimming.CharacterEllipsis,
            });

        var qtyBlock = new TextBlock
        {
            Text = item.재고량 > 0 ? $"재고 {item.재고량}" : "",
            FontSize = AppTheme.FontXS, FontFamily = FontR,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 0, 0),
        };

        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("56,*,Auto") };
        row.Children.Add(badge);
        row.Children.Add(nameStack); nameStack.SetValue(Grid.ColumnProperty, 1);
        row.Children.Add(qtyBlock);  qtyBlock.SetValue(Grid.ColumnProperty, 2);

        string darkBg = DarkenColor(color, 0.85);
        var card = new Border
        {
            Background      = new SolidColorBrush(Color.Parse(darkBg)),
            BorderBrush     = new SolidColorBrush(Color.Parse(color + "66")),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(5),
            Padding         = new Thickness(8, 6),
            Margin          = new Thickness(0, 1),
            Child           = row,
            Cursor          = new Cursor(StandardCursorType.Hand),
        };
        card.Tag = color;

        var captured = item;
        card.AddHandler(InputElement.PointerPressedEvent, (object? _, PointerPressedEventArgs ev) =>
        {
            if (ev.GetCurrentPoint(card).Properties.IsLeftButtonPressed)
            {
                HighlightCard(card, color);
                LoadReagentToForm(captured);
                UpdateUsagePanel(captured);
                UpdateAnalytePanel(captured);
            }
        }, RoutingStrategies.Tunnel);

        return card;
    }

    // ── 초자 목록 ────────────────────────────────────────────────────────────
    private void RenderGlassList(string query)
    {
        var all = GlasswareService.GetAll();
        var filtered = string.IsNullOrEmpty(query)
            ? all
            : all.Where(g =>
                g.품목명.ToLower().Contains(query) ||
                g.규격.ToLower().Contains(query)   ||
                g.용도.ToLower().Contains(query)).ToList();

        _listCount.Text = $"초자  {filtered.Count} / {all.Count}건";
        foreach (var item in filtered)
            _listPanel.Children.Add(MakeGlassCard(item));
    }

    private Border MakeGlassCard(Glassware item)
    {
        string color = GetChosungColor(item.품목명);
        var (statusBg, statusFg) = item.상태 switch
        {
            "파손" => ("#665500", "#ffee00"),
            "폐기" => ("#444444", "#aaaaaa"),
            _      => (color, "#ffffff"),
        };

        var badge = new Border
        {
            Background      = new SolidColorBrush(Color.Parse(statusBg)),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(5, 2),
            Margin          = new Thickness(0, 0, 8, 0),
            MinWidth        = 44,
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = item.상태, FontSize = AppTheme.FontXS, FontFamily = FontM,
                Foreground = new SolidColorBrush(Color.Parse(statusFg)),
                TextAlignment = TextAlignment.Center,
            },
        };

        var nameStack = new StackPanel { Spacing = 1 };
        nameStack.Children.Add(new TextBlock
        {
            Text = item.품목명, FontSize = AppTheme.FontMD, FontFamily = FontM,
            Foreground = AppTheme.FgPrimary,
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        if (!string.IsNullOrEmpty(item.규격))
            nameStack.Children.Add(new TextBlock
            {
                Text = $"{item.규격}  {item.재질}", FontSize = AppTheme.FontXS, FontFamily = FontR,
                Foreground = AppTheme.FgDimmed,
            });

        var qtyBlock = new TextBlock
        {
            Text = $"×{item.수량}", FontSize = AppTheme.FontSM, FontFamily = FontR,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 0, 0),
        };

        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("56,*,Auto") };
        row.Children.Add(badge);
        row.Children.Add(nameStack); nameStack.SetValue(Grid.ColumnProperty, 1);
        row.Children.Add(qtyBlock);  qtyBlock.SetValue(Grid.ColumnProperty, 2);

        string darkBg = DarkenColor(color, 0.85);
        var card = new Border
        {
            Background      = new SolidColorBrush(Color.Parse(darkBg)),
            BorderBrush     = new SolidColorBrush(Color.Parse(color + "66")),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(5),
            Padding         = new Thickness(8, 6),
            Margin          = new Thickness(0, 1),
            Child           = row,
            Cursor          = new Cursor(StandardCursorType.Hand),
        };
        card.Tag = color;

        var captured = item;
        card.AddHandler(InputElement.PointerPressedEvent, (object? _, PointerPressedEventArgs ev) =>
        {
            if (ev.GetCurrentPoint(card).Properties.IsLeftButtonPressed)
            {
                HighlightCard(card, color);
                LoadGlassToForm(captured);
            }
        }, RoutingStrategies.Tunnel);

        return card;
    }

    private void HighlightCard(Border card, string color)
    {
        if (_selectedCard != null && _selectedCard != card)
        {
            string prev = _selectedCard.Tag as string ?? "#555";
            _selectedCard.BorderBrush     = new SolidColorBrush(Color.Parse(prev + "66"));
            _selectedCard.BorderThickness = new Thickness(1);
        }
        _selectedCard = card;
        card.BorderBrush     = new SolidColorBrush(Color.Parse(color));
        card.BorderThickness = new Thickness(1.5);
    }

    // =========================================================================
    // Show2 — 사용량 현황 패널
    // =========================================================================
    private Control BuildUsagePanel()
    {
        var scroll = new ScrollViewer
        {
            Content = _usagePanel,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };
        return new Border { Padding = new Thickness(14, 12), Child = scroll };
    }

    private void ShowUsagePlaceholder()
    {
        _usagePanel.Children.Clear();
        _usagePanel.Children.Add(new TextBlock
        {
            Text = "시약을 선택하면\n사용량 현황이 표시됩니다",
            FontSize = AppTheme.FontBase, FontFamily = FontR,
            Foreground = AppTheme.FgDimmed,
            TextAlignment = TextAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 30, 0, 0),
        });
    }

    private void UpdateUsagePanel(Reagent item)
    {
        _usagePanel.Children.Clear();

        // 제목
        _usagePanel.Children.Add(new TextBlock
        {
            Text = "사용량 현황", FontSize = AppTheme.FontBase, FontWeight = FontWeight.Bold,
            FontFamily = FontM, Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 2),
        });
        _usagePanel.Children.Add(new TextBlock
        {
            Text = item.품목명, FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            FontFamily = FontM,
            Foreground = new SolidColorBrush(Color.Parse(GetChosungColor(item.품목명))),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 4),
        });
        if (!string.IsNullOrEmpty(item.영문명))
            _usagePanel.Children.Add(new TextBlock
            {
                Text = item.영문명, FontSize = AppTheme.FontSM, FontFamily = FontR,
                Foreground = new SolidColorBrush(Color.Parse("#667788")),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 14),
            });
        else
            _usagePanel.Children.Add(new Border { Height = 10 }); // spacer

        int max = item.최대적정보유량;

        _usagePanel.Children.Add(MakeProgressRow("당월 사용량", item.당월사용량, max, "#4488ff"));
        _usagePanel.Children.Add(MakeProgressRow("전월 사용량", item.전월사용량, max, "#44aacc"));
        _usagePanel.Children.Add(MakeProgressRow("재고량",      item.재고량,     max, "#44cc88"));
        _usagePanel.Children.Add(MakeProgressRow("적정 사용량", item.적정사용량, max, "#ccaa44"));

        if (max == 0)
            _usagePanel.Children.Add(new Border
            {
                Margin = new Thickness(0, 6, 0, 0),
                Padding = new Thickness(8, 6),
                Background = new SolidColorBrush(Color.Parse("#18182a")),
                CornerRadius = new CornerRadius(4),
                Child = new TextBlock
                {
                    Text = "💡 Show3 폼에서 최대 적정보유량을 입력하면\n진행률이 표시됩니다",
                    FontSize = AppTheme.FontXS, FontFamily = FontR,
                    Foreground = AppTheme.FgDimmed,
                    TextWrapping = TextWrapping.Wrap,
                },
            });
    }

    private static Control MakeProgressRow(string label, int value, int max, string barColor)
    {
        bool over = max > 0 && value > max;
        double pct = max > 0 ? Math.Clamp((double)value / max, 0, 1.0) : 0;
        string fillColor = over ? "#cc4444" : barColor;

        // 헤더: [레이블] ... [value / max]
        var header = new Grid { ColumnDefinitions = new ColumnDefinitions("90,*,Auto"), Margin = new Thickness(0, 0, 0, 3) };
        header.Children.Add(new TextBlock
        {
            Text = label, FontSize = AppTheme.FontSM, FontFamily = FontR,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        });
        var valTb = new TextBlock
        {
            Text = max > 0 ? $"{value:N0} / {max:N0}" : $"{value:N0}",
            FontSize = AppTheme.FontSM, FontFamily = FontR,
            Foreground = new SolidColorBrush(Color.Parse(over ? "#ff9999" : "#777788")),
            VerticalAlignment = VerticalAlignment.Center,
        };
        valTb.SetValue(Grid.ColumnProperty, 2);
        header.Children.Add(valTb);

        // 프로그레스 바
        double fill = pct;
        double empty = 1.0 - pct;

        var barGrid = new Grid();
        barGrid.ColumnDefinitions.Add(new ColumnDefinition(fill * 100 + 0.01, GridUnitType.Star));
        barGrid.ColumnDefinitions.Add(new ColumnDefinition(empty * 100 + 0.01, GridUnitType.Star));

        var fillBorder = new Border
        {
            Background   = new SolidColorBrush(Color.Parse(fillColor)),
            Height       = 10,
            CornerRadius = pct >= 1.0 ? new CornerRadius(3) : new CornerRadius(3, 0, 0, 3),
        };
        var emptyBorder = new Border
        {
            Background   = AppTheme.BgPrimary,
            Height       = 10,
            CornerRadius = pct <= 0 ? new CornerRadius(3) : new CornerRadius(0, 3, 3, 0),
        };
        fillBorder.SetValue(Grid.ColumnProperty, 0);
        emptyBorder.SetValue(Grid.ColumnProperty, 1);
        barGrid.Children.Add(fillBorder);
        barGrid.Children.Add(emptyBorder);

        var track = new Border { CornerRadius = new CornerRadius(3), ClipToBounds = true, Child = barGrid };

        var sp = new StackPanel { Spacing = 2, Margin = new Thickness(0, 0, 0, 10) };
        sp.Children.Add(header);
        sp.Children.Add(track);
        return sp;
    }

    // =========================================================================
    // Show4 — 분석항목 연결 패널
    // =========================================================================
    private Control BuildAnalytePanel()
    {
        var scroll = new ScrollViewer
        {
            Content = _analytePanel,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };
        return new Border { Padding = new Thickness(14, 12), Child = scroll };
    }

    private void ShowAnalytePlaceholder()
    {
        _analytePanel.Children.Clear();
        _analytePanel.Children.Add(new TextBlock
        {
            Text = "시약을 선택하면\n연결된 분석항목이 표시됩니다",
            FontSize = AppTheme.FontBase, FontFamily = FontR,
            Foreground = AppTheme.FgDimmed,
            TextAlignment = TextAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 30, 0, 0),
        });
    }

    private void UpdateAnalytePanel(Reagent item)
    {
        _analytePanel.Children.Clear();

        _analytePanel.Children.Add(new TextBlock
        {
            Text = "분석항목 연결", FontSize = AppTheme.FontBase, FontWeight = FontWeight.Bold,
            FontFamily = FontM, Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 4),
        });
        _analytePanel.Children.Add(new TextBlock
        {
            Text = $"{item.품목명} — 시료당 소요량 기반 적정사용량 산정",
            FontSize = AppTheme.FontXS, FontFamily = FontR,
            Foreground = AppTheme.FgDimmed,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10),
        });

        // 연결된 분석항목 목록
        var links = ReagentAnalyteService.GetByReagentId(item.Id);
        int totalEstimate = 0;

        if (links.Count > 0)
        {
            // 헤더
            var headerGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,60,60,60,30"),
                Margin = new Thickness(0, 0, 0, 4),
            };
            AddColText(headerGrid, "분석항목", 0, AppTheme.FgDimmed, FontWeight.SemiBold);
            AddColText(headerGrid, "소요량", 1, AppTheme.FgDimmed, FontWeight.SemiBold);
            AddColText(headerGrid, "의뢰수", 2, AppTheme.FgDimmed, FontWeight.SemiBold);
            AddColText(headerGrid, "산정량", 3, AppTheme.FgDimmed, FontWeight.SemiBold);
            _analytePanel.Children.Add(headerGrid);

            foreach (var link in links)
            {
                int monthlyCount = ETA.Services.SERVICE1.AnalysisRequestService.GetMonthlyAnalyteCount(link.분석항목);
                int estimateEA  = (int)Math.Ceiling(link.시료당소요량 * monthlyCount);
                int estimateDay = (int)Math.Ceiling(link.일일소요량 * 22);
                int estimate    = estimateEA + estimateDay;
                totalEstimate += estimate;

                var rowGrid = new Grid
                {
                    ColumnDefinitions = new ColumnDefinitions("*,52,52,50,58,28"),
                    Margin = new Thickness(0, 1, 0, 1),
                };

                // 항목명
                AddColText(rowGrid, link.분석항목, 0, AppTheme.FgPrimary, FontWeight.Normal);
                // g/EA
                AddColText(rowGrid, link.시료당소요량 > 0 ? link.시료당소요량.ToString("F2") : "-", 1, AppTheme.FgMuted, FontWeight.Normal);
                // g/Day
                AddColText(rowGrid, link.일일소요량 > 0 ? link.일일소요량.ToString("F2") : "-", 2, AppTheme.FgMuted, FontWeight.Normal);
                // 이번달 의뢰 건수
                AddColText(rowGrid, monthlyCount.ToString(), 3,
                    monthlyCount > 0 ? new SolidColorBrush(Color.Parse("#4488ff")) : AppTheme.FgMuted,
                    FontWeight.Normal);
                // 적정사용량 산정
                AddColText(rowGrid, estimate.ToString(), 4,
                    new SolidColorBrush(Color.Parse("#ccaa44")), FontWeight.SemiBold);

                // 삭제 버튼
                var btnDel = new Button
                {
                    Content = "✕", FontSize = AppTheme.FontXS,
                    Padding = new Thickness(4, 2),
                    Background = Brushes.Transparent,
                    Foreground = new SolidColorBrush(Color.Parse("#884444")),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                };
                var capturedLink = link;
                var capturedItem = item;
                btnDel.Click += (_, _) =>
                {
                    ReagentAnalyteService.Delete(capturedLink.Id);
                    UpdateAnalytePanel(capturedItem);
                };
                Grid.SetColumn(btnDel, 5);
                rowGrid.Children.Add(btnDel);

                var rowBorder = new Border
                {
                    Child = rowGrid,
                    Padding = new Thickness(4, 4),
                    CornerRadius = new CornerRadius(3),
                    Background = new SolidColorBrush(Color.Parse("#141428")),
                    Margin = new Thickness(0, 1),
                };
                _analytePanel.Children.Add(rowBorder);
            }
        }
        else
        {
            _analytePanel.Children.Add(new TextBlock
            {
                Text = "연결된 분석항목이 없습니다.\nShow3 폼에서 분석항목을 추가하세요.",
                FontSize = AppTheme.FontSM, FontFamily = FontR,
                Foreground = AppTheme.FgDimmed,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 10, 0, 0),
            });
        }

        // 합계 표시
        if (totalEstimate > 0)
        {
            _analytePanel.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(0, 8) });
            var sumGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,52,52,50,58,28"),
            };
            AddColText(sumGrid, "합계 (적정사용량)", 0, AppTheme.FgPrimary, FontWeight.Bold);
            AddColText(sumGrid, totalEstimate.ToString(), 4,
                new SolidColorBrush(Color.Parse("#ffaa44")), FontWeight.Bold);
            _analytePanel.Children.Add(sumGrid);

            // 재고 대비 안내
            string status;
            IBrush statusColor;
            if (item.재고량 >= totalEstimate)
            {
                status = $"재고({item.재고량}) >= 적정({totalEstimate}) — 충분";
                statusColor = new SolidColorBrush(Color.Parse("#44cc88"));
            }
            else
            {
                status = $"재고({item.재고량}) < 적정({totalEstimate}) — 부족 ({totalEstimate - item.재고량} 추가 필요)";
                statusColor = new SolidColorBrush(Color.Parse("#ff6666"));
            }
            _analytePanel.Children.Add(new TextBlock
            {
                Text = status, FontSize = AppTheme.FontSM, FontFamily = FontM,
                Foreground = statusColor,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 6, 0, 0),
            });
        }
    }

    private static void AddColText(Grid grid, string text, int col, IBrush fg, FontWeight weight)
    {
        var tb = new TextBlock
        {
            Text = text, FontSize = AppTheme.FontSM, FontFamily = FontR,
            Foreground = fg, FontWeight = weight,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(tb, col);
        grid.Children.Add(tb);
    }

    // =========================================================================
    // Show3 — 폼 패널
    // =========================================================================
    private Control BuildFormPanel()
    {
        var scroll = new ScrollViewer
        {
            Content = _formPanel,
            VerticalScrollBarVisibility   = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };
        return new Border { Padding = new Thickness(10), Child = scroll };
    }

    // ── 시약 폼 ──────────────────────────────────────────────────────────────
    private void LoadReagentToForm(Reagent item)   => BuildReagentForm(item);
    private void ShowReagentNewForm()              => BuildReagentForm(null);

    private void BuildReagentForm(Reagent? item)
    {
        bool isNew = item == null;
        _editingReagent = item;
        _editingGlass   = null;
        _formPanel.Children.Clear();

        // 제목
        string titleText  = isNew ? "➕ 시약 신규 등록" : $"✏️  {item!.품목명}";
        string titleColor = isNew ? "#88ccaa" : GetChosungColor(item!.품목명);
        _formPanel.Children.Add(new TextBlock
        {
            Text = titleText, FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            FontFamily = FontM, Foreground = new SolidColorBrush(Color.Parse(titleColor)),
            Margin = new Thickness(0, 0, 0, 10), TextWrapping = TextWrapping.Wrap,
        });

        string v(string f)
        {
            if (isNew) return f switch
            {
                "위험등급" => "일반", "상태" => "정상",
                "재고량" or "적정사용량" or "최대적정보유량" or "당월사용량" or "전월사용량" => "0",
                _ => "",
            };
            return f switch
            {
                "ITEM_NO"       => item!.ITEM_NO,
                "품목명"        => item!.품목명,
                "영문명"        => item!.영문명,
                "화학식"        => item!.화학식,
                "CAS번호"       => item!.CAS번호,
                "규격"          => item!.규격,
                "단위"          => item!.단위,
                "위험등급"      => item!.위험등급,
                "GHS"           => item!.GHS,
                "보관조건"      => item!.보관조건,
                "제조사"        => item!.제조사,
                "만료일"        => item!.만료일,
                "비고"          => item!.비고,
                "상태"          => item!.상태,
                "재고량"        => item!.재고량.ToString(),
                "적정사용량"    => item!.적정사용량.ToString(),
                "최대적정보유량" => item!.최대적정보유량.ToString(),
                "당월사용량"    => item!.당월사용량.ToString(),
                "전월사용량"    => item!.전월사용량.ToString(),
                _               => "",
            };
        }

        // ── 기본 정보 (4열 그리드) ───────────────────────────────────────
        AddSectionHeader("기본 정보");
        _formPanel.Children.Add(BuildFieldGrid(new (string, Control, string)[]
        {
            ("ITEM NO.", _rItemNo,  v("ITEM_NO")),
            ("국문명 *",  _rName,   v("품목명")),
            ("영문명",    _rEng,    v("영문명")),
            ("화학식",    _rFormula, v("화학식")),
            ("CAS NO.",  _rCas,    v("CAS번호")),
            ("규격",      _rSpec,   v("규격")),
            ("단위",      _rUnit,   v("단위")),
            ("등급",      _rRisk,   v("위험등급")),
            ("GHS",       _rGhs,    v("GHS")),
            ("보관조건",  _rStorage, v("보관조건")),
            ("제조사",    _rMaker,  v("제조사")),
            ("만료일",    _rExpiry, v("만료일")),
            ("비고",      _rNote,   v("비고")),
            ("상태",      _rStatus, v("상태")),
        }));

        // ── 재고 / 사용량 (4열 그리드) ───────────────────────────────────
        AddSectionHeader("재고 / 사용량");
        _formPanel.Children.Add(BuildFieldGrid(new (string, Control, string)[]
        {
            ("재고량",         _rStock,    v("재고량")),
            ("적정 사용량",    _rOptUse,   v("적정사용량")),
            ("최대 적정보유량", _rMaxStock, v("최대적정보유량")),
            ("당월 사용량",    _rCurrUse,  v("당월사용량")),
            ("전월 사용량",    _rPrevUse,  v("전월사용량")),
        }));

        // ── 분석항목 연결 (기존 시약 수정 시만) ──────────────────────────
        if (!isNew)
        {
            AddSectionHeader("분석항목 연결");

            // ── 검색 입력 ──
            var selectedName    = new string[] { "" };
            var ignoreTextChg   = new bool[] { false };
            List<(string Full, string Short)> analyteList = new();
            try { analyteList = ETA.Services.SERVICE1.AnalysisRequestService.GetOrderedAnalytes(); }
            catch { }

            var tbSearch = MakeTxb("분석항목 검색 또는 직접 입력...");

            var suggestSource = new System.Collections.ObjectModel.ObservableCollection<string>();
            var suggestFullNames = new List<string>();

            var lbSuggest = new ListBox
            {
                MaxHeight = 160,
                Background = new SolidColorBrush(Color.Parse("#1a1a2e")),
                FontFamily = FontR, FontSize = AppTheme.FontSM,
                ItemsSource = suggestSource,
            };
            var suggestBorder = new Border
            {
                Child = lbSuggest,
                IsVisible = false,
                Background = new SolidColorBrush(Color.Parse("#1a1a2e")),
                BorderBrush = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 0, 4),
            };

            tbSearch.TextChanged += (_, _) =>
            {
                if (ignoreTextChg[0]) return;
                var q = tbSearch.Text?.Trim() ?? "";
                selectedName[0] = q;
                if (q.Length == 0) { suggestBorder.IsVisible = false; return; }
                var filtered = analyteList
                    .Where(a => a.Full.Contains(q, StringComparison.OrdinalIgnoreCase)
                             || a.Short.Contains(q, StringComparison.OrdinalIgnoreCase))
                    .Take(20).ToList();
                suggestSource.Clear();
                suggestFullNames.Clear();
                foreach (var (full, sh) in filtered)
                {
                    suggestSource.Add(full == sh ? full : $"{sh} ({full})");
                    suggestFullNames.Add(full);
                }
                suggestBorder.IsVisible = suggestSource.Count > 0;
            };

            lbSuggest.SelectionChanged += (_, _) =>
            {
                int idx = lbSuggest.SelectedIndex;
                if (idx < 0 || idx >= suggestFullNames.Count) return;
                selectedName[0] = suggestFullNames[idx];
                ignoreTextChg[0] = true;
                tbSearch.Text = suggestSource[idx];
                ignoreTextChg[0] = false;
                suggestBorder.IsVisible = false;
            };

            // ── g/EA + g/Day 입력 ──
            var tbAmountEA  = new TextBox { Text = "0", Width = 70, FontFamily = FontR, FontSize = AppTheme.FontSM, Watermark = "0" };
            var tbAmountDay = new TextBox { Text = "0", Width = 70, FontFamily = FontR, FontSize = AppTheme.FontSM, Watermark = "0" };
            var unitTxt = string.IsNullOrEmpty(item!.단위) ? item.규격 : item.단위;
            TextBlock MiniLbl(string t) => new TextBlock
            {
                Text = t, FontSize = AppTheme.FontSM, FontFamily = FontR,
                Foreground = AppTheme.FgDimmed, VerticalAlignment = VerticalAlignment.Center,
            };

            var amountRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            amountRow.Children.Add(MiniLbl("g/EA:"));
            amountRow.Children.Add(tbAmountEA);
            amountRow.Children.Add(MiniLbl("  g/Day:"));
            amountRow.Children.Add(tbAmountDay);
            if (!string.IsNullOrWhiteSpace(unitTxt))
                amountRow.Children.Add(MiniLbl($"  ({unitTxt})"));

            var btnAdd = MakeBtn("+ 추가", "#1a2a3a", "#88aaff");
            var capturedItem = item;
            btnAdd.Click += (_, _) =>
            {
                string analyteName = selectedName[0].Trim();
                if (string.IsNullOrEmpty(analyteName)) return;
                double.TryParse(tbAmountEA.Text?.Trim(),  out double ea);
                double.TryParse(tbAmountDay.Text?.Trim(), out double day);

                var existing = ReagentAnalyteService.GetByReagentId(capturedItem.Id);
                if (existing.Any(e => e.분석항목 == analyteName))
                {
                    var existingLink = existing.First(e => e.분석항목 == analyteName);
                    existingLink.시료당소요량 = ea;
                    existingLink.일일소요량   = day;
                    ReagentAnalyteService.Update(existingLink);
                }
                else
                {
                    ReagentAnalyteService.Insert(new ReagentAnalyte
                    {
                        시약Id = capturedItem.Id,
                        분석항목 = analyteName,
                        시료당소요량 = ea,
                        일일소요량   = day,
                    });
                }
                ignoreTextChg[0] = true;
                tbSearch.Text = "";
                ignoreTextChg[0] = false;
                selectedName[0] = "";
                tbAmountEA.Text  = "0";
                tbAmountDay.Text = "0";
                suggestBorder.IsVisible = false;
                UpdateAnalytePanel(capturedItem);
            };
            amountRow.Children.Add(btnAdd);

            _formPanel.Children.Add(tbSearch);
            _formPanel.Children.Add(suggestBorder);
            _formPanel.Children.Add(amountRow);
        }

        // ── 버튼 ─────────────────────────────────────────────────────────
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(0, 10, 0, 0) };
        var btnSave = MakeBtn("💾 저장", "#1a3a2a", "#88ffaa");
        btnSave.Click += (_, _) => SaveReagent();
        btnRow.Children.Add(btnSave);

        if (!isNew)
        {
            var btnDel = MakeBtn("🗑 삭제", "#3a1a1a", "#ff8888");
            btnDel.Click += (_, _) =>
            {
                ReagentService.Delete(item!.Id);
                _editingReagent = null;
                RenderList();
                ClearDetail();
            };
            btnRow.Children.Add(btnDel);
        }
        _formPanel.Children.Add(btnRow);
    }

    /// <summary>
    /// (label, control, value) 배열을 받아 4열 그리드로 배치
    /// </summary>
    private Grid BuildFieldGrid(IEnumerable<(string Label, Control Ctrl, string Value)> fields)
    {
        const int Cols = 4;
        var list = fields.ToList();
        int rows = (list.Count + Cols - 1) / Cols;

        var grid = new Grid { Margin = new Thickness(0, 0, 0, 2) };
        for (int c = 0; c < Cols; c++)
            grid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        for (int r = 0; r < rows; r++)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        for (int i = 0; i < list.Count; i++)
        {
            var (label, ctrl, value) = list[i];
            int row = i / Cols, col = i % Cols;

            // 기존 부모에서 분리
            if (ctrl.Parent is Avalonia.Controls.Panel p) p.Children.Remove(ctrl);

            if (ctrl is TextBox tb)      tb.Text = value;
            else if (ctrl is ComboBox cb) SelectCombo(cb, value);

            var cell = new StackPanel
            {
                Spacing = 2,
                Margin  = new Thickness(col == 0 ? 0 : 4, 0, col == Cols - 1 ? 0 : 4, 8),
            };
            cell.Children.Add(new TextBlock
            {
                Text = label, FontSize = AppTheme.FontXS, FontFamily = FontR,
                Foreground = AppTheme.FgDimmed,
            });
            cell.Children.Add(ctrl);

            Grid.SetRow(cell,    row);
            Grid.SetColumn(cell, col);
            grid.Children.Add(cell);
        }
        return grid;
    }

    private void SaveReagent()
    {
        if (string.IsNullOrWhiteSpace(_rName.Text)) return;

        var item = _editingReagent ?? new Reagent { 등록일 = DateTime.Today.ToString("yyyy-MM-dd") };
        item.ITEM_NO      = _rItemNo.Text?.Trim()   ?? "";
        item.품목명       = _rName.Text.Trim();
        item.영문명       = _rEng.Text?.Trim()      ?? "";
        item.화학식       = _rFormula.Text?.Trim()  ?? "";
        item.CAS번호      = _rCas.Text?.Trim()      ?? "";
        item.규격         = _rSpec.Text?.Trim()     ?? "";
        item.단위         = _rUnit.Text?.Trim()     ?? "";
        item.위험등급     = ComboVal(_rRisk)        ?? "일반";
        item.GHS          = _rGhs.Text?.Trim()      ?? "";
        item.보관조건     = _rStorage.Text?.Trim()  ?? "";
        item.제조사       = _rMaker.Text?.Trim()    ?? "";
        item.만료일       = _rExpiry.Text?.Trim()   ?? "";
        item.비고         = _rNote.Text?.Trim()     ?? "";
        item.상태         = ComboVal(_rStatus)      ?? "정상";
        item.재고량       = IntVal(_rStock);
        item.적정사용량    = IntVal(_rOptUse);
        item.최대적정보유량 = IntVal(_rMaxStock);
        item.당월사용량    = IntVal(_rCurrUse);
        item.전월사용량    = IntVal(_rPrevUse);

        if (_editingReagent == null) ReagentService.Insert(item);
        else                         ReagentService.Update(item);

        UpdateUsagePanel(item);
        RenderList();
    }

    // ── 초자 폼 ──────────────────────────────────────────────────────────────
    private void LoadGlassToForm(Glassware item)
    {
        _editingGlass   = item;
        _editingReagent = null;
        _formPanel.Children.Clear();

        _formPanel.Children.Add(new TextBlock
        {
            Text = $"✏️ 초자 수정 — {item.품목명}", FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            FontFamily = FontM, Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 10),
        });

        AddField("품목명", _gName,     item.품목명);
        AddField("용도",   _gUse,      item.용도);
        AddField("규격",   _gSpec,     item.규격);
        AddField("재질",   _gMaterial, item.재질);
        AddField("수량",   _gQty,      item.수량.ToString());
        AddField("단가",   _gPrice,    item.단가 > 0 ? item.단가.ToString("F0") : "");
        AddField("비고",   _gNote,     item.비고);
        AddField("상태",   _gStatus,   item.상태);

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(0, 10, 0, 0) };
        var btnSave = MakeBtn("💾 저장", "#1a3a2a", "#88ffaa");
        var btnDel  = MakeBtn("🗑 삭제",  "#3a1a1a", "#ff8888");
        btnSave.Click += (_, _) => SaveGlass();
        btnDel.Click  += (_, _) => { GlasswareService.Delete(item.Id); _editingGlass = null; RenderList(); ClearDetail(); };
        btnRow.Children.Add(btnSave);
        btnRow.Children.Add(btnDel);
        _formPanel.Children.Add(btnRow);
    }

    private void ShowGlassNewForm()
    {
        _editingGlass   = null;
        _editingReagent = null;
        _formPanel.Children.Clear();

        _formPanel.Children.Add(new TextBlock
        {
            Text = "➕ 초자 신규 등록", FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            FontFamily = FontM, Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 10),
        });

        AddField("품목명", _gName,     "");
        AddField("용도",   _gUse,      "");
        AddField("규격",   _gSpec,     "");
        AddField("재질",   _gMaterial, "유리");
        AddField("수량",   _gQty,      "0");
        AddField("단가",   _gPrice,    "");
        AddField("비고",   _gNote,     "");
        AddField("상태",   _gStatus,   "정상");

        var btnSave = MakeBtn("💾 저장", "#1a3a2a", "#88ffaa");
        btnSave.Click += (_, _) => SaveGlass();
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 10, 0, 0) };
        btnRow.Children.Add(btnSave);
        _formPanel.Children.Add(btnRow);
    }

    private void SaveGlass()
    {
        if (string.IsNullOrWhiteSpace(_gName.Text)) return;
        var item = _editingGlass ?? new Glassware { 등록일 = DateTime.Today.ToString("yyyy-MM-dd") };
        item.품목명 = _gName.Text.Trim();
        item.용도   = _gUse.Text?.Trim()   ?? "";
        item.규격   = _gSpec.Text?.Trim()  ?? "";
        item.재질   = ComboVal(_gMaterial) ?? "유리";
        item.수량   = IntVal(_gQty);
        item.단가   = decimal.TryParse(_gPrice.Text, out decimal p) ? p : 0m;
        item.비고   = _gNote.Text?.Trim()  ?? "";
        item.상태   = ComboVal(_gStatus)   ?? "정상";

        if (_editingGlass == null) GlasswareService.Insert(item);
        else                       GlasswareService.Update(item);

        RenderList();
        if (_editingGlass == null) ShowGlassNewForm();
    }

    private void ClearDetail()
    {
        _formPanel.Children.Clear();
        _editingReagent = null;
        _editingGlass   = null;
        ShowUsagePlaceholder();

        if (_category == "시약")      ShowReagentNewForm();
        else if (_category == "초자") ShowGlassNewForm();
    }

    private void SwitchCategory(string cat)
    {
        _category = cat;
        _searchBox.Text = "";
        RenderList();

        ShowUsagePlaceholder();
        if (cat == "초자") ShowGlassNewForm();
        else               ShowReagentNewForm();
    }

    private void AddSectionHeader(string text)
    {
        _formPanel.Children.Add(new Border
        {
            Margin = new Thickness(0, 6, 0, 4),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = new TextBlock
            {
                Text = text, FontSize = AppTheme.FontSM, FontFamily = FontM,
                Foreground = new SolidColorBrush(Color.Parse("#667788")),
                Margin = new Thickness(0, 0, 0, 3),
            },
        });
    }

    private void AddField(string label, Control ctrl, string value)
    {
        if (ctrl.Parent is Avalonia.Controls.Panel p) p.Children.Remove(ctrl);

        if (ctrl is TextBox tb) tb.Text = value;
        else if (ctrl is ComboBox cb) SelectCombo(cb, value);

        var sp = new StackPanel { Spacing = 2, Margin = new Thickness(0, 0, 0, 6) };
        sp.Children.Add(new TextBlock
        {
            Text = label, FontSize = AppTheme.FontXS, FontFamily = FontR,
            Foreground = AppTheme.FgDimmed,
        });
        sp.Children.Add(ctrl);
        _formPanel.Children.Add(sp);
    }

    // =========================================================================
    // Excel → DB 자동 가져오기 (DB가 비어 있을 때만)
    // =========================================================================
    private static void ImportReagentsIfEmpty()
    {
        if (ReagentService.Count() > 0) return;

        var reagents = RiskExcelService.GetReagents();
        var history  = RiskExcelService.GetReagentHistory();

        for (int i = 0; i < reagents.Count; i++)
        {
            var r = reagents[i];

            // 해당 시약의 최근 이력에서 재고/사용중 가져오기
            var latest = history
                .Where(h => h.ReagentIndex == i)
                .OrderByDescending(h => h.일자)
                .FirstOrDefault();

            var item = new Reagent
            {
                ITEM_NO = r.ITEM_NO,
                품목명  = r.국문명,
                영문명  = r.영문명,
                화학식  = r.화학식,
                CAS번호 = r.CAS번호,
                규격    = $"{r.규격}{r.단위}",
                단위    = r.단위,
                위험등급 = r.등급,
                재고량  = latest?.재고 ?? 0,
                등록일  = DateTime.Today.ToString("yyyy-MM-dd"),
            };
            ReagentService.Insert(item);

            // 이력 데이터 가져오기 (입고 또는 출고가 있는 날만)
            var itemHistory = history
                .Where(h => h.ReagentIndex == i && (h.입고 > 0 || h.출고 > 0))
                .ToList();
            foreach (var h in itemHistory)
            {
                ReagentHistoryService.Insert(new ReagentHistory
                {
                    시약Id = item.Id,
                    일자   = h.일자.ToString("yyyy-MM-dd"),
                    입고   = h.입고,
                    출고   = h.출고,
                    재고   = h.재고,
                    사용중 = h.사용중,
                });
            }
        }
    }

    // =========================================================================
    // 유틸
    // =========================================================================
    private static string GetChosungColor(string name)
    {
        if (string.IsNullOrEmpty(name)) return "#555555";
        char c = name[0];
        if (c >= 0xAC00 && c <= 0xD7A3)
        {
            int idx = (c - 0xAC00) / (21 * 28);
            if (idx < ChosungColors.Length) return ChosungColors[idx];
        }
        return "#555566";
    }

    private static string DarkenColor(string hex, double factor)
    {
        try
        {
            var c = Color.Parse(hex);
            return $"#{(int)(c.R * (1 - factor)):x2}{(int)(c.G * (1 - factor)):x2}{(int)(c.B * (1 - factor)):x2}";
        }
        catch { return "#1a1a1a"; }
    }

    private static string TruncateFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return "?";
        int dot = formula.IndexOf('·');
        if (dot > 0) formula = formula[..dot].Trim();
        return formula.Length > 18 ? formula[..16] + "…" : formula;
    }

    private static string? ComboVal(ComboBox cb)
        => (cb.SelectedItem as ComboBoxItem)?.Content?.ToString();

    private static int IntVal(TextBox tb)
        => int.TryParse(tb.Text, out int v) ? v : 0;

    private static TextBox MakeTxb(string hint) => new()
    {
        Watermark       = hint,
        FontFamily      = FontR,
        FontSize        = AppTheme.FontMD,
        Background      = AppTheme.BgPrimary,
        Foreground      = AppTheme.FgPrimary,
        BorderBrush     = AppTheme.BorderSubtle,
        BorderThickness = new Thickness(1),
        CornerRadius    = new CornerRadius(4),
        Padding         = new Thickness(8, 5),
    };

    private static ComboBox MakeCombo(string[] items)
    {
        var cb = new ComboBox
        {
            FontFamily  = FontR, FontSize = AppTheme.FontMD,
            Background  = AppTheme.BgPrimary,
            Foreground  = AppTheme.FgPrimary,
            BorderBrush = AppTheme.BorderSubtle,
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        foreach (var s in items) cb.Items.Add(new ComboBoxItem { Content = s });
        cb.SelectedIndex = 0;
        return cb;
    }


    private static Button MakeBtn(string text, string bg, string fg) => new()
    {
        Content         = text,
        Height          = 30,
        Padding         = new Thickness(14, 0),
        FontFamily      = FontR,
        FontSize        = AppTheme.FontBase,
        Background      = new SolidColorBrush(Color.Parse(bg)),
        Foreground      = new SolidColorBrush(Color.Parse(fg)),
        BorderThickness = new Thickness(0),
        CornerRadius    = new CornerRadius(4),
        Cursor          = new Cursor(StandardCursorType.Hand),
    };

    private static void SelectCombo(ComboBox cb, string value)
    {
        foreach (var item in cb.Items)
            if (item is ComboBoxItem ci && ci.Content?.ToString() == value)
            { cb.SelectedItem = ci; return; }
        cb.SelectedIndex = 0;
    }
}
