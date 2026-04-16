using Avalonia;
using Avalonia.Animation;
using Avalonia.Animation.Easings;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.ViewModels;
using ETA.Views.Pages;
using ETA.Views.Pages.PAGE1;
using ETA.Views.Pages.PAGE2;
using ETA.Views.Pages.PAGE3;
using ETA.Views.Pages.Common;
using ETA.Views.Pages.PAGE5;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views;

public partial class MainPage : Window
{
    public static string CurrentEmployeeId { get; set; } = "";

    private string _currentMode = "None";
    
    // ── 창 위치/레이아웃 관리 ──────────────────────────────────────────
    private WindowPositionManager? _positionManager;
    private const string LayoutStorageModePrefix = "Mode_";

    // 페이지는 처음 진입 시점에 Lazy 생성
    private AnalysisPage?      _analysisPage;
    private ContractPage?      _contractPage;
    private AgentTreePage?     _agentTreePage;
    private MyTaskPage?        _myTaskPage;
    private WasteCompanyPage?      _wasteCompanyPage;
    private WasteNameReconcilePage?            _wasteNameReconcilePage;
    private WaterQualityNameReconcilePage?     _waterQualityNameReconcilePage;
    private WasteSampleListPage?         _wasteSampleListPage;
    private ProcessingFacilityPage?      _processingFacilityPage;
    private ResultSubmitMeasurePage?     _resultSubmitMeasurePage;
    private TestReportPage?              _resultSubmitMeasureTestReport;
    private WasteAnalysisInputPage?      _wasteAnalysisInputPage;
    private EcotoxicityPage?             _ecotoxicityPage;
    private ResultSubmitErpPage?         _resultSubmitErpPage;
    private AccessPage?                  _accessPage;
    private AiDocClassificationPage?     _aiDocClassificationPage;
    private ParserGeneratorPage?          _parserGeneratorPage;
    private ResultSubmitZero4Page?       _resultSubmitZero4Page;
    private PurchasePage?      _purchasePage;
    private RepairPage?       _repairPage;
    private RiskManagePage?   _riskPage;
    private SchedulePage?     _schedulePage;
    private TestReportPage?         _testReportPage;
    private WasteTestReportPage?    _wasteTestReportPage;
    private ReportsPanel?           _reportsPanel;           // Content4: 출력 보관함
    private DbMigrationPage?           _dbMigrationPage;
    private DbMigrationPointPanel?     _dbMigrationPointPanel;
    private DbMigrationMappingPanel?   _dbMigrationMappingPanel;
    private DbMigrationPage?           _dbMigrationTargetPage;  // Show4: 변경 후 업체 선택
    private string?                    _migrationOldName;
    private string?                    _migrationNewName;

    // ── 견적/의뢰서 전용 4-패널 ──────────────────────────────────────────
    // Content1: 발행내역 트리  Content2: 신규작성 폼
    // Content3: 분석항목 체크  Content4: 계약업체 목록
    private QuotationHistoryPanel? _quotationHistoryPanel;

    // ── 견적발행 전용 패널 ────────────────────────────────────────────────
    private QuotationHistoryPanel?                    _issuingHistoryPanel;
    private StackPanel?                               _issuingChecklistPanel;  // Show4
    private readonly List<ETA.Models.QuotationIssue> _issuingChecklist = [];   // 추가된 항목
    private ScrollViewer?                             _issuingChecklistScroll;
    private QuotationDetailPanel?  _quotationDetailPanel;   // Content2: 세부내역
    private QuotationNewPanel?     _quotationNewPanel;      // Content2: 신규작성
    private QuotationCheckPanel?   _quotationCheckPanel;
    private QuotationPage?         _quotationPage;
    private OrderRequestEditPanel? _orderRequestEditPanel;  // Content2: 의뢰서 편집

    private System.Action? _bt1SaveAction;

    // ── Show2/Show3 영속 래퍼 (직원정보 등 — 내부 CrossFade, 외부 슬라이드 방지) ──
    private TransitioningContentControl? _show2AgentWrapper;
    private TransitioningContentControl? _show3AgentWrapper;

    // ── 마지막으로 표시된 issue 캐시 (트리뷰 선택 null 시 복원용) ─────
    private ETA.Models.QuotationIssue? _lastShownIssue;

    // ── 분석의뢰 상세/목록 패널 ──────────────────────────────────────
    private AnalysisRequestDetailPanel?  _analysisRequestDetailPanel;
    private AnalysisRequestListPanel?    _analysisRequestListPanel;
    // ── 분석계획 (처리시설별 시료 × 분석항목 체크박스) ─────────────────────────
    private string[] _analysisPlanItems = Array.Empty<string>();
    private Dictionary<string, string[]> _facilitySamples = new();

    // facility → bool[sampleCount][itemCount]
    private readonly Dictionary<string, List<bool[]>> _facilityPlanState = new();
    private int _analysisPlanSelectedDay = -1; // -1=전체, 0=월..6=일, -2=BASE
    private string? _selectedAnalysisInfoForMapping;
    private int _pendingMappingItemIndex = -1;
    private string _pendingMappingItemName = "";

    public MainPage()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        ApplyTheme(0);

        // 애니메이션 토글 버튼 초기 상태 설정 & 비활성화 처리
        Dispatcher.UIThread.Post(() =>
        {
            var btn = this.FindControl<Control>("btnAnimations");
            if (btn != null)
            {
                var prop = btn.GetType().GetProperty("IsChecked");
                if (prop != null)
                    prop.SetValue(btn, App.AnimationsEnabled);
            }

            // 애니메이션 비활성화 (활성화되지 않은 경우)
            if (!App.AnimationsEnabled)
            {
                DisableTransitionsRecursive(this);
            }
        }, DispatcherPriority.Loaded);

        // QuotationNewPanel의 ListPanelChanged → Show2.Content 연결 (분석결과입력과 동일)
        if (_quotationNewPanel != null)
            _quotationNewPanel.ListPanelChanged += panel => { Show2.Content = panel; LogContentChange("Show2", panel); };

        // QuotationPage 회사 선택 시 Show4에 분석항목+단가 패널 표시
        if (_quotationPage != null)
        {
            _quotationPage.CompanySelected += company =>
            {
                Show4.Content = BuildAnalytePricePanel(company.C_CompanyName);
            };
        }

        // WindowPositionManager 초기화 - 완전 비활성화 (스턱 방지)
        _positionManager = null; // new WindowPositionManager(CurrentUserManager.Instance.CurrentUserId);

        // 윈도우 이벤트 연결
        this.Opened += MainPage_Opened;
        this.Closing += MainPage_Closing;

        // 한글 IME 마지막 글자 다음 입력창 이동 버그 수정
        // 포인터 클릭/Tab/Enter로 포커스가 이동하기 전에 IME 조합 상태를 초기화하여
        // 조합 중이던 마지막 글자가 다음 TextBox로 넘어가지 않게 함
        AddHandler(InputElement.PointerPressedEvent, OnPointerPressedImeReset,
            RoutingStrategies.Tunnel, handledEventsToo: false);
        AddHandler(InputElement.KeyDownEvent, OnKeyDownImeReset,
            RoutingStrategies.Tunnel, handledEventsToo: false);


        // GridSplitter 실시간 레이아웃 저장 이벤트 연결
        this.Loaded += SetupSplitterEvents;

        // SubMenu 버튼 호버 시 텍스트 그라데이션 효과
        TextShimmer.AttachHover(BT1); TextShimmer.AttachHover(BT2);
        TextShimmer.AttachHover(BT3); TextShimmer.AttachHover(BT4);
        TextShimmer.AttachHover(BT5); TextShimmer.AttachHover(BT6);
        TextShimmer.AttachHover(BT7); TextShimmer.AttachHover(BT8);

        // Show1~4 패널 호버 shimmer
        this.Loaded += (_, _) =>
        {
            var show1Border = this.FindControl<Border>("Show1Border");
            var show2Border = this.FindControl<Border>("Content2Border");
            var show3Border = this.FindControl<Border>("LowerBorder");
            var content4Border = this.FindControl<Border>("Content4Border");
            if (show1Border != null)    TextShimmer.AttachPanelHover(show1Border);
            if (show2Border != null)    TextShimmer.AttachPanelHover(show2Border);
            if (show3Border != null)    TextShimmer.AttachPanelHover(show3Border);
            if (content4Border != null) TextShimmer.AttachPanelHover(content4Border);

            // 디폴트 화면: 비용부담금 > 폐수배출업소
            WasteCompany_Click(null, new Avalonia.Interactivity.RoutedEventArgs());
        };
    }

    // 분석항목+단가 패널 생성 함수
    private Control BuildAnalytePricePanel(string companyName)
    {
        var analytes = AnalysisService.GetAllItems();
        var prices = ContractService.GetContractPrices(companyName)
            .ToDictionary(x => x.Analyte, x => x.Price);

        var stack = new StackPanel { Spacing = 2, Margin = new Thickness(8) };
        foreach (var a in analytes)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
            row.Children.Add(new TextBlock
            {
                Text = a.Analyte,
                Width = 120,
                FontSize = 14,
                Foreground = AppTheme.FgPrimary,
            });
            string priceStr = prices.TryGetValue(a.Analyte, out var p) && decimal.TryParse(p, out var d) && d > 0
                ? $"{d:N0}"
                : "—";
            row.Children.Add(new TextBlock
            {
                Text = priceStr,
                Width = 60,
                FontSize = 14,
                Foreground = AppTheme.FgSuccess,
                TextAlignment = Avalonia.Media.TextAlignment.Right,
            });
            stack.Children.Add(row);
        }
        return new ScrollViewer { Content = stack, VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto };
    }

    // ── 글로벌 키보드 단축키 ──
    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.KeyModifiers == KeyModifiers.Control && e.Key == Key.S)
        {
            e.Handled = true;
            BT1_Click(null, new RoutedEventArgs());
            return;
        }
    }

    // ── Show3 접기/펼치기 ──
    private GridLength _show3SavedHeight = new GridLength(192, GridUnitType.Star);
    private bool _show3Collapsed;

    // ── 레이아웃 모드: "default" = Show3 하단 전체폭, "split" = Show4 풀하이트 ──
    private string _layoutMode = "default";

    private void BtnToggleShow3_Click(object? sender, RoutedEventArgs e)
    {
        var grid = this.FindControl<Grid>("RightSplitGrid");
        if (grid == null) return;

        var rowDef = grid.RowDefinitions[2]; // Show3 row
        if (_show3Collapsed)
        {
            rowDef.Height = _show3SavedHeight;
            LowerBorder.IsVisible = true;
            VerticalSplitter.IsVisible = true;
            btnToggleShow3.Content = "▼";
            _show3Collapsed = false;
        }
        else
        {
            _show3SavedHeight = rowDef.Height;
            rowDef.Height = new GridLength(0);
            LowerBorder.IsVisible = false;
            VerticalSplitter.IsVisible = false;
            btnToggleShow3.Content = "▲";
            _show3Collapsed = true;
        }
    }

    private void BtnLayoutToggle_Click(object? sender, RoutedEventArgs e)
    {
        // Show3 버튼: default ↔ split (Show4 풀하이트)
        var newMode = _layoutMode == "split" ? "default" : "split";
        SetPanelLayout(newMode);
    }

    private void BtnLayoutToggle2_Click(object? sender, RoutedEventArgs e)
    {
        // Show2 버튼: default ↔ expand2 (Show2 풀하이트)
        var newMode = _layoutMode == "expand2" ? "default" : "expand2";
        SetPanelLayout(newMode);
    }

    private void BtnLayoutToggle4_Click(object? sender, RoutedEventArgs e)
    {
        // Show4 버튼: default ↔ split (Show4 풀하이트)
        var newMode = _layoutMode == "split" ? "default" : "split";
        SetPanelLayout(newMode);
    }

    private void BtnLayoutToggle4Top_Click(object? sender, RoutedEventArgs e)
    {
        // Show4 상단 버튼: default ↔ expand2top (Show4 하단이동 + Show2 상단 확장)
        var newMode = _layoutMode == "expand2top" ? "default" : "expand2top";
        SetPanelLayout(newMode);
    }

    /// <summary>
    /// 레이아웃 모드 전환.
    /// "default":    Show2+Show4 상단, Show3 하단 전체폭
    /// "split":      Show2+Show3 좌측 상하, Show4 우측 풀하이트
    /// "expand2":    Show2 좌측 풀하이트, Show4+Show3 우측 상하
    /// "expand2top": Show2 상단 전체폭, Show3+Show4 하단 좌우
    /// </summary>
    private void SetPanelLayout(string mode)
    {
        _layoutMode = mode;
        var grid = this.FindControl<Grid>("RightSplitGrid");
        if (grid == null) return;

        var c2 = this.FindControl<Border>("Content2Border");
        var c4 = this.FindControl<Border>("Content4Border");
        var lb = this.FindControl<Border>("LowerBorder");
        var hs = this.FindControl<GridSplitter>("HorizontalSplitter");
        var vs = this.FindControl<GridSplitter>("VerticalSplitter");
        var tb = this.FindControl<Button>("btnToggleShow3");
        if (c2 == null || c4 == null || lb == null || hs == null || vs == null) return;

        if (mode == "split")
        {
            // Show2: Row=0, Col=0
            Grid.SetRow(c2, 0); Grid.SetColumn(c2, 0);
            Grid.SetRowSpan(c2, 1); Grid.SetColumnSpan(c2, 1);

            // Show3: Row=2, Col=0 (Show2 아래)
            Grid.SetRow(lb, 2); Grid.SetColumn(lb, 0);
            Grid.SetRowSpan(lb, 1); Grid.SetColumnSpan(lb, 1);

            // Show4: Row=0, Col=2, RowSpan=3 (풀하이트)
            Grid.SetRow(c4, 0); Grid.SetColumn(c4, 2);
            Grid.SetRowSpan(c4, 3); Grid.SetColumnSpan(c4, 1);

            // 좌우 스플리터: 전체 높이
            Grid.SetRow(hs, 0); Grid.SetColumn(hs, 1);
            Grid.SetRowSpan(hs, 3); Grid.SetColumnSpan(hs, 1);

            // 상하 스플리터: 좌측 열만
            Grid.SetRow(vs, 1); Grid.SetColumn(vs, 0);
            Grid.SetRowSpan(vs, 1); Grid.SetColumnSpan(vs, 1);

            if (tb != null)
            {
                Grid.SetRow(tb, 1); Grid.SetColumn(tb, 0);
                Grid.SetColumnSpan(tb, 1);
            }
        }
        else if (mode == "expand2")
        {
            // Show2: Row=0, Col=0, RowSpan=3 (풀하이트)
            Grid.SetRow(c2, 0); Grid.SetColumn(c2, 0);
            Grid.SetRowSpan(c2, 3); Grid.SetColumnSpan(c2, 1);

            // Show4: Row=0, Col=2
            Grid.SetRow(c4, 0); Grid.SetColumn(c4, 2);
            Grid.SetRowSpan(c4, 1); Grid.SetColumnSpan(c4, 1);

            // Show3: Row=2, Col=2 (Show4 아래)
            Grid.SetRow(lb, 2); Grid.SetColumn(lb, 2);
            Grid.SetRowSpan(lb, 1); Grid.SetColumnSpan(lb, 1);

            // 좌우 스플리터: 전체 높이
            Grid.SetRow(hs, 0); Grid.SetColumn(hs, 1);
            Grid.SetRowSpan(hs, 3); Grid.SetColumnSpan(hs, 1);

            // 상하 스플리터: 우측 열만
            Grid.SetRow(vs, 1); Grid.SetColumn(vs, 2);
            Grid.SetRowSpan(vs, 1); Grid.SetColumnSpan(vs, 1);

            if (tb != null)
            {
                Grid.SetRow(tb, 1); Grid.SetColumn(tb, 2);
                Grid.SetColumnSpan(tb, 1);
            }
        }
        else if (mode == "expand2top")
        {
            // Show2: Row=0, Col=0, ColSpan=3 (상단 전체폭 확장)
            Grid.SetRow(c2, 0); Grid.SetColumn(c2, 0);
            Grid.SetRowSpan(c2, 1); Grid.SetColumnSpan(c2, 3);

            // Show4: Row=2, Col=2 (하단 우측)
            Grid.SetRow(c4, 2); Grid.SetColumn(c4, 2);
            Grid.SetRowSpan(c4, 1); Grid.SetColumnSpan(c4, 1);

            // Show3: Row=2, Col=0 (하단 좌측)
            Grid.SetRow(lb, 2); Grid.SetColumn(lb, 0);
            Grid.SetRowSpan(lb, 1); Grid.SetColumnSpan(lb, 1);

            // 좌우 스플리터: 하단 행 (Show3 ↔ Show4 사이)
            Grid.SetRow(hs, 2); Grid.SetColumn(hs, 1);
            Grid.SetRowSpan(hs, 1); Grid.SetColumnSpan(hs, 1);

            // 상하 스플리터: 전체 폭
            Grid.SetRow(vs, 1); Grid.SetColumn(vs, 0);
            Grid.SetRowSpan(vs, 1); Grid.SetColumnSpan(vs, 3);

            if (tb != null)
            {
                Grid.SetRow(tb, 1); Grid.SetColumn(tb, 0);
                Grid.SetColumnSpan(tb, 3);
            }
        }
        else // "default"
        {
            // Show2: Row=0, Col=0
            Grid.SetRow(c2, 0); Grid.SetColumn(c2, 0);
            Grid.SetRowSpan(c2, 1); Grid.SetColumnSpan(c2, 1);

            // Show4: Row=0, Col=2
            Grid.SetRow(c4, 0); Grid.SetColumn(c4, 2);
            Grid.SetRowSpan(c4, 1); Grid.SetColumnSpan(c4, 1);

            // Show3: Row=2, Col=0, ColSpan=3 (전체 폭)
            Grid.SetRow(lb, 2); Grid.SetColumn(lb, 0);
            Grid.SetRowSpan(lb, 1); Grid.SetColumnSpan(lb, 3);

            // 좌우 스플리터: 상단만
            Grid.SetRow(hs, 0); Grid.SetColumn(hs, 1);
            Grid.SetRowSpan(hs, 1); Grid.SetColumnSpan(hs, 1);

            // 상하 스플리터: 전체 폭
            Grid.SetRow(vs, 1); Grid.SetColumn(vs, 0);
            Grid.SetRowSpan(vs, 1); Grid.SetColumnSpan(vs, 3);

            if (tb != null)
            {
                Grid.SetRow(tb, 1); Grid.SetColumn(tb, 0);
                Grid.SetColumnSpan(tb, 3);
            }
        }

        // 버튼 아이콘 갱신
        btnLayoutToggle.Content     = mode == "split"      ? "◢" : "◤";
        btnLayoutToggle2.Content    = mode == "expand2"    ? "◤" : "◣";
        btnLayoutToggle4.Content    = mode == "split"      ? "◥" : "◤";
        if (btnLayoutToggle4Top != null)
            btnLayoutToggle4Top.Content = mode == "expand2top" ? "▲" : "▼";
    }

    // ── 로그인 → 메인 전환 확장 애니메이션 ────────────────────────────────
    private double _expandFromW, _expandFromH;
    private PixelPoint _expandFromPos;
    private bool _shouldAnimate;

    public void AnimateExpand(double fromW, double fromH, PixelPoint fromPos)
    {
        _expandFromW = fromW;
        _expandFromH = fromH;
        _expandFromPos = fromPos;
        _shouldAnimate = true;
    }

    private async void RunExpandAnimation()
    {
        if (!_shouldAnimate) return;
        _shouldAnimate = false;

        double targetW = _expandFromW * 1.5;
        double targetH = _expandFromH * 1.5;

        // 화면 크기 제한
        var screen = Screens.Primary ?? Screens.All.FirstOrDefault();
        if (screen != null)
        {
            var wa = screen.WorkingArea;
            double maxW = wa.Width / screen.Scaling;
            double maxH = wa.Height / screen.Scaling;
            if (targetW > maxW) targetW = maxW;
            if (targetH > maxH) targetH = maxH;
        }

        int steps = 30;
        double duration = 400; // ms
        double stepMs = duration / steps;

        double startW = _expandFromW;
        double startH = _expandFromH;
        double startX = _expandFromPos.X;
        double startY = _expandFromPos.Y;

        // 최종 위치: 중앙 유지
        double scaling = screen?.Scaling ?? 1.0;
        double endX = startX - (targetW - startW) * scaling / 2;
        double endY = startY - (targetH - startH) * scaling / 2;

        // 화면 밖으로 나가지 않도록 보정
        if (endX < 0) endX = 0;
        if (endY < 0) endY = 0;

        for (int i = 1; i <= steps; i++)
        {
            double t = (double)i / steps;
            // CubicEaseOut: t = 1 - (1-t)^3
            double ease = 1 - Math.Pow(1 - t, 3);

            Width  = startW + (targetW - startW) * ease;
            Height = startH + (targetH - startH) * ease;
            Position = new PixelPoint(
                (int)(startX + (endX - startX) * ease),
                (int)(startY + (endY - startY) * ease));

            await Task.Delay((int)stepMs);
        }

        // 최종 값 보정
        Width = targetW;
        Height = targetH;
        Position = new PixelPoint((int)endX, (int)endY);
    }

    /// <summary>
    /// <summary>
    /// 한글 IME 조합 중 포커스 이동 시 마지막 글자가 다음 입력창으로 넘어가는 버그 방지.
    /// 포인터 클릭/Tab/Enter 등 모든 포커스 변경 전에 IME 상태를 초기화.
    /// </summary>
    private void OnPointerPressedImeReset(object? sender, PointerPressedEventArgs e)
    {
        var focused = FocusManager?.GetFocusedElement();
        if (focused is not TextBox) return;
        if (e.Source is TextBox clickedBox && clickedBox == focused) return;

        ResetImeComposition(focused as TextBox);
    }

    /// <summary>Tab/Enter 키로 포커스가 이동할 때 IME 조합 상태 초기화</summary>
    private void OnKeyDownImeReset(object? sender, KeyEventArgs e)
    {
        // Tab, Shift+Tab, Enter 키만 처리
        if (e.Key != Key.Tab && e.Key != Key.Return) return;

        var focused = FocusManager?.GetFocusedElement() as TextBox;
        if (focused == null) return;

        // 포커스 이동 직전에 IME를 리셋하되, 키 입력 처리 직후에 포커스 이동이 일어나도록
        // PreviewKeyDown(Tunnel) → KeyDown(Bubble) 순서를 이용
        ResetImeComposition(focused);
    }

    /// <summary>IME 조합 상태를 강제로 초기화 (한글 입력 안정화)</summary>
    private static void ResetImeComposition(TextBox? tb)
    {
        if (tb == null) return;

        try
        {
            // IME 클라이언트를 종료했다가 재생성 → 조합 중이던 글자를 현재 TextBox에서 flush
            Avalonia.Input.InputMethod.SetIsInputMethodEnabled(tb, false);
            // 약간의 딜레이 후 재활성화 (Avalonia의 버그 우회)
            Dispatcher.UIThread.Post(() =>
                Avalonia.Input.InputMethod.SetIsInputMethodEnabled(tb, true),
                DispatcherPriority.Input);
        }
        catch { }
    }

    private void MainPage_Opened(object? sender, EventArgs e)
    {
        // ── 직원정보 메뉴: 허용된 사번만 활성화 ──────────────────────────
        var allowedIds = new System.Collections.Generic.HashSet<string>
        {
            "201000308",  // 정승욱
            "202104002",  // 방찬미
            "201101234",  // 박은지
        };
        if (MenuItemAgentInfo != null)
            MenuItemAgentInfo.IsEnabled = allowedIds.Contains(CurrentEmployeeId);

        // ── 접근관리: 201000308 (정승욱) 전용 ───────────────────────────────
        if (MenuItemAccess != null)
            MenuItemAccess.IsVisible = CurrentEmployeeId == "201000308";

        RunExpandAnimation();
        _ = LoadProfileInfoAsync();

        // ── AccessService: AgentAccess 테이블 생성 + 메뉴 접근 필터 적용 ──
        AccessService.EnsureTable();
        ApplyMenuAccessFilter();

        // 업무분장 자동 연장 (6개월 후까지, 프로그레스바 표시)
        {
            var overlay = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 8,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 8),
            };
            var pb = new ProgressBar
            {
                Minimum = 0, Maximum = 1, Value = 0,
                Height = 6, Width = 200, CornerRadius = new CornerRadius(3),
                Foreground = new SolidColorBrush(Color.Parse("#ff8844")),
                Background = new SolidColorBrush(Color.Parse("#333")),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            var lbl = new TextBlock
            {
                FontSize = 11, Foreground = new SolidColorBrush(Color.Parse("#ff8844")),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Text = "분장표 동기화 중...",
            };
            overlay.Children.Add(pb);
            overlay.Children.Add(lbl);

            // MainPage의 최상위 패널에 추가
            if (this.Content is Panel rootPanel)
                rootPanel.Children.Add(overlay);

            _ = Task.Run(() =>
            {
                AnalysisRequestService.AutoExtendAssignmentsToToday((pct, name) =>
                {
                    Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                    {
                        pb.Value = pct;
                        lbl.Text = string.IsNullOrEmpty(name) ? "분장표 동기화 중..." : $"{(int)(pct * 100)}% {name}";
                    });
                });
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    if (this.Content is Panel rp)
                        rp.Children.Remove(overlay);
                });
            });
        }

        // 저장된 글자 크기 복원
        var savedScale = LoadFontScale();
        ApplyFontScale(savedScale);
        if (sldFontSize != null) sldFontSize.Value = savedScale;
        if (txbFontScale != null) txbFontScale.Text = $"{(int)(savedScale * 100)}%";

        // MenuItem 호버 시 텍스트 그라데이션 효과 (최상위 + 서브메뉴 열릴 때 동적 부착)
        foreach (var mi in this.GetVisualDescendants().OfType<MenuItem>())
        {
            TextShimmer.AttachHover(mi);
            mi.SubmenuOpened += (_, _) =>
            {
                // 서브메뉴가 팝업으로 열리면 비주얼 트리에 새 MenuItem이 생김
                Dispatcher.UIThread.Post(() =>
                {
                    foreach (var sub in mi.GetVisualDescendants().OfType<MenuItem>())
                        TextShimmer.AttachIfNew(sub);
                }, DispatcherPriority.Background);
            };
        }

    }

    // ── AccessService 기반 메뉴 접근 필터 적용 ─────────────────────────────
    private void ApplyMenuAccessFilter()
    {
        if (string.IsNullOrEmpty(CurrentEmployeeId)) return;

        // 201000308은 모든 메뉴 허용 (관리자)
        if (CurrentEmployeeId == "201000308") return;

        var allowed = AccessService.GetAllowedMenus(CurrentEmployeeId);

        // 메뉴키 → MenuItem 매핑
        var menuMap = new Dictionary<string, MenuItem?>
        {
            ["Agent"]                    = this.FindControl<MenuItem>("MenuItemAgentInfo"),
            ["Contract"]                 = this.FindControl<MenuItem>("MenuItemContract"),
            ["Quotation"]                = this.FindControl<MenuItem>("MenuItemQuotation"),
            ["QuotationIssue"]           = this.FindControl<MenuItem>("MenuItemQuotationIssue"),
            ["TestReport"]               = this.FindControl<MenuItem>("MenuItemTestReport"),
            ["DbMigration"]              = this.FindControl<MenuItem>("MenuItemDbMigration"),
            ["WasteCompany"]             = this.FindControl<MenuItem>("MenuItemWasteCompany"),
            ["WasteSampleList"]          = this.FindControl<MenuItem>("MenuItemWasteSampleList"),
            ["WasteTestReport"]          = this.FindControl<MenuItem>("MenuItemWasteTestReport"),
            ["ProcessingFacility"]       = this.FindControl<MenuItem>("MenuItemProcessingFacility"),
            ["WasteAnalysisInputWater"]  = this.FindControl<MenuItem>("MenuItemWasteAnalysisInputWater"),
            ["WasteAnalysisInputWaste"]  = this.FindControl<MenuItem>("MenuItemWasteAnalysisInputWaste"),
            ["ResultSubmitMeasure"]      = this.FindControl<MenuItem>("MenuItemResultSubmitMeasure"),
            ["ResultSubmitErp"]          = this.FindControl<MenuItem>("MenuItemResultSubmitErp"),
            ["Repair"]                   = this.FindControl<MenuItem>("MenuItemRepair"),
            ["Purchase"]                 = this.FindControl<MenuItem>("MenuItemPurchase"),
            ["Schedule"]                 = this.FindControl<MenuItem>("MenuItemSchedule2"),
            ["Reagent"]                  = this.FindControl<MenuItem>("MenuItemReagent"),
            ["WaterQualityNameReconcile"]= this.FindControl<MenuItem>("MenuItemWaterQualityNameReconcile"),
            ["WasteNameReconcile"]       = this.FindControl<MenuItem>("MenuItemWasteNameReconcile"),
            ["Access"]                   = this.FindControl<MenuItem>("MenuItemAccess"),
        };

        foreach (var (key, mi) in menuMap)
        {
            if (mi == null) continue;
            bool isAllowed = allowed.Contains(key);
            mi.IsVisible = isAllowed;
            mi.IsEnabled = isAllowed;
        }

        // 부모 메뉴: 자식이 모두 숨겨지면 부모도 숨김
        var parentChildKeys = new Dictionary<string, string[]>
        {
            ["MenuItemWaterCenter"]   = new[] { "Agent", "Contract", "Quotation", "QuotationIssue", "TestReport", "DbMigration" },
            ["MenuItemWaste"]         = new[] { "WasteCompany", "WasteSampleList", "WasteTestReport" },
            ["MenuItemFacility"]      = new[] { "ProcessingFacility" },
            ["MenuItemAnalysisInput"] = new[] { "WasteAnalysisInputWater", "WasteAnalysisInputWaste" },
            ["MenuItemResultSubmit"]  = new[] { "ResultSubmitMeasure", "ResultSubmitErp" },
            ["MenuItemSchedule"]      = new[] { "Repair", "Purchase", "Schedule" },
            ["MenuItemRisk"]          = new[] { "Reagent" },
            ["PermissionMenu"]        = new[] { "WaterQualityNameReconcile", "WasteNameReconcile", "Access" },
        };

        foreach (var (parentName, childKeys) in parentChildKeys)
        {
            var parent = this.FindControl<MenuItem>(parentName);
            if (parent == null) continue;
            bool anyAllowed = childKeys.Any(k => allowed.Contains(k));
            parent.IsVisible = anyAllowed;
            parent.IsEnabled = anyAllowed;
        }
    }

    private async Task LoadProfileInfoAsync()
    {
        try
        {
            var empId = CurrentEmployeeId;
            if (string.IsNullOrEmpty(empId)) return;

            // DB 쿼리 + 파일 I/O → 백그라운드 스레드
            var (me, photoFullPath) = await Task.Run(() =>
            {
                var agents = AgentService.GetAllItems();
                var found  = agents.FirstOrDefault(a => a.사번 == empId);
                if (found == null) return (found, (string?)null);

                string? fp = null;
                if (!string.IsNullOrEmpty(found.PhotoPath))
                {
                    ETA.Services.SERVICE1.AgentService.EnsurePhotoLocal(found.사번, found.PhotoPath);
                    var candidate = Path.IsPathRooted(found.PhotoPath)
                        ? found.PhotoPath
                        : Path.Combine(ETA.Services.SERVICE1.AgentService.GetPhotoDirectory(), found.PhotoPath);
                    if (File.Exists(candidate)) fp = candidate;
                }
                return (found, fp);
            });

            if (me == null) return;

            // UI 업데이트 (이미 UI 스레드로 복귀됨)
            if (profileName != null)
                profileName.Text = me.성명;

            if (profilePhoto != null && photoFullPath != null)
            {
                profilePhoto.Source    = new Avalonia.Media.Imaging.Bitmap(photoFullPath);
                profilePhoto.IsVisible = true;
            }
        }
        catch (Exception ex)
        {
        }
    }

    private void ProfileName_Click(object? sender, PointerPressedEventArgs e)
        => MyTask_Click(sender, new RoutedEventArgs());

    private void ProfileName_MenuClick(object? sender, RoutedEventArgs e)
        => MyTask_Click(sender, new RoutedEventArgs());

    private void Logout_Click(object? sender, RoutedEventArgs e)
    {
        ETA.Services.Common.CurrentUserManager.Instance.ResetToWindowsUser();
        MainPage.CurrentEmployeeId = "";
        var login = new Login();
        login.Show();
        Close();
    }

    private void MainPage_Closing(object? sender, WindowClosingEventArgs e)
    {
        if (!string.IsNullOrEmpty(_currentMode) && _currentMode != "None")
        {
            SaveCurrentModeLayout();
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  분석항목 리스트 (드래그 앤 드랍용)
    // ══════════════════════════════════════════════════════════════════════

    private ListBox?        _analysisItemsListBox;
    private Control?        _content4Container;
    private ContentControl? _show4AgentWrapper;  // Show4 내부 스왑용 영구 wrapper
    private DateTime    _content4QueryStart = DateTime.Today;
    private DateTime    _content4QueryEnd   = DateTime.Today;

    private Control CreateAnalysisItemsListBox()
    {
        if (_content4Container != null) return _content4Container;

        // ── 리스트박스 ────────────────────────────────────────────────────
        _analysisItemsListBox = new ListBox
        {
            Background      = AppTheme.BgPrimary,
            BorderThickness = new Thickness(0),
            SelectionMode   = SelectionMode.Multiple,
            Margin          = new Thickness(2, 0, 2, 2),
        };
        LoadAnalysisItems();

        // ── 날짜 범위 행 ──────────────────────────────────────────────────
        var txbRange = new TextBlock
        {
            Text              = DateTime.Today.ToString("yyyy-MM-dd"),
            FontSize          = AppTheme.FontSM,
            Foreground        = AppTheme.FgMuted,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            MinWidth          = 160,
            FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        };

        var btnCal = new Button
        {
            Content         = "📅",
            Width           = 28,
            Height          = 22,
            FontSize        = AppTheme.FontBase,
            Padding         = new Thickness(0),
            Background      = new SolidColorBrush(Color.Parse("#2a3a4a")),
            Foreground      = AppTheme.FgMuted,
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderDefault,
        };
        ToolTip.SetTip(btnCal, "기간 선택 (드래그로 범위 설정)");

        var btnToday = new Button
        {
            Content         = "오늘",
            Width           = 42,
            Height          = 22,
            FontSize        = AppTheme.FontSM,
            Padding         = new Thickness(4, 0),
            Background      = new SolidColorBrush(Color.Parse("#3a5a3a")),
            Foreground      = AppTheme.FgMuted,
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderMuted,
        };

        var dateRow = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing     = 5,
            Margin      = new Thickness(4, 4, 4, 2),
        };
        dateRow.Children.Add(txbRange);
        dateRow.Children.Add(btnCal);
        dateRow.Children.Add(btnToday);

        // ── 인라인 달력 ───────────────────────────────────────────────────
        var calendar = new Avalonia.Controls.Calendar
        {
            SelectionMode = Avalonia.Controls.CalendarSelectionMode.SingleRange,
            IsVisible     = false,
            Margin        = new Thickness(4, 0, 4, 2),
            DisplayDate   = DateTime.Today,
            Focusable     = true,
        };
        calendar.SelectedDates.Add(DateTime.Today);
        calendar.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape && calendar.IsVisible)
            { calendar.IsVisible = false; e.Handled = true; }
        };

        // ── 이벤트 연결 ───────────────────────────────────────────────────
        btnCal.Click += (_, _) => { calendar.IsVisible = !calendar.IsVisible; if (calendar.IsVisible) calendar.Focus(); };

        btnToday.Click += (_, _) =>
        {
            _content4QueryStart = _content4QueryEnd = DateTime.Today;
            txbRange.Text       = DateTime.Today.ToString("yyyy-MM-dd");
            calendar.IsVisible  = false;
            UpdateAssignmentInfo();
        };

        calendar.SelectedDatesChanged += (_, _) =>
        {
            if (calendar.SelectedDates.Count == 0) return;
            var dates = calendar.SelectedDates.Cast<DateTime>().ToList();
            _content4QueryStart = dates.Min();
            _content4QueryEnd   = dates.Max();
            txbRange.Text = _content4QueryStart == _content4QueryEnd
                ? _content4QueryStart.ToString("yyyy-MM-dd")
                : $"{_content4QueryStart:yyyy-MM-dd} ~ {_content4QueryEnd:yyyy-MM-dd}";
            UpdateAssignmentInfo();
            if (_content4QueryStart != _content4QueryEnd)
                calendar.IsVisible = false;
        };

        // ── 검색창 ────────────────────────────────────────────────────────
        var searchBox = new TextBox
        {
            Watermark       = "\ud83d\udd0d \ubd84\uc11d\ud56d\ubaa9 \uac80\uc0c9...",
            FontSize        = AppTheme.FontMD,
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Height          = 32,
            Margin          = new Thickness(4, 2, 4, 4),
            CornerRadius    = new CornerRadius(6),
            Padding         = new Thickness(8, 5),
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderDefault,
            Background      = AppTheme.BgPrimary,
            Foreground      = AppTheme.FgPrimary,
        };
        // 한글 IME 안전: TextChanged → 400ms 디바운스 후 검색
        DispatcherTimer? searchTimer = null;
        searchBox.TextChanged += (_, _) =>
        {
            searchTimer?.Stop();
            searchTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            searchTimer.Tick += (_, _) =>
            {
                searchTimer!.Stop();
                if (_analysisItemsListBox == null) return;
                var q = searchBox.Text?.Trim().ToLower() ?? "";
                foreach (var item in _analysisItemsListBox.Items.OfType<ListBoxItem>())
                {
                    var name = item.Tag as string ?? "";
                    item.IsVisible = string.IsNullOrEmpty(q) || name.ToLower().Contains(q);
                }
            };
            searchTimer.Start();
        };

        // ── 분석항목 컨테이너 (DockPanel) ────────────────────────────────
        var header = new StackPanel { Spacing = 0 };
        header.Children.Add(dateRow);
        header.Children.Add(calendar);
        header.Children.Add(searchBox);

        var dock = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(header, Dock.Top);
        dock.Children.Add(header);
        dock.Children.Add(_analysisItemsListBox);

        // ── 계약업체 리스트 컨테이너 ─────────────────────────────────────
        var contractStack = new StackPanel { Spacing = 0 };
        var contractScroll = new ScrollViewer
        {
            Content = contractStack,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };

        void LoadContracts()
        {
            contractStack.Children.Clear();
            try
            {
                var contracts = ContractService.GetAllContracts();
                bool odd = false;
                foreach (var c in contracts)
                {
                    var row = new Border
                    {
                        Background = new SolidColorBrush(Color.Parse(odd ? "#1a1a28" : "#1e1e30")),
                        Padding    = new Thickness(8, 4),
                    };
                    odd = !odd;

                    var inner = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };

                    var name = new TextBlock
                    {
                        Text              = c.C_CompanyName,
                        FontSize          = AppTheme.FontBase,
                        Foreground        = AppTheme.FgSecondary,
                        FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                        VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                        TextTrimming      = TextTrimming.CharacterEllipsis,
                    };
                    Grid.SetColumn(name, 0);
                    inner.Children.Add(name);

                    if (!string.IsNullOrEmpty(c.C_Abbreviation))
                    {
                        var (bg, fg) = BadgeColorHelper.GetBadgeColor(c.C_ContractType ?? "");
                        var badge = new Border
                        {
                            Background        = new SolidColorBrush(Color.Parse(bg)),
                            CornerRadius      = new CornerRadius(3),
                            Padding           = new Thickness(4, 1),
                            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                            [Grid.ColumnProperty] = 1,
                            Child             = new TextBlock
                            {
                                Text       = c.C_Abbreviation,
                                FontSize   = AppTheme.FontXS,
                                Foreground = new SolidColorBrush(Color.Parse(fg)),
                                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                            },
                        };
                        inner.Children.Add(badge);
                    }

                    row.Child = inner;
                    contractStack.Children.Add(row);
                }
                if (contracts.Count == 0)
                    contractStack.Children.Add(new TextBlock
                    {
                        Text = "계약업체 없음", FontSize = AppTheme.FontBase,
                        Foreground = AppTheme.FgDimmed,
                        Margin = new Thickness(8, 8),
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    });
            }
            catch { }
        }

        // ── 탭 토글 버튼 행 ───────────────────────────────────────────────
        var btnAnalysis = new Button
        {
            Content         = "분석항목",
            FontSize        = AppTheme.FontBase,
            Height          = 26,
            Padding         = new Thickness(12, 0),
            BorderThickness = new Thickness(0),
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background      = new SolidColorBrush(Color.Parse("#3a4a6a")),
            Foreground      = new SolidColorBrush(Color.Parse("#ddeeff")),
        };
        var btnContract = new Button
        {
            Content         = "계약업체",
            FontSize        = AppTheme.FontBase,
            Height          = 26,
            Padding         = new Thickness(12, 0),
            BorderThickness = new Thickness(0),
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
            Foreground      = AppTheme.FgMuted,
        };

        // ── 일반업무 리스트 컨테이너 ────────────────────────────────────
        var generalTaskStack = new StackPanel { Spacing = 0 };
        var generalTaskScroll = new ScrollViewer
        {
            Content = generalTaskStack,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };

        // ── 인라인 업무 추가/수정 폼 빌더 ──────────────────────────────────
        void ShowInlineTaskForm(ETA.Models.GeneralTask? existing, Action onDone)
        {
            bool isEdit = existing != null;
            generalTaskStack.Children.Clear();

            var agentFont  = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
            var agentFontM = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

            var form = new StackPanel { Spacing = 6, Margin = new Thickness(6) };

            form.Children.Add(new TextBlock
            {
                Text = isEdit ? "✏️ 일반업무 수정" : "➕ 신규 일반업무 추가",
                FontSize = AppTheme.FontMD, FontWeight = FontWeight.Bold,
                FontFamily = agentFontM,
                Foreground = AppTheme.FgWarn,
            });

            var txtName = new TextBox
            {
                Text = existing?.업무명 ?? "",
                Watermark = "업무명",
                FontFamily = agentFont, FontSize = AppTheme.FontBase,
                Background = AppTheme.BorderSeparator,
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderDefault,
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 4),
            };
            form.Children.Add(txtName);

            var txtContent = new TextBox
            {
                Text = existing?.내용 ?? "",
                Watermark = "상세 내용",
                AcceptsReturn = true,
                TextWrapping = TextWrapping.Wrap,
                MinHeight = 60,
                FontFamily = agentFont, FontSize = AppTheme.FontBase,
                Background = AppTheme.BorderSeparator,
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderDefault,
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 4),
            };
            form.Children.Add(txtContent);

            // 담당자 콤보박스
            var agentRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 6 };
            agentRow.Children.Add(new TextBlock
            {
                Text = "담당자", FontSize = AppTheme.FontSM, FontFamily = agentFont,
                Foreground = Res("FgMuted"),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            });
            var agents = ETA.Services.SERVICE1.AgentService.GetAllItems().OrderBy(a => a.성명).ToList();
            var cboAgent = new ComboBox
            {
                Width = 160, FontSize = AppTheme.FontBase, FontFamily = agentFont,
                Background = AppTheme.BorderSeparator,
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderDefault,
                CornerRadius = new CornerRadius(4),
            };
            foreach (var a in agents)
                cboAgent.Items.Add(new ComboBoxItem { Content = a.성명, Tag = a.사번 });
            if (isEdit && !string.IsNullOrEmpty(existing!.담당자id))
                for (int i = 0; i < cboAgent.Items.Count; i++)
                    if (cboAgent.Items[i] is ComboBoxItem ci && ci.Tag?.ToString() == existing.담당자id)
                    { cboAgent.SelectedIndex = i; break; }
            agentRow.Children.Add(cboAgent);
            form.Children.Add(agentRow);

            // 마감일
            var deadlineRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 6 };
            deadlineRow.Children.Add(new TextBlock
            {
                Text = "마감일", FontSize = AppTheme.FontSM, FontFamily = agentFont,
                Foreground = Res("FgMuted"),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            });
            var txtDeadline = new TextBox
            {
                Text = existing?.마감일 ?? "",
                Watermark = "yyyy-MM-dd",
                Width = 120, FontFamily = agentFont, FontSize = AppTheme.FontBase,
                Background = AppTheme.BorderSeparator,
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderDefault,
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 3),
            };
            deadlineRow.Children.Add(txtDeadline);
            form.Children.Add(deadlineRow);

            // 상태 토글 (수정시)
            string currentStatus = existing?.상태 ?? "대기";
            if (isEdit)
            {
                var statusRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 4 };
                statusRow.Children.Add(new TextBlock
                {
                    Text = "상태", FontSize = AppTheme.FontSM, FontFamily = agentFont,
                    Foreground = Res("FgMuted"),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                });
                foreach (var st in new[] { "대기", "진행", "완료" })
                {
                    var capturedSt = st;
                    var (sbg, sfg, sbd) = st switch
                    {
                        "진행" => ("#2a3a1a", "#88cc88", "#3a6a3a"),
                        "완료" => ("#1a2a3a", "#88aacc", "#336699"),
                        _      => ("#2a2a2a", "#ccaa88", "#665533"),
                    };
                    bool active = st == currentStatus;
                    var btnSt = new Button
                    {
                        Content = st, FontSize = AppTheme.FontXS, FontFamily = agentFont,
                        Padding = new Thickness(6, 2),
                        CornerRadius = new CornerRadius(6),
                        BorderThickness = new Thickness(1),
                        Background  = active ? new SolidColorBrush(Color.Parse(sbg)) : AppTheme.BgSecondary,
                        Foreground  = active ? new SolidColorBrush(Color.Parse(sfg)) : AppTheme.BorderMuted,
                        BorderBrush = active ? new SolidColorBrush(Color.Parse(sbd)) : AppTheme.BorderMuted,
                    };
                    btnSt.Click += (_, _) =>
                    {
                        currentStatus = capturedSt;
                        foreach (var child in statusRow.Children.OfType<Button>())
                        {
                            var s = child.Content?.ToString() ?? "";
                            var (b2, f2, d2) = s switch
                            {
                                "진행" => ("#2a3a1a", "#88cc88", "#3a6a3a"),
                                "완료" => ("#1a2a3a", "#88aacc", "#336699"),
                                _      => ("#2a2a2a", "#ccaa88", "#665533"),
                            };
                            bool on2 = s == capturedSt;
                            child.Background  = on2 ? new SolidColorBrush(Color.Parse(b2)) : AppTheme.BgSecondary;
                            child.Foreground  = on2 ? new SolidColorBrush(Color.Parse(f2)) : AppTheme.BorderMuted;
                            child.BorderBrush = on2 ? new SolidColorBrush(Color.Parse(d2)) : AppTheme.BorderMuted;
                        }
                    };
                    statusRow.Children.Add(btnSt);
                }
                form.Children.Add(statusRow);
            }

            // 버튼 행
            var btnRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 4, Margin = new Thickness(0, 6, 0, 0) };
            var btnSave = new Button
            {
                Content = isEdit ? "💾 수정" : "💾 저장",
                Height = 26, Padding = new Thickness(12, 0),
                FontFamily = agentFont, FontSize = AppTheme.FontSM,
                Background = AppTheme.BgActiveGreen,
                Foreground = AppTheme.FgSuccess,
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            };
            btnSave.Click += (_, _) =>
            {
                if (string.IsNullOrWhiteSpace(txtName.Text)) return;
                var selItem = cboAgent.SelectedItem as ComboBoxItem;
                string agentId   = selItem?.Tag?.ToString() ?? "";
                string agentName = selItem?.Content?.ToString() ?? "";

                if (isEdit)
                {
                    existing!.업무명 = txtName.Text.Trim();
                    existing!.내용   = txtContent.Text?.Trim() ?? "";
                    existing!.마감일 = txtDeadline.Text?.Trim() ?? "";
                    existing!.상태   = currentStatus;
                    if (!string.IsNullOrEmpty(agentId)) { existing!.담당자id = agentId; existing!.담당자명 = agentName; }
                    if (currentStatus == "완료" && string.IsNullOrEmpty(existing.완료일시))
                        existing.완료일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    ETA.Services.Common.GeneralTaskService.Update(existing);
                }
                else
                {
                    var t = new ETA.Models.GeneralTask
                    {
                        업무명   = txtName.Text.Trim(),
                        내용     = txtContent.Text?.Trim() ?? "",
                        배정자   = MainPage.CurrentEmployeeId,
                        담당자id = agentId,
                        담당자명 = agentName,
                        마감일   = txtDeadline.Text?.Trim() ?? "",
                        등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    };
                    ETA.Services.Common.GeneralTaskService.Insert(t);
                }
                onDone();
            };
            btnRow.Children.Add(btnSave);

            if (isEdit)
            {
                var btnDel = new Button
                {
                    Content = "🗑 삭제", Height = 26, Padding = new Thickness(8, 0),
                    FontFamily = agentFont, FontSize = AppTheme.FontSM,
                    Background = AppTheme.BgDanger,
                    Foreground = AppTheme.FgDanger,
                    BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
                };
                btnDel.Click += (_, _) =>
                {
                    ETA.Services.Common.GeneralTaskService.Delete(existing!.Id);
                    onDone();
                };
                btnRow.Children.Add(btnDel);
            }

            var btnCancel = new Button
            {
                Content = "← 취소", Height = 26, Padding = new Thickness(8, 0),
                FontFamily = agentFont, FontSize = AppTheme.FontSM,
                Background = AppTheme.BorderSubtle,
                Foreground = AppTheme.FgMuted,
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            };
            btnCancel.Click += (_, _) => onDone();
            btnRow.Children.Add(btnCancel);

            form.Children.Add(btnRow);
            generalTaskStack.Children.Add(form);
        }

        void LoadGeneralTaskList()
        {
            generalTaskStack.Children.Clear();
            try
            {
                var tasks = ETA.Services.Common.GeneralTaskService.GetAll();
                // 업무명 기준 그룹화 — 같은 업무에 여러 명 배정 가능
                var grouped = tasks.GroupBy(t => t.업무명).OrderBy(g => g.Key).ToList();
                var agentFont = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
                var agentFontM = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
                bool odd = false;

                foreach (var g in grouped)
                {
                    var row = new Border
                    {
                        Background = new SolidColorBrush(Color.Parse(odd ? "#1a1a28" : "#1e1e30")),
                        Padding    = new Thickness(8, 5),
                        Cursor     = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    };
                    odd = !odd;

                    var inner = new StackPanel { Spacing = 2 };

                    // 업무명
                    inner.Children.Add(new TextBlock
                    {
                        Text              = g.Key,
                        FontSize          = AppTheme.FontBase,
                        FontWeight        = FontWeight.SemiBold,
                        Foreground        = AppTheme.FgSecondary,
                        FontFamily        = agentFontM,
                        TextTrimming      = TextTrimming.CharacterEllipsis,
                    });

                    // 배정인원 표시
                    var chipsPanel = new WrapPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 0) };
                    foreach (var t in g)
                    {
                        if (string.IsNullOrWhiteSpace(t.담당자명)) continue;
                        var (bg, fg) = BadgeColorHelper.GetBadgeColor(t.담당자명);
                        var statusColor = t.상태 switch
                        {
                            "완료"  => "#2a4a2a",
                            "진행"  or "진행중" => "#2a3a4a",
                            _       => "#3a2a3a",
                        };
                        chipsPanel.Children.Add(new Border
                        {
                            Background      = new SolidColorBrush(Color.Parse(statusColor)),
                            CornerRadius    = new CornerRadius(8),
                            Padding         = new Thickness(6, 1),
                            Margin          = new Thickness(0, 0, 3, 2),
                            Child = new TextBlock
                            {
                                Text       = $"{t.담당자명} ({t.상태})",
                                FontSize   = AppTheme.FontXS,
                                FontFamily = agentFont,
                                Foreground = new SolidColorBrush(Color.Parse(fg)),
                            },
                        });
                    }
                    if (chipsPanel.Children.Count == 0)
                        chipsPanel.Children.Add(new TextBlock
                        {
                            Text = "미배정", FontSize = AppTheme.FontXS, FontFamily = agentFont,
                            Foreground = AppTheme.FgDimmed,
                        });
                    inner.Children.Add(chipsPanel);

                    row.Child = inner;

                    // 클릭 시 수정 폼 — 그룹의 첫 번째 항목 수정
                    var firstTask = g.First();
                    row.PointerPressed += (_, _) => ShowInlineTaskForm(firstTask, () => LoadGeneralTaskList());

                    generalTaskStack.Children.Add(row);
                }

                if (grouped.Count == 0)
                    generalTaskStack.Children.Add(new TextBlock
                    {
                        Text = "일반업무 없음", FontSize = AppTheme.FontBase,
                        Foreground = AppTheme.FgDimmed,
                        Margin = new Thickness(8, 8),
                        FontFamily = agentFont,
                    });

                // 하단: 추가 버튼
                var btnAdd = new Button
                {
                    Content         = "＋ 업무 추가",
                    FontSize        = AppTheme.FontBase,
                    Height          = 28,
                    Padding         = new Thickness(12, 0),
                    BorderThickness = new Thickness(0),
                    CornerRadius    = new CornerRadius(4),
                    FontFamily      = agentFontM,
                    Background      = AppTheme.BgActiveGreen,
                    Foreground      = AppTheme.FgSuccess,
                    Margin          = new Thickness(4, 6),
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left,
                };
                btnAdd.Click += (_, _) => ShowInlineTaskForm(null, () => LoadGeneralTaskList());
                generalTaskStack.Children.Add(btnAdd);
            }
            catch { }
        }

        var btnGenTask = new Button
        {
            Content         = "일반업무",
            FontSize        = AppTheme.FontBase,
            Height          = 26,
            Padding         = new Thickness(12, 0),
            BorderThickness = new Thickness(0),
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
            Foreground      = AppTheme.FgMuted,
        };

        var tabRow = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing     = 2,
            Margin      = new Thickness(0, 0, 0, 2),
        };
        tabRow.Children.Add(btnAnalysis);
        tabRow.Children.Add(btnContract);
        tabRow.Children.Add(btnGenTask);

        // 콘텐츠 영역 — 토글로 전환
        var contentArea = new Grid();
        contentArea.Children.Add(dock);
        contentArea.Children.Add(contractScroll);
        contentArea.Children.Add(generalTaskScroll);
        dock.IsVisible             = true;
        contractScroll.IsVisible   = false;
        generalTaskScroll.IsVisible = false;

        void SetActiveTab(Button active)
        {
            foreach (var btn in new[] { btnAnalysis, btnContract, btnGenTask })
            {
                bool on = btn == active;
                btn.Background = on ? Res("TabActiveBg") : Res("SubBtnBg");
                btn.Foreground = on ? Res("TabActiveFg") : Res("FgMuted");
            }
        }

        btnAnalysis.Click += (_, _) =>
        {
            dock.IsVisible             = true;
            contractScroll.IsVisible   = false;
            generalTaskScroll.IsVisible = false;
            SetActiveTab(btnAnalysis);
        };
        btnContract.Click += (_, _) =>
        {
            dock.IsVisible             = false;
            contractScroll.IsVisible   = true;
            generalTaskScroll.IsVisible = false;
            SetActiveTab(btnContract);
            LoadContracts();
        };
        btnGenTask.Click += (_, _) =>
        {
            dock.IsVisible             = false;
            contractScroll.IsVisible   = false;
            generalTaskScroll.IsVisible = true;
            SetActiveTab(btnGenTask);
            LoadGeneralTaskList();
        };

        var outerDock = new DockPanel { LastChildFill = true };
        DockPanel.SetDock(tabRow, Dock.Top);
        outerDock.Children.Add(tabRow);
        outerDock.Children.Add(contentArea);

        _content4Container = outerDock;
        return _content4Container;
    }

    private void LoadAnalysisItems()
    {
        if (_analysisItemsListBox == null) return;

        _analysisItemsListBox.Items.Clear();

        // 분장표준처리 테이블 컬럼 순서 + 약칭(row2)으로 가져오기 (비분석항목 제외)
        var analytes = AnalysisRequestService.GetOrderedAnalytes()
            .Where(a => !a.fullName.StartsWith("_")
                     && a.fullName != "기타업무"
                     && a.fullName != "담당계약업체"
                     && a.fullName != "항목명")
            .ToList();

        foreach (var (fullName, shortName) in analytes)
        {
            var (badgeBg, badgeFg) = BadgeColorHelper.GetBadgeColor(shortName);

            // 카드 UI: Border + Grid로 구성
            var cardBorder = new Border
            {
                Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
                BorderThickness = new Thickness(1),
                BorderBrush     = AppTheme.BorderMuted,
                CornerRadius    = new CornerRadius(3),
                Padding         = new Thickness(6, 2),
                Margin          = new Thickness(0, 1),
                Cursor          = new Cursor(StandardCursorType.Hand),
            };
            ToolTip.SetTip(cardBorder, $"{fullName} ({shortName})");

            var cardGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"),
            };

            // Col 0: 약칭 배지
            var catBadge = new Border
            {
                Background        = new SolidColorBrush(Color.Parse(badgeBg)),
                CornerRadius      = new CornerRadius(3),
                Padding           = new Thickness(4, 1),
                Margin            = new Thickness(0, 0, 5, 0),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text       = shortName,
                    FontSize   = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                }
            };
            Grid.SetColumn(catBadge, 0);
            cardGrid.Children.Add(catBadge);

            // Col 1: 항목명
            var nameBlock = new TextBlock
            {
                Text              = fullName,
                FontSize          = AppTheme.FontBase,
                FontWeight        = FontWeight.SemiBold,
                Foreground        = Res("AppFg"),
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                TextTrimming      = TextTrimming.CharacterEllipsis,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            Grid.SetColumn(nameBlock, 1);
            cardGrid.Children.Add(nameBlock);

            // Col 2: 할당 배지 (UpdateAssignmentInfo에서 갱신)
            var assignBadge = new TextBlock
            {
                Text              = "·",
                FontSize          = AppTheme.FontXS,
                Foreground        = Res("FgMuted"),
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Margin            = new Thickness(4, 0, 0, 0),
            };
            Grid.SetColumn(assignBadge, 2);
            cardGrid.Children.Add(assignBadge);

            cardBorder.Tag = fullName;
            cardBorder.Child = cardGrid;

            // ✅ 중요: Border에 직접 PointerPressed 이벤트 등록 (ListBox 이벤트와 상충 방지)
            cardBorder.PointerPressed += (s, e) =>
            {
                if (e.GetCurrentPoint(cardBorder).Properties.IsLeftButtonPressed)
                {
                    var analyte = (cardBorder.Tag as string) ?? "";
                    if (!string.IsNullOrEmpty(analyte))
                    {
                        e.Handled = true;  // 이벤트 버블링 차단 → 중복 드래그 방지
                        var data = new DataObject();
                        data.Set("analyte", analyte);
                        DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                    }
                }
            };

            var listBoxItem = new ListBoxItem
            {
                Content = cardBorder,
                Tag = fullName,
            };

            _analysisItemsListBox.Items.Add(listBoxItem);
        }

        // 할당 정보 비동기로 업데이트
        UpdateAssignmentInfo();
    }

    private static string GetAgentCategoryBg(string cat) => cat.Trim() switch
    {
        var c when c.Contains("유기")                       => "#1a2a3a",
        var c when c.Contains("무기")                       => "#2a1a3a",
        var c when c.Contains("부유")                       => "#1a3a2a",
        var c when c.Contains("질소") || c.Contains("인")   => "#3a2a1a",
        var c when c.Contains("금속")                       => "#2a3a1a",
        var c when c.Contains("대장") || c.Contains("세균") => "#3a1a1a",
        var c when c.Contains("pH")  || c.Contains("수소") => "#1a3a3a",
        _                                                   => "#2a2a3a"
    };
    private static string GetAgentCategoryFg(string cat) => cat.Trim() switch
    {
        var c when c.Contains("유기")                       => "#88aaff",
        var c when c.Contains("무기")                       => "#cc88ff",
        var c when c.Contains("부유")                       => "#88ccaa",
        var c when c.Contains("질소") || c.Contains("인")   => "#ccaa88",
        var c when c.Contains("금속")                       => "#aacc88",
        var c when c.Contains("대장") || c.Contains("세균") => "#ff8888",
        var c when c.Contains("pH")  || c.Contains("수소") => "#88ddcc",
        _                                                   => "#aaaacc"
    };

    /// <summary>각 카드에 할당 정보를 표시 (현재 선택된 기간 기준)</summary>
    private void UpdateAssignmentInfo()
    {
        if (_analysisItemsListBox == null) return;

        foreach (var listBoxItem in _analysisItemsListBox.Items.OfType<ListBoxItem>())
        {
            if (listBoxItem.Content is Border border && border.Child is Grid grid)
            {
                var analyte = listBoxItem.Tag as string;
                if (string.IsNullOrEmpty(analyte)) continue;

                var assignees = _content4QueryStart == _content4QueryEnd
                    ? AnalysisRequestService.GetAssigneesForAnalyteOnDate(analyte, _content4QueryStart)
                    : AnalysisRequestService.GetAssigneesForAnalyteInRange(analyte, _content4QueryStart, _content4QueryEnd);

                var assignBadge = grid.Children.OfType<TextBlock>()
                    .FirstOrDefault(tb => Grid.GetColumn(tb) == 2);
                if (assignBadge != null)
                {
                    var baseTip = ToolTip.GetTip(border)?.ToString()?.Split('\n')[0] ?? "";
                    if (assignees.Count == 0)
                    {
                        assignBadge.Text       = "미할당";
                        assignBadge.Foreground = Res("FgMuted");
                        ToolTip.SetTip(border, baseTip);
                    }
                    else
                    {
                        // 이름별 일수 집계 (중복 포함)
                        var grouped = assignees
                            .GroupBy(a => a)
                            .OrderByDescending(g => g.Count())
                            .ToList();
                        int uniqueCount = grouped.Count;

                        // 배지: 이름을 모두 나열
                        assignBadge.Text       = string.Join(", ", grouped.Select(g => g.Key));
                        assignBadge.Foreground = AppTheme.FgSuccess;

                        // 툴팁: 이름(N일) 형식으로 중복 표시
                        bool isRange = _content4QueryStart != _content4QueryEnd;
                        var detail = isRange
                            ? string.Join(", ", grouped.Select(g =>
                                g.Count() > 1 ? $"{g.Key}({g.Count()}일)" : g.Key))
                            : string.Join(", ", grouped.Select(g => g.Key));
                        ToolTip.SetTip(border, $"{baseTip}\n담당: {detail}");
                    }
                }
            }
        }
    }

    private void OnAnalysisItemPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            // Content4의 카드에서 드래그 시작
            var listBox = sender as ListBox;
            if (listBox == null) return;

            // 클릭된 control을 따라가서 ListBoxItem 찾기
            var source = e.Source as Control;
            ListBoxItem? targetItem = null;

            // visual tree를 따라 올라가면서 ListBoxItem 찾기
            var current = source;
            while (current != null)
            {
                if (current is ListBoxItem lbi)
                {
                    targetItem = lbi;
                    break;
                }
                current = current.Parent as Control;
            }

            if (targetItem != null && targetItem.Tag is string analyte)
            {
                // 드래그 데이터 설정
                var data = new DataObject();
                data.Set("analyte", analyte);
                DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
            }
        }
    }


    // ══════════════════════════════════════════════════════════════════════
    //  프로그레스 오버레이 헬퍼 + shimmer 애니메이션
    // ══════════════════════════════════════════════════════════════════════

    private DispatcherTimer? _shimmerTimer;

    public void ShowProgress(string message)
    {
        if (txtProgressMsg != null) txtProgressMsg.Text = message;
        if (pbGlobal != null) pbGlobal.IsIndeterminate = true;
        if (progressOverlay != null)
        {
            progressOverlay.Opacity = 0;
            progressOverlay.IsVisible = true;
            progressOverlay.Opacity = 1;
        }
        StartShimmer();
    }

    public void HideProgress()
    {
        StopShimmer();
        if (progressOverlay != null)
            progressOverlay.IsVisible = false;
    }

    private void StartShimmer()
    {
        var shimmer = this.FindControl<Border>("shimmerOverlay");
        if (shimmer == null) return;

        _shimmerTimer?.Stop();
        var transform = shimmer.RenderTransform as TranslateTransform ?? new TranslateTransform(-300, 0);
        shimmer.RenderTransform = transform;

        _shimmerTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(16) };
        int frame = 0;
        _shimmerTimer.Tick += (_, _) =>
        {
            frame++;
            double t = (double)(frame % 110) / 110; // 110 frames = ~1.8초
            double easedT = t < 0.5 ? 4 * t * t * t : 1 - Math.Pow(-2 * t + 2, 3) / 2;
            transform.X = -300 + 600 * easedT;
        };
        _shimmerTimer.Start();
    }

    private void StopShimmer()
    {
        _shimmerTimer?.Stop();
        _shimmerTimer = null;
    }

    public async Task RunWithProgress(string message, Func<Task> work)
    {
        ShowProgress(message);
        try { await Task.Run(work); }
        finally { HideProgress(); }
    }

    public async Task RunWithProgress(string message, Func<Action<int, string>, Task> work)
    {
        ShowProgress(message);
        if (pbGlobal != null)
        {
            pbGlobal.IsIndeterminate = false;
            pbGlobal.Maximum = 100;
        }
        try
        {
            await Task.Run(() => work((pct, msg) =>
            {
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    if (pbGlobal != null) pbGlobal.Value = pct;
                    if (txtProgressMsg != null) txtProgressMsg.Text = msg;
                });
            }));
        }
        finally
        {
            if (pbGlobal != null) pbGlobal.IsIndeterminate = true;
            HideProgress();
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  레이아웃 헬퍼
    // ══════════════════════════════════════════════════════════════════════

    private void SetLeftPanelWidth(double width)
    {
        var grid = this.FindControl<Grid>("MainSplitGrid");
        if (grid != null && grid.ColumnDefinitions.Count > 0)
            grid.ColumnDefinitions[0].Width = new GridLength(width);
    }

    private void SetContentLayout(
        double content2Star = 1, double content4Star = 1,
        double upperStar   = 4, double lowerStar    = 1)
    {
        // 다른 모드 진입 시 하단분할 레이아웃 복원
        ResetDefaultSplitLayout();

        var rightGrid = this.FindControl<Grid>("RightSplitGrid");
        if (rightGrid == null) return;

        // 행: Show2(상) / 스플리터 / Show3(하)
        if (rightGrid.RowDefinitions.Count >= 3)
        {
            rightGrid.RowDefinitions[0].Height = new GridLength(upperStar, GridUnitType.Star);
            rightGrid.RowDefinitions[2].Height = lowerStar > 0
                ? new GridLength(lowerStar, GridUnitType.Star)
                : new GridLength(0);

            var lowerBorder = this.FindControl<Border>("LowerBorder");
            if (lowerBorder != null) lowerBorder.IsVisible = lowerStar > 0;

            var vSplitter = this.FindControl<GridSplitter>("VerticalSplitter");
            if (vSplitter != null) vSplitter.IsVisible = lowerStar > 0;

            var toggle = this.FindControl<Button>("btnToggleShow3");
            if (toggle != null) toggle.IsVisible = lowerStar > 0 || _show3Collapsed;
        }

        // 열: Show2+Show3(좌) / 스플리터 / Show4(우)
        if (rightGrid.ColumnDefinitions.Count >= 3)
        {
            rightGrid.ColumnDefinitions[0].Width = new GridLength(content2Star, GridUnitType.Star);
            rightGrid.ColumnDefinitions[2].Width = content4Star > 0
                ? new GridLength(content4Star, GridUnitType.Star)
                : new GridLength(0);

            var content4Border = this.FindControl<Border>("Content4Border");
            if (content4Border != null) content4Border.IsVisible = content4Star > 0;

            var hSplitter = this.FindControl<GridSplitter>("HorizontalSplitter");
            if (hSplitter != null) hSplitter.IsVisible = content4Star > 0;
        }
    }

    /// <summary>Show4가 숨겨져 있으면 최소 비율로 보이게 복원</summary>
    private void EnsureShow4Visible()
    {
        var rightGrid = this.FindControl<Grid>("RightSplitGrid");
        if (rightGrid == null || rightGrid.ColumnDefinitions.Count < 3) return;

        var col2 = rightGrid.ColumnDefinitions[2];
        if (col2.Width.Value < 0.1)
        {
            col2.Width = new GridLength(2, GridUnitType.Star);
            var c4 = this.FindControl<Border>("Content4Border");
            if (c4 != null) c4.IsVisible = true;
            var hs = this.FindControl<GridSplitter>("HorizontalSplitter");
            if (hs != null) hs.IsVisible = true;
        }
    }

    private bool _bottomSplitActive;

    /// <summary>Show2를 전체폭 상단, Show3+Show4를 하단 병렬로 배치</summary>
    private void ApplyBottomSplitLayout()
    {
        var rightGrid = this.FindControl<Grid>("RightSplitGrid");
        var c2 = this.FindControl<Border>("Content2Border");
        var c4 = this.FindControl<Border>("Content4Border");
        var lower = this.FindControl<Border>("LowerBorder");
        var hSplitter = this.FindControl<GridSplitter>("HorizontalSplitter");
        if (rightGrid == null || c2 == null || c4 == null || lower == null) return;

        // Show2: Row=0, 전체 컬럼
        Grid.SetColumnSpan(c2, 3);

        // Show4: Row=2, Col=2 (하단 우측)
        Grid.SetRow(c4, 2);
        Grid.SetColumn(c4, 2);
        c4.IsVisible = true;
        c4.Margin = new Thickness(1, 0, 0, 0);

        // Show3: Row=2, Col=0 (하단 좌측, ColumnSpan 1)
        Grid.SetColumnSpan(lower, 1);

        // 상단 좌우 스플리터 숨기기, 하단에서 좌우 분할 사용
        if (hSplitter != null)
        {
            Grid.SetRow(hSplitter, 2);
            hSplitter.IsVisible = true;
        }

        // 컬럼 비율: Show3(좌) 5 : Show4(우) 2
        if (rightGrid.ColumnDefinitions.Count >= 3)
        {
            rightGrid.ColumnDefinitions[0].Width = new GridLength(5, GridUnitType.Star);
            rightGrid.ColumnDefinitions[2].Width = new GridLength(2, GridUnitType.Star);
        }

        _bottomSplitActive = true;
    }

    /// <summary>기본 레이아웃 복원 (Show2+Show4 상단, Show3 하단 전체폭)</summary>
    private void ResetDefaultSplitLayout()
    {
        if (!_bottomSplitActive) return;

        var rightGrid = this.FindControl<Grid>("RightSplitGrid");
        var c2 = this.FindControl<Border>("Content2Border");
        var c4 = this.FindControl<Border>("Content4Border");
        var lower = this.FindControl<Border>("LowerBorder");
        var hSplitter = this.FindControl<GridSplitter>("HorizontalSplitter");
        if (rightGrid == null || c2 == null || c4 == null || lower == null) return;

        // Show2: Row=0, Col=0
        Grid.SetColumnSpan(c2, 1);

        // Show4: Row=0, Col=2 (상단 우측)
        Grid.SetRow(c4, 0);
        Grid.SetColumn(c4, 2);
        c4.Margin = new Thickness(1, 0, 0, 0);

        // Show3: Row=2, ColumnSpan=3 (하단 전체폭)
        Grid.SetColumnSpan(lower, 3);

        // 스플리터 원위치
        if (hSplitter != null)
        {
            Grid.SetRow(hSplitter, 0);
            hSplitter.IsVisible = true;
        }

        _bottomSplitActive = false;
    }


    // ══════════════════════════════════════════════════════════════════════
    //  메뉴 클릭
    // ══════════════════════════════════════════════════════════════════════

    private void MyTask_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "MyTask";

        if (_myTaskPage == null)
        {
            _myTaskPage = new MyTaskPage();
            _myTaskPage.DetailPanelChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _myTaskPage.EditPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
            _myTaskPage.StatsPanelChanged += panel =>
            {
                Show4.Content = panel;
                LogContentChange("Show4", panel);
            };
        }

        Show1.Content = _myTaskPage;
        LogContentChange("Show1", _myTaskPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = null;
        LogContentChange("Show4", null);

        _myTaskPage.LoadData();
        SetSubMenu("새로고침", "", "", "", "", "", "");
        SetLeftPanelWidth(280);
        SetContentLayout(content2Star: 6, content4Star: 3, upperStar: 6, lowerStar: 4);
        RestoreModeLayout("MyTask");
    }

    private void Agent_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Agent";

        if (_agentTreePage == null)
        {
            _agentTreePage = new AgentTreePage();

            // Show2 영속 래퍼: 직원 선택 시 전체 슬라이드 대신 CrossFade만 적용
            _show2AgentWrapper ??= new TransitioningContentControl
            {
                PageTransition = new CrossFade(TimeSpan.FromMilliseconds(200))
            };
            _agentTreePage.DetailPanelChanged += panel =>
            {
                _show2AgentWrapper!.Content = panel;
                LogContentChange("Show2(inner)", panel);
            };
        }

        // Show3 영속 래퍼: 타임라인 차트도 CrossFade 전환
        _show3AgentWrapper ??= new TransitioningContentControl
        {
            PageTransition = new CrossFade(TimeSpan.FromMilliseconds(200))
        };
        _agentTreePage.Show3ContentRequest = ctrl =>
        {
            _show3AgentWrapper!.Content = ctrl;
            LogContentChange("Show3(inner)", ctrl);
        };
        // Show4: 영구 wrapper ContentControl 안에서 내용만 교체 (TransitioningContentControl 재사용 버그 방지)
        if (_show4AgentWrapper == null)
            _show4AgentWrapper = new ContentControl();
        _show4AgentWrapper.Content = CreateAnalysisItemsListBox();  // 기본: 분석항목 컨테이너

        _agentTreePage.Show4ContentRequest = ctrl =>
        {
            _show4AgentWrapper!.Content = ctrl ?? _content4Container;
            LogContentChange("Show4", ctrl ?? _content4Container as Control);
        };

        Show1.Content = _agentTreePage;
        LogContentChange("Show1", _agentTreePage);
        // Show2/Show3: 영속 래퍼 사용 → 직원 선택 시 부드러운 CrossFade만 적용
        _show2AgentWrapper!.Content = null;
        Show2.Content = _show2AgentWrapper;
        LogContentChange("Show2", _show2AgentWrapper);
        _show3AgentWrapper!.Content = null;
        Show3.Content = _show3AgentWrapper;
        LogContentChange("Show3", _show3AgentWrapper);

        // ✅ Content4: wrapper를 Show4에 고정 (이후 내용은 wrapper.Content 교체)
        Show4.Content = _show4AgentWrapper;
        LogContentChange("Show4", _analysisItemsListBox);

        // AgentTreePage에 ListBox 참조 전달
        _agentTreePage.AnalysisItemsListBox = _analysisItemsListBox;
        _agentTreePage.Show4RefreshRequest = () => UpdateAssignmentInfo();
        
        _agentTreePage.LoadData();
        _bt1SaveAction = _agentTreePage.SaveSelected;

        SetSubMenu("저장", "새로고침", "직원 추가", "", "엑셀 내보내기", "인쇄", "업무분장표");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Agent");
    }

    private void WasteCompany_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteCompany";

        if (_wasteCompanyPage == null)
        {
            _wasteCompanyPage = new WasteCompanyPage();

            // 업소정보(편집 패널) → Show3
            _wasteCompanyPage.DetailPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };

            // 업체 선택 → Show2(차트) + Show4(목록) 자료조회 기능 직접 구현
            _wasteCompanyPage.CompanySelected += company =>
            {
                List<ETA.Models.WasteAnalysisResult> results;
                try { results = ETA.Services.SERVICE2.WasteDataService.GetResults(company.업체명); }
                catch (Exception ex)
                {
                    Show2.Content = new TextBlock
                    {
                        Text = $"조회 오류: {ex.Message}",
                        FontSize = AppTheme.FontBase, FontFamily = _wasteFont,
                        Foreground = AppTheme.FgDanger,
                        Margin = new Thickness(12),
                    };
                    return;
                }

                // Show2 = 분석결과 추이 차트
                Show2.Content = BuildWasteBarLinePanel(company.업체명, results);
                LogContentChange("Show2", Show2.Content as Control);

                // Show4 = 전 항목 분석결과 테이블
                Show4.Content = BuildAllResultsPanel(company.업체명, results, company.약칭);
                LogContentChange("Show4", Show4.Content as Control);
            };

        }

        Show1.Content = _wasteCompanyPage;
        LogContentChange("Show1", _wasteCompanyPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _ = _wasteCompanyPage.LoadDataAsync();
        _bt1SaveAction = _wasteCompanyPage.SaveSelected;

        SetSubMenu("저장", "새로고침", "업소 등록", "엑셀 업로드", "", "통계 보기", "설정");
        SetLeftPanelWidth(300);
        SetContentLayout(content2Star: 2, content4Star: 1, upperStar: 6, lowerStar: 4);

        RestoreModeLayout("WasteCompany");
    }

    // =========================================================================
    // 의뢰내역 리스트 패널 (Show4)
    // =========================================================================
    private static readonly FontFamily _wasteFont =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    private bool _wasteChartIsBar = true; // Bar/Line 토글 상태 유지
    private int  _wasteChartCount = 30; // 그래프 조회 수량

    private Control BuildOrderHistoryPanel()
    {
        var root = new StackPanel { Spacing = 0 };

        // 헤더
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "📋  의뢰내역",
                FontSize   = AppTheme.FontLG, FontWeight = FontWeight.SemiBold,
                FontFamily = _wasteFont, Foreground = AppTheme.FgMuted,
            }
        });

        // 날짜별 데이터
        var scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        var listPanel = new StackPanel { Spacing = 1 };

        try
        {
            var dates = ETA.Services.SERVICE2.WasteSampleService.GetDates();   // 역순
            foreach (var date in dates)
            {
                var rows = ETA.Services.SERVICE2.WasteSampleService.GetByDate(date);
                if (rows.Count == 0) continue;

                // 날짜 헤더
                DateTime.TryParse(date, out var d);
                listPanel.Children.Add(new Border
                {
                    Background = AppTheme.BgInput,
                    Padding    = new Thickness(8, 3),
                    Child = new TextBlock
                    {
                        Text       = $"{d:yyyy-MM-dd} ({DayKr(d)})  {rows.Count}건",
                        FontSize   = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
                        FontFamily = _wasteFont, Foreground = AppTheme.FgMuted,
                    }
                });

                // 각 행
                foreach (var r in rows)
                {
                    var (gColor, gBg) = r.구분 switch
                    {
                        "여수" => ("#88aacc", "#1a1e2a"),
                        "율촌" => ("#aaccaa", "#1a2a1a"),
                        "세풍" => ("#ccaa88", "#2a1e14"),
                        _      => ("#aaaaaa", "#1e1e1e"),
                    };

                    var rowGrid = new Grid
                    {
                        ColumnDefinitions = new ColumnDefinitions("70,50,*,60"),
                        Background        = Brush.Parse(gBg),
                        Margin            = new Thickness(0, 0, 0, 1),
                    };

                    void AddCell(int col, string text, string fg, double size = 0, FontWeight fw = FontWeight.Normal)
                    {
                        var tb = new TextBlock
                        {
                            Text      = text, FontSize = size > 0 ? size : AppTheme.FontMD, FontFamily = _wasteFont,
                            Foreground = Brush.Parse(fg), FontWeight = fw,
                            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                            Margin = new Thickness(6, 3),
                            TextTrimming = TextTrimming.CharacterEllipsis,
                        };
                        rowGrid.Children.Add(tb);
                        Grid.SetColumn(tb, col);
                    }

                    AddCell(0, r.SN,    "#aaccff");
                    AddCell(1, r.구분,  gColor);
                    AddCell(2, r.업체명, "#dddddd");
                    AddCell(3, r.확인자, "#888888");

                    listPanel.Children.Add(rowGrid);
                }
            }
        }
        catch (Exception ex)
        {
            listPanel.Children.Add(new TextBlock
            {
                Text = $"오류: {ex.Message}", FontSize = AppTheme.FontSM,
                FontFamily = _wasteFont, Foreground = AppTheme.FgDanger,
                Margin = new Thickness(8),
            });
        }

        scroll.Content = listPanel;

        return new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,*"),
            Children       = { root, new Border { [Grid.RowProperty] = 1, Child = scroll } },
        };
    }

    private static string DayKr(DateTime d) => d.DayOfWeek switch
    {
        DayOfWeek.Monday    => "월",
        DayOfWeek.Tuesday   => "화",
        DayOfWeek.Wednesday => "수",
        DayOfWeek.Thursday  => "목",
        DayOfWeek.Friday    => "금",
        DayOfWeek.Saturday  => "토",
        _                   => "일",
    };

    // =========================================================================
    // 자료 조회 — Show2(그래프) + Show3(목록)
    // =========================================================================
    private void ShowWasteCompanyData()
    {
        var company = _wasteCompanyPage?.SelectedCompany;
        if (company == null)
        {
            Show3.Content = new TextBlock
            {
                Text = "왼쪽 트리에서 업체를 먼저 선택하세요",
                FontSize = AppTheme.FontMD, FontFamily = _wasteFont,
                Foreground = AppTheme.FgMuted,
                Margin = new Thickness(12),
            };
            return;
        }

        List<ETA.Models.WasteAnalysisResult> results;
        try { results = ETA.Services.SERVICE2.WasteDataService.GetResults(company.업체명); }
        catch (Exception ex)
        {
            Show3.Content = new TextBlock
            {
                Text = $"조회 오류: {ex.Message}",
                FontSize = AppTheme.FontBase, FontFamily = _wasteFont,
                Foreground = AppTheme.FgDanger,
                Margin = new Thickness(12),
            };
            return;
        }

        // Show3 = 목록 테이블
        Show3.Content = BuildDataListPanel(company.업체명, results, company.약칭);
        LogContentChange("Show3", Show3.Content as Control);

        // 하단 패널 보이도록
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 6, lowerStar: 4);
    }

    private Control BuildWasteBarLinePanel(string 업체명, List<ETA.Models.WasteAnalysisResult> results)
    {
        int cnt = _wasteChartCount;
        var charts = new List<WasteSingleSeriesChart>();
        foreach (var (label, color, getValue) in WasteBarLineChartControl.Series)
        {
            var withValue = results.Where(r => getValue(r).HasValue).ToList();
            var recent    = withValue.Count > cnt ? withValue.Skip(withValue.Count - cnt).ToList() : withValue;
            charts.Add(new WasteSingleSeriesChart(label, color, getValue, recent));
        }

        // 상단 바: Bar/Line 토글
        var topBar = new WrapPanel { Margin = new Thickness(4, 3), VerticalAlignment = VerticalAlignment.Center };

        topBar.Children.Add(new TextBlock
        {
            Text = $"{업체명}  분석결과 추이", FontSize = AppTheme.FontBase, FontFamily = _wasteFont,
            Foreground = AppTheme.FgMuted, FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 12, 0),
        });

        var btnBar = new Avalonia.Controls.Button
        {
            Content = "Bar", FontSize = AppTheme.FontSM, FontFamily = _wasteFont,
            Padding = new Thickness(8, 2), Margin = new Thickness(0, 0, 2, 0),
            Background = _wasteChartIsBar ? AppTheme.BorderMuted : AppTheme.BorderSubtle,
            Foreground = _wasteChartIsBar ? AppTheme.FgPrimary : AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(3),
        };
        var btnLine = new Avalonia.Controls.Button
        {
            Content = "Line", FontSize = AppTheme.FontSM, FontFamily = _wasteFont,
            Padding = new Thickness(8, 2), Margin = new Thickness(0, 0, 8, 0),
            Background = _wasteChartIsBar ? AppTheme.BorderSubtle : AppTheme.BorderMuted,
            Foreground = _wasteChartIsBar ? AppTheme.FgMuted : AppTheme.FgPrimary,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(3),
        };

        // 이전 상태 복원
        if (!_wasteChartIsBar)
            foreach (var c in charts) c.SetBarMode(false);

        btnBar.Click += (_, _) =>
        {
            _wasteChartIsBar = true;
            foreach (var c in charts) c.SetBarMode(true);
            btnBar.Background  = AppTheme.BorderMuted; btnBar.Foreground  = AppTheme.FgPrimary;
            btnLine.Background = AppTheme.BorderSubtle; btnLine.Foreground = AppTheme.FgMuted;
        };
        btnLine.Click += (_, _) =>
        {
            _wasteChartIsBar = false;
            foreach (var c in charts) c.SetBarMode(false);
            btnLine.Background = AppTheme.BorderMuted; btnLine.Foreground = AppTheme.FgPrimary;
            btnBar.Background  = AppTheme.BorderSubtle; btnBar.Foreground  = AppTheme.FgMuted;
        };
        topBar.Children.Add(btnBar);
        topBar.Children.Add(btnLine);

        // 그래프 수량 드롭박스
        var cntCombo = new ComboBox
        {
            FontSize = AppTheme.FontSM, FontFamily = _wasteFont,
            Padding = new Thickness(4, 1), Margin = new Thickness(4, 0, 0, 0),
            MinWidth = 55, MaxHeight = 24,
            VerticalAlignment = VerticalAlignment.Center,
            ItemsSource = new[] { 10, 20, 30, 40, 50 },
            SelectedItem = _wasteChartCount,
        };
        cntCombo.SelectionChanged += (_, _) =>
        {
            if (cntCombo.SelectedItem is int v)
            {
                _wasteChartCount = v;
                Show2.Content = BuildWasteBarLinePanel(업체명, results);
            }
        };
        topBar.Children.Add(cntCombo);

        // 4열 × 2행 그리드
        var chartGrid = new Grid
        {
            RowDefinitions    = new RowDefinitions("*,*"),
            ColumnDefinitions = new ColumnDefinitions("*,*,*,*"),
            Margin = new Thickness(2),
        };
        for (int i = 0; i < charts.Count; i++)
        {
            var border = new Border
            {
                Child           = charts[i],
                BorderBrush     = AppTheme.BgSecondary,
                BorderThickness = new Thickness(0.5),
                Margin          = new Thickness(1),
            };
            Grid.SetRow(border, i / 4);
            Grid.SetColumn(border, i % 4);
            chartGrid.Children.Add(border);
        }

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
        var topBorder = new Border { Background = Brush.Parse("#16161e"), Child = topBar };
        root.Children.Add(topBorder);
        root.Children.Add(chartGrid);
        Grid.SetRow(chartGrid, 1);
        return root;
    }

    private Control BuildDataListPanel(string 업체명, List<ETA.Models.WasteAnalysisResult> results, string 약칭 = "")
    {
        var root = new StackPanel { Spacing = 0 };

        // ── 헤더 ────────────────────────────────────────────────────────────
        var titleRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 6,
            Margin      = new Thickness(10, 6),
        };
        if (!string.IsNullOrWhiteSpace(약칭))
        {
            var (bg, fg, bd) = WasteCompanyPage.GetChosungBadgeColorPublic(약칭);
            titleRow.Children.Add(new Border
            {
                Background      = Brush.Parse(bg),
                BorderBrush     = Brush.Parse(bd),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(8),
                Padding         = new Thickness(6, 2),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text       = 약칭,
                    FontSize   = AppTheme.FontSM, FontFamily = _wasteFont,
                    Foreground = Brush.Parse(fg),
                }
            });
        }
        titleRow.Children.Add(new TextBlock
        {
            Text       = $"📋  {업체명}  분석결과 내역  ({results.Count}건)",
            FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            FontFamily = _wasteFont, Foreground = AppTheme.FgMuted,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        });
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Child        = titleRow,
        });

        if (results.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "분석결과 없음",
                FontSize = AppTheme.FontBase, FontFamily = _wasteFont,
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(12, 8),
            });
            return root;
        }

        // ── 항목별 평균 계산 ─────────────────────────────────────────────────
        double? Avg(Func<ETA.Models.WasteAnalysisResult, double?> f)
        {
            var vals = results.Where(r => f(r).HasValue).Select(r => f(r)!.Value).ToList();
            return vals.Count > 0 ? vals.Average() : null;
        }
        string H(string name, string avg) => $"{name}\n({avg})";

        // ── 컬럼 헤더 (평균 포함) ────────────────────────────────────────────
        root.Children.Add(MakeDataRow(
            "채수일",
            H("BOD",         Fmt(Avg(r => r.BOD))),
            H("TOC(TC-IC)",  Fmt(Avg(r => r.TOC_TCIC))),
            H("TOC(NPOC)",   Fmt(Avg(r => r.TOC_NPOC))),
            H("SS",          Fmt(Avg(r => r.SS))),
            H("T-N",         Fmt(Avg(r => r.TN))),
            H("T-P",         FmtTP(Avg(r => r.TP))),
            H("Phenols",     Fmt(Avg(r => r.Phenols))),
            H("N-Hexan",     Fmt(Avg(r => r.NHexan))),
            isHeader: true));

        // ── 데이터 행 (역순: 최근 → 오래된 순) ─────────────────────────────
        bool alt = false;
        foreach (var r in results.AsEnumerable().Reverse())
        {
            root.Children.Add(MakeDataRow(
                r.채수일,
                Fmt(r.BOD),     Fmt(r.TOC_TCIC), Fmt(r.TOC_NPOC),
                Fmt(r.SS),      Fmt(r.TN),        FmtTP(r.TP),
                Fmt(r.Phenols), Fmt(r.NHexan),
                isHeader: false, alt: alt));
            alt = !alt;
        }

        return new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    private static string Fmt(double? v) =>
        v.HasValue ? v.Value.ToString("F1") : "—";

    private static string FmtTP(double? v) =>
        v.HasValue ? v.Value.ToString("F3") : "—";

    private static Border MakeDataRow(
        string 날짜, string bod, string tocTcic, string tocNpoc,
        string ss,   string tn,  string tp,      string phenols, string nhexan,
        bool isHeader, bool alt = false)
    {
        var cols = new[] { "100,*,*,*,*,*,*,*,*" };
        var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("100,*,*,*,*,*,*,*,*") };
        var bg   = isHeader ? "#22223a" : alt ? "#1a1e28" : "#161620";
        var fg   = isHeader ? "#8899bb" : "#cccccc";
        var fw   = isHeader ? FontWeight.SemiBold : FontWeight.Normal;

        void Cell(int col, string text)
        {
            var tb = new TextBlock
            {
                Text = text, FontSize = AppTheme.FontSM, FontFamily = _wasteFont,
                FontWeight = fw,
                Foreground = text == "—" ? AppTheme.BorderMuted : Brush.Parse(fg),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                HorizontalAlignment = col == 0 ? Avalonia.Layout.HorizontalAlignment.Left : Avalonia.Layout.HorizontalAlignment.Right,
                Margin = new Thickness(col == 0 ? 8 : 4, 3, col == 0 ? 4 : 8, 3),
                TextTrimming = isHeader ? TextTrimming.None : TextTrimming.CharacterEllipsis,
                TextWrapping = isHeader ? TextWrapping.Wrap : TextWrapping.NoWrap,
            };
            Grid.SetColumn(tb, col);
            grid.Children.Add(tb);
        }

        Cell(0, 날짜); Cell(1, bod);    Cell(2, tocTcic); Cell(3, tocNpoc);
        Cell(4, ss);   Cell(5, tn);     Cell(6, tp);      Cell(7, phenols);
        Cell(8, nhexan);

        return new Border
        {
            Background = Brush.Parse(bg),
            Child      = grid,
            Margin     = new Thickness(0, 0, 0, 1),
        };
    }

    // ── 전 항목 분석결과 테이블 (Show4) ─────────────────────────────────
    private Control BuildAllResultsPanel(string 업체명, List<ETA.Models.WasteAnalysisResult> results, string 약칭 = "")
    {
        var root = new StackPanel { Spacing = 0 };

        // 타이틀
        var titleRow = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing = 6,
            Margin = new Thickness(8, 6),
        };
        if (!string.IsNullOrEmpty(약칭))
        {
            var (bg, fg, bd) = WasteCompanyPage.GetChosungBadgeColorPublic(약칭);
            titleRow.Children.Add(new Border
            {
                Background      = Brush.Parse(bg),
                BorderBrush     = Brush.Parse(bd),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(8),
                Padding         = new Thickness(6, 2),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text       = 약칭,
                    FontSize   = AppTheme.FontSM, FontFamily = _wasteFont,
                    Foreground = Brush.Parse(fg),
                }
            });
        }
        titleRow.Children.Add(new TextBlock
        {
            Text       = $"{업체명}  분석결과 내역  ({results.Count}건)",
            FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            FontFamily = _wasteFont, Foreground = AppTheme.FgMuted,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        });
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Child        = titleRow,
        });

        if (results.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "분석결과 없음",
                FontSize = AppTheme.FontBase, FontFamily = _wasteFont,
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(12, 8),
            });
            return root;
        }

        double? Avg(Func<ETA.Models.WasteAnalysisResult, double?> f)
        {
            var vals = results.Where(r => f(r).HasValue).Select(r => f(r)!.Value).ToList();
            return vals.Count > 0 ? vals.Average() : null;
        }

        // 컬럼 정의: (헤더, 셀렉터)
        var cols = new (string label, Func<ETA.Models.WasteAnalysisResult, double?> sel)[]
        {
            ("BOD",     r => r.BOD),
            ("TOC",     r => r.TOC_TCIC ?? r.TOC_NPOC),
            ("SS",      r => r.SS),
            ("T-N",     r => r.TN),
            ("T-P",     r => r.TP),
            ("Phenols", r => r.Phenols),
            ("N-Hexan", r => r.NHexan),
        };

        // 헤더 행
        root.Children.Add(MakeAllRow("분석일", cols.Select(c => $"{c.label}\n({Fmt(Avg(c.sel))})").ToArray(), isHeader: true));

        // 데이터 행 (최근→과거)
        bool alt = false;
        foreach (var r in results.AsEnumerable().Reverse())
        {
            root.Children.Add(MakeAllRow(r.채수일, cols.Select(c => Fmt(c.sel(r))).ToArray(), isHeader: false, alt: alt));
            alt = !alt;
        }

        return new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    private static Border MakeAllRow(string 날짜, string[] cells, bool isHeader, bool alt = false)
    {
        var colDefs = "100," + string.Join(",", Enumerable.Repeat("*", cells.Length));
        var grid = new Grid { ColumnDefinitions = new ColumnDefinitions(colDefs) };
        var bg = isHeader ? "#22223a" : alt ? "#1a1e28" : "#161620";
        var fg = isHeader ? "#8899bb" : "#cccccc";
        var fw = isHeader ? FontWeight.SemiBold : FontWeight.Normal;

        void Cell(int col, string text)
        {
            var tb = new TextBlock
            {
                Text = text, FontSize = AppTheme.FontSM, FontFamily = _wasteFont,
                FontWeight = fw,
                Foreground = text == "—" ? AppTheme.BorderMuted : Brush.Parse(fg),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                HorizontalAlignment = col == 0
                    ? Avalonia.Layout.HorizontalAlignment.Left
                    : Avalonia.Layout.HorizontalAlignment.Right,
                Margin = new Thickness(col == 0 ? 8 : 2, 4, col == 0 ? 4 : 6, 4),
                TextWrapping = isHeader ? TextWrapping.Wrap : TextWrapping.NoWrap,
            };
            Grid.SetColumn(tb, col);
            grid.Children.Add(tb);
        }

        Cell(0, 날짜);
        for (int i = 0; i < cells.Length; i++) Cell(i + 1, cells[i]);

        return new Border
        {
            Background = Brush.Parse(bg),
            Child      = grid,
            Margin     = new Thickness(0, 0, 0, 1),
        };
    }

    // ── TOC 전용 패널 (Show4) ────────────────────────────────────────────
    private Control BuildTocListPanel(string 업체명, List<ETA.Models.WasteAnalysisResult> results, string 약칭 = "")
    {
        var root = new StackPanel { Spacing = 0 };

        // ── 타이틀 ──────────────────────────────────────────────────────
        var titleRow = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing = 6,
            Margin = new Thickness(8, 6),
        };
        if (!string.IsNullOrEmpty(약칭))
        {
            var (bg, fg, bd) = WasteCompanyPage.GetChosungBadgeColorPublic(약칭);
            titleRow.Children.Add(new Border
            {
                Background      = Brush.Parse(bg),
                BorderBrush     = Brush.Parse(bd),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(8),
                Padding         = new Thickness(6, 2),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text       = 약칭,
                    FontSize   = AppTheme.FontSM, FontFamily = _wasteFont,
                    Foreground = Brush.Parse(fg),
                }
            });
        }
        titleRow.Children.Add(new TextBlock
        {
            Text       = $"{업체명}  TOC 내역  ({results.Count}건)",
            FontSize   = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            FontFamily = _wasteFont, Foreground = AppTheme.FgMuted,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        });
        root.Children.Add(new Border
        {
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Child        = titleRow,
        });

        if (results.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "TOC 데이터 없음",
                FontSize = AppTheme.FontBase, FontFamily = _wasteFont,
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(12, 8),
            });
            return root;
        }

        // ── 평균 계산 ────────────────────────────────────────────────────
        double? AvgT(Func<ETA.Models.WasteAnalysisResult, double?> f)
        {
            var vals = results.Where(r => f(r).HasValue).Select(r => f(r)!.Value).ToList();
            return vals.Count > 0 ? vals.Average() : null;
        }

        // ── 헤더 행 ──────────────────────────────────────────────────────
        root.Children.Add(MakeTocRow("분석일",
            $"TOC(TC-IC)\n({Fmt(AvgT(r => r.TOC_TCIC))})",
            $"TOC(NPOC)\n({Fmt(AvgT(r => r.TOC_NPOC))})",
            isHeader: true));

        // ── 데이터 행 (최근→과거) ────────────────────────────────────────
        bool alt = false;
        foreach (var r in results.AsEnumerable().Reverse())
        {
            root.Children.Add(MakeTocRow(r.채수일,
                Fmt(r.TOC_TCIC), Fmt(r.TOC_NPOC),
                isHeader: false, alt: alt));
            alt = !alt;
        }

        return new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    private static Border MakeTocRow(string 날짜, string tcic, string npoc,
        bool isHeader, bool alt = false)
    {
        var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("120,*,*") };
        var bg = isHeader ? "#22223a" : alt ? "#1a1e28" : "#161620";
        var fg = isHeader ? "#8899bb" : "#cccccc";
        var fw = isHeader ? FontWeight.SemiBold : FontWeight.Normal;

        void Cell(int col, string text)
        {
            var tb = new TextBlock
            {
                Text = text, FontSize = AppTheme.FontSM, FontFamily = _wasteFont,
                FontWeight = fw,
                Foreground = text == "—" ? AppTheme.BorderMuted : Brush.Parse(fg),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                HorizontalAlignment = col == 0
                    ? Avalonia.Layout.HorizontalAlignment.Left
                    : Avalonia.Layout.HorizontalAlignment.Right,
                Margin = new Thickness(col == 0 ? 8 : 4, 4, col == 0 ? 4 : 8, 4),
                TextWrapping = isHeader ? TextWrapping.Wrap : TextWrapping.NoWrap,
            };
            Grid.SetColumn(tb, col);
            grid.Children.Add(tb);
        }

        Cell(0, 날짜); Cell(1, tcic); Cell(2, npoc);

        return new Border
        {
            Background = Brush.Parse(bg),
            Child      = grid,
            Margin     = new Thickness(0, 0, 0, 1),
        };
    }

    private void WasteNameReconcile_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteNameReconcile";

        _wasteNameReconcilePage ??= new WasteNameReconcilePage();

        Show1.Content = _wasteNameReconcilePage.LeftPanel;
        LogContentChange("Show1", _wasteNameReconcilePage.LeftPanel);
        Show2.Content = _wasteNameReconcilePage.CenterPanel;
        LogContentChange("Show2", _wasteNameReconcilePage.CenterPanel);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = _wasteNameReconcilePage.RightPanel;
        LogContentChange("Show4", _wasteNameReconcilePage.RightPanel);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "", "", "", "", "설정");
        SetLeftPanelWidth(280);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);

        RestoreModeLayout("WasteNameReconcile");
    }

    private void WaterQualityNameReconcile_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WaterQualityNameReconcile";

        _waterQualityNameReconcilePage ??= new WaterQualityNameReconcilePage();

        Show1.Content = _waterQualityNameReconcilePage.LeftPanel;
        LogContentChange("Show1", _waterQualityNameReconcilePage.LeftPanel);
        Show2.Content = _waterQualityNameReconcilePage.CenterPanel;
        LogContentChange("Show2", _waterQualityNameReconcilePage.CenterPanel);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = _waterQualityNameReconcilePage.RightPanel;
        LogContentChange("Show4", _waterQualityNameReconcilePage.RightPanel);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "", "", "", "", "설정");
        SetLeftPanelWidth(280);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);

        RestoreModeLayout("WaterQualityNameReconcile");
    }

    private void WasteSampleList_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteSampleList";

        if (_wasteSampleListPage == null)
        {
            _wasteSampleListPage = new WasteSampleListPage();
            _wasteSampleListPage.DetailPanelChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _wasteSampleListPage.DetailPanelSilentChanged += panel =>
            {
                var saved = Show2.PageTransition;
                Show2.PageTransition = null;
                Show2.Content = panel;
                Show2.PageTransition = saved;
            };
        }

        Show1.Content = _wasteSampleListPage;
        LogContentChange("Show1", _wasteSampleListPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show4.Content = _wasteSampleListPage.CompanyTreePanel;
        LogContentChange("Show4", _wasteSampleListPage.CompanyTreePanel as Control);
        Show3.Content = null;
        LogContentChange("Show3", null);
        _wasteSampleListPage.LoadData();
        _wasteSampleListPage.LoadCompanyTree();
        _bt1SaveAction = null; // WasteSampleList 저장은 BT1_Click에서 직접 처리

        SetSubMenu("저장", "새로고침", "", "", "", "", "");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 8, lowerStar: 2);

        RestoreModeLayout("WasteSampleList", minLowerStar: 2);
    }

    private void ProcessingFacility_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "ProcessingFacility";

        if (_processingFacilityPage == null)
        {
            _processingFacilityPage = new ProcessingFacilityPage();
            _processingFacilityPage.ResultGridChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
        }

        Show1.Content = _processingFacilityPage;
        LogContentChange("Show1", _processingFacilityPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _bt1SaveAction = () => _processingFacilityPage.Save();

        SetSubMenu("저장", "새로고침", "", "", "", "", "");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 0, lowerStar: 1);

        RestoreModeLayout("ProcessingFacility");
    }

    private void WasteAnalysisInput_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteAnalysisInput";

        if (_wasteAnalysisInputPage == null)
        {
            _wasteAnalysisInputPage = new WasteAnalysisInputPage();
            _wasteAnalysisInputPage.ListPanelChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _wasteAnalysisInputPage.EditPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
            _wasteAnalysisInputPage.StatsPanelChanged += panel =>
            {
                // StatsPanelChanged → Show1 (미매칭 수동매칭 UI 인라인 표시)
                Show1.Content = panel;
                LogContentChange("Show1", panel);
            };
        }

        // 메뉴 Tag로 입력 모드 설정 (수질분석센터/비용부담금/처리시설)
        string? inputMode = (sender as MenuItem)?.Tag?.ToString();
        if (!string.IsNullOrEmpty(inputMode))
            _wasteAnalysisInputPage.SetInputMode(inputMode);

        // WasteAnalysisInputPage: Show4(하단 우측)에 배치, Show1은 수동매칭 인라인용으로 비워둠
        Show4.Content = _wasteAnalysisInputPage;
        LogContentChange("Show4", _wasteAnalysisInputPage);
        Show1.Content = null;
        Show2.Content = null; Show3.Content = null;
        string attachLabel = inputMode == "비용부담금/처리시설" ? "파일첨부" : "";
        SetSubMenu("새로고침", "검증", "입력", "출력", attachLabel, "AI파서", "");
        SetLeftPanelWidth(320);

        // 레이아웃: Show2 전체폭 상단, Show3+Show4 하단 병렬 (Show4 = WasteAnalysisInputPage)
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 6, lowerStar: 4);
        RestoreModeLayout("WasteAnalysisInput", minLowerStar: 2);
        ApplyBottomSplitLayout();

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _wasteAnalysisInputPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void Ecotoxicity_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Ecotoxicity";

        if (_ecotoxicityPage == null)
        {
            _ecotoxicityPage = new EcotoxicityPage();
            _ecotoxicityPage.StatsPanelChanged += panel =>
            {
                Show1.Content = panel;
                LogContentChange("Show1", panel);
            };
            _ecotoxicityPage.ListPanelChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _ecotoxicityPage.EditPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
        }

        Show1.Content = null;
        Show2.Content = null; Show3.Content = null; Show4.Content = null;
        SetSubMenu("새로고침", "시험기록부 출력", "", "", "", "", "");
        SetLeftPanelWidth(250);
        SetContentLayout(content2Star: 6, content4Star: 4, upperStar: 9, lowerStar: 1);
        RestoreModeLayout("Ecotoxicity");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _ecotoxicityPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void ResultSubmitMeasure_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "ResultSubmitMeasure";

        if (_resultSubmitMeasureTestReport == null)
        {
            _resultSubmitMeasureTestReport = new TestReportPage { IsMeasurerMode = true };
            _resultSubmitMeasureTestReport.ResultListChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _resultSubmitMeasureTestReport.EditPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
        }

        Show1.Content = _resultSubmitMeasureTestReport;
        LogContentChange("Show1", _resultSubmitMeasureTestReport);
        Show2.Content = null; Show3.Content = null; Show4.Content = null;
        SetSubMenu("새로고침", "", "", "", "", "", "측정인 LOGIN", "자료TO측정인");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 8.5, lowerStar: 1.5);
        RestoreModeLayout("ResultSubmitMeasure");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _resultSubmitMeasureTestReport.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void ResultSubmitErp_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "ResultSubmitErp";
        if (_resultSubmitErpPage == null)
        {
            _resultSubmitErpPage = new ResultSubmitErpPage();
            _resultSubmitErpPage.Show2ContentRequest = ctrl => Show2.Content = ctrl;
            _resultSubmitErpPage.Show3ContentRequest = ctrl => Show3.Content = ctrl;
        }
        Show1.Content = _resultSubmitErpPage;
        LogContentChange("Show1", _resultSubmitErpPage);
        Show3.Content = null; Show4.Content = null;
        _resultSubmitErpPage.RefreshShow2();         // Excel 테이블 → Show2
        SetSubMenu("저장", "새로고침", "", "", "", "", "");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 5, content4Star: 0, upperStar: 8, lowerStar: 2);
        RestoreModeLayout("ResultSubmitErp");
    }

    private void ResultSubmitZero4_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "ResultSubmitZero4";
        _resultSubmitZero4Page ??= new ResultSubmitZero4Page();
        Show1.Content = _resultSubmitZero4Page;
        LogContentChange("Show1", _resultSubmitZero4Page);
        Show2.Content = null; Show3.Content = null; Show4.Content = null;
        SetSubMenu("저장", "새로고침", "", "", "", "", "");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 8, lowerStar: 2);
        RestoreModeLayout("ResultSubmitZero4");
    }

    // 처리시설 목록 (분석계획 Show1) — 엑셀에서 동적 로딩
    private string[] _facilityNames = Array.Empty<string>();
    private string _selectedFacilityPlan = "중흥처리시설";

    private void AnalysisPlan_Click(object sender, RoutedEventArgs e)
    {
        _currentMode = "AnalysisPlan";
        _analysisPlanSelectedDay = -1;

        // DB에서 시설/시료/항목 로딩
        LoadAnalysisPlanFromDb();

        Show1.Content = BuildFacilityListPanel();
        LogContentChange("Show1", Show1.Content as Control);

        // 첫 번째 시설 기본 선택
        if (string.IsNullOrEmpty(_selectedFacilityPlan) && _facilityNames.Length > 0)
            _selectedFacilityPlan = _facilityNames[0];

        var checkPanel = BuildAnalysisPlanPanel();
        Show2.Content = checkPanel;
        LogContentChange("Show2", checkPanel);

        if (!string.IsNullOrEmpty(_selectedFacilityPlan))
            Show3.Content = BuildFacilityEditPanel(_selectedFacilityPlan);
        else
            Show3.Content = null;
        LogContentChange("Show3", Show3.Content as Control);
        Show4.Content = BuildAnalysisInfoMappingPanel();
        LogContentChange("Show4", Show4.Content as Control);

        _bt1SaveAction = SaveAnalysisPlanWithProgress;
        SetSubMenu("저장", "월", "화", "수", "목", "금", "토", "일", "BASE");
        SetLeftPanelWidth(200);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 8, lowerStar: 2);
        RestoreModeLayout("AnalysisPlan");
    }

    private Dictionary<string, (string 약칭, int 순서)> _facilitySettings = new();

    private Control BuildFacilityListPanel()
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        _facilitySettings = FacilityResultService.GetFacilitySettings();
        var root = new StackPanel { Spacing = 0 };

        // ── 헤더 ────────────────────────────────────────────────────────
        root.Children.Add(new Border
        {
            Background = AppTheme.BgPrimary, BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1), Padding = new Thickness(10, 8),
            Child = new TextBlock
            {
                Text = "처리시설", FontFamily = font, FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgMuted,
            }.BindMD(),
        });

        // ── 시설 리스트 (드래그 순서 변경) ────────────────────────────────
        var listStack = new StackPanel { Spacing = 0 };
        DragDrop.SetAllowDrop(listStack, true);
        listStack.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
        {
            e.DragEffects = e.Data.Contains("fac-reorder") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        });
        listStack.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
        {
            if (!e.Data.Contains("fac-reorder")) return;
            var draggedName = e.Data.Get("fac-reorder") as string;
            if (draggedName == null) return;

            // 드롭 위치 찾기
            var dropPos = e.GetPosition(listStack);
            int dropIdx = listStack.Children.Count;
            double y = 0;
            for (int i = 0; i < listStack.Children.Count; i++)
            {
                var child = listStack.Children[i];
                if (dropPos.Y < y + child.Bounds.Height / 2) { dropIdx = i; break; }
                y += child.Bounds.Height;
            }

            var list = _facilityNames.ToList();
            int oldIdx = list.IndexOf(draggedName);
            if (oldIdx < 0 || oldIdx == dropIdx) return;
            list.RemoveAt(oldIdx);
            if (dropIdx > oldIdx) dropIdx--;
            list.Insert(Math.Min(dropIdx, list.Count), draggedName);
            _facilityNames = list.ToArray();

            FacilityResultService.SaveFacilityOrder(_facilityNames);
            Show1.Content = BuildFacilityListPanel();
            LogContentChange("Show1", Show1.Content as Control);
        });

        foreach (var name in _facilityNames)
        {
            var facilityName = name;
            var isSelected = facilityName == _selectedFacilityPlan;
            _facilitySettings.TryGetValue(facilityName, out var setting);

            // 약칭 뱃지 + 시설명
            var namePanel = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 6 };
            if (!string.IsNullOrEmpty(setting.약칭))
            {
                var (bg, fg) = BadgeColorHelper.GetBadgeColor(setting.약칭);
                namePanel.Children.Add(new Border
                {
                    Background = Brush.Parse(bg), CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(5, 2),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = setting.약칭, FontFamily = font, FontSize = 10,
                        Foreground = Brush.Parse(fg),
                    }
                });
            }
            namePanel.Children.Add(new TextBlock
            {
                Text = facilityName, FontFamily = font,
                Foreground = isSelected ? AppTheme.FgPrimary : AppTheme.FgMuted,
                FontWeight = isSelected ? FontWeight.SemiBold : FontWeight.Normal,
                VerticalAlignment = VerticalAlignment.Center,
            }.BindMD());

            // 드래그 핸들(☰) + 약칭뱃지 + 시설명
            var handleTb = new TextBlock
            {
                Text = "\u2630", FontSize = 12,
                Foreground = AppTheme.FgDimmed,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0),
                Cursor = new Cursor(StandardCursorType.SizeAll),
            };

            var rowPanel = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal };
            rowPanel.Children.Add(handleTb);
            foreach (var child in namePanel.Children.ToList())
            {
                namePanel.Children.Remove(child);
                rowPanel.Children.Add(child);
            }

            var item = new Border
            {
                Background = isSelected ? AppTheme.BgCard : AppTheme.BgSecondary,
                BorderBrush = isSelected ? AppTheme.BorderAccent : Brushes.Transparent,
                BorderThickness = new Thickness(3, 0, 0, 0),
                Padding = new Thickness(8, 10),
                Cursor = new Cursor(StandardCursorType.Hand),
                Tag = facilityName,
                Child = rowPanel,
            };

            // 우클릭 → 컨텍스트 메뉴 (Flyout 사용)
            var deleteMenuItem = new MenuItem { Header = "🗑️ 삭제" };
            deleteMenuItem.Click += (_, _) =>
            {
                try
                {
                    FacilityResultService.DeleteFacility(facilityName);
                    _facilityNames = _facilityNames.Where(f => f != facilityName).ToArray();
                    _selectedFacilityPlan = _facilityNames.Length > 0 ? _facilityNames[0] : "";
                    Show1.Content = BuildFacilityListPanel();
                    LogContentChange("Show1", Show1.Content as Control);
                    Show2.Content = BuildAnalysisPlanPanel();
                    LogContentChange("Show2", Show2.Content as Control);
                    Debug.WriteLine($"✅ 처리시설 '{facilityName}' 삭제 완료");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"❌ 처리시설 삭제 실패: {ex.Message}");
                }
            };
            var contextFlyout = new MenuFlyout { Items = { deleteMenuItem } };
            FlyoutBase.SetAttachedFlyout(item, contextFlyout);

            item.PointerPressed += (_, e) =>
            {
                if (_facilityDragStarted) return;

                // 우클릭 (오른쪽 마우스 버튼)
                if (e.GetCurrentPoint(item).Properties.IsRightButtonPressed)
                {
                    e.Handled = true;
                    FlyoutBase.ShowAttachedFlyout(item);
                    return;
                }

                // 좌클릭 (왼쪽 마우스 버튼)
                _selectedFacilityPlan = facilityName;
                Show1.Content = BuildFacilityListPanel();
                LogContentChange("Show1", Show1.Content as Control);
                Show2.Content = BuildAnalysisPlanPanel();
                LogContentChange("Show2", Show2.Content as Control);
                Show3.Content = BuildFacilityEditPanel(facilityName);
                LogContentChange("Show3", Show3.Content as Control);
            };

            // 드래그 핸들러
            AttachFacilityDrag(item, facilityName);
            listStack.Children.Add(item);
        }
        root.Children.Add(listStack);

        // ── 시설 추가 버튼 ────────────────────────────────────────────────
        var addBtn = new Button
        {
            Content = "+ 시설 추가", FontFamily = font,
            Foreground = AppTheme.FgLink, Background = AppTheme.BgSecondary,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            Padding = new Thickness(12, 10), Margin = new Thickness(0, 4, 0, 0),
            Cursor = new Cursor(StandardCursorType.Hand),
        };
        addBtn.Click += async (_, _) =>
        {
            var dlg = new Window
            {
                Title = "시설 추가", Width = 350, Height = 150,
                WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = false,
            };
            var tb = new TextBox { Watermark = "시설명 입력", Margin = new Thickness(20, 20, 20, 10) };
            var okBtn = new Button { Content = "추가", HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 20) };
            okBtn.Click += (_, _) => dlg.Close(tb.Text?.Trim());
            dlg.Content = new StackPanel { Children = { tb, okBtn } };
            var result = await dlg.ShowDialog<string?>(this);
            if (!string.IsNullOrWhiteSpace(result))
            {
                try
                {
                    FacilityResultService.AddFacility(result);
                    LoadAnalysisPlanFromDb(_analysisPlanSelectedDay == -2 ? -1 : _analysisPlanSelectedDay);
                    _selectedFacilityPlan = result;
                    Show1.Content = BuildFacilityListPanel();
                    Show2.Content = BuildAnalysisPlanPanel();
                }
                catch (Exception ex) { }
            }
        };
        root.Children.Add(addBtn);

        return new ScrollViewer
        {
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Content = root,
        };
    }

    // ── 시설 드래그 순서 변경 ─────────────────────────────────────────────
    private bool _facilityDragStarted;

    private void AttachFacilityDrag(Border item, string facilityName)
    {
        Point? pressPos = null;
        item.PointerPressed += (_, e) =>
        {
            if (!e.GetCurrentPoint(item).Properties.IsLeftButtonPressed) return;
            pressPos = e.GetCurrentPoint(item).Position;
            _facilityDragStarted = false;
        };
        item.PointerMoved += async (_, e) =>
        {
            if (pressPos == null || _facilityDragStarted) return;
            if (!e.GetCurrentPoint(item).Properties.IsLeftButtonPressed) { pressPos = null; return; }
            var diff = e.GetCurrentPoint(item).Position - pressPos.Value;
            if (Math.Abs(diff.Y) > 8)
            {
                _facilityDragStarted = true;
                pressPos = null;
                var data = new DataObject();
                data.Set("fac-reorder", facilityName);
                await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                _facilityDragStarted = false;
            }
        };
        item.PointerReleased += (_, _) => { pressPos = null; _facilityDragStarted = false; };
    }

    // ── Show3: 시설 편집 패널 ─────────────────────────────────────────────
    private Control BuildFacilityEditPanel(string facilityName)
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        _facilitySettings.TryGetValue(facilityName, out var setting);

        var stack = new StackPanel { Spacing = 12, Margin = new Thickness(16) };

        stack.Children.Add(new TextBlock
        {
            Text = "시설 설정", FontFamily = font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary, FontSize = 15,
        });

        // 시설명
        stack.Children.Add(new TextBlock { Text = "시설명", FontFamily = font, Foreground = AppTheme.FgMuted, FontSize = 12 });
        var nameBox = new TextBox { Text = facilityName, FontFamily = font, Margin = new Thickness(0, 2, 0, 0) };
        stack.Children.Add(nameBox);

        // 약칭
        stack.Children.Add(new TextBlock { Text = "약칭 (뱃지)", FontFamily = font, Foreground = AppTheme.FgMuted, FontSize = 12 });
        var aliasBox = new TextBox { Text = setting.약칭, FontFamily = font, Watermark = "예: 중흥", Margin = new Thickness(0, 2, 0, 0) };
        stack.Children.Add(aliasBox);

        // 적용 버튼
        var applyBtn = new Button
        {
            Content = "적용", FontFamily = font, Padding = new Thickness(20, 8),
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 8, 0, 0),
        };
        applyBtn.Click += (_, _) =>
        {
            try
            {
                var newName = nameBox.Text?.Trim() ?? facilityName;
                var newAlias = aliasBox.Text?.Trim() ?? "";

                // 약칭 저장
                FacilityResultService.SaveFacilityAlias(facilityName, newAlias);

                // 시설명 변경
                if (newName != facilityName && !string.IsNullOrEmpty(newName))
                {
                    FacilityResultService.RenameFacility(facilityName, newName);
                    _selectedFacilityPlan = newName;
                }

                // UI 갱신
                LoadAnalysisPlanFromDb(_analysisPlanSelectedDay == -2 ? -1 : _analysisPlanSelectedDay);
                Show1.Content = BuildFacilityListPanel();
                Show2.Content = BuildAnalysisPlanPanel();
                Show3.Content = BuildFacilityEditPanel(_selectedFacilityPlan);
            }
            catch (Exception ex) { }
        };
        stack.Children.Add(applyBtn);

        return new Border
        {
            Background = AppTheme.BgPrimary,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = stack,
        };
    }

    // ── 저장 (프로그레스바 포함) ─────────────────────────────────────────
    // ── Show4: 분석항목 편집 패널 ─────────────────────────────────────────
    // ── DB 분석정보에서 분석항목 조회 ────────────────────────────────────────
    private List<string> GetDbAnalysisItems()
    {
        var items = new List<string>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `Analyte` FROM `분석정보` WHERE `Analyte` IS NOT NULL AND `Analyte` <> '' ORDER BY `Category`, `_id`";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var analyte = reader[0]?.ToString();
                if (!string.IsNullOrWhiteSpace(analyte))
                    items.Add(analyte);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[분석정보 조회 오류] {ex.Message}");
        }
        return items;
    }

    // ── Show4: DB 분석정보 매칭 패널 ───────────────────────────────────────
    private Control BuildDbAnalysisInfoPanel()
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var dbItems = GetDbAnalysisItems();

        var stack = new StackPanel { Spacing = 8, Margin = new Thickness(16) };
        stack.Children.Add(new TextBlock
        {
            Text = "분석정보 관리", FontFamily = font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary, FontSize = 15,
        });

        // ── DB 분석항목 리스트 (클릭하면 매칭) ────────────────────────────
        var listStack = new StackPanel { Spacing = 2 };
        var listScroll = new ScrollViewer
        {
            Content = listStack, MaxHeight = 400,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };

        foreach (var analyte in dbItems)
        {
            var border = new Border
            {
                Background = (!string.IsNullOrEmpty(_selectedAnalysisInfoForMapping) && analyte == _selectedAnalysisInfoForMapping)
                    ? AppTheme.BgActiveBlue
                    : AppTheme.BgCard,
                BorderBrush = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(8, 6),
                MinHeight = 32,
                Cursor = new Cursor(StandardCursorType.Hand),
                Tag = analyte,
            };

            var tb = new TextBlock
            {
                Text = analyte, FontFamily = font,
                Foreground = string.IsNullOrEmpty(_selectedAnalysisInfoForMapping)
                    ? AppTheme.FgPrimary
                    : (analyte == _selectedAnalysisInfoForMapping ? AppTheme.FgSuccess : AppTheme.FgPrimary),
                VerticalAlignment = VerticalAlignment.Center,
            };
            border.Child = tb;

            var capturedAnalyte = analyte;
            var capturedBorder = border;
            border.PointerPressed += (_, _) =>
            {
                if (_pendingMappingItemIndex >= 0)
                {
                    SaveAliasByIndex(_pendingMappingItemIndex, capturedAnalyte);
                    _selectedAnalysisInfoForMapping = null;
                    _pendingMappingItemIndex = -1;
                    Show4.Content = null;
                    return;
                }
                _selectedAnalysisInfoForMapping = capturedAnalyte;
                Show4.Content = BuildDbAnalysisInfoPanel();
            };

            listStack.Children.Add(border);
        }
        stack.Children.Add(listScroll);

        // ── 상태 표시 ────────────────────────────────────────────────────
        var statusRow = new Border
        {
            Background = AppTheme.BgInput,
            BorderBrush = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1),
            Padding = new Thickness(8, 6),
            Margin = new Thickness(0, 8, 0, 0),
        };

        var statusTb = new TextBlock
        {
            Text = string.IsNullOrEmpty(_selectedAnalysisInfoForMapping)
                ? "분석항목을 선택하세요"
                : $"선택됨: {_selectedAnalysisInfoForMapping}",
            FontFamily = font,
            Foreground = AppTheme.FgSecondary,
            VerticalAlignment = VerticalAlignment.Center,
        };
        statusRow.Child = statusTb;
        stack.Children.Add(statusRow);

        return stack;
    }

    // ── 기존 분석항목 편집 패널 (비활성) ──────────────────────────────────────
    private Control BuildAnalysisItemEditorPanel()
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var allItems = FacilityResultService.GetAnalysisItems(activeOnly: false);

        var stack = new StackPanel { Spacing = 8, Margin = new Thickness(16) };
        stack.Children.Add(new TextBlock
        {
            Text = "분석항목 관리 (비활성)", FontFamily = font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgDimmed, FontSize = 15,
        });

        // ── 항목 리스트 (체크박스 + 이름 + 순서 버튼) ──────────────────
        var listStack = new StackPanel { Spacing = 2 };
        foreach (var item in allItems.OrderBy(i => i.순서))
        {
            var row = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto,Auto"),
                MinHeight = 32,
            };

            // 활성 체크박스
            var cb = new CheckBox
            {
                IsChecked = item.활성,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0),
            };
            var capturedId = item.Id;
            cb.IsCheckedChanged += (_, _) =>
            {
                FacilityResultService.SetItemActive(capturedId, cb.IsChecked == true);
            };
            Grid.SetColumn(cb, 0);
            row.Children.Add(cb);

            // 항목명
            var nameTb = new TextBlock
            {
                Text = item.항목명, FontFamily = font,
                Foreground = item.활성 ? AppTheme.FgPrimary : AppTheme.FgDimmed,
                VerticalAlignment = VerticalAlignment.Center,
            };
            Grid.SetColumn(nameTb, 1);
            row.Children.Add(nameTb);

            // 위로 버튼
            var upBtn = new Button
            {
                Content = "\u25b2", Padding = new Thickness(6, 2), FontSize = 10,
                Background = Brushes.Transparent, Foreground = AppTheme.FgMuted,
            };
            upBtn.Click += (_, _) =>
            {
                var ordered = FacilityResultService.GetAnalysisItems(false).OrderBy(i => i.순서).ToList();
                var idx = ordered.FindIndex(i => i.Id == capturedId);
                if (idx > 0)
                {
                    (ordered[idx], ordered[idx - 1]) = (ordered[idx - 1], ordered[idx]);
                    FacilityResultService.SaveItemOrder(ordered.Select(i => i.Id).ToList());
                    Show4.Content = BuildAnalysisItemEditorPanel();
                }
            };
            Grid.SetColumn(upBtn, 2);
            row.Children.Add(upBtn);

            // 아래로 버튼
            var downBtn = new Button
            {
                Content = "\u25bc", Padding = new Thickness(6, 2), FontSize = 10,
                Background = Brushes.Transparent, Foreground = AppTheme.FgMuted,
            };
            downBtn.Click += (_, _) =>
            {
                var ordered = FacilityResultService.GetAnalysisItems(false).OrderBy(i => i.순서).ToList();
                var idx = ordered.FindIndex(i => i.Id == capturedId);
                if (idx >= 0 && idx < ordered.Count - 1)
                {
                    (ordered[idx], ordered[idx + 1]) = (ordered[idx + 1], ordered[idx]);
                    FacilityResultService.SaveItemOrder(ordered.Select(i => i.Id).ToList());
                    Show4.Content = BuildAnalysisItemEditorPanel();
                }
            };
            Grid.SetColumn(downBtn, 3);
            row.Children.Add(downBtn);

            listStack.Children.Add(row);
        }
        stack.Children.Add(new ScrollViewer
        {
            Content = listStack, MaxHeight = 350,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        });

        // ── 새 항목 추가 ────────────────────────────────────────────────
        var addRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(0, 8, 0, 0) };
        var addBox = new TextBox { Watermark = "새 항목명", FontFamily = font };
        Grid.SetColumn(addBox, 0);
        addRow.Children.Add(addBox);

        var addBtn = new Button
        {
            Content = "추가", FontFamily = font, Padding = new Thickness(12, 6),
            Margin = new Thickness(6, 0, 0, 0),
        };
        addBtn.Click += (_, _) =>
        {
            var name = addBox.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(name))
            {
                try
                {
                    FacilityResultService.AddAnalysisItem(name);
                    Show4.Content = BuildAnalysisItemEditorPanel();
                }
                catch (Exception ex) { }
            }
        };
        Grid.SetColumn(addBtn, 1);
        addRow.Children.Add(addBtn);
        stack.Children.Add(addRow);

        // ── 적용 버튼 (헤더 갱신) ────────────────────────────────────────
        var applyBtn = new Button
        {
            Content = "적용 (헤더 갱신)", FontFamily = font, Padding = new Thickness(20, 8),
            Margin = new Thickness(0, 12, 0, 0),
        };
        applyBtn.Click += (_, _) =>
        {
            FacilityResultService.InvalidateItemsCache();
            _analysisPlanItems = FacilityResultService.AnalysisPlanItemNames;
            LoadAnalysisPlanFromDb(_analysisPlanSelectedDay == -2 ? -1 : _analysisPlanSelectedDay);
            Show2.Content = BuildAnalysisPlanPanel();
            LogContentChange("Show2", Show2.Content as Control);
        };
        stack.Children.Add(applyBtn);

        return new Border
        {
            Background = AppTheme.BgPrimary, BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = stack,
        };
    }

    private async void SaveAnalysisPlanWithProgress()
    {
        var facility = _selectedFacilityPlan;
        if (string.IsNullOrEmpty(facility)) return;
        if (!_facilitySamples.TryGetValue(facility, out var samples)) return;
        if (!_facilityPlanState.TryGetValue(facility, out var checkRows)) return;

        // 프로그레스바 표시 (0→100%)
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var pb = new Avalonia.Controls.ProgressBar
        {
            Minimum = 0, Maximum = 100, Value = 0,
            Height = 6, Foreground = AppTheme.FgInfo,
        };
        var pctText = new TextBlock
        {
            Text = "0%", FontFamily = font, Foreground = AppTheme.FgMuted,
            HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 8),
        };
        var savePanel = new StackPanel
        {
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            Margin = new Thickness(40),
            Children = { pb, pctText },
        };
        Show2.Content = savePanel;

        try
        {
            var total = samples.Length;
            var day = _analysisPlanSelectedDay;
            var isBase = day == -2;

            await Task.Run(() =>
            {
                var cols = FacilityResultService.PlanDbCols;
                using var conn = DbConnectionFactory.CreateConnection();
                conn.Open();

                for (int si = 0; si < total; si++)
                {
                    if (si >= checkRows.Count) continue;
                    var checks = checkRows[si];
                    int dayStart = isBase ? 0 : day;
                    int dayEnd   = isBase ? 6 : day;

                    for (int d = dayStart; d <= dayEnd; d++)
                    {
                        using var cmd = conn.CreateCommand();
                        var sets = new List<string>();
                        for (int i = 0; i < cols.Length && i < checks.Length; i++)
                        {
                            sets.Add($"{cols[i]} = @v{i}");
                            cmd.Parameters.AddWithValue($"@v{i}", checks[i] ? "O" : "");
                        }
                        cmd.CommandText = $"UPDATE `처리시설_분석계획` SET {string.Join(", ", sets)} WHERE 시설명=@f AND 시료명=@s AND 요일=@d";
                        cmd.Parameters.AddWithValue("@f", facility);
                        cmd.Parameters.AddWithValue("@s", samples[si]);
                        cmd.Parameters.AddWithValue("@d", d);
                        cmd.ExecuteNonQuery();
                    }

                    int pct = (int)((si + 1) * 100.0 / total);
                    Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                    {
                        pb.Value = pct;
                        pctText.Text = $"{pct}%";
                    });
                }
            });
            pb.Value = 100;
            pctText.Text = "100% 완료";
            await Task.Delay(300);
        }
        catch (Exception ex)
        {
        }
        finally
        {
            Show2.Content = BuildAnalysisPlanPanel();
            LogContentChange("Show2", Show2.Content as Control);
        }
    }

    private Control BuildAnalysisPlanPanel()
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var facility = _selectedFacilityPlan;

        if (!_facilitySamples.TryGetValue(facility, out var samples))
            samples = Array.Empty<string>();

        // 상태 초기화 (처음 진입 시)
        if (!_facilityPlanState.TryGetValue(facility, out var stateRows))
        {
            stateRows = samples.Select(_ => new bool[_analysisPlanItems.Length]).ToList();
            _facilityPlanState[facility] = stateRows;
        }

        // 컬럼 정의: 시료명(160) + 항목 12개(각 58)
        const int firstColW = 160;
        const int itemColW  = 58;
        string colDef = _analysisPlanItems.Length > 0
            ? firstColW + "," + string.Join(",", Enumerable.Repeat(itemColW, _analysisPlanItems.Length))
            : firstColW.ToString();

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        // ── 고정 헤더 ──────────────────────────────────────────────────────
        var headerGrid = new Grid { ColumnDefinitions = new ColumnDefinitions(colDef) };
        headerGrid.Children.Add(new TextBlock
        {
            Text = "시료명", FontFamily = font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 4),
        }.BindSM());
        for (int i = 0; i < _analysisPlanItems.Length; i++)
        {
            var itemName = _analysisPlanItems[i];
            var hdrBorder = new Border
            {
                Cursor = new Cursor(StandardCursorType.Hand),
                Background = Brushes.Transparent,
                Child = new TextBlock
                {
                    Text = itemName, FontFamily = font, FontWeight = FontWeight.SemiBold,
                    Foreground = AppTheme.FgMuted,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(2, 4),
                }.BindXS(),
            };
            DragDrop.SetAllowDrop(hdrBorder, true);
            hdrBorder.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
            {
                e.DragEffects = e.Data?.Contains("analysisinfo-item") == true ? DragDropEffects.Move : DragDropEffects.None;
                e.Handled = true;
            });
            hdrBorder.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
            {
                if (e.Data?.Contains("analysisinfo-item") != true) return;
                var analysisInfo = e.Data.Get("analysisinfo-item") as string;
                if (!string.IsNullOrWhiteSpace(analysisInfo))
                    MapAnalysisPlanItemToAnalysisInfo(i, analysisInfo);
                _selectedAnalysisInfoForMapping = null;
                e.Handled = true;
            });
            hdrBorder.PointerPressed += (_, _) =>
            {
                if (!string.IsNullOrEmpty(_selectedAnalysisInfoForMapping))
                {
                    SaveAliasByIndex(_pendingMappingItemIndex, _selectedAnalysisInfoForMapping);
                    _selectedAnalysisInfoForMapping = null;
                    _pendingMappingItemIndex = -1;
                    Show4.Content = null;
                    return;
                }

                _pendingMappingItemIndex = i;
                _pendingMappingItemName = itemName;
                Show4.Content = BuildDbAnalysisInfoPanel();
                LogContentChange("Show4", Show4.Content as Control);
                SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 6, lowerStar: 4);
            };
            Grid.SetColumn(hdrBorder, i + 1);
            headerGrid.Children.Add(hdrBorder);
        }
        var headerBorder = new Border
        {
            Background = AppTheme.BgCard,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = headerGrid,
        };
        Grid.SetRow(headerBorder, 0);
        root.Children.Add(headerBorder);

        // ── 모드 표시 ────────────────────────────────────────────────────
        bool isBaseMode = _analysisPlanSelectedDay == -2;

        // ── 시료 행 (드래그 순서 변경) ─────────────────────────────────────
        var bodyStack = new StackPanel { Spacing = 1 };
        DragDrop.SetAllowDrop(bodyStack, true);
        bodyStack.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
        {
            e.DragEffects = e.Data.Contains("sample-reorder") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        });
        bodyStack.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
        {
            if (!e.Data.Contains("sample-reorder")) return;
            var draggedSample = e.Data.Get("sample-reorder") as string;
            if (draggedSample == null || !_facilitySamples.TryGetValue(facility, out var curSamples)) return;

            var dropPos = e.GetPosition(bodyStack);
            int dropIdx = curSamples.Length;
            double y = 0;
            for (int i = 0; i < bodyStack.Children.Count; i++)
            {
                var child = bodyStack.Children[i];
                if (child.Tag is not string) continue; // 시료 행만
                if (dropPos.Y < y + child.Bounds.Height / 2) { dropIdx = i; break; }
                y += child.Bounds.Height;
            }

            var list = curSamples.ToList();
            int oldIdx = list.IndexOf(draggedSample);
            if (oldIdx < 0 || oldIdx == dropIdx) return;
            list.RemoveAt(oldIdx);
            if (dropIdx > oldIdx) dropIdx--;
            list.Insert(Math.Min(dropIdx, list.Count), draggedSample);
            _facilitySamples[facility] = list.ToArray();

            FacilityResultService.SaveSampleOrder(facility, list.ToArray());
            // 상태도 같은 순서로 재배치
            if (_facilityPlanState.TryGetValue(facility, out var stRows))
            {
                var newState = new List<bool[]>();
                foreach (var sn in list)
                {
                    int origIdx = Array.IndexOf(curSamples, sn);
                    newState.Add(origIdx >= 0 && origIdx < stRows.Count ? stRows[origIdx] : new bool[_analysisPlanItems.Length]);
                }
                _facilityPlanState[facility] = newState;
            }
            Show2.Content = BuildAnalysisPlanPanel();
            LogContentChange("Show2", Show2.Content as Control);
        });

        for (int si = 0; si < samples.Length; si++)
        {
            int sampleIdx = si;
            var sampleName = samples[si];
            var checks     = si < stateRows.Count ? stateRows[si] : new bool[_analysisPlanItems.Length];
            bool hasAnyCheck = checks.Any(c => c);

            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDef),
                MinHeight  = 28,
                Background = hasAnyCheck ? AppTheme.BgSecondary : AppTheme.BgPrimary,
                Opacity    = hasAnyCheck ? 1.0 : 0.5,
                Tag        = sampleName, // 드래그 식별용
            };

            // 우클릭 → 컨텍스트 메뉴 (Flyout 사용)
            var deleteSampleMenuItem = new MenuItem { Header = "🗑️ 삭제" };
            deleteSampleMenuItem.Click += (_, _) =>
            {
                try
                {
                    FacilityResultService.DeleteSample(facility, sampleName);
                    if (_facilitySamples.TryGetValue(facility, out var curSamples))
                    {
                        _facilitySamples[facility] = curSamples.Where(s => s != sampleName).ToArray();
                    }
                    if (_facilityPlanState.TryGetValue(facility, out var stRows))
                    {
                        var newState = stRows.Where((_, i) => i != sampleIdx).ToList();
                        _facilityPlanState[facility] = newState;
                    }
                    Show2.Content = BuildAnalysisPlanPanel();
                    LogContentChange("Show2", Show2.Content as Control);
                    Debug.WriteLine($"✅ 시료 '{sampleName}' 삭제 완료");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"❌ 시료 삭제 실패: {ex.Message}");
                }
            };
            var sampleContextFlyout = new MenuFlyout { Items = { deleteSampleMenuItem } };
            FlyoutBase.SetAttachedFlyout(rowGrid, sampleContextFlyout);

            rowGrid.PointerPressed += (_, e) =>
            {
                // 우클릭 (오른쪽 마우스 버튼)
                if (e.GetCurrentPoint(rowGrid).Properties.IsRightButtonPressed)
                {
                    e.Handled = true;
                    FlyoutBase.ShowAttachedFlyout(rowGrid);
                }
            };

            bool isMainView = _analysisPlanSelectedDay == -1; // 전체 뷰

            if (isMainView || isBaseMode)
            {
                // 전체/BASE: 시료명 편집 + ☰ 드래그 핸들(전체만)
                var nameRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal };

                if (isMainView)
                {
                    // ☰ 드래그 핸들 (전체 뷰에서만)
                    var capturedName = sampleName;
                    var handle = new TextBlock
                    {
                        Text = "\u2630", FontSize = 10, Foreground = AppTheme.FgDimmed,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(4, 0, 4, 0),
                        Cursor = new Cursor(StandardCursorType.SizeAll),
                    };
                    Point? pressPos = null;
                    bool dragStarted = false;
                    handle.PointerPressed += (_, e) =>
                    {
                        if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed) return;
                        pressPos = e.GetCurrentPoint(handle).Position;
                        dragStarted = false;
                        e.Handled = true;
                    };
                    handle.PointerMoved += async (_, e) =>
                    {
                        if (pressPos == null || dragStarted) return;
                        if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed) { pressPos = null; return; }
                        var diff = e.GetCurrentPoint(handle).Position - pressPos.Value;
                        if (Math.Abs(diff.Y) > 8)
                        {
                            dragStarted = true; pressPos = null;
                            var data = new DataObject();
                            data.Set("sample-reorder", capturedName);
                            await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                            dragStarted = false;
                        }
                    };
                    handle.PointerReleased += (_, _) => { pressPos = null; dragStarted = false; };
                    nameRow.Children.Add(handle);

                    // ✕ 삭제 버튼 (전체 뷰에서만)
                    var deleteBtn = new Button
                    {
                        Content = "✕",
                        FontSize = 14,
                        Foreground = AppTheme.FgMuted,
                        Background = Brushes.Transparent,
                        BorderThickness = new Thickness(0),
                        Padding = new Thickness(4, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                        Cursor = new Cursor(StandardCursorType.Hand),
                    };
                    deleteBtn.Click += (_, _) =>
                    {
                        try
                        {
                            using var conn = DbConnectionFactory.CreateConnection();
                            conn.Open();
                            using var cmd = conn.CreateCommand();
                            cmd.CommandText = $"DELETE FROM `처리시설_분석계획` WHERE 시설명=@f AND 시료명=@s";
                            cmd.Parameters.AddWithValue("@f", facility);
                            cmd.Parameters.AddWithValue("@s", sampleName);
                            cmd.ExecuteNonQuery();
                            // UI 새로고침
                            Show2.Content = BuildAnalysisPlanPanel();
                            LogContentChange("Show2", Show2.Content as Control);
                        }
                        catch (Exception ex) { Debug.WriteLine($"시료명 삭제 실패: {ex.Message}"); }
                    };
                    nameRow.Children.Add(deleteBtn);
                }

                // 시료명 편집 TextBox
                var nameBox = new TextBox
                {
                    Text = sampleName, FontFamily = font,
                    Foreground = AppTheme.FgPrimary,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, 0), MinHeight = 26,
                    BorderThickness = new Thickness(0, 0, 0, 1),
                    BorderBrush = AppTheme.BorderSubtle,
                    Background = Brushes.Transparent,
                };
                var oldName = sampleName;
                nameBox.LostFocus += (_, _) =>
                {
                    var newName = nameBox.Text?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(newName) && newName != oldName)
                    {
                        try
                        {
                            FacilityResultService.RenameSampleAllDays(facility, oldName, newName);
                            var list = _facilitySamples[facility].ToList();
                            list[sampleIdx] = newName;
                            _facilitySamples[facility] = list.ToArray();
                        }
                        catch (Exception ex) { }
                    }
                };
                nameRow.Children.Add(nameBox);
                Grid.SetColumn(nameRow, 0);
                rowGrid.Children.Add(nameRow);
            }
            else
            {
                // 개별 요일: 읽기전용
                var nameTb = new TextBlock
                {
                    Text = sampleName, FontFamily = font,
                    Foreground = AppTheme.FgPrimary,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin = new Thickness(6, 0),
                }.BindSM();
                Grid.SetColumn(nameTb, 0);
                rowGrid.Children.Add(nameTb);
            }

            for (int ii = 0; ii < _analysisPlanItems.Length; ii++)
            {
                int itemIdx = ii;
                var cb = new CheckBox
                {
                    IsChecked = checks[itemIdx],
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                };
                cb.IsCheckedChanged += (_, _) =>
                {
                    checks[itemIdx] = cb.IsChecked == true;
                    if (isBaseMode)
                    {
                        try
                        {
                            FacilityResultService.ApplyBaseToAllDays(facility, new[] { sampleName }, new List<bool[]> { checks });
                            // BASE 모드에서는 즉시 DB 반영 후 UI를 다시 그려서 일괄 변경이 확실히 보이도록 함
                            Show2.Content = BuildAnalysisPlanPanel();
                            LogContentChange("Show2", Show2.Content as Control);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"BASE 모드 즉시 반영 실패: {ex.Message}");
                        }
                    }
                };
                Grid.SetColumn(cb, itemIdx + 1);
                rowGrid.Children.Add(cb);
            }

            bodyStack.Children.Add(rowGrid);
        }

        // ── 시료 추가 버튼 ──────────────────────────────────────────────────
        {
            var addSampleBtn = new Button
            {
                Content = "+ 시료 추가",
                FontFamily = font,
                Foreground = AppTheme.FgLink,
                Background = AppTheme.BgSecondary,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 4, 0, 0),
                Padding = new Thickness(6),
                Cursor = new Cursor(StandardCursorType.Hand),
            };
            addSampleBtn.Click += async (_, _) =>
            {
                // 시설이 선택되어 있지 않으면 시설 선택 다이얼로그
                var targetFacility = facility;
                if (string.IsNullOrEmpty(targetFacility) || _facilityNames.Length == 0)
                {
                    return;
                }
                if (_facilityNames.Length > 1 && string.IsNullOrEmpty(_selectedFacilityPlan))
                {
                    var facDlg = new Window
                    {
                        Title = "시설 선택", Width = 300, Height = 200,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = false,
                    };
                    var facList = new ListBox { ItemsSource = _facilityNames, Margin = new Thickness(10) };
                    var facOk = new Button { Content = "선택", HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 10) };
                    facOk.Click += (_, _) => facDlg.Close(facList.SelectedItem as string);
                    facDlg.Content = new StackPanel { Children = { facList, facOk } };
                    var picked = await facDlg.ShowDialog<string?>(this);
                    if (string.IsNullOrWhiteSpace(picked)) return;
                    targetFacility = picked;
                    _selectedFacilityPlan = picked;
                }

                // 시료명 입력
                var dlg = new Window
                {
                    Title = $"시료 추가 ({targetFacility})", Width = 300, Height = 140,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner, CanResize = false,
                };
                var tb = new TextBox { Watermark = "시료명 입력", Margin = new Thickness(20, 20, 20, 10) };
                var okBtn = new Button { Content = "추가", HorizontalAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 0, 0, 20) };
                okBtn.Click += (_, _) => dlg.Close(tb.Text?.Trim());
                dlg.Content = new StackPanel { Children = { tb, okBtn } };
                var result = await dlg.ShowDialog<string?>(this);
                if (!string.IsNullOrWhiteSpace(result))
                {
                    try
                    {
                        FacilityResultService.AddSampleToFacility(targetFacility, result);
                        LoadAnalysisPlanFromDb(_analysisPlanSelectedDay == -2 ? -1 : _analysisPlanSelectedDay);
                        Show1.Content = BuildFacilityListPanel();
                        Show2.Content = BuildAnalysisPlanPanel();
                    }
                    catch (Exception ex) { }
                }
            };
            bodyStack.Children.Add(addSampleBtn);
        }

        if (samples.Length == 0)
            bodyStack.Children.Add(new TextBlock
            {
                Text = "시설을 선택하세요.", FontFamily = font,
                Foreground = AppTheme.FgWarn, Margin = new Thickness(10),
            }.BindMD());

        var scroll = new ScrollViewer
        {
            Content = bodyStack,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        return root;
    }

    // ── DB에서 분석계획 데이터 로딩 ───────────────────────────────────────

    /// <summary>DB에서 시설/시료/항목/체크 상태 로딩</summary>
    private void LoadAnalysisPlanFromDb(int dayIdx = -1)
    {
        try
        {
            _analysisPlanItems = FacilityResultService.AnalysisPlanItemNames;

            var (facilities, samples) = FacilityResultService.GetAnalysisPlanStructure();
            _facilityNames = facilities;
            _facilitySamples = samples;

            _facilityPlanState.Clear();
            var state = FacilityResultService.GetAnalysisPlanState(dayIdx);
            foreach (var kv in state)
                _facilityPlanState[kv.Key] = kv.Value;

            if (_facilityNames.Length > 0 && string.IsNullOrEmpty(_selectedFacilityPlan))
                _selectedFacilityPlan = _facilityNames[0];

        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[LoadAnalysisPlanFromDb] {ex.Message}");
        }
    }

    private void MapAnalysisPlanItemToAnalysisInfo(int itemIndex, string analysisInfoName)
    {
        var items = FacilityResultService.GetAnalysisItems(activeOnly: true);
        if (itemIndex < 0 || itemIndex >= items.Count) return;

        var targetItem = items[itemIndex];
        if (targetItem.항목명 == analysisInfoName) return;

        FacilityResultService.RenameAnalysisItem(targetItem.Id, analysisInfoName);
        _analysisPlanItems = FacilityResultService.AnalysisPlanItemNames;
        Show2.Content = BuildAnalysisPlanPanel();
        Show4.Content = BuildAnalysisInfoMappingPanel();
        LogContentChange("Show2", Show2.Content as Control);
        LogContentChange("Show4", Show4.Content as Control);
    }

    private void SaveAliasByIndex(int itemIndex, string analyteAlias)
    {
        var targetName = !string.IsNullOrEmpty(_pendingMappingItemName)
            ? _pendingMappingItemName
            : (itemIndex >= 0 && itemIndex < _analysisPlanItems.Length ? _analysisPlanItems[itemIndex] : null);
        if (string.IsNullOrEmpty(targetName)) return;
        var allItems = FacilityResultService.GetAnalysisItems(activeOnly: false);
        var item = allItems.FirstOrDefault(x => x.항목명 == targetName);
        if (item == null) return;
        FacilityResultService.SaveAnalyteAlias(item.Id, analyteAlias);
    }

    private Control BuildAnalysisInfoMappingPanel()
    {
        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var analytes = AnalysisRequestService.GetOrderedAnalytes();

        var stack = new StackPanel { Spacing = 10, Margin = new Thickness(16) };
        stack.Children.Add(new TextBlock
        {
            Text = "분석정보 항목 매핑",
            FontFamily = font,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
            FontSize = 16,
        });
        stack.Children.Add(new TextBlock
        {
            Text = "Show2의 처리시설 분석계획 컬럼 헤더에 분석정보 항목을 드래그하여 매핑하세요."
                 + " 컬럼명은 유지되며 표시명은 분석정보 기준으로 갱신됩니다.",
            FontFamily = font,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
        });

        var searchBox = new TextBox
        {
            Watermark = "분석정보 항목 검색...",
            FontFamily = font,
            Height = 32,
            Margin = new Thickness(0, 6, 0, 0),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            Background = AppTheme.BgPrimary,
            Foreground = AppTheme.FgPrimary,
        };

        var listStack = new StackPanel { Spacing = 2 };
        void RefreshList()
        {
            listStack.Children.Clear();
            var query = searchBox.Text?.Trim().ToLower() ?? "";
            foreach (var (fullName, alias) in analytes)
            {
                if (!string.IsNullOrEmpty(query) && !fullName.ToLower().Contains(query) && !alias.ToLower().Contains(query))
                    continue;

                var label = fullName;
                if (!string.IsNullOrWhiteSpace(alias) && alias != fullName)
                    label += $" ({alias})";

                var row = new Border
                {
                    Background = AppTheme.BgSecondary,
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(10, 8),
                    Cursor = new Cursor(StandardCursorType.Hand),
                };
                var rowText = new TextBlock
                {
                    Text = label,
                    FontFamily = font,
                    Foreground = AppTheme.FgPrimary,
                    TextWrapping = TextWrapping.Wrap,
                };
                row.Child = rowText;
                row.Tag = fullName;

                Point? pressPos = null;
                bool dragStarted = false;
                row.PointerPressed += (_, e) =>
                {
                    if (!e.GetCurrentPoint(row).Properties.IsLeftButtonPressed) return;
                    pressPos = e.GetCurrentPoint(row).Position;
                    _selectedAnalysisInfoForMapping = fullName;
                    RefreshSelectionVisuals();
                    e.Handled = true;
                };

                row.PointerMoved += async (_, e) =>
                {
                    if (pressPos == null || dragStarted) return;
                    if (!e.GetCurrentPoint(row).Properties.IsLeftButtonPressed) { pressPos = null; return; }
                    var diff = e.GetCurrentPoint(row).Position - pressPos.Value;
                    if (Math.Abs(diff.X) > 8 || Math.Abs(diff.Y) > 8)
                    {
                        dragStarted = true;
                        pressPos = null;
                        var data = new DataObject();
                        data.Set("analysisinfo-item", fullName);
                        await DragDrop.DoDragDrop(e, data, DragDropEffects.Move);
                        dragStarted = false;
                    }
                };
                row.PointerReleased += (_, _) => { pressPos = null; dragStarted = false; };
                row.PointerPressed += (_, e) => pressPos = e.GetCurrentPoint(row).Position;

                listStack.Children.Add(row);
            }
        }

        var selectedInfoText = new TextBlock
        {
            Text = "선택된 항목: 없음",
            FontFamily = font,
            Foreground = AppTheme.FgMuted,
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        };

        void UpdateSelectedInfoText()
        {
            selectedInfoText.Text = string.IsNullOrEmpty(_selectedAnalysisInfoForMapping)
                ? "선택된 항목: 없음"
                : $"선택된 항목: {_selectedAnalysisInfoForMapping} (헤더 클릭 시 매핑)";
        }

        void RefreshSelectionVisuals()
        {
            foreach (var child in listStack.Children.OfType<Border>())
            {
                if (child.Tag is string tag && tag == _selectedAnalysisInfoForMapping)
                {
                    child.Background = AppTheme.BgActiveBlue;
                }
                else
                {
                    child.Background = AppTheme.BgCard;
                }
            }
            UpdateSelectedInfoText();
        }

        searchBox.TextChanged += (_, _) => RefreshList();
        RefreshList();

        stack.Children.Add(searchBox);
        stack.Children.Add(selectedInfoText);

        UpdateSelectedInfoText();

        stack.Children.Add(new ScrollViewer
        {
            Content = listStack,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            MaxHeight = 520,
        });

        RefreshSelectionVisuals();

        return new Border
        {
            Background = AppTheme.BgPrimary,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child = stack,
        };
    }

    /// <summary>현재 선택된 시설의 체크 상태를 DB에 저장</summary>
    private void SaveAnalysisPlanToDb()
    {
        var facility = _selectedFacilityPlan;
        if (string.IsNullOrEmpty(facility) || !_facilitySamples.TryGetValue(facility, out var samples))
            return;
        if (!_facilityPlanState.TryGetValue(facility, out var checkRows))
            return;

        try
        {
            if (_analysisPlanSelectedDay == -2)
            {
                // BASE 모드: 전체 요일에 일괄 적용
                FacilityResultService.ApplyBaseToAllDays(facility, samples, checkRows);
            }
            else if (_analysisPlanSelectedDay >= 0)
            {
                FacilityResultService.SaveAnalysisPlanState(facility, samples, checkRows, _analysisPlanSelectedDay);
            }
            else
            {
            }
        }
        catch (Exception ex)
        {
        }
    }

    private void SetAnalysisPlanDay(int dayIdx)
    {
        _analysisPlanSelectedDay = _analysisPlanSelectedDay == dayIdx ? -1 : dayIdx;
        // BASE(-2) 모드는 전체(-1) 뷰로 로딩 (OR 합산)
        LoadAnalysisPlanFromDb(_analysisPlanSelectedDay == -2 ? -1 : _analysisPlanSelectedDay);
        Show1.Content = BuildFacilityListPanel();
        LogContentChange("Show1", Show1.Content as Control);
        Show2.Content = BuildAnalysisPlanPanel();
        LogContentChange("Show2", Show2.Content as Control);
    }


    private void Analysis_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Analysis";
        _analysisPage ??= new AnalysisPage();
        Show1.Content = _analysisPage;
        Show2.Content = null;
        Show4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("분석 시작", "새로고침", "데이터 추가", "선택 삭제", "엑셀 내보내기", "인쇄", "설정");
        SetLeftPanelWidth(380);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 1, lowerStar: 0);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Analysis");
    }

    private void Contract_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Contract";

        if (_contractPage == null)
        {
            _contractPage = new ContractPage();
            _contractPage.ParentMainPage = this;
            _contractPage.DetailPanelChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _contractPage.EditPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
            _contractPage.StatsPanelChanged += panel =>
            {
                Show4.Content = panel;
                LogContentChange("Show4", panel);
            };
        }

        // 최초 진입 시 단가·수량 컬럼 보장 (없는 컬럼 자동 추가)
        try { ContractService.EnsureContractPriceColumns(); }
        catch (Exception ex) { }
        try { ContractService.EnsureContractQuantityColumns(); }
        catch (Exception ex) { }

        Show1.Content = _contractPage;
        LogContentChange("Show1", _contractPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _contractPage.LoadData();
        _bt1SaveAction = _contractPage.SaveSelected;

        SetSubMenu("저장", "새로고침", "업체 추가", "선택 삭제", "Excel 가져오기", "인쇄", "설정");
        SetLeftPanelWidth(350);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 7, lowerStar: 3);
        // Show3: 단가표, Show4: 단가 편집폼
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Contract");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _contractPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    // ── 리스크관리 ────────────────────────────────────────────────────────────
    private void Reagent_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "RiskManage";
        _riskPage ??= new RiskManagePage();
        _riskPage.SelectCategory("시약");

        Show1.Content = _riskPage.TreeControl;
        LogContentChange("Show1", _riskPage.TreeControl);
        Show2.Content = _riskPage.UsageControl;
        LogContentChange("Show2", _riskPage.UsageControl);
        Show3.Content = _riskPage.FormControl;
        LogContentChange("Show3", _riskPage.FormControl);
        Show4.Content = _riskPage.DetailControl;
        LogContentChange("Show4", _riskPage.DetailControl);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "삭제", "", "", "", "", "", "");
        SetLeftPanelWidth(320);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 5, lowerStar: 5);
        RestoreModeLayout("RiskManage", minLowerStar: 3);
    }

    private void Glassware_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "RiskManage";
        _riskPage ??= new RiskManagePage();
        _riskPage.SelectCategory("초자");

        Show1.Content = _riskPage.TreeControl;
        LogContentChange("Show1", _riskPage.TreeControl);
        Show2.Content = _riskPage.UsageControl;
        LogContentChange("Show2", _riskPage.UsageControl);
        Show3.Content = _riskPage.FormControl;
        LogContentChange("Show3", _riskPage.FormControl);
        Show4.Content = _riskPage.DetailControl;
        LogContentChange("Show4", _riskPage.DetailControl);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "삭제", "", "", "", "", "", "");
        SetLeftPanelWidth(320);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 5, lowerStar: 5);
        RestoreModeLayout("RiskManage", minLowerStar: 3);
    }

    // ── 보수요청 ──────────────────────────────────────────────────────────────
    private void Repair_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Repair";
        _repairPage ??= new RepairPage();

        Show1.Content = _repairPage.TreeControl;
        LogContentChange("Show1", _repairPage.TreeControl);
        Show2.Content = _repairPage.ListControl;
        LogContentChange("Show2", _repairPage.ListControl);
        Show3.Content = _repairPage.FormControl;
        LogContentChange("Show3", _repairPage.FormControl);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "승인", "반려", "완료", "삭제", "", "설정");
        SetLeftPanelWidth(220);
        // Content2(목록) 위, Content3(폼) 아래 30% 표시
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 7, lowerStar: 3);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Repair", minLowerStar: 3);

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _repairPage.Refresh(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    // ── 견적/의뢰서 ───────────────────────────────────────────────────────
    private void Quotation_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Quotation";
        try { ETA.Services.SERVICE1.QuotationService.EnsureTradeStatementColumn(); } catch { }

        // ── 초기화 순서 중요: CheckPanel → Page → NewPanel → DetailPanel → HistoryPanel

        // Content3: 분석항목 체크박스 (먼저 생성)
        _quotationCheckPanel ??= new QuotationCheckPanel();

        // Content4: 계약업체 목록
        _quotationPage ??= new QuotationPage();

        // Content2-A: 신규작성 폼
        if (_quotationNewPanel == null)
        {
            _quotationNewPanel = new QuotationNewPanel();
            // 저장 완료 → 히스토리 갱신 후 저장된 issue를 DetailPanel에 표시
            _quotationNewPanel.SaveCompleted += savedIssue =>
            {
                _lastShownIssue = savedIssue;
                _quotationHistoryPanel?.LoadData();
                // Content1 트리뷰 강제 갱신
                Show1.Content = null;
                Show1.Content = _quotationHistoryPanel;
                LogContentChange("Show1", _quotationHistoryPanel);
                // DB에서 최신 row 재조회 후 DetailPanel 갱신
                _quotationDetailPanel?.ShowIssue(savedIssue);
                Show2.Content = _quotationDetailPanel;
                LogContentChange("Show2", _quotationDetailPanel);
            };
        }
        // CheckPanel → 모드에 따라 NewPanel / DetailPanel 연동
        _quotationCheckPanel.SelectionChanged    -= OnCheckSelectionChanged;
        _quotationCheckPanel.SelectionChanged    += OnCheckSelectionChanged;
        _quotationCheckPanel.AnalysisRecordSaved -= OnAnalysisRecordSaved;
        _quotationCheckPanel.AnalysisRecordSaved += OnAnalysisRecordSaved;
        _quotationCheckPanel.IssueSaved          -= OnIssueSaved;
        _quotationCheckPanel.IssueSaved          += OnIssueSaved;
        // 업체 → NewPanel 연동
        _quotationPage.CompanySelected -= OnCompanySelected;
        _quotationPage.CompanySelected += OnCompanySelected;

        // Content2-B: 세부내역 패널
        if (_quotationDetailPanel == null)
        {
            _quotationDetailPanel = new QuotationDetailPanel();
            _quotationDetailPanel.CheckPanel = _quotationCheckPanel;

            // 🥕 당근: 재활용 → NewPanel 에서 신규 번호·날짜로 작성
            _quotationDetailPanel.CarrotRequested += issue =>
            {
                var target = issue ?? _lastShownIssue;
                if (target == null) return;
                _lastShownIssue = target;
                _quotationNewPanel!.LoadFromIssue(target);
                Show2.Content = _quotationNewPanel;
                LogContentChange("Show2", _quotationNewPanel);
            };

            // 💾 저장 완료 → Show1 트리뷰 리프레시
            _quotationDetailPanel.SaveCompleted += () =>
            {
                _quotationHistoryPanel?.LoadData();
                Show1.Content = null;
                Show1.Content = _quotationHistoryPanel;
                LogContentChange("Show1", _quotationHistoryPanel);
            };

            // ✏️ 오작성 수정: DetailPanel에서 직접 편집 (더 이상 NewPanel로 이동 안 함)

            // ESC 취소 → DetailPanel 복귀 + 마지막 issue 재표시
            _quotationNewPanel!.EscapeCancelled += () =>
            {
                if (_lastShownIssue != null)
                    _quotationDetailPanel?.ShowIssue(_lastShownIssue);
                Show2.Content = _quotationDetailPanel;
                LogContentChange("Show2", _quotationDetailPanel);
            };

            // 📋 의뢰서 작성: 편집 패널로 전환
            _quotationDetailPanel.OrderRequestEditRequested += (issue, samples, quotedItems) =>
            {
                if (_orderRequestEditPanel == null)
                {
                    _orderRequestEditPanel = new OrderRequestEditPanel();
                    _orderRequestEditPanel.SubmitCompleted += () =>
                    {
                        // 업데이트 완료 → Show1 갱신 후 세부내역으로 복귀
                        _quotationHistoryPanel?.LoadData();
                        Show1.Content = null;
                        Show1.Content = _quotationHistoryPanel;
                        LogContentChange("Show1", _quotationHistoryPanel);
                        Show2.Content = _quotationDetailPanel;
                        LogContentChange("Show2", _quotationDetailPanel);
                    };
                    _orderRequestEditPanel.Cancelled += () =>
                    {
                        Show2.Content = _quotationDetailPanel;
                        LogContentChange("Show2", _quotationDetailPanel);
                    };
                }
                _orderRequestEditPanel.Load(issue, samples, quotedItems);
                Show2.Content = _orderRequestEditPanel;
                LogContentChange("Show2", _orderRequestEditPanel);
            };
        }
        _quotationDetailPanel.CheckPanel = _quotationCheckPanel;

        // Content1: 발행내역 + 분석의뢰내역 토글 트리
        if (_quotationHistoryPanel == null)
        {
            _quotationHistoryPanel = new QuotationHistoryPanel();

            // 견적발행내역 노드 선택
            _quotationHistoryPanel.IssueSelected += issue =>
            {
                _lastShownIssue = issue;
                _quotationDetailPanel!.ShowIssue(issue);
                Show2.Content = _quotationDetailPanel;
                LogContentChange("Show2", _quotationDetailPanel);
                Show3.Content = _quotationPage;
                LogContentChange("Show3", _quotationPage);
                // 편집 대상 설정
                _quotationCheckPanel!.CurrentIssue = issue;
            };

            // 분석의뢰내역 탭으로 전환됨 → Content4: 의뢰 리스트 + TODO 패널
            _quotationHistoryPanel.AnalysisTabActivated += () =>
            {
                EnsureAnalysisRequestListPanel();
                Show3.Content = _analysisRequestListPanel;
                LogContentChange("Show3", _analysisRequestListPanel);
            };

            // 견적발행내역 탭으로 복귀 → Show3: 계약업체 목록
            _quotationHistoryPanel.QuotationTabActivated += () =>
            {
                Show3.Content = _quotationPage;
                LogContentChange("Show3", _quotationPage);
            };

            // 분석의뢰내역 노드 선택
            _quotationHistoryPanel.AnalysisRequestSelected += record =>
            {
                _analysisRequestDetailPanel ??= new AnalysisRequestDetailPanel();
                _analysisRequestDetailPanel.CheckPanel = _quotationCheckPanel;
                _analysisRequestDetailPanel.ShowRecord(record);
                Show2.Content = _analysisRequestDetailPanel;
                LogContentChange("Show2", _analysisRequestDetailPanel);
                EnsureAnalysisRequestListPanel();
                Show3.Content = _analysisRequestListPanel;
                LogContentChange("Show3", _analysisRequestListPanel);
                _analysisRequestListPanel!.AddRecord(record);
                // 편집 대상 설정
                _quotationCheckPanel!.CurrentAnalysisRecord = record;
            };
        }

        Show1.Content = _quotationHistoryPanel;
        LogContentChange("Show1", _quotationHistoryPanel);
        Show2.Content = _quotationDetailPanel;
        LogContentChange("Show2", _quotationDetailPanel);
        Show3.Content = _quotationPage;
        LogContentChange("Show3", _quotationPage);
        Show4.Content = _quotationCheckPanel;
        LogContentChange("Show4", _quotationCheckPanel);
        _bt1SaveAction = null;

        _quotationHistoryPanel.LoadData();
        _quotationCheckPanel.LoadData();
        _quotationPage.LoadData();

        SetSubMenu("새로고침", "신규 작성", "삭제", "엑셀 내보내기", "인쇄", "", "설정");
        SetLeftPanelWidth(430);
        // Content2(세부내역) 50% : Content4(업체목록) 50%
        // 하단(Content3 분석항목) ≈ 23%  (13 : 4 → 76% : 24%)
        // Show2(세부내역) 60% | Show4(분석항목 트리뷰) 40% / 하단 Show3(업체목록) 23%
        SetContentLayout(content2Star: 6, content4Star: 4, upperStar: 13, lowerStar: 3);

        // 저장된 레이아웃 복원
        RestoreModeLayout("Quotation", minLowerStar: 3);
    }

    // ── 분석의뢰 리스트 패널 — 생성 + Show4→Show2 이벤트 1회 연결 ──────────
    private void EnsureAnalysisRequestListPanel()
    {
        if (_analysisRequestListPanel != null) return;
        _analysisRequestListPanel = new AnalysisRequestListPanel();
        _analysisRequestListPanel.RecordSelected += rec =>
        {
            _analysisRequestDetailPanel ??= new AnalysisRequestDetailPanel();
            _analysisRequestDetailPanel.CheckPanel = _quotationCheckPanel;
            _analysisRequestDetailPanel.ShowRecord(rec);
            Show2.Content = _analysisRequestDetailPanel;
            LogContentChange("Show2 (Show3 선택)", _analysisRequestDetailPanel);
            if (_quotationCheckPanel != null)
                _quotationCheckPanel.CurrentAnalysisRecord = rec;
        };
    }

    // ── 견적발행 ──────────────────────────────────────────────────────────
    private void QuotationIssuing_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "QuotationIssue";
        try { ETA.Services.SERVICE1.QuotationService.EnsureQuotationIssueTable(); } catch { }
        try { ETA.Services.SERVICE1.QuotationService.EnsureTradeStatementColumn(); } catch { }

        // Show1: 견적발행내역 트리 (IssuingMode = true → 클릭 시 Show4 체크리스트에 추가)
        if (_issuingHistoryPanel == null)
        {
            _issuingHistoryPanel = new QuotationHistoryPanel { IssuingMode = true };
            _issuingHistoryPanel.IssueAddedToList += issue =>
            {
                AddToIssuingChecklist(issue);
            };
        }

        // Show4: 체크리스트 컨테이너
        _issuingChecklist.Clear();

        Show1.Content = _issuingHistoryPanel;
        LogContentChange("Show1", _issuingHistoryPanel);
        Show2.Content = null;
        Show3.Content = null;
        Show4.Content = BuildIssuingChecklistContainer();
        _bt1SaveAction = null;

        _issuingHistoryPanel.LoadData();

        SetSubMenu("새로고침", "", "", "", "", "거래명세서 발행", "");
        SetLeftPanelWidth(430);
        SetContentLayout(content2Star: 7, content4Star: 3, upperStar: 13, lowerStar: 4);
        RestoreModeLayout("QuotationIssue", minLowerStar: 4);
    }

    private Border BuildIssuingChecklistContainer()
    {
        _issuingChecklistPanel = new StackPanel { Spacing = 4 };
        _issuingChecklistScroll = new ScrollViewer
        {
            Content = _issuingChecklistPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };

        var header = new TextBlock
        {
            Text = "발행 대상 목록",
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontXL,
            Foreground = Res("AppFg"),
            Margin = new Thickness(8, 6, 0, 4),
        };

        var clearBtn = new Button
        {
            Content = "전체 지우기",
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontBase,
            Padding = new Thickness(6, 2),
            Background = Res("SubBtnBg"),
            Foreground = Res("FgMuted"),
            Margin = new Thickness(0, 0, 8, 0),
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        clearBtn.Click += (_, _) =>
        {
            _issuingChecklist.Clear();
            _issuingChecklistPanel!.Children.Clear();
            Show2.Content = null;
        };

        var headerRow = new Grid();
        headerRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        headerRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        Grid.SetColumn(header, 0);
        Grid.SetColumn(clearBtn, 1);
        headerRow.Children.Add(header);
        headerRow.Children.Add(clearBtn);

        var root = new StackPanel { Spacing = 0 };
        root.Children.Add(headerRow);
        root.Children.Add(new Border
        {
            BorderBrush = Res("InputBorder"),
            BorderThickness = new Thickness(0, 1, 0, 0),
            Margin = new Thickness(0, 0, 0, 4),
        });
        root.Children.Add(_issuingChecklistScroll);

        return new Border
        {
            Background = Res("PanelBg"),
            Child = root,
        };
    }

    private void AddToIssuingChecklist(ETA.Models.QuotationIssue issue)
    {
        // 중복 방지
        if (_issuingChecklist.Any(x => x.Id == issue.Id)) return;
        _issuingChecklist.Add(issue);

        if (_issuingChecklistPanel == null) return;

        var cb = new CheckBox
        {
            IsChecked = true,
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontMD,
            Foreground = Res("AppFg"),
            Content = $"{issue.약칭}  {issue.시료명}  [{issue.견적번호}]",
            Tag = issue,
            Margin = new Thickness(4, 2),
        };
        cb.IsCheckedChanged += (_, _) => RefreshIssuingPreview();
        _issuingChecklistPanel.Children.Add(cb);

        RefreshIssuingPreview();
    }

    private void RefreshIssuingPreview()
    {
        // 체크된 항목 수집
        var checkedIssues = _issuingChecklistPanel?.Children
            .OfType<CheckBox>()
            .Where(cb => cb.IsChecked == true && cb.Tag is ETA.Models.QuotationIssue)
            .Select(cb => (ETA.Models.QuotationIssue)cb.Tag!)
            .ToList() ?? new List<ETA.Models.QuotationIssue>();

        if (checkedIssues.Count == 0)
        {
            Show2.Content = new TextBlock
            {
                Text = "체크된 항목이 없습니다.",
                FontFamily = _tradeFont,
                FontSize = AppTheme.FontLG,
                Foreground = Res("FgMuted"),
                Margin = new Thickness(16),
            };
            return;
        }

        // 항목별 수량/금액 집계
        var aggData = Task.Run(() =>
        {
            return ETA.Services.SERVICE1.QuotationService.AggregateIssueItems(checkedIssues);
        });

        aggData.ContinueWith(t =>
        {
            Avalonia.Threading.Dispatcher.UIThread.Post(() => BuildIssuingPreviewPanel(checkedIssues, t.Result));
        });
    }

    private void BuildIssuingPreviewPanel(
        List<ETA.Models.QuotationIssue> issues,
        List<(string 항목, int 수량, decimal 금액)> rows)
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        // 헤더
        var companies = string.Join(", ", issues.Select(i => i.약칭).Distinct());
        root.Children.Add(new TextBlock
        {
            Text = $"업체: {companies}",
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontLG,
            Foreground = Res("AppFg"),
        });
        root.Children.Add(new TextBlock
        {
            Text = $"견적번호: {string.Join(", ", issues.Select(i => i.견적번호))}",
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontBase,
            Foreground = Res("FgMuted"),
        });
        root.Children.Add(new Border
        {
            BorderBrush = Res("InputBorder"),
            BorderThickness = new Thickness(0, 1, 0, 0),
            Margin = new Thickness(0, 4, 0, 4),
        });

        // 테이블 헤더
        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        grid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        grid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        TextBlock Cell(string text, bool isHeader = false) => new()
        {
            Text = text,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = isHeader ? 12 : 11,
            Foreground = Res(isHeader ? "AppFg" : "FgMuted"),
            Margin = new Thickness(4, 2),
            TextAlignment = TextAlignment.Right,
        };

        var hItem   = Cell("항목", true);
        var hQty    = Cell("수량", true);
        var hAmt    = Cell("금액", true);
        hItem.TextAlignment = TextAlignment.Left;
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        Grid.SetRow(hItem, 0); Grid.SetColumn(hItem, 0);
        Grid.SetRow(hQty,  0); Grid.SetColumn(hQty,  1);
        Grid.SetRow(hAmt,  0); Grid.SetColumn(hAmt,  2);
        grid.Children.Add(hItem);
        grid.Children.Add(hQty);
        grid.Children.Add(hAmt);

        decimal totalAmt = 0;
        for (int i = 0; i < rows.Count; i++)
        {
            var (항목, 수량, 금액) = rows[i];
            totalAmt += 금액;
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
            int row = i + 1;
            var cItem = Cell(항목);    cItem.TextAlignment = TextAlignment.Left;
            var cQty  = Cell(수량.ToString("N0"));
            var cAmt  = Cell(금액.ToString("N0"));
            Grid.SetRow(cItem, row); Grid.SetColumn(cItem, 0); grid.Children.Add(cItem);
            Grid.SetRow(cQty,  row); Grid.SetColumn(cQty,  1); grid.Children.Add(cQty);
            Grid.SetRow(cAmt,  row); Grid.SetColumn(cAmt,  2); grid.Children.Add(cAmt);
        }

        // 합계
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        int totalRow = rows.Count + 1;
        var tLabel = Cell("합계", true); tLabel.TextAlignment = TextAlignment.Left;
        var tAmt   = Cell(totalAmt.ToString("N0"), true);
        Grid.SetRow(tLabel, totalRow); Grid.SetColumn(tLabel, 0); grid.Children.Add(tLabel);
        Grid.SetRow(tAmt,   totalRow); Grid.SetColumn(tAmt,   2); grid.Children.Add(tAmt);

        root.Children.Add(grid);

        Show2.Content = new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };
    }

    private void Purchase_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Purchase";
        _purchasePage ??= new PurchasePage();

        Show1.Content = _purchasePage.TreeControl;
        LogContentChange("Show1", _purchasePage.TreeControl);
        Show2.Content = _purchasePage.ListControl;
        LogContentChange("Show2", _purchasePage.ListControl);
        Show3.Content = _purchasePage.FormControl;
        LogContentChange("Show3", _purchasePage.FormControl);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "엑셀 내보내기", "승인", "반려", "완료", "삭제", "설정");
        SetLeftPanelWidth(250);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 8, lowerStar: 2);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("Purchase");
    }


    // ── 출장/일정 관리 ────────────────────────────────────────────────────────
    private void Schedule_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Schedule";
        _schedulePage ??= new SchedulePage();

        Show1.Content = _schedulePage.TreeControl;
        Show2.Content = _schedulePage.CalendarControl;
        Show3.Content = _schedulePage.FormControl;
        Show4.Content = null;

        LogContentChange("Show1", _schedulePage.TreeControl);
        SetSubMenu("저장", "새로고침", "", "", "", "", "");
        SetLeftPanelWidth(250);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 6, lowerStar: 4);
        RestoreModeLayout("Schedule");

        Avalonia.Threading.Dispatcher.UIThread.Post(() => _schedulePage.LoadData());
    }

    private void TestReport_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "TestReport";

        if (_testReportPage == null)
        {
            _testReportPage = new TestReportPage();
            _testReportPage.ResultListChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
            _testReportPage.EditPanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
        }

        Show1.Content = _testReportPage;
        LogContentChange("Show1", _testReportPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);

        // Content4: 출력 보관함 (Reports 폴더)
        _reportsPanel ??= new ReportsPanel();
        _reportsPanel.LoadFiles();
        Show4.Content = _reportsPanel;
        LogContentChange("Show4", _reportsPanel);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "삭제", "엑셀 출력", "", "일괄 엑셀");

        SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 8.5, lowerStar: 1.5);
        
        // 저장된 레이아웃 복원
        RestoreModeLayout("TestReport");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _testReportPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void WasteTestReport_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteTestReport";

        if (_wasteTestReportPage == null)
        {
            _wasteTestReportPage = new WasteTestReportPage();
            _wasteTestReportPage.ResultGridChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
        }

        Show1.Content = _wasteTestReportPage;
        LogContentChange("Show1", _wasteTestReportPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "", "", "", "");
        SetContentLayout(content2Star: 8, content4Star: 0, upperStar: 9, lowerStar: 1);
        RestoreModeLayout("WasteTestReport");

        Avalonia.Threading.Dispatcher.UIThread.Post(
            () => _wasteTestReportPage.LoadData(),
            Avalonia.Threading.DispatcherPriority.Render);
    }

    private void DbMigration_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "DbMigration";

        // Show1: 변경 전 업체 선택 (구 이름)
        if (_dbMigrationPage == null)
        {
            _dbMigrationPage = new DbMigrationPage { PanelTitle = "🔴 변경 전 업체 (구 이름)" };
            _dbMigrationPage.CompanySelected += company =>
            {
                _migrationOldName = company?.C_CompanyName;
                RefreshMigrationShow2();
            };
        }

        // Show4: 변경 후 업체 선택 (새 이름)
        if (_dbMigrationTargetPage == null)
        {
            _dbMigrationTargetPage = new DbMigrationPage { PanelTitle = "🟢 변경 후 업체 (새 이름)" };
            _dbMigrationTargetPage.CompanySelected += company =>
            {
                _migrationNewName = company?.C_CompanyName;
                RefreshMigrationShow2();
            };
        }

        _migrationOldName = null;
        _migrationNewName = null;

        Show1.Content = _dbMigrationPage;
        LogContentChange("Show1", _dbMigrationPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = _dbMigrationTargetPage;
        LogContentChange("Show4", _dbMigrationTargetPage);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "", "", "", "");
        SetContentLayout(content2Star: 4, content4Star: 5, upperStar: 8.5, lowerStar: 1.5);
        RestoreModeLayout("DbMigration");

        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        {
            _dbMigrationPage.LoadData();
            _dbMigrationTargetPage.LoadData();
        }, Avalonia.Threading.DispatcherPriority.Render);
    }

    private void RefreshMigrationShow2()
    {
        if (string.IsNullOrWhiteSpace(_migrationOldName) || string.IsNullOrWhiteSpace(_migrationNewName))
        {
            Show2.Content = null;
            LogContentChange("Show2", null);
            return;
        }
        var panel = BuildMigrationConfirmPanel(_migrationOldName, _migrationNewName);
        Show2.Content = panel;
        LogContentChange("Show2", panel);
    }

    private Border BuildMigrationConfirmPanel(string oldName, string newName)
    {
        var font  = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var fontR = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        // 결과 메시지 TextBlock (나중에 버튼 핸들러에서 참조)
        var txbMigResult = new TextBlock
        {
            Text       = "",
            FontSize   = AppTheme.FontLG,
            FontFamily = fontR,
            Foreground = AppTheme.FgSuccess,
            Margin     = new Thickness(0, 12, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        };

        var btnExecute = new Button
        {
            Content    = "▶ 변경 실행",
            FontSize   = AppTheme.FontXL,
            FontFamily = font,
            FontWeight = FontWeight.Bold,
            Background = AppTheme.BorderActive,
            Foreground = Brush.Parse("#88ff88"),
            BorderBrush = AppTheme.BorderActive,
            BorderThickness = new Thickness(1),
            Padding    = new Thickness(16, 8),
            CornerRadius = new CornerRadius(5),
            Margin     = new Thickness(0, 16, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left,
        };

        // 캡처용 로컬 변수 (클로저)
        var capturedOld = oldName;
        var capturedNew = newName;

        btnExecute.Click += (_, _) =>
        {
            btnExecute.IsEnabled = false;
            txbMigResult.Text    = "처리 중...";
            txbMigResult.Foreground = Brush.Parse("#eeeeaa");

            var (updatedRows, error) = ETA.Services.SERVICE1.CompanyRenameService.RenameCompany(capturedOld, capturedNew);

            if (string.IsNullOrEmpty(error))
            {
                txbMigResult.Text       = $"✅ 완료! {updatedRows}건 변경";
                txbMigResult.Foreground = AppTheme.FgSuccess;
                _dbMigrationPage?.LoadData();
                _dbMigrationTargetPage?.LoadData();
            }
            else
            {
                txbMigResult.Text       = $"❌ 오류: {error}";
                txbMigResult.Foreground = AppTheme.FgDanger;
                btnExecute.IsEnabled    = true;
            }
        };

        var content = new StackPanel
        {
            Spacing = 6,
            Margin  = new Thickness(16),
            Children =
            {
                new TextBlock
                {
                    Text       = "업체명 변경 확인",
                    FontSize   = AppTheme.FontXL, FontWeight = FontWeight.Bold,
                    FontFamily = font,
                    Foreground = AppTheme.FgPrimary,
                    Margin     = new Thickness(0, 0, 0, 8),
                },
                new TextBlock
                {
                    Text       = $"변경 전:  {oldName}",
                    FontSize   = AppTheme.FontLG, FontFamily = fontR,
                    Foreground = AppTheme.FgDanger,
                },
                new TextBlock
                {
                    Text       = $"변경 후:  {newName}",
                    FontSize   = AppTheme.FontLG, FontFamily = fontR,
                    Foreground = Brush.Parse("#99ff99"),
                },
                new TextBlock
                {
                    Text       = "아래 테이블의 업체명이 일괄 변경됩니다:",
                    FontSize   = AppTheme.FontMD, FontFamily = fontR,
                    Foreground = AppTheme.FgMuted,
                    Margin     = new Thickness(0, 10, 0, 0),
                },
                new TextBlock
                {
                    Text       = "  • 견적발행내역\n  • 거래명세서발행내역\n  • 계약 DB\n  • 수질분석센터_결과\n  • 시료명칭(컬럼명)",
                    FontSize   = AppTheme.FontMD, FontFamily = fontR,
                    Foreground = AppTheme.FgSecondary,
                    LineHeight = 20,
                },
                btnExecute,
                txbMigResult,
            },
        };

        return new Border
        {
            Background   = Brush.Parse("#1a1e2a"),
            BorderBrush  = AppTheme.BorderAccent,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Child        = new ScrollViewer
            {
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                Content = content,
            },
        };
    }

    private void Admin_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Admin";
        Show1.Content = null;
        Show2.Content = null;
        Show4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("사용자 관리", "권한 설정", "로그 확인", "백업하기", "시스템 설정", "통계", "종료");
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
    }

    private void Permission_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Permission";
        Show1.Content = null;
        Show2.Content = null;
        Show4.Content = null;
        _bt1SaveAction = null;
        SetSubMenu("", "", "", "", "", "", "");
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
    }

    private void Access_Click(object? sender, RoutedEventArgs e)
    {
        if (CurrentEmployeeId != "201000308") return;
        _currentMode = "Access";

        _accessPage ??= new AccessPage();

        Show1.Content = _accessPage.Show1;
        Show2.Content = _accessPage.Show2;
        Show3.Content = null;
        Show4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "", "", "", "", "");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 2, content4Star: 0, upperStar: 1, lowerStar: 0);
        RestoreModeLayout("Access");
    }


    private void AiDocClassification_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "AiDocClassification";

        if (_aiDocClassificationPage == null)
        {
            _aiDocClassificationPage = new AiDocClassificationPage();
            _aiDocClassificationPage.ListPanelChanged   += panel => Show2.Content = panel;
            _aiDocClassificationPage.DetailPanelChanged += panel => Show3.Content = panel;
            _aiDocClassificationPage.StatsPanelChanged  += panel => Show4.Content = panel;
        }

        Show1.Content = _aiDocClassificationPage;
        Show2.Content = null;
        Show3.Content = null;
        _aiDocClassificationPage.RefreshShow4();
        _bt1SaveAction = null;

        SetSubMenu("학습", "새로고침", "", "", "", "", "");
        SetLeftPanelWidth(220);
        SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 7, lowerStar: 3);
        RestoreModeLayout("AiDocClassification");
    }

    private void ParserGenerator_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "ParserGenerator";

        if (_parserGeneratorPage == null)
        {
            _parserGeneratorPage = new ParserGeneratorPage();
            _parserGeneratorPage.ListPanelChanged   += panel => Show2.Content = panel;
            _parserGeneratorPage.DetailPanelChanged += panel => Show3.Content = panel;
            _parserGeneratorPage.StatsPanelChanged  += panel => Show4.Content = panel;
        }

        Show1.Content = _parserGeneratorPage;
        Show2.Content = null;
        Show3.Content = null;
        Show4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("파일 추가", "", "", "", "", "", "");
        SetLeftPanelWidth(220);
        SetContentLayout(content2Star: 8, content4Star: 2, upperStar: 7, lowerStar: 3);
        RestoreModeLayout("ParserGenerator");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  서브메뉴 버튼
    // ══════════════════════════════════════════════════════════════════════

    // ── 모드 → 네비게이션 표시명 매핑 ──
    private static readonly Dictionary<string, (string Main, string Sub)> NavMap = new()
    {
        ["Agent"]                    = ("직원관리", "업무분장"),
        ["Contract"]                 = ("계약관리", "계약"),
        ["Analysis"]                 = ("분석관리", "분석"),
        ["TestReport"]               = ("분석관리", "시험성적서"),
        ["Quotation"]                = ("견적관리", "견적/발행"),
        ["QuotationIssue"]           = ("견적관리", "분석의뢰"),
        ["WasteCompany"]             = ("폐수관리", "배출업소"),
        ["WasteNameReconcile"]       = ("폐수관리", "약칭매칭(폐수)"),
        ["WaterQualityNameReconcile"]= ("폐수관리", "약칭매칭(수질)"),
        ["WasteSampleList"]          = ("폐수관리", "시료관리"),
        ["ProcessingFacility"]       = ("폐수관리", "처리시설"),
        ["WasteAnalysisInput"]       = ("폐수관리", "분석입력"),
        ["ResultSubmitMeasure"]      = ("결과제출", "측정인"),
        ["ResultSubmitErp"]          = ("결과제출", "ERP"),
        ["ResultSubmitZero4"]        = ("결과제출", "Zero4"),
        ["Schedule"]                 = ("일정관리", "스케줄"),
        ["Purchase"]                 = ("업무지원", "구매요청"),
        ["Repair"]                   = ("업무지원", "보수요청"),
        ["RiskManage"]               = ("업무지원", "위험관리"),
        ["Access"]                   = ("권한관리", "접근관리"),
        ["MyTask"]                   = ("나의업무", "할 일"),
        ["DbMigration"]              = ("설정", "DB 마이그레이션"),
    };

    private void UpdateNavPath()
    {
        if (NavMap.TryGetValue(_currentMode, out var nav))
            txtNavPath.Text = $"{nav.Main}  ›  {nav.Sub}";
        else
            txtNavPath.Text = "";
    }

    private void SetSubMenu(string bt1, string bt2, string bt3,
                            string bt4, string bt5, string bt6,
                            string bt7 = "", string bt8 = "", string bt9 = "")
    {
        SetBtn(BT1, bt1); SetBtn(BT2, bt2); SetBtn(BT3, bt3);
        SetBtn(BT4, bt4); SetBtn(BT5, bt5);
        SetBtn(BT6, bt6); SetBtn(BT7, bt7); SetBtn(BT8, bt8); SetBtn(BT9, bt9);
        SubMenu.IsVisible = new[] { bt1, bt2, bt3, bt4, bt5, bt6, bt7, bt8, bt9 }
            .Any(s => !string.IsNullOrWhiteSpace(s));
        UpdateNavPath();

        // (shimmer 제거됨)
    }

    private static void SetBtn(Avalonia.Controls.Button btn, string label)
    {
        btn.IsVisible = !string.IsNullOrWhiteSpace(label);
        btn.Content   = label;
    }

    private async void BT1_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Schedule":        _schedulePage?.SaveEntry();         break;
            case "Purchase":        _purchasePage?.Refresh();           break;
            case "TestReport":      _testReportPage?.LoadData();        break;
            case "ResultSubmitMeasure": _resultSubmitMeasureTestReport?.LoadData(); break;
            case "Repair":          _repairPage?.Refresh();             break;
            case "RiskManage":      _riskPage?.Refresh();               break;
            case "WasteSampleList":
                if (_wasteSampleListPage != null)
                {
                    SaveProgressPanel.IsVisible = true;
                    pbSave.Value = 0;
                    txtSaveProgress.Text = "저장 중...";
                    var prog = new Progress<(int done, int total)>(p =>
                    {
                        pbSave.Value = p.total > 0 ? (double)p.done / p.total : 0;
                        txtSaveProgress.Text = $"{p.done}/{p.total}";
                    });
                    _wasteSampleListPage.SavePending(prog);
                    pbSave.Value = 1;
                    txtSaveProgress.Text = "완료";
                    await Task.Delay(1200);
                    SaveProgressPanel.IsVisible = false;
                }
                break;
            case "MyTask":          _myTaskPage?.LoadData();             break;
            case "Access":
                _accessPage = new AccessPage();
                Show1.Content = _accessPage.Show1;
                Show2.Content = _accessPage.Show2;
                break;
            case "WasteAnalysisInput": _wasteAnalysisInputPage?.LoadData(); break;
            case "AiDocClassification":
                if (_aiDocClassificationPage != null)
                    _ = _aiDocClassificationPage.LearnAsync();
                break;
            case "ParserGenerator":
                if (_parserGeneratorPage != null)
                    _ = _parserGeneratorPage.UploadAndAnalyzeAsync();
                break;
            default: _bt1SaveAction?.Invoke();                          break;
        }
    }

    private void BT2_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":        _agentTreePage?.LoadData();      break;
            case "WasteCompany":      _ = _wasteCompanyPage?.LoadDataAsync();    break;
            case "WasteNameReconcile": _wasteNameReconcilePage?.Reload(); break;
            case "Schedule":     _schedulePage?.LoadData();        break;
            case "Contract":     _contractPage?.LoadData();       break;
            case "Purchase":     _purchasePage?.ExportCsv();      break;
            case "TestReport":   _testReportPage?.SaveCsv();      break;
            case "Quotation":
                // BT2 = 신규 작성 → Content2 를 NewPanel 로 교체
                _quotationNewPanel?.Clear();
                if (_quotationCheckPanel != null)
                {
                    _quotationCheckPanel.CurrentAnalysisRecord = null;
                    _quotationCheckPanel.CurrentIssue = null;
                }
                Show2.Content = _quotationNewPanel;
                break;
            case "WasteSampleList": _wasteSampleListPage?.LoadData(); _wasteSampleListPage?.LoadCompanyTree(forceReload: true); break;
            case "WasteAnalysisInput": _wasteAnalysisInputPage?.VerifyData(); break;
            case "Repair":          _repairPage?.ApproveSelected();     break;
            case "RiskManage":      _riskPage?.DeleteSelected();        break;
            case "AnalysisPlan":    SetAnalysisPlanDay(0); break; // 월
            case "AiDocClassification": AiDocClassification_Click(null, null!); break;
            case "Ecotoxicity":     _ = _ecotoxicityPage?.ExportTestReportAsync(); break;
            default: break;
        }
    }

    private void BT3_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":        _agentTreePage?.ShowAddPanel();           break;
            case "Contract":     _contractPage?.ShowAddPanel();            break;
            case "WasteCompany": _wasteCompanyPage?.ShowAddPanel();        break;
            case "Purchase":   _purchasePage?.ApproveSelected();         break;
            case "Quotation":  _quotationPage?.LoadData(); _quotationHistoryPanel?.LoadData(); break;
            case "TestReport": _ = _testReportPage?.DeleteSampleAsync(); break;
            case "WasteAnalysisInput": _ = _wasteAnalysisInputPage?.ImportData(); break;
            case "Repair":          _repairPage?.RejectSelected();           break;
            case "AnalysisPlan":    SetAnalysisPlanDay(1); break; // 화
            default: break;
        }
    }

    private async void BT4_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":    if (_agentTreePage  != null) await _agentTreePage.DeleteSelectedAsync();  break;
            case "Contract": if (_contractPage   != null) await _contractPage.DeleteSelectedAsync();   break;
            case "Purchase": _purchasePage?.RejectSelected();   break;
            case "TestReport": _testReportPage?.OpenPrintWindow();   break;
            case "WasteAnalysisInput": _wasteAnalysisInputPage?.ExportData(); break;
            case "Repair":          _repairPage?.CompleteSelected();   break;
            case "Quotation":       await ExportQuotationAsync(); break;
            case "AnalysisPlan":    SetAnalysisPlanDay(2); break; // 수
            default: break;
        }
    }

    private void BT5_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "WasteCompany":  ShowWasteCompanyData();            break;
            case "Purchase":      _purchasePage?.CompleteSelected();  break;
            case "TestReport":    _testReportPage?.OpenPrintWindow();   break;
            case "Repair":        _repairPage?.DeleteSelected();      break;
            case "Contract":      _ = ImportContractFromExcelAsync(); break;
            case "WasteAnalysisInput": _wasteAnalysisInputPage?.AttachExcel(); break;
            case "AnalysisPlan":  SetAnalysisPlanDay(3); break; // 목
            default: break;
        }
    }

    private async Task ImportContractFromExcelAsync()
    {
        // 파일 선택
        var dialog = new Avalonia.Platform.Storage.FilePickerOpenOptions
        {
            Title            = "ETA DB Excel 파일 선택",
            AllowMultiple    = false,
            FileTypeFilter   = new[]
            {
                new Avalonia.Platform.Storage.FilePickerFileType("Excel 파일")
                {
                    Patterns = new[] { "*.xlsm", "*.xlsx" }
                }
            }
        };

        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(dialog);
        if (files.Count == 0) return;

        var filePath = files[0].Path.LocalPath;

        // 진행 표시 (Show2 임시 메시지)
        Show2.Content = new TextBlock
        {
            Text              = "⏳  Excel 가져오는 중...",
            FontSize          = AppTheme.FontLG, Foreground = Res("FgMuted"),
            Margin            = new Thickness(12),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        var (ok, errCount) = await Task.Run(() => ContractService.ImportFromExcel(filePath));

        Show2.Content = new TextBlock
        {
            Text              = errCount < 0
                                    ? "❌  Excel 파일을 열 수 없습니다."
                                    : $"✅  가져오기 완료 — 성공 {ok}건{(errCount > 0 ? $", 오류 {errCount}건" : "")}\n트리에서 업체를 선택하면 단가가 표시됩니다.",
            FontSize          = AppTheme.FontLG,
            Foreground        = errCount < 0 ? Brushes.OrangeRed : AppTheme.FgSuccess,
            Margin            = new Thickness(12),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        // 트리 새로고침
        _contractPage?.LoadData();
    }

    private void BT6_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Purchase":        _purchasePage?.DeleteSelected();        break;
            case "TestReport":      _testReportPage?.BatchPrintExcel();     break;
            case "Quotation":       ShowTradeStatementEditor();             break;
            case "QuotationIssue":  IssueTradeStatementFromChecklist();     break;
            case "WasteAnalysisInput": _ = _wasteAnalysisInputPage?.OnAiParserButtonClick(); break;
            case "AnalysisPlan":    SetAnalysisPlanDay(4); break; // 금
            default: break;
        }
    }

    private async Task ExportQuotationAsync()
    {
        var issue = _lastShownIssue;
        if (issue == null)
        {
            Show2.Content = new TextBlock
            {
                Text = "견적서를 먼저 선택하세요.",
                FontFamily = _tradeFont, FontSize = AppTheme.FontLG,
                Foreground = Brushes.OrangeRed, Margin = new Thickness(16),
            };
            return;
        }

        var window = TopLevel.GetTopLevel(this) as Window;
        var picker = new Avalonia.Platform.Storage.FilePickerSaveOptions
        {
            Title             = "견적서 저장",
            SuggestedFileName = $"견적서_{issue.약칭}_{issue.견적번호}",
            FileTypeChoices   = new[]
            {
                new Avalonia.Platform.Storage.FilePickerFileType("Excel")
                    { Patterns = new[] { "*.xlsx" } }
            }
        };
        var file = window != null ? await window.StorageProvider.SaveFilePickerAsync(picker) : null;
        if (file == null) return;

        var (ok, msg) = await Task.Run(() =>
            ETA.Services.SERVICE1.QuotationService.ExportQuotation(issue, file.Path.LocalPath));

        Show2.Content = new TextBlock
        {
            Text = ok ? $"견적서 저장 완료\n{file.Path.LocalPath}" : $"오류: {msg}",
            FontFamily = _tradeFont, FontSize = AppTheme.FontLG,
            Foreground = ok ? AppTheme.FgSuccess : Brushes.OrangeRed,
            Margin = new Thickness(16),
            TextWrapping = TextWrapping.Wrap,
        };
    }

    private async void IssueTradeStatementFromChecklist()
    {
        var checkedIssues = _issuingChecklistPanel?.Children
            .OfType<CheckBox>()
            .Where(cb => cb.IsChecked == true && cb.Tag is ETA.Models.QuotationIssue)
            .Select(cb => (ETA.Models.QuotationIssue)cb.Tag!)
            .ToList() ?? new List<ETA.Models.QuotationIssue>();

        if (checkedIssues.Count == 0)
        {
            Show2.Content = new TextBlock
            {
                Text = "발행할 항목을 체크리스트에서 선택하세요.",
                FontFamily = _tradeFont,
                FontSize = AppTheme.FontLG,
                Foreground = Brushes.OrangeRed,
                Margin = new Thickness(16),
            };
            return;
        }

        // 분석 미완료 항목 확인
        var quotNos = checkedIssues.Select(i => i.견적번호).Distinct().ToList();
        var incomplete = await Task.Run(() =>
            ETA.Services.SERVICE1.AnalysisRecordService.GetIncompleteItems(quotNos));

        if (incomplete.Count > 0)
        {
            bool proceed = await ShowIncompleteWarningAsync(incomplete);
            if (!proceed) return;
        }

        // 기존 거래명세서 발행 플로우 재사용
        _ = IssueTradeStatementAsync(checkedIssues);
    }

    private async Task<bool> ShowIncompleteWarningAsync(List<(string 시료명, string 항목)> items)
    {
        bool proceed = false;
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner == null) return true;

        var itemLines = string.Join("\n", items.Select(x => $"  • {x.시료명} : {x.항목}"));

        var btnProceed = new Button
        {
            Content = "그래도 발행",
            Padding = new Thickness(14, 6),
            Margin = new Thickness(6, 0, 0, 0),
            CornerRadius = new CornerRadius(4),
            Background = new SolidColorBrush(Color.Parse("#b91c1c")),
            Foreground = Brushes.White,
            BorderThickness = new Thickness(0),
            FontFamily = _tradeFont,
        };
        var btnCancel = new Button
        {
            Content = "취소",
            Padding = new Thickness(14, 6),
            CornerRadius = new CornerRadius(4),
            Background = Res("SubBtnBg"),
            Foreground = Res("FgMuted"),
            BorderThickness = new Thickness(0),
            FontFamily = _tradeFont,
        };

        var dlg = new Window
        {
            Title = "분석 미완료 경고",
            Width = 440,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = Res("PanelBg"),
            CanResize = false,
            Content = new StackPanel
            {
                Margin = new Thickness(18, 16, 18, 18),
                Spacing = 10,
                Children =
                {
                    new TextBlock
                    {
                        Text = "⚠  분석이 완료되지 않았습니다",
                        FontSize = AppTheme.FontXL,
                        FontWeight = FontWeight.Bold,
                        Foreground = Brushes.Orange,
                        FontFamily = _tradeFont,
                    },
                    new Border
                    {
                        Background = Res("PanelInnerBg"),
                        CornerRadius = new CornerRadius(4),
                        Padding = new Thickness(10, 8),
                        MaxHeight = 260,
                        Child = new ScrollViewer
                        {
                            Content = new TextBlock
                            {
                                Text = itemLines,
                                FontSize = AppTheme.FontMD,
                                Foreground = Res("AppFg"),
                                FontFamily = _tradeFont,
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                            },
                        },
                    },
                    new TextBlock
                    {
                        Text = "그래도 발행하시겠습니까?",
                        FontSize = AppTheme.FontLG,
                        Foreground = Res("FgMuted"),
                        FontFamily = _tradeFont,
                    },
                    new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        HorizontalAlignment = HorizontalAlignment.Right,
                        Children = { btnCancel, btnProceed },
                    },
                },
            },
        };

        btnProceed.Click += (_, _) => { proceed = true; dlg.Close(); };
        btnCancel.Click  += (_, _) => { proceed = false; dlg.Close(); };

        await dlg.ShowDialog(owner);
        return proceed;
    }

    private async Task IssueTradeStatementAsync(List<ETA.Models.QuotationIssue> issues)
    {
        var window = TopLevel.GetTopLevel(this) as Window;
        var picker = new Avalonia.Platform.Storage.FilePickerSaveOptions
        {
            Title           = "거래명세서 저장",
            SuggestedFileName = $"거래명세서_{DateTime.Today:yyyyMMdd}",
            FileTypeChoices = new[]
            {
                new Avalonia.Platform.Storage.FilePickerFileType("Excel")
                    { Patterns = new[] { "*.xlsx" } }
            }
        };

        var file = window != null
            ? await window.StorageProvider.SaveFilePickerAsync(picker)
            : null;
        if (file == null) return;

        string path = file.Path.LocalPath;

        Show2.Content = new TextBlock
        {
            Text = "거래명세서 생성 중...",
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontLG,
            Foreground = Res("FgMuted"),
            Margin = new Thickness(16),
        };

        var captured = issues;
        var (ok, msg, supply, vat, total) = await Task.Run(() =>
            ETA.Services.SERVICE1.QuotationService.ExportTradingStatement(captured, path));

        if (!ok)
        {
            Show2.Content = new TextBlock
            {
                Text = $"오류: {msg}",
                FontFamily = _tradeFont,
                FontSize = AppTheme.FontLG,
                Foreground = Brushes.OrangeRed,
                Margin = new Thickness(16),
            };
            return;
        }

        // DB 저장
        string statementNo = await Task.Run(() =>
        {
            ETA.Services.SERVICE1.QuotationService.EnsureTradeStatementTable();
            string no = $"MS-{DateTime.Today:yyyyMMdd}-{DateTime.Now:HHmmss}";
            var quotNos     = captured.Select(i => i.견적번호);
            var abbr        = captured.First().약칭;
            var itemDataDict = ETA.Services.SERVICE1.QuotationService.BuildTradeStatementItemData(captured);
            ETA.Services.SERVICE1.QuotationService.InsertTradeStatement(
                captured.First().업체명, abbr, no, quotNos, supply, vat, total, itemDataDict);
            ETA.Services.SERVICE1.QuotationService.SetTradeStatementNo(
                captured.Select(i => i.Id).ToList(), no);
            return no;
        });

        // 트리 아이콘 갱신
        _issuingHistoryPanel?.RefreshIssueIcons(captured.Select(i => i.Id).ToHashSet());

        Show2.Content = new TextBlock
        {
            Text = $"발행 완료\n거래명세서번호: {statementNo}\n공급가액: {supply:N0}\n부가세: {vat:N0}\n합계: {total:N0}",
            FontFamily = _tradeFont,
            FontSize = AppTheme.FontLG,
            Foreground = AppTheme.FgSuccess,
            Margin = new Thickness(16),
        };
    }

    // =========================================================================
    // 거래명세서 편집 패널 (BT6)
    // =========================================================================
    private static readonly FontFamily _tradeFont =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private void ShowTradeStatementEditor()
    {
        var selected = _quotationHistoryPanel?.GetSelectedIssues()
                       ?? new List<ETA.Models.QuotationIssue>();

        if (selected.Count == 0)
        {
            Show4.Content = new TextBlock
            {
                Text = "⚠️  Show1에서 Ctrl+클릭으로\n견적서를 먼저 선택하세요.",
                FontSize = AppTheme.FontMD, Foreground = Brushes.DarkOrange, FontFamily = _tradeFont,
                Margin = new Thickness(12), TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            return;
        }

        // ── 체크박스별 상태 ──────────────────────────────────────────────────
        var checkBoxes = new List<(ETA.Models.QuotationIssue Issue, CheckBox CB)>();

        // ── Show2 미리보기 갱신 함수 ─────────────────────────────────────────
        void RefreshPreview()
        {
            var checked_ = checkBoxes
                .Where(x => x.CB.IsChecked == true)
                .Select(x => x.Issue)
                .ToList();

            if (checked_.Count == 0)
            {
                Show2.Content = new TextBlock
                {
                    Text = "체크된 항목이 없습니다.",
                    FontSize = AppTheme.FontMD, Foreground = Res("FgMuted"), FontFamily = _tradeFont,
                    Margin = new Thickness(12),
                };
                return;
            }

            decimal supplyTotal = checked_.Sum(i => i.총금액);
            decimal vat         = Math.Round(supplyTotal * 0.1m, 0);
            decimal grand       = supplyTotal + vat;
            string  company     = checked_.First().업체명;

            var preview = new StackPanel { Spacing = 6, Margin = new Thickness(12) };

            preview.Children.Add(new TextBlock
            {
                Text = "📄  거래명세서 미리보기",
                FontSize = AppTheme.FontXL, FontFamily = _tradeFont, FontWeight = FontWeight.SemiBold,
                Foreground = Res("AppFg"),
            });
            preview.Children.Add(new Border { Height = 1, Background = Res("InputBorder"), Margin = new Thickness(0,2,0,4) });
            preview.Children.Add(new TextBlock { Text = $"공급받는자 : {company}", FontSize = AppTheme.FontBase, FontFamily = _tradeFont, Foreground = Res("FgMuted") });
            preview.Children.Add(new TextBlock { Text = $"발행일     : {DateTime.Today:yyyy-MM-dd}", FontSize = AppTheme.FontBase, FontFamily = _tradeFont, Foreground = Res("FgMuted") });
            preview.Children.Add(new TextBlock { Text = $"건수       : {checked_.Count}건", FontSize = AppTheme.FontBase, FontFamily = _tradeFont, Foreground = Res("FgMuted") });

            // 견적번호 목록
            var noList = new TextBlock
            {
                Text = "견적번호 : " + string.Join(", ", checked_.Select(i => i.견적번호)),
                FontSize = AppTheme.FontSM, FontFamily = _tradeFont,
                Foreground = Res("FgMuted"),
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            };
            preview.Children.Add(noList);

            preview.Children.Add(new Border { Height = 1, Background = Res("InputBorder"), Margin = new Thickness(0,6,0,4) });

            // 금액 행
            void AddAmtRow(string label, decimal amt, bool bold = false)
            {
                var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
                row.Children.Add(new TextBlock { Text = label, Width = 90, FontSize = AppTheme.FontMD, FontFamily = _tradeFont, Foreground = Res("FgMuted") });
                var tb = new TextBlock
                {
                    Text = amt.ToString("N0") + " 원",
                    FontSize = AppTheme.FontLG, FontFamily = _tradeFont,
                    Foreground = bold ? Res("AppFg") : Res("FgMuted"),
                    FontWeight = bold ? FontWeight.SemiBold : FontWeight.Normal,
                };
                row.Children.Add(tb);
                preview.Children.Add(row);
            }

            AddAmtRow("공급가액", supplyTotal);
            AddAmtRow("부가세(10%)", vat);
            AddAmtRow("합  계", grand, bold: true);

            Show2.Content = new ScrollViewer { Content = preview, VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto };
        }

        // ── Show4 패널 구성 ──────────────────────────────────────────────────
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*,Auto") };

        // 헤더
        root.Children.Add(new Border
        {
            Background = Res("PanelBg"),
            Padding = new Thickness(10, 8),
            [Grid.RowProperty] = 0,
            Child = new TextBlock
            {
                Text = "📋  거래명세서 항목 선택",
                FontSize = AppTheme.FontLG, FontFamily = _tradeFont, FontWeight = FontWeight.SemiBold,
                Foreground = Res("AppFg"),
            }
        });

        // 체크박스 리스트
        var listPanel = new StackPanel { Spacing = 4, Margin = new Thickness(8, 6) };

        foreach (var issue in selected)
        {
            var cb = new CheckBox
            {
                IsChecked = true,
                Margin = new Thickness(0, 2),
                Content = new StackPanel
                {
                    Orientation = Orientation.Vertical, Spacing = 2,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = $"{issue.약칭}  {issue.시료명}",
                            FontSize = AppTheme.FontBase, FontFamily = _tradeFont, Foreground = Res("AppFg"),
                            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                        },
                        new TextBlock
                        {
                            Text = $"{issue.견적번호}  |  {issue.총금액:N0} 원",
                            FontSize = AppTheme.FontSM, FontFamily = _tradeFont,
                            Foreground = Res("FgMuted"),
                        },
                    }
                },
            };
            cb.IsCheckedChanged += (_, _) => RefreshPreview();
            checkBoxes.Add((issue, cb));
            listPanel.Children.Add(cb);
        }

        var scroll = new ScrollViewer
        {
            Content = listPanel,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            [Grid.RowProperty] = 1,
        };
        root.Children.Add(scroll);

        // 발행 버튼
        var btnIssue = new Button
        {
            Content = "🖨  발행 (Excel + DB 저장)",
            Height = 38, Margin = new Thickness(8, 6),
            FontSize = AppTheme.FontMD, FontFamily = _tradeFont,
            Background = Res("SubBtnBg"),
            Foreground = Res("AppFg"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            [Grid.RowProperty] = 2,
        };
        btnIssue.Click += (_, _) => _ = IssueTradeStatementAsync(checkBoxes);
        root.Children.Add(btnIssue);

        Show4.Content = root;
        // Show4 영역 표시
        SetContentLayout(content2Star: 7, content4Star: 3, upperStar: 13, lowerStar: 4);

        // 초기 미리보기
        RefreshPreview();
    }

    private async Task IssueTradeStatementAsync(
        List<(ETA.Models.QuotationIssue Issue, CheckBox CB)> checkBoxes)
    {
        var checkedIssues = checkBoxes
            .Where(x => x.CB.IsChecked == true)
            .Select(x => x.Issue)
            .ToList();

        if (checkedIssues.Count == 0)
        {
            Show2.Content = new TextBlock
            {
                Text = "⚠️  체크된 항목이 없습니다.",
                FontSize = AppTheme.FontLG, Foreground = Brushes.Orange, FontFamily = _tradeFont,
                Margin = new Thickness(12),
            };
            return;
        }

        // 분석 미완료 항목 확인
        var incompleteQuotNos = checkedIssues.Select(i => i.견적번호).Distinct().ToList();
        var incompleteItems = await Task.Run(() =>
            ETA.Services.SERVICE1.AnalysisRecordService.GetIncompleteItems(incompleteQuotNos));
        if (incompleteItems.Count > 0)
        {
            bool proceed = await ShowIncompleteWarningAsync(incompleteItems);
            if (!proceed) return;
        }

        string companyName = checkedIssues.First().업체명;
        string statementNo = $"TS-{companyName[..Math.Min(4, companyName.Length)]}-{DateTime.Today:yyyyMMdd}-{DateTime.Now:HHmm}";

        // 파일 저장 경로 선택
        var dialog = new Avalonia.Platform.Storage.FilePickerSaveOptions
        {
            Title             = "거래명세서 Excel 저장",
            SuggestedFileName = $"거래명세서_{companyName}_{DateTime.Today:yyyyMMdd}.xlsx",
            FileTypeChoices   = new[]
            {
                new Avalonia.Platform.Storage.FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx" } }
            }
        };
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;
        var file = await topLevel.StorageProvider.SaveFilePickerAsync(dialog);
        if (file == null) return;

        var savePath = file.Path.LocalPath;

        Show2.Content = new TextBlock
        {
            Text = $"⏳  발행 중... ({checkedIssues.Count}건)",
            FontSize = AppTheme.FontLG, Foreground = Res("FgMuted"), FontFamily = _tradeFont,
            Margin = new Thickness(12),
        };

        var captured = checkedIssues.ToList();
        var (ok, msg, supply, vat, total) = await Task.Run(
            () => ETA.Services.SERVICE1.QuotationService.ExportTradingStatement(captured, savePath));

        if (ok)
        {
            // 거래명세서발행내역 DB 저장
            var quotNos      = captured.Select(i => i.견적번호);
            var abbr         = captured.First().약칭;
            var itemDataDict = ETA.Services.SERVICE1.QuotationService.BuildTradeStatementItemData(captured);
            await Task.Run(() =>
            {
                ETA.Services.SERVICE1.QuotationService.EnsureTradeStatementTable();
                ETA.Services.SERVICE1.QuotationService.InsertTradeStatement(
                    companyName, abbr, statementNo, quotNos, supply, vat, total, itemDataDict);
                ETA.Services.SERVICE1.QuotationService.SetTradeStatementNo(
                    captured.Select(i => i.Id), statementNo);
            });

            // 트리 아이콘 갱신
            _quotationHistoryPanel?.RefreshIssueIcons(captured.Select(i => i.Id).ToHashSet());

            Show2.Content = new TextBlock
            {
                Text = $"✅  발행 완료!\n거래명세서번호: {statementNo}\n{checkedIssues.Count}건  합계: {total:N0} 원\n\n{savePath}",
                FontSize = AppTheme.FontLG, Foreground = AppTheme.FgSuccess, FontFamily = _tradeFont,
                Margin = new Thickness(12), TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            };
            try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true }); }
            catch { }
        }
        else
        {
            Show2.Content = new TextBlock
            {
                Text = $"❌  {msg}",
                FontSize = AppTheme.FontLG, Foreground = Brushes.OrangeRed, FontFamily = _tradeFont,
                Margin = new Thickness(12),
            };
        }
    }

    private void BT7_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":
                if (_agentTreePage != null)
                    Show2.Content = _agentTreePage.BuildAssignmentChart();
                break;
            case "Purchase":
                _purchasePage?.ShowSettings(this);
                break;
            case "ResultSubmitMeasure":
                new MeasurerLoginWindow().Show(this);
                break;
            case "AnalysisPlan": SetAnalysisPlanDay(5); break; // 토
            default:
                break;
        }
    }

    private void BT8_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "ResultSubmitMeasure":
                new DataToMeasurerWindow().Show(this);
                break;
            case "AnalysisPlan": SetAnalysisPlanDay(6); break; // 일
            default:
                break;
        }
    }

    private void BT9_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "AnalysisPlan": SetAnalysisPlanDay(-2); break; // BASE
            default:
                break;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  테마 (DynamicResource 전체 교체 방식)
    // ══════════════════════════════════════════════════════════════════════

    private void OnThemeSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (cboTheme == null) return;
        ApplyTheme(cboTheme.SelectedIndex);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  글자 크기 슬라이더
    // ══════════════════════════════════════════════════════════════════════

    private void SldFontSize_ValueChanged(object? sender,
        Avalonia.Controls.Primitives.RangeBaseValueChangedEventArgs e)
    {
        var scale = e.NewValue;
        ApplyFontScale(scale);
        if (txbFontScale != null)
            txbFontScale.Text = $"{(int)(scale * 100)}%";
        SaveFontScale(scale);
    }

    /// <summary>8개 폰트 크기 키를 scale 배율로 일괄 갱신</summary>
    private void ApplyFontScale(double scale)
    {
        var sizes = new (string Key, double Base)[]
        {
            ("FontSizeXS",     9.0),
            ("FontSizeSM",    10.0),
            ("FontSizeBase",  11.0),
            ("FontSizeMD",    12.0),
            ("FontSizeLG",    13.0),
            ("FontSizeXL",    14.0),
            ("FontSizeTitle", 22.0),
            ("FontSizeHuge",  28.0),
        };
        foreach (var (key, baseSize) in sizes)
        {
            double newVal = Math.Round(baseSize * scale, 1);
            this.Resources[key] = newVal;
            // Application.Current.Resources 에도 동기화 → 정적 코드비하인드 BindFs() 지원
            if (Application.Current != null)
                Application.Current.Resources[key] = newVal;
        }
    }

    // ── 글자 크기 설정 저장/복원 (AppData/ETA/Users/{id}/ui_settings.json) ──

    private string FontScaleSettingsPath =>
        System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ETA", "Users", CurrentEmployeeId, "ui_settings.json");

    private void SaveFontScale(double scale)
    {
        try
        {
            var dir = System.IO.Path.GetDirectoryName(FontScaleSettingsPath)!;
            System.IO.Directory.CreateDirectory(dir);
            System.IO.File.WriteAllText(FontScaleSettingsPath,
                $"{{\"fontScale\":{scale.ToString(System.Globalization.CultureInfo.InvariantCulture)}}}");
        }
        catch (Exception ex) { }
    }

    private double LoadFontScale()
    {
        try
        {
            if (!System.IO.File.Exists(FontScaleSettingsPath)) return 1.0;
            var json = System.IO.File.ReadAllText(FontScaleSettingsPath);
            using var doc = System.Text.Json.JsonDocument.Parse(json);
            if (doc.RootElement.TryGetProperty("fontScale", out var el))
                return el.GetDouble();
        }
        catch (Exception ex) { }
        return 1.0;
    }

    private void ApplyTheme(int themeIndex)
    {
        // ── Avalonia 기본 테마 변형 ──
        var isLight = themeIndex == 1; // Soft Light만 Light variant
        var variant = isLight ? ThemeVariant.Light : ThemeVariant.Dark;
        if (Application.Current is Application app)
            app.RequestedThemeVariant = variant;
        this.RequestedThemeVariant = variant;

        // ── 팔레트 선택 ──
        var palette = themeIndex switch
        {
            1 => LightPalette(),
            2 => OceanPalette(),
            3 => ForestPalette(),
            4 => DeadpoolPalette(),
            5 => BumblebeePalette(),
            6 => HulkPalette(),
            7 => IronManPalette(),
            _ => DarkPalette(),
        };

        // ── Window.Resources + Application.Current.Resources 동시 교체 ──
        foreach (var (key, color) in palette)
        {
            var brush = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse(color));
            this.Resources[key] = brush;
            if (Application.Current != null)
                Application.Current.Resources[key] = brush;
        }

        // ── 모든 테마에서 shimmer 활성화 ──
        Dispatcher.UIThread.Post(() => TextShimmer.AttachAll(this), DispatcherPriority.Background);
    }

    /// <summary>현재 테마 브러시를 Window.Resources 에서 읽는 헬퍼</summary>
    private Brush Res(string key, string fallback = "#888888")
    {
        if (this.Resources.TryGetResource(key, null, out var v) && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 다크 팔레트 ──────────────────────────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> DarkPalette() => new()
    {
        // ── 레이아웃 ──
        ["AppBg"]            = "#1e1e26",
        ["PanelBg"]          = "#39383f",
        ["PanelInnerBg"]     = "#2d2d35",
        ["MenuBarBg"]        = "#444444",
        ["MenuItemBg"]       = "#444444",
        ["MenuItemHover"]    = "#3a3a5a",
        ["SubMenuBg"]        = "#2a2a3a",
        ["SubMenuItemBg"]    = "#35343c",
        ["SubMenuItemHover"] = "#3a3a5a",
        ["SubBtnBg"]         = "#342f2f",
        ["SubBtnHover"]      = "#3a3a5a",
        ["SplitterColor"]    = "#020202",
        ["AppFg"]            = "#d8d8e0",
        ["FgMuted"]          = "#999aaa",
        ["FgHover"]          = "#88aaff",
        ["TreeFg"]           = "#d0d0d8",
        ["TreeSelBg"]        = "#1e3a5a",
        ["TreeSelFg"]        = "#e8e8f0",
        ["GridHeaderBg"]     = "#2a2a32",
        ["GridRowBg"]        = "#35343c",
        ["GridRowAltBg"]     = "#2d2d35",
        ["InputBg"]          = "#2d2d35",
        ["InputBorder"]      = "#555566",
        ["InputFg"]          = "#e8e8e8",
        ["TreeHoverBg"]      = "#252535",
        ["Panel4Bg"]         = "#252535",
        ["TabActiveBg"]      = "#1e3a5a",
        ["TabActiveFg"]      = "#88aaee",
        // ── 버튼 ──
        ["BtnBg"]            = "#2a2a3a",
        ["BtnFg"]            = "#cccccc",
        ["BtnBorder"]        = "#404050",
        ["BtnHoverBg"]       = "#3a3a5a",
        ["BtnPrimaryBg"]     = "#1e3a5a",
        ["BtnPrimaryFg"]     = "#88aaee",
        ["BtnPrimaryBorder"] = "#336699",
        ["BtnPrimaryHover"]  = "#2a4a6a",
        ["BtnDangerBg"]      = "#4a2a2a",
        ["BtnDangerFg"]      = "#ff8888",
        ["BtnDangerBorder"]  = "#663333",
        ["BtnDangerHover"]   = "#5a3333",
        ["BtnSuccessBg"]     = "#1a3a2a",
        ["BtnSuccessFg"]     = "#88cc88",
        ["BtnSuccessBorder"] = "#2a6a2a",
        ["BtnSuccessHover"]  = "#2a4a3a",
        // ── Theme* (AppTheme 헬퍼 + DynamicResource 연동) ──
        ["ThemeBorderDefault"]   = "#555577",
        ["ThemeBorderSubtle"]    = "#333344",
        ["ThemeBorderMuted"]     = "#404050",
        ["ThemeBorderActive"]    = "#2a6a2a",
        ["ThemeBorderInfo"]      = "#336699",
        ["ThemeBorderAccent"]    = "#3a5a8a",
        ["ThemeBorderDanger"]    = "#663333",
        ["ThemeBorderWarn"]      = "#6a5a3a",
        ["ThemeBorderSeparator"] = "#3a3a4a",
        ["ThemeBorderPanel"]     = "#3a3a4a",
        ["ThemeFgPrimary"]       = "#ffffff",
        ["ThemeFgSecondary"]     = "#cccccc",
        ["ThemeFgMuted"]         = "#888899",
        ["ThemeFgDimmed"]        = "#555566",
        ["ThemeFgSuccess"]       = "#88cc88",
        ["ThemeFgInfo"]          = "#88aacc",
        ["ThemeFgDanger"]        = "#ff6666",
        ["ThemeFgWarn"]          = "#ffaa44",
        ["ThemeFgLink"]          = "#88aaff",
        ["ThemeBgPrimary"]       = "#1e1e26",
        ["ThemeBgSecondary"]     = "#252535",
        ["ThemeBgCard"]          = "#2a2a3a",
        ["ThemeBgInput"]         = "#22223a",
        ["ThemeBgActiveGreen"]   = "#1a3a1a",
        ["ThemeBgActiveBlue"]    = "#1e3a5a",
        ["ThemeBgDanger"]        = "#4a2a2a",
    };

    // ── 라이트 팔레트 ─────────────────────────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> LightPalette() => new()
    {
        // ── 레이아웃 (눈부시지 않은 소프트 라이트) ──
        ["AppBg"]            = "#eef1f6",
        ["PanelBg"]          = "#f5f7fa",
        ["PanelInnerBg"]     = "#edf0f5",
        ["MenuBarBg"]        = "#334155",
        ["MenuItemBg"]       = "#334155",
        ["MenuItemHover"]    = "#1e293b",
        ["SubMenuBg"]        = "#1e293b",
        ["SubMenuItemBg"]    = "#334155",
        ["SubMenuItemHover"] = "#475569",
        ["SubBtnBg"]         = "#dde3ec",
        ["SubBtnHover"]      = "#c8d0dc",
        ["SplitterColor"]    = "#c8d0dc",
        ["AppFg"]            = "#1e293b",
        ["FgMuted"]          = "#64748b",
        ["FgHover"]          = "#2563eb",
        ["TreeFg"]           = "#1e293b",
        ["TreeSelBg"]        = "#dbeafe",
        ["TreeSelFg"]        = "#1e40af",
        ["GridHeaderBg"]     = "#dde3ec",
        ["GridRowBg"]        = "#f5f7fa",
        ["GridRowAltBg"]     = "#edf0f5",
        ["InputBg"]          = "#f5f7fa",
        ["InputBorder"]      = "#94a3b8",
        ["InputFg"]          = "#1e293b",
        ["TreeHoverBg"]      = "#dde3ec",
        ["Panel4Bg"]         = "#edf0f5",
        ["TabActiveBg"]      = "#dbeafe",
        ["TabActiveFg"]      = "#1e40af",
        // ── 버튼 (라이트) ──
        ["BtnBg"]            = "#e4e8ef",
        ["BtnFg"]            = "#334155",
        ["BtnBorder"]        = "#c4cdd8",
        ["BtnHoverBg"]       = "#d1d9e4",
        ["BtnPrimaryBg"]     = "#2563eb",
        ["BtnPrimaryFg"]     = "#ffffff",
        ["BtnPrimaryBorder"] = "#1e40af",
        ["BtnPrimaryHover"]  = "#1d4ed8",
        ["BtnDangerBg"]      = "#fef2f2",
        ["BtnDangerFg"]      = "#dc2626",
        ["BtnDangerBorder"]  = "#fca5a5",
        ["BtnDangerHover"]   = "#fee2e2",
        ["BtnSuccessBg"]     = "#f0fdf4",
        ["BtnSuccessFg"]     = "#16a34a",
        ["BtnSuccessBorder"] = "#86efac",
        ["BtnSuccessHover"]  = "#dcfce7",
        // ── Theme* (라이트 모드 — 소프트 톤) ──
        ["ThemeBorderDefault"]   = "#b8c4d4",
        ["ThemeBorderSubtle"]    = "#d1d9e4",
        ["ThemeBorderMuted"]     = "#c4cdd8",
        ["ThemeBorderActive"]    = "#3b82f6",
        ["ThemeBorderInfo"]      = "#3b82f6",
        ["ThemeBorderAccent"]    = "#6366f1",
        ["ThemeBorderDanger"]    = "#ef4444",
        ["ThemeBorderWarn"]      = "#f59e0b",
        ["ThemeBorderSeparator"] = "#d1d9e4",
        ["ThemeBorderPanel"]     = "#c4cdd8",
        ["ThemeFgPrimary"]       = "#1e293b",
        ["ThemeFgSecondary"]     = "#475569",
        ["ThemeFgMuted"]         = "#64748b",
        ["ThemeFgDimmed"]        = "#94a3b8",
        ["ThemeFgSuccess"]       = "#16a34a",
        ["ThemeFgInfo"]          = "#2563eb",
        ["ThemeFgDanger"]        = "#dc2626",
        ["ThemeFgWarn"]          = "#d97706",
        ["ThemeFgLink"]          = "#2563eb",
        ["ThemeBgPrimary"]       = "#eef1f6",
        ["ThemeBgSecondary"]     = "#e4e8ef",
        ["ThemeBgCard"]          = "#f5f7fa",
        ["ThemeBgInput"]         = "#f5f7fa",
        ["ThemeBgActiveGreen"]   = "#dcfce7",
        ["ThemeBgActiveBlue"]    = "#dbeafe",
        ["ThemeBgDanger"]        = "#fef2f2",
    };

    // ── Ocean 팔레트 (딥 블루 다크) ──────────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> OceanPalette() => new()
    {
        ["AppBg"]            = "#0f172a",
        ["PanelBg"]          = "#1e293b",
        ["PanelInnerBg"]     = "#162032",
        ["MenuBarBg"]        = "#1e3a5a",
        ["MenuItemBg"]       = "#1e3a5a",
        ["MenuItemHover"]    = "#2a4a6a",
        ["SubMenuBg"]        = "#162032",
        ["SubMenuItemBg"]    = "#1e293b",
        ["SubMenuItemHover"] = "#2a4a6a",
        ["SubBtnBg"]         = "#1e293b",
        ["SubBtnHover"]      = "#2a4a6a",
        ["SplitterColor"]    = "#0c1220",
        ["AppFg"]            = "#d0d8e4",
        ["FgMuted"]          = "#8a99b0",
        ["FgHover"]          = "#60a5fa",
        ["TreeFg"]           = "#bcc8d8",
        ["TreeSelBg"]        = "#1e40af",
        ["TreeSelFg"]        = "#dde4f0",
        ["GridHeaderBg"]     = "#1a2744",
        ["GridRowBg"]        = "#1e293b",
        ["GridRowAltBg"]     = "#162032",
        ["InputBg"]          = "#162032",
        ["InputBorder"]      = "#334155",
        ["InputFg"]          = "#e2e8f0",
        ["TreeHoverBg"]      = "#1a2744",
        ["Panel4Bg"]         = "#162032",
        ["TabActiveBg"]      = "#1e40af",
        ["TabActiveFg"]      = "#93c5fd",
        ["BtnBg"]            = "#1e293b",
        ["BtnFg"]            = "#cbd5e1",
        ["BtnBorder"]        = "#334155",
        ["BtnHoverBg"]       = "#2a4a6a",
        ["BtnPrimaryBg"]     = "#1e40af",
        ["BtnPrimaryFg"]     = "#93c5fd",
        ["BtnPrimaryBorder"] = "#3b82f6",
        ["BtnPrimaryHover"]  = "#2563eb",
        ["BtnDangerBg"]      = "#450a0a",
        ["BtnDangerFg"]      = "#f87171",
        ["BtnDangerBorder"]  = "#7f1d1d",
        ["BtnDangerHover"]   = "#991b1b",
        ["BtnSuccessBg"]     = "#052e16",
        ["BtnSuccessFg"]     = "#4ade80",
        ["BtnSuccessBorder"] = "#166534",
        ["BtnSuccessHover"]  = "#14532d",
        ["ThemeBorderDefault"]   = "#334155",
        ["ThemeBorderSubtle"]    = "#1e293b",
        ["ThemeBorderMuted"]     = "#2a3a5a",
        ["ThemeBorderActive"]    = "#3b82f6",
        ["ThemeBorderInfo"]      = "#3b82f6",
        ["ThemeBorderAccent"]    = "#6366f1",
        ["ThemeBorderDanger"]    = "#7f1d1d",
        ["ThemeBorderWarn"]      = "#78350f",
        ["ThemeBorderSeparator"] = "#1e293b",
        ["ThemeBorderPanel"]     = "#334155",
        ["ThemeFgPrimary"]       = "#e2e8f0",
        ["ThemeFgSecondary"]     = "#cbd5e1",
        ["ThemeFgMuted"]         = "#94a3b8",
        ["ThemeFgDimmed"]        = "#475569",
        ["ThemeFgSuccess"]       = "#4ade80",
        ["ThemeFgInfo"]          = "#60a5fa",
        ["ThemeFgDanger"]        = "#f87171",
        ["ThemeFgWarn"]          = "#fbbf24",
        ["ThemeFgLink"]          = "#93c5fd",
        ["ThemeBgPrimary"]       = "#0f172a",
        ["ThemeBgSecondary"]     = "#162032",
        ["ThemeBgCard"]          = "#1e293b",
        ["ThemeBgInput"]         = "#162032",
        ["ThemeBgActiveGreen"]   = "#052e16",
        ["ThemeBgActiveBlue"]    = "#172554",
        ["ThemeBgDanger"]        = "#450a0a",
    };

    // ── Forest 팔레트 (따뜻한 다크 그린) ──────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> ForestPalette() => new()
    {
        ["AppBg"]            = "#1a1e1a",
        ["PanelBg"]          = "#2a302a",
        ["PanelInnerBg"]     = "#222822",
        ["MenuBarBg"]        = "#2d3a2d",
        ["MenuItemBg"]       = "#2d3a2d",
        ["MenuItemHover"]    = "#3a4a3a",
        ["SubMenuBg"]        = "#222822",
        ["SubMenuItemBg"]    = "#2a302a",
        ["SubMenuItemHover"] = "#3a4a3a",
        ["SubBtnBg"]         = "#2a302a",
        ["SubBtnHover"]      = "#3a4a3a",
        ["SplitterColor"]    = "#141814",
        ["AppFg"]            = "#d0d8d0",
        ["FgMuted"]          = "#7a8a7a",
        ["FgHover"]          = "#86efac",
        ["TreeFg"]           = "#b8c8b8",
        ["TreeSelBg"]        = "#1a5a2a",
        ["TreeSelFg"]        = "#dde8dd",
        ["GridHeaderBg"]     = "#243024",
        ["GridRowBg"]        = "#2a302a",
        ["GridRowAltBg"]     = "#222822",
        ["InputBg"]          = "#222822",
        ["InputBorder"]      = "#3a4a3a",
        ["InputFg"]          = "#e0e8e0",
        ["TreeHoverBg"]      = "#243024",
        ["Panel4Bg"]         = "#222822",
        ["TabActiveBg"]      = "#1a5a2a",
        ["TabActiveFg"]      = "#86efac",
        ["BtnBg"]            = "#2a302a",
        ["BtnFg"]            = "#b4c8b4",
        ["BtnBorder"]        = "#3a4a3a",
        ["BtnHoverBg"]       = "#3a4a3a",
        ["BtnPrimaryBg"]     = "#1a5a2a",
        ["BtnPrimaryFg"]     = "#86efac",
        ["BtnPrimaryBorder"] = "#22c55e",
        ["BtnPrimaryHover"]  = "#166534",
        ["BtnDangerBg"]      = "#3a1a1a",
        ["BtnDangerFg"]      = "#f87171",
        ["BtnDangerBorder"]  = "#663333",
        ["BtnDangerHover"]   = "#5a2a2a",
        ["BtnSuccessBg"]     = "#0f3a1a",
        ["BtnSuccessFg"]     = "#86efac",
        ["BtnSuccessBorder"] = "#22c55e",
        ["BtnSuccessHover"]  = "#166534",
        ["ThemeBorderDefault"]   = "#3a4a3a",
        ["ThemeBorderSubtle"]    = "#2a342a",
        ["ThemeBorderMuted"]     = "#344434",
        ["ThemeBorderActive"]    = "#22c55e",
        ["ThemeBorderInfo"]      = "#3b82f6",
        ["ThemeBorderAccent"]    = "#4ade80",
        ["ThemeBorderDanger"]    = "#663333",
        ["ThemeBorderWarn"]      = "#6a5a3a",
        ["ThemeBorderSeparator"] = "#2a342a",
        ["ThemeBorderPanel"]     = "#3a4a3a",
        ["ThemeFgPrimary"]       = "#e0e8e0",
        ["ThemeFgSecondary"]     = "#b4c8b4",
        ["ThemeFgMuted"]         = "#8a9a8a",
        ["ThemeFgDimmed"]        = "#5a6a5a",
        ["ThemeFgSuccess"]       = "#86efac",
        ["ThemeFgInfo"]          = "#60a5fa",
        ["ThemeFgDanger"]        = "#f87171",
        ["ThemeFgWarn"]          = "#fbbf24",
        ["ThemeFgLink"]          = "#86efac",
        ["ThemeBgPrimary"]       = "#1a1e1a",
        ["ThemeBgSecondary"]     = "#222822",
        ["ThemeBgCard"]          = "#2a302a",
        ["ThemeBgInput"]         = "#222822",
        ["ThemeBgActiveGreen"]   = "#0f3a1a",
        ["ThemeBgActiveBlue"]    = "#1a2a3a",
        ["ThemeBgDanger"]        = "#3a1a1a",
    };

    // ── 데드풀 팔레트 (마룬/크림슨 + 제트블랙, 고대비) ──────────────────
    private static System.Collections.Generic.Dictionary<string, string> DeadpoolPalette() => new()
    {
        ["AppBg"]            = "#0e0606",
        ["PanelBg"]          = "#1c0c0c",
        ["PanelInnerBg"]     = "#140808",
        ["MenuBarBg"]        = "#2a0e0e",
        ["MenuItemBg"]       = "#2a0e0e",
        ["MenuItemHover"]    = "#4a1616",
        ["SubMenuBg"]        = "#140808",
        ["SubMenuItemBg"]    = "#1c0c0c",
        ["SubMenuItemHover"] = "#4a1616",
        ["SubBtnBg"]         = "#1c0c0c",
        ["SubBtnHover"]      = "#4a1616",
        ["SplitterColor"]    = "#040101",
        ["AppFg"]            = "#e0c8c8",
        ["FgMuted"]          = "#8a5555",
        ["FgHover"]          = "#dd3333",
        ["TreeFg"]           = "#ccaaaa",
        ["TreeSelBg"]        = "#6a1818",
        ["TreeSelFg"]        = "#ffd0d0",
        ["GridHeaderBg"]     = "#200a0a",
        ["GridRowBg"]        = "#1c0c0c",
        ["GridRowAltBg"]     = "#140808",
        ["InputBg"]          = "#140808",
        ["InputBorder"]      = "#4a2020",
        ["InputFg"]          = "#eed8d8",
        ["TreeHoverBg"]      = "#2a0e0e",
        ["Panel4Bg"]         = "#140808",
        ["TabActiveBg"]      = "#6a1818",
        ["TabActiveFg"]      = "#ff4040",
        ["BtnBg"]            = "#1c0c0c",
        ["BtnFg"]            = "#bb8888",
        ["BtnBorder"]        = "#4a2020",
        ["BtnHoverBg"]       = "#4a1616",
        ["BtnPrimaryBg"]     = "#6a1818",
        ["BtnPrimaryFg"]     = "#ff4040",
        ["BtnPrimaryBorder"] = "#bb1111",
        ["BtnPrimaryHover"]  = "#7a2222",
        ["BtnDangerBg"]      = "#500808",
        ["BtnDangerFg"]      = "#ff3333",
        ["BtnDangerBorder"]  = "#991111",
        ["BtnDangerHover"]   = "#601010",
        ["BtnSuccessBg"]     = "#1c0c0c",
        ["BtnSuccessFg"]     = "#ee7777",
        ["BtnSuccessBorder"] = "#773333",
        ["BtnSuccessHover"]  = "#2a1414",
        ["ThemeBorderDefault"]   = "#4a2020",
        ["ThemeBorderSubtle"]    = "#2a1010",
        ["ThemeBorderMuted"]     = "#381818",
        ["ThemeBorderActive"]    = "#bb1111",
        ["ThemeBorderInfo"]      = "#aa2222",
        ["ThemeBorderAccent"]    = "#881111",
        ["ThemeBorderDanger"]    = "#991111",
        ["ThemeBorderWarn"]      = "#774422",
        ["ThemeBorderSeparator"] = "#2a0e0e",
        ["ThemeBorderPanel"]     = "#381414",
        ["ThemeFgPrimary"]       = "#eed8d8",
        ["ThemeFgSecondary"]     = "#bb8888",
        ["ThemeFgMuted"]         = "#8a5555",
        ["ThemeFgDimmed"]        = "#553030",
        ["ThemeFgSuccess"]       = "#ee7777",
        ["ThemeFgInfo"]          = "#dd4444",
        ["ThemeFgDanger"]        = "#ff3333",
        ["ThemeFgWarn"]          = "#ee8844",
        ["ThemeFgLink"]          = "#ff4040",
        ["ThemeBgPrimary"]       = "#0e0606",
        ["ThemeBgSecondary"]     = "#140808",
        ["ThemeBgCard"]          = "#1c0c0c",
        ["ThemeBgInput"]         = "#120707",
        ["ThemeBgActiveGreen"]   = "#1c0c0c",
        ["ThemeBgActiveBlue"]    = "#2a1010",
        ["ThemeBgDanger"]        = "#500808",
    };

    // ── 범블비 팔레트 (오렌지 옐로우 + 건메탈/차콜 + 시안 포인트) ────────
    private static System.Collections.Generic.Dictionary<string, string> BumblebeePalette() => new()
    {
        ["AppBg"]            = "#0e0c02",
        ["PanelBg"]          = "#1c1804",
        ["PanelInnerBg"]     = "#141202",
        ["MenuBarBg"]        = "#2c2606",
        ["MenuItemBg"]       = "#2c2606",
        ["MenuItemHover"]    = "#443808",
        ["SubMenuBg"]        = "#141202",
        ["SubMenuItemBg"]    = "#1c1804",
        ["SubMenuItemHover"] = "#443808",
        ["SubBtnBg"]         = "#1c1804",
        ["SubBtnHover"]      = "#443808",
        ["SplitterColor"]    = "#060500",
        ["AppFg"]            = "#f5eeb0",
        ["FgMuted"]          = "#a09020",
        ["FgHover"]          = "#ffe600",
        ["TreeFg"]           = "#eee080",
        ["TreeSelBg"]        = "#665500",
        ["TreeSelFg"]        = "#ffee55",
        ["GridHeaderBg"]     = "#1e1c04",
        ["GridRowBg"]        = "#1c1804",
        ["GridRowAltBg"]     = "#141202",
        ["InputBg"]          = "#141202",
        ["InputBorder"]      = "#554a08",
        ["InputFg"]          = "#fff8cc",
        ["TreeHoverBg"]      = "#2c2606",
        ["Panel4Bg"]         = "#141202",
        ["TabActiveBg"]      = "#665500",
        ["TabActiveFg"]      = "#ffdd00",
        ["BtnBg"]            = "#1c1804",
        ["BtnFg"]            = "#ddd050",
        ["BtnBorder"]        = "#554a08",
        ["BtnHoverBg"]       = "#443808",
        ["BtnPrimaryBg"]     = "#665500",
        ["BtnPrimaryFg"]     = "#ffdd00",
        ["BtnPrimaryBorder"] = "#ddbb00",
        ["BtnPrimaryHover"]  = "#776608",
        ["BtnDangerBg"]      = "#3a1a04",
        ["BtnDangerFg"]      = "#ff6622",
        ["BtnDangerBorder"]  = "#994410",
        ["BtnDangerHover"]   = "#4a2408",
        ["BtnSuccessBg"]     = "#1a2808",
        ["BtnSuccessFg"]     = "#aadd22",
        ["BtnSuccessBorder"] = "#668c11",
        ["BtnSuccessHover"]  = "#223410",
        ["ThemeBorderDefault"]   = "#554a08",
        ["ThemeBorderSubtle"]    = "#332e04",
        ["ThemeBorderMuted"]     = "#443a06",
        ["ThemeBorderActive"]    = "#ddbb00",
        ["ThemeBorderInfo"]      = "#44ccee",
        ["ThemeBorderAccent"]    = "#ccaa00",
        ["ThemeBorderDanger"]    = "#994410",
        ["ThemeBorderWarn"]      = "#aa8800",
        ["ThemeBorderSeparator"] = "#2c2606",
        ["ThemeBorderPanel"]     = "#443a06",
        ["ThemeFgPrimary"]       = "#fff8cc",
        ["ThemeFgSecondary"]     = "#ddd050",
        ["ThemeFgMuted"]         = "#a09020",
        ["ThemeFgDimmed"]        = "#665820",
        ["ThemeFgSuccess"]       = "#aadd22",
        ["ThemeFgInfo"]          = "#44ccee",
        ["ThemeFgDanger"]        = "#ff6622",
        ["ThemeFgWarn"]          = "#ffbb00",
        ["ThemeFgLink"]          = "#ffe600",
        ["ThemeBgPrimary"]       = "#0e0c02",
        ["ThemeBgSecondary"]     = "#141202",
        ["ThemeBgCard"]          = "#1c1804",
        ["ThemeBgInput"]         = "#121002",
        ["ThemeBgActiveGreen"]   = "#1a2808",
        ["ThemeBgActiveBlue"]    = "#141e2c",
        ["ThemeBgDanger"]        = "#3a1a04",
    };

    // ── 헐크 팔레트 (강렬한 녹색 + 퍼플 포인트) ────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> HulkPalette() => new()
    {
        ["AppBg"]            = "#040e04",
        ["PanelBg"]          = "#0c1c0c",
        ["PanelInnerBg"]     = "#081408",
        ["MenuBarBg"]        = "#143014",
        ["MenuItemBg"]       = "#143014",
        ["MenuItemHover"]    = "#1e4a1e",
        ["SubMenuBg"]        = "#081408",
        ["SubMenuItemBg"]    = "#0c1c0c",
        ["SubMenuItemHover"] = "#1e4a1e",
        ["SubBtnBg"]         = "#0c1c0c",
        ["SubBtnHover"]      = "#1e4a1e",
        ["SplitterColor"]    = "#020602",
        ["AppFg"]            = "#c0f0c0",
        ["FgMuted"]          = "#449944",
        ["FgHover"]          = "#44ff44",
        ["TreeFg"]           = "#88dd88",
        ["TreeSelBg"]        = "#116611",
        ["TreeSelFg"]        = "#88ff88",
        ["GridHeaderBg"]     = "#0a220a",
        ["GridRowBg"]        = "#0c1c0c",
        ["GridRowAltBg"]     = "#081408",
        ["InputBg"]          = "#081408",
        ["InputBorder"]      = "#1a551a",
        ["InputFg"]          = "#ccffcc",
        ["TreeHoverBg"]      = "#143014",
        ["Panel4Bg"]         = "#081408",
        ["TabActiveBg"]      = "#116611",
        ["TabActiveFg"]      = "#44ff44",
        ["BtnBg"]            = "#0c1c0c",
        ["BtnFg"]            = "#66cc66",
        ["BtnBorder"]        = "#1a551a",
        ["BtnHoverBg"]       = "#1e4a1e",
        ["BtnPrimaryBg"]     = "#116611",
        ["BtnPrimaryFg"]     = "#44ff44",
        ["BtnPrimaryBorder"] = "#22bb22",
        ["BtnPrimaryHover"]  = "#1a7a1a",
        ["BtnDangerBg"]      = "#280e28",
        ["BtnDangerFg"]      = "#dd55cc",
        ["BtnDangerBorder"]  = "#883388",
        ["BtnDangerHover"]   = "#331633",
        ["BtnSuccessBg"]     = "#083808",
        ["BtnSuccessFg"]     = "#44ee44",
        ["BtnSuccessBorder"] = "#22cc22",
        ["BtnSuccessHover"]  = "#0e4a0e",
        ["ThemeBorderDefault"]   = "#1a551a",
        ["ThemeBorderSubtle"]    = "#0e330e",
        ["ThemeBorderMuted"]     = "#144414",
        ["ThemeBorderActive"]    = "#22bb22",
        ["ThemeBorderInfo"]      = "#8844cc",
        ["ThemeBorderAccent"]    = "#33dd33",
        ["ThemeBorderDanger"]    = "#883388",
        ["ThemeBorderWarn"]      = "#558811",
        ["ThemeBorderSeparator"] = "#143014",
        ["ThemeBorderPanel"]     = "#1e4a1e",
        ["ThemeFgPrimary"]       = "#ccffcc",
        ["ThemeFgSecondary"]     = "#66cc66",
        ["ThemeFgMuted"]         = "#449944",
        ["ThemeFgDimmed"]        = "#226622",
        ["ThemeFgSuccess"]       = "#44ee44",
        ["ThemeFgInfo"]          = "#aa66dd",
        ["ThemeFgDanger"]        = "#dd55cc",
        ["ThemeFgWarn"]          = "#ccbb11",
        ["ThemeFgLink"]          = "#44ff44",
        ["ThemeBgPrimary"]       = "#040e04",
        ["ThemeBgSecondary"]     = "#081408",
        ["ThemeBgCard"]          = "#0c1c0c",
        ["ThemeBgInput"]         = "#061206",
        ["ThemeBgActiveGreen"]   = "#083808",
        ["ThemeBgActiveBlue"]    = "#0e1430",
        ["ThemeBgDanger"]        = "#280e28",
    };

    // ── 아이언맨 팔레트 (크롬/메탈릭 실버 + 레드/골드/아크리액터 블루) ───
    private static System.Collections.Generic.Dictionary<string, string> IronManPalette() => new()
    {
        ["AppBg"]            = "#18181c",
        ["PanelBg"]          = "#26262c",
        ["PanelInnerBg"]     = "#1e1e24",
        ["MenuBarBg"]        = "#303038",
        ["MenuItemBg"]       = "#303038",
        ["MenuItemHover"]    = "#3c3c46",
        ["SubMenuBg"]        = "#1e1e24",
        ["SubMenuItemBg"]    = "#26262c",
        ["SubMenuItemHover"] = "#3c3c46",
        ["SubBtnBg"]         = "#26262c",
        ["SubBtnHover"]      = "#3c3c46",
        ["SplitterColor"]    = "#0a0a0e",
        ["AppFg"]            = "#e0dde4",
        ["FgMuted"]          = "#8888a0",
        ["FgHover"]          = "#dd3322",
        ["TreeFg"]           = "#c8c8d4",
        ["TreeSelBg"]        = "#5a1818",
        ["TreeSelFg"]        = "#ffd0c0",
        ["GridHeaderBg"]     = "#222228",
        ["GridRowBg"]        = "#26262c",
        ["GridRowAltBg"]     = "#1e1e24",
        ["InputBg"]          = "#1e1e24",
        ["InputBorder"]      = "#4a4a58",
        ["InputFg"]          = "#e8e8f0",
        ["TreeHoverBg"]      = "#303038",
        ["Panel4Bg"]         = "#1e1e24",
        ["TabActiveBg"]      = "#5a1818",
        ["TabActiveFg"]      = "#ff4433",
        ["BtnBg"]            = "#26262c",
        ["BtnFg"]            = "#b0b0c0",
        ["BtnBorder"]        = "#4a4a58",
        ["BtnHoverBg"]       = "#3c3c46",
        ["BtnPrimaryBg"]     = "#5a1818",
        ["BtnPrimaryFg"]     = "#ff4433",
        ["BtnPrimaryBorder"] = "#cc2211",
        ["BtnPrimaryHover"]  = "#6a2222",
        ["BtnDangerBg"]      = "#4a1010",
        ["BtnDangerFg"]      = "#ff3322",
        ["BtnDangerBorder"]  = "#aa1111",
        ["BtnDangerHover"]   = "#5a1818",
        ["BtnSuccessBg"]     = "#0e2a3a",
        ["BtnSuccessFg"]     = "#44ccff",
        ["BtnSuccessBorder"] = "#2288bb",
        ["BtnSuccessHover"]  = "#163848",
        ["ThemeBorderDefault"]   = "#4a4a58",
        ["ThemeBorderSubtle"]    = "#303038",
        ["ThemeBorderMuted"]     = "#3a3a44",
        ["ThemeBorderActive"]    = "#cc2211",
        ["ThemeBorderInfo"]      = "#44aaee",
        ["ThemeBorderAccent"]    = "#ddaa22",
        ["ThemeBorderDanger"]    = "#aa1111",
        ["ThemeBorderWarn"]      = "#ddaa22",
        ["ThemeBorderSeparator"] = "#303038",
        ["ThemeBorderPanel"]     = "#3a3a44",
        ["ThemeFgPrimary"]       = "#e8e8f0",
        ["ThemeFgSecondary"]     = "#b0b0c0",
        ["ThemeFgMuted"]         = "#8888a0",
        ["ThemeFgDimmed"]        = "#555566",
        ["ThemeFgSuccess"]       = "#44ccff",
        ["ThemeFgInfo"]          = "#44aaee",
        ["ThemeFgDanger"]        = "#ff3322",
        ["ThemeFgWarn"]          = "#ddaa22",
        ["ThemeFgLink"]          = "#dd3322",
        ["ThemeBgPrimary"]       = "#18181c",
        ["ThemeBgSecondary"]     = "#1e1e24",
        ["ThemeBgCard"]          = "#26262c",
        ["ThemeBgInput"]         = "#1c1c22",
        ["ThemeBgActiveGreen"]   = "#0e2a3a",
        ["ThemeBgActiveBlue"]    = "#0e1e3a",
        ["ThemeBgDanger"]        = "#4a1010",
    };

    // ── Quotation 이벤트 핸들러 (중복 구독 방지용 named handler) ─────────
    private void OnCheckSelectionChanged(System.Collections.Generic.List<ETA.Models.AnalysisItem> items)
    {
        var names = items.Select(a => a.Analyte);

        // Show2 가 신규/오작성수정 패널이면 항상 NewPanel 으로 라우팅
        if (Show2.Content == _quotationNewPanel)
        {
            _quotationNewPanel?.SetSelectedAnalytes(items);
        }
        else if (_quotationCheckPanel?.CurrentAnalysisRecord != null)
        {
            _analysisRequestDetailPanel?.PreviewCheckedItems(names);
        }
        else if (_quotationCheckPanel?.CurrentIssue != null)
        {
            _quotationDetailPanel?.PreviewCheckedItems(names);
        }
        else
        {
            _quotationNewPanel?.SetSelectedAnalytes(items);
        }
    }

    private void OnAnalysisRecordSaved(ETA.Views.Pages.PAGE1.AnalysisRequestRecord rec)
    {
        // 저장(또는 취소) 후 DB에서 다시 불러와 Show2 갱신
        _analysisRequestDetailPanel?.ShowRecord(rec);
    }

    private void OnIssueSaved(ETA.Models.QuotationIssue issue)
    {
        // 저장(또는 취소) 후 DB에서 다시 불러와 Show2 갱신
        _quotationDetailPanel?.ShowIssue(issue);
    }

    private void OnCompanySelected(ETA.Models.Contract company)
    {
        _quotationNewPanel?.SetCompany(company);
        // Show4 분석항목에 선택된 업체 정보 전달 (금액 표시용)
        if (_quotationCheckPanel != null)
            _quotationCheckPanel.SelectedCompany = company;
    }

    private void LogContentChange(string contentName, object? content)
    {

        // 패널 내용이 세팅되면 shimmer sweep 재생
        if (content != null)
        {
            Control? target = contentName switch
            {
                "Show1" => Show1,
                "Show2" or "Show2(inner)" => Show2,
                "Show3" => Show3,
                "Show4" => Show4,
                _ => null,
            };
            if (target is Control targetCtrl)
                Dispatcher.UIThread.Post(() => TextShimmer.AttachAll(targetCtrl), DispatcherPriority.Background);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  창 위치/레이아웃 저장 및 복원
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 현재 모드의 레이아웃 정보 저장
    /// </summary>
    private void SaveCurrentModeLayout()
    {
        if (_positionManager == null || string.IsNullOrEmpty(_currentMode) || _currentMode == "None")
        {
            return;
        }

        try
        {
            string modeKey = LayoutStorageModePrefix + _currentMode;

            var layout = new PageLayoutInfo();

            // 윈도우 위치/크기
            layout.WindowX = this.Position.X;
            layout.WindowY = this.Position.Y;
            layout.WindowWidth = this.Width;
            layout.WindowHeight = this.Height;

            // 왼쪽 패널 너비 조회
            var mainGrid = this.FindControl<Grid>("MainSplitGrid");
            if (mainGrid != null && mainGrid.ColumnDefinitions.Count > 0)
            {
                var colDef = mainGrid.ColumnDefinitions[0];
                if (colDef.Width.IsAbsolute)
                    layout.LeftPanelWidth = colDef.Width.Value;
            }

            // 오른쪽 상단/하단 + 좌우 분할 비율 (플랫 그리드)
            var rightGrid = this.FindControl<Grid>("RightSplitGrid");
            if (rightGrid != null)
            {
                if (rightGrid.RowDefinitions.Count >= 3)
                {
                    layout.UpperStar = rightGrid.RowDefinitions[0].Height.Value;
                    layout.LowerStar = rightGrid.RowDefinitions[2].Height.Value;
                }
                if (rightGrid.ColumnDefinitions.Count >= 3)
                {
                    layout.Content2Star = rightGrid.ColumnDefinitions[0].Width.Value;
                    layout.Content4Star = rightGrid.ColumnDefinitions[2].Width.Value;
                }
            }

            layout.SavedAt = DateTime.Now;
            _positionManager?.SavePageLayout(modeKey, layout); // null 체크 - WindowPositionManager 비활성화 시
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>
    /// 특정 모드의 저장된 레이아웃 정보 복원
    /// </summary>
    private void RestoreModeLayout(string modeName, double minLowerStar = 0)
    {
        if (_positionManager == null)
        {
            return;
        }

        try
        {
            string modeKey = LayoutStorageModePrefix + modeName;

            var layout = _positionManager?.GetPageLayout(modeKey); // null 체크 - WindowPositionManager 비활성화 시
            if (layout == null)
            {
                return;
            }


            // 왼쪽 패널 너비 복원
            if (layout.LeftPanelWidth > 0)
                SetLeftPanelWidth(layout.LeftPanelWidth);

            // 레이아웃 비율 복원 (minLowerStar 이상 보장)
            SetContentLayout(
                content2Star: layout.Content2Star,
                content4Star: layout.Content4Star,
                upperStar: layout.UpperStar,
                lowerStar: Math.Max(layout.LowerStar, minLowerStar));

        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>
    /// 사용자 변경 시 WindowPositionManager 업데이트
    /// (로그인 후 호출)
    /// </summary>
    public void UpdateCurrentUser(string newUserId)
    {
        if (!string.IsNullOrWhiteSpace(newUserId))
        {
            CurrentUserManager.Instance.SetCurrentUser(newUserId);
            // 새 사용자로 WindowPositionManager 재초기화 - 완전 비활성화 (스턱 방지)
            _positionManager = null; // new WindowPositionManager(newUserId);
        }
    }

    /// <summary>
    /// 로그 파일 경로 조회 (디버깅용)
    /// </summary>
    public string GetPositionLogFilePath()
    {
        return _positionManager?.GetLogFilePath() ?? "Unknown";
    }

    // ═════════════════════════════════════════════════════════════════════════
    //  GridSplitter 실시간 레이아웃 저장
    // ═════════════════════════════════════════════════════════════════════════

    private System.Threading.Timer? _layoutSaveTimer;

    /// <summary>
    /// GridSplitter 이벤트 핸들러 연결 (창 로드 후 한번만 실행)
    /// </summary>
    private void SetupSplitterEvents(object? sender, EventArgs e)
    {
        try
        {
            // 메인 좌우 스플리터 (Show1 ↔ Show2,3,4)
            var mainSplitter = this.FindControl<GridSplitter>("MainSplitter");
            if (mainSplitter != null)
            {
                mainSplitter.DragCompleted += OnSplitterMoved;
            }

            // 수평 스플리터 (Show2,3 ↔ Show4)
            var horizontalSplitter = this.FindControl<GridSplitter>("HorizontalSplitter");
            if (horizontalSplitter != null)
            {
                horizontalSplitter.DragCompleted += OnSplitterMoved;
            }

            // 수직 스플리터 (Show2 ↔ Show3)
            var verticalSplitter = this.FindControl<GridSplitter>("VerticalSplitter");
            if (verticalSplitter != null)
            {
                verticalSplitter.DragCompleted += OnSplitterMoved;
            }
        }
        catch (Exception ex)
        {
        }
    }

    /// <summary>
    /// GridSplitter 이동 완료 시 지연 저장 (300ms 후)
    /// </summary>
    private void OnSplitterMoved(object? sender, VectorEventArgs e)
    {
        // 기존 타이머가 있으면 취소
        _layoutSaveTimer?.Dispose();

        // 300ms 후 레이아웃 저장 (연속 드래그 중 중복 저장 방지)
        _layoutSaveTimer = new System.Threading.Timer(
            callback: _ =>
            {
                Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
                {
                    SaveCurrentModeLayout();
                });
            },
            state: null,
            dueTime: TimeSpan.FromMilliseconds(300),
            period: System.Threading.Timeout.InfiniteTimeSpan);
    }

    private void TEST_Click(object? sender, RoutedEventArgs e) { }

    // ── 애니메이션 토글 ────────────────────────────────────────────────────────
    /// <summary>모든 컨트롤의 Transitions 제거 (재귀)</summary>
    private void DisableTransitionsRecursive(Control? control)
    {
        if (control == null) return;

        // Transitions 제거
        if (control.Transitions != null && control.Transitions.Count > 0)
        {
            control.Transitions.Clear();
        }

        // 자식 컨트롤 재귀 처리
        foreach (var child in control.GetVisualChildren().OfType<Control>())
        {
            DisableTransitionsRecursive(child);
        }
    }

    /// <summary>애니메이션 토글</summary>
    public void ToggleAnimations()
    {
        App.AnimationsEnabled = !App.AnimationsEnabled;

        if (App.AnimationsEnabled)
        {
            Debug.WriteLine("🎬 애니메이션 활성화");
        }
        else
        {
            Debug.WriteLine("⏸️ 애니메이션 비활성화 중...");
            // Show1~4의 Transitions 제거
            foreach (var show in new[] { Show1, Show2, Show3, Show4 })
            {
                if (show != null && show.Transitions != null)
                    show.Transitions.Clear();
            }
            DisableTransitionsRecursive(this);
        }

        // UI 업데이트 - 리플렉션으로 IsChecked 설정
        var btn = this.FindControl<Control>("btnAnimations");
        if (btn != null)
        {
            var prop = btn.GetType().GetProperty("IsChecked");
            if (prop != null)
                prop.SetValue(btn, App.AnimationsEnabled);
        }
    }

    private void BtnAnimations_Click(object? sender, RoutedEventArgs e)
    {
        ToggleAnimations();
    }
}
