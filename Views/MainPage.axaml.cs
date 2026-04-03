using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
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
using Microsoft.Data.Sqlite;
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
    private WasteDataQueryPage?      _wasteDataQueryPage;
    private WasteNameReconcilePage?            _wasteNameReconcilePage;
    private WaterQualityNameReconcilePage?     _waterQualityNameReconcilePage;
    private WasteSampleListPage?         _wasteSampleListPage;
    private ProcessingFacilityPage?      _processingFacilityPage;
    private ResultSubmitMeasurePage?     _resultSubmitMeasurePage;
    private TestReportPage?              _resultSubmitMeasureTestReport;
    private ResultSubmitErpPage?         _resultSubmitErpPage;
    private ResultSubmitZero4Page?       _resultSubmitZero4Page;
    private PurchasePage?      _purchasePage;
    private RepairPage?       _repairPage;
    private RiskManagePage?   _riskPage;
    private SchedulePage?     _schedulePage;
    private TestReportPage?         _testReportPage;
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

    // ── 마지막으로 표시된 issue 캐시 (트리뷰 선택 null 시 복원용) ─────
    private ETA.Models.QuotationIssue? _lastShownIssue;

    // ── 분석의뢰 상세/목록 패널 ──────────────────────────────────────
    private AnalysisRequestDetailPanel?  _analysisRequestDetailPanel;
    private AnalysisRequestListPanel?    _analysisRequestListPanel;

    public MainPage()
    {
        InitializeComponent();
        DataContext = new MainWindowViewModel();
        ApplyTheme(true);
        
        // WindowPositionManager 초기화
        _positionManager = new WindowPositionManager(CurrentUserManager.Instance.CurrentUserId);
        
        // 윈도우 이벤트 연결
        this.Opened += MainPage_Opened;
        this.Closing += MainPage_Closing;
    }

    private void MainPage_Opened(object? sender, EventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainPage] Opened 이벤트");
        LoadProfileInfo();

        // 서버 연결 시 오늘까지 업무분장 자동 연장
        if (DbConnectionFactory.IsMariaDb)
            AnalysisRequestService.AutoExtendAssignmentsToToday();

        // 저장된 글자 크기 복원
        var savedScale = LoadFontScale();
        ApplyFontScale(savedScale);
        if (sldFontSize != null) sldFontSize.Value = savedScale;
        if (txbFontScale != null) txbFontScale.Text = $"{(int)(savedScale * 100)}%";
    }

    private void LoadProfileInfo()
    {
        try
        {
            var empId = CurrentEmployeeId;
            if (string.IsNullOrEmpty(empId)) return;

            var agents = AgentService.GetAllItems();
            var me = agents.FirstOrDefault(a => a.사번 == empId);
            if (me == null) return;

            if (profileName != null)
                profileName.Text = me.성명;

            if (profilePhoto != null && !string.IsNullOrEmpty(me.PhotoPath))
            {
                // 로컬 파일 없으면 DB에서 가져와 캐시
                ETA.Services.SERVICE1.AgentService.EnsurePhotoLocal(me.사번, me.PhotoPath);

                var fullPath = Path.IsPathRooted(me.PhotoPath)
                    ? me.PhotoPath
                    : Path.Combine(ETA.Services.SERVICE1.AgentService.GetPhotoDirectory(), me.PhotoPath);
                if (File.Exists(fullPath))
                {
                    profilePhoto.Source   = new Avalonia.Media.Imaging.Bitmap(fullPath);
                    profilePhoto.IsVisible = true;
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Profile] 로드 실패: {ex.Message}");
        }
    }

    private void ProfileName_Click(object? sender, PointerPressedEventArgs e)
        => MyTask_Click(sender, new RoutedEventArgs());

    private void MainPage_Closing(object? sender, WindowClosingEventArgs e)
    {
        System.Diagnostics.Debug.WriteLine("[MainPage] Closing 이벤트 - 현재 모드 레이아웃 저장");
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
            Background      = new SolidColorBrush(Color.Parse("#1e1e2e")),
            BorderThickness = new Thickness(0),
            SelectionMode   = SelectionMode.Multiple,
            Margin          = new Thickness(2, 0, 2, 2),
        };
        LoadAnalysisItems();

        // ── 날짜 범위 행 ──────────────────────────────────────────────────
        var txbRange = new TextBlock
        {
            Text              = DateTime.Today.ToString("yyyy-MM-dd"),
            FontSize          = 10,
            Foreground        = new SolidColorBrush(Color.Parse("#aaaaaa")),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            MinWidth          = 160,
            FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
        };

        var btnCal = new Button
        {
            Content         = "📅",
            Width           = 28,
            Height          = 22,
            FontSize        = 11,
            Padding         = new Thickness(0),
            Background      = new SolidColorBrush(Color.Parse("#2a3a4a")),
            Foreground      = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555555")),
        };
        ToolTip.SetTip(btnCal, "기간 선택 (드래그로 범위 설정)");

        var btnToday = new Button
        {
            Content         = "오늘",
            Width           = 42,
            Height          = 22,
            FontSize        = 10,
            Padding         = new Thickness(4, 0),
            Background      = new SolidColorBrush(Color.Parse("#3a5a3a")),
            Foreground      = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#666666")),
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
        };
        calendar.SelectedDates.Add(DateTime.Today);

        // ── 이벤트 연결 ───────────────────────────────────────────────────
        btnCal.Click += (_, _) => calendar.IsVisible = !calendar.IsVisible;

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

        // ── 분석항목 컨테이너 (DockPanel) ────────────────────────────────
        var header = new StackPanel { Spacing = 0 };
        header.Children.Add(dateRow);
        header.Children.Add(calendar);

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
                        FontSize          = 11,
                        Foreground        = new SolidColorBrush(Color.Parse("#cccccc")),
                        FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
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
                                FontSize   = 9,
                                Foreground = new SolidColorBrush(Color.Parse(fg)),
                                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
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
                        Text = "계약업체 없음", FontSize = 11,
                        Foreground = new SolidColorBrush(Color.Parse("#555")),
                        Margin = new Thickness(8, 8),
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    });
            }
            catch { }
        }

        // ── 탭 토글 버튼 행 ───────────────────────────────────────────────
        var btnAnalysis = new Button
        {
            Content         = "분석항목",
            FontSize        = 11,
            Height          = 26,
            Padding         = new Thickness(12, 0),
            BorderThickness = new Thickness(0),
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Background      = new SolidColorBrush(Color.Parse("#3a4a6a")),
            Foreground      = new SolidColorBrush(Color.Parse("#ddeeff")),
        };
        var btnContract = new Button
        {
            Content         = "계약업체",
            FontSize        = 11,
            Height          = 26,
            Padding         = new Thickness(12, 0),
            BorderThickness = new Thickness(0),
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
            Foreground      = new SolidColorBrush(Color.Parse("#888899")),
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

            var agentFont  = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
            var agentFontM = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

            var form = new StackPanel { Spacing = 6, Margin = new Thickness(6) };

            form.Children.Add(new TextBlock
            {
                Text = isEdit ? "✏️ 일반업무 수정" : "➕ 신규 일반업무 추가",
                FontSize = 12, FontWeight = FontWeight.Bold,
                FontFamily = agentFontM,
                Foreground = new SolidColorBrush(Color.Parse("#ccbb88")),
            });

            var txtName = new TextBox
            {
                Text = existing?.업무명 ?? "",
                Watermark = "업무명",
                FontFamily = agentFont, FontSize = 11,
                Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
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
                FontFamily = agentFont, FontSize = 11,
                Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 4),
            };
            form.Children.Add(txtContent);

            // 담당자 콤보박스
            var agentRow = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal, Spacing = 6 };
            agentRow.Children.Add(new TextBlock
            {
                Text = "담당자", FontSize = 10, FontFamily = agentFont,
                Foreground = Res("FgMuted"),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            });
            var agents = ETA.Services.SERVICE1.AgentService.GetAllItems().OrderBy(a => a.성명).ToList();
            var cboAgent = new ComboBox
            {
                Width = 160, FontSize = 11, FontFamily = agentFont,
                Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
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
                Text = "마감일", FontSize = 10, FontFamily = agentFont,
                Foreground = Res("FgMuted"),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            });
            var txtDeadline = new TextBox
            {
                Text = existing?.마감일 ?? "",
                Watermark = "yyyy-MM-dd",
                Width = 120, FontFamily = agentFont, FontSize = 11,
                Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
                Foreground = Res("AppFg"),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
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
                    Text = "상태", FontSize = 10, FontFamily = agentFont,
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
                        Content = st, FontSize = 9, FontFamily = agentFont,
                        Padding = new Thickness(6, 2),
                        CornerRadius = new CornerRadius(6),
                        BorderThickness = new Thickness(1),
                        Background  = active ? new SolidColorBrush(Color.Parse(sbg)) : new SolidColorBrush(Color.Parse("#222")),
                        Foreground  = active ? new SolidColorBrush(Color.Parse(sfg)) : new SolidColorBrush(Color.Parse("#666")),
                        BorderBrush = active ? new SolidColorBrush(Color.Parse(sbd)) : new SolidColorBrush(Color.Parse("#444")),
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
                            child.Background  = on2 ? new SolidColorBrush(Color.Parse(b2)) : new SolidColorBrush(Color.Parse("#222"));
                            child.Foreground  = on2 ? new SolidColorBrush(Color.Parse(f2)) : new SolidColorBrush(Color.Parse("#666"));
                            child.BorderBrush = on2 ? new SolidColorBrush(Color.Parse(d2)) : new SolidColorBrush(Color.Parse("#444"));
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
                FontFamily = agentFont, FontSize = 10,
                Background = new SolidColorBrush(Color.Parse("#1a3a2a")),
                Foreground = new SolidColorBrush(Color.Parse("#88ee88")),
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
                    FontFamily = agentFont, FontSize = 10,
                    Background = new SolidColorBrush(Color.Parse("#3a1a1a")),
                    Foreground = new SolidColorBrush(Color.Parse("#ee8888")),
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
                FontFamily = agentFont, FontSize = 10,
                Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
                Foreground = new SolidColorBrush(Color.Parse("#aaaaaa")),
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
                var agentFont = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
                var agentFontM = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");
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
                        FontSize          = 11,
                        FontWeight        = FontWeight.SemiBold,
                        Foreground        = new SolidColorBrush(Color.Parse("#cccccc")),
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
                                FontSize   = 9,
                                FontFamily = agentFont,
                                Foreground = new SolidColorBrush(Color.Parse(fg)),
                            },
                        });
                    }
                    if (chipsPanel.Children.Count == 0)
                        chipsPanel.Children.Add(new TextBlock
                        {
                            Text = "미배정", FontSize = 9, FontFamily = agentFont,
                            Foreground = new SolidColorBrush(Color.Parse("#555566")),
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
                        Text = "일반업무 없음", FontSize = 11,
                        Foreground = new SolidColorBrush(Color.Parse("#555")),
                        Margin = new Thickness(8, 8),
                        FontFamily = agentFont,
                    });

                // 하단: 추가 버튼
                var btnAdd = new Button
                {
                    Content         = "＋ 업무 추가",
                    FontSize        = 11,
                    Height          = 28,
                    Padding         = new Thickness(12, 0),
                    BorderThickness = new Thickness(0),
                    CornerRadius    = new CornerRadius(4),
                    FontFamily      = agentFontM,
                    Background      = new SolidColorBrush(Color.Parse("#2a3a2a")),
                    Foreground      = new SolidColorBrush(Color.Parse("#88cc88")),
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
            FontSize        = 11,
            Height          = 26,
            Padding         = new Thickness(12, 0),
            BorderThickness = new Thickness(0),
            FontFamily      = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
            Foreground      = new SolidColorBrush(Color.Parse("#888899")),
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

        // 분장표준처리 테이블 컬럼 순서 + 약칭(row2)으로 가져오기
        var analytes = AnalysisRequestService.GetOrderedAnalytes();
        Debug.WriteLine($"[LoadAnalysisItems] 로드된 분석항목 수: {analytes.Count}");

        foreach (var (fullName, shortName) in analytes)
        {
            var (badgeBg, badgeFg) = BadgeColorHelper.GetBadgeColor(shortName);

            // 카드 UI: Border + Grid로 구성
            var cardBorder = new Border
            {
                Background      = new SolidColorBrush(Color.Parse("#2a2a38")),
                BorderThickness = new Thickness(1),
                BorderBrush     = new SolidColorBrush(Color.Parse("#444455")),
                CornerRadius    = new CornerRadius(3),
                Padding         = new Thickness(6, 3),
                Margin          = new Thickness(2, 1),
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
                    FontSize   = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                }
            };
            Grid.SetColumn(catBadge, 0);
            cardGrid.Children.Add(catBadge);

            // Col 1: 항목명
            var nameBlock = new TextBlock
            {
                Text              = fullName,
                FontSize          = 11,
                FontWeight        = FontWeight.SemiBold,
                Foreground        = new SolidColorBrush(Color.Parse("#a0d060")),
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                TextTrimming      = TextTrimming.CharacterEllipsis,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            };
            Grid.SetColumn(nameBlock, 1);
            cardGrid.Children.Add(nameBlock);

            // Col 2: 할당 배지 (UpdateAssignmentInfo에서 갱신)
            var assignBadge = new TextBlock
            {
                Text              = "·",
                FontSize          = 9,
                Foreground        = Res("FgMuted"),
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
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
                        System.Diagnostics.Debug.WriteLine($"[DragStart] {analyte}");
                    }
                }
            };

            // ListBoxItem에 카드 넣기
            var listBoxItem = new ListBoxItem
            {
                Content = cardBorder,
                Tag = fullName,
                Padding = new Thickness(0),
                Margin = new Thickness(0)
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
                        assignBadge.Foreground = new SolidColorBrush(Color.Parse("#88cc88"));

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
        var rightGrid = this.FindControl<Grid>("RightSplitGrid");
        if (rightGrid != null && rightGrid.RowDefinitions.Count >= 3)
        {
            rightGrid.RowDefinitions[0].Height = new GridLength(upperStar, GridUnitType.Star);
            rightGrid.RowDefinitions[2].Height = lowerStar > 0
                ? new GridLength(lowerStar, GridUnitType.Star)
                : new GridLength(0);

            var lowerBorder = this.FindControl<Border>("LowerBorder");
            if (lowerBorder != null) lowerBorder.IsVisible = lowerStar > 0;

            var vSplitter = this.FindControl<GridSplitter>("VerticalSplitter");
            if (vSplitter != null) vSplitter.IsVisible = lowerStar > 0;
        }

        var upperGrid = this.FindControl<Grid>("UpperContentGrid");
        if (upperGrid != null && upperGrid.ColumnDefinitions.Count >= 3)
        {
            upperGrid.ColumnDefinitions[0].Width = new GridLength(content2Star, GridUnitType.Star);
            upperGrid.ColumnDefinitions[2].Width = content4Star > 0
                ? new GridLength(content4Star, GridUnitType.Star)
                : new GridLength(0);

            var content4Border = this.FindControl<Border>("Content4Border");
            if (content4Border != null) content4Border.IsVisible = content4Star > 0;

            var hSplitter = this.FindControl<GridSplitter>("HorizontalSplitter");
            if (hSplitter != null) hSplitter.IsVisible = content4Star > 0;
        }
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
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
        RestoreModeLayout("MyTask");
    }

    private void Agent_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "Agent";

        if (_agentTreePage == null)
        {
            _agentTreePage = new AgentTreePage();
            _agentTreePage.DetailPanelChanged += panel =>
            {
                Show2.Content = panel;
                LogContentChange("Show2", panel);
            };
        }

        // Show3: 직원 선택 시 타임라인 차트
        _agentTreePage.Show3ContentRequest = ctrl =>
        {
            Show3.Content = ctrl;
            LogContentChange("Show3", ctrl);
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
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);

        // ✅ Content4: wrapper를 Show4에 고정 (이후 내용은 wrapper.Content 교체)
        Show4.Content = _show4AgentWrapper;
        LogContentChange("Show4", _analysisItemsListBox);

        // AgentTreePage에 ListBox 참조 전달
        _agentTreePage.AnalysisItemsListBox = _analysisItemsListBox;
        
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

            // 업소정보(편집 패널) → Show4 (기존 Show2에서 이동)
            _wasteCompanyPage.DetailPanelChanged += panel =>
            {
                Show4.Content = panel;
                LogContentChange("Show4", panel);
            };

            // 업체 선택 → Show2(차트) + Show3(목록) 자료조회 기능 직접 구현
            _wasteCompanyPage.CompanySelected += company =>
            {
                List<ETA.Models.WasteAnalysisResult> results;
                try { results = ETA.Services.SERVICE2.WasteDataService.GetResults(company.업체명); }
                catch (Exception ex)
                {
                    Show2.Content = new TextBlock
                    {
                        Text = $"조회 오류: {ex.Message}",
                        FontSize = 11, FontFamily = _wasteFont,
                        Foreground = Brush.Parse("#cc4444"),
                        Margin = new Thickness(12),
                    };
                    return;
                }

                // Show2 = 분석결과 추이 차트
                Show2.Content = BuildWasteBarLinePanel(company.업체명, results);
                LogContentChange("Show2", Show2.Content as Control);

                // Show3 = 분석결과 목록 테이블
                Show3.Content = BuildDataListPanel(company.업체명, results, company.약칭);
                LogContentChange("Show3", Show3.Content as Control);

                SetContentLayout(content2Star: 2, content4Star: 1, upperStar: 5, lowerStar: 5);
            };

            _wasteCompanyPage.OrderSaved += () => { };
        }

        Show1.Content = _wasteCompanyPage;
        LogContentChange("Show1", _wasteCompanyPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = null;
        LogContentChange("Show4", null);
        _wasteCompanyPage.LoadData();
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
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    private Control BuildOrderHistoryPanel()
    {
        var root = new StackPanel { Spacing = 0 };

        // 헤더
        root.Children.Add(new Border
        {
            Background   = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Padding      = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text       = "📋  의뢰내역",
                FontSize   = 11, FontWeight = FontWeight.SemiBold,
                FontFamily = _wasteFont, Foreground = Brush.Parse("#8899bb"),
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
                    Background = Brush.Parse("#22223a"),
                    Padding    = new Thickness(8, 3),
                    Child = new TextBlock
                    {
                        Text       = $"{d:yyyy-MM-dd} ({DayKr(d)})  {rows.Count}건",
                        FontSize   = 10, FontWeight = FontWeight.SemiBold,
                        FontFamily = _wasteFont, Foreground = Brush.Parse("#8899bb"),
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

                    void AddCell(int col, string text, string fg, int size = 10, FontWeight fw = FontWeight.Normal)
                    {
                        var tb = new TextBlock
                        {
                            Text      = text, FontSize = size, FontFamily = _wasteFont,
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
                Text = $"오류: {ex.Message}", FontSize = 10,
                FontFamily = _wasteFont, Foreground = Brush.Parse("#cc4444"),
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
                FontSize = 12, FontFamily = _wasteFont,
                Foreground = Brush.Parse("#888888"),
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
                FontSize = 11, FontFamily = _wasteFont,
                Foreground = Brush.Parse("#cc4444"),
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
        // 8개 항목 개별 차트 생성 (항목별로 값이 있는 것만 최근 10건)
        var charts = new List<WasteSingleSeriesChart>();
        foreach (var (label, color, getValue) in WasteBarLineChartControl.Series)
        {
            var withValue = results.Where(r => getValue(r).HasValue).ToList();
            var recent    = withValue.Count > 10 ? withValue.Skip(withValue.Count - 10).ToList() : withValue;
            charts.Add(new WasteSingleSeriesChart(label, color, getValue, recent));
        }

        // 상단 바: Bar/Line 토글
        var topBar = new WrapPanel { Margin = new Thickness(4, 3), VerticalAlignment = VerticalAlignment.Center };

        topBar.Children.Add(new TextBlock
        {
            Text = $"{업체명}  분석결과 추이", FontSize = 11, FontFamily = _wasteFont,
            Foreground = Brush.Parse("#8899bb"), FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 0, 12, 0),
        });

        var btnBar = new Avalonia.Controls.Button
        {
            Content = "Bar", FontSize = 10, FontFamily = _wasteFont,
            Padding = new Thickness(8, 2), Margin = new Thickness(0, 0, 2, 0),
            Background = Brush.Parse("#3a3a5a"), Foreground = Brush.Parse("#ffffff"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(3),
        };
        var btnLine = new Avalonia.Controls.Button
        {
            Content = "Line", FontSize = 10, FontFamily = _wasteFont,
            Padding = new Thickness(8, 2), Margin = new Thickness(0, 0, 8, 0),
            Background = Brush.Parse("#2a2a3a"), Foreground = Brush.Parse("#888888"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(3),
        };
        btnBar.Click += (_, _) =>
        {
            foreach (var c in charts) c.SetBarMode(true);
            btnBar.Background  = Brush.Parse("#3a3a5a"); btnBar.Foreground  = Brush.Parse("#ffffff");
            btnLine.Background = Brush.Parse("#2a2a3a"); btnLine.Foreground = Brush.Parse("#888888");
        };
        btnLine.Click += (_, _) =>
        {
            foreach (var c in charts) c.SetBarMode(false);
            btnLine.Background = Brush.Parse("#3a3a5a"); btnLine.Foreground = Brush.Parse("#ffffff");
            btnBar.Background  = Brush.Parse("#2a2a3a"); btnBar.Foreground  = Brush.Parse("#888888");
        };
        topBar.Children.Add(btnBar);
        topBar.Children.Add(btnLine);

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
                BorderBrush     = new SolidColorBrush(Color.Parse("#252535")),
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
                    FontSize   = 10, FontFamily = _wasteFont,
                    Foreground = Brush.Parse(fg),
                }
            });
        }
        titleRow.Children.Add(new TextBlock
        {
            Text       = $"📋  {업체명}  분석결과 내역  ({results.Count}건)",
            FontSize   = 11, FontWeight = FontWeight.SemiBold,
            FontFamily = _wasteFont, Foreground = Brush.Parse("#8899bb"),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        });
        root.Children.Add(new Border
        {
            Background   = Brush.Parse("#1a1a28"),
            CornerRadius = new CornerRadius(6, 6, 0, 0),
            Child        = titleRow,
        });

        if (results.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "분석결과 없음",
                FontSize = 11, FontFamily = _wasteFont,
                Foreground = Brush.Parse("#555555"),
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
                Text = text, FontSize = 10, FontFamily = _wasteFont,
                FontWeight = fw,
                Foreground = text == "—" ? Brush.Parse("#444444") : Brush.Parse(fg),
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

    private void WasteDataQuery_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteDataQuery";

        if (_wasteDataQueryPage == null)
        {
            _wasteDataQueryPage = new WasteDataQueryPage();
            _wasteDataQueryPage.CompanySelected += company =>
            {
                List<ETA.Models.WasteAnalysisResult> results;
                try { results = ETA.Services.SERVICE2.WasteDataService.GetResults(company.업체명); }
                catch (Exception ex)
                {
                    Show3.Content = new TextBlock
                    {
                        Text = $"조회 오류: {ex.Message}",
                        FontSize = 11, FontFamily = _wasteFont,
                        Foreground = Brush.Parse("#cc4444"),
                        Margin = new Thickness(12),
                    };
                    return;
                }

                // Show2 = Bar/Line 차트 (최근 10건)
                Show2.Content = BuildWasteBarLinePanel(company.업체명, results);
                LogContentChange("Show2", Show2.Content as Control);

                Show3.Content = BuildDataListPanel(company.업체명, results, company.약칭);
                LogContentChange("Show3", Show3.Content as Control);
            };
        }

        Show1.Content = _wasteDataQueryPage;
        LogContentChange("Show1", _wasteDataQueryPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = null;
        LogContentChange("Show3", null);
        Show4.Content = BuildOrderHistoryPanel();
        LogContentChange("Show4", Show4.Content as Control);
        _wasteDataQueryPage.LoadData();
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "", "", "", "", "", "설정");
        SetLeftPanelWidth(300);
        SetContentLayout(content2Star: 1, content4Star: 1, upperStar: 6, lowerStar: 4);

        RestoreModeLayout("WasteDataQuery");
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
        }

        Show1.Content = _wasteSampleListPage;
        LogContentChange("Show1", _wasteSampleListPage);
        Show2.Content = null;
        LogContentChange("Show2", null);
        Show3.Content = _wasteSampleListPage.CompanyTreePanel;
        LogContentChange("Show3", _wasteSampleListPage.CompanyTreePanel as Control);
        Show4.Content = BuildOrderHistoryPanel();
        LogContentChange("Show4", Show4.Content as Control);
        _wasteSampleListPage.LoadData();
        _wasteSampleListPage.LoadCompanyTree();
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "날짜 추가", "", "", "", "", "");
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
        SetContentLayout(content2Star: 0, content4Star: 0, upperStar: 3, lowerStar: 7);

        RestoreModeLayout("ProcessingFacility");
    }

    private void WasteAnalysisInput_Click(object? sender, RoutedEventArgs e)
    {
        _currentMode = "WasteAnalysisInput";
        Show1.Content = BuildAnalysisInputCategoryPanel();
        LogContentChange("Show1", Show1.Content as Control);
        Show2.Content = null;
        Show3.Content = null;
        Show4.Content = null;
        SetSubMenu("", "", "", "", "", "", "");
        SetLeftPanelWidth(260);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 8, lowerStar: 2);
        RestoreModeLayout("WasteAnalysisInput");
    }

    private Control BuildAnalysisInputCategoryPanel()
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(10) };

        root.Children.Add(new TextBlock
        {
            Text = "분석결과입력",
            FontSize = 14, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = (Brush)Application.Current!.Resources["AppFg"]!,
            Margin = new Thickness(0, 0, 0, 8),
        });

        var categories = new (string Label, string Icon, string Bg, string Fg, string Bd)[]
        {
            ("수질분석센터", "🧪", "#1a1a2a", "#aaccff", "#3a5a8a"),
            ("폐수배출업소", "🏭", "#1a2a1a", "#aaccaa", "#3a6a3a"),
            ("처리시설",     "⚙️", "#2a1a2a", "#ccaaff", "#6a3a8a"),
        };

        foreach (var c in categories)
        {
            var captured = c;
            var btn = new Button
            {
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Padding = new Thickness(14, 10),
                Background = new SolidColorBrush(Color.Parse(c.Bg)),
                Foreground = new SolidColorBrush(Color.Parse(c.Fg)),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse(c.Bd)),
                CornerRadius = new CornerRadius(6),
                Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Content = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 10,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = c.Icon, FontSize = 18,
                            VerticalAlignment = VerticalAlignment.Center,
                        },
                        new TextBlock
                        {
                            Text = c.Label, FontSize = 13,
                            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                            Foreground = new SolidColorBrush(Color.Parse(c.Fg)),
                            VerticalAlignment = VerticalAlignment.Center,
                        },
                    },
                },
            };
            btn.Click += (_, _) => OnAnalysisInputCategorySelected(captured.Label);
            root.Children.Add(btn);
        }

        return root;
    }

    private void OnAnalysisInputCategorySelected(string category)
    {
        // TODO: 카테고리별 분석결과 입력 화면 구현
        Show2.Content = new TextBlock
        {
            Text = $"📋 {category} — 분석결과 입력",
            FontSize = 13,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = (Brush)Application.Current!.Resources["AppFg"]!,
            Margin = new Thickness(10),
        };
        LogContentChange("Show2", Show2.Content as Control);
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
            _contractPage.PricePanelChanged += panel =>
            {
                Show3.Content = panel;
                LogContentChange("Show3", panel);
            };
        }

        // 최초 진입 시 분석단가 컬럼 보장 (없는 컬럼 자동 추가)
        try { ContractService.EnsureContractPriceColumns(); }
        catch (Exception ex) { Debug.WriteLine($"[Contract] EnsureContractPriceColumns 오류: {ex.Message}"); }

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
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 7, lowerStar: 3);
        // Show3: 단가 항목 편집 폼 영역 (계약 선택 후 항목 클릭 시 로드)
        
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
        Show4.Content = null;
        LogContentChange("Show4", null);
        _bt1SaveAction = null;

        SetSubMenu("새로고침", "삭제", "", "", "", "", "", "");
        SetLeftPanelWidth(320);
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 5, lowerStar: 5);
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
        Show4.Content = null;
        LogContentChange("Show4", null);
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

            // ✏️ 오작성 수정: 기존 Id 덮어쓰기 — 메타 수정
            _quotationDetailPanel.CorrectRequested += issue =>
            {
                var target = issue ?? _lastShownIssue;
                if (target == null) return;
                _lastShownIssue = target;
                _quotationNewPanel!.LoadFromIssueCorrect(target);
                Show2.Content = _quotationNewPanel;
                LogContentChange("Show2", _quotationNewPanel);
            };

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
                _orderRequestEditPanel ??= new OrderRequestEditPanel();
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
                Show4.Content = _quotationPage;
                LogContentChange("Show4", _quotationPage);
                // 편집 대상 설정
                _quotationCheckPanel!.CurrentIssue = issue;
            };

            // 분석의뢰내역 탭으로 전환됨 → Content4: 의뢰 리스트 + TODO 패널
            _quotationHistoryPanel.AnalysisTabActivated += () =>
            {
                _analysisRequestListPanel ??= new AnalysisRequestListPanel();
                Show4.Content = _analysisRequestListPanel;
                LogContentChange("Show4", _analysisRequestListPanel);
            };

            // 견적발행내역 탭으로 복귀 → Content4: 계약업체 목록 + TODO 패널
            _quotationHistoryPanel.QuotationTabActivated += () =>
            {
                Show4.Content = _quotationPage;
                LogContentChange("Show4", _quotationPage);
            };

            // 거래명세서 발행내역 탭으로 전환 → Show2/Show4 초기화
            _quotationHistoryPanel.StatementTabActivated += () =>
            {
                Show2.Content = null;
                Show4.Content = null;
                LogContentChange("Show4", null);
            };

            // 분석의뢰내역 노드 선택
            _quotationHistoryPanel.AnalysisRequestSelected += record =>
            {
                _analysisRequestDetailPanel ??= new AnalysisRequestDetailPanel();
                _analysisRequestDetailPanel.CheckPanel = _quotationCheckPanel;
                _analysisRequestDetailPanel.ShowRecord(record);
                Show2.Content = _analysisRequestDetailPanel;
                LogContentChange("Show2", _analysisRequestDetailPanel);
                _analysisRequestListPanel ??= new AnalysisRequestListPanel();
                Show4.Content = _analysisRequestListPanel;
                LogContentChange("Show4", _analysisRequestListPanel);
                _analysisRequestListPanel.AddRecord(record);
                // 편집 대상 설정
                _quotationCheckPanel!.CurrentAnalysisRecord = record;
            };
        }

        Show1.Content = _quotationHistoryPanel;
        LogContentChange("Show1", _quotationHistoryPanel);
        Show2.Content = _quotationDetailPanel;
        LogContentChange("Show2", _quotationDetailPanel);
        Show3.Content = _quotationCheckPanel;
        LogContentChange("Show3", _quotationCheckPanel);
        Show4.Content = _quotationPage;
        LogContentChange("Show4", _quotationPage);
        _bt1SaveAction = null;

        _quotationHistoryPanel.LoadData();
        _quotationCheckPanel.LoadData();
        _quotationPage.LoadData();

        SetSubMenu("새로고침", "신규 작성", "삭제", "엑셀 내보내기", "인쇄", "", "설정");
        SetLeftPanelWidth(430);
        // Content2(세부내역) 50% : Content4(업체목록) 50%
        // 하단(Content3 분석항목) ≈ 23%  (13 : 4 → 76% : 24%)
        SetContentLayout(content2Star: 7, content4Star: 3, upperStar: 13, lowerStar: 4);
        
        // 저장된 레이아웃 복원 (Show3 항상 표시 보장)
        RestoreModeLayout("Quotation", minLowerStar: 4);
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
            FontSize = 14,
            Foreground = Res("AppFg"),
            Margin = new Thickness(8, 6, 0, 4),
        };

        var clearBtn = new Button
        {
            Content = "전체 지우기",
            FontFamily = _tradeFont,
            FontSize = 11,
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
            FontSize = 12,
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
                FontSize = 13,
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
            FontSize = 13,
            Foreground = Res("AppFg"),
        });
        root.Children.Add(new TextBlock
        {
            Text = $"견적번호: {string.Join(", ", issues.Select(i => i.견적번호))}",
            FontFamily = _tradeFont,
            FontSize = 11,
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
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
        var font  = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");
        var fontR = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

        // 결과 메시지 TextBlock (나중에 버튼 핸들러에서 참조)
        var txbMigResult = new TextBlock
        {
            Text       = "",
            FontSize   = 13,
            FontFamily = fontR,
            Foreground = Brush.Parse("#aaddaa"),
            Margin     = new Thickness(0, 12, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        };

        var btnExecute = new Button
        {
            Content    = "▶ 변경 실행",
            FontSize   = 14,
            FontFamily = font,
            FontWeight = FontWeight.Bold,
            Background = Brush.Parse("#2a4a2a"),
            Foreground = Brush.Parse("#88ff88"),
            BorderBrush = Brush.Parse("#448844"),
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
                txbMigResult.Foreground = Brush.Parse("#aaddaa");
                _dbMigrationPage?.LoadData();
                _dbMigrationTargetPage?.LoadData();
            }
            else
            {
                txbMigResult.Text       = $"❌ 오류: {error}";
                txbMigResult.Foreground = Brush.Parse("#ff8888");
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
                    FontSize   = 16, FontWeight = FontWeight.Bold,
                    FontFamily = font,
                    Foreground = Brush.Parse("#e0e0e0"),
                    Margin     = new Thickness(0, 0, 0, 8),
                },
                new TextBlock
                {
                    Text       = $"변경 전:  {oldName}",
                    FontSize   = 13, FontFamily = fontR,
                    Foreground = Brush.Parse("#ff9999"),
                },
                new TextBlock
                {
                    Text       = $"변경 후:  {newName}",
                    FontSize   = 13, FontFamily = fontR,
                    Foreground = Brush.Parse("#99ff99"),
                },
                new TextBlock
                {
                    Text       = "아래 테이블의 업체명이 일괄 변경됩니다:",
                    FontSize   = 12, FontFamily = fontR,
                    Foreground = Brush.Parse("#aaaaaa"),
                    Margin     = new Thickness(0, 10, 0, 0),
                },
                new TextBlock
                {
                    Text       = "  • 견적발행내역\n  • 거래명세서발행내역\n  • 계약 DB\n  • 분석의뢰및결과\n  • 시료명칭(컬럼명)",
                    FontSize   = 12, FontFamily = fontR,
                    Foreground = Brush.Parse("#cccccc"),
                    LineHeight = 20,
                },
                btnExecute,
                txbMigResult,
            },
        };

        return new Border
        {
            Background   = Brush.Parse("#1a1e2a"),
            BorderBrush  = Brush.Parse("#334466"),
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
        // TODO: 권한관리 페이지 구현
        Show1.Content = null;
        Show2.Content = null;
        Show4.Content = null;
        _bt1SaveAction = null;

        SetSubMenu("저장", "새로고침", "삭제", "설정", "통계", "종료", "도움말");
        SetContentLayout(content2Star: 1, content4Star: 0, upperStar: 1, lowerStar: 0);
    }


    // ══════════════════════════════════════════════════════════════════════
    //  서브메뉴 버튼
    // ══════════════════════════════════════════════════════════════════════

    private void SetSubMenu(string bt1, string bt2, string bt3,
                            string bt4, string bt5, string bt6,
                            string bt7 = "", string bt8 = "")
    {
        SetBtn(BT1, bt1); SetBtn(BT2, bt2); SetBtn(BT3, bt3);
        SetBtn(BT4, bt4); SetBtn(BT5, bt5);
        SetBtn(BT6, bt6); SetBtn(BT7, bt7); SetBtn(BT8, bt8);
        SubMenu.IsVisible = new[] { bt1, bt2, bt3, bt4, bt5, bt6, bt7, bt8 }
            .Any(s => !string.IsNullOrWhiteSpace(s));
    }

    private static void SetBtn(Avalonia.Controls.Button btn, string label)
    {
        btn.IsVisible = !string.IsNullOrWhiteSpace(label);
        btn.Content   = label;
    }

    private void BT1_Click(object? sender, RoutedEventArgs e)
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
                _wasteSampleListPage?.LoadData();
                _wasteSampleListPage?.LoadCompanyTree();
                break;
            case "MyTask":          _myTaskPage?.LoadData();             break;
            default: _bt1SaveAction?.Invoke();                          break;
        }
    }

    private void BT2_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":        _agentTreePage?.LoadData();      break;
            case "WasteCompany":      _wasteCompanyPage?.LoadData();    break;
            case "WasteDataQuery":    _wasteDataQueryPage?.LoadData(); break;
            case "WasteNameReconcile": _wasteNameReconcilePage?.Reload(); break;
            case "Schedule":     _schedulePage?.LoadData();        break;
            case "Contract":     _contractPage?.LoadData();       break;
            case "Purchase":     _purchasePage?.ExportCsv();      break;
            case "TestReport":   _testReportPage?.SaveCsv();      break;
            case "Quotation":
                // BT2 = 신규 작성 → Content2 를 NewPanel 로 교체
                _quotationNewPanel?.Clear();
                _quotationCheckPanel?.ClearAll();
                if (_quotationCheckPanel != null)
                {
                    _quotationCheckPanel.CurrentAnalysisRecord = null;
                    _quotationCheckPanel.CurrentIssue = null;
                }
                Show2.Content = _quotationNewPanel;
                break;
            case "WasteSampleList": _wasteSampleListPage?.AddNewDate(); break;
            case "Repair":       _repairPage?.ApproveSelected();  break;
            case "RiskManage":   _riskPage?.DeleteSelected();     break;
            default: Debug.WriteLine($"[{_currentMode}] BT2");   break;
        }
    }

    private void BT3_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "Agent":      _agentTreePage?.ShowAddPanel();           break;
            case "Contract":   _contractPage?.ShowAddPanel();            break;
            case "Purchase":   _purchasePage?.ApproveSelected();         break;
            case "Quotation":  _quotationPage?.LoadData(); _quotationHistoryPanel?.LoadData(); break;
            case "TestReport": _ = _testReportPage?.DeleteSampleAsync(); break;
            case "Repair":     _repairPage?.RejectSelected();            break;
            default: Debug.WriteLine($"[{_currentMode}] BT3");          break;
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
            case "Repair":   _repairPage?.CompleteSelected();   break;
            case "Quotation": await ExportQuotationAsync(); break;
            default: Debug.WriteLine($"[{_currentMode}] BT4"); break;
        }
    }

    private void BT5_Click(object? sender, RoutedEventArgs e)
    {
        switch (_currentMode)
        {
            case "WasteCompany": ShowWasteCompanyData();            break;
            case "Purchase":     _purchasePage?.CompleteSelected();  break;
            case "TestReport":   _testReportPage?.OpenPrintWindow();   break;
            case "Repair":       _repairPage?.DeleteSelected();      break;
            case "Contract":     _ = ImportContractFromExcelAsync(); break;
            default: Debug.WriteLine($"[{_currentMode}] BT5");      break;
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
            FontSize          = 13, Foreground = Res("FgMuted"),
            Margin            = new Thickness(12),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        var (ok, errCount) = await Task.Run(() => ContractService.ImportFromExcel(filePath));

        Show2.Content = new TextBlock
        {
            Text              = errCount < 0
                                    ? "❌  Excel 파일을 열 수 없습니다."
                                    : $"✅  가져오기 완료 — 성공 {ok}건{(errCount > 0 ? $", 오류 {errCount}건" : "")}\n트리에서 업체를 선택하면 단가가 표시됩니다.",
            FontSize          = 13,
            Foreground        = errCount < 0 ? Brushes.OrangeRed : new SolidColorBrush(Color.Parse("#16a34a")),
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
            case "Purchase":   _purchasePage?.DeleteSelected();        break;
            case "TestReport": _testReportPage?.BatchPrintExcel();     break;
            case "Quotation":       ShowTradeStatementEditor();    break;
            case "QuotationIssue":  IssueTradeStatementFromChecklist(); break;
            default: Debug.WriteLine($"[{_currentMode}] BT6");        break;
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
                FontFamily = _tradeFont, FontSize = 13,
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
            FontFamily = _tradeFont, FontSize = 13,
            Foreground = ok ? new SolidColorBrush(Color.Parse("#16a34a")) : Brushes.OrangeRed,
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
                FontSize = 13,
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
                        FontSize = 15,
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
                                FontSize = 12,
                                Foreground = Res("AppFg"),
                                FontFamily = _tradeFont,
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                            },
                        },
                    },
                    new TextBlock
                    {
                        Text = "그래도 발행하시겠습니까?",
                        FontSize = 13,
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
            FontSize = 13,
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
                FontSize = 13,
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
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.Parse("#16a34a")),
            Margin = new Thickness(16),
        };
    }

    // =========================================================================
    // 거래명세서 편집 패널 (BT6)
    // =========================================================================
    private static readonly FontFamily _tradeFont =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    private void ShowTradeStatementEditor()
    {
        var selected = _quotationHistoryPanel?.GetSelectedIssues()
                       ?? new List<ETA.Models.QuotationIssue>();

        if (selected.Count == 0)
        {
            Show4.Content = new TextBlock
            {
                Text = "⚠️  Show1에서 Ctrl+클릭으로\n견적서를 먼저 선택하세요.",
                FontSize = 12, Foreground = Brushes.DarkOrange, FontFamily = _tradeFont,
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
                    FontSize = 12, Foreground = Res("FgMuted"), FontFamily = _tradeFont,
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
                FontSize = 14, FontFamily = _tradeFont, FontWeight = FontWeight.SemiBold,
                Foreground = Res("AppFg"),
            });
            preview.Children.Add(new Border { Height = 1, Background = Res("InputBorder"), Margin = new Thickness(0,2,0,4) });
            preview.Children.Add(new TextBlock { Text = $"공급받는자 : {company}", FontSize = 11, FontFamily = _tradeFont, Foreground = Res("FgMuted") });
            preview.Children.Add(new TextBlock { Text = $"발행일     : {DateTime.Today:yyyy-MM-dd}", FontSize = 11, FontFamily = _tradeFont, Foreground = Res("FgMuted") });
            preview.Children.Add(new TextBlock { Text = $"건수       : {checked_.Count}건", FontSize = 11, FontFamily = _tradeFont, Foreground = Res("FgMuted") });

            // 견적번호 목록
            var noList = new TextBlock
            {
                Text = "견적번호 : " + string.Join(", ", checked_.Select(i => i.견적번호)),
                FontSize = 10, FontFamily = _tradeFont,
                Foreground = Res("FgMuted"),
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            };
            preview.Children.Add(noList);

            preview.Children.Add(new Border { Height = 1, Background = Res("InputBorder"), Margin = new Thickness(0,6,0,4) });

            // 금액 행
            void AddAmtRow(string label, decimal amt, bool bold = false)
            {
                var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
                row.Children.Add(new TextBlock { Text = label, Width = 90, FontSize = 12, FontFamily = _tradeFont, Foreground = Res("FgMuted") });
                var tb = new TextBlock
                {
                    Text = amt.ToString("N0") + " 원",
                    FontSize = 13, FontFamily = _tradeFont,
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
                FontSize = 13, FontFamily = _tradeFont, FontWeight = FontWeight.SemiBold,
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
                            FontSize = 11, FontFamily = _tradeFont, Foreground = Res("AppFg"),
                            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                        },
                        new TextBlock
                        {
                            Text = $"{issue.견적번호}  |  {issue.총금액:N0} 원",
                            FontSize = 10, FontFamily = _tradeFont,
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
            FontSize = 12, FontFamily = _tradeFont,
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
                FontSize = 13, Foreground = Brushes.Orange, FontFamily = _tradeFont,
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
            FontSize = 13, Foreground = Res("FgMuted"), FontFamily = _tradeFont,
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
                FontSize = 13, Foreground = new SolidColorBrush(Color.Parse("#16a34a")), FontFamily = _tradeFont,
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
                FontSize = 13, Foreground = Brushes.OrangeRed, FontFamily = _tradeFont,
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
            default:
                Debug.WriteLine($"[{_currentMode}] BT7");
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
            default:
                Debug.WriteLine($"[{_currentMode}] BT8");
                break;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  테마 (DynamicResource 전체 교체 방식)
    // ══════════════════════════════════════════════════════════════════════

    private void OnThemeChanged(object? sender, RoutedEventArgs e)
    {
        if (tglTheme == null) return;
        ApplyTheme(tglTheme.IsChecked == true);
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
        catch (Exception ex) { Debug.WriteLine($"[FontScale] 저장 실패: {ex.Message}"); }
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
        catch (Exception ex) { Debug.WriteLine($"[FontScale] 로드 실패: {ex.Message}"); }
        return 1.0;
    }

    private void ApplyTheme(bool isDark)
    {
        // ── Avalonia 기본 테마 변형도 함께 변경 ──────────────────────────
        var variant = isDark ? ThemeVariant.Dark : ThemeVariant.Light;
        if (Application.Current is Application app)
            app.RequestedThemeVariant = variant;
        this.RequestedThemeVariant = variant;

        // ── Window.Resources + Application.Current.Resources 동시 교체 ──
        // → 하위 페이지 코드에서 AppRes() 로 현재 테마 색상 조회 가능
        var palette = isDark ? DarkPalette() : LightPalette();
        foreach (var (key, color) in palette)
        {
            var brush = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse(color));
            this.Resources[key] = brush;
            if (Application.Current != null)
                Application.Current.Resources[key] = brush;
        }
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
        ["AppBg"]            = "#1e1e26",
        ["PanelBg"]          = "#39383f",
        ["PanelInnerBg"]     = "#2d2d35",
        ["MenuBarBg"]        = "#444444",
        ["MenuItemBg"]       = "#444444",
        ["MenuItemHover"]    = "#012800",
        ["SubMenuBg"]        = "#012800",
        ["SubMenuItemBg"]    = "#444444",   // ★ 서브메뉴 항목 통일
        ["SubMenuItemHover"] = "#023d00",
        ["SubBtnBg"]         = "#342f2f",
        ["SubBtnHover"]      = "#012800",
        ["SplitterColor"]    = "#020202",
        ["AppFg"]            = "#ffffff",
        ["FgMuted"]          = "#aaaaaa",
        ["FgHover"]          = "#fd0f0f",
        ["TreeFg"]           = "#e0e0e0",
        ["TreeSelBg"]        = "#014a00",
        ["TreeSelFg"]        = "#ffffff",
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
    };

    // ── 라이트 팔레트 ─────────────────────────────────────────────────────
    private static System.Collections.Generic.Dictionary<string, string> LightPalette() => new()
    {
        ["AppBg"]            = "#f0f2f5",
        ["PanelBg"]          = "#ffffff",
        ["PanelInnerBg"]     = "#f8f9fb",
        ["MenuBarBg"]        = "#2d6a4f",
        ["MenuItemBg"]       = "#2d6a4f",
        ["MenuItemHover"]    = "#1b4332",
        ["SubMenuBg"]        = "#1b4332",
        ["SubMenuItemBg"]    = "#2d6a4f",   // ★ 라이트도 통일
        ["SubMenuItemHover"] = "#145a32",
        ["SubBtnBg"]         = "#e2e8f0",
        ["SubBtnHover"]      = "#cbd5e1",
        ["SplitterColor"]    = "#cbd5e1",
        ["AppFg"]            = "#1a1a2e",
        ["FgMuted"]          = "#64748b",
        ["FgHover"]          = "#dc2626",
        ["TreeFg"]           = "#1e293b",
        ["TreeSelBg"]        = "#bbf7d0",
        ["TreeSelFg"]        = "#14532d",
        ["GridHeaderBg"]     = "#e2e8f0",
        ["GridRowBg"]        = "#ffffff",
        ["GridRowAltBg"]     = "#f8fafc",
        ["InputBg"]          = "#ffffff",
        ["InputBorder"]      = "#94a3b8",
        ["InputFg"]          = "#1e293b",
        ["TreeHoverBg"]      = "#e2e8f0",
        ["Panel4Bg"]         = "#f0f4f8",
        ["TabActiveBg"]      = "#dbeafe",
        ["TabActiveFg"]      = "#1e40af",
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
        => _quotationNewPanel?.SetCompany(company);

    private void LogContentChange(string contentName, object? content)
    {
        Debug.WriteLine($"[{DateTime.Now:HH:mm:ss}] ContentChange: {contentName} = {content?.GetType().Name ?? "null"}");
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
            return;

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

            // 오른쪽 상단/하단 분할 비율
            var rightGrid = this.FindControl<Grid>("RightSplitGrid");
            if (rightGrid != null && rightGrid.RowDefinitions.Count >= 3)
            {
                layout.UpperStar = rightGrid.RowDefinitions[0].Height.Value;
                layout.LowerStar = rightGrid.RowDefinitions[2].Height.Value;
            }

            // Content2/Content4 분할 비율
            var upperGrid = this.FindControl<Grid>("UpperContentGrid");
            if (upperGrid != null && upperGrid.ColumnDefinitions.Count >= 3)
            {
                layout.Content2Star = upperGrid.ColumnDefinitions[0].Width.Value;
                layout.Content4Star = upperGrid.ColumnDefinitions[2].Width.Value;
            }

            layout.SavedAt = DateTime.Now;
            _positionManager.SavePageLayout(modeKey, layout);
            System.Diagnostics.Debug.WriteLine($"[MainPage] 저장: {modeKey} - {layout}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MainPage] 레이아웃 저장 오류: {ex.Message}");
        }
    }

    /// <summary>
    /// 특정 모드의 저장된 레이아웃 정보 복원
    /// </summary>
    private void RestoreModeLayout(string modeName, double minLowerStar = 0)
    {
        if (_positionManager == null)
            return;

        try
        {
            string modeKey = LayoutStorageModePrefix + modeName;
            var layout = _positionManager.GetPageLayout(modeKey);
            if (layout == null)
            {
                System.Diagnostics.Debug.WriteLine($"[MainPage] 복원할 레이아웃 없음: {modeKey}");
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

            System.Diagnostics.Debug.WriteLine($"[MainPage] 복원: {modeKey} - {layout}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MainPage] 레이아웃 복원 오류: {ex.Message}");
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
            // 새 사용자로 WindowPositionManager 재초기화
            _positionManager = new WindowPositionManager(newUserId);
            System.Diagnostics.Debug.WriteLine($"[MainPage] 사용자 변경: {newUserId}");
        }
    }

    /// <summary>
    /// 로그 파일 경로 조회 (디버깅용)
    /// </summary>
    public string GetPositionLogFilePath()
    {
        return _positionManager?.GetLogFilePath() ?? "Unknown";
    }

    private void TEST_Click(object? sender, RoutedEventArgs e) { }
}
