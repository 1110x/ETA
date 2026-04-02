using Avalonia;
using Avalonia.Controls;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE1;

public partial class AgentTreePage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;
    public Action<Control?>? Show3ContentRequest { get; set; }
    public Action<Control?>? Show4ContentRequest { get; set; }
    public ListBox? AnalysisItemsListBox { get; set; }

    // ── 상태 ────────────────────────────────────────────────────────────────
    private Agent?      _selectedAgent;
    private StackPanel? _detailPanel;
    private bool        _isAddMode  = false;

    private static bool CanEdit =>
        MainPage.CurrentEmployeeId == "201000308";

    // 사진 미리보기 Image 컨트롤 (저장 시 PhotoPath 접근용)
    private Image?      _photoImage;
    private string      _pendingPhotoPath = "";   // 선택했지만 아직 저장 안 된 경로

    // 업무 분장 저장 시 참조 (BuildAssignmentArea에서 설정)
    private ListBox?      _assignmentListBox;
    private Canvas?       _timelineCanvas;
    private ScrollViewer? _timelineScroll;
    private DateTime    _assignmentRangeStart = DateTime.Today;
    private DateTime    _assignmentRangeEnd   = DateTime.Today;

    // 담당항목/업체/기타업무 패널 (Show2)
    private WrapPanel?  _항목ChipsPanel;
    private WrapPanel?  _업체ChipsPanel;
    private StackPanel? _miscTaskPanel;

    // ── 업무분장표 상태 ─────────────────────────────────────────────────────
    private DateTime _chartRangeStart;
    private DateTime _chartRangeEnd;
    private Canvas?  _chartCanvas;

    // ── 타임라인 상수 ───────────────────────────────────────────────────────
    private const double TL_DAY_W      = 14.0;
    private const double TL_ROW_H      = 22.0;
    private const double TL_LABEL_W    = 110.0;
    private const double TL_HEADER_H   = 18.0;
    private const double TL_MILESTONE_H = 60.0;

    // ── 업무분장표 상수 ─────────────────────────────────────────────────────
    private const double GC_LABEL_W   = 56.0;
    private const double GC_DAY_W     = 22.0;
    private const double GC_ROW_H     = 28.0;   // 컴팩트 행
    private const double GC_BAR_H     = 18.0;   // 바 높이
    private const double GC_DATE_H    = 20.0;   // 하단 날짜 영역
    private const double GC_HANDLE_R  = 5.0;    // 드래그 핸들 반지름
    private const double GC_VBRANCH_H = 22.0;   // V분기 높이
    private const double GC_VBRANCH_W = 36.0;   // V분기 퍼짐

    public AgentTreePage()
    {
        InitializeComponent();
    }

    // =========================================================================
    // 데이터 로드
    // =========================================================================
    public void LoadData()
    {
        Log("LoadData() 시작");
        AgentTreeView.Items.Clear();
        _selectedAgent        = null;
        _isAddMode            = false;
        _pendingPhotoPath     = "";
        _assignmentListBox    = null;
        _assignmentRangeStart = DateTime.Today;
        _assignmentRangeEnd   = DateTime.Today;
        DetailPanelChanged?.Invoke(null);

        try
        {
            var items = AgentService.GetAllItems().OrderBy(a => a.입사일).ToList();
            foreach (var item in items)
                AgentTreeView.Items.Add(CreateTreeItem(item));
            Log($"로드 완료 → {items.Count}명");
        }
        catch (Exception ex) { Log("★ 크래시 ★ " + ex.Message); }
    }

    // =========================================================================
    // TreeViewItem 생성
    // =========================================================================
    private TreeViewItem CreateTreeItem(Agent agent)
    {
        var headerPanel = new StackPanel
        {
            Orientation       = Orientation.Horizontal,
            Spacing           = 9,
            VerticalAlignment = VerticalAlignment.Center,
        };

        // 원형 사진 (또는 이니셜 원)
        headerPanel.Children.Add(MakePhotoCircle(agent, 30));

        // 직급 배지 — 이름 앞에 배치
        if (!string.IsNullOrWhiteSpace(agent.직급))
        {
            var (bg, fg) = BadgeColorHelper.GetBadgeColor(agent.직급);
            headerPanel.Children.Add(new Border
            {
                Background        = new SolidColorBrush(Color.Parse(bg)),
                CornerRadius      = new CornerRadius(3),
                Padding           = new Thickness(5, 2),
                VerticalAlignment = VerticalAlignment.Center,
                Child             = new TextBlock
                {
                    Text              = agent.직급,
                    FontSize          = 10,
                    Foreground        = new SolidColorBrush(Color.Parse(fg)),
                    FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    VerticalAlignment = VerticalAlignment.Center,
                }
            });
        }

        // 이름
        headerPanel.Children.Add(new TextBlock
        {
            Text              = agent.성명,
            FontSize          = 13,
            FontFamily        = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Foreground        = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        return new TreeViewItem
        {
            Tag    = agent,
            Header = headerPanel,
        };
    }

    // 원형 사진 또는 이니셜 원 헬퍼
    private static Control MakePhotoCircle(Agent agent, double size)
    {
        var photoDir = AgentService.GetPhotoDirectory();
        if (!string.IsNullOrEmpty(agent.PhotoPath))
        {
            // 로컬 파일 없으면 DB에서 가져와 캐시
            AgentService.EnsurePhotoLocal(agent.사번, agent.PhotoPath);

            var fullPath = Path.IsPathRooted(agent.PhotoPath)
                ? agent.PhotoPath
                : Path.Combine(photoDir, agent.PhotoPath);
            if (File.Exists(fullPath))
            {
                try
                {
                    using var stream = File.OpenRead(fullPath);
                    var bmp = new Bitmap(stream);
                    return new Ellipse
                    {
                        Width  = size*0.7,
                        Height = size*0.7,
                        Fill   = new ImageBrush(bmp) { Stretch = Stretch.UniformToFill },
                    };
                }
                catch { }
            }
        }

        // 이니셜 원
        var initial = agent.성명.Length > 0 ? agent.성명[0].ToString() : "?";
        var key     = !string.IsNullOrWhiteSpace(agent.직급) ? agent.직급 : agent.성명;
        var (ibg, ifg) = BadgeColorHelper.GetBadgeColor(key);
        return new Border
        {
            Width        = size,
            Height       = size,
            CornerRadius = new CornerRadius(size / 2),
            Background   = new SolidColorBrush(Color.Parse(ibg)),
            ClipToBounds = true,
            Child        = new TextBlock
            {
                Text                = initial,
                FontSize            = size * 0.45,
                Foreground          = new SolidColorBrush(Color.Parse(ifg)),
                FontFamily          = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            }
        };
    }

    // =========================================================================
    // 트리 선택 → 수정 모드
    // =========================================================================
    public void AgentTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        Agent? agent = null;
        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is Agent a1) agent = a1;
            else if (e.AddedItems[0] is Agent a2) agent = a2;
        }
        if (agent == null) return;

        _selectedAgent    = agent;
        _isAddMode        = false;
        _pendingPhotoPath = "";
        _detailPanel      = BuildEditPanel(agent);
        DetailPanelChanged?.Invoke(_detailPanel);
        Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
        Log($"선택: {agent.성명}");
    }

    // =========================================================================
    // 직원 추가 패널  (MainPage BT3)
    // =========================================================================
    public void ShowAddPanel()
    {
        if (!CanEdit) return;
        _selectedAgent           = null;
        _isAddMode               = true;
        _pendingPhotoPath        = "";
        AgentTreeView.SelectedItem = null;
        _detailPanel             = BuildAddPanel();
        DetailPanelChanged?.Invoke(_detailPanel);
        Log("추가 모드");
    }

    // =========================================================================
    // 저장  (MainPage BT1)
    // =========================================================================
    public void SaveSelected()
    {
        if (!CanEdit) return;
        if (_isAddMode) SaveAdd();
        else            SaveEdit();
    }

    // =========================================================================
    // 삭제  (MainPage BT4)
    // =========================================================================
    public async Task DeleteSelectedAsync()
    {
        if (!CanEdit) return;
        if (_selectedAgent == null)
        {
            Log("삭제 스킵: 선택 없음");
            return;
        }

        // 확인 다이얼로그
        var dlg = new Window
        {
            Title           = "삭제 확인",
            Width           = 320,
            Height          = 150,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize       = false,
            Background      = new SolidColorBrush(Color.Parse("#2d2d2d")),
        };

        bool confirmed = false;

        var yesBtn = new Button
        {
            Content    = "삭제",
            Width      = 80,
            Background = new SolidColorBrush(Color.Parse("#c0392b")),
            Foreground = Brushes.White,
        };
        var noBtn = new Button
        {
            Content    = "취소",
            Width      = 80,
            Background = new SolidColorBrush(Color.Parse("#444")),
            Foreground = Brushes.White,
        };

        yesBtn.Click += (_, _) => { confirmed = true;  dlg.Close(); };
        noBtn.Click  += (_, _) => { confirmed = false; dlg.Close(); };

        dlg.Content = new StackPanel
        {
            Margin  = new Thickness(20),
            Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text       = $"'{_selectedAgent.성명}' 직원을 삭제하시겠습니까?",
                    Foreground = AppRes("AppFg"),
                    FontSize   = 13,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap
                },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing     = 12,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Children    = { yesBtn, noBtn }
                }
            }
        };

        // 부모 Window 찾기
        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null)
            await dlg.ShowDialog(owner);
        else
            dlg.Show();

        if (!confirmed) return;

        bool ok = AgentService.Delete(_selectedAgent);
        Log(ok ? $"✅ 삭제 성공: {_selectedAgent.성명}" : $"❌ 삭제 실패: {_selectedAgent.성명}");

        if (ok)
        {
            // 트리에서 해당 항목 제거
            var toRemove = AgentTreeView.Items
                .OfType<TreeViewItem>()
                .FirstOrDefault(i => i.Tag == _selectedAgent);
            if (toRemove != null) AgentTreeView.Items.Remove(toRemove);

            _selectedAgent    = null;
            _detailPanel      = null;
            _pendingPhotoPath = "";
            DetailPanelChanged?.Invoke(null);
        }
    }

    // =========================================================================
    // 패널 — 수정 모드
    // =========================================================================
    private StackPanel BuildEditPanel(Agent agent)
    {
        bool ro = !CanEdit;
        var root = MakeRootPanel(ro
            ? $"👁  {agent.성명} — 정보 조회 (읽기 전용)"
            : $"✏️  {agent.성명} — 정보 수정");

        // 사진 + 성명/직급/사번 (상단)
        var topRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, Margin = new Thickness(0, 0, 0, 4) };
        topRow.Children.Add(BuildPhotoArea(agent.PhotoPath, agent.사번));
        var topInfo = new StackPanel { Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
        topInfo.Children.Add(BuildFieldRow("성명", agent.성명, isReadOnly: true, isLocked: true));
        topInfo.Children.Add(BuildFieldRow("직급", agent.직급, isReadOnly: ro));
        topInfo.Children.Add(BuildFieldRow("사번", agent.사번, isReadOnly: ro));
        topRow.Children.Add(topInfo);
        root.Children.Add(topRow);

        // 기본정보 2열 배치
        root.Children.Add(BuildFieldGrid(
            BuildFieldRow("직무",           agent.직무,           isReadOnly: ro),
            BuildFieldRow("입사일",         agent.입사일표시,     isReadOnly: true),
            BuildFieldRow("자격사항",       agent.자격사항,       isReadOnly: ro),
            BuildFieldRow("Email",          agent.Email,          isReadOnly: ro),
            BuildFieldRow("측정인고유번호", agent.측정인고유번호, isReadOnly: ro)
        ));

        // ── 담당 분석항목 ──
        var itemSection = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#1a1a2a")),
            CornerRadius = new CornerRadius(4),
            Padding      = new Thickness(8, 6),
            Margin       = new Thickness(0, 6, 0, 0),
        };
        var itemInner = new StackPanel { Spacing = 3 };
        itemInner.Children.Add(BuildChipSectionHeader("🧪 담당 분석항목",
            () => Show4ContentRequest?.Invoke(BuildShow4AnalyteList(agent)),
            date => Show4ContentRequest?.Invoke(BuildAssignWithAttachPanel(agent, date))));
        _항목ChipsPanel = new WrapPanel { Orientation = Orientation.Horizontal };
        RefreshItemChips(agent);
        itemInner.Children.Add(_항목ChipsPanel);
        itemSection.Child = itemInner;
        root.Children.Add(itemSection);

        // ── 담당 계약업체 ──
        var contractSection = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#1a2a1a")),
            CornerRadius = new CornerRadius(4),
            Padding      = new Thickness(8, 6),
            Margin       = new Thickness(0, 6, 0, 0),
        };
        var contractInner = new StackPanel { Spacing = 3 };
        contractInner.Children.Add(BuildChipSectionHeader("🏢 담당 계약업체",
            () => Show4ContentRequest?.Invoke(BuildShow4ContractList(agent))));
        _업체ChipsPanel = new WrapPanel { Orientation = Orientation.Horizontal };
        RefreshContractChips(agent);
        contractInner.Children.Add(_업체ChipsPanel);
        contractSection.Child = contractInner;
        root.Children.Add(contractSection);

        // ── 기타업무 ──
        var taskSection = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#2a1a2a")),
            CornerRadius = new CornerRadius(4),
            Padding      = new Thickness(8, 6),
            Margin       = new Thickness(0, 6, 0, 0),
        };
        var taskInner = new StackPanel { Spacing = 3 };
        _miscTaskPanel = new StackPanel { Spacing = 3 };
        taskInner.Children.Add(BuildChipSectionHeader("📋 기타업무",
            () => Show3ContentRequest?.Invoke(BuildMiscTaskForm(agent))));
        taskInner.Children.Add(_miscTaskPanel);
        taskSection.Child = taskInner;
        root.Children.Add(taskSection);
        RefreshMiscTasks(agent);

        // ── 과거 분장이력 ──
        root.Children.Add(BuildHistorySection(agent));

        return root;
    }

    // =========================================================================
    // 패널 — 추가 모드
    // =========================================================================
    private StackPanel BuildAddPanel()
    {
        var root = MakeRootPanel("➕  신규 직원 추가");

        var topRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, Margin = new Thickness(0, 0, 0, 4) };
        topRow.Children.Add(BuildPhotoArea(""));
        var topInfo = new StackPanel { Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
        topInfo.Children.Add(BuildFieldRow("성명", "", hint: "이름 입력 (필수)"));
        topInfo.Children.Add(BuildFieldRow("직급", ""));
        topInfo.Children.Add(BuildFieldRow("사번", ""));
        topRow.Children.Add(topInfo);
        root.Children.Add(topRow);

        root.Children.Add(BuildFieldGrid(
            BuildFieldRow("직무",           ""),
            BuildFieldRow("입사일",         "", hint: "예) 2024-01-01"),
            BuildFieldRow("자격사항",       ""),
            BuildFieldRow("Email",          ""),
            BuildFieldRow("측정인고유번호", "")
        ));

        return root;
    }

    // =========================================================================
    // 분장 영역 빌드 (드래그 앤 드랍 + 월 조회)
    // =========================================================================
    private Control BuildAssignmentArea(Agent agent)
    {
        var border = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#1a2a1a")),
            CornerRadius = new CornerRadius(6),
            Padding      = new Thickness(10, 8),
            Margin       = new Thickness(0, 8, 0, 0)
        };

        var stack = new StackPanel { Spacing = 4 };

        // ── 제목 ──────────────────────────────────────────────────────────
        stack.Children.Add(new TextBlock
        {
            Text       = "📋 업무 분장",
            FontSize   = 12,
            FontWeight = FontWeight.Bold,
            Foreground = new SolidColorBrush(Color.Parse("#7cd87c")),
            FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"
        });

        // ── 날짜 범위 표시 + 달력 버튼 ──────────────────────────────────
        var dateRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 6,
            Margin      = new Thickness(0, 4, 0, 0)
        };

        var txbDateRange = new TextBlock
        {
            Text              = DateTime.Today.ToString("yyyy-MM-dd"),
            FontSize          = 11,
            Foreground        = new SolidColorBrush(Color.Parse("#aaa")),
            VerticalAlignment = VerticalAlignment.Center,
            MinWidth          = 170,
            FontFamily        = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"
        };

        var btnCal = new Button
        {
            Content         = "📅",
            Width           = 30,
            Height          = 24,
            FontSize        = 12,
            Padding         = new Thickness(0),
            Background      = new SolidColorBrush(Color.Parse("#2a3a4a")),
            Foreground      = new SolidColorBrush(Color.Parse("#aaa")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555")),
        };
        ToolTip.SetTip(btnCal, "기간 선택 (드래그로 범위 설정)");

        var btnToday = new Button
        {
            Content         = "오늘",
            Width           = 44,
            Height          = 24,
            FontSize        = 10,
            Padding         = new Thickness(4, 0),
            Background      = new SolidColorBrush(Color.Parse("#3a5a3a")),
            Foreground      = new SolidColorBrush(Color.Parse("#aaa")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#666"))
        };

        dateRow.Children.Add(txbDateRange);
        dateRow.Children.Add(btnCal);
        dateRow.Children.Add(btnToday);
        stack.Children.Add(dateRow);

        // ── 달력 (inline, 토글) ──────────────────────────────────────────
        var calendar = new Avalonia.Controls.Calendar
        {
            SelectionMode = Avalonia.Controls.CalendarSelectionMode.SingleRange,
            IsVisible     = false,
            Margin        = new Thickness(0, 2, 0, 0),
            DisplayDate   = DateTime.Today,
        };
        calendar.SelectedDates.Add(DateTime.Today);
        stack.Children.Add(calendar);

        // ── 타임라인 캔버스 ───────────────────────────────────────────────
        var timelineCanvas = new Canvas { Width = 300, Height = 60 };
        _timelineCanvas = timelineCanvas;

        var timelineScroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Height  = 160,
            Margin  = new Thickness(0, 4, 0, 0),
            Content = timelineCanvas,
        };
        _timelineScroll = timelineScroll;
        DragDrop.SetAllowDrop(timelineScroll, true);
        _assignmentListBox = null; // no longer used

        // 날짜 범위 상태 (클로저로 공유)
        DateTime rangeStart = DateTime.Today;
        DateTime rangeEnd   = DateTime.Today;

        // 드래그 드랍 — 중복 체크 포함 async 핸들러
        timelineScroll.AddHandler(DragDrop.DropEvent, async (object? sender, DragEventArgs e) =>
        {
            if (!e.Data.Contains("analyte")) return;

            var analyte = e.Data.Get("analyte") as string;
            if (string.IsNullOrEmpty(analyte)) return;

            // 전체 fullName 확인 (analyte는 item.Analyte = shortName일 수 있음)
            var existing = AnalysisRequestService.GetAssigneesForAnalyteOnDate(analyte, rangeStart);

            if (existing.Count > 0 && !existing.Contains(agent.성명))
            {
                // 중복 경고 다이얼로그
                bool doUpdate = await ShowDuplicateWarningAsync(analyte, existing);
                if (!doUpdate) { Log($"드랍 취소: {analyte} → 중복"); return; }
            }

            AnalysisRequestService.AddAssignment(
                agent.사번, analyte, rangeStart, rangeStart.AddMonths(1));
            RefreshTimeline(timelineCanvas, agent, rangeStart, rangeEnd);
            Log($"분장 추가: {agent.성명} ← {analyte}");
        });

        RefreshTimeline(timelineCanvas, agent, rangeStart, rangeEnd);
        stack.Children.Add(timelineScroll);

        // ── 저장 프로그래스 바 (저장 중에만 표시) ───────────────────────
        var saveProgress = new ProgressBar
        {
            IsIndeterminate = true,
            IsVisible       = false,
            Height          = 4,
            Margin          = new Thickness(0, 6, 0, 0),
            Foreground      = new SolidColorBrush(Color.Parse("#7cd87c")),
            Background      = new SolidColorBrush(Color.Parse("#1a2a1a")),
        };
        stack.Children.Add(saveProgress);

        // ── 분장 저장 버튼 ───────────────────────────────────────────────
        var btnAssignSave = new Button
        {
            Content             = "💾 분장 저장",
            Height              = 28,
            FontSize            = 11,
            FontFamily          = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Background          = new SolidColorBrush(Color.Parse("#2a5a2a")),
            Foreground          = AppRes("AppFg"),
            BorderThickness     = new Thickness(1),
            BorderBrush         = new SolidColorBrush(Color.Parse("#4a8a4a")),
            CornerRadius        = new CornerRadius(4),
            Padding             = new Thickness(10, 0),
            Margin              = new Thickness(0, 4, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        stack.Children.Add(btnAssignSave);

        // ── 이벤트 ───────────────────────────────────────────────────────
        btnCal.Click += (_, _) =>
            calendar.IsVisible = !calendar.IsVisible;

        btnToday.Click += (_, _) =>
        {
            rangeStart = rangeEnd = DateTime.Today;
            _assignmentRangeStart = _assignmentRangeEnd = DateTime.Today;
            txbDateRange.Text  = rangeStart.ToString("yyyy-MM-dd");
            calendar.IsVisible = false;
            RefreshTimeline(timelineCanvas, agent, rangeStart, rangeEnd);
            Log($"오늘 조회: {rangeStart:yyyy-MM-dd}");
        };

        calendar.SelectedDatesChanged += (_, _) =>
        {
            if (calendar.SelectedDates.Count == 0) return;
            var dates = calendar.SelectedDates.Cast<DateTime>().ToList();
            rangeStart = dates.Min();
            rangeEnd   = dates.Max();
            _assignmentRangeStart = rangeStart;
            _assignmentRangeEnd   = rangeEnd;
            txbDateRange.Text = rangeStart == rangeEnd
                ? rangeStart.ToString("yyyy-MM-dd")
                : $"{rangeStart:yyyy-MM-dd} ~ {rangeEnd:yyyy-MM-dd}";
            RefreshTimeline(timelineCanvas, agent, rangeStart, rangeEnd);
            // 시작·종료 날짜가 다르면 범위 선택 완료 → 달력 자동 닫기
            if (rangeStart != rangeEnd)
                calendar.IsVisible = false;
            Log($"기간 조회: {rangeStart:yyyy-MM-dd} ~ {rangeEnd:yyyy-MM-dd}");
        };

        btnAssignSave.Click += async (_, _) =>
        {
            saveProgress.IsVisible  = true;
            btnAssignSave.IsEnabled = false;

            // Assignments are already saved via drag-drop; refresh timeline
            await Task.Run(() => { /* no-op: assignments saved on drop */ });

            RefreshTimeline(timelineCanvas, agent, rangeStart, rangeEnd);
            saveProgress.IsVisible  = false;
            btnAssignSave.IsEnabled = true;
            Log($"분장 새로고침 완료 ({rangeStart:yyyy-MM-dd} ~ {rangeEnd:yyyy-MM-dd})");
        };

        border.Child = stack;
        return border;
    }

    // 중복 경고 다이얼로그
    private async Task<bool> ShowDuplicateWarningAsync(string analyte, List<string> existingAssignees)
    {
        var names = string.Join(", ", existingAssignees);
        var dlg = new Window
        {
            Title                 = "중복 할당 경고",
            Width                 = 360,
            Height                = 160,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize             = false,
            Background            = new SolidColorBrush(Color.Parse("#2d2d2d")),
        };

        bool result = false;
        var yesBtn = new Button
        {
            Content    = "업데이트",
            Width      = 90,
            Background = new SolidColorBrush(Color.Parse("#3a5a2a")),
            Foreground = Brushes.White,
        };
        var noBtn = new Button
        {
            Content    = "취소",
            Width      = 70,
            Background = new SolidColorBrush(Color.Parse("#444")),
            Foreground = Brushes.White,
        };
        yesBtn.Click += (_, _) => { result = true;  dlg.Close(); };
        noBtn.Click  += (_, _) => { result = false; dlg.Close(); };

        dlg.Content = new StackPanel
        {
            Margin  = new Thickness(20),
            Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text         = $"⚠️ '{analyte}'은(는) 이미 {names}에게 할당되어 있습니다.\n업데이트 하시겠습니까?",
                    Foreground   = AppRes("AppFg"),
                    FontSize     = 12,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    FontFamily   = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
                },
                new StackPanel
                {
                    Orientation         = Orientation.Horizontal,
                    Spacing             = 12,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Children            = { yesBtn, noBtn }
                }
            }
        };

        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();
        return result;
    }

    private void LoadAssignments(ListBox listBox, Agent agent, DateTime start, DateTime end)
    {
        listBox.Items.Clear();
        var assignments = start == end
            ? AnalysisRequestService.GetAssignmentsForAgent(agent.사번, start)
            : AnalysisRequestService.GetAssignmentsForAgentRange(agent.사번, start, end);
        RenderAssignmentItems(listBox, assignments);
    }

    private void LoadAssignments(ListBox listBox, Agent agent, DateTime queryDate)
    {
        listBox.Items.Clear();
        var assignments = AnalysisRequestService.GetAssignmentsForAgent(agent.사번, queryDate);
        RenderAssignmentItems(listBox, assignments);
    }

    private void RenderAssignmentItems(ListBox listBox, List<(string FullName, string ShortName)> assignments)
    {
        listBox.Items.Clear();

        if (assignments.Count == 0)
        {
            listBox.Items.Add(new ListBoxItem
            {
                Content   = "⏳ 할당된 항목 없음",
                IsEnabled = false,
                Foreground = new SolidColorBrush(Color.Parse("#666666"))
            });
            return;
        }

        foreach (var (fullName, shortName) in assignments)
        {
            var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
            topRow.Children.Add(new Border
            {
                Background    = new SolidColorBrush(Color.Parse("#1a2a3a")),
                CornerRadius  = new CornerRadius(3),
                Padding       = new Thickness(5, 1),
                Margin        = new Thickness(0, 0, 6, 0),
                [Grid.ColumnProperty] = 0,
                Child = new TextBlock
                {
                    Text       = shortName,
                    FontSize   = 9,
                    Foreground = new SolidColorBrush(Color.Parse("#7ab4cc")),
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                },
            });
            topRow.Children.Add(new TextBlock
            {
                Text              = fullName,
                FontSize          = 11,
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                Foreground        = new SolidColorBrush(Color.Parse("#dddddd")),
                TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 1,
            });

            listBox.Items.Add(new ListBoxItem
            {
                Content = topRow,
                Tag     = fullName,
                Padding = new Thickness(4, 3),
            });
        }
    }
    private void RefreshTimeline(Canvas canvas, Agent agent, DateTime start, DateTime end)
    {
        canvas.Children.Clear();

        var calendar = AnalysisRequestService.GetAssignmentCalendar(agent.사번, start, end);
        if (calendar.Count == 0)
        {
            var noItem = new TextBlock
            {
                Text       = "할당된 항목 없음",
                FontSize   = 11,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#666")),
            };
            Canvas.SetLeft(noItem, 8.0);
            Canvas.SetTop(noItem, 8.0);
            canvas.Children.Add(noItem);
            canvas.Width  = 300;
            canvas.Height = 40;
            return;
        }

        int numDays    = (int)(end - start).TotalDays + 1;
        double canvasW = TL_LABEL_W + numDays * TL_DAY_W + 4;
        double canvasH = TL_HEADER_H + calendar.Count * TL_ROW_H + 4;
        canvas.Width  = canvasW;
        canvas.Height = canvasH;

        // ── 날짜 헤더 ─────────────────────────────────────────────────────
        for (int d = 0; d < numDays; d++)
        {
            var dt         = start.AddDays(d);
            bool isMonStart = dt.Day == 1;
            bool isSun      = dt.DayOfWeek == DayOfWeek.Sunday;
            bool isSat      = dt.DayOfWeek == DayOfWeek.Saturday;
            string label    = isMonStart ? $"{dt.Month}/{dt.Day}" : (dt.Day % 5 == 0 ? dt.Day.ToString() : "");
            if (!string.IsNullOrEmpty(label) || isMonStart)
            {
                var lbl = new TextBlock
                {
                    Text       = label,
                    FontSize   = 8,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Foreground = new SolidColorBrush(Color.Parse(isMonStart ? "#aaccff" : "#888")),
                };
                Canvas.SetLeft(lbl, TL_LABEL_W + d * TL_DAY_W);
                Canvas.SetTop(lbl, 1.0);
                canvas.Children.Add(lbl);
            }
            if (isSat || isSun)
            {
                var tint = new Avalonia.Controls.Shapes.Rectangle
                {
                    Width  = TL_DAY_W,
                    Height = canvasH - TL_HEADER_H,
                    Fill   = new SolidColorBrush(Color.Parse(isSun ? "#1a1a10" : "#151520")),
                };
                Canvas.SetLeft(tint, TL_LABEL_W + d * TL_DAY_W);
                Canvas.SetTop(tint, TL_HEADER_H);
                canvas.Children.Add(tint);
            }
        }

        // ── 항목 행 ────────────────────────────────────────────────────────
        int row = 0;
        foreach (var entry in calendar.OrderBy(x => x.FullName))
        {
            double y = TL_HEADER_H + row * TL_ROW_H;

            // Row separator
            var sep = new Avalonia.Controls.Shapes.Rectangle
            {
                Width  = canvasW,
                Height = 1,
                Fill   = new SolidColorBrush(Color.Parse("#2a2a3a")),
            };
            Canvas.SetLeft(sep, 0.0);
            Canvas.SetTop(sep, y);
            canvas.Children.Add(sep);

            // Label — 약칭 기반 색상
            var (catBg, catFg) = BadgeColorHelper.GetBadgeColor(entry.ShortName);
            var labelBorder = new Border
            {
                Width        = TL_LABEL_W - 4,
                Height       = TL_ROW_H - 2,
                Background   = new SolidColorBrush(Color.Parse(catBg)),
                CornerRadius = new CornerRadius(2),
                Padding      = new Thickness(3, 1),
                Child        = new TextBlock
                {
                    Text              = entry.ShortName,
                    FontSize          = 9,
                    FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground        = new SolidColorBrush(Color.Parse(catFg)),
                    TextTrimming      = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                },
            };
            Canvas.SetLeft(labelBorder, 2.0);
            Canvas.SetTop(labelBorder, y + 1);
            canvas.Children.Add(labelBorder);

            // Day bars
            var dateSet = new HashSet<DateTime>(entry.Dates);
            for (int d = 0; d < numDays; d++)
            {
                if (!dateSet.Contains(start.AddDays(d))) continue;
                var bar = new Avalonia.Controls.Shapes.Rectangle
                {
                    Width   = TL_DAY_W - 1,
                    Height  = TL_ROW_H - 4,
                    Fill    = new SolidColorBrush(Color.Parse(catBg)),
                    Opacity = 0.7,
                    RadiusX = 1,
                    RadiusY = 1,
                };
                Canvas.SetLeft(bar, TL_LABEL_W + d * TL_DAY_W);
                Canvas.SetTop(bar, y + 2);
                canvas.Children.Add(bar);
            }
            row++;
        }
    }

    // Category color helper (local, for timeline)
    private static (string Bg, string Fg) GetCategoryColor(string category)
    {
        return category.Trim() switch
        {
            var c when c.Contains("유기")                       => ("#1a2a3a", "#88aaff"),
            var c when c.Contains("무기")                       => ("#2a1a3a", "#cc88ff"),
            var c when c.Contains("부유")                       => ("#1a3a2a", "#88ccaa"),
            var c when c.Contains("질소") || c.Contains("인")   => ("#3a2a1a", "#ccaa88"),
            var c when c.Contains("금속")                       => ("#2a3a1a", "#aacc88"),
            var c when c.Contains("대장") || c.Contains("세균") => ("#3a1a1a", "#ff8888"),
            var c when c.Contains("pH")  || c.Contains("수소") => ("#1a3a3a", "#88ddcc"),
            _                                                   => ("#2a2a3a", "#aaaacc"),
        };
    }

    private StackPanel BuildPhotoArea(string photoPath, string? 사번 = null)
    {
        var panel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 12,
            Margin      = new Thickness(0, 0, 0, 4)
        };

        // 사진 미리보기
        _photoImage = new Image
        {
            Width   = 80,
            Height  = 100,
            Stretch = Stretch.UniformToFill,
        };

        // 사진 테두리
        var photoBorder = new Border
        {
            Width           = 80,
            Height          = 100,
            CornerRadius    = new CornerRadius(6),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555577")),
            Background      = new SolidColorBrush(Color.Parse("#252525")),
            ClipToBounds    = true,
            Child           = _photoImage
        };

        // 초기 사진 로드 (로컬 없으면 DB에서 가져옴)
        LoadPhotoToImage(_photoImage, photoPath, 사번);

        // 버튼들
        var btnPanel = new StackPanel { Spacing = 6, VerticalAlignment = VerticalAlignment.Center };

        var uploadBtn = new Button
        {
            Content         = "📷 사진 업로드",
            FontSize        = 11,
            FontFamily      = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Background      = new SolidColorBrush(Color.Parse("#3a4a6a")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };

        var removeBtn = new Button
        {
            Content         = "🗑 사진 제거",
            FontSize        = 11,
            FontFamily      = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Background      = new SolidColorBrush(Color.Parse("#4a3a3a")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };

        uploadBtn.Click += async (_, _) => await PickPhotoAsync();
        removeBtn.Click += (_, _) =>
        {
            _pendingPhotoPath = "";
            _photoImage!.Source = null;
        };

        btnPanel.Children.Add(uploadBtn);
        btnPanel.Children.Add(removeBtn);
        btnPanel.Children.Add(new TextBlock
        {
            Text       = "jpg / png / bmp",
            FontSize   = 10,
            Foreground = new SolidColorBrush(Color.Parse("#666666"))
        });

        panel.Children.Add(photoBorder);
        panel.Children.Add(btnPanel);
        return panel;
    }

    // ── 사진 파일 선택 ────────────────────────────────────────────────────────
    private async Task PickPhotoAsync()
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title          = "사진 선택",
            AllowMultiple  = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("이미지")
                {
                    Patterns = new[] { "*.jpg", "*.jpeg", "*.png", "*.bmp" }
                }
            }
        });

        if (files.Count == 0) return;

        var srcPath = files[0].Path.LocalPath;
        var ext     = Path.GetExtension(srcPath).ToLower();

        // 파일명: 사번.확장자 (추가 모드는 패널에서 사번 읽기, 없으면 임시명)
        string 사번 = "";
        if (!_isAddMode)
            사번 = _selectedAgent?.사번 ?? "";
        else if (_detailPanel != null)
            사번 = ReadFieldFromPanel(_detailPanel, "사번");

        var fileName = string.IsNullOrEmpty(사번)
            ? $"temp_{DateTime.Now:yyyyMMddHHmmss}{ext}"
            : $"{사번}{ext}";

        var destPath = Path.Combine(AgentService.GetPhotoDirectory(), fileName);
        File.Copy(srcPath, destPath, overwrite: true);

        // DB에는 파일명만 저장 (절대경로 X)
        _pendingPhotoPath = fileName;

        // 미리보기 갱신 (절대경로로 표시)
        if (_photoImage != null)
            LoadPhotoToImage(_photoImage, destPath);

        Log($"사진 선택: 파일명={fileName}");
    }

    // ── 이미지 로드 헬퍼 (파일명 또는 절대경로 모두 처리) ────────────────────
    private static void LoadPhotoToImage(Image img, string pathOrFileName, string? 사번 = null)
    {
        if (string.IsNullOrEmpty(pathOrFileName)) { img.Source = null; return; }

        // 로컬 파일 없으면 DB에서 가져와 캐시
        if (!string.IsNullOrEmpty(사번))
            AgentService.EnsurePhotoLocal(사번, pathOrFileName);

        // 절대경로가 아니면 PhotoDirectory와 조합
        var fullPath = Path.IsPathRooted(pathOrFileName)
            ? pathOrFileName
            : Path.Combine(AgentService.GetPhotoDirectory(), pathOrFileName);

        if (!File.Exists(fullPath)) { img.Source = null; return; }
        try
        {
            using var stream = File.OpenRead(fullPath);
            img.Source = new Bitmap(stream);
        }
        catch { img.Source = null; }
    }

    // ── 패널에서 특정 필드 값 읽기 (재귀) ──────────────────────────────────
    private static string ReadFieldFromPanel(Control parent, string fieldLabel)
    {
        IEnumerable<Control> children = parent switch
        {
            Panel p => p.Children,
            _ => Array.Empty<Control>(),
        };
        foreach (var child in children)
        {
            if (child is StackPanel sp && sp.Children.Count >= 2
                && sp.Children[0] is TextBlock lbl && sp.Children[1] is TextBox tb)
            {
                var label = lbl.Text?.Replace("🔒 ", "").Trim() ?? "";
                if (label == fieldLabel) return tb.Text ?? "";
            }
            else
            {
                var found = ReadFieldFromPanel(child, fieldLabel);
                if (!string.IsNullOrEmpty(found)) return found;
            }
        }
        return "";
    }

    // =========================================================================
    // 수정 저장
    // =========================================================================
    private void SaveEdit()
    {
        if (_selectedAgent == null || _detailPanel == null)
        {
            Log("저장 스킵: 선택 없음");
            return;
        }

        SyncPanelToAgent(_detailPanel, _selectedAgent, includeReadOnly: false);

        // 사진 경로 반영
        if (!string.IsNullOrEmpty(_pendingPhotoPath))
            _selectedAgent.PhotoPath = _pendingPhotoPath;

        bool ok = AgentService.Update(_selectedAgent);
        Log(ok ? $"✅ 수정 저장: {_selectedAgent.성명}" : $"❌ 수정 실패: {_selectedAgent.성명}");

        if (ok)
        {
            // 사진이 있으면 DB에 바이너리도 저장 (다른 PC 동기화용)
            SyncPhotoToDb(_selectedAgent.사번, _selectedAgent.PhotoPath);
            _pendingPhotoPath = "";

            // 업무 분장 타임라인 새로고침
            if (_timelineCanvas != null && _selectedAgent != null)
            {
                RefreshTimeline(_timelineCanvas, _selectedAgent, _assignmentRangeStart, _assignmentRangeEnd);
            }
        }
    }

    // =========================================================================
    // 추가 저장
    // =========================================================================
    private void SaveAdd()
    {
        if (_detailPanel == null) return;

        var newAgent = new Agent();
        SyncPanelToAgent(_detailPanel, newAgent, includeReadOnly: true);

        if (string.IsNullOrWhiteSpace(newAgent.성명))
        {
            Log("❌ 성명 없음 → 추가 취소");
            return;
        }

        newAgent.Original성명 = newAgent.성명;
        if (!string.IsNullOrEmpty(_pendingPhotoPath))
            newAgent.PhotoPath = _pendingPhotoPath;

        bool ok = AgentService.Insert(newAgent);
        Log(ok ? $"✅ 추가 성공: {newAgent.성명}" : $"❌ 추가 실패: {newAgent.성명}");

        if (ok)
        {
            // 사진이 있으면 DB에 바이너리도 저장 (다른 PC 동기화용)
            SyncPhotoToDb(newAgent.사번, newAgent.PhotoPath);
            AgentTreeView.Items.Add(CreateTreeItem(newAgent));
            _isAddMode        = false;
            _pendingPhotoPath = "";
            _detailPanel      = null;
            DetailPanelChanged?.Invoke(null);
        }
    }

    // =========================================================================
    // UI → Agent 동기화
    // =========================================================================
    private static void SyncPanelToAgent(StackPanel panel, Agent agent, bool includeReadOnly)
    {
        SyncFieldsRecursive(panel, agent, includeReadOnly);
    }

    private static void SyncFieldsRecursive(Control parent, Agent agent, bool includeReadOnly)
    {
        IEnumerable<Control> children = parent switch
        {
            Panel p => p.Children,
            _ => Array.Empty<Control>(),
        };
        foreach (var child in children)
        {
            if (child is StackPanel sp && sp.Children.Count >= 2
                && sp.Children[0] is TextBlock lbl && sp.Children[1] is TextBox tb)
            {
                if (tb.IsReadOnly && !includeReadOnly) continue;
                var label = lbl.Text?.Replace("🔒 ", "").Trim() ?? "";

                switch (label)
                {
                    case "성명":           agent.성명           = tb.Text ?? ""; break;
                    case "직급":           agent.직급           = tb.Text ?? ""; break;
                    case "직무":           agent.직무           = tb.Text ?? ""; break;
                    case "사번":           agent.사번           = tb.Text ?? ""; break;
                    case "자격사항":       agent.자격사항       = tb.Text ?? ""; break;
                    case "Email":          agent.Email          = tb.Text ?? ""; break;
                    case "측정인고유번호": agent.측정인고유번호 = tb.Text ?? ""; break;
                    case "입사일":
                        if (DateOnly.TryParse(tb.Text, out var d)) agent.입사일 = d;
                        break;
                }
            }
            else
            {
                SyncFieldsRecursive(child, agent, includeReadOnly);
            }
        }
    }

    // =========================================================================
    // UI 헬퍼
    // =========================================================================
    private static StackPanel MakeRootPanel(string title)
    {
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(4) };
        root.Children.Add(new TextBlock
        {
            Text       = title,
            FontSize   = 15,
            FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 M",
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 0, 0, 4)
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = new SolidColorBrush(Color.Parse("#555555")),
            Margin     = new Thickness(0, 0, 0, 4)
        });
        return root;
    }

    private static StackPanel BuildFieldRow(string label, string value,
                                            bool isReadOnly = false,
                                            bool isLocked   = false,
                                            string hint     = "")
    {
        var panel = new StackPanel { Spacing = 2 };

        panel.Children.Add(new TextBlock
        {
            Text              = (isLocked ? "🔒 " : "") + label,
            FontSize          = 10,
            FontFamily        = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Foreground        = isLocked
                                    ? new SolidColorBrush(Color.Parse("#888888"))
                                    : AppRes("FgMuted"),
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            FontSize        = 12,
            FontFamily      = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            IsReadOnly      = isReadOnly,
            Watermark       = hint,
            Background      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#252525"))
                                  : new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#666666"))
                                  : AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#333333"))
                                  : new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 4)
        });

        return panel;
    }

    /// <summary>필드를 2열 Grid로 배치</summary>
    private static Grid BuildFieldGrid(params StackPanel[] fields)
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,8,*"),
            Margin = new Thickness(0, 2, 0, 2),
        };
        for (int i = 0; i < fields.Length; i++)
        {
            int col = (i % 2) * 2;     // 0 또는 2
            int row = i / 2;
            if (grid.RowDefinitions.Count <= row)
                grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
            Grid.SetColumn(fields[i], col);
            Grid.SetRow(fields[i], row);
            grid.Children.Add(fields[i]);
        }
        return grid;
    }

    // =========================================================================
    // 담당항목/업체 칩 섹션 헬퍼
    // =========================================================================
    private Control BuildChipSectionHeader(string title, Action? onAdd, Action<DateTime>? onAssignDate = null)
    {
        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto"),
            Margin = new Thickness(0, 10, 0, 2),
        };
        row.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = 11, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Foreground = new SolidColorBrush(Color.Parse("#aaaacc")),
            VerticalAlignment = VerticalAlignment.Center,
        });
        if (CanEdit)
        {
            var btnStack = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 3, [Grid.ColumnProperty] = 1 };

            if (onAssignDate != null)
            {
                var calPicker = new CalendarDatePicker
                {
                    Width = 200,
                    Watermark = "배정 날짜",
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                };
                var flyout = new Flyout { Content = calPicker };
                calPicker.SelectedDateChanged += (_, _) =>
                {
                    if (calPicker.SelectedDate.HasValue)
                    {
                        flyout.Hide();
                        onAssignDate(calPicker.SelectedDate.Value);
                        calPicker.SelectedDate = null;
                    }
                };
                var btnCal = new Button
                {
                    Content = "📅", FontSize = 11, Width = 24, Height = 24,
                    Padding = new Thickness(0),
                    Background = new SolidColorBrush(Color.Parse("#2a3a2a")),
                    Foreground = new SolidColorBrush(Color.Parse("#88cc88")),
                    BorderThickness = new Thickness(1),
                    BorderBrush = new SolidColorBrush(Color.Parse("#3a6a3a")),
                    CornerRadius = new CornerRadius(4),
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Flyout = flyout,
                };
                btnStack.Children.Add(btnCal);
            }

            var btnAdd = new Button
            {
                Content = "+", FontSize = 13, Width = 24, Height = 24,
                Padding = new Thickness(0),
                Background = new SolidColorBrush(Color.Parse("#2a3a5a")),
                Foreground = new SolidColorBrush(Color.Parse("#88aaee")),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#3a5a8a")),
                CornerRadius = new CornerRadius(4),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
            };
            btnAdd.Click += (_, _) => onAdd?.Invoke();
            btnStack.Children.Add(btnAdd);
            row.Children.Add(btnStack);
        }
        return row;
    }

    /// <summary>agent.담당항목목록의 모든 항목을 지정 날짜의 분장표준처리에 배정</summary>
    private static void AssignItemsForDate(Agent agent, DateTime date)
    {
        if (agent.담당항목목록.Count == 0) return;
        foreach (var item in agent.담당항목목록)
            AnalysisRequestService.AddAssignment(agent.사번, item, date, date);
        Debug.WriteLine($"[AssignItemsForDate] {agent.성명} → {agent.담당항목} @ {date:yyyy-MM-dd}");
    }

    // =========================================================================
    // QAQC 첨부 파일 헬퍼
    // =========================================================================
    private static string GetQaqcDir(string yearMonth)
    {
        var dataDir = Path.GetDirectoryName(DbPathHelper.DbPath)!;
        return Path.Combine(dataDir, "QAQC", yearMonth);
    }

    private static string GetShortForQaqc(string fullName)
    {
        var info = AnalysisRequestService.GetStandardDaysInfo();
        if (info.TryGetValue(fullName, out var meta) && !string.IsNullOrEmpty(meta.shortName))
            return meta.shortName;
        // 파일명 안전 처리
        var safe = string.Concat(fullName.Where(c => char.IsLetterOrDigit(c) || c == '_'));
        return safe.Length > 0 ? safe[..Math.Min(safe.Length, 20)] : "item";
    }

    private static bool HasQaqcFile(string shortName)
    {
        var qaqcRoot = Path.Combine(Path.GetDirectoryName(DbPathHelper.DbPath)!, "QAQC");
        if (!Directory.Exists(qaqcRoot)) return false;
        // 전체 하위 폴더에서 패턴 검색 (월 무관)
        return Directory.GetFiles(qaqcRoot, $"????-??-??_{shortName}_*", SearchOption.AllDirectories).Length > 0;
    }

    // =========================================================================
    // 배정 + 첨부 패널 (Show4)
    // =========================================================================
    private Control BuildAssignWithAttachPanel(Agent agent, DateTime date)
    {
        var ym      = date.ToString("yyyy-MM");
        var qaqcDir = GetQaqcDir(ym);
        var font    = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

        var root = new StackPanel { Spacing = 6, Margin = new Thickness(6) };

        // 헤더
        root.Children.Add(new TextBlock
        {
            Text = $"📅 {date:yyyy년 MM월} 배정 + 첨부",
            FontSize = 12, FontWeight = FontWeight.SemiBold,
            FontFamily = font, Foreground = new SolidColorBrush(Color.Parse("#aaccee")),
        });
        root.Children.Add(new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#333355")), Margin = new Thickness(0, 0, 0, 4) });

        var listStack = new StackPanel { Spacing = 6 };

        foreach (var fullName in agent.담당항목목록.ToList())
        {
            var capturedFull  = fullName;
            var capturedShort = GetShortForQaqc(fullName);
            var info     = AnalysisRequestService.GetStandardDaysInfo();
            bool hasMeta = info.TryGetValue(fullName, out var meta);
            string displayShort = hasMeta && !string.IsNullOrEmpty(meta.shortName) ? meta.shortName : capturedShort;
            var (chipBg, chipFg) = BadgeColorHelper.GetBadgeColor(displayShort);

            var itemBorder = new Border
            {
                Background      = new SolidColorBrush(Color.Parse("#1a1a2e")),
                BorderBrush     = new SolidColorBrush(Color.Parse("#333355")),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(4),
                Padding         = new Thickness(6, 4),
            };
            var itemStack = new StackPanel { Spacing = 3 };

            // 항목명 행
            var nameRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
            nameRow.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.Parse(chipBg)),
                CornerRadius = new CornerRadius(3), Padding = new Thickness(5, 1),
                Margin       = new Thickness(0, 0, 6, 0),
                [Grid.ColumnProperty] = 0,
                Child = new TextBlock
                {
                    Text = displayShort, FontSize = 9,
                    FontFamily = font, Foreground = new SolidColorBrush(Color.Parse(chipFg)),
                },
            });
            nameRow.Children.Add(new TextBlock
            {
                Text = capturedFull, FontSize = 10, FontFamily = font,
                Foreground = new SolidColorBrush(Color.Parse("#ccccee")),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                [Grid.ColumnProperty] = 1,
            });

            // 📎 버튼
            var btnAttach = new Button
            {
                Content = "📎", FontSize = 11, Width = 24, Height = 24,
                Padding = new Thickness(0),
                Background = new SolidColorBrush(Color.Parse("#2a3a5a")),
                Foreground = new SolidColorBrush(Color.Parse("#88aaee")),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#3a5a8a")),
                CornerRadius = new CornerRadius(4),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 2,
            };
            nameRow.Children.Add(btnAttach);
            itemStack.Children.Add(nameRow);

            // 첨부 파일 목록
            var fileStack = new StackPanel { Spacing = 2, Margin = new Thickness(4, 2, 0, 0) };

            void RefreshFiles()
            {
                fileStack.Children.Clear();
                if (!Directory.Exists(qaqcDir)) return;
                var files = Directory.GetFiles(qaqcDir, $"????-??-??_{capturedShort}_*");
                foreach (var fPath in files)
                {
                    var fName   = Path.GetFileName(fPath);
                    var display = fName;
                    var fRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
                    fRow.Children.Add(new TextBlock
                    {
                        Text = "📄 " + display, FontSize = 9, FontFamily = font,
                        Foreground = new SolidColorBrush(Color.Parse("#88ee88")),
                        VerticalAlignment = VerticalAlignment.Center,
                    });
                    var capturedF = fPath;
                    var delBtn = new Button
                    {
                        Content = "×", FontSize = 9, Width = 16, Height = 16, Padding = new Thickness(0),
                        Background = new SolidColorBrush(Color.Parse("#3a1a1a")),
                        Foreground = new SolidColorBrush(Color.Parse("#ee8888")),
                        BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(2),
                        HorizontalContentAlignment = HorizontalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center,
                    };
                    delBtn.Click += (_, _) =>
                    {
                        try { File.Delete(capturedF); } catch { }
                        RefreshFiles();
                        RefreshItemChips(agent);
                    };
                    fRow.Children.Add(delBtn);
                    fileStack.Children.Add(fRow);
                }
                if (files.Length == 0)
                    fileStack.Children.Add(new TextBlock
                    {
                        Text = "첨부 없음", FontSize = 9, FontFamily = font,
                        Foreground = new SolidColorBrush(Color.Parse("#555577")),
                    });
            }
            RefreshFiles();

            btnAttach.Click += async (_, _) =>
            {
                var topLevel = TopLevel.GetTopLevel(btnAttach);
                if (topLevel == null) return;
                var picked = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
                {
                    Title = $"{capturedFull} 첨부파일 선택",
                    AllowMultiple = false,
                });
                if (picked.Count == 0) return;
                Directory.CreateDirectory(qaqcDir);
                var src  = picked[0].Path.LocalPath;
                var ext      = Path.GetExtension(src);
                var dateStr  = date.ToString("yyyy-MM-dd");
                var dest     = Path.Combine(qaqcDir,
                    $"{dateStr}_{capturedShort}_{agent.성명}{ext}");
                File.Copy(src, dest, overwrite: true);
                RefreshFiles();
                RefreshItemChips(agent);
            };

            itemStack.Children.Add(fileStack);
            itemBorder.Child = itemStack;
            listStack.Children.Add(itemBorder);
        }

        root.Children.Add(new ScrollViewer
        {
            Content = listStack,
            MaxHeight = 320,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        });

        // 배정 완료 버튼
        var btnDone = new Button
        {
            Content = "✅ 배정 완료",
            Height = 32, Padding = new Thickness(16, 0),
            FontFamily = font, FontSize = 11,
            Background = new SolidColorBrush(Color.Parse("#1a3a2a")),
            Foreground = new SolidColorBrush(Color.Parse("#88ee88")),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#2a6a3a")),
            CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 4, 0, 0),
        };
        btnDone.Click += (_, _) =>
        {
            AssignItemsForDate(agent, date);
            RefreshItemChips(agent);
            Show4ContentRequest?.Invoke(BuildShow4AnalyteList(agent));
        };
        root.Children.Add(btnDone);

        return new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    private void RefreshItemChips(Agent agent)
    {
        if (_항목ChipsPanel == null) return;
        _항목ChipsPanel.Children.Clear();
        if (agent.담당항목목록.Count == 0)
        {
            _항목ChipsPanel.Children.Add(new TextBlock
            {
                Text = "없음", FontSize = 10,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#555")),
                Margin = new Thickness(0, 2),
            });
            return;
        }
        var shortNames = AnalysisRequestService.GetShortNames();
        foreach (var item in agent.담당항목목록.ToList())
        {
            var captured   = item;
            // 분장표준처리 ROW2 약칭 우선, fallback으로 GetShortForQaqc
            string shortName = shortNames.TryGetValue(item, out var sn) && !string.IsNullOrEmpty(sn)
                ? sn : GetShortForQaqc(item);
            bool certified = HasQaqcFile(shortName);
            string display = !string.IsNullOrEmpty(shortName) ? shortName : item;
            string label   = certified ? $"🏅 {display}" : display;
            string bg      = certified ? "#2a3a1a" : "#1a2a3a";
            string fg      = certified ? "#aaccaa" : "#88aacc";
            string bd      = certified ? "#4a8a2a" : "#3366aa";
            _항목ChipsPanel.Children.Add(BuildRemovableChip(label, bg, fg, bd, () =>
            {
                var list = agent.담당항목목록.ToList();
                list.Remove(captured);
                agent.담당항목 = string.Join(",", list);
                AgentService.Update(agent);
                AssignItemsForDate(agent, DateTime.Today);
                RefreshItemChips(agent);
            }));
        }
    }

    private void RefreshContractChips(Agent agent)
    {
        if (_업체ChipsPanel == null) return;
        _업체ChipsPanel.Children.Clear();
        if (agent.담당업체목록.Count == 0)
        {
            _업체ChipsPanel.Children.Add(new TextBlock
            {
                Text = "없음", FontSize = 10,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#555")),
                Margin = new Thickness(0, 2),
            });
            return;
        }
        foreach (var abbr in agent.담당업체목록.ToList())
        {
            var captured = abbr;
            _업체ChipsPanel.Children.Add(BuildRemovableChip(captured, "#1a3a2a", "#88ccaa", "#2a6633", () =>
            {
                var list = agent.담당업체목록.ToList();
                list.Remove(captured);
                agent.담당업체 = string.Join(",", list);
                AgentService.Update(agent);
                RefreshContractChips(agent);
            }));
        }
    }

    private static Border BuildRemovableChip(string label, string bg, string fg, string border, Action? onRemove)
    {
        var chip = new Border
        {
            Background      = new SolidColorBrush(Color.Parse(bg)),
            BorderBrush     = new SolidColorBrush(Color.Parse(border)),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(10),
            Padding         = new Thickness(7, 2),
            Margin          = new Thickness(0, 2, 4, 2),
        };
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 3 };
        row.Children.Add(new TextBlock
        {
            Text = label, FontSize = 10,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Foreground = new SolidColorBrush(Color.Parse(fg)),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var btnX = new Button
        {
            Content = "×", FontSize = 11, Padding = new Thickness(2, 0),
            Background = Brushes.Transparent, BorderThickness = new Thickness(0),
            Foreground = new SolidColorBrush(Color.Parse("#886888")),
            VerticalAlignment = VerticalAlignment.Center,
        };
        btnX.Click += (_, _) => onRemove?.Invoke();
        row.Children.Add(btnX);
        chip.Child = row;
        return chip;
    }

    // =========================================================================
    // 과거 분장이력 섹션
    // =========================================================================
    private Control BuildHistorySection(Agent agent)
    {
        var border = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#1a1a2a")),
            CornerRadius = new CornerRadius(4),
            Padding      = new Thickness(8, 6),
            Margin       = new Thickness(0, 10, 0, 0),
        };
        var stack = new StackPanel { Spacing = 3 };
        stack.Children.Add(new TextBlock
        {
            Text = "📊 과거 분장이력 (전체)",
            FontSize = 11, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Foreground = new SolidColorBrush(Color.Parse("#aaaacc")),
            Margin = new Thickness(0, 0, 0, 4),
        });

        var histGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,8,*"),
        };
        try
        {
            var assignments = AnalysisRequestService.GetAssignmentDaysForAgentAll(agent.사번);
            if (assignments.Count == 0)
            {
                histGrid.Children.Add(new TextBlock
                {
                    Text = "이력 없음", FontSize = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Foreground = new SolidColorBrush(Color.Parse("#555")),
                });
            }
            else
            {
                var allAnalytes = AnalysisService.GetAllItems()
                    .ToDictionary(a => a.Analyte, a => a.Category, StringComparer.OrdinalIgnoreCase);
                var shortNameMap = AnalysisRequestService.GetShortNames();
                var groups = assignments
                    .GroupBy(a => a.FullName)
                    .Select(g =>
                    {
                        // 분장표준처리 ROW2 약칭 직접 매칭
                        string shortName = shortNameMap.TryGetValue(g.Key, out var sn) && !string.IsNullOrEmpty(sn)
                            ? sn : g.First().ShortName;
                        // 여전히 전체이름과 같으면 fallback
                        if (shortName == g.Key) shortName = g.Key.Length > 6 ? g.Key[..6] : g.Key;
                        return (Short: shortName, Full: g.Key, Count: g.Count());
                    })
                    .OrderByDescending(x => x.Count)
                    .ToList();

                int rowIdx = 0;
                for (int i = 0; i < groups.Count; i++)
                {
                    var g = groups[i];
                    int col = (i % 2) * 2;
                    int r   = i / 2;
                    if (histGrid.RowDefinitions.Count <= r)
                        histGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

                    allAnalytes.TryGetValue(g.Full, out var cat);
                    var (catBg, catFg) = GetCategoryColor(cat ?? "");
                    var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"), Margin = new Thickness(0, 1) };
                    row.Children.Add(new Border
                    {
                        Background   = new SolidColorBrush(Color.Parse(catBg)),
                        CornerRadius = new CornerRadius(3),
                        Padding      = new Thickness(4, 1), Margin = new Thickness(0, 0, 4, 0),
                        [Grid.ColumnProperty] = 0,
                        Child = new TextBlock
                        {
                            Text = g.Short, FontSize = 9,
                            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                            Foreground = new SolidColorBrush(Color.Parse(catFg)),
                        },
                    });
                    row.Children.Add(new TextBlock
                    {
                        Text = $"{g.Count}일",
                        FontSize = 9,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                        Foreground = new SolidColorBrush(Color.Parse("#7ab4cc")),
                        VerticalAlignment = VerticalAlignment.Center,
                        HorizontalAlignment = HorizontalAlignment.Right,
                        [Grid.ColumnProperty] = 2,
                    });
                    Grid.SetColumn(row, col);
                    Grid.SetRow(row, r);
                    histGrid.Children.Add(row);
                }
            }
        }
        catch (Exception ex) { Log("이력 로드 실패: " + ex.Message); }

        stack.Children.Add(new ScrollViewer
        {
            Content = histGrid,
            MaxHeight = 160,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        });
        border.Child = stack;
        return border;
    }

    // =========================================================================
    // Show3 — 탭 구조: 분석항목 | 담당업체 | 기타업무
    // =========================================================================
    private string _show3Tab = "분석항목";

    private Control BuildShow3Timeline(Agent agent)
    {
        var outer = new Grid { RowDefinitions = new RowDefinitions("Auto,*"), Margin = new Thickness(4) };

        // ── 탭 바 ──────────────────────────────────────────────────────────────
        var tabBar = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, Margin = new Thickness(0, 0, 0, 6) };
        var tabs = new (string Key, string Icon, string Label, string Bg, string Fg, string Bd)[]
        {
            ("분석항목", "🧪", "분석항목", "#1a1a2a", "#aaccff", "#3a5a8a"),
            ("담당업체", "🏢", "담당업체", "#1a2a1a", "#aaccaa", "#3a6a3a"),
            ("기타업무", "📋", "기타업무", "#2a1a2a", "#ccaaff", "#6a3a8a"),
        };
        var contentPanel = new StackPanel { Spacing = 4 };
        Grid.SetRow(contentPanel, 1);

        foreach (var t in tabs)
        {
            var key = t.Key;
            bool active = key == _show3Tab;
            var btn = new Button
            {
                Content     = $"{t.Icon} {t.Label}",
                FontSize    = 11,
                FontFamily  = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                Padding     = new Thickness(10, 4),
                CornerRadius = new CornerRadius(8, 8, 0, 0),
                BorderThickness = new Thickness(1, 1, 1, 0),
                Background  = active ? new SolidColorBrush(Color.Parse(t.Bg)) : new SolidColorBrush(Color.Parse("#1a1a1a")),
                Foreground  = active ? new SolidColorBrush(Color.Parse(t.Fg)) : new SolidColorBrush(Color.Parse("#666")),
                BorderBrush = active ? new SolidColorBrush(Color.Parse(t.Bd)) : new SolidColorBrush(Color.Parse("#333")),
                Cursor      = new Cursor(StandardCursorType.Hand),
            };
            btn.Click += (_, _) =>
            {
                _show3Tab = key;
                Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
            };
            tabBar.Children.Add(btn);
        }
        Grid.SetRow(tabBar, 0);
        outer.Children.Add(tabBar);

        // ── 탭 콘텐츠 ──────────────────────────────────────────────────────────
        switch (_show3Tab)
        {
            case "분석항목":
                contentPanel.Children.Add(BuildShow3AnalyteTab(agent));
                break;
            case "담당업체":
                contentPanel.Children.Add(BuildShow3ContractTab(agent));
                break;
            case "기타업무":
                contentPanel.Children.Add(BuildShow3MiscTaskTab(agent));
                break;
        }
        outer.Children.Add(contentPanel);
        return outer;
    }

    // ── 분석항목 탭 (기존 타임라인) ──────────────────────────────────────────
    private Control BuildShow3AnalyteTab(Agent agent)
    {
        var root = new StackPanel { Spacing = 4 };

        root.Children.Add(new TextBlock
        {
            Text = $"📈 업무 분장 타임라인 — {agent.성명}",
            FontSize = 13, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccff")),
            Margin = new Thickness(0, 0, 0, 2),
        });

        var dateRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        var txbRange = new TextBlock
        {
            Text = $"{DateTime.Today:yyyy-MM-dd}",
            FontSize = 11,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Foreground = new SolidColorBrush(Color.Parse("#aaa")),
            VerticalAlignment = VerticalAlignment.Center,
            MinWidth = 170,
        };
        var btnToday = new Button
        {
            Content = "오늘", Width = 48, Height = 24, FontSize = 10,
            Padding = new Thickness(4, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Background = new SolidColorBrush(Color.Parse("#2a3a5a")),
            Foreground = new SolidColorBrush(Color.Parse("#88aacc")),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#3a5a8a")),
            CornerRadius = new CornerRadius(4),
        };
        var btnMonth = new Button
        {
            Content = "이번달", Width = 52, Height = 24, FontSize = 10,
            Padding = new Thickness(4, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Background = new SolidColorBrush(Color.Parse("#2a3a2a")),
            Foreground = new SolidColorBrush(Color.Parse("#88cc88")),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#3a6a3a")),
            CornerRadius = new CornerRadius(4),
        };
        dateRow.Children.Add(txbRange);
        dateRow.Children.Add(btnToday);
        dateRow.Children.Add(btnMonth);
        root.Children.Add(dateRow);

        var canvas = new Canvas { Width = 400, Height = 100 };
        _timelineCanvas = canvas;

        var scroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Height  = 300,
            Margin  = new Thickness(0, 4, 0, 0),
            Content = canvas,
        };
        _timelineScroll = scroll;
        DragDrop.SetAllowDrop(scroll, true);

        DateTime rangeStart = DateTime.Today;
        DateTime rangeEnd   = DateTime.Today;

        scroll.AddHandler(DragDrop.DropEvent, async (object? sender, DragEventArgs e) =>
        {
            if (!e.Data.Contains("analyte")) return;
            var analyte = e.Data.Get("analyte") as string;
            if (string.IsNullOrEmpty(analyte)) return;

            var existing = AnalysisRequestService.GetAssigneesForAnalyteOnDate(analyte, rangeStart);
            if (existing.Count > 0 && !existing.Contains(agent.성명))
            {
                bool doUpdate = await ShowDuplicateWarningAsync(analyte, existing);
                if (!doUpdate) return;
            }

            AnalysisRequestService.AddAssignment(agent.사번, analyte, rangeStart, rangeStart.AddMonths(1));
            RefreshTimeline(canvas, agent, rangeStart, rangeEnd);
        });

        btnToday.Click += (_, _) =>
        {
            rangeStart = rangeEnd = DateTime.Today;
            _assignmentRangeStart = _assignmentRangeEnd = DateTime.Today;
            txbRange.Text = rangeStart.ToString("yyyy-MM-dd");
            RefreshTimeline(canvas, agent, rangeStart, rangeEnd);
        };

        btnMonth.Click += (_, _) =>
        {
            rangeStart = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            rangeEnd   = rangeStart.AddMonths(1).AddDays(-1);
            _assignmentRangeStart = rangeStart;
            _assignmentRangeEnd   = rangeEnd;
            txbRange.Text = $"{rangeStart:yyyy-MM-dd} ~ {rangeEnd:yyyy-MM-dd}";
            RefreshTimeline(canvas, agent, rangeStart, rangeEnd);
        };

        RefreshTimeline(canvas, agent, rangeStart, rangeEnd);
        root.Children.Add(scroll);
        return root;
    }

    // ── 담당업체 탭 ─────────────────────────────────────────────────────────
    private Control BuildShow3ContractTab(Agent agent)
    {
        var root = new StackPanel { Spacing = 6 };
        root.Children.Add(new TextBlock
        {
            Text = $"🏢 담당 계약업체 — {agent.성명}",
            FontSize = 13, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccaa")),
            Margin = new Thickness(0, 0, 0, 4),
        });

        if (agent.담당업체목록.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "배정된 계약업체 없음", FontSize = 11,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#666")),
            });
            return root;
        }

        foreach (var name in agent.담당업체목록)
        {
            var (bg, fg) = BadgeColorHelper.GetBadgeColor(name);
            var row = new Border
            {
                Background   = new SolidColorBrush(Color.Parse("#1a2a1a")),
                BorderBrush  = new SolidColorBrush(Color.Parse("#3a6a3a")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(8, 6),
                Margin       = new Thickness(0, 1),
            };
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            sp.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.Parse(bg)),
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(6, 2),
                Child = new TextBlock
                {
                    Text = name, FontSize = 11,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground = new SolidColorBrush(Color.Parse(fg)),
                },
            });
            row.Child = sp;
            root.Children.Add(row);
        }
        return root;
    }

    // ── 기타업무 탭 ─────────────────────────────────────────────────────────
    private Control BuildShow3MiscTaskTab(Agent agent)
    {
        var root = new StackPanel { Spacing = 6 };

        // 헤더 + 신규 추가 버튼
        var hdr = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Margin = new Thickness(0, 0, 0, 4) };
        hdr.Children.Add(new TextBlock
        {
            Text = $"📋 기타업무 — {agent.성명}",
            FontSize = 13, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = new SolidColorBrush(Color.Parse("#ccaaff")),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var btnAdd = new Button
        {
            Content = "➕ 신규 추가", FontSize = 11, Height = 28,
            Padding = new Thickness(10, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Background = new SolidColorBrush(Color.Parse("#2a1a3a")),
            Foreground = new SolidColorBrush(Color.Parse("#cc88ff")),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#6a3a8a")),
            CornerRadius = new CornerRadius(4),
            Cursor = new Cursor(StandardCursorType.Hand),
            [Grid.ColumnProperty] = 1,
        };
        btnAdd.Click += (_, _) =>
        {
            // 인라인 폼 표시
            Show3ContentRequest?.Invoke(BuildMiscTaskForm(agent));
        };
        hdr.Children.Add(btnAdd);
        root.Children.Add(hdr);

        // 업무 목록
        var tasks = MiscTaskService.GetByAgent(agent.사번);
        if (tasks.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "배정된 업무 없음", FontSize = 11,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#666")),
                Margin = new Thickness(0, 8),
            });
            return root;
        }

        foreach (var t in tasks)
        {
            var statusStyle = t.상태 switch
            {
                "진행" => (Bg: "#1a2a1a", Fg: "#88cc88", Bd: "#3a6a3a", Badge: "#2a5a2a"),
                "완료" => (Bg: "#1a1a2a", Fg: "#88aacc", Bd: "#336699", Badge: "#2a3a5a"),
                _      => (Bg: "#2a2a1a", Fg: "#ccaa88", Bd: "#665533", Badge: "#4a3a1a"),
            };

            var card = new Border
            {
                Background      = new SolidColorBrush(Color.Parse(statusStyle.Bg)),
                BorderBrush     = new SolidColorBrush(Color.Parse(statusStyle.Bd)),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 8),
                Margin          = new Thickness(0, 2),
            };
            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto,Auto") };

            // 상태 뱃지
            grid.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.Parse(statusStyle.Badge)),
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(6, 2),
                Margin       = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = t.상태, FontSize = 10,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground = new SolidColorBrush(Color.Parse(statusStyle.Fg)),
                },
            });

            // 정보
            var info = new StackPanel { Spacing = 2, [Grid.ColumnProperty] = 1 };
            info.Children.Add(new TextBlock
            {
                Text = t.업무명, FontSize = 12,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                Foreground = new SolidColorBrush(Color.Parse(statusStyle.Fg)),
            });
            var meta = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            if (!string.IsNullOrEmpty(t.마감일))
                meta.Children.Add(new TextBlock
                {
                    Text = $"마감: {t.마감일}", FontSize = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Foreground = new SolidColorBrush(Color.Parse("#888")),
                });
            if (!string.IsNullOrEmpty(t.내용))
                meta.Children.Add(new TextBlock
                {
                    Text = t.내용, FontSize = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Foreground = new SolidColorBrush(Color.Parse("#777")),
                    TextTrimming = TextTrimming.CharacterEllipsis, MaxWidth = 200,
                });
            if (meta.Children.Count > 0) info.Children.Add(meta);
            grid.Children.Add(info);

            // 수정 버튼
            var capturedTask = t;
            var btnEdit = new Button
            {
                Content = "✏️", Width = 28, Height = 28, Padding = new Thickness(0),
                Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#555")),
                CornerRadius = new CornerRadius(4),
                Cursor = new Cursor(StandardCursorType.Hand),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
                [Grid.ColumnProperty] = 2,
            };
            btnEdit.Click += (_, _) => Show3ContentRequest?.Invoke(BuildMiscTaskForm(agent, capturedTask));
            grid.Children.Add(btnEdit);

            // 삭제 버튼
            var taskId = t.Id;
            var btnDel = new Button
            {
                Content = "🗑", Width = 28, Height = 28, Padding = new Thickness(0),
                Background = new SolidColorBrush(Color.Parse("#3a1a1a")),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.Parse("#663333")),
                CornerRadius = new CornerRadius(4),
                Cursor = new Cursor(StandardCursorType.Hand),
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 3,
            };
            btnDel.Click += (_, _) =>
            {
                MiscTaskService.Delete(taskId);
                RefreshMiscTasks(agent);
                _show3Tab = "기타업무";
                Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
            };
            grid.Children.Add(btnDel);

            card.Child = grid;
            root.Children.Add(card);
        }
        return root;
    }

    // =========================================================================
    // 기타업무 — Show2 목록 + Show3 폼
    // =========================================================================
    private void RefreshMiscTasks(Agent agent)
    {
        if (_miscTaskPanel == null) return;
        _miscTaskPanel.Children.Clear();

        var tasks = MiscTaskService.GetByAgent(agent.사번);
        if (tasks.Count == 0)
        {
            _miscTaskPanel.Children.Add(new TextBlock
            {
                Text = "배정된 업무 없음", FontSize = 10,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#555")),
                Margin = new Thickness(0, 2),
            });
            return;
        }

        foreach (var t in tasks)
        {
            var statusColor = t.상태 switch
            {
                "진행" => ("#2a3a1a", "#88cc88", "#3a6a3a"),
                "완료" => ("#1a2a3a", "#88aacc", "#336699"),
                _      => ("#2a2a2a", "#ccaa88", "#665533"),
            };
            var card = new Border
            {
                Background   = new SolidColorBrush(Color.Parse(statusColor.Item1)),
                BorderBrush  = new SolidColorBrush(Color.Parse(statusColor.Item3)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(6, 3),
                Margin       = new Thickness(0, 1),
            };
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };

            // 상태 뱃지
            row.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.Parse(statusColor.Item3)),
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(4, 1),
                Margin       = new Thickness(0, 0, 6, 0),
                Child = new TextBlock
                {
                    Text = t.상태, FontSize = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground = new SolidColorBrush(Color.Parse(statusColor.Item2)),
                },
            });

            // 업무명 + 마감일
            var info = new StackPanel { [Grid.ColumnProperty] = 1 };
            info.Children.Add(new TextBlock
            {
                Text = t.업무명, FontSize = 10,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse(statusColor.Item2)),
                TextTrimming = TextTrimming.CharacterEllipsis,
            });
            if (!string.IsNullOrEmpty(t.마감일))
                info.Children.Add(new TextBlock
                {
                    Text = $"마감: {t.마감일}", FontSize = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Foreground = new SolidColorBrush(Color.Parse("#777")),
                });
            row.Children.Add(info);

            // 수정 버튼
            var capturedTask = t;
            var capturedAgent = agent;
            var btnEdit = new Button
            {
                Content = "✏️", Width = 24, Height = 24, Padding = new Thickness(0),
                Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                Foreground = new SolidColorBrush(Color.Parse("#88aacc")),
                VerticalAlignment = VerticalAlignment.Center,
                Cursor = new Cursor(StandardCursorType.Hand),
                [Grid.ColumnProperty] = 2,
            };
            btnEdit.Click += (_, _) => Show3ContentRequest?.Invoke(BuildMiscTaskForm(capturedAgent, capturedTask));
            row.Children.Add(btnEdit);

            card.Child = row;
            _miscTaskPanel.Children.Add(card);
        }
    }

    /// <summary>기타업무 등록/수정 폼 (Show3에 표시)</summary>
    private Control BuildMiscTaskForm(Agent agent, MiscTask? existing = null)
    {
        bool isEdit = existing != null;
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(4) };

        root.Children.Add(new TextBlock
        {
            Text = isEdit ? $"✏️ 업무 수정 — {agent.성명}" : $"➕ 신규 업무 배정 — {agent.성명}",
            FontSize = 13, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = new SolidColorBrush(Color.Parse("#ccaaff")),
        });

        var txtName = new TextBox
        {
            Text = existing?.업무명 ?? "",
            Watermark = "업무명",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            FontSize = 12,
            Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
        };
        root.Children.Add(txtName);

        var txtContent = new TextBox
        {
            Text = existing?.내용 ?? "",
            Watermark = "상세 내용",
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 80,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            FontSize = 12,
            Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
        };
        root.Children.Add(txtContent);

        // 마감일
        var deadlineRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        deadlineRow.Children.Add(new TextBlock
        {
            Text = "마감일", FontSize = 11,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var txtDeadline = new TextBox
        {
            Text = existing?.마감일 ?? "",
            Watermark = "yyyy-MM-dd",
            Width = 140,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            FontSize = 12,
            Background = new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 4),
        };
        deadlineRow.Children.Add(txtDeadline);
        root.Children.Add(deadlineRow);

        // 상태 토글 (수정시)
        string currentStatus = existing?.상태 ?? "대기";
        TextBlock? statusLbl = null;
        if (isEdit)
        {
            var statusRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            statusRow.Children.Add(new TextBlock
            {
                Text = "상태", FontSize = 11,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
            });
            foreach (var st in new[] { "대기", "진행", "완료" })
            {
                var capturedSt = st;
                var (bg, fg, bd) = st switch
                {
                    "진행" => ("#2a3a1a", "#88cc88", "#3a6a3a"),
                    "완료" => ("#1a2a3a", "#88aacc", "#336699"),
                    _      => ("#2a2a2a", "#ccaa88", "#665533"),
                };
                bool active = st == currentStatus;
                var btnSt = new Button
                {
                    Content = st, FontSize = 10,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Padding = new Thickness(8, 3),
                    CornerRadius = new CornerRadius(8),
                    BorderThickness = new Thickness(1),
                    Background  = active ? new SolidColorBrush(Color.Parse(bg)) : new SolidColorBrush(Color.Parse("#222")),
                    Foreground  = active ? new SolidColorBrush(Color.Parse(fg)) : new SolidColorBrush(Color.Parse("#666")),
                    BorderBrush = active ? new SolidColorBrush(Color.Parse(bd)) : new SolidColorBrush(Color.Parse("#444")),
                };
                btnSt.Click += (_, _) =>
                {
                    currentStatus = capturedSt;
                    // 모든 상태 버튼 비활성 스타일로 리셋
                    foreach (var child in statusRow.Children.OfType<Button>())
                    {
                        var s = child.Content?.ToString() ?? "";
                        var (b2, f2, d2) = s switch
                        {
                            "진행" => ("#2a3a1a", "#88cc88", "#3a6a3a"),
                            "완료" => ("#1a2a3a", "#88aacc", "#336699"),
                            _      => ("#2a2a2a", "#ccaa88", "#665533"),
                        };
                        bool isActive = s == capturedSt;
                        child.Background  = isActive ? new SolidColorBrush(Color.Parse(b2)) : new SolidColorBrush(Color.Parse("#222"));
                        child.Foreground  = isActive ? new SolidColorBrush(Color.Parse(f2)) : new SolidColorBrush(Color.Parse("#666"));
                        child.BorderBrush = isActive ? new SolidColorBrush(Color.Parse(d2)) : new SolidColorBrush(Color.Parse("#444"));
                    }
                };
                statusRow.Children.Add(btnSt);
            }
            root.Children.Add(statusRow);
        }

        // 버튼 행
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(0, 8, 0, 0) };
        var btnSave = new Button
        {
            Content = isEdit ? "💾 수정" : "💾 저장",
            Height = 30, Padding = new Thickness(16, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            FontSize = 11,
            Background = new SolidColorBrush(Color.Parse("#1a3a2a")),
            Foreground = new SolidColorBrush(Color.Parse("#88ee88")),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnSave.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(txtName.Text)) return;
            if (isEdit)
            {
                existing!.업무명 = txtName.Text.Trim();
                existing!.내용   = txtContent.Text?.Trim() ?? "";
                existing!.마감일 = txtDeadline.Text?.Trim() ?? "";
                existing!.상태   = currentStatus;
                if (currentStatus == "완료" && string.IsNullOrEmpty(existing.완료일시))
                    existing.완료일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                MiscTaskService.Update(existing);
            }
            else
            {
                var t = new MiscTask
                {
                    업무명   = txtName.Text.Trim(),
                    내용     = txtContent.Text?.Trim() ?? "",
                    배정자   = ETA.Views.MainPage.CurrentEmployeeId,
                    담당자id = agent.사번,
                    담당자명 = agent.성명,
                    마감일   = txtDeadline.Text?.Trim() ?? "",
                    등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                };
                MiscTaskService.Insert(t);
            }
            RefreshMiscTasks(agent);
            _show3Tab = "기타업무";
            Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
        };
        btnRow.Children.Add(btnSave);

        if (isEdit)
        {
            var btnDel = new Button
            {
                Content = "🗑 삭제", Height = 30, Padding = new Thickness(10, 0),
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                FontSize = 11,
                Background = new SolidColorBrush(Color.Parse("#3a1a1a")),
                Foreground = new SolidColorBrush(Color.Parse("#ee8888")),
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            };
            btnDel.Click += (_, _) =>
            {
                MiscTaskService.Delete(existing!.Id);
                RefreshMiscTasks(agent);
                Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
            };
            btnRow.Children.Add(btnDel);
        }

        var btnBack = new Button
        {
            Content = "← 돌아가기", Height = 30, Padding = new Thickness(10, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            FontSize = 11,
            Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
            Foreground = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => { _show3Tab = "기타업무"; Show3ContentRequest?.Invoke(BuildShow3Timeline(agent)); };
        btnRow.Children.Add(btnBack);

        root.Children.Add(btnRow);
        return root;
    }

    // =========================================================================
    // Show4 — 분석항목 선택 리스트
    // =========================================================================
    private Control BuildShow4AnalyteList(Agent agent)
    {
        var root = new StackPanel { Spacing = 3, Margin = new Thickness(4) };

        // 헤더
        var hdr = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*"), Margin = new Thickness(0, 0, 0, 6) };
        var btnBack = new Button
        {
            Content = "← 돌아가기", FontSize = 10,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Padding = new Thickness(8, 4),
            Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
            Foreground = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => Show4ContentRequest?.Invoke(null);
        hdr.Children.Add(btnBack);
        hdr.Children.Add(new TextBlock
        {
            Text = "분석항목 선택", FontSize = 12, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccff")),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0),
            [Grid.ColumnProperty] = 1,
        });
        root.Children.Add(hdr);
        root.Children.Add(new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#444")) });

        // 분장표준처리 컬럼 순서대로 항목 목록
        var analytes = AnalysisRequestService.GetOrderedAnalytes();
        foreach (var (fullName, shortName) in analytes)
        {
            bool assigned = agent.담당항목목록.Contains(fullName);
            var (badgeBg, badgeFg) = BadgeColorHelper.GetBadgeColor(shortName);

            var row = new Border
            {
                Background   = assigned ? new SolidColorBrush(Color.Parse(badgeBg)) : Brushes.Transparent,
                CornerRadius = new CornerRadius(4),
                Padding      = new Thickness(6, 4),
                Margin       = new Thickness(0, 1),
                Cursor       = new Cursor(StandardCursorType.Hand),
            };
            var inner = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };

            // 배지: 분장표준처리 약칭
            inner.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.Parse(badgeBg)),
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(5, 1), Margin = new Thickness(0, 0, 8, 0),
                [Grid.ColumnProperty] = 0,
                Child = new TextBlock
                {
                    Text = shortName, FontSize = 9,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                },
            });

            // 전체 항목명
            inner.Children.Add(new TextBlock
            {
                Text = fullName, FontSize = 10,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = assigned ? new SolidColorBrush(Color.Parse(badgeFg)) : AppRes("AppFg"),
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 1,
            });

            if (assigned)
                inner.Children.Add(new TextBlock
                {
                    Text = "✓", FontSize = 11,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                    Foreground = new SolidColorBrush(Color.Parse("#88ee88")),
                    VerticalAlignment = VerticalAlignment.Center,
                    [Grid.ColumnProperty] = 2,
                });

            row.Child = inner;

            var capturedFull  = fullName;
            row.PointerPressed += (_, _) =>
            {
                var list = agent.담당항목목록.ToList();
                if (list.Contains(capturedFull)) list.Remove(capturedFull);
                else list.Add(capturedFull);
                agent.담당항목 = string.Join(",", list);
                AgentService.Update(agent);
                AssignItemsForDate(agent, DateTime.Today);
                RefreshItemChips(agent);
                Show4ContentRequest?.Invoke(BuildShow4AnalyteList(agent));
            };
            root.Children.Add(row);
        }

        return new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // =========================================================================
    // Show4 — 계약업체 선택 리스트 (현행 계약만)
    // =========================================================================
    private Control BuildShow4ContractList(Agent agent)
    {
        var root = new StackPanel { Spacing = 3, Margin = new Thickness(4) };

        // 헤더
        var hdr = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*"), Margin = new Thickness(0, 0, 0, 6) };
        var btnBack = new Button
        {
            Content = "← 돌아가기", FontSize = 10,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
            Padding = new Thickness(8, 4),
            Background = new SolidColorBrush(Color.Parse("#2a2a3a")),
            Foreground = new SolidColorBrush(Color.Parse("#aaaaaa")),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => Show4ContentRequest?.Invoke(null);
        hdr.Children.Add(btnBack);
        hdr.Children.Add(new TextBlock
        {
            Text = "계약업체 선택 (현행 계약)", FontSize = 12, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccaa")),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0),
            [Grid.ColumnProperty] = 1,
        });
        root.Children.Add(hdr);
        root.Children.Add(new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#444")) });

        var contracts = ContractService.GetAllContracts()
            .Where(c => c.DaysLeft == null || c.DaysLeft >= 0)
            .OrderBy(c => c.C_CompanyName).ToList();

        if (contracts.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "현행 계약 없음", FontSize = 10,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                Foreground = new SolidColorBrush(Color.Parse("#666")),
                Margin = new Thickness(0, 8),
            });
        }
        else
        {
            foreach (var c in contracts)
            {
                string abbr = string.IsNullOrWhiteSpace(c.C_Abbreviation) ? c.C_CompanyName : c.C_Abbreviation;
                bool assigned = agent.담당업체목록.Contains(abbr);
                var (bg, fg) = BadgeColorHelper.GetBadgeColor(abbr);

                var row = new Border
                {
                    Background   = assigned ? new SolidColorBrush(Color.Parse("#1a3a2a")) : Brushes.Transparent,
                    CornerRadius = new CornerRadius(4),
                    Padding      = new Thickness(6, 4),
                    Margin       = new Thickness(0, 1),
                    Cursor       = new Cursor(StandardCursorType.Hand),
                };
                var inner = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
                inner.Children.Add(new Border
                {
                    Background   = new SolidColorBrush(Color.Parse(bg)),
                    CornerRadius = new CornerRadius(3),
                    Padding      = new Thickness(5, 1), Margin = new Thickness(0, 0, 6, 0),
                    [Grid.ColumnProperty] = 0,
                    Child = new TextBlock
                    {
                        Text = abbr, FontSize = 9,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                        Foreground = new SolidColorBrush(Color.Parse(fg)),
                    },
                });
                inner.Children.Add(new TextBlock
                {
                    Text = c.C_CompanyName, FontSize = 10,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R"),
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    [Grid.ColumnProperty] = 1,
                });
                if (assigned)
                    inner.Children.Add(new TextBlock
                    {
                        Text = "✓", FontSize = 11,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
                        Foreground = new SolidColorBrush(Color.Parse("#88ee88")),
                        VerticalAlignment = VerticalAlignment.Center,
                        [Grid.ColumnProperty] = 2,
                    });
                row.Child = inner;

                var capturedAbbr = abbr;
                row.PointerPressed += (_, _) =>
                {
                    var list = agent.담당업체목록.ToList();
                    if (list.Contains(capturedAbbr)) list.Remove(capturedAbbr);
                    else list.Add(capturedAbbr);
                    agent.담당업체 = string.Join(",", list);
                    AgentService.Update(agent);
                    RefreshContractChips(agent);
                    Show4ContentRequest?.Invoke(BuildShow4ContractList(agent));
                };
                root.Children.Add(row);
            }
        }

        return new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // =========================================================================
    //  업무분장표 — Gantt 차트
    // =========================================================================
    public Control BuildAssignmentChart()
    {
        var today = DateTime.Today;
        _chartRangeStart = new DateTime(today.Year, today.Month, 1);
        _chartRangeEnd   = _chartRangeStart.AddMonths(1).AddDays(-1);

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };

        // ── 헤더 ────────────────────────────────────────────────────────────
        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(8, 6, 8, 4) };
        var kbFont = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
        var kbFontM = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

        header.Children.Add(new TextBlock {
            Text = "업무분장표", FontSize = 14, FontWeight = FontWeight.Bold,
            FontFamily = kbFontM, Foreground = new SolidColorBrush(Color.Parse("#aaccff")),
            VerticalAlignment = VerticalAlignment.Center });

        Button MakeBtn(string text, string bg, string fg, string bd, int w = 30) => new()
        {
            Content = text, Width = w, Height = 24, FontSize = 10,
            Padding = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = kbFont, Cursor = new Cursor(StandardCursorType.Hand),
            Background = new SolidColorBrush(Color.Parse(bg)),
            Foreground = new SolidColorBrush(Color.Parse(fg)),
            BorderThickness = new Thickness(1),
            BorderBrush = new SolidColorBrush(Color.Parse(bd)),
        };

        var btnPrev  = MakeBtn("◀", "#2a3a5a", "#88aacc", "#3a5a8a");
        var btnNext  = MakeBtn("▶", "#2a3a5a", "#88aacc", "#3a5a8a");
        var btnToday = MakeBtn("이번달", "#2a3a2a", "#88cc88", "#3a6a3a", 52);
        var btnApply = MakeBtn("반영", "#4a2a1a", "#ffaa66", "#6a4a2a", 48);
        btnApply.FontFamily = kbFontM;

        var txbMonth = new TextBlock {
            Text = _chartRangeStart.ToString("yyyy년 MM월"), FontSize = 12,
            FontFamily = kbFont, Foreground = new SolidColorBrush(Color.Parse("#ccc")),
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 0, 0) };

        header.Children.Add(btnPrev);
        header.Children.Add(txbMonth);
        header.Children.Add(btnNext);
        header.Children.Add(btnToday);
        header.Children.Add(btnApply);
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // ── 본문: StackPanel 기반 행 ────────────────────────────────────────
        var body = new StackPanel { Spacing = 0 };
        var scroll = new ScrollViewer {
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Content = body, Margin = new Thickness(4, 0, 4, 4) };
        Grid.SetRow(scroll, 1);
        root.Children.Add(scroll);

        RefreshAssignmentChartSlider(body);

        // ── 이벤트 ─────────────────────────────────────────────────────────
        void Navigate()
        {
            txbMonth.Text = _chartRangeStart.ToString("yyyy년 MM월");
            RefreshAssignmentChartSlider(body);
            btnApply.IsEnabled = true; btnApply.Content = "반영";
        }
        btnPrev.Click += (_, _) =>
        { _chartRangeStart = _chartRangeStart.AddMonths(-1); _chartRangeEnd = _chartRangeStart.AddMonths(1).AddDays(-1); Navigate(); };
        btnNext.Click += (_, _) =>
        { _chartRangeStart = _chartRangeStart.AddMonths(1); _chartRangeEnd = _chartRangeStart.AddMonths(1).AddDays(-1); Navigate(); };
        btnToday.Click += (_, _) =>
        { _chartRangeStart = new DateTime(today.Year, today.Month, 1); _chartRangeEnd = _chartRangeStart.AddMonths(1).AddDays(-1); Navigate(); };
        btnApply.Click += (_, _) =>
        {
            AnalysisRequestService.AutoExtendAssignmentsToToday();
            RefreshAssignmentChartSlider(body);
            RefreshShow3AfterChartUpdate();
            btnApply.IsEnabled = false; btnApply.Content = "반영됨";
        };

        return root;
    }

    /// <summary>슬라이더 기반 업무분장표 렌더링</summary>
    private void RefreshAssignmentChartSlider(StackPanel body)
    {
        body.Children.Clear();
        var kbFont  = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");
        var kbFontM = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

        // 표시 범위: 해당월 -5일 ~ +5일
        DateTime dispStart = _chartRangeStart.AddDays(-5);
        DateTime dispEnd   = _chartRangeEnd.AddDays(5);
        int dispDays = (int)(dispEnd - dispStart).TotalDays + 1;
        int monthDays = (int)(_chartRangeEnd - _chartRangeStart).TotalDays + 1;

        var spans    = AnalysisRequestService.GetAssignmentChartData(dispStart, dispEnd);
        var analytes = AnalysisRequestService.GetOrderedAnalytes()
            .Where(a => !a.fullName.StartsWith("_") && a.fullName != "기타업무" && a.fullName != "담당계약업체")
            .ToList();
        if (analytes.Count == 0) return;

        // 담당자 색상
        var mc = new Dictionary<string, (string Bg, string Fg)>();
        string[] bgP = ["#2a4a6a","#4a2a5a","#2a5a3a","#5a3a2a","#3a3a5a","#4a5a2a","#5a2a4a","#2a5a5a","#5a4a3a","#3a5a4a"];
        string[] fgP = ["#88bbee","#cc88dd","#88dd99","#ddaa88","#9999dd","#bbcc88","#dd88bb","#88ddcc","#ddbb88","#88ccaa"];
        int ci = 0;

        double trackW = dispDays * GC_DAY_W;

        // ── 날짜 눈금 (상단) ────────────────────────────────────────────────
        var dateRow = new Grid { ColumnDefinitions = new ColumnDefinitions($"{GC_LABEL_W},*"), Height = GC_DATE_H };
        var dateCanvas = new Canvas { Width = trackW, Height = GC_DATE_H, ClipToBounds = true };
        Grid.SetColumn(dateCanvas, 1);
        for (int d = 0; d < dispDays; d++)
        {
            var dt = dispStart.AddDays(d);
            bool inMonth = dt >= _chartRangeStart && dt <= _chartRangeEnd;
            bool isSun = dt.DayOfWeek == DayOfWeek.Sunday;
            bool isSat = dt.DayOfWeek == DayOfWeek.Saturday;
            string color = !inMonth ? "#444" : isSun ? "#cc6666" : isSat ? "#6688cc" : "#999";
            dateCanvas.Children.Add(new TextBlock
            {
                Text = dt.Day.ToString(), FontSize = 8, Width = GC_DAY_W,
                FontFamily = kbFont, TextAlignment = TextAlignment.Center,
                Foreground = new SolidColorBrush(Color.Parse(color)),
                [Canvas.LeftProperty] = d * GC_DAY_W, [Canvas.TopProperty] = 2.0,
            });
        }
        dateRow.Children.Add(dateCanvas);
        body.Children.Add(dateRow);

        // ── 항목별 행 ───────────────────────────────────────────────────────
        foreach (var (fullName, shortName) in analytes)
        {
            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions($"{GC_LABEL_W},*"),
                Height = GC_ROW_H, Margin = new Thickness(0, 0, 0, 1),
            };

            // 라벨
            var (lbBg, lbFg) = BadgeColorHelper.GetBadgeColor(shortName);
            var label = new Border
            {
                Height = GC_BAR_H, CornerRadius = new CornerRadius(3),
                Background = new SolidColorBrush(Color.Parse(lbBg)),
                Padding = new Thickness(3, 0), VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = shortName, FontSize = 8, FontFamily = kbFontM,
                    Foreground = new SolidColorBrush(Color.Parse(lbFg)),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                },
            };
            Grid.SetColumn(label, 0);
            rowGrid.Children.Add(label);

            // 트랙 캔버스
            var track = new Canvas { Width = trackW, Height = GC_ROW_H, ClipToBounds = true };
            Grid.SetColumn(track, 1);
            double barY = (GC_ROW_H - GC_BAR_H) / 2;

            // 월 외 구간 어둡게
            int preOff = (int)(_chartRangeStart - dispStart).TotalDays;
            int postOff = (int)(_chartRangeEnd - dispStart).TotalDays + 1;
            if (preOff > 0)
                track.Children.Add(new Avalonia.Controls.Shapes.Rectangle
                { Width = preOff * GC_DAY_W, Height = GC_ROW_H, Fill = new SolidColorBrush(Color.Parse("#0a0a0a")), Opacity = 0.5 });
            if (postOff < dispDays)
                track.Children.Add(new Avalonia.Controls.Shapes.Rectangle
                {
                    Width = (dispDays - postOff) * GC_DAY_W, Height = GC_ROW_H,
                    Fill = new SolidColorBrush(Color.Parse("#0a0a0a")), Opacity = 0.5,
                    [Canvas.LeftProperty] = postOff * GC_DAY_W,
                });

            // 오늘 하이라이트
            if (DateTime.Today >= dispStart && DateTime.Today <= dispEnd)
            {
                int ti = (int)(DateTime.Today - dispStart).TotalDays;
                track.Children.Add(new Avalonia.Controls.Shapes.Rectangle
                { Width = GC_DAY_W, Height = GC_ROW_H, Fill = new SolidColorBrush(Color.Parse("#1a2a1a")),
                  [Canvas.LeftProperty] = ti * GC_DAY_W });
            }

            // 수평 트랙 라인
            track.Children.Add(new Avalonia.Controls.Shapes.Line
            {
                StartPoint = new Point(0, barY + GC_BAR_H / 2),
                EndPoint = new Point(trackW, barY + GC_BAR_H / 2),
                Stroke = new SolidColorBrush(Color.Parse("#333")), StrokeThickness = 1,
            });

            // 스팬 바 (담당자별 색상 바 + 이름)
            var itemSpans = spans
                .Where(s => string.Equals(s.FullName, fullName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(s => s.Start).ToList();

            foreach (var sp in itemSpans)
            {
                if (!mc.TryGetValue(sp.Manager, out var c))
                { c = (bgP[ci % bgP.Length], fgP[ci % fgP.Length]); mc[sp.Manager] = c; ci++; }

                double sx = Math.Max(0, (sp.Start - dispStart).TotalDays) * GC_DAY_W;
                double ex = (Math.Min(dispDays - 1, (sp.End - dispStart).TotalDays) + 1) * GC_DAY_W;
                double w = ex - sx;
                if (w < 1) continue;

                track.Children.Add(new Border
                {
                    Width = w, Height = GC_BAR_H,
                    Background = new SolidColorBrush(Color.Parse(c.Item1)),
                    CornerRadius = new CornerRadius(3), Padding = new Thickness(4, 0), ClipToBounds = true,
                    Child = new TextBlock
                    {
                        Text = sp.Manager, FontSize = 9, FontFamily = kbFont,
                        Foreground = new SolidColorBrush(Color.Parse(c.Item2)),
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    [Canvas.LeftProperty] = sx, [Canvas.TopProperty] = barY,
                });
            }

            // 경계 슬라이더 (드래그 가능한 경계선 + 메모 라벨)
            for (int i = 0; i < itemSpans.Count - 1; i++)
            {
                var ls = itemSpans[i];
                var rs = itemSpans[i + 1];
                DateTime bnd = rs.Start;
                double bx = (bnd - dispStart).TotalDays * GC_DAY_W;

                // 경계 수직선
                track.Children.Add(new Avalonia.Controls.Shapes.Line
                {
                    StartPoint = new Point(bx, barY - 2), EndPoint = new Point(bx, barY + GC_BAR_H + 2),
                    Stroke = new SolidColorBrush(Color.Parse("#ffaa44")), StrokeThickness = 2,
                });

                // 메모 라벨 (엑셀 메모 스타일 — 경계선 위)
                string memoText = $"{ls.Manager} → {rs.Manager}\n{bnd:MM/dd}";
                var memo = new Border
                {
                    Background = new SolidColorBrush(Color.Parse("#332a1a")),
                    BorderBrush = new SolidColorBrush(Color.Parse("#664a2a")),
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(4, 1), IsVisible = false,
                    Child = new TextBlock
                    {
                        Text = memoText, FontSize = 8, FontFamily = kbFont,
                        Foreground = new SolidColorBrush(Color.Parse("#ffcc88")),
                    },
                };
                Canvas.SetLeft(memo, bx - 30);
                Canvas.SetTop(memo, barY - 32);
                track.Children.Add(memo);

                // 드래그 핸들 (투명 넓은 영역)
                var capturedBody = body;
                var capturedFull = fullName;
                var capturedLMgr = ls.Manager;
                var capturedRMgr = rs.Manager;
                var capturedBnd  = bnd;
                var capturedMemo = memo;

                var handle = new Border
                {
                    Width = 14, Height = GC_BAR_H + 10,
                    Background = Brushes.Transparent,
                    Cursor = new Cursor(StandardCursorType.SizeWestEast),
                };
                Canvas.SetLeft(handle, bx - 7);
                Canvas.SetTop(handle, barY - 5);
                track.Children.Add(handle);

                double dragSx = 0; bool dragging = false;

                handle.PointerEntered += (_, _) => capturedMemo.IsVisible = true;
                handle.PointerExited += (_, _) => { if (!dragging) capturedMemo.IsVisible = false; };

                handle.PointerPressed += (s, e) =>
                {
                    if (s is not Border b) return;
                    dragging = true; dragSx = e.GetPosition(track).X;
                    e.Pointer.Capture(b); e.Handled = true;
                };
                handle.PointerMoved += (s, e) =>
                {
                    if (!dragging || s is not Border b) return;
                    double cur = e.GetPosition(track).X;
                    double nc = Canvas.GetLeft(b) + 7 + (cur - dragSx);
                    double minX = preOff * GC_DAY_W + GC_DAY_W;
                    double maxX = postOff * GC_DAY_W - GC_DAY_W;
                    nc = Math.Clamp(nc, minX, maxX);
                    Canvas.SetLeft(b, nc - 7);
                    dragSx = cur;
                    int dayIdx = (int)(nc / GC_DAY_W);
                    var nd = dispStart.AddDays(dayIdx);
                    ((capturedMemo.Child as TextBlock)!).Text = $"{capturedLMgr} → {capturedRMgr}\n{nd:MM/dd}";
                    Canvas.SetLeft(capturedMemo, nc - 30);
                    capturedMemo.IsVisible = true;
                    e.Handled = true;
                };
                handle.PointerReleased += (s, e) =>
                {
                    if (!dragging || s is not Border b) return;
                    dragging = false; e.Pointer.Capture(null);
                    capturedMemo.IsVisible = false;
                    double fc = Canvas.GetLeft(b) + 7;
                    int dayIdx = (int)(fc / GC_DAY_W);
                    var nb = dispStart.AddDays(dayIdx);
                    if (nb != capturedBnd && nb >= _chartRangeStart && nb <= _chartRangeEnd)
                    {
                        if (nb > capturedBnd)
                            AnalysisRequestService.UpdateAssignmentByName(capturedFull, capturedLMgr, capturedBnd, nb.AddDays(-1));
                        else
                            AnalysisRequestService.UpdateAssignmentByName(capturedFull, capturedRMgr, nb, capturedBnd.AddDays(-1));
                        RefreshAssignmentChartSlider(capturedBody);
                        RefreshShow3AfterChartUpdate();
                    }
                    else RefreshAssignmentChartSlider(capturedBody);
                    e.Handled = true;
                };
            }

            // "+" 미배정 구간 (미래)
            DateTime gapStart = itemSpans.Count > 0 ? itemSpans[^1].End.AddDays(1) : _chartRangeStart;
            if (gapStart <= DateTime.Today && DateTime.Today < _chartRangeEnd)
                gapStart = DateTime.Today.AddDays(1);
            if (gapStart <= _chartRangeEnd)
            {
                double gx = (gapStart - dispStart).TotalDays * GC_DAY_W;
                double ge = (postOff) * GC_DAY_W;
                double gcx = (gx + ge) / 2;
                var captFull2 = fullName; var captStart2 = gapStart; var captBody2 = body;
                var plus = new Button
                {
                    Width = 20, Height = 20, Content = "+", FontSize = 12,
                    Padding = new Thickness(0), CornerRadius = new CornerRadius(10),
                    Background = new SolidColorBrush(Color.Parse("#2a3a2a")),
                    Foreground = new SolidColorBrush(Color.Parse("#88cc88")),
                    BorderBrush = new SolidColorBrush(Color.Parse("#4a6a4a")),
                    BorderThickness = new Thickness(1), FontFamily = kbFontM,
                    Cursor = new Cursor(StandardCursorType.Hand),
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center,
                };
                Canvas.SetLeft(plus, gcx - 10);
                Canvas.SetTop(plus, barY);
                track.Children.Add(plus);

                var names = AgentService.GetAllNames();
                var lb = new ListBox { MaxHeight = 200, MinWidth = 120, ItemsSource = names,
                    FontSize = 11, FontFamily = kbFont,
                    Background = new SolidColorBrush(Color.Parse("#1a1a2a")),
                    Foreground = new SolidColorBrush(Color.Parse("#ccc")) };
                var fly = new Flyout { Content = lb, Placement = PlacementMode.Bottom };
                plus.Flyout = fly;
                lb.SelectionChanged += (_, _) =>
                {
                    if (lb.SelectedItem is not string nm) return;
                    AnalysisRequestService.UpdateAssignmentByName(captFull2, nm, captStart2, _chartRangeEnd);
                    fly.Hide();
                    RefreshAssignmentChartSlider(captBody2);
                    RefreshShow3AfterChartUpdate();
                };
            }

            rowGrid.Children.Add(track);
            body.Children.Add(rowGrid);
        }
    }

    /// <summary>분장표 변경 후 Show3 (타임라인/이력) 리프레시</summary>
    private void RefreshShow3AfterChartUpdate()
    {
        if (_selectedAgent != null)
        {
            Show3ContentRequest?.Invoke(BuildShow3Timeline(_selectedAgent));
            if (_항목ChipsPanel != null)
                RefreshItemChips(_selectedAgent);
        }
    }

    /// <summary>사진 파일을 DB에 바이너리로 동기화 (다른 PC에서도 사진 보이게)</summary>
    private static void SyncPhotoToDb(string 사번, string photoPath)
    {
        if (string.IsNullOrEmpty(사번) || string.IsNullOrEmpty(photoPath)) return;
        try
        {
            var dir = AgentService.GetPhotoDirectory();
            var fullPath = Path.IsPathRooted(photoPath) ? photoPath : Path.Combine(dir, photoPath);
            if (File.Exists(fullPath))
            {
                var data = File.ReadAllBytes(fullPath);
                AgentService.SavePhotoToDb(사번, data);
            }
        }
        catch { }
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [AgentTree] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("Logs/AgentDebug.log", line + Environment.NewLine); } catch { }
    }
}
