using Avalonia;
using Avalonia.Controls;
using Ellipse = Avalonia.Controls.Shapes.Ellipse;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using System;
using System.IO;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.Views;
using ETA.Views.Controls;

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
    public Action? Show4RefreshRequest { get; set; }
    public ListBox? AnalysisItemsListBox { get; set; }

    // ── 상태 ────────────────────────────────────────────────────────────────
    private Agent?      _selectedAgent;
    private StackPanel? _detailPanel;
    private bool        _isAddMode  = false;

    // ── 권한 관리 ────────────────────────────────────────────────────────────
    // 편집/직급수정: 정승욱, 박은지
    private static readonly HashSet<string> _editAllowedNames =
        new(StringComparer.Ordinal) { "정승욱", "박은지" };
    // 삭제: 정승욱, 박은지, 방찬미, 이예림
    private static readonly HashSet<string> _deleteAllowedNames =
        new(StringComparer.Ordinal) { "정승욱", "박은지", "방찬미", "이예림" };

    private static string? _currentUserNameCache;
    private static string CurrentUserName
    {
        get
        {
            if (_currentUserNameCache != null) return _currentUserNameCache;
            var empId = MainPage.CurrentEmployeeId;
            var agent = AgentService.GetAllItems().FirstOrDefault(a => a.사번 == empId);
            _currentUserNameCache = agent?.성명 ?? "";
            return _currentUserNameCache;
        }
    }

    /// <summary>직급 수정 등 편집 권한 (정승욱, 박은지)</summary>
    private static bool CanEdit => _editAllowedNames.Contains(CurrentUserName);

    /// <summary>직원 삭제 권한 (정승욱, 박은지, 방찬미, 이예림)</summary>
    private static bool CanDelete => _deleteAllowedNames.Contains(CurrentUserName);

    // 사진 미리보기 Image 컨트롤 (저장 시 PhotoPath 접근용)
    private Image?      _photoImage;
    private string      _pendingPhotoPath = "";   // 선택했지만 아직 저장 안 된 경로

    // 업무 분장 저장 시 참조 (BuildAssignmentArea에서 설정)
    private ListBox?      _assignmentListBox;
    private Canvas?       _timelineCanvas;
    private ScrollViewer? _timelineScroll;
    private DateTime    _assignmentRangeStart = DateTime.Today;
    private DateTime    _assignmentRangeEnd   = DateTime.Today;

    // 담당항목/업체/기타업무/일반업무 패널 (Show2)
    private WrapPanel?  _항목ChipsPanel;
    private WrapPanel?  _업체ChipsPanel;
    private StackPanel? _miscTaskPanel;
    private StackPanel? _generalTaskPanel;

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
    private const double GC_LABEL_W   = 120.0;
    private const double GC_DAY_W     = 22.0;
    private const double GC_ROW_H     = 38.0;   // 행 높이 (여유 있게)
    private const double GC_BAR_H     = 24.0;   // 바 높이
    private const double GC_DATE_H    = 20.0;   // 하단 날짜 영역
    private string _chartTab = "분석항목";          // 현재 업무분장표 탭
    private StackPanel? _chartSummaryPanel;           // wire-v01 요약 스트립
    // 드래그로 변경된 경계 — (항목명, 원래시작일, 원래종료일, 새시작일, 새종료일, 담당자)
    private readonly List<(string FullName, string Manager, DateTime OrigStart, DateTime OrigEnd, DateTime NewStart, DateTime NewEnd)> _chartPendingChanges = new();
    private readonly List<Action<DbConnection, DbTransaction, List<string>>> _chartPendingActions = new();
    private ProgressBar? _chartProgressBar;
    private TextBlock? _chartProgressLabel;
    /// <summary>로컬 캐시 — null이면 DB에서 로드, 편집 시 로컬만 수정</summary>
    private List<AnalysisRequestService.AssignmentSpan>? _chartCachedSpans;

    public AgentTreePage()
    {
        InitializeComponent();
        DeptFilterBox.ItemsSource   = _deptItems;
        DeptFilterBox.SelectedIndex = 0;
    }

    // =========================================================================
    // 검색 / 부서 필터
    // =========================================================================
    private static readonly string[] _deptItems =
        { "전체", "수질분석센터", "처리시설", "일반업무", "기타" };

    private DispatcherTimer? _searchDebounce;
    private string _lastAgentSearch = "";
    private string _deptFilter = "전체";

    private void AgentSearchBox_TextChanged(object? sender, Avalonia.Controls.TextChangedEventArgs e)
    {
        _searchDebounce?.Stop();
        _searchDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
        _searchDebounce.Tick += (_, _) =>
        {
            _searchDebounce!.Stop();
            DoAgentSearch();
        };
        _searchDebounce.Start();
    }

    protected override void OnLoaded(RoutedEventArgs e)
    {
        base.OnLoaded(e);
    }

    private void DeptFilter_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        _deptFilter = DeptFilterBox.SelectedItem?.ToString() ?? "전체";
        ApplyFilters();
    }

    private void DoAgentSearch()
    {
        var q = AgentSearchBox.Text?.Trim().ToLower() ?? "";
        if (q == _lastAgentSearch) return;
        _lastAgentSearch = q;
        ApplyFilters();
    }

    private void ApplyFilters()
    {
        var q    = AgentSearchBox.Text?.Trim().ToLower() ?? "";
        var dept = _deptFilter;
        foreach (var tvi in AgentTreeView.Items.OfType<TreeViewItem>())
        {
            if (tvi.Tag is not Agent agent) continue;
            bool nameMatch = string.IsNullOrEmpty(q)
                || (agent.성명?.ToLower().Contains(q) ?? false)
                || (agent.사번?.ToLower().Contains(q) ?? false)
                || (agent.직급?.ToLower().Contains(q) ?? false);
            bool deptMatch = dept == "전체" || agent.부서 == dept;
            tvi.IsVisible = nameMatch && deptMatch;
        }
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
        _assignmentRangeStart = DateTime.Today;
        _assignmentRangeEnd   = DateTime.Today;
        DetailPanelChanged?.Invoke(null);

        try
        {
            var allItems = AgentService.GetAllItems().OrderBy(a => a.입사일).ToList();

            // 편집 권한자(관리자)는 전체 조회, 일반 사용자는 본인 부서만
            var dept = ETA.Services.Common.CurrentUserManager.Instance.CurrentDepartment;
            var items = (CanEdit || string.IsNullOrWhiteSpace(dept))
                ? allItems
                : allItems.Where(a => a.부서 == dept).ToList();

            foreach (var item in items)
                AgentTreeView.Items.Add(CreateTreeItem(item));
            Log($"로드 완료 → {items.Count}명 (부서필터: {(CanEdit ? "전체" : dept)})");
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
                    FontSize          = AppTheme.FontSM,
                    Foreground        = new SolidColorBrush(Color.Parse(fg)),
                    FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    VerticalAlignment = VerticalAlignment.Center,
                }
            });
        }

        // 이름
        headerPanel.Children.Add(new TextBlock
        {
            Text              = agent.성명,
            FontSize          = AppTheme.FontLG,
            FontFamily        = "avares://ETA/Assets/Fonts#Pretendard",
            Foreground        = AppRes("AppFg"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        var tvi = new TreeViewItem
        {
            Tag    = agent,
            Header = headerPanel,
        };
        TextShimmer.AttachHover(tvi);

        // 직원을 업무분장표로 드래그 가능하게
        tvi.PointerPressed += async (s, e) =>
        {
            if (e.ClickCount > 1) return; // 더블클릭 무시
            var props = e.GetCurrentPoint((Control)s!).Properties;
            if (!props.IsLeftButtonPressed) return;

            var data = new DataObject();
            data.Set("agent_name", agent.성명);
            await DragDrop.DoDragDrop(e, data, DragDropEffects.Copy);
        };

        return tvi;
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
                FontFamily          = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
        if (_selectedAgent == null) { Log("삭제 스킵: 선택 없음"); return; }
        await DeleteAgentAsync(_selectedAgent);
    }

    private async Task DeleteAgentAsync(Agent agent)
    {
        if (!CanDelete) return;

        // 확인 다이얼로그
        var dlg = new Window
        {
            Title           = "삭제 확인",
            Width           = 320,
            Height          = 150,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize       = false,
            Background      = AppTheme.BgSecondary,
        };

        bool confirmed = false;

        var yesBtn = new Button
        {
            Content    = "삭제",
            Width      = 80,
            Background = AppTheme.FgDanger,
            Foreground = Brushes.White,
        };
        var noBtn = new Button
        {
            Content    = "취소",
            Width      = 80,
            Background = AppTheme.BorderMuted,
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
                    Text       = $"'{agent.성명}' 직원을 삭제하시겠습니까?",
                    Foreground = AppRes("AppFg"),
                    FontSize   = AppTheme.FontLG,
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

        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null)
            await dlg.ShowDialog(owner);
        else
            dlg.Show();

        if (!confirmed) return;

        bool ok = AgentService.Delete(agent);
        Log(ok ? $"✅ 삭제 성공: {agent.성명}" : $"❌ 삭제 실패: {agent.성명}");

        if (ok)
        {
            var toRemove = AgentTreeView.Items
                .OfType<TreeViewItem>()
                .FirstOrDefault(i => i.Tag == agent);
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

        // 사진 + 성명/직급/사번 한 줄 (상단)
        var topRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, Margin = new Thickness(0, 0, 0, 4) };
        topRow.Children.Add(BuildPhotoArea(agent.PhotoPath, agent.사번));
        var topInfo = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,8,*,8,*"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        topInfo.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        var f성명 = BuildFieldRow("성명", agent.성명, isReadOnly: true, isLocked: true);
        var f직급 = BuildFieldRow("직급", agent.직급, isReadOnly: ro);
        var f사번 = BuildFieldRow("사번", agent.사번, isReadOnly: ro);
        Grid.SetColumn(f성명, 0); Grid.SetColumn(f직급, 2); Grid.SetColumn(f사번, 4);
        topInfo.Children.Add(f성명); topInfo.Children.Add(f직급); topInfo.Children.Add(f사번);
        topRow.Children.Add(topInfo);
        root.Children.Add(topRow);

        // ── 개발자 모드: 선택 인원으로 로그인 ──
        if (CanEdit && agent.사번 != MainPage.CurrentEmployeeId)
        {
            var devLoginBtn = new Button
            {
                Content    = $"🔑 [{agent.성명}](으)로 로그인 (개발자 모드)",
                FontSize   = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Background = AppTheme.BorderInfo,
                Foreground = new SolidColorBrush(Color.Parse("#8cf")),
                Padding    = new Thickness(10, 4),
                Margin     = new Thickness(0, 4, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left,
                Cursor     = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            };
            devLoginBtn.Click += (_, _) =>
            {
                var prevId = MainPage.CurrentEmployeeId;
                MainPage.CurrentEmployeeId = agent.사번;
                CurrentUserManager.Instance.SetCurrentUser(agent.사번);
                _currentUserNameCache = null; // 권한 캐시 초기화

                // 타이틀 바에 개발자 모드 표시
                if (this.VisualRoot is Window win)
                    win.Title = $"ETA — 개발자 모드: {agent.성명} ({agent.사번})  [원래: {prevId}]";

                Log($"[DevMode] {prevId} → {agent.사번} ({agent.성명}) 로 전환");

                // UI 새로고침
                _detailPanel = BuildEditPanel(agent);
                DetailPanelChanged?.Invoke(_detailPanel);
            };
            root.Children.Add(devLoginBtn);
        }

        // 직무/자격사항/Email 한 줄 + 부서/입사일/측정인고유번호 한 줄
        root.Children.Add(BuildFieldGrid3(
            BuildFieldRow("직무",     agent.직무,     isReadOnly: ro),
            BuildFieldRow("자격사항", agent.자격사항, isReadOnly: ro),
            BuildFieldRow("Email",    agent.Email,    isReadOnly: ro)
        ));
        root.Children.Add(BuildFieldGrid(
            BuildComboRow("부서", new[] { "수질분석센터", "처리시설", "일반업무", "기타" }, agent.부서, isReadOnly: ro),
            BuildFieldRow("측정인고유번호", agent.측정인고유번호, isReadOnly: ro)
        ));
        root.Children.Add(BuildFieldGrid(
            BuildFieldRow("입사일",         agent.입사일표시,     isReadOnly: true),
            new StackPanel()
        ));

        // ── 직원 삭제 버튼 (권한 있는 경우만) ──
        if (CanDelete)
        {
            var deleteBtn = new Button
            {
                Content             = $"🗑  {agent.성명} 삭제",
                FontSize            = AppTheme.FontBase,
                FontFamily          = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Background          = AppTheme.BgDanger,
                Foreground          = AppTheme.FgDanger,
                BorderBrush         = AppTheme.BorderDanger,
                BorderThickness     = new Thickness(1),
                CornerRadius        = new CornerRadius(4),
                Padding             = new Thickness(10, 4),
                Margin              = new Thickness(0, 6, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Left,
                Cursor              = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            };
            var capturedAgent = agent;
            deleteBtn.Click += async (_, _) => await DeleteAgentAsync(capturedAgent);
            root.Children.Add(deleteBtn);
        }

        // ── 담당 분석항목 ──
        var itemSection = new Border
        {
            Background   = AppTheme.BgPrimary,
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

        // 드래그앤드랍: Show4에서 분석항목을 끌어다 놓으면 배정
        DragDrop.SetAllowDrop(itemSection, true);
        itemSection.AddHandler(DragDrop.DragOverEvent, (object? s, DragEventArgs e) =>
        {
            e.DragEffects = e.Data.Contains("analyte") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        });
        itemSection.AddHandler(DragDrop.DropEvent, (object? s, DragEventArgs e) =>
        {
            if (!e.Data.Contains("analyte")) return;
            var analyte = e.Data.Get("analyte") as string;
            if (string.IsNullOrEmpty(analyte)) return;

            // 이미 배정된 항목이면 무시
            if (agent.담당항목목록.Contains(analyte))
            {
                Log($"드랍 무시: {analyte} 이미 배정됨");
                return;
            }

            // 담당항목에 추가
            var list = agent.담당항목목록.ToList();
            list.Add(analyte);
            agent.담당항목 = string.Join(",", list);
            AgentService.Update(agent);

            // 분장표준처리에 금일부터 배정
            AssignItemsForDate(agent, DateTime.Today);
            RefreshItemChips(agent);
            Show4RefreshRequest?.Invoke();
            Log($"드랍 배정: {agent.성명} ← {analyte}");
        });

        root.Children.Add(itemSection);

        // ── 담당 계약업체 ──
        var contractSection = new Border
        {
            Background   = AppTheme.BgActiveGreen,
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

        // ── 일반업무 ──
        var generalTaskSection = new Border
        {
            Background   = new SolidColorBrush(Color.Parse("#1a2a2a")),
            CornerRadius = new CornerRadius(4),
            Padding      = new Thickness(8, 6),
            Margin       = new Thickness(0, 6, 0, 0),
        };
        var generalTaskInner = new StackPanel { Spacing = 3 };
        _generalTaskPanel = new StackPanel { Spacing = 3 };
        generalTaskInner.Children.Add(BuildChipSectionHeader("📝 담당 일반업무",
            () => Show4ContentRequest?.Invoke(BuildShow4GeneralTaskAssignList(agent))));
        generalTaskInner.Children.Add(_generalTaskPanel);
        generalTaskSection.Child = generalTaskInner;
        root.Children.Add(generalTaskSection);
        RefreshGeneralTasks(agent);

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
        var topInfo = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,8,*,8,*"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        topInfo.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        var an성명 = BuildFieldRow("성명", "", hint: "이름 입력 (필수)");
        var an직급 = BuildFieldRow("직급", "");
        var an사번 = BuildFieldRow("사번", "");
        Grid.SetColumn(an성명, 0); Grid.SetColumn(an직급, 2); Grid.SetColumn(an사번, 4);
        topInfo.Children.Add(an성명); topInfo.Children.Add(an직급); topInfo.Children.Add(an사번);
        topRow.Children.Add(topInfo);
        root.Children.Add(topRow);

        root.Children.Add(BuildFieldGrid3(
            BuildFieldRow("직무",     ""),
            BuildFieldRow("자격사항", ""),
            BuildFieldRow("Email",    "")
        ));
        root.Children.Add(BuildFieldGrid(
            BuildComboRow("부서", new[] { "수질분석센터", "처리시설", "일반업무", "기타" }, ""),
            BuildFieldRow("측정인고유번호", "")
        ));
        root.Children.Add(BuildFieldGrid(
            BuildFieldRow("입사일",         "", hint: "예) 2024-01-01"),
            new StackPanel()
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
            Background   = AppTheme.BgActiveGreen,
            CornerRadius = new CornerRadius(6),
            Padding      = new Thickness(10, 8),
            Margin       = new Thickness(0, 8, 0, 0)
        };

        var stack = new StackPanel { Spacing = 4 };

        // ── 제목 ──────────────────────────────────────────────────────────
        stack.Children.Add(new TextBlock
        {
            Text       = "📋 업무 분장",
            FontSize   = AppTheme.FontMD,
            FontWeight = FontWeight.Bold,
            Foreground = new SolidColorBrush(Color.Parse("#7cd87c")),
            FontFamily = "avares://ETA/Assets/Fonts#Pretendard"
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
            FontSize          = AppTheme.FontBase,
            Foreground        = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            MinWidth          = 170,
            FontFamily        = "avares://ETA/Assets/Fonts#Pretendard"
        };

        var btnCal = new Button
        {
            Content         = "📅",
            Width           = 30,
            Height          = 24,
            FontSize        = AppTheme.FontMD,
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
            Width           = 44,
            Height          = 24,
            FontSize        = AppTheme.FontSM,
            Padding         = new Thickness(4, 0),
            Background      = new SolidColorBrush(Color.Parse("#3a5a3a")),
            Foreground      = AppTheme.FgMuted,
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderMuted
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
            Focusable     = true,
        };
        calendar.SelectedDates.Add(DateTime.Today);
        calendar.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape && calendar.IsVisible)
            { calendar.IsVisible = false; e.Handled = true; }
        };
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

        // 날짜 범위 상태 (클로저로 공유)
        DateTime rangeStart = DateTime.Today;
        DateTime rangeEnd   = DateTime.Today;

        // DragOver 핸들러 — 드롭 허용 표시 (이것이 없으면 드롭 거부됨)
        timelineScroll.AddHandler(DragDrop.DragOverEvent, (object? sender, DragEventArgs e) =>
        {
            e.DragEffects = e.Data.Contains("analyte") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        });

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
            Show4RefreshRequest?.Invoke();
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
            Background      = AppTheme.BgActiveGreen,
        };
        stack.Children.Add(saveProgress);

        // ── 분장 저장 버튼 ───────────────────────────────────────────────
        var btnAssignSave = new Button
        {
            Content             = "💾 분장 저장",
            Height              = 28,
            FontSize            = AppTheme.FontBase,
            FontFamily          = "avares://ETA/Assets/Fonts#Pretendard",
            Background          = AppTheme.BorderActive,
            Foreground          = AppRes("AppFg"),
            BorderThickness     = new Thickness(1),
            BorderBrush         = AppTheme.BorderActive,
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
            Background            = AppTheme.BgSecondary,
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
            Background = AppTheme.BorderMuted,
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
                    FontSize     = AppTheme.FontMD,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    FontFamily   = "avares://ETA/Assets/Fonts#Pretendard",
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
                Foreground = AppTheme.FgDimmed
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
                    FontSize   = AppTheme.FontXS,
                    Foreground = new SolidColorBrush(Color.Parse("#7ab4cc")),
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                },
            });
            topRow.Children.Add(new TextBlock
            {
                Text              = fullName,
                FontSize          = AppTheme.FontBase,
                FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground        = AppTheme.FgPrimary,
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
                FontSize   = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
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
                    FontSize   = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
                Fill   = AppTheme.BorderSubtle,
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
                    FontSize          = AppTheme.FontXS,
                    FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
            BorderBrush     = AppTheme.BorderDefault,
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
            FontSize        = AppTheme.FontBase,
            FontFamily      = "avares://ETA/Assets/Fonts#Pretendard",
            Background      = new SolidColorBrush(Color.Parse("#3a4a6a")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };

        var removeBtn = new Button
        {
            Content         = "🗑 사진 제거",
            FontSize        = AppTheme.FontBase,
            FontFamily      = "avares://ETA/Assets/Fonts#Pretendard",
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
            FontSize   = AppTheme.FontSM,
            Foreground = AppTheme.FgDimmed
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
            else if (child is StackPanel sp2 && sp2.Children.Count >= 2
                && sp2.Children[0] is TextBlock lbl2 && sp2.Children[1] is ComboBox cb)
            {
                var label2 = lbl2.Text?.Trim() ?? "";
                if (label2 == "부서")
                    agent.부서 = cb.SelectedItem?.ToString() ?? "";
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
            FontSize   = AppTheme.FontXL,
            FontFamily = "avares://ETA/Assets/Fonts#Pretendard",
            Foreground = AppRes("AppFg"),
            Margin     = new Thickness(0, 0, 0, 4)
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = AppTheme.BorderDefault,
            Margin     = new Thickness(0, 0, 0, 4)
        });
        return root;
    }

    private static StackPanel BuildComboRow(string label, string[] items, string selected,
                                             bool isReadOnly = false)
    {
        var panel = new StackPanel { Spacing = 2 };
        panel.Children.Add(new TextBlock
        {
            Text       = label,
            FontSize   = AppTheme.FontSM,
            FontFamily = "avares://ETA/Assets/Fonts#Pretendard",
            Foreground = AppRes("FgMuted"),
        });
        var cb = new ComboBox
        {
            ItemsSource             = items,
            SelectedItem            = items.Contains(selected) ? selected : (items.Length > 0 ? items[0] : null),
            FontSize                = AppTheme.FontMD,
            FontFamily              = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            IsEnabled               = !isReadOnly,
            Background              = isReadOnly ? AppTheme.BgSecondary : AppTheme.BgInput,
            Foreground              = isReadOnly ? AppTheme.FgMuted     : AppRes("AppFg"),
            BorderThickness         = new Thickness(1),
            BorderBrush             = isReadOnly ? AppTheme.BorderSubtle : AppTheme.BorderDefault,
            CornerRadius            = new CornerRadius(4),
            Padding                 = new Thickness(8, 4),
            HorizontalAlignment     = HorizontalAlignment.Stretch,
        };
        panel.Children.Add(cb);
        return panel;
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
            FontSize          = AppTheme.FontSM,
            FontFamily        = "avares://ETA/Assets/Fonts#Pretendard",
            Foreground        = isLocked ? AppTheme.FgMuted : AppRes("FgMuted"),
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            FontSize        = AppTheme.FontMD,
            FontFamily      = "avares://ETA/Assets/Fonts#Pretendard",
            IsReadOnly      = isReadOnly,
            Watermark       = hint,
            Background      = isReadOnly ? AppTheme.BgSecondary : AppTheme.BgInput,
            Foreground      = isReadOnly ? AppTheme.FgMuted     : AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly ? AppTheme.BorderSubtle : AppTheme.BorderDefault,
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

    /// <summary>필드를 3열 Grid로 배치</summary>
    private static Grid BuildFieldGrid3(params StackPanel[] fields)
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,8,*,8,*"),
            Margin = new Thickness(0, 2, 0, 2),
        };
        for (int i = 0; i < fields.Length; i++)
        {
            int col = (i % 3) * 2;     // 0, 2, 4
            int row = i / 3;
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
        var wrapper = new StackPanel { Spacing = 0 };

        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto"),
            Margin = new Thickness(0, 10, 0, 2),
        };
        row.Children.Add(new TextBlock
        {
            Text = title,
            FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        });

        // 인라인 캘린더 (토글)
        Border? calBorder = null;

        if (CanEdit)
        {
            var btnStack = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 3, [Grid.ColumnProperty] = 1 };

            if (onAssignDate != null)
            {
                var btnCal = new Button
                {
                    Content = "📅", FontSize = AppTheme.FontBase, Width = 24, Height = 24,
                    Padding = new Thickness(0),
                    Background = AppTheme.BgActiveGreen,
                    Foreground = AppTheme.FgSuccess,
                    BorderThickness = new Thickness(1),
                    BorderBrush = AppTheme.BorderActive,
                    CornerRadius = new CornerRadius(4),
                    HorizontalContentAlignment = HorizontalAlignment.Center,
                    VerticalContentAlignment = VerticalAlignment.Center,
                };
                calBorder = InlineCalendarHelper.Create(onAssignDate, btnCal);
                btnStack.Children.Add(btnCal);
            }

            var btnAdd = new Button
            {
                Content = "+", FontSize = AppTheme.FontLG, Width = 24, Height = 24,
                Padding = new Thickness(0),
                Background = AppTheme.BgActiveBlue,
                Foreground = AppTheme.FgInfo,
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderAccent,
                CornerRadius = new CornerRadius(4),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
            };
            btnAdd.Click += (_, _) => onAdd?.Invoke();
            btnStack.Children.Add(btnAdd);
            row.Children.Add(btnStack);
        }

        wrapper.Children.Add(row);
        if (calBorder != null) wrapper.Children.Add(calBorder);
        return wrapper;
    }

    /// <summary>agent.담당항목목록의 모든 항목을 지정 날짜의 분장표준처리에 배정</summary>
    private static void AssignItemsForDate(Agent agent, DateTime date)
    {
        if (agent.담당항목목록.Count == 0) return;
        foreach (var item in agent.담당항목목록)
            AnalysisRequestService.AddAssignment(agent.사번, item, date, date);
    }

    // =========================================================================
    // QAQC 첨부 파일 헬퍼
    // =========================================================================
    private static string GetQaqcDir(string yearMonth)
    {
        var dataDir = Path.GetDirectoryName(DbPathHelper.PhotoDirectory)!;
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
        var qaqcRoot = Path.Combine(Path.GetDirectoryName(DbPathHelper.PhotoDirectory)!, "QAQC");
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
        var font    = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        var root = new StackPanel { Spacing = 6, Margin = new Thickness(6) };

        // 헤더
        root.Children.Add(new TextBlock
        {
            Text = $"📅 {date:yyyy년 MM월} 배정 + 첨부",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            FontFamily = font, Foreground = new SolidColorBrush(Color.Parse("#aaccee")),
        });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(0, 0, 0, 4) });

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
                BorderBrush     = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(4),
                Padding         = new Thickness(6, 4),
            };
            var itemStack = new StackPanel { Spacing = 3 };

            // 항목명 행
            var nameRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
            nameRow.Children.Add(new Border
            {
                Background      = new SolidColorBrush(Color.Parse(chipBg)),
                BorderBrush     = new SolidColorBrush(Color.Parse(chipFg)),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(10),
                Padding         = new Thickness(6, 1, 8, 1),
                Margin          = new Thickness(0, 0, 4, 0),
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 0,
                Child = new TextBlock
                {
                    Text              = displayShort,
                    FontSize          = AppTheme.FontSM,
                    FontWeight        = FontWeight.Medium,
                    FontFamily        = font,
                    Foreground        = new SolidColorBrush(Color.Parse(chipFg)),
                    VerticalAlignment = VerticalAlignment.Center,
                },
            });
            nameRow.Children.Add(new TextBlock
            {
                Text = capturedFull, FontSize = AppTheme.FontSM, FontFamily = font,
                Foreground = new SolidColorBrush(Color.Parse("#ccccee")),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                [Grid.ColumnProperty] = 1,
            });

            // 📎 버튼
            var btnAttach = new Button
            {
                Content = "📎", FontSize = AppTheme.FontBase, Width = 24, Height = 24,
                Padding = new Thickness(0),
                Background = AppTheme.BgActiveBlue,
                Foreground = AppTheme.FgInfo,
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderAccent,
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
                        Text = "📄 " + display, FontSize = AppTheme.FontXS, FontFamily = font,
                        Foreground = AppTheme.FgSuccess,
                        VerticalAlignment = VerticalAlignment.Center,
                    });
                    var capturedF = fPath;
                    var delBtn = new Button
                    {
                        Content = "×", FontSize = AppTheme.FontXS, Width = 16, Height = 16, Padding = new Thickness(0),
                        Background = AppTheme.BgDanger,
                        Foreground = AppTheme.FgDanger,
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
                        Text = "첨부 없음", FontSize = AppTheme.FontXS, FontFamily = font,
                        Foreground = AppTheme.FgDimmed,
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
            FontFamily = font, FontSize = AppTheme.FontBase,
            Background = AppTheme.BgActiveGreen,
            Foreground = AppTheme.FgSuccess,
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderActive,
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
                Text = "없음", FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
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
                // 분장표준처리에서 금일부터 이 항목의 담당 제거
                AnalysisRequestService.ClearAnalyteFromAgent(agent.사번, captured, DateTime.Today);
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
                Text = "없음", FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
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
            Cursor          = new Cursor(StandardCursorType.Hand),
        };
        var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 3 };
        row.Children.Add(new TextBlock
        {
            Text = label, FontSize = AppTheme.FontSM,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = new SolidColorBrush(Color.Parse(fg)),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var btnX = new Button
        {
            Content = "×", FontSize = AppTheme.FontBase, Padding = new Thickness(2, 0),
            Background = Brushes.Transparent, BorderThickness = new Thickness(0),
            Foreground = new SolidColorBrush(Color.Parse("#dd6688")),
            VerticalAlignment = VerticalAlignment.Center,
            IsVisible = false,
        };
        btnX.Click += (_, e) => { e.Handled = true; onRemove?.Invoke(); };
        chip.PointerEntered += (_, _) => btnX.IsVisible = true;
        chip.PointerExited  += (_, _) => btnX.IsVisible = false;
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
            Background   = AppTheme.BgPrimary,
            CornerRadius = new CornerRadius(4),
            Padding      = new Thickness(8, 6),
            Margin       = new Thickness(0, 10, 0, 0),
        };
        var stack = new StackPanel { Spacing = 3 };
        stack.Children.Add(new TextBlock
        {
            Text = "📊 과거 분장이력 (전체)",
            FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 4),
        });

        var histGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,4,*,4,*,4,*,4,*"),
        };
        try
        {
            var assignments = AnalysisRequestService.GetAssignmentDaysForAgentAll(agent.사번);
            if (assignments.Count == 0)
            {
                histGrid.Children.Add(new TextBlock
                {
                    Text = "이력 없음", FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppTheme.FgDimmed,
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

                for (int i = 0; i < groups.Count; i++)
                {
                    var g = groups[i];
                    int col = (i % 5) * 2;   // 0, 2, 4, 6, 8 (5개 데이터 열)
                    int r   = i / 5;
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
                            Text = g.Short, FontSize = AppTheme.FontXS,
                            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                            Foreground = new SolidColorBrush(Color.Parse(catFg)),
                        },
                    });
                    row.Children.Add(new TextBlock
                    {
                        Text = $"{g.Count}일",
                        FontSize = AppTheme.FontXS,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
            ("일반업무", "📝", "일반업무", "#2a2a1a", "#ccbb88", "#6a5a3a"),
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
                FontSize    = AppTheme.FontBase,
                FontFamily  = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Padding     = new Thickness(10, 4),
                CornerRadius = new CornerRadius(8, 8, 0, 0),
                BorderThickness = new Thickness(1, 1, 1, 0),
                Background  = active ? new SolidColorBrush(Color.Parse(t.Bg)) : AppTheme.BgPrimary,
                Foreground  = active ? new SolidColorBrush(Color.Parse(t.Fg)) : AppTheme.BorderMuted,
                BorderBrush = active ? new SolidColorBrush(Color.Parse(t.Bd)) : AppTheme.BorderSubtle,
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
            case "일반업무":
                contentPanel.Children.Add(BuildShow3GeneralTaskTab(agent));
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
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgInfo,
            Margin = new Thickness(0, 0, 0, 2),
        });

        var dateRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        var txbRange = new TextBlock
        {
            Text = $"{DateTime.Today:yyyy-MM-dd}",
            FontSize = AppTheme.FontBase,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
            MinWidth = 170,
        };
        var btnToday = new Button
        {
            Content = "오늘", Width = 48, Height = 24, FontSize = AppTheme.FontSM,
            Padding = new Thickness(4, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background = AppTheme.BgActiveBlue,
            Foreground = AppTheme.FgInfo,
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderAccent,
            CornerRadius = new CornerRadius(4),
        };
        var btnMonth = new Button
        {
            Content = "이번달", Width = 52, Height = 24, FontSize = AppTheme.FontSM,
            Padding = new Thickness(4, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background = AppTheme.BgActiveGreen,
            Foreground = AppTheme.FgSuccess,
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderActive,
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

        // DragOver 핸들러 — 드롭 허용 표시
        scroll.AddHandler(DragDrop.DragOverEvent, (object? sender, DragEventArgs e) =>
        {
            e.DragEffects = e.Data.Contains("analyte") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        });

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
            Show4RefreshRequest?.Invoke();
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
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccaa")),
            Margin = new Thickness(0, 0, 0, 4),
        });

        if (agent.담당업체목록.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "배정된 계약업체 없음", FontSize = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
            });
            return root;
        }

        foreach (var name in agent.담당업체목록)
        {
            var (bg, fg) = BadgeColorHelper.GetBadgeColor(name);
            var row = new Border
            {
                Background   = AppTheme.BgActiveGreen,
                BorderBrush  = AppTheme.BorderActive,
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
                    Text = name, FontSize = AppTheme.FontBase,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = new SolidColorBrush(Color.Parse(fg)),
                },
            });
            row.Child = sp;
            root.Children.Add(row);
        }
        return root;
    }

    // ── 일반업무 탭 — 전체 업무 목록 + 직원 배정 인라인 ──────────────────────
    private Control BuildShow3GeneralTaskTab(Agent agent)
    {
        var kbR = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var kbM = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        var root = new StackPanel { Spacing = 4 };

        root.Children.Add(new TextBlock
        {
            Text = $"📝 일반업무 배정 — {agent.성명}",
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = kbM, Foreground = AppTheme.FgWarn,
            Margin = new Thickness(0, 0, 0, 6),
        });

        // 설명 텍스트
        root.Children.Add(new TextBlock
        {
            Text = "행을 클릭하면 이 직원의 배정을 추가/해제합니다.",
            FontSize = AppTheme.FontXS, FontFamily = kbR,
            Foreground = new SolidColorBrush(Color.Parse("#778")),
            Margin = new Thickness(0, 0, 0, 6),
        });

        // 전체 업무 × 이 직원 배정 상태
        var allTasks   = GeneralTaskService.GetAll();
        var grouped    = allTasks.GroupBy(t => t.업무명).OrderBy(g => g.Key).ToList();
        var myTaskIds  = new HashSet<string>(
            GeneralTaskService.GetByAgent(agent.사번).Select(t => t.업무명),
            StringComparer.OrdinalIgnoreCase);

        if (grouped.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "등록된 업무 없음 — 하단 '＋ 신규 업무' 버튼으로 추가하세요.",
                FontSize = AppTheme.FontSM, FontFamily = kbR,
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(0, 8),
            });
        }
        else
        {
            var listPanel = new StackPanel { Spacing = 2 };

            void RebuildList()
            {
                listPanel.Children.Clear();
                var updatedMyIds = new HashSet<string>(
                    GeneralTaskService.GetByAgent(agent.사번).Select(t => t.업무명),
                    StringComparer.OrdinalIgnoreCase);
                var updatedAll = GeneralTaskService.GetAll()
                    .GroupBy(t => t.업무명).OrderBy(g => g.Key).ToList();

                foreach (var g in updatedAll)
                {
                    bool assigned = updatedMyIds.Contains(g.Key);
                    var tasksInGroup = g.ToList();
                    var capturedKey  = g.Key;

                    var row = new Border
                    {
                        Background      = assigned
                            ? new SolidColorBrush(Color.Parse("#1a3a22"))
                            : new SolidColorBrush(Color.Parse("#222226")),
                        BorderBrush     = assigned
                            ? new SolidColorBrush(Color.Parse("#3a7a4a"))
                            : new SolidColorBrush(Color.Parse("#383840")),
                        BorderThickness = new Thickness(1),
                        CornerRadius    = new CornerRadius(5),
                        Padding         = new Thickness(8, 6),
                        Cursor          = new Cursor(StandardCursorType.Hand),
                    };

                    var inner = new Grid
                    {
                        ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto,Auto"),
                    };

                    // 토글 아이콘
                    inner.Children.Add(new TextBlock
                    {
                        Text = assigned ? "✅" : "⬜",
                        FontSize = AppTheme.FontXL,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 8, 0),
                    });

                    // 업무명 + 배정 인원 뱃지
                    var nameCol = new StackPanel { Spacing = 2, [Grid.ColumnProperty] = 1,
                        VerticalAlignment = VerticalAlignment.Center };
                    nameCol.Children.Add(new TextBlock
                    {
                        Text = g.Key, FontSize = AppTheme.FontBase, FontFamily = kbM,
                        Foreground = assigned
                            ? new SolidColorBrush(Color.Parse("#88eebb"))
                            : AppRes("AppFg"),
                    });
                    var chips = new WrapPanel { Orientation = Orientation.Horizontal };
                    foreach (var t in tasksInGroup.Where(t => !string.IsNullOrWhiteSpace(t.담당자명)))
                    {
                        var (cbg, cfg) = BadgeColorHelper.GetBadgeColor(t.담당자명);
                        chips.Children.Add(new Border
                        {
                            Background = new SolidColorBrush(Color.Parse(cbg)),
                            CornerRadius = new CornerRadius(6),
                            Padding = new Thickness(5, 1),
                            Margin = new Thickness(0, 0, 3, 0),
                            Child = new TextBlock
                            {
                                Text = t.담당자명, FontSize = AppTheme.FontXS, FontFamily = kbR,
                                Foreground = new SolidColorBrush(Color.Parse(cfg)),
                            },
                        });
                    }
                    if (chips.Children.Count > 0) nameCol.Children.Add(chips);
                    inner.Children.Add(nameCol);

                    // 수정 버튼 (업무 자체 내용 편집)
                    var myRecord = tasksInGroup.FirstOrDefault(t => t.담당자id == agent.사번);
                    if (assigned && myRecord != null)
                    {
                        var capturedRecord = myRecord;
                        var btnEdit = new Button
                        {
                            Content = "✏️", Width = 26, Height = 26, Padding = new Thickness(0),
                            Background = AppTheme.BorderSubtle,
                            BorderThickness = new Thickness(1),
                            BorderBrush = AppTheme.BorderDefault,
                            CornerRadius = new CornerRadius(4),
                            Cursor = new Cursor(StandardCursorType.Hand),
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(4, 0, 0, 0),
                            [Grid.ColumnProperty] = 2,
                        };
                        btnEdit.Click += (_, e) =>
                        {
                            e.Handled = true;
                            Show3ContentRequest?.Invoke(BuildGeneralTaskForm(agent, capturedRecord));
                        };
                        inner.Children.Add(btnEdit);
                    }

                    row.Child = inner;
                    row.PointerPressed += (_, _) =>
                    {
                        if (assigned)
                        {
                            var toRemove = tasksInGroup.FirstOrDefault(t => t.담당자id == agent.사번);
                            if (toRemove != null) GeneralTaskService.Delete(toRemove.Id);
                        }
                        else
                        {
                            var template = tasksInGroup.First();
                            GeneralTaskService.Insert(new GeneralTask
                            {
                                업무명   = capturedKey,
                                내용     = template.내용,
                                배정자   = ETA.Views.MainPage.CurrentEmployeeId,
                                담당자id = agent.사번,
                                담당자명 = agent.성명,
                                마감일   = template.마감일,
                                등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                            });
                        }
                        RefreshGeneralTasks(agent);
                        RebuildList();
                    };

                    listPanel.Children.Add(row);
                }
            }

            RebuildList();
            root.Children.Add(listPanel);
        }

        // 구분선 + 신규 업무 추가
        root.Children.Add(new Border
        {
            Height = 1, Background = AppTheme.BorderSubtle,
            Margin = new Thickness(0, 8, 0, 4),
        });
        var btnAdd = new Button
        {
            Content = "＋ 신규 업무 추가", FontSize = AppTheme.FontBase, Height = 32,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            FontFamily = kbM,
            Background = new SolidColorBrush(Color.Parse("#2a2a1a")),
            Foreground = AppTheme.FgWarn,
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderWarn,
            CornerRadius = new CornerRadius(4),
            Cursor = new Cursor(StandardCursorType.Hand),
        };
        btnAdd.Click += (_, _) => Show3ContentRequest?.Invoke(BuildGeneralTaskForm(agent));
        root.Children.Add(btnAdd);

        return root;
    }

    /// <summary>일반업무 등록/수정 폼 (Show3에 표시)</summary>
    private Control BuildGeneralTaskForm(Agent agent, GeneralTask? existing = null)
    {
        bool isEdit = existing != null;
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(4) };

        root.Children.Add(new TextBlock
        {
            Text = isEdit ? $"✏️ 일반업무 수정 — {agent.성명}" : $"➕ 신규 일반업무 배정 — {agent.성명}",
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgWarn,
        });

        var txtName = new TextBox
        {
            Text = existing?.업무명 ?? "",
            Watermark = "업무명",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
        };
        root.Children.Add(txtContent);

        // 마감일
        var deadlineRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        deadlineRow.Children.Add(new TextBlock
        {
            Text = "마감일", FontSize = AppTheme.FontBase,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var txtDeadline = new TextBox
        {
            Text = existing?.마감일 ?? "",
            Watermark = "yyyy-MM-dd",
            Width = 140,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 4),
        };
        deadlineRow.Children.Add(txtDeadline);
        root.Children.Add(deadlineRow);

        // 상태 토글 (수정시)
        string currentStatus = existing?.상태 ?? "대기";
        if (isEdit)
        {
            var statusRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            statusRow.Children.Add(new TextBlock
            {
                Text = "상태", FontSize = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
                    Content = st, FontSize = AppTheme.FontSM,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Padding = new Thickness(8, 3),
                    CornerRadius = new CornerRadius(8),
                    BorderThickness = new Thickness(1),
                    Background  = active ? new SolidColorBrush(Color.Parse(bg)) : AppTheme.BgSecondary,
                    Foreground  = active ? new SolidColorBrush(Color.Parse(fg)) : AppTheme.BorderMuted,
                    BorderBrush = active ? new SolidColorBrush(Color.Parse(bd)) : AppTheme.BorderMuted,
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
                        bool isActive = s == capturedSt;
                        child.Background  = isActive ? new SolidColorBrush(Color.Parse(b2)) : AppTheme.BgSecondary;
                        child.Foreground  = isActive ? new SolidColorBrush(Color.Parse(f2)) : AppTheme.BorderMuted;
                        child.BorderBrush = isActive ? new SolidColorBrush(Color.Parse(d2)) : AppTheme.BorderMuted;
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase,
            Background = AppTheme.BgActiveGreen,
            Foreground = AppTheme.FgSuccess,
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
                GeneralTaskService.Update(existing);
            }
            else
            {
                var t = new GeneralTask
                {
                    업무명   = txtName.Text.Trim(),
                    내용     = txtContent.Text?.Trim() ?? "",
                    배정자   = ETA.Views.MainPage.CurrentEmployeeId,
                    담당자id = agent.사번,
                    담당자명 = agent.성명,
                    마감일   = txtDeadline.Text?.Trim() ?? "",
                    등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                };
                GeneralTaskService.Insert(t);
            }
            _show3Tab = "일반업무";
            Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
        };
        btnRow.Children.Add(btnSave);

        if (isEdit)
        {
            var btnDel = new Button
            {
                Content = "🗑 해당 업무 삭제", Height = 30, Padding = new Thickness(10, 0),
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                FontSize = AppTheme.FontBase,
                Background = AppTheme.BgDanger,
                Foreground = AppTheme.FgDanger,
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            };
            btnDel.Click += (_, _) =>
            {
                GeneralTaskService.Delete(existing!.Id);
                _show3Tab = "일반업무";
                Show3ContentRequest?.Invoke(BuildShow3Timeline(agent));
            };
            btnRow.Children.Add(btnDel);
        }

        var btnBack = new Button
        {
            Content = "← 돌아가기", Height = 30, Padding = new Thickness(10, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase,
            Background = AppTheme.BorderSubtle,
            Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => { _show3Tab = "일반업무"; Show3ContentRequest?.Invoke(BuildShow3Timeline(agent)); };
        btnRow.Children.Add(btnBack);

        root.Children.Add(btnRow);
        return root;
    }

    // ── Show4에서 호출: 담당자 선택 가능한 일반업무 추가/수정 폼 ─────────────
    public void ShowGeneralTaskForm(GeneralTask? existing, Action? onSaved)
    {
        bool isEdit = existing != null;
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(4) };

        root.Children.Add(new TextBlock
        {
            Text = isEdit ? "✏️ 일반업무 수정" : "➕ 신규 일반업무 추가",
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgWarn,
        });

        var txtName = new TextBox
        {
            Text = existing?.업무명 ?? "",
            Watermark = "업무명",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
        };
        root.Children.Add(txtContent);

        // 담당자 선택 콤보박스
        var agentRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        agentRow.Children.Add(new TextBlock
        {
            Text = "담당자", FontSize = AppTheme.FontBase,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var agents = AgentService.GetAllItems().OrderBy(a => a.성명).ToList();
        var cboAgent = new ComboBox
        {
            Width = 180, FontSize = AppTheme.FontMD,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
        };
        foreach (var a in agents)
            cboAgent.Items.Add(new ComboBoxItem { Content = a.성명, Tag = a.사번 });
        // 기존 담당자 선택
        if (isEdit && !string.IsNullOrEmpty(existing!.담당자id))
        {
            for (int i = 0; i < cboAgent.Items.Count; i++)
                if (cboAgent.Items[i] is ComboBoxItem ci && ci.Tag?.ToString() == existing.담당자id)
                { cboAgent.SelectedIndex = i; break; }
        }
        agentRow.Children.Add(cboAgent);
        root.Children.Add(agentRow);

        // 마감일
        var deadlineRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        deadlineRow.Children.Add(new TextBlock
        {
            Text = "마감일", FontSize = AppTheme.FontBase,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var txtDeadline = new TextBox
        {
            Text = existing?.마감일 ?? "",
            Watermark = "yyyy-MM-dd",
            Width = 140,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 4),
        };
        deadlineRow.Children.Add(txtDeadline);
        root.Children.Add(deadlineRow);

        // 상태 토글 (수정시)
        string currentStatus = existing?.상태 ?? "대기";
        if (isEdit)
        {
            var statusRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            statusRow.Children.Add(new TextBlock
            {
                Text = "상태", FontSize = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
                    Content = st, FontSize = AppTheme.FontSM,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Padding = new Thickness(8, 3),
                    CornerRadius = new CornerRadius(8),
                    BorderThickness = new Thickness(1),
                    Background  = active ? new SolidColorBrush(Color.Parse(bg)) : AppTheme.BgSecondary,
                    Foreground  = active ? new SolidColorBrush(Color.Parse(fg)) : AppTheme.BorderMuted,
                    BorderBrush = active ? new SolidColorBrush(Color.Parse(bd)) : AppTheme.BorderMuted,
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
                        bool isActive = s == capturedSt;
                        child.Background  = isActive ? new SolidColorBrush(Color.Parse(b2)) : AppTheme.BgSecondary;
                        child.Foreground  = isActive ? new SolidColorBrush(Color.Parse(f2)) : AppTheme.BorderMuted;
                        child.BorderBrush = isActive ? new SolidColorBrush(Color.Parse(d2)) : AppTheme.BorderMuted;
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase,
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
                if (!string.IsNullOrEmpty(agentId))
                {
                    existing!.담당자id = agentId;
                    existing!.담당자명 = agentName;
                }
                if (currentStatus == "완료" && string.IsNullOrEmpty(existing.완료일시))
                    existing.완료일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                GeneralTaskService.Update(existing);
            }
            else
            {
                var t = new GeneralTask
                {
                    업무명   = txtName.Text.Trim(),
                    내용     = txtContent.Text?.Trim() ?? "",
                    배정자   = ETA.Views.MainPage.CurrentEmployeeId,
                    담당자id = agentId,
                    담당자명 = agentName,
                    마감일   = txtDeadline.Text?.Trim() ?? "",
                    등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                };
                GeneralTaskService.Insert(t);
            }
            onSaved?.Invoke();
        };
        btnRow.Children.Add(btnSave);

        if (isEdit)
        {
            var btnDel = new Button
            {
                Content = "🗑 해당 업무 삭제", Height = 30, Padding = new Thickness(10, 0),
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                FontSize = AppTheme.FontBase,
                Background = AppTheme.BgDanger,
                Foreground = AppTheme.FgDanger,
                BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            };
            btnDel.Click += (_, _) =>
            {
                GeneralTaskService.Delete(existing!.Id);
                onSaved?.Invoke();
            };
            btnRow.Children.Add(btnDel);
        }

        var btnBack = new Button
        {
            Content = "← 취소", Height = 30, Padding = new Thickness(10, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase,
            Background = AppTheme.BorderSubtle,
            Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => onSaved?.Invoke();
        btnRow.Children.Add(btnBack);

        root.Children.Add(btnRow);
        Show4ContentRequest?.Invoke(new ScrollViewer
        {
            Content = root,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        });
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
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = new SolidColorBrush(Color.Parse("#ccaaff")),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var btnAdd = new Button
        {
            Content = "➕ 신규 추가", FontSize = AppTheme.FontBase, Height = 28,
            Padding = new Thickness(10, 0),
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background = new SolidColorBrush(Color.Parse("#2a1a3a")),
            Foreground = new SolidColorBrush(Color.Parse("#cc88ff")),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderAccent,
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
                Text = "배정된 업무 없음", FontSize = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
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
                    Text = t.상태, FontSize = AppTheme.FontSM,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = new SolidColorBrush(Color.Parse(statusStyle.Fg)),
                },
            });

            // 정보
            var info = new StackPanel { Spacing = 2, [Grid.ColumnProperty] = 1 };
            info.Children.Add(new TextBlock
            {
                Text = t.업무명, FontSize = AppTheme.FontMD,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = new SolidColorBrush(Color.Parse(statusStyle.Fg)),
            });
            var meta = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            if (!string.IsNullOrEmpty(t.마감일))
                meta.Children.Add(new TextBlock
                {
                    Text = $"마감: {t.마감일}", FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppTheme.FgMuted,
                });
            if (!string.IsNullOrEmpty(t.내용))
                meta.Children.Add(new TextBlock
                {
                    Text = t.내용, FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppTheme.FgDimmed,
                    TextTrimming = TextTrimming.CharacterEllipsis, MaxWidth = 200,
                });
            if (meta.Children.Count > 0) info.Children.Add(meta);
            grid.Children.Add(info);

            // 수정 버튼
            var capturedTask = t;
            var btnEdit = new Button
            {
                Content = "✏️", Width = 28, Height = 28, Padding = new Thickness(0),
                Background = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderDefault,
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
                Background = AppTheme.BgDanger,
                BorderThickness = new Thickness(1),
                BorderBrush = AppTheme.BorderDanger,
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
                Text = "배정된 업무 없음", FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
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
                    Text = t.상태, FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = new SolidColorBrush(Color.Parse(statusColor.Item2)),
                },
            });

            // 업무명 + 마감일
            var info = new StackPanel { [Grid.ColumnProperty] = 1 };
            info.Children.Add(new TextBlock
            {
                Text = t.업무명, FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = new SolidColorBrush(Color.Parse(statusColor.Item2)),
                TextTrimming = TextTrimming.CharacterEllipsis,
            });
            if (!string.IsNullOrEmpty(t.마감일))
                info.Children.Add(new TextBlock
                {
                    Text = $"마감: {t.마감일}", FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppTheme.FgDimmed,
                });
            row.Children.Add(info);

            // 수정 버튼
            var capturedTask = t;
            var capturedAgent = agent;
            var btnEdit = new Button
            {
                Content = "✏️", Width = 24, Height = 24, Padding = new Thickness(0),
                Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                Foreground = AppTheme.FgInfo,
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
            FontSize = AppTheme.FontLG, FontWeight = FontWeight.Bold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = new SolidColorBrush(Color.Parse("#ccaaff")),
        });

        var txtName = new TextBox
        {
            Text = existing?.업무명 ?? "",
            Watermark = "업무명",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
        };
        root.Children.Add(txtContent);

        // 마감일
        var deadlineRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        deadlineRow.Children.Add(new TextBlock
        {
            Text = "마감일", FontSize = AppTheme.FontBase,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        var txtDeadline = new TextBox
        {
            Text = existing?.마감일 ?? "",
            Watermark = "yyyy-MM-dd",
            Width = 140,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontMD,
            Background = AppTheme.BorderSeparator,
            Foreground = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush = AppTheme.BorderDefault,
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 4),
        };
        deadlineRow.Children.Add(txtDeadline);
        root.Children.Add(deadlineRow);

        // 상태 토글 (수정시)
        string currentStatus = existing?.상태 ?? "대기";
        if (isEdit)
        {
            var statusRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            statusRow.Children.Add(new TextBlock
            {
                Text = "상태", FontSize = AppTheme.FontBase,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
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
                    Content = st, FontSize = AppTheme.FontSM,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Padding = new Thickness(8, 3),
                    CornerRadius = new CornerRadius(8),
                    BorderThickness = new Thickness(1),
                    Background  = active ? new SolidColorBrush(Color.Parse(bg)) : AppTheme.BgSecondary,
                    Foreground  = active ? new SolidColorBrush(Color.Parse(fg)) : AppTheme.BorderMuted,
                    BorderBrush = active ? new SolidColorBrush(Color.Parse(bd)) : AppTheme.BorderMuted,
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
                        child.Background  = isActive ? new SolidColorBrush(Color.Parse(b2)) : AppTheme.BgSecondary;
                        child.Foreground  = isActive ? new SolidColorBrush(Color.Parse(f2)) : AppTheme.BorderMuted;
                        child.BorderBrush = isActive ? new SolidColorBrush(Color.Parse(d2)) : AppTheme.BorderMuted;
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase,
            Background = AppTheme.BgActiveGreen,
            Foreground = AppTheme.FgSuccess,
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
                Content = "🗑 해당 업무 삭제", Height = 30, Padding = new Thickness(10, 0),
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                FontSize = AppTheme.FontBase,
                Background = AppTheme.BgDanger,
                Foreground = AppTheme.FgDanger,
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
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize = AppTheme.FontBase,
            Background = AppTheme.BorderSubtle,
            Foreground = AppTheme.FgMuted,
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
            Content = "← 돌아가기", FontSize = AppTheme.FontSM,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Padding = new Thickness(8, 4),
            Background = AppTheme.BorderSubtle,
            Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => Show4ContentRequest?.Invoke(null);
        hdr.Children.Add(btnBack);
        hdr.Children.Add(new TextBlock
        {
            Text = "분석항목 선택", FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0),
            [Grid.ColumnProperty] = 1,
        });
        root.Children.Add(hdr);
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderMuted });

        // 분석정보 기준 항목 목록 (약칭 포함)
        var analyteInfos = AnalysisService.GetAllItems()
            .Where(a => !string.IsNullOrWhiteSpace(a.Analyte))
            .ToList();
        foreach (var ai in analyteInfos)
        {
            string fullName  = ai.Analyte;
            string shortName = string.IsNullOrWhiteSpace(ai.약칭) ? ai.Analyte : ai.약칭;
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
                    Text = shortName, FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                },
            });

            // 전체 항목명
            inner.Children.Add(new TextBlock
            {
                Text = fullName, FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = assigned ? new SolidColorBrush(Color.Parse(badgeFg)) : AppRes("AppFg"),
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 1,
            });

            if (assigned)
                inner.Children.Add(new TextBlock
                {
                    Text = "✓", FontSize = AppTheme.FontBase,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppTheme.FgSuccess,
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
            Content = "← 돌아가기", FontSize = AppTheme.FontSM,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Padding = new Thickness(8, 4),
            Background = AppTheme.BorderSubtle,
            Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => Show4ContentRequest?.Invoke(null);
        hdr.Children.Add(btnBack);
        hdr.Children.Add(new TextBlock
        {
            Text = "계약업체 선택 (현행 계약)", FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccaa")),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0),
            [Grid.ColumnProperty] = 1,
        });
        root.Children.Add(hdr);
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderMuted });

        var contracts = ContractService.GetAllContracts()
            .Where(c => c.DaysLeft == null || c.DaysLeft >= 0)
            .OrderBy(c => c.C_CompanyName).ToList();

        if (contracts.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "현행 계약 없음", FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
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
                    Background   = assigned ? AppTheme.BgActiveGreen : Brushes.Transparent,
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
                        Text = abbr, FontSize = AppTheme.FontXS,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                        Foreground = new SolidColorBrush(Color.Parse(fg)),
                    },
                });
                inner.Children.Add(new TextBlock
                {
                    Text = c.C_CompanyName, FontSize = AppTheme.FontSM,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppRes("AppFg"),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    [Grid.ColumnProperty] = 1,
                });
                if (assigned)
                    inner.Children.Add(new TextBlock
                    {
                        Text = "✓", FontSize = AppTheme.FontBase,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                        Foreground = AppTheme.FgSuccess,
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
    //  Show4: 일반업무 배정 리스트 (Show2 "+" 클릭 시)
    // =========================================================================
    private Control BuildShow4GeneralTaskAssignList(Agent agent)
    {
        var root = new StackPanel { Spacing = 3, Margin = new Thickness(4) };

        // 헤더
        var hdr = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*"), Margin = new Thickness(0, 0, 0, 6) };
        var btnBack = new Button
        {
            Content = "← 돌아가기", FontSize = AppTheme.FontSM,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Padding = new Thickness(8, 4),
            Background = AppTheme.BorderSubtle,
            Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnBack.Click += (_, _) => Show4ContentRequest?.Invoke(null);
        hdr.Children.Add(btnBack);
        hdr.Children.Add(new TextBlock
        {
            Text = $"일반업무 배정 — {agent.성명}", FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Foreground = new SolidColorBrush(Color.Parse("#aaccbb")),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0),
            [Grid.ColumnProperty] = 1,
        });
        root.Children.Add(hdr);
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderMuted });

        // 전체 업무 가져오기 — 업무명 기준 그룹
        var allTasks = GeneralTaskService.GetAll();
        var grouped = allTasks.GroupBy(t => t.업무명).OrderBy(g => g.Key).ToList();

        if (grouped.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "등록된 일반업무 없음\n하단 '＋ 신규 업무' 버튼으로 먼저 업무를 추가하세요.",
                FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(0, 8),
            });
        }
        else
        {
            foreach (var g in grouped)
            {
                var tasksInGroup = g.ToList();
                bool assigned = tasksInGroup.Any(t => t.담당자id == agent.사번);

                var row = new Border
                {
                    Background   = assigned ? AppTheme.BgActiveGreen : Brushes.Transparent,
                    CornerRadius = new CornerRadius(4),
                    Padding      = new Thickness(6, 5),
                    Margin       = new Thickness(0, 1),
                    Cursor       = new Cursor(StandardCursorType.Hand),
                };

                var inner = new StackPanel { Spacing = 2 };

                // 업무명 + 체크마크
                var nameRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
                nameRow.Children.Add(new TextBlock
                {
                    Text = g.Key, FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = assigned
                        ? new SolidColorBrush(Color.Parse("#aaeebb"))
                        : AppRes("AppFg"),
                    TextTrimming = TextTrimming.CharacterEllipsis,
                });
                if (assigned)
                    nameRow.Children.Add(new TextBlock
                    {
                        Text = "✓", FontSize = AppTheme.FontBase,
                        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                        Foreground = AppTheme.FgSuccess,
                        VerticalAlignment = VerticalAlignment.Center,
                        [Grid.ColumnProperty] = 1,
                    });
                inner.Children.Add(nameRow);

                // 배정된 인원 뱃지
                var chipsPanel = new WrapPanel { Orientation = Orientation.Horizontal };
                foreach (var t in tasksInGroup)
                {
                    if (string.IsNullOrWhiteSpace(t.담당자명)) continue;
                    var (bg, fg) = BadgeColorHelper.GetBadgeColor(t.담당자명);
                    chipsPanel.Children.Add(new Border
                    {
                        Background   = new SolidColorBrush(Color.Parse(bg)),
                        CornerRadius = new CornerRadius(6),
                        Padding      = new Thickness(5, 1),
                        Margin       = new Thickness(0, 0, 3, 0),
                        Child = new TextBlock
                        {
                            Text = t.담당자명, FontSize = AppTheme.FontXS,
                            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                            Foreground = new SolidColorBrush(Color.Parse(fg)),
                        },
                    });
                }
                if (chipsPanel.Children.Count > 0)
                    inner.Children.Add(chipsPanel);

                row.Child = inner;

                var capturedTaskName = g.Key;
                row.PointerPressed += (_, _) =>
                {
                    if (assigned)
                    {
                        // 이미 배정 → 제거
                        var toRemove = tasksInGroup.FirstOrDefault(t => t.담당자id == agent.사번);
                        if (toRemove != null) GeneralTaskService.Delete(toRemove.Id);
                    }
                    else
                    {
                        // 미배정 → 추가 (기존 업무 기반 복제)
                        var template = tasksInGroup.First();
                        var newTask = new GeneralTask
                        {
                            업무명   = capturedTaskName,
                            내용     = template.내용,
                            배정자   = ETA.Views.MainPage.CurrentEmployeeId,
                            담당자id = agent.사번,
                            담당자명 = agent.성명,
                            마감일   = template.마감일,
                            등록일시 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        };
                        GeneralTaskService.Insert(newTask);
                    }
                    RefreshGeneralTasks(agent);
                    Show4ContentRequest?.Invoke(BuildShow4GeneralTaskAssignList(agent));
                };
                root.Children.Add(row);
            }
        }

        // 하단: 신규 업무 추가 버튼
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderSubtle, Margin = new Thickness(0, 6) });
        var btnNewTask = new Button
        {
            Content = "＋ 신규 업무 추가", FontSize = AppTheme.FontBase, Height = 30,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            Background = new SolidColorBrush(Color.Parse("#2a2a1a")),
            Foreground = AppTheme.FgWarn,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
        };
        btnNewTask.Click += (_, _) =>
        {
            ShowGeneralTaskForm(null, () =>
            {
                RefreshGeneralTasks(agent);
                Show4ContentRequest?.Invoke(BuildShow4GeneralTaskAssignList(agent));
            });
        };
        root.Children.Add(btnNewTask);

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
        // 이번달 기준: 전월 20일 ~ +2개월 10일 (4월→3/20~6/10, 약 80일)
        _chartRangeStart = new DateTime(today.Year, today.Month, 1).AddDays(-11); // 전월 20일
        _chartRangeEnd   = new DateTime(today.Year, today.Month, 1).AddMonths(2).AddDays(9); // +2개월 10일
        bool isH1 = today.Month <= 6;

        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };
        var kbFont = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        // ── Row 0: 헤더 (wire-v01 pill 스타일) ───────────────────────────────
        var header = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(12, 10, 12, 6) };

        header.Children.Add(new TextBlock {
            Text = "업무분장표", FontSize = AppTheme.FontXL + 4, FontWeight = FontWeight.Bold,
            FontFamily = kbFont, Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 12, 0) });

        // wire-v01: pill 버튼 팩토리 (디자인 시스템 상태 토큰 사용)
        Button MakePill(string text, BadgeStatus status, bool active = true, int w = 52)
        {
            var (bg, fg, bd) = active
                ? StatusBadge.GetBrushes(status)
                : (AppTheme.StatusMutedBg, AppTheme.StatusMutedFg, AppTheme.StatusMutedBorder);
            return new Button
            {
                Content = text, Width = w, Height = 26, FontSize = AppTheme.FontSM,
                Padding = new Thickness(0), CornerRadius = new CornerRadius(999),
                FontFamily = kbFont, Cursor = new Cursor(StandardCursorType.Hand),
                Background = bg, Foreground = fg,
                BorderThickness = new Thickness(1.2), BorderBrush = bd,
                FontWeight = active ? FontWeight.SemiBold : FontWeight.Normal,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
            };
        }

        var btnPrev  = MakePill("◀",     BadgeStatus.Info,   active: true,  w: 32);
        var btnNext  = MakePill("▶",     BadgeStatus.Info,   active: true,  w: 32);
        var btnToday = MakePill("이번달", BadgeStatus.Accent, active: true,  w: 60);
        var btnH1    = MakePill("상반기", BadgeStatus.Ok,     active: isH1);
        var btnH2    = MakePill("하반기", BadgeStatus.Ok,     active: !isH1);
        var btnApply = MakePill("반영",   BadgeStatus.Accent, active: true,  w: 54);

        var txbMonth = new TextBlock {
            Text = $"{_chartRangeStart:M/dd} – {_chartRangeEnd:M/dd}", FontSize = AppTheme.FontMD,
            FontFamily = kbFont, Foreground = AppTheme.FgSecondary, FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 8, 0) };

        header.Children.Add(btnPrev);
        header.Children.Add(txbMonth);
        header.Children.Add(btnNext);
        header.Children.Add(btnToday);
        header.Children.Add(btnH1);
        header.Children.Add(btnH2);
        header.Children.Add(btnApply);
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // ── Row 1: 카테고리 탭 (wire-v01 active=Info, inactive=Muted) ─────────
        var tabBar = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, Margin = new Thickness(12, 0, 12, 6) };
        string[] tabNames = ["분석항목", "계약업체", "일반업무"];
        var tabBtns = new Button[tabNames.Length];
        for (int t = 0; t < tabNames.Length; t++)
        {
            bool active = tabNames[t] == _chartTab;
            var tb = MakePill(tabNames[t], BadgeStatus.Info, active: active, w: 76);
            tabBtns[t] = tb;
            tabBar.Children.Add(tb);
        }
        Grid.SetRow(tabBar, 2);
        root.Children.Add(tabBar);

        // ── Row 1: Summary strip (wire-v01) ────────────────────────────────
        var summaryStrip = new Border
        {
            Background = AppTheme.BgCard, BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 1, 0, 1),
            Padding = new Thickness(12, 8), Margin = new Thickness(0, 0, 0, 0),
        };
        var summaryPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        summaryStrip.Child = summaryPanel;
        _chartSummaryPanel = summaryPanel;
        Grid.SetRow(summaryStrip, 1);
        root.Children.Add(summaryStrip);

        // ── Row 3: 헤더+본문 공통 래퍼 (같은 부모 → 컬럼 정렬 보장) ──────────
        var chartWrapper = new Grid
        {
            Margin          = new Thickness(4, 0, 4, 4),
            RowDefinitions  = new RowDefinitions("Auto,*"),
        };

        var fixedHeader = new StackPanel { Spacing = 0 };
        Grid.SetRow(fixedHeader, 0);
        chartWrapper.Children.Add(fixedHeader);

        var body = new StackPanel { Spacing = 0 };
        var scroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Padding         = new Thickness(0),
            BorderThickness = new Thickness(0),
            Content         = body,
        };
        Grid.SetRow(scroll, 1);
        chartWrapper.Children.Add(scroll);

        Grid.SetRow(chartWrapper, 3);
        root.Children.Add(chartWrapper);

        // 탭별 렌더링
        void RefreshBody()
        {
            switch (_chartTab)
            {
                case "분석항목": RefreshChartTable(fixedHeader, body); break;
                case "계약업체": fixedHeader.Children.Clear(); RefreshChartContractTab(body); break;
                case "일반업무": fixedHeader.Children.Clear(); RefreshChartGeneralTab(body); break;
            }
        }
        RefreshBody();

        // ── 이벤트 ─────────────────────────────────────────────────────────
        void ApplyPillStyle(Button btn, BadgeStatus status, bool active, FontWeight? weight = null)
        {
            var (bg, fg, bd) = active
                ? StatusBadge.GetBrushes(status)
                : (AppTheme.StatusMutedBg, AppTheme.StatusMutedFg, AppTheme.StatusMutedBorder);
            btn.Background = bg; btn.Foreground = fg; btn.BorderBrush = bd;
            btn.FontWeight = weight ?? (active ? FontWeight.SemiBold : FontWeight.Normal);
        }

        void UpdateHalfBtnStyle()
        {
            bool h1 = _chartRangeStart.Month <= 6;
            ApplyPillStyle(btnH1, BadgeStatus.Ok, active: h1);
            ApplyPillStyle(btnH2, BadgeStatus.Ok, active: !h1);
        }

        void Navigate()
        {
            txbMonth.Text = $"{_chartRangeStart:M/dd} – {_chartRangeEnd:M/dd}";
            _chartCachedSpans = null; // 범위 변경 시 DB 재로드
            UpdateHalfBtnStyle();
            RefreshBody();
            btnApply.IsEnabled = true; btnApply.Content = "반영";
        }

        // ◀ ▶: 1개월씩 이동
        btnPrev.Click += (_, _) =>
        { _chartRangeStart = _chartRangeStart.AddMonths(-1); _chartRangeEnd = _chartRangeEnd.AddMonths(-1); Navigate(); };
        btnNext.Click += (_, _) =>
        { _chartRangeStart = _chartRangeStart.AddMonths(1); _chartRangeEnd = _chartRangeEnd.AddMonths(1); Navigate(); };
        btnToday.Click += (_, _) =>
        {
            var t = DateTime.Today;
            _chartRangeStart = new DateTime(t.Year, t.Month, 1).AddDays(-11);
            _chartRangeEnd = new DateTime(t.Year, t.Month, 1).AddMonths(2).AddDays(9);
            Navigate();
        };
        btnH1.Click += (_, _) =>
        {
            _chartRangeStart = new DateTime(_chartRangeStart.Year, 1, 1);
            _chartRangeEnd   = new DateTime(_chartRangeStart.Year, 6, 30);
            Navigate();
        };
        btnH2.Click += (_, _) =>
        {
            _chartRangeStart = new DateTime(_chartRangeStart.Year, 7, 1);
            _chartRangeEnd   = new DateTime(_chartRangeStart.Year, 12, 31);
            Navigate();
        };
        btnApply.Click += async (_, _) =>
        {
            if (_chartPendingActions.Count == 0) return;
            btnApply.IsEnabled = false;
            var pb = _chartProgressBar;
            var pl = _chartProgressLabel;
            if (pb != null) { pb.IsVisible = true; pb.Value = 0; }
            if (pl != null) { pl.IsVisible = true; pl.Text = "반영 중..."; }

            var actions = _chartPendingActions.ToList();
            _chartPendingActions.Clear();
            _chartPendingChanges.Clear();

            // 단일 커넥션/트랜잭션으로 일괄 반영 — N회 왕복을 1회로
            await Task.Run(() =>
            {
                try
                {
                    using var conn = DbConnectionFactory.CreateConnection();
                    conn.Open();
                    var cols = DbConnectionFactory.GetColumnNames(conn, "분장표준처리");
                    using var tx = conn.BeginTransaction();

                    int step = Math.Max(1, actions.Count / 20); // 진행률 UI 업데이트 5%마다
                    for (int i = 0; i < actions.Count; i++)
                    {
                        actions[i](conn, tx, cols);
                        if (i == actions.Count - 1 || i % step == 0)
                        {
                            var pct = (double)(i + 1) / actions.Count;
                            Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                            {
                                if (pb != null) pb.Value = pct;
                                if (pl != null) pl.Text = $"반영 중... {(int)(pct * 100)}%";
                            });
                        }
                    }
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[AgentTreePage] 반영 일괄 처리 오류: {ex}");
                }
            });

            if (pb != null) pb.IsVisible = false;
            if (pl != null) pl.IsVisible = false;
            _chartCachedSpans = null; // DB에서 다시 로드
            RefreshBody();
            RefreshShow3AfterChartUpdate();
            btnApply.IsEnabled = true; btnApply.Content = "반영됨";
        };
        // 프로그레스바 (헤더에 인라인) — wire-v01 accent
        var progressBar = new ProgressBar
        {
            Minimum = 0, Maximum = 1, Value = 0,
            Height = 6, Width = 150, CornerRadius = new CornerRadius(3),
            Background = new SolidColorBrush(Color.Parse("#333")),
            VerticalAlignment = VerticalAlignment.Center,
            IsVisible = false,
        };
        progressBar.Classes.Add("accent");
        var progressLabel = new TextBlock
        {
            FontSize = AppTheme.FontXS, FontFamily = kbFont,
            Foreground = AppTheme.StatusAccentFg,
            VerticalAlignment = VerticalAlignment.Center,
            IsVisible = false,
        };
        _chartProgressBar = progressBar;
        _chartProgressLabel = progressLabel;
        header.Children.Add(progressBar);
        header.Children.Add(progressLabel);

        // 탭 클릭
        void UpdateTabStyle()
        {
            for (int i = 0; i < tabNames.Length; i++)
                ApplyPillStyle(tabBtns[i], BadgeStatus.Info, active: tabNames[i] == _chartTab);
        }
        for (int t = 0; t < tabNames.Length; t++)
        {
            var tn = tabNames[t];
            tabBtns[t].Click += (_, _) => { _chartTab = tn; UpdateTabStyle(); RefreshBody(); };
        }

        return root;
    }

    // ── 초성별 색상 매핑 ─────────────────────────────────────────────────
    // 초성별 shimmer 기본 색상 (밝은 톤 — 텍스트 + shimmer sweep 용)
    private static readonly Color[] _chosungBaseColors =
    [
        Color.Parse("#70a8d8"), // ㄱ — 스카이
        Color.Parse("#a088c8"), // ㄴ — 라벤더
        Color.Parse("#70b8b0"), // ㄷ — 틸
        Color.Parse("#c8a878"), // ㄹ — 샌드
        Color.Parse("#8888c0"), // ㅁ — 슬레이트
        Color.Parse("#a0a870"), // ㅂ — 올리브
        Color.Parse("#c88888"), // ㅅ — 코랄
        Color.Parse("#78b0a0"), // ㅇ — 세이지
        Color.Parse("#88b070"), // ㅈ — 모스
        Color.Parse("#b080a8"), // ㅊ — 모브
        Color.Parse("#7898b8"), // ㅋ — 스틸
        Color.Parse("#a07898"), // ㅌ — 플럼
        Color.Parse("#b8a080"), // ㅍ — 톤
        Color.Parse("#80b088"), // ㅎ — 민트
    ];
    private static readonly char[] _chosungTable =
        ['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ'];

    private static Color GetChosungColor(string name)
    {
        if (string.IsNullOrEmpty(name)) return Color.Parse("#666");
        char first = name[0];
        if (first >= 0xAC00 && first <= 0xD7A3)
        {
            int idx = (first - 0xAC00) / (21 * 28);
            int colorIdx = _chosungTable[idx] switch
            {
                'ㄱ' or 'ㄲ' => 0, 'ㄴ' => 1, 'ㄷ' or 'ㄸ' => 2, 'ㄹ' => 3,
                'ㅁ' => 4, 'ㅂ' or 'ㅃ' => 5, 'ㅅ' or 'ㅆ' => 6, 'ㅇ' => 7,
                'ㅈ' or 'ㅉ' => 8, 'ㅊ' => 9, 'ㅋ' => 10, 'ㅌ' => 11,
                'ㅍ' => 12, 'ㅎ' => 13, _ => 7,
            };
            return _chosungBaseColors[colorIdx];
        }
        return Color.Parse("#888");
    }

    /// <summary>업무분장표 테이블 렌더링 — 분석항목 탭 (2개월)</summary>
    private void RefreshChartTable(StackPanel fixedHeader, StackPanel body)
    {
        fixedHeader.Children.Clear();
        body.Children.Clear();
        var kbFont = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        DateTime rangeEnd = _chartRangeEnd;
        int totalDays = (int)(rangeEnd - _chartRangeStart).TotalDays + 1;

        // 캐시가 없으면 DB에서 로드, 있으면 캐시 사용 (로컬 편집 반영)
        if (_chartCachedSpans == null)
            _chartCachedSpans = AnalysisRequestService.GetAssignmentChartData(_chartRangeStart, rangeEnd);
        var spans = _chartCachedSpans;
        var analytes = AnalysisRequestService.GetOrderedAnalytes()
            .Where(a => !a.fullName.StartsWith("_") && a.fullName != "기타업무" && a.fullName != "담당계약업체" && a.fullName != "항목명")
            .ToList();
        if (analytes.Count == 0) return;

        RefreshSummaryStrip(spans, analytes.Select(a => a.fullName).ToList(), rangeEnd);

        double trackW = totalDays * GC_DAY_W;

        // ── 월 구분 헤더 ────────────────────────────────────────────────
        var monthHeaderRow = new Grid { ColumnDefinitions = new ColumnDefinitions($"{GC_LABEL_W},*"), Height = 18 };
        var monthHeaderCanvas = new Canvas { Width = trackW, Height = 18, ClipToBounds = true };
        Grid.SetColumn(monthHeaderCanvas, 1);
        // 범위에 포함되는 월 구분 표시
        var monthColors = new[] { "#90b0d0", "#b0a090", "#a0c090", "#c0a0b0" };
        int dayOffset = 0;
        var curMonth = new DateTime(_chartRangeStart.Year, _chartRangeStart.Month, 1);
        int mIdx = 0;
        while (curMonth <= rangeEnd)
        {
            var nextMonth = curMonth.AddMonths(1);
            // 이 월이 범위 내에서 차지하는 일수
            var mStart = curMonth < _chartRangeStart ? _chartRangeStart : curMonth;
            var mEnd = nextMonth.AddDays(-1) > rangeEnd ? rangeEnd : nextMonth.AddDays(-1);
            int mDays = (int)(mEnd - mStart).TotalDays + 1;
            double mw = mDays * GC_DAY_W;

            monthHeaderCanvas.Children.Add(new TextBlock
            {
                Text = $"{curMonth:M월}", FontSize = AppTheme.FontXS, FontWeight = FontWeight.Bold,
                Width = mw, FontFamily = kbFont, TextAlignment = TextAlignment.Center,
                Foreground = new SolidColorBrush(Color.Parse(monthColors[mIdx % monthColors.Length])),
                [Canvas.LeftProperty] = dayOffset * GC_DAY_W, [Canvas.TopProperty] = 1.0,
            });
            if (mIdx > 0)
            {
                monthHeaderCanvas.Children.Add(new Avalonia.Controls.Shapes.Line
                {
                    StartPoint = new Point(dayOffset * GC_DAY_W, 0),
                    EndPoint = new Point(dayOffset * GC_DAY_W, 18),
                    Stroke = new SolidColorBrush(Color.Parse("#333840")), StrokeThickness = 1,
                });
            }
            dayOffset += mDays;
            curMonth = nextMonth;
            mIdx++;
        }
        monthHeaderRow.Children.Add(monthHeaderCanvas);
        fixedHeader.Children.Add(monthHeaderRow);

        // ── 날짜 헤더 ────────────────────────────────────────────────────
        var dateRow = new Grid { ColumnDefinitions = new ColumnDefinitions($"{GC_LABEL_W},*"), Height = GC_DATE_H };
        var dateCanvas = new Canvas { Width = trackW, Height = GC_DATE_H, ClipToBounds = true };
        Grid.SetColumn(dateCanvas, 1);
        for (int d = 0; d < totalDays; d++)
        {
            var dt = _chartRangeStart.AddDays(d);
            bool isSun = dt.DayOfWeek == DayOfWeek.Sunday;
            bool isSat = dt.DayOfWeek == DayOfWeek.Saturday;
            bool isToday = dt.Date == DateTime.Today;
            string color = isToday ? "#70d070" : isSun ? "#c07070" : isSat ? "#7090c0" : "#808890";
            dateCanvas.Children.Add(new TextBlock
            {
                Text = dt.Day.ToString(), FontSize = AppTheme.FontXS, Width = GC_DAY_W,
                FontFamily = kbFont, TextAlignment = TextAlignment.Center,
                Foreground = new SolidColorBrush(Color.Parse(color)),
                FontWeight = isToday ? FontWeight.Bold : FontWeight.Normal,
                [Canvas.LeftProperty] = d * GC_DAY_W, [Canvas.TopProperty] = 2.0,
            });
            // 월 1일 경계선
            if (dt.Day == 1)
            {
                dateCanvas.Children.Add(new Avalonia.Controls.Shapes.Line
                {
                    StartPoint = new Point(d * GC_DAY_W, 0),
                    EndPoint = new Point(d * GC_DAY_W, GC_DATE_H),
                    Stroke = new SolidColorBrush(Color.Parse("#333840")), StrokeThickness = 1,
                });
            }
        }
        dateRow.Children.Add(dateCanvas);
        fixedHeader.Children.Add(dateRow);

        // ── 항목별 행 ───────────────────────────────────────────────────────
        foreach (var (fullName, shortName) in analytes)
        {
            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions($"{GC_LABEL_W},*"),
                Height = GC_ROW_H, Margin = new Thickness(0, 0, 0, 1),
            };

            // 라벨: 약칭 배지 + 전체명 (Grid로 구성 — 긴 이름이 트랙 영역으로 넘치지 않도록 * 컬럼에서 ellipsis 트리밍)
            // 트랙 바보다 앞에 그려지도록 ZIndex=1, 솔리드 배경으로 뒤편 가림 방지
            var (lbBg, lbFg) = BadgeColorHelper.GetBadgeColor(shortName);
            var labelPanel = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,4,*"),
                Width = GC_LABEL_W - 4,
                ClipToBounds = true,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2, 0),
                Background = AppTheme.BgPrimary,
                ZIndex = 1,
            };
            var badge = new Border
            {
                Height = 16, MinWidth = 32, CornerRadius = new CornerRadius(3),
                Background = new SolidColorBrush(Color.Parse(lbBg)),
                Padding = new Thickness(4, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = shortName, FontSize = AppTheme.FontXS, FontFamily = kbFont,
                    Foreground = new SolidColorBrush(Color.Parse(lbFg)),
                    VerticalAlignment = VerticalAlignment.Center,
                    TextAlignment = TextAlignment.Center,
                },
            };
            Grid.SetColumn(badge, 0);
            labelPanel.Children.Add(badge);
            var nameTb = new TextBlock
            {
                Text = fullName, FontSize = AppTheme.FontXS, FontFamily = kbFont,
                Foreground = AppTheme.FgSecondary,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
            };
            Grid.SetColumn(nameTb, 2);
            labelPanel.Children.Add(nameTb);
            Grid.SetColumn(labelPanel, 0);

            // 트랙 캔버스
            var track = new Canvas { Width = trackW, Height = GC_ROW_H, ClipToBounds = true, UseLayoutRounding = true };
            Grid.SetColumn(track, 1);
            double barY = (GC_ROW_H - GC_BAR_H) / 2;

            // 드래그&드롭: 직원을 트랙에 놓으면 해당 바 구간의 담당자를 변경
            if (CanEdit)
            {
                DragDrop.SetAllowDrop(track, true);
                var captFullDrop = fullName;
                var captHeaderDrop = fixedHeader; var captBodyDrop = body;
                track.AddHandler(DragDrop.DragOverEvent, (object? _, DragEventArgs de) =>
                {
                    de.DragEffects = de.Data.Contains("agent_name") ? DragDropEffects.Copy : DragDropEffects.None;
                });
                track.AddHandler(DragDrop.DropEvent, (object? _, DragEventArgs de) =>
                {
                    if (!de.Data.Contains("agent_name")) return;
                    string? dropName = de.Data.Get("agent_name") as string;
                    if (string.IsNullOrEmpty(dropName)) return;

                    // 드롭 위치 → 날짜 계산
                    var pos = de.GetPosition(track);
                    int dayIdx = (int)(pos.X / GC_DAY_W);
                    DateTime dropDate = _chartRangeStart.AddDays(dayIdx);

                    // 해당 날짜를 포함하는 스팬 찾기
                    var hitSpan = spans.FirstOrDefault(sp =>
                        string.Equals(sp.FullName, captFullDrop, StringComparison.OrdinalIgnoreCase)
                        && dropDate >= sp.Start && dropDate <= sp.End);

                    if (hitSpan != null)
                    {
                        var hs = hitSpan;
                        ApplySpanChangeToCache(captFullDrop, dropName, hs.Start, hs.End);
                        _chartPendingActions.Add((c, t, cols) =>
                            AnalysisRequestService.UpdateAssignmentByName(c, t, captFullDrop, dropName, hs.Start, hs.End, cols));
                    }
                    else
                    {
                        var dd = dropDate;
                        ApplySpanChangeToCache(captFullDrop, dropName, dd, dd);
                        _chartPendingActions.Add((c, t, cols) =>
                            AnalysisRequestService.UpdateAssignmentByName(c, t, captFullDrop, dropName, dd, dd, cols));
                    }
                    RefreshChartTable(captHeaderDrop, captBodyDrop);
                });
            }

            // 오늘 하이라이트 (붉은색 세로선)
            if (DateTime.Today >= _chartRangeStart && DateTime.Today <= rangeEnd)
            {
                int ti = (int)(DateTime.Today - _chartRangeStart).TotalDays;
                double todayX = ti * GC_DAY_W + GC_DAY_W / 2;
                track.Children.Add(new Avalonia.Controls.Shapes.Line
                {
                    StartPoint = new Point(todayX, 0), EndPoint = new Point(todayX, GC_ROW_H),
                    Stroke = new SolidColorBrush(Color.Parse("#cc4444")), StrokeThickness = 1.5,
                    Opacity = 0.7,
                });
            }

            // 월 경계선 (1일마다)
            for (int d = 0; d < totalDays; d++)
            {
                if (_chartRangeStart.AddDays(d).Day == 1)
                    track.Children.Add(new Avalonia.Controls.Shapes.Line
                    {
                        StartPoint = new Point(d * GC_DAY_W, 0),
                        EndPoint = new Point(d * GC_DAY_W, GC_ROW_H),
                        Stroke = new SolidColorBrush(Color.Parse("#282c30")), StrokeThickness = 1,
                    });
            }

            // 수평 트랙 라인
            track.Children.Add(new Avalonia.Controls.Shapes.Line
            {
                StartPoint = new Point(0, barY + GC_BAR_H / 2),
                EndPoint = new Point(trackW, barY + GC_BAR_H / 2),
                Stroke = AppTheme.BorderSubtle, StrokeThickness = 1,
            });

            // 스팬 바 — solid 배경
            var itemSpans = spans
                .Where(s => string.Equals(s.FullName, fullName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(s => s.Start).ToList();

            foreach (var sp in itemSpans)
            {
                var baseColor = GetChosungColor(sp.Manager);
                double sx = Math.Round(Math.Max(0, (sp.Start - _chartRangeStart).Days) * GC_DAY_W);
                double ex = Math.Round((Math.Min(totalDays - 1, (sp.End - _chartRangeStart).Days) + 1) * GC_DAY_W);
                double w = ex - sx;
                if (w < 1) continue;

                var bgBrush = new SolidColorBrush(Color.FromArgb(180, baseColor.R, baseColor.G, baseColor.B));
                var borderBrush = new SolidColorBrush(Color.FromArgb(255, baseColor.R, baseColor.G, baseColor.B));

                var captHeader = fixedHeader; var captBody = body; var captFull = fullName; var captSp = sp;
                var bar = new Border
                {
                    Width = w, Height = GC_BAR_H,
                    Background = bgBrush,
                    BorderBrush = borderBrush, BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4), Padding = new Thickness(4, 0), ClipToBounds = true,
                    Cursor = CanEdit ? new Cursor(StandardCursorType.Hand) : Cursor.Default,
                    Child = new TextBlock
                    {
                        Text = sp.Manager, FontSize = AppTheme.FontXS, FontFamily = kbFont,
                        Foreground = Brushes.White, FontWeight = FontWeight.SemiBold,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        VerticalAlignment = VerticalAlignment.Center,
                        TextAlignment = TextAlignment.Center,
                    },
                    [Canvas.LeftProperty] = sx, [Canvas.TopProperty] = barY,
                };

                // 더블클릭 → 담당자 변경 Flyout
                if (CanEdit)
                {
                    bar.PointerPressed += (s, e) =>
                    {
                        if (e.ClickCount < 2) return;
                        e.Handled = true;
                        var names = AgentService.GetAllNames();
                        var lb = new ListBox { MaxHeight = 200, MinWidth = 120, ItemsSource = names,
                            FontSize = AppTheme.FontBase, FontFamily = kbFont,
                            Background = AppTheme.BgPrimary, Foreground = AppTheme.FgSecondary };
                        var fly = new Flyout { Content = lb, Placement = PlacementMode.Bottom };
                        lb.SelectionChanged += (_, _) =>
                        {
                            if (lb.SelectedItem is not string nm) return;
                            var cFull = captFull; var cStart = captSp.Start; var cEnd = captSp.End;
                            ApplySpanChangeToCache(cFull, nm, cStart, cEnd);
                            _chartPendingActions.Add((c, t, cols) =>
                                AnalysisRequestService.UpdateAssignmentByName(c, t, cFull, nm, cStart, cEnd, cols));
                            fly.Hide();
                            RefreshChartTable(captHeader, captBody);
                        };
                        fly.ShowAt((Control)s!);
                    };

                    // 우클릭 → 담당자 삭제
                    bar.PointerPressed += (s, e) =>
                    {
                        if (!e.GetCurrentPoint((Control)s!).Properties.IsRightButtonPressed) return;
                        e.Handled = true;
                        var delBtn = new Button
                        {
                            Content = $"{captSp.Manager} 삭제 ({captSp.Start:M/d}~{captSp.End:M/d})",
                            FontSize = AppTheme.FontSM, FontFamily = kbFont,
                            Background = new SolidColorBrush(Color.Parse("#3a1a1a")),
                            Foreground = new SolidColorBrush(Color.Parse("#ff6666")),
                            BorderBrush = new SolidColorBrush(Color.Parse("#6a2a2a")),
                            Padding = new Thickness(12, 6),
                        };
                        var fly = new Flyout { Content = delBtn, Placement = PlacementMode.Bottom };
                        delBtn.Click += (_, _) =>
                        {
                            var cFull = captFull; var cStart = captSp.Start; var cEnd = captSp.End;
                            ClearSpanFromCache(cFull, cStart, cEnd);
                            _chartPendingActions.Add((c, t, cols) =>
                                AnalysisRequestService.ClearAssignmentByName(c, t, cFull, cStart, cEnd));
                            fly.Hide();
                            RefreshChartTable(captHeader, captBody);
                        };
                        fly.ShowAt((Control)s!);
                    };
                }
                track.Children.Add(bar);
            }

            // ── 각 바의 양쪽 끝 드래그 핸들 (◁ 좌측=시작일, ▷ 우측=종료일) ──
            if (CanEdit)
            {
                for (int idx = 0; idx < itemSpans.Count; idx++)
                {
                    var sp2 = itemSpans[idx];
                    double spSx = Math.Round(Math.Max(0, (sp2.Start - _chartRangeStart).Days) * GC_DAY_W);
                    double spEx = Math.Round((Math.Min(totalDays - 1, (sp2.End - _chartRangeStart).Days) + 1) * GC_DAY_W);

                    // 이전/다음 바 참조 (경계 제한용)
                    var prevSpan = idx > 0 ? itemSpans[idx - 1] : (AnalysisRequestService.AssignmentSpan?)null;
                    var nextSpan = idx < itemSpans.Count - 1 ? itemSpans[idx + 1] : (AnalysisRequestService.AssignmentSpan?)null;

                    // --- 좌측 핸들 (시작일 변경) ---
                    AddEdgeHandle(track, spSx, barY, sp2, fullName, isLeft: true,
                        prevSpan, nextSpan, kbFont, fixedHeader, body, rangeEnd);

                    // --- 우측 핸들 (종료일 변경) ---
                    AddEdgeHandle(track, spEx, barY, sp2, fullName, isLeft: false,
                        prevSpan, nextSpan, kbFont, fixedHeader, body, rangeEnd);
                }
            }

            // 경계 세로선
            for (int i = 0; i < itemSpans.Count - 1; i++)
            {
                DateTime bnd = itemSpans[i + 1].Start;
                double bx = (bnd - _chartRangeStart).TotalDays * GC_DAY_W;
                track.Children.Add(new Avalonia.Controls.Shapes.Line
                {
                    StartPoint = new Point(bx, barY - 1), EndPoint = new Point(bx, barY + GC_BAR_H + 1),
                    Stroke = new SolidColorBrush(Color.FromArgb(50, 160, 170, 190)), StrokeThickness = 1,
                });
            }

            rowGrid.Children.Add(track);
            // 라벨 패널을 마지막에 추가 → 트랙(바)보다 앞에 렌더링되어 가려지지 않음
            rowGrid.Children.Add(labelPanel);
            body.Children.Add(rowGrid);
        }
    }

    /// <summary>바 양쪽 끝 드래그 핸들 생성 (좌측=시작일, 우측=종료일)</summary>
    /// <summary>캐시 스팬을 로컬 수정 (화면 즉시 반영용)</summary>
    private void ApplySpanChangeToCache(string fullName, string manager, DateTime newStart, DateTime newEnd)
    {
        if (_chartCachedSpans == null) return;
        // 해당 fullName의 겹치는 스팬 제거 후 새 스팬 추가
        _chartCachedSpans.RemoveAll(s =>
            string.Equals(s.FullName, fullName, StringComparison.OrdinalIgnoreCase)
            && s.Start <= newEnd && s.End >= newStart);
        if (!string.IsNullOrEmpty(manager))
        {
            var info = AnalysisRequestService.GetStandardDaysInfo();
            string shortName = info.TryGetValue(fullName, out var v) ? v.shortName : fullName;
            _chartCachedSpans.Add(new AnalysisRequestService.AssignmentSpan(fullName, shortName, manager, newStart, newEnd));
        }
    }

    /// <summary>캐시에서 스팬 구간 삭제 (화면 즉시 반영용)</summary>
    private void ClearSpanFromCache(string fullName, DateTime start, DateTime end)
    {
        if (_chartCachedSpans == null) return;
        var toModify = _chartCachedSpans
            .Where(s => string.Equals(s.FullName, fullName, StringComparison.OrdinalIgnoreCase)
                && s.Start <= end && s.End >= start)
            .ToList();
        foreach (var s in toModify)
        {
            _chartCachedSpans.Remove(s);
            // 잘려서 남는 부분 재추가
            if (s.Start < start)
                _chartCachedSpans.Add(new AnalysisRequestService.AssignmentSpan(s.FullName, s.ShortName, s.Manager, s.Start, start.AddDays(-1)));
            if (s.End > end)
                _chartCachedSpans.Add(new AnalysisRequestService.AssignmentSpan(s.FullName, s.ShortName, s.Manager, end.AddDays(1), s.End));
        }
    }

    private void AddEdgeHandle(Canvas track, double edgeX, double barY,
        AnalysisRequestService.AssignmentSpan span, string fullName, bool isLeft,
        AnalysisRequestService.AssignmentSpan? prevSpan, AnalysisRequestService.AssignmentSpan? nextSpan,
        FontFamily kbFont, StackPanel fixedHeader, StackPanel body, DateTime chartRangeEnd)
    {
        const double HANDLE_W = 10;
        var handle = new Avalonia.Controls.Shapes.Rectangle
        {
            Width = HANDLE_W, Height = GC_BAR_H + 6,
            Fill = new SolidColorBrush(Color.FromArgb(0, 255, 255, 255)),
            Cursor = new Cursor(StandardCursorType.SizeWestEast),
            [Canvas.LeftProperty] = edgeX - HANDLE_W / 2, [Canvas.TopProperty] = barY - 3,
        };

        var baseColor = GetChosungColor(span.Manager);
        var hoverColor = Color.FromArgb(100, baseColor.R, baseColor.G, baseColor.B);

        handle.PointerEntered += (s, _) =>
            ((Avalonia.Controls.Shapes.Rectangle)s!).Fill = new SolidColorBrush(hoverColor);
        handle.PointerExited += (s, _) =>
            ((Avalonia.Controls.Shapes.Rectangle)s!).Fill = new SolidColorBrush(Color.FromArgb(0, 255, 255, 255));

        var captSpan = span; var captFn = fullName;
        bool dragging = false; double dragStartX = 0;
        DateTime origDate = isLeft ? span.Start : span.End;

        handle.PointerPressed += (s, e) =>
        {
            dragging = true;
            dragStartX = e.GetPosition(track).X;
            origDate = isLeft ? captSpan.Start : captSpan.End;
            e.Pointer.Capture((IInputElement)s!);
            e.Handled = true;
        };

        handle.PointerMoved += (s, e) =>
        {
            if (!dragging) return;
            double curX = e.GetPosition(track).X;
            int dayDelta = (int)Math.Round((curX - dragStartX) / GC_DAY_W);
            if (dayDelta == 0) return;

            DateTime newDate = origDate.AddDays(dayDelta);

            if (isLeft)
            {
                // 시작일: 이전 바 시작일+1 이후, 자기 종료일 이전
                DateTime minDate = prevSpan != null ? prevSpan.Start.AddDays(1) : _chartRangeStart;
                if (newDate < minDate) newDate = minDate;
                if (newDate > captSpan.End) newDate = captSpan.End;
            }
            else
            {
                // 종료일: 자기 시작일 이후, 다음 바 종료일 이전
                DateTime maxDate = nextSpan != null ? nextSpan.End : chartRangeEnd;
                if (newDate < captSpan.Start) newDate = captSpan.Start;
                if (newDate > maxDate) newDate = maxDate;
            }

            double newPx = isLeft
                ? (newDate - _chartRangeStart).TotalDays * GC_DAY_W
                : ((newDate - _chartRangeStart).TotalDays + 1) * GC_DAY_W;
            Canvas.SetLeft((Control)s!, newPx - HANDLE_W / 2);

            // 인디케이터 라인
            string tag = $"edge_{captFn}_{(isLeft ? "L" : "R")}_{captSpan.Start:yyyyMMdd}";
            var indicator = track.Children.OfType<Avalonia.Controls.Shapes.Line>()
                .FirstOrDefault(l => l.Tag is string t && t == tag);
            if (indicator != null)
            {
                indicator.StartPoint = new Point(newPx, barY - 3);
                indicator.EndPoint = new Point(newPx, barY + GC_BAR_H + 3);
            }
            else
            {
                track.Children.Add(new Avalonia.Controls.Shapes.Line
                {
                    StartPoint = new Point(newPx, barY - 3),
                    EndPoint = new Point(newPx, barY + GC_BAR_H + 3),
                    Stroke = new SolidColorBrush(Color.Parse("#ff8844")), StrokeThickness = 2, Tag = tag,
                });
            }
        };

        handle.PointerReleased += (s, e) =>
        {
            if (!dragging) return;
            dragging = false;
            e.Pointer.Capture(null);

            double curX = e.GetPosition(track).X;
            int dayDelta = (int)Math.Round((curX - dragStartX) / GC_DAY_W);
            if (dayDelta == 0) return;

            DateTime newDate = origDate.AddDays(dayDelta);
            DateTime newStart, newEnd;

            if (isLeft)
            {
                DateTime minDate = prevSpan != null ? prevSpan.Start.AddDays(1) : _chartRangeStart;
                if (newDate < minDate) newDate = minDate;
                if (newDate > captSpan.End) newDate = captSpan.End;
                newStart = newDate;
                newEnd = captSpan.End;
            }
            else
            {
                DateTime maxDate = nextSpan != null ? nextSpan.End : chartRangeEnd;
                if (newDate < captSpan.Start) newDate = captSpan.Start;
                if (newDate > maxDate) newDate = maxDate;
                newStart = captSpan.Start;
                newEnd = newDate;
            }

            bool shiftHeld = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

            if (shiftHeld)
            {
                // Shift+드래그: 전체 항목 일괄
                var captNewStart = newStart; var captNewEnd = newEnd;
                var captIsLeft = isLeft; var captCaptSpan = captSpan;
                var captChartStart = _chartRangeStart; var captChartEnd = chartRangeEnd;

                // 캐시에서 전체 항목 즉시 수정 (화면 반영)
                if (_chartCachedSpans != null)
                {
                    var analytes = _chartCachedSpans.Select(sp => sp.FullName).Distinct().ToList();
                    foreach (var fn in analytes)
                    {
                        var itemSpans = _chartCachedSpans
                            .Where(sp2 => string.Equals(sp2.FullName, fn, StringComparison.OrdinalIgnoreCase))
                            .OrderBy(sp2 => sp2.Start).ToList();
                        var target = captIsLeft
                            ? itemSpans.FirstOrDefault(sp2 => sp2.Start == captCaptSpan.Start)
                            : itemSpans.FirstOrDefault(sp2 => sp2.End == captCaptSpan.End);
                        if (target == null) continue;

                        DateTime tNewStart = captIsLeft ? captNewStart : target.Start;
                        DateTime tNewEnd = captIsLeft ? target.End : captNewEnd;

                        if (captIsLeft && tNewStart > target.Start)
                            ClearSpanFromCache(fn, target.Start, tNewStart.AddDays(-1));
                        if (!captIsLeft && tNewEnd < target.End)
                            ClearSpanFromCache(fn, tNewEnd.AddDays(1), target.End);
                        ApplySpanChangeToCache(fn, target.Manager, tNewStart, tNewEnd);
                    }
                }

                // DB 반영 액션 — 항목별로 분리 등록 (프로그레스바 표시용)
                var allSpansSnap = AnalysisRequestService.GetAssignmentChartData(captChartStart, captChartEnd);
                var analytes2 = allSpansSnap.Select(sp => sp.FullName).Distinct().ToList();
                foreach (var fn2 in analytes2)
                {
                    var fnLocal = fn2;
                    var itemSpans2 = allSpansSnap
                        .Where(sp2 => string.Equals(sp2.FullName, fnLocal, StringComparison.OrdinalIgnoreCase))
                        .OrderBy(sp2 => sp2.Start).ToList();
                    var target2 = captIsLeft
                        ? itemSpans2.FirstOrDefault(sp2 => sp2.Start == captCaptSpan.Start)
                        : itemSpans2.FirstOrDefault(sp2 => sp2.End == captCaptSpan.End);
                    if (target2 == null) continue;

                    DateTime tNewStart2 = captIsLeft ? captNewStart : target2.Start;
                    DateTime tNewEnd2 = captIsLeft ? target2.End : captNewEnd;
                    var tgt = target2;

                    _chartPendingActions.Add((c, t, cols) =>
                    {
                        if (captIsLeft && tNewStart2 > tgt.Start)
                            AnalysisRequestService.ClearAssignmentByName(c, t, fnLocal, tgt.Start, tNewStart2.AddDays(-1));
                        if (!captIsLeft && tNewEnd2 < tgt.End)
                            AnalysisRequestService.ClearAssignmentByName(c, t, fnLocal, tNewEnd2.AddDays(1), tgt.End);
                        AnalysisRequestService.UpdateAssignmentByName(c, t, fnLocal, tgt.Manager, tNewStart2, tNewEnd2, cols);
                    });
                }
            }
            else
            {
                // 개별 — 캐시 즉시 수정 (화면 반영)
                if (isLeft && newStart > captSpan.Start)
                    ClearSpanFromCache(captFn, captSpan.Start, newStart.AddDays(-1));
                if (!isLeft && newEnd < captSpan.End)
                    ClearSpanFromCache(captFn, newEnd.AddDays(1), captSpan.End);
                ApplySpanChangeToCache(captFn, captSpan.Manager, newStart, newEnd);

                // DB 반영 액션 등록
                var captNewStart = newStart; var captNewEnd = newEnd;
                var captNewDate = newDate;
                _chartPendingActions.Add((c, t, cols) =>
                {
                    if (isLeft && captNewStart > captSpan.Start)
                        AnalysisRequestService.ClearAssignmentByName(c, t, captFn, captSpan.Start, captNewStart.AddDays(-1));
                    if (!isLeft && captNewEnd < captSpan.End)
                        AnalysisRequestService.ClearAssignmentByName(c, t, captFn, captNewEnd.AddDays(1), captSpan.End);

                    if (prevSpan != null && isLeft)
                        AnalysisRequestService.UpdateAssignmentByName(c, t, captFn, prevSpan.Manager, prevSpan.Start, captNewDate.AddDays(-1), cols);
                    if (nextSpan != null && !isLeft)
                        AnalysisRequestService.UpdateAssignmentByName(c, t, captFn, nextSpan.Manager, captNewDate.AddDays(1), nextSpan.End, cols);
                    AnalysisRequestService.UpdateAssignmentByName(c, t, captFn, captSpan.Manager, captNewStart, captNewEnd, cols);
                });
            }

            // 화면 즉시 갱신 (캐시 기반)
            RefreshChartTable(fixedHeader, body);
            RefreshShow3AfterChartUpdate();
        };

        track.Children.Add(handle);
    }

    /// <summary>wire-v01 요약 스트립 — 배정 상태 카운트 배지</summary>
    private void RefreshSummaryStrip(
        List<AnalysisRequestService.AssignmentSpan> spans,
        List<string> analyteFullNames,
        DateTime rangeEnd)
    {
        if (_chartSummaryPanel == null) return;
        _chartSummaryPanel.Children.Clear();

        var today = DateTime.Today;
        var itemsWithSpan = spans.Select(s => s.FullName).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet();
        int totalItems = analyteFullNames.Count;
        int assignedItems = analyteFullNames.Count(n => itemsWithSpan.Contains(n));
        int unassignedItems = totalItems - assignedItems;
        int endingSoon = spans.Count(s => s.End >= today && s.End <= today.AddDays(7));
        int managers = spans.Select(s => s.Manager).Where(m => !string.IsNullOrWhiteSpace(m))
                            .Distinct(StringComparer.OrdinalIgnoreCase).Count();

        _chartSummaryPanel.Children.Add(StatusBadge.Info($"항목 {totalItems}", withIcon: false));
        _chartSummaryPanel.Children.Add(StatusBadge.Ok($"배정 {assignedItems}"));
        if (unassignedItems > 0)
            _chartSummaryPanel.Children.Add(StatusBadge.Bad($"미배정 {unassignedItems}"));
        else
            _chartSummaryPanel.Children.Add(StatusBadge.Muted($"미배정 {unassignedItems}"));
        if (endingSoon > 0)
            _chartSummaryPanel.Children.Add(StatusBadge.Warn($"7일 내 만료 {endingSoon}"));
        _chartSummaryPanel.Children.Add(StatusBadge.Accent($"담당자 {managers}"));
    }

    /// <summary>업무분장표 — 계약업체 탭</summary>
    private void RefreshChartContractTab(StackPanel body)
    {
        body.Children.Clear();
        var kbFont = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        // Agent별 담당업체 목록 조회
        var agents = AgentService.GetAllItems().Where(a => !string.IsNullOrWhiteSpace(a.담당업체)).ToList();
        if (agents.Count == 0)
        {
            body.Children.Add(new TextBlock
            {
                Text = "배정된 계약업체가 없습니다.", FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgMuted, Margin = new Thickness(12, 20),
            });
            return;
        }

        // 직원별 행: 이름 | 담당업체 목록
        foreach (var agent in agents.OrderBy(a => a.성명))
        {
            var rowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions($"{GC_LABEL_W},*"),
                MinHeight = GC_ROW_H, Margin = new Thickness(0, 0, 0, 1),
            };

            // 이름 배지
            var nameColor = GetChosungColor(agent.성명);
            var label = new Border
            {
                Height = GC_BAR_H, CornerRadius = new CornerRadius(3),
                Background = new SolidColorBrush(Color.FromArgb(25, nameColor.R, nameColor.G, nameColor.B)),
                BorderBrush = new SolidColorBrush(Color.FromArgb(50, nameColor.R, nameColor.G, nameColor.B)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(3, 0), VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = agent.성명, FontSize = AppTheme.FontXS, FontFamily = kbFont,
                    Foreground = new SolidColorBrush(nameColor),
                    VerticalAlignment = VerticalAlignment.Center,
                },
            };
            Grid.SetColumn(label, 0);
            rowGrid.Children.Add(label);

            // 업체 칩 목록 — wire-v01 통일 StatusBadge
            var chips = new WrapPanel { Margin = new Thickness(4, 2), };
            foreach (var company in agent.담당업체목록)
            {
                var chip = ETA.Views.Controls.StatusBadge.Info(company, withIcon: false);
                chip.Margin = new Thickness(2, 1);
                chips.Children.Add(chip);
            }
            Grid.SetColumn(chips, 1);
            rowGrid.Children.Add(chips);
            body.Children.Add(rowGrid);
        }
    }

    /// <summary>업무분장표 — 일반업무 탭</summary>
    private void RefreshChartGeneralTab(StackPanel body)
    {
        body.Children.Clear();
        var kbFont = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        var tasks = ETA.Services.Common.GeneralTaskService.GetAll();
        if (tasks.Count == 0)
        {
            body.Children.Add(new TextBlock
            {
                Text = "등록된 일반업무가 없습니다.", FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgMuted, Margin = new Thickness(12, 20),
            });
            return;
        }

        // 헤더
        var hdr = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,80,80,80"),
            Height = GC_DATE_H, Margin = new Thickness(0, 0, 0, 2),
        };
        string[] hLabels = ["업무명", "담당자", "상태", "마감일"];
        for (int i = 0; i < hLabels.Length; i++)
        {
            var tb = new TextBlock
            {
                Text = hLabels[i], FontSize = AppTheme.FontXS, FontFamily = kbFont,
                Foreground = AppTheme.FgMuted, FontWeight = FontWeight.SemiBold,
                Margin = new Thickness(4, 0),
            };
            Grid.SetColumn(tb, i);
            hdr.Children.Add(tb);
        }
        body.Children.Add(hdr);

        // 행
        int idx = 0;
        foreach (var task in tasks.OrderBy(t => t.상태 == "완료" ? 1 : 0).ThenBy(t => t.마감일))
        {
            var row = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,80,80,80"),
                MinHeight = GC_ROW_H, Margin = new Thickness(0, 0, 0, 1),
                Background = idx++ % 2 == 0 ? AppRes("GridRowBg") : AppRes("GridRowAltBg"),
            };

            bool done = task.상태 == "완료";
            var fgColor = done ? AppTheme.FgMuted : AppTheme.FgSecondary;

            row.Children.Add(new TextBlock
            {
                Text = task.업무명, FontSize = AppTheme.FontSM, FontFamily = kbFont,
                Foreground = fgColor, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0), TextTrimming = TextTrimming.CharacterEllipsis,
            });

            var nameBlock = new TextBlock
            {
                Text = task.담당자명, FontSize = AppTheme.FontSM, FontFamily = kbFont,
                Foreground = fgColor, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
            };
            Grid.SetColumn(nameBlock, 1);
            row.Children.Add(nameBlock);

            string statusIcon = task.상태 switch { "완료" => "✅", "진행" => "🔵", _ => "⏳" };
            var statusBlock = new TextBlock
            {
                Text = $"{statusIcon} {task.상태}", FontSize = AppTheme.FontXS, FontFamily = kbFont,
                Foreground = fgColor, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
            };
            Grid.SetColumn(statusBlock, 2);
            row.Children.Add(statusBlock);

            var deadlineBlock = new TextBlock
            {
                Text = task.마감일, FontSize = AppTheme.FontXS, FontFamily = kbFont,
                Foreground = fgColor, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0),
            };
            Grid.SetColumn(deadlineBlock, 3);
            row.Children.Add(deadlineBlock);

            body.Children.Add(row);
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

    // =========================================================================
    // 일반업무 — Show2 목록 (4열 그리드)
    // =========================================================================
    private void RefreshGeneralTasks(Agent agent)
    {
        if (_generalTaskPanel == null) return;
        _generalTaskPanel.Children.Clear();

        var tasks = GeneralTaskService.GetByAgent(agent.사번);
        if (tasks.Count == 0)
        {
            _generalTaskPanel.Children.Add(new TextBlock
            {
                Text = "배정된 일반업무 없음", FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = AppTheme.FgDimmed,
                Margin = new Thickness(0, 2),
            });
            return;
        }

        const int Cols = 4;
        var taskGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,4,*,4,*,4,*") };

        for (int i = 0; i < tasks.Count; i++)
        {
            var t = tasks[i];
            int gridCol = (i % Cols) * 2;   // 0, 2, 4, 6
            int gridRow = i / Cols;
            if (taskGrid.RowDefinitions.Count <= gridRow)
                taskGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

            var statusColor = t.상태 switch
            {
                "진행" => ("#1a2a2a", "#66ccbb", "#2a6a5a"),
                "완료" => ("#1a2a3a", "#88aacc", "#336699"),
                _      => ("#2a2a2a", "#cccc88", "#666633"),
            };

            var capturedTask  = t;
            var capturedAgent = agent;
            var card = new Border
            {
                Background      = new SolidColorBrush(Color.Parse(statusColor.Item1)),
                BorderBrush     = new SolidColorBrush(Color.Parse(statusColor.Item3)),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(4),
                Padding         = new Thickness(5, 3),
                Margin          = new Thickness(0, 1),
                Cursor          = new Cursor(StandardCursorType.Hand),
                [Grid.ColumnProperty] = gridCol,
                [Grid.RowProperty]    = gridRow,
            };
            // 호버시 × 버튼 표시
            var btnXTask = new Button
            {
                Content = "×", FontSize = AppTheme.FontBase, Padding = new Thickness(2, 0),
                Background = Brushes.Transparent, BorderThickness = new Thickness(0),
                Foreground = new SolidColorBrush(Color.Parse("#dd6688")),
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Right,
                IsVisible = false,
            };
            var taskId = t.Id;
            btnXTask.Click += (_, e) =>
            {
                e.Handled = true;
                GeneralTaskService.Delete(taskId);
                RefreshGeneralTasks(agent);
            };

            var inner = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            var nameStack = new StackPanel { Spacing = 1 };
            nameStack.Children.Add(new TextBlock
            {
                Text = t.업무명, FontSize = AppTheme.FontSM,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = new SolidColorBrush(Color.Parse(statusColor.Item2)),
                TextTrimming = TextTrimming.CharacterEllipsis,
            });
            var sub = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            sub.Children.Add(new Border
            {
                Background   = new SolidColorBrush(Color.Parse(statusColor.Item3)),
                CornerRadius = new CornerRadius(2),
                Padding      = new Thickness(3, 0),
                Child = new TextBlock
                {
                    Text = t.상태, FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = new SolidColorBrush(Color.Parse(statusColor.Item2)),
                },
            });
            if (!string.IsNullOrEmpty(t.마감일))
                sub.Children.Add(new TextBlock
                {
                    Text = t.마감일, FontSize = AppTheme.FontXS,
                    FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                    Foreground = AppTheme.FgDimmed,
                    VerticalAlignment = VerticalAlignment.Center,
                });
            nameStack.Children.Add(sub);
            inner.Children.Add(nameStack);
            inner.Children.Add(btnXTask);
            Grid.SetColumn(btnXTask, 1);
            card.Child = inner;
            card.PointerEntered += (_, _) => btnXTask.IsVisible = true;
            card.PointerExited  += (_, _) => btnXTask.IsVisible = false;
            card.PointerPressed += (_, _) => Show3ContentRequest?.Invoke(BuildGeneralTaskForm(capturedAgent, capturedTask));
            taskGrid.Children.Add(card);
        }

        _generalTaskPanel.Children.Add(taskGrid);
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [AgentTree] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText("Logs/AgentDebug.log", line + Environment.NewLine); } catch { }
        }
    }
}
