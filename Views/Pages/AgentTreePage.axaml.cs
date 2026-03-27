using Avalonia;
using Avalonia.Controls;
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

namespace ETA.Views.Pages;

public partial class AgentTreePage : UserControl
{
    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;
    public ListBox? AnalysisItemsListBox { get; set; }

    // ── 상태 ────────────────────────────────────────────────────────────────
    private Agent?      _selectedAgent;
    private StackPanel? _detailPanel;
    private bool        _isAddMode  = false;

    // 사진 미리보기 Image 컨트롤 (저장 시 PhotoPath 접근용)
    private Image?      _photoImage;
    private string      _pendingPhotoPath = "";   // 선택했지만 아직 저장 안 된 경로

    // 업무 분장 저장 시 참조 (BuildAssignmentArea에서 설정)
    private ListBox?    _assignmentListBox;
    private DateTime    _assignmentRangeStart = DateTime.Today;
    private DateTime    _assignmentRangeEnd   = DateTime.Today;

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
        string icon = agent.기타 switch
        {
            "0" => "🥷", "1" => "👨‍💼", "2" => "👩‍⚖️", _ => "👤"
        };

        return new TreeViewItem
        {
            Tag    = agent,
            Header = new StackPanel
            {
                Orientation       = Orientation.Horizontal,
                Spacing           = 8,
                VerticalAlignment = VerticalAlignment.Center,
                Children =
                {
                    new TextBlock { Text = icon, FontSize = 20, VerticalAlignment = VerticalAlignment.Center },
                    new TextBlock
                    {
                        Text              = agent.성명,
                        FontSize          = 13,
                        FontFamily        = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
                        Foreground        = Brushes.WhiteSmoke,
                        VerticalAlignment = VerticalAlignment.Center
                    }
                }
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
        Log($"선택: {agent.성명}");
    }

    // =========================================================================
    // 직원 추가 패널  (MainPage BT3)
    // =========================================================================
    public void ShowAddPanel()
    {
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
        if (_isAddMode) SaveAdd();
        else            SaveEdit();
    }

    // =========================================================================
    // 삭제  (MainPage BT4)
    // =========================================================================
    public async Task DeleteSelectedAsync()
    {
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
                    Foreground = Brushes.WhiteSmoke,
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
        var root = MakeRootPanel($"✏️  {agent.성명} — 정보 수정");

        // 사진 영역
        root.Children.Add(BuildPhotoArea(agent.PhotoPath));

        root.Children.Add(BuildFieldRow("성명",           agent.성명,           isReadOnly: true, isLocked: true));
        root.Children.Add(BuildFieldRow("직급",           agent.직급));
        root.Children.Add(BuildFieldRow("직무",           agent.직무));
        root.Children.Add(BuildFieldRow("사번",           agent.사번));
        root.Children.Add(BuildFieldRow("입사일",         agent.입사일표시,     isReadOnly: true));
        root.Children.Add(BuildFieldRow("자격사항",       agent.자격사항));
        root.Children.Add(BuildFieldRow("Email",          agent.Email));
        root.Children.Add(BuildFieldRow("측정인고유번호", agent.측정인고유번호));

        // 분장 영역
        root.Children.Add(BuildAssignmentArea(agent));

        return root;
    }

    // =========================================================================
    // 패널 — 추가 모드
    // =========================================================================
    private StackPanel BuildAddPanel()
    {
        var root = MakeRootPanel("➕  신규 직원 추가");

        root.Children.Add(BuildPhotoArea(""));

        root.Children.Add(BuildFieldRow("성명",           "", hint: "이름 입력 (필수)"));
        root.Children.Add(BuildFieldRow("직급",           ""));
        root.Children.Add(BuildFieldRow("직무",           ""));
        root.Children.Add(BuildFieldRow("사번",           ""));
        root.Children.Add(BuildFieldRow("입사일",         "", hint: "예) 2024-01-01"));
        root.Children.Add(BuildFieldRow("자격사항",       ""));
        root.Children.Add(BuildFieldRow("Email",          ""));
        root.Children.Add(BuildFieldRow("측정인고유번호", ""));

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

        // ── 드랍 영역 ────────────────────────────────────────────────────
        var dropListBox = new ListBox
        {
            Background      = new SolidColorBrush(Color.Parse("#2a2a3a")),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555")),
            Height          = 130,
            Margin          = new Thickness(0, 4, 0, 0)
        };

        DragDrop.SetAllowDrop(dropListBox, true);
        _assignmentListBox = dropListBox;

        // 날짜 범위 상태 (클로저로 공유)
        DateTime rangeStart = DateTime.Today;
        DateTime rangeEnd   = DateTime.Today;

        // 드래그 드랍 — 중복 체크 포함 async 핸들러
        dropListBox.AddHandler(DragDrop.DropEvent, async (object? sender, DragEventArgs e) =>
        {
            if (sender is not ListBox lb) return;
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
            LoadAssignments(lb, agent, rangeStart, rangeEnd);
            Log($"분장 추가: {agent.성명} ← {analyte}");
        });

        dropListBox.Tag = agent;

        LoadAssignments(dropListBox, agent, rangeStart, rangeEnd);
        stack.Children.Add(dropListBox);

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
            Foreground          = Brushes.WhiteSmoke,
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
            LoadAssignments(dropListBox, agent, rangeStart, rangeEnd);
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
            LoadAssignments(dropListBox, agent, rangeStart, rangeEnd);
            // 시작·종료 날짜가 다르면 범위 선택 완료 → 달력 자동 닫기
            if (rangeStart != rangeEnd)
                calendar.IsVisible = false;
            Log($"기간 조회: {rangeStart:yyyy-MM-dd} ~ {rangeEnd:yyyy-MM-dd}");
        };

        btnAssignSave.Click += async (_, _) =>
        {
            var analytes = dropListBox.Items
                .OfType<ListBoxItem>()
                .Select(i => i.Tag as string)
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            saveProgress.IsVisible  = true;
            btnAssignSave.IsEnabled = false;

            await Task.Run(() =>
            {
                AnalysisRequestService.ClearAssignmentsForAgent(
                    agent.사번, rangeStart, rangeEnd);
                foreach (var a in analytes)
                    AnalysisRequestService.AddAssignment(
                        agent.사번, a!, rangeStart, rangeEnd);
            });

            saveProgress.IsVisible  = false;
            btnAssignSave.IsEnabled = true;
            Log($"분장 저장 완료: {analytes.Count}개 ({rangeStart:yyyy-MM-dd} ~ {rangeEnd:yyyy-MM-dd})");
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
                    Foreground   = Brushes.WhiteSmoke,
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
    private StackPanel BuildPhotoArea(string photoPath)
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

        // 초기 사진 로드
        LoadPhotoToImage(_photoImage, photoPath);

        // 버튼들
        var btnPanel = new StackPanel { Spacing = 6, VerticalAlignment = VerticalAlignment.Center };

        var uploadBtn = new Button
        {
            Content         = "📷 사진 업로드",
            FontSize        = 11,
            FontFamily      = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Background      = new SolidColorBrush(Color.Parse("#3a4a6a")),
            Foreground      = Brushes.WhiteSmoke,
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
            Foreground      = Brushes.WhiteSmoke,
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
    private static void LoadPhotoToImage(Image img, string pathOrFileName)
    {
        if (string.IsNullOrEmpty(pathOrFileName)) { img.Source = null; return; }

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

    // ── 패널에서 특정 필드 값 읽기 ───────────────────────────────────────────
    private static string ReadFieldFromPanel(StackPanel panel, string fieldLabel)
    {
        foreach (var child in panel.Children.OfType<StackPanel>())
        {
            if (child.Children.Count < 2) continue;
            var label = (child.Children[0] as TextBlock)?.Text ?? "";
            label = label.Replace("🔒 ", "").Replace("    ", "").Replace(" :", "").Trim();
            if (label == fieldLabel && child.Children[1] is TextBox tb)
                return tb.Text ?? "";
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
            _pendingPhotoPath = "";

            // 업무 분장 동기화: 기간 내 기존 분장 초기화 후 현재 목록으로 재등록
            if (_assignmentListBox != null)
            {
                var analytes = _assignmentListBox.Items
                    .OfType<ListBoxItem>()
                    .Select(i => i.Tag as string)
                    .Where(t => !string.IsNullOrEmpty(t))
                    .ToList();

                AnalysisRequestService.ClearAssignmentsForAgent(
                    _selectedAgent.사번, _assignmentRangeStart, _assignmentRangeEnd);

                foreach (var a in analytes)
                    AnalysisRequestService.AddAssignment(
                        _selectedAgent.사번, a!, _assignmentRangeStart, _assignmentRangeEnd);

                Log($"분장 저장: {analytes.Count}개 ({_assignmentRangeStart:yyyy-MM-dd} ~ {_assignmentRangeEnd:yyyy-MM-dd})");
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
        foreach (var child in panel.Children.OfType<StackPanel>())
        {
            if (child.Children.Count < 2) continue;
            var tb = child.Children[1] as TextBox;
            if (tb == null) continue;
            if (tb.IsReadOnly && !includeReadOnly) continue;

            var label = (child.Children[0] as TextBlock)?.Text ?? "";
            label = label.Replace("🔒 ", "").Replace("    ", "").Replace(" :", "").Trim();

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
            Foreground = Brushes.WhiteSmoke,
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
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };

        panel.Children.Add(new TextBlock
        {
            Text              = (isLocked ? "🔒 " : "    ") + label + " :",
            Width             = 140,
            FontSize          = 12,
            FontFamily        = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Foreground        = isLocked
                                    ? new SolidColorBrush(Color.Parse("#888888"))
                                    : Brushes.LightGray,
            VerticalAlignment = VerticalAlignment.Center
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            Width           = 260,
            FontSize        = 12,
            FontFamily      = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            IsReadOnly      = isReadOnly,
            Watermark       = hint,
            Background      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#252525"))
                                  : new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#666666"))
                                  : Brushes.WhiteSmoke,
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#333333"))
                                  : new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 4)
        });

        return panel;
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [AgentTree] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("AgentDebug.log", line + Environment.NewLine); } catch { }
    }
}
