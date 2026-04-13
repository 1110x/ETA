using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 물품구매 페이지
///   TreeControl  → Show1 (왼쪽: 연/월 트리)
///   ListControl  → Show2 (상단: 해당 월 리스트)
///   FormControl  → Show3 (하단: 신규 추가 폼)
/// </summary>
public class PurchasePage
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── MainPage 에서 각 영역에 꽂을 컨트롤 ─────────────────────────────────
    public Control TreeControl { get; }
    public Control ListControl { get; }
    public Control FormControl { get; }

    // ── 내부 상태 ────────────────────────────────────────────────────────────
    private readonly StackPanel _listPanel;
    private readonly StackPanel _treePanel;     // 연/월 노드가 들어갈 패널
    private List<PurchaseItem>  _items = new();
    private int _filterYear  = 0;
    private int _filterMonth = 0;

    // 현재 선택된 월 노드 (하이라이트 토글용)
    private Border?      _selectedMonthBorder;

    // 현재 선택된 리스트 행
    private PurchaseItem? _selectedItem;
    private Border?       _selectedRowBorder;

    // 저장 버튼 (신규/수정 모드 전환용)
    private Button _saveBtn = null!;

    // 폼 입력 컨트롤
    private readonly ComboBox _cbCategory;
    private readonly TextBox  _tbItem;
    private readonly TextBox  _tbQty;
    private readonly TextBox  _tbRemark;
    private readonly TextBox  _tbRequester;

    // =========================================================================
    public PurchasePage()
    {
        // ── 날짜 트리 (왼쪽) ─────────────────────────────────────────────────
        _treePanel = new StackPanel { Spacing = 2 };

        var treeScroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Content                       = _treePanel
        };

        TreeControl = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,*"),
            Children =
            {
                // 트리 제목
                new Border
                {
                    Background = AppTheme.BorderSubtle,
                    Padding    = new Thickness(10, 8),
                    Child      = new TextBlock
                    {
                        Text       = "📅  구매 내역",
                        FontFamily = Font,
                        FontSize   = AppTheme.FontMD,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = AppTheme.FgInfo,
                    }
                },
                treeScroll
            }
        };
        Grid.SetRow(treeScroll, 1);

        // ── 리스트 (중앙 상단) ───────────────────────────────────────────────
        _listPanel = new StackPanel { Spacing = 0 };

        var listScroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Content                       = _listPanel
        };

        ListControl = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,*"),
            Children = { BuildListHeader(), listScroll }
        };
        Grid.SetRow(listScroll, 1);

        // ── 신규 추가 폼 (하단) ──────────────────────────────────────────────
        _cbCategory  = new ComboBox
        {
            ItemsSource   = new[] { "소모품", "장비", "시약", "기타" },
            SelectedIndex = 0,
            FontFamily    = Font,
            FontSize      = AppTheme.FontMD,
        };
        _tbItem      = MakeTextBox("품목명 입력",  180);
        _tbQty       = MakeTextBox("수량",          60);
        _tbRemark    = MakeTextBox("비고 (선택)",  180);
        _tbRequester = MakeTextBox("요청자 이름",  120);

        _saveBtn = new Button
        {
            Content         = "➕  요청 추가",
            FontFamily      = Font,
            FontSize        = AppTheme.FontMD,
            Background      = AppTheme.BorderActive,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(14, 6),
            VerticalAlignment = VerticalAlignment.Bottom,
        };
        _saveBtn.Click += SaveBtn_Click;

        FormControl = new Border
        {
            Padding = new Thickness(12),
            Child   = new StackPanel
            {
                Spacing  = 10,
                Children =
                {
                    new TextBlock
                    {
                        Text       = "📋  신규 물품 요청",
                        FontFamily = Font,
                        FontSize   = AppTheme.FontXL,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = AppRes("AppFg"),
                    },
                    new Border { Height = 1, Background = AppTheme.BorderDefault },
                    new WrapPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Children    =
                        {
                            InlineField("구분",   _cbCategory,  110),
                            InlineField("품목",   _tbItem,      180),
                            InlineField("수량",   _tbQty,        60),
                            InlineField("요청자", _tbRequester, 120),
                            InlineField("비고",   _tbRemark,    180),
                            // 버튼도 인라인으로
                            new StackPanel
                            {
                                Orientation   = Orientation.Vertical,
                                Spacing       = 4,
                                Margin        = new Thickness(0, 0, 0, 0),
                                VerticalAlignment = VerticalAlignment.Bottom,
                                Children      = { new TextBlock { Text = " ", FontSize = AppTheme.FontBase }, _saveBtn }
                            }
                        }
                    }
                }
            }
        };

        LoadData();
    }

    // =========================================================================
    // 전체 로드 (트리 + 리스트)
    // =========================================================================
    public void LoadData()
    {
        RefreshTree();

        // 필터가 설정돼 있으면 유지, 없으면 전체 또는 이번 달
        if (_filterYear == 0)
        {
            _filterYear  = DateTime.Today.Year;
            _filterMonth = DateTime.Today.Month;
        }
        LoadListByMonth(_filterYear, _filterMonth);
    }

    // =========================================================================
    // 트리 갱신  연도 > 월 구조
    // =========================================================================
    private void RefreshTree()
    {
        _treePanel.Children.Clear();
        _selectedMonthBorder = null;

        var summary = PurchaseService.GetMonthSummary();
        if (summary.Count == 0)
        {
            _treePanel.Children.Add(new TextBlock
            {
                Text       = "내역 없음",
                FontFamily = Font,
                FontSize   = AppTheme.FontBase,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Thickness(12, 8),
            });
            return;
        }

        // 연도별 그룹
        var byYear = summary.GroupBy(x => x.Year).OrderByDescending(g => g.Key);

        foreach (var yearGroup in byYear)
        {
            int year = yearGroup.Key;

            // ── 연도 노드 ────────────────────────────────────────────────────
            var monthPanel = new StackPanel { Spacing = 1, IsVisible = true };

            var yearHeader = new Border
            {
                Background  = AppTheme.BorderSubtle,
                Padding     = new Thickness(10, 6),
                Cursor      = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child       = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing     = 6,
                    Children    =
                    {
                        new TextBlock
                        {
                            Text       = "📁",
                            FontSize   = AppTheme.FontMD,
                            VerticalAlignment = VerticalAlignment.Center,
                        },
                        new TextBlock
                        {
                            Text       = $"{year}년",
                            FontFamily = Font,
                            FontSize   = AppTheme.FontLG,
                            FontWeight = FontWeight.SemiBold,
                            Foreground = new SolidColorBrush(Color.Parse("#ccccee")),
                            VerticalAlignment = VerticalAlignment.Center,
                        }
                    }
                }
            };

            // 연도 클릭 → 월 목록 펼치기/접기
            yearHeader.PointerPressed += (_, _) =>
                monthPanel.IsVisible = !monthPanel.IsVisible;

            yearHeader.PointerEntered += (_, _) =>
                yearHeader.Background = AppTheme.BorderMuted;
            yearHeader.PointerExited += (_, _) =>
                yearHeader.Background = AppTheme.BorderSubtle;

            _treePanel.Children.Add(yearHeader);
            _treePanel.Children.Add(monthPanel);

            // ── 월 노드 ──────────────────────────────────────────────────────
            foreach (var (_, month, count) in yearGroup.OrderByDescending(x => x.Month))
            {
                int y = year, m = month, c = count;

                var monthBorder = new Border
                {
                    Padding    = new Thickness(28, 5, 10, 5),
                    Background = (y == _filterYear && m == _filterMonth)
                                    ? new SolidColorBrush(Color.Parse("#3a3a60"))
                                    : Brushes.Transparent,
                    Cursor     = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    Child      = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        Spacing     = 6,
                        Children    =
                        {
                            new TextBlock
                            {
                                Text       = $"{m:D2}월",
                                FontFamily = Font,
                                FontSize   = AppTheme.FontMD,
                                Foreground = AppRes("AppFg"),
                                Width      = 36,
                                VerticalAlignment = VerticalAlignment.Center,
                            },
                            new Border
                            {
                                Background   = new SolidColorBrush(Color.Parse("#3a4a5a")),
                                CornerRadius = new CornerRadius(8),
                                Padding      = new Thickness(6, 1),
                                Child        = new TextBlock
                                {
                                    Text       = $"{c}건",
                                    FontFamily = Font,
                                    FontSize   = AppTheme.FontSM,
                                    Foreground = AppTheme.FgInfo,
                                }
                            }
                        }
                    }
                };

                // 선택 하이라이트 초기값
                if (y == _filterYear && m == _filterMonth)
                    _selectedMonthBorder = monthBorder;

                monthBorder.PointerEntered += (_, _) =>
                {
                    if (monthBorder != _selectedMonthBorder)
                        monthBorder.Background = new SolidColorBrush(Color.Parse("#2a2a4a"));
                };
                monthBorder.PointerExited += (_, _) =>
                {
                    if (monthBorder != _selectedMonthBorder)
                        monthBorder.Background = Brushes.Transparent;
                };

                monthBorder.PointerPressed += (_, _) =>
                {
                    // 이전 선택 해제
                    if (_selectedMonthBorder != null)
                        _selectedMonthBorder.Background = Brushes.Transparent;

                    // 새 선택 하이라이트
                    monthBorder.Background = new SolidColorBrush(Color.Parse("#3a3a60"));
                    _selectedMonthBorder   = monthBorder;

                    // 리스트 갱신
                    LoadListByMonth(y, m);
                };

                monthPanel.Children.Add(monthBorder);
            }
        }
    }

    // =========================================================================
    // 월별 리스트 로드
    // =========================================================================
    private void LoadListByMonth(int year, int month)
    {
        _filterYear  = year;
        _filterMonth = month;
        _items       = PurchaseService.GetByMonth(year, month);
        RefreshList();
    }

    // =========================================================================
    // 리스트 갱신
    // =========================================================================
    private void RefreshList()
    {
        _listPanel.Children.Clear();

        // 리스트 상단에 현재 필터 표시
        _listPanel.Children.Add(new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1e2030")),
            Padding    = new Thickness(12, 6),
            Child      = new TextBlock
            {
                Text       = _filterYear > 0
                                 ? $"📋  {_filterYear}년 {_filterMonth:D2}월  —  {_items.Count}건"
                                 : "📋  전체",
                FontFamily = Font,
                FontSize   = AppTheme.FontMD,
                Foreground = new SolidColorBrush(Color.Parse("#8888bb")),
            }
        });

        if (_items.Count == 0)
        {
            _listPanel.Children.Add(new TextBlock
            {
                Text                = "해당 월의 요청 내역이 없습니다.",
                FontFamily          = Font,
                FontSize            = AppTheme.FontMD,
                Foreground          = AppTheme.BorderDefault,
                Margin              = new Thickness(12, 20),
                HorizontalAlignment = HorizontalAlignment.Center
            });
            return;
        }

        bool odd = false;
        int seq = _items.Count;          // 최신(맨 위)이 가장 큰 번호
        foreach (var item in _items)
        {
            _listPanel.Children.Add(BuildDataRow(item, odd, seq--));
            odd = !odd;
        }
    }

    // =========================================================================
    // 리스트 헤더
    // =========================================================================
    private static Border BuildListHeader() => new Border
    {
        Background = AppTheme.BorderSubtle,
        Padding    = new Thickness(8, 6),
        Child      = new Grid
        {
            ColumnDefinitions = ColDefs,
            Children =
            {
                HeaderCell("번호",  0), HeaderCell("구분",  1),
                HeaderCell("품목",  2), HeaderCell("수량",  3),
                HeaderCell("요청자",4), HeaderCell("요청일",5),
                HeaderCell("상태",  6), HeaderCell("비고",  7),
            }
        }
    };

    // =========================================================================
    // 데이터 행
    // =========================================================================
    private Border BuildDataRow(PurchaseItem item, bool odd, int seq)
    {
        var bgColor = odd ? Color.Parse("#2e2e3a") : Color.Parse("#252530");
        var bg      = new SolidColorBrush(bgColor);

        var statusColor = item.상태 switch
        {
            "승인" => Color.Parse("#1a6a1a"),
            "완료" => Color.Parse("#1a3a6a"),
            "반려" => Color.Parse("#6a1a1a"),
            _     => Color.Parse("#4a4a2a")
        };

        var row = new Border
        {
            Background = bg,
            Padding    = new Thickness(8, 5),
            Child      = new Grid
            {
                ColumnDefinitions = ColDefs,
                Children =
                {
                    DataCell($"{seq}",                      0, Brushes.Gray),
                    DataCell(item.구분,                     1),
                    DataCell(item.품목,                     2, AppRes("AppFg"), bold: true),
                    DataCell($"{item.수량}",                3),
                    DataCell(item.요청자,                   4),
                    DataCell(item.요청일.ToString("MM-dd"), 5, AppRes("FgMuted")),
                    BuildStatusBadge(item.상태, statusColor,6),
                    DataCell(item.비고,                     7, AppRes("FgMuted")),
                }
            }
        };

        // ── 이벤트: 터널링(Preview) 방식으로 자식이 소비해도 받을 수 있게 ──
        row.AddHandler(
            Avalonia.Input.InputElement.PointerPressedEvent,
            (object? s, Avalonia.Input.PointerPressedEventArgs ev) =>
            {
                SelectRow(row, item, bg);
                ev.Handled = false;   // 다른 핸들러도 계속 동작
            },
            Avalonia.Interactivity.RoutingStrategies.Tunnel);

        row.PointerEntered += (_, _) =>
        {
            if (row != _selectedRowBorder)
                row.Background = AppTheme.BorderMuted;
        };
        row.PointerExited += (_, _) =>
        {
            if (row != _selectedRowBorder)
                row.Background = bg;
        };
        return row;
    }

    private static Control BuildStatusBadge(string status, Color color, int col)
    {
        var b = new Border
        {
            Background          = new SolidColorBrush(color),
            CornerRadius        = new CornerRadius(10),
            Padding             = new Thickness(8, 2),
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment   = VerticalAlignment.Center,
            Child               = new TextBlock
            {
                Text = status, FontFamily = Font, FontSize = AppTheme.FontSM, Foreground = AppRes("AppFg")
            }
        };
        Grid.SetColumn(b, col);
        return b;
    }

    // =========================================================================
    // 행 선택
    // =========================================================================
    private void SelectRow(Border row, PurchaseItem item, SolidColorBrush normalBg)
    {
        // 이전 선택 해제
        if (_selectedRowBorder != null)
        {
            if (_selectedRowBorder.Tag is SolidColorBrush prev)
                _selectedRowBorder.Background = prev;
        }

        // 새 행 선택
        row.Tag            = normalBg;
        row.Background     = new SolidColorBrush(Color.Parse("#4a4a70"));
        _selectedRowBorder = row;
        _selectedItem      = item;

        // ── 폼에 선택된 항목 데이터 로드 (수정 모드) ──────────────────────
        LoadToForm(item);
    }

    /// <summary>선택된 항목을 폼에 채운다 (수정 모드)</summary>
    private void LoadToForm(PurchaseItem item)
    {
        // 구분 콤보박스 선택
        var items = _cbCategory.ItemsSource as string[];
        if (items != null)
        {
            int idx = Array.IndexOf(items, item.구분);
            _cbCategory.SelectedIndex = idx >= 0 ? idx : 0;
        }

        _tbItem.Text      = item.품목;
        _tbQty.Text       = item.수량.ToString();
        _tbRequester.Text = item.요청자;
        _tbRemark.Text    = item.비고;

        // 저장버튼 텍스트 → 수정 모드 표시
        _saveBtn.Content = "✏️  수정 저장";
        _saveBtn.Background = new SolidColorBrush(Color.Parse("#2a4a5a"));

    }

    /// <summary>폼 초기화 (신규 모드로 전환)</summary>
    private void ClearForm()
    {
        _selectedItem      = null;
        _selectedRowBorder = null;
        _cbCategory.SelectedIndex = 0;
        _tbItem.Text = _tbQty.Text = _tbRequester.Text = _tbRemark.Text = "";
        _saveBtn.Content = "➕  요청 추가";
        _saveBtn.Background = AppTheme.BorderActive;
    }

    // =========================================================================
    // 공개 액션 메서드  (MainPage BT 에서 호출)
    // =========================================================================

    /// BT1 — 새로고침
    public void Refresh()
    {
        _selectedItem      = null;
        _selectedRowBorder = null;
        RefreshTree();
        LoadListByMonth(_filterYear, _filterMonth);
    }

    /// BT2 — 엑셀(CSV) 내보내기
    public void ExportCsv()
    {
        if (_items.Count == 0) return;

        try
        {
            var root    = System.IO.Path.GetFullPath(
                              System.IO.Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
            var dir     = System.IO.Path.Combine(root, "Data", "Export");
            System.IO.Directory.CreateDirectory(dir);
            var file    = System.IO.Path.Combine(dir,
                              $"물품구매_{_filterYear}{_filterMonth:D2}_{DateTime.Now:HHmmss}.csv");

            var lines   = new System.Collections.Generic.List<string>
            {
                "번호,구분,품목,수량,요청자,요청일,상태,비고"
            };
            foreach (var it in _items)
                lines.Add($"{it.Id},{it.구분},{it.품목},{it.수량},{it.요청자}," +
                          $"{it.요청일:yyyy-MM-dd},{it.상태},{it.비고}");

            System.IO.File.WriteAllLines(file, lines, System.Text.Encoding.UTF8);

            // 파일 탐색기에서 열기
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName        = file,
                UseShellExecute = true
            });
        }
        catch (Exception ex) { }
    }

    /// BT3 — 승인
    public void ApproveSelected()  => ChangeStatus("승인");

    /// BT4 — 반려
    public void RejectSelected()   => ChangeStatus("반려");

    /// BT5 — 완료
    public void CompleteSelected() => ChangeStatus("완료");

    /// BT6 — 삭제
    public void DeleteSelected()
    {
        if (_selectedItem == null) return;

        bool ok = PurchaseService.Delete(_selectedItem.Id);

        if (ok)
        {
            _selectedItem      = null;
            _selectedRowBorder = null;
            RefreshTree();
            LoadListByMonth(_filterYear, _filterMonth);
        }
    }

    // ── 상태 변경 공통 ────────────────────────────────────────────────────────
    private void ChangeStatus(string status)
    {
        if (_selectedItem == null)
        {
            return;
        }

        bool ok = PurchaseService.UpdateStatus(_selectedItem.Id, status);

        if (ok)
        {
            _selectedItem.상태 = status;
            _selectedItem      = null;
            _selectedRowBorder = null;
            LoadListByMonth(_filterYear, _filterMonth);
        }
    }

    // =========================================================================
    // 저장 버튼
    // =========================================================================
    private void SaveBtn_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        var 품목 = _tbItem.Text?.Trim() ?? "";
        if (string.IsNullOrEmpty(품목)) { _tbItem.BorderBrush = Brushes.Red; return; }
        _tbItem.BorderBrush = AppTheme.BorderDefault;

        if (!int.TryParse(_tbQty.Text?.Trim(), out int qty) || qty <= 0) qty = 1;

        // ── 수정 모드 ─────────────────────────────────────────────────────
        if (_selectedItem != null)
        {
            _selectedItem.구분   = _cbCategory.SelectedItem?.ToString() ?? "기타";
            _selectedItem.품목   = 품목;
            _selectedItem.수량   = qty;
            _selectedItem.비고   = _tbRemark.Text?.Trim()    ?? "";
            _selectedItem.요청자 = _tbRequester.Text?.Trim() ?? "";

            bool ok = PurchaseService.Update(_selectedItem.Id,
                _selectedItem.구분, _selectedItem.품목, _selectedItem.수량,
                _selectedItem.요청자, _selectedItem.비고);

            if (ok) { ClearForm(); RefreshTree(); LoadListByMonth(_filterYear, _filterMonth); }
            return;
        }

        // ── 신규 모드 ─────────────────────────────────────────────────────
        var item = new PurchaseItem
        {
            구분   = _cbCategory.SelectedItem?.ToString() ?? "기타",
            품목   = 품목,
            수량   = qty,
            비고   = _tbRemark.Text?.Trim()    ?? "",
            요청자 = _tbRequester.Text?.Trim() ?? "",
            요청일 = DateTime.Today,
            상태   = "대기",
        };

        if (!PurchaseService.Insert(item)) return;

        ClearForm();
        RefreshTree();
        LoadListByMonth(item.요청일.Year, item.요청일.Month);
    }

    /// BT7 — 설정 (선택 항목 상세 팝업)
    public void ShowSettings(Avalonia.Controls.Window? owner)
    {
        if (_selectedItem == null)
        {
            // 선택 없으면 현재 필터 정보 팝업
            ShowInfoPopup(owner,
                "ℹ️  현재 필터",
                $"조회 기간 : {_filterYear}년 {_filterMonth:D2}월\n" +
                $"총 건수   : {_items.Count}건\n" +
                $"대기 : {_items.Count(x => x.상태 == "대기")}건  " +
                $"승인 : {_items.Count(x => x.상태 == "승인")}건  " +
                $"완료 : {_items.Count(x => x.상태 == "완료")}건  " +
                $"반려 : {_items.Count(x => x.상태 == "반려")}건");
            return;
        }

        ShowInfoPopup(owner,
            $"📦  {_selectedItem.품목} — 상세",
            $"번호   : {_selectedItem.Id}\n" +
            $"구분   : {_selectedItem.구분}\n" +
            $"품목   : {_selectedItem.품목}\n" +
            $"수량   : {_selectedItem.수량}\n" +
            $"요청자 : {_selectedItem.요청자}\n" +
            $"요청일 : {_selectedItem.요청일:yyyy-MM-dd}\n" +
            $"상태   : {_selectedItem.상태}\n" +
            $"비고   : {_selectedItem.비고}");
    }

    private static void ShowInfoPopup(Avalonia.Controls.Window? owner, string title, string message)
    {
        var dlg = new Avalonia.Controls.Window
        {
            Title   = title,
            Width   = 340,
            Height  = 280,
            CanResize = false,
            WindowStartupLocation = Avalonia.Controls.WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgSecondary,
        };

        var closeBtn = new Avalonia.Controls.Button
        {
            Content         = "닫기",
            Width           = 80,
            HorizontalAlignment = HorizontalAlignment.Right,
            Background      = AppTheme.BorderMuted,
            Foreground      = Brushes.White,
            BorderThickness = new Thickness(0),
        };
        closeBtn.Click += (_, _) => dlg.Close();

        dlg.Content = new Avalonia.Controls.StackPanel
        {
            Margin   = new Thickness(20),
            Spacing  = 12,
            Children =
            {
                new Avalonia.Controls.TextBlock
                {
                    Text         = title,
                    FontFamily   = Font,
                    FontSize     = AppTheme.FontXL,
                    FontWeight   = FontWeight.SemiBold,
                    Foreground   = AppRes("AppFg"),
                },
                new Avalonia.Controls.Border
                {
                    Height     = 1,
                    Background = AppTheme.BorderDefault,
                },
                new Avalonia.Controls.TextBlock
                {
                    Text         = message,
                    FontFamily   = Font,
                    FontSize     = AppTheme.FontMD,
                    Foreground   = AppTheme.FgSecondary,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                closeBtn
            }
        };

        if (owner != null) dlg.ShowDialog(owner);
        else               dlg.Show();
    }

    // =========================================================================
    // 헬퍼
    // =========================================================================
    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    private static ColumnDefinitions ColDefs =>
        new ColumnDefinitions("50,70,180,50,80,70,70,*");

    private static TextBlock HeaderCell(string text, int col)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontBase,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(tb, col); return tb;
    }

    private static TextBlock DataCell(string text, int col, IBrush? fg = null, bool bold = false)
    {
        var tb = new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = AppTheme.FontMD,
            FontWeight   = bold ? FontWeight.SemiBold : FontWeight.Normal,
            Foreground   = fg ?? AppTheme.FgSecondary,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(tb, col); return tb;
    }

    private static StackPanel InlineField(string label, Control input, double w)
    {
        if (input is TextBox tb)  tb.Width  = w;
        if (input is ComboBox cb) cb.Width  = w;
        return new StackPanel
        {
            Orientation = Orientation.Vertical,
            Spacing     = 4,
            Margin      = new Thickness(0, 0, 14, 0),
            Children    =
            {
                new TextBlock { Text = label, FontFamily = Font, FontSize = AppTheme.FontBase,
                                Foreground = AppTheme.FgMuted },
                input
            }
        };
    }

    private static TextBox MakeTextBox(string hint, double w) => new TextBox
    {
        Watermark = hint, Width = w, FontFamily = Font, FontSize = AppTheme.FontMD,
        Background      = AppTheme.BorderSeparator,
        Foreground      = AppRes("AppFg"),
        BorderThickness = new Thickness(1),
        BorderBrush     = AppTheme.BorderDefault,
        CornerRadius    = new CornerRadius(4),
        Padding         = new Thickness(8, 4),
    };
}