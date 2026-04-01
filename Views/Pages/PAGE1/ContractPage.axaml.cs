using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using Avalonia.VisualTree;

namespace ETA.Views.Pages.PAGE1;

/// <summary>약칭이 있을 때만 뱃지 표시용</summary>



public partial class ContractPage : UserControl
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;
    public event Action<Control?>? PricePanelChanged;  // Show3 (현재 미사용)

    // ── 상태 ────────────────────────────────────────────────────────────────
    private Contract?   _selectedContract;
    private StackPanel? _detailPanel;        // 내부 StackPanel (SyncPanelToContract용)
    private bool        _isAddMode = false;
    private bool        _showActive = true;   // true=현행계약, false=종료계약
    public MainPage? ParentMainPage { get; set; }

    // 단가 in-memory 편집본 (서브메뉴 저장 시 DB 반영)
    private readonly Dictionary<string, string>    _pendingPrices       = new();
    // Show2 단가 표시용 TextBlock 참조 (항목명 → 표시 TextBlock)
    private readonly Dictionary<string, TextBlock> _priceDisplayBlocks  = new();

    // 계약구분 ComboBox 참조 (저장 시 값 읽기용)
    private ComboBox? _contractTypeComboBox;

    // 측정인 관리 패널 상태
    private string?    _selectedMeasCompany;
    private TextBox?   _txbMeasAlias;
    private TextBox?   _txbMeasAmount;
    private ComboBox?  _cmbMeasQuotType;
    private StackPanel? _measEditForm;

    public ContractPage()
    {
        InitializeComponent();
    }

    // =========================================================================
    // 데이터 로드
    // =========================================================================
    public void LoadData()
    {
        Log("LoadData() 시작");
        ContractTreeView.Items.Clear();
        _selectedContract = null;
        _isAddMode        = false;
        DetailPanelChanged?.Invoke(null);
        PricePanelChanged?.Invoke(null);

        try
        {
            var today = DateTime.Today;
            var items = ContractService.GetAllContracts()
                            .Where(c =>
                            {
                                bool active = c.C_ContractStart.HasValue && c.C_ContractEnd.HasValue
                                              && c.C_ContractStart.Value.Date <= today
                                              && c.C_ContractEnd.Value.Date   >= today;
                                return _showActive ? active : !active;
                            })
                            .OrderBy(c => c.C_CompanyName)
                            .ToList();

            foreach (var item in items)
                ContractTreeView.Items.Add(CreateTreeItem(item));

            Log($"로드 완료 → {items.Count}건 ({(_showActive ? "현행" : "종료")})");
        }
        catch (Exception ex) { Log("★ 크래시 ★ " + ex.Message); }
    }

    private void TglContractStatus_Changed(object? sender, RoutedEventArgs e)
    {
        _showActive = tglContractStatus.IsChecked == true;
        LoadData();
    }

    // =========================================================================
    // TreeViewItem 생성
    // =========================================================================
    private static TreeViewItem CreateTreeItem(Contract contract)
    {
        string icon = contract.C_ContractType switch
        {
            "위탁" => "🤝",
            "용역" => "📋",
            "구매" => "🛒",
            _     => "🏢"
        };

        return new TreeViewItem
        {
            Tag    = contract,
            Header = new StackPanel
            {
                Orientation       = Orientation.Horizontal,
                Spacing           = 8,
                VerticalAlignment = VerticalAlignment.Center,
                Children =
                {
                    new TextBlock
                    {
                        Text     = icon,
                        FontSize = 18,
                        VerticalAlignment = VerticalAlignment.Center
                    },
                    new StackPanel
                    {
                        Orientation = Orientation.Vertical,
                        Spacing     = 1,
                        VerticalAlignment = VerticalAlignment.Center,
                        Children =
                        {
                            new TextBlock
                            {
                                Text       = contract.C_CompanyName,
                                FontSize   = 13,
                                FontFamily = Font,
                                Foreground = AppRes("AppFg"),
                            },
                            new StackPanel
                            {
                                Orientation = Orientation.Horizontal, Spacing = 4,
                                Children =
                                {
                                    string.IsNullOrEmpty(contract.C_Abbreviation)
                                        ? (Control)new TextBlock()
                                        : new Border
                                        {
                                            Background   = Brush.Parse(BadgeColorHelper.GetBadgeColor(contract.C_Abbreviation).Bg),
                                            CornerRadius = new CornerRadius(3),
                                            Padding      = new Thickness(4, 1),
                                            VerticalAlignment = VerticalAlignment.Center,
                                            Child = new TextBlock
                                            {
                                                Text       = contract.C_Abbreviation,
                                                FontSize   = 9, FontFamily = Font,
                                                Foreground = Brush.Parse(BadgeColorHelper.GetBadgeColor(contract.C_Abbreviation).Fg),
                                            }
                                        },
                                    new TextBlock
                                    {
                                        Text       = contract.C_ContractType,
                                        FontSize   = 10, FontFamily = Font,
                                        Foreground = new SolidColorBrush(Color.Parse("#888888")),
                                        VerticalAlignment = VerticalAlignment.Center,
                                    }
                                }
                            }
                        }
                    }
                }
            }
        };
    }

    // =========================================================================
    // 트리 선택 → 수정 패널
    // =========================================================================
    public void ContractTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        Contract? contract = null;

        if (e.AddedItems.Count > 0)
        {
            if (e.AddedItems[0] is TreeViewItem tvi && tvi.Tag is Contract c1) contract = c1;
            else if (e.AddedItems[0] is Contract c2) contract = c2;
        }

        if (contract == null) return;

        _selectedContract = contract;
        _isAddMode        = false;
        _pendingPrices.Clear();
        _priceDisplayBlocks.Clear();
        PricePanelChanged?.Invoke(null);

        // ① 계약정보 패널 즉시 표시 (DB 없음 — 빠름)
        var (root, scroll, priceContainer) = BuildInfoPanel(contract);
        _detailPanel = root;
        DetailPanelChanged?.Invoke(scroll);

        // ② 단가 데이터 백그라운드 로드 후 UI append
        var captured = contract;
        _ = Task.Run(() => ContractService.GetContractPrices(captured.C_CompanyName))
                .ContinueWith(t =>
                {
                    if (t.Exception != null) { Log("단가 로드 오류: " + t.Exception.Message); return; }
                    var prices = t.Result;
                    Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                    {
                        if (_selectedContract?.C_CompanyName != captured.C_CompanyName) return;
                        AppendPriceTable(priceContainer, captured.C_CompanyName, prices);
                    });
                }, TaskContinuationOptions.None);

        Log($"선택: {contract.C_CompanyName}");
    }

    // =========================================================================
    // 계약업체 추가 패널  (MainPage BT3)
    // =========================================================================
    public void ShowAddPanel()
    {
        _selectedContract              = null;
        _isAddMode                     = true;
        _pendingPrices.Clear();
        _priceDisplayBlocks.Clear();
        ContractTreeView.SelectedItem  = null;
        _detailPanel                   = BuildAddPanel();
        DetailPanelChanged?.Invoke(_detailPanel);
        PricePanelChanged?.Invoke(null);
        Log("추가 모드");
    }

    // =========================================================================
    // 저장  (MainPage BT1)
    // =========================================================================
    public void SaveSelected()
    {
        if (_isAddMode) SaveAdd();
        else
        {
            SaveEdit();
            SavePendingPrices();
        }
    }

    // =========================================================================
    // 삭제  (MainPage BT4)
    // =========================================================================
    public async Task DeleteSelectedAsync()
    {
        if (_selectedContract == null) { Log("삭제 스킵: 선택 없음"); return; }

        var confirmed = await ShowConfirmDialog(
            $"'{_selectedContract.C_CompanyName}' 계약업체를 삭제하시겠습니까?");
        if (!confirmed) return;

        bool ok = ContractService.Delete(_selectedContract);
        Log(ok ? $"✅ 삭제: {_selectedContract.C_CompanyName}"
               : $"❌ 삭제 실패: {_selectedContract.C_CompanyName}");

        if (ok)
        {
            var toRemove = ContractTreeView.Items
                .OfType<TreeViewItem>()
                .FirstOrDefault(i => i.Tag == _selectedContract);
            if (toRemove != null) ContractTreeView.Items.Remove(toRemove);

            _selectedContract = null;
            _detailPanel      = null;
            DetailPanelChanged?.Invoke(null);
        }
    }

    // 단가 접근 허용 사용자
    private static readonly HashSet<string> _priceAllowedUsers =
        new(StringComparer.OrdinalIgnoreCase) { "정승욱", "박은지", "방찬미" };

    private static string FormatPrice(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "—";
        if (decimal.TryParse(raw, out var d)) return ((long)Math.Round(d)).ToString("N0");
        return raw;
    }

    // =========================================================================
    // 패널 — 수정 모드  (계약정보만, 단가는 백그라운드 로드)
    // =========================================================================
    /// <returns>(innerStackPanel, scrollWrapper, priceContainer)</returns>
    private (StackPanel root, ScrollViewer scroll, StackPanel priceContainer) BuildInfoPanel(Contract c)
    {
        var root = MakeRootPanel($"🏢  {c.C_CompanyName} — 계약 정보");
        var grid = MakeTwoColumnGrid();

        AddGridRow(grid, 0, "업체명",           c.C_CompanyName,      isReadOnly: true, isLocked: true);
        AddGridRow(grid, 1, "약칭",             c.C_Abbreviation);
        AddGridRowComboBox(grid, 2, "계약구분", c.C_ContractType);
        // 계약일수: C_ContractEnd - 오늘 (남은 일수)
        string remainDays = c.C_ContractEnd.HasValue
            ? (c.C_ContractEnd.Value.Date - DateTime.Today).Days.ToString()
            : "";

        // 계약금액: decimal 직접 포맷 (null 안전)
        string amountStr = c.C_ContractAmountVATExcluded.HasValue
            ? c.C_ContractAmountVATExcluded.Value.ToString("N0") + " 원"
            : "";

        AddGridRow(grid, 3, "계약시작",         c.C_ContractStartStr, isReadOnly: true);
        AddGridRow(grid, 4, "계약종료",         c.C_ContractEndStr,   isReadOnly: true);
        AddGridRow(grid, 5, "잔여일수",         remainDays,           isReadOnly: true);
        AddGridRow(grid, 6, "계약금액(VAT별도)", amountStr);
        AddGridRow(grid, 7, "주소",             c.C_Address);
        AddGridRow(grid, 8, "대표자",           c.C_Representative);
        AddGridRow(grid, 9, "시설별",           c.C_FacilityType);
        AddGridRow(grid,10, "종류별",           c.C_CategoryType);
        AddGridRow(grid,11, "주생산품",         c.C_MainProduct);
        AddGridRow(grid,12, "담당자",           c.C_ContactPerson);
        AddGridRow(grid,13, "연락처",           c.C_PhoneNumber);
        AddGridRow(grid,14, "이메일",           c.C_Email);
        root.Children.Add(grid);

        // ── 단가 섹션 컨테이너 (백그라운드 로드 후 채워짐) ───────────────────
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = new SolidColorBrush(Color.Parse("#444444")),
            Margin     = new Thickness(0, 16, 0, 8),
        });

        var currentUser = ETA.Services.Common.CurrentUserManager.Instance.CurrentUserId;
        var priceContainer = new StackPanel { Spacing = 4 };

        if (_priceAllowedUsers.Contains(currentUser))
        {
            root.Children.Add(new TextBlock
            {
                Text       = $"📊  분석 단가 — {c.C_CompanyName}",
                FontSize   = 13, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = new SolidColorBrush(Color.Parse("#8899bb")),
                Margin     = new Thickness(0, 0, 0, 4),
            });
            root.Children.Add(new TextBlock
            {
                Text       = "항목 클릭 → 오른쪽에서 편집  |  서브메뉴 [저장] → 서버 반영",
                FontSize   = 10, FontFamily = Font,
                Foreground = new SolidColorBrush(Color.Parse("#666666")),
                Margin     = new Thickness(0, 0, 0, 6),
            });
            // 로딩 표시
            priceContainer.Children.Add(new TextBlock
            {
                Text       = "⏳ 단가 로드 중...",
                FontSize   = 10, FontFamily = Font,
                Foreground = new SolidColorBrush(Color.Parse("#666666")),
            });
        }
        else
        {
            root.Children.Add(new TextBlock
            {
                Text       = "🔒  단가 정보 — 접근 권한 없음",
                FontSize   = 12, FontFamily = Font,
                Foreground = new SolidColorBrush(Color.Parse("#775555")),
            });
        }

        root.Children.Add(priceContainer);

        var scroll = new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        return (root, scroll, priceContainer);
    }

    /// <summary>백그라운드에서 받은 단가 데이터를 priceContainer에 렌더링</summary>
    private void AppendPriceTable(StackPanel priceContainer, string companyName,
                                  List<(string Analyte, string Price)> prices)
    {
        priceContainer.Children.Clear();

        if (prices.Count == 0)
        {
            priceContainer.Children.Add(new TextBlock
            {
                Text       = "단가 정보 없음 — Excel 가져오기로 불러오세요.",
                FontSize   = 11, FontFamily = Font,
                Foreground = new SolidColorBrush(Color.Parse("#888888")),
            });
            return;
        }

        const int COLS = 4;
        var priceGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,*,*,*"),
            ColumnSpacing     = 6,
            RowSpacing        = 3,
        };
        int rowCount = (prices.Count + COLS - 1) / COLS;
        for (int i = 0; i < rowCount; i++)
            priceGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        for (int i = 0; i < prices.Count; i++)
        {
            var (analyte, price) = prices[i];
            int col = i % COLS;
            int row = i / COLS;

            var priceBlock = new TextBlock
            {
                Text              = FormatPrice(price),
                Width             = 60, FontSize = 10, FontFamily = Font,
                Foreground        = string.IsNullOrEmpty(price)
                                        ? new SolidColorBrush(Color.Parse("#555555"))
                                        : new SolidColorBrush(Color.Parse("#aaddaa")),
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment     = Avalonia.Media.TextAlignment.Right,
            };
            _priceDisplayBlocks[analyte] = priceBlock;

            var cell = new Border
            {
                Background      = new SolidColorBrush(Color.Parse("#1e1e2e")),
                CornerRadius    = new CornerRadius(3),
                Padding         = new Thickness(6, 3),
                Cursor          = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child           = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing     = 4,
                    Children    =
                    {
                        new TextBlock
                        {
                            Text              = analyte,
                            FontSize          = 9, FontFamily = Font,
                            Foreground        = AppRes("FgMuted"),
                            VerticalAlignment = VerticalAlignment.Center,
                            TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
                            Width             = 110,
                        },
                        priceBlock,
                    }
                }
            };

            var capturedAnalyte = analyte;
            cell.PointerPressed += (_, _) => ShowPriceEditor(companyName, capturedAnalyte);

            Grid.SetColumn(cell, col);
            Grid.SetRow(cell, row);
            priceGrid.Children.Add(cell);
        }

        priceContainer.Children.Add(priceGrid);
    }

    // =========================================================================
    // 패널 — 추가 모드 (업체명 입력 가능)
    // =========================================================================
    private StackPanel BuildAddPanel()
    {
        var root = MakeRootPanel("➕  신규 계약업체 추가");
        var grid = MakeTwoColumnGrid();

        AddGridRow(grid, 0, "업체명",           "", hint: "업체명 입력 (필수)");
        AddGridRow(grid, 1, "약칭",             "");

        // ★ 계약구분 → ComboBox (row 2)
        AddGridRowComboBox(grid, 2, "계약구분", "");

        AddGridRow(grid, 3, "계약시작",         "", hint: "예) 2024-01-01");
        AddGridRow(grid, 4, "계약종료",         "", hint: "예) 2024-12-31");
        AddGridRow(grid, 5, "계약일수",         "");
        AddGridRow(grid, 6, "계약금액(VAT별도)", "");
        AddGridRow(grid, 7, "주소",             "");
        AddGridRow(grid, 8, "대표자",           "");
        AddGridRow(grid, 9, "시설별",           "");
        AddGridRow(grid,10, "종류별",           "");
        AddGridRow(grid,11, "주생산품",         "");
        AddGridRow(grid,12, "담당자",           "");
        AddGridRow(grid,13, "연락처",           "");
        AddGridRow(grid,14, "이메일",           "");

        root.Children.Add(grid);
        return root;
    }

    // =========================================================================
    // 수정 저장
    // =========================================================================
    private void SaveEdit()
    {
        if (_selectedContract == null || _detailPanel == null)
        {
            Log("저장 스킵: 선택 없음"); return;
        }

        SyncPanelToContract(_detailPanel, _selectedContract, includeReadOnly: false);

        bool ok = ContractService.Update(_selectedContract);
        Log(ok ? $"✅ 수정 저장: {_selectedContract.C_CompanyName}"
               : $"❌ 수정 실패: {_selectedContract.C_CompanyName}");
    }

    // =========================================================================
    // 추가 저장
    // =========================================================================
    private void SaveAdd()
    {
        if (_detailPanel == null) return;

        var newContract = new Contract();
        SyncPanelToContract(_detailPanel, newContract, includeReadOnly: true);

        if (string.IsNullOrWhiteSpace(newContract.C_CompanyName))
        {
            Log("❌ 업체명 없음 → 추가 취소"); return;
        }

        newContract.OriginalCompanyName = newContract.C_CompanyName;

        bool ok = ContractService.Insert(newContract);
        Log(ok ? $"✅ 추가: {newContract.C_CompanyName}"
               : $"❌ 추가 실패: {newContract.C_CompanyName}");

        if (ok)
        {
            ContractTreeView.Items.Add(CreateTreeItem(newContract));
            _isAddMode   = false;
            _detailPanel = null;
            DetailPanelChanged?.Invoke(null);
        }
    }

    // =========================================================================
    // UI → Contract 동기화
    // =========================================================================
    private void SyncPanelToContract(StackPanel panel, Contract c, bool includeReadOnly)
    {
        var grid = panel.Children.OfType<Grid>().FirstOrDefault();
        if (grid == null) return;

        foreach (var child in grid.Children.OfType<StackPanel>())
        {
            if (child.Children.Count < 2) continue;

            var labelBlock = child.Children[0] as TextBlock;
            var label = labelBlock?.Text ?? "";
            label = label.Replace("🔒 ", "").Replace("    ", "").Replace(" :", "").Trim();

            // ★ ComboBox 처리 (계약구분)
            if (child.Children[1] is ComboBox cb)
            {
                if (label == "계약구분")
                    c.C_ContractType = cb.SelectedItem?.ToString() ?? "";
                continue;
            }

            // TextBox 처리
            var tb = child.Children[1] as TextBox;
            if (tb == null) continue;
            if (tb.IsReadOnly && !includeReadOnly) continue;

            switch (label)
            {
                case "업체명":           c.C_CompanyName    = tb.Text ?? ""; break;
                case "약칭":             c.C_Abbreviation   = tb.Text ?? ""; break;
                case "주소":             c.C_Address        = tb.Text ?? ""; break;
                case "대표자":           c.C_Representative = tb.Text ?? ""; break;
                case "시설별":           c.C_FacilityType   = tb.Text ?? ""; break;
                case "종류별":           c.C_CategoryType   = tb.Text ?? ""; break;
                case "주생산품":         c.C_MainProduct    = tb.Text ?? ""; break;
                case "담당자":           c.C_ContactPerson  = tb.Text ?? ""; break;
                case "연락처":           c.C_PhoneNumber    = tb.Text ?? ""; break;
                case "이메일":           c.C_Email          = tb.Text ?? ""; break;
                case "계약일수":
                    if (int.TryParse(tb.Text, out int days)) c.C_ContractDays = days;
                    break;
                case "계약금액(VAT별도)":
                    var amtStr = tb.Text?.Replace(",", "").Replace("원", "").Trim() ?? "";
                    if (decimal.TryParse(amtStr, out decimal amt))
                        c.C_ContractAmountVATExcluded = amt;
                    break;
                case "계약시작":
                    if (DateTime.TryParse(tb.Text, out var s)) c.C_ContractStart = s;
                    break;
                case "계약종료":
                    if (DateTime.TryParse(tb.Text, out var e2)) c.C_ContractEnd = e2;
                    break;
            }
        }
    }

    // =========================================================================
    // UI 헬퍼
    // =========================================================================
    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 R");

    private static StackPanel MakeRootPanel(string title)
    {
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(4) };
        root.Children.Add(new TextBlock
        {
            Text       = title,
            FontSize   = 15,
            FontFamily = Font,
            FontWeight = FontWeight.SemiBold,
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

    private static Grid MakeTwoColumnGrid()
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,*"),
            ColumnSpacing     = 16,
            RowSpacing        = 8,
        };
        for (int i = 0; i < 15; i++)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        return grid;
    }

    private static void AddGridRow(Grid grid, int row, string label, string value,
                                   bool isReadOnly = false, bool isLocked = false,
                                   string hint = "")
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

        panel.Children.Add(new TextBlock
        {
            Text              = (isLocked ? "🔒 " : "    ") + label + " :",
            Width             = 130,
            FontSize          = 12,
            FontFamily        = Font,
            Foreground        = isLocked
                                    ? new SolidColorBrush(Color.Parse("#888888"))
                                    : AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        panel.Children.Add(new TextBox
        {
            Text            = value ?? "",
            Width           = 200,
            FontSize        = 12,
            FontFamily      = Font,
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
            Padding         = new Thickness(8, 4),
        });

        int col     = row % 2;
        int gridRow = row / 2;
        Grid.SetColumn(panel, col);
        Grid.SetRow(panel, gridRow);
        grid.Children.Add(panel);
    }

    // =========================================================================
    // ★ 계약구분 ComboBox 행 추가
    //   - 분석단가 테이블의 FS100 이후 컬럼을 항목으로 채움
    //   - 기본 선택값 : FS25 (없으면 첫 번째 항목)
    //   - 선택 변경 시 : Show3 에 해당 컬럼 데이터 테이블 표시
    // =========================================================================
    private void AddGridRowComboBox(Grid grid, int row, string label, string currentValue)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

        panel.Children.Add(new TextBlock
        {
            Text              = "    " + label + " :",
            Width             = 130,
            FontSize          = 12,
            FontFamily        = Font,
            Foreground        = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });

        // 계약구분 목록 (계약 DB C_ContractType 고유값, 없으면 기본값)
        List<string> columns;
        try
        {
            columns = QuotationService.GetContractTypes();
        }
        catch
        {
            columns = new List<string> { "위탁", "용역", "구매", "기타" };
        }

        var comboBox = new ComboBox
        {
            Width           = 200,
            FontSize        = 12,
            FontFamily      = Font,
            Background      = new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555577")),
            Padding         = new Thickness(8, 4),
        };

        foreach (var col in columns)
            comboBox.Items.Add(col);

        // 기본 선택값 결정: currentValue 우선, 없으면 FS25, 그것도 없으면 첫 번째
        if (!string.IsNullOrEmpty(currentValue) && columns.Contains(currentValue))
            comboBox.SelectedItem = currentValue;
        else if (columns.Contains("FS25"))
            comboBox.SelectedItem = "FS25";
        else if (columns.Count > 0)
            comboBox.SelectedIndex = 0;

        // 필드 저장용 참조 보관
        _contractTypeComboBox = comboBox;

        panel.Children.Add(comboBox);

        int col2    = row % 2;
        int gridRow = row / 2;
        Grid.SetColumn(panel, col2);
        Grid.SetRow(panel, gridRow);
        grid.Children.Add(panel);

    }

    // =========================================================================
    // 확인 다이얼로그
    // =========================================================================
    private async Task<bool> ShowConfirmDialog(string message)
    {
        var dlg = new Window
        {
            Title   = "삭제 확인",
            Width   = 340, Height = 150,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.Parse("#2d2d2d")),
        };

        bool result = false;
        var yes = new Button { Content = "삭제", Width = 80,
                               Background = new SolidColorBrush(Color.Parse("#c0392b")),
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };
        var no  = new Button { Content = "취소", Width = 80,
                               Background = new SolidColorBrush(Color.Parse("#444")),
                               Foreground = Brushes.White, BorderThickness = new Thickness(0) };

        yes.Click += (_, _) => { result = true;  dlg.Close(); };
        no.Click  += (_, _) => { result = false; dlg.Close(); };

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock { Text = message, Foreground = AppRes("AppFg"),
                                FontSize = 13, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12,
                                 HorizontalAlignment = HorizontalAlignment.Right,
                                 Children = { yes, no } }
            }
        };

        var owner = TopLevel.GetTopLevel(this) as Window;
        if (owner != null) await dlg.ShowDialog(owner);
        else dlg.Show();
        return result;
    }

    // =========================================================================
    // 측정인 업체 관리 패널 (ActiveContentPage1 에 표시)
    // =========================================================================
    public void ShowMeasurerPanel()
    {
        _selectedContract = null;
        _isAddMode        = false;
        ContractTreeView.SelectedItem = null;
        DetailPanelChanged?.Invoke(BuildMeasurerPanel());
    }

    private StackPanel BuildMeasurerPanel()
    {
        var root = MakeRootPanel("📍  측정인 업체 관리");

        // 업체 목록
        var companies = MeasurerService.GetCompanies();
        if (companies.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "측정인 DB에 업체가 없습니다.\n먼저 로그인하여 데이터를 수집하세요.",
                FontSize = 11, FontFamily = Font,
                Foreground = Brush.Parse("#f0c040"),
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            });
            return root;
        }

        root.Children.Add(new TextBlock
        {
            Text = $"업체 {companies.Count}개 — 선택하여 정보를 편집하세요.",
            FontSize = 10, FontFamily = Font, Foreground = Brush.Parse("#888"),
            Margin = new Thickness(0, 0, 0, 4),
        });

        var listBox = new ListBox
        {
            MaxHeight = 220,
            Background = Brush.Parse("#13131f"),
            BorderBrush = Brush.Parse("#333"), BorderThickness = new Thickness(1),
        };
        foreach (var c in companies)
        {
            var info = MeasurerService.GetCompanyInfo(c);
            var item = new ListBoxItem
            {
                Tag = c,
                Content = new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 8,
                    Children =
                    {
                        new TextBlock { Text = c, FontSize = 11, FontFamily = Font,
                                        Foreground = Brush.Parse("#ddd") },
                        string.IsNullOrEmpty(info.약칭)
                            ? (Control)new TextBlock()
                            : new Border
                            {
                                Background = Brush.Parse("#1a3a5a"), CornerRadius = new CornerRadius(3),
                                Padding = new Thickness(4, 1),
                                Child = new TextBlock { Text = info.약칭, FontSize = 9,
                                                        FontFamily = Font, Foreground = Brush.Parse("#88aadd") }
                            }
                    }
                }
            };
            listBox.Items.Add(item);
        }

        _measEditForm = new StackPanel { Spacing = 8, Margin = new Thickness(0, 10, 0, 0) };

        listBox.SelectionChanged += (_, _) =>
        {
            if (listBox.SelectedItem is ListBoxItem li && li.Tag is string company)
                RefreshMeasurerEditForm(company);
        };

        root.Children.Add(listBox);
        root.Children.Add(_measEditForm);
        return root;
    }

    private void RefreshMeasurerEditForm(string company)
    {
        _selectedMeasCompany = company;
        _measEditForm!.Children.Clear();

        var (alias, amount, quotType) = MeasurerService.GetCompanyInfo(company);

        _measEditForm.Children.Add(new Border
        {
            Height = 1, Background = Brush.Parse("#444"), Margin = new Thickness(0, 0, 0, 4)
        });
        _measEditForm.Children.Add(new TextBlock
        {
            Text = company, FontSize = 12, FontFamily = Font,
            Foreground = Brush.Parse("#aaccff"), FontWeight = FontWeight.SemiBold,
        });

        // 약칭
        var aliasRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        aliasRow.Children.Add(MakeLabel("약칭"));
        _txbMeasAlias = MakeTextBox(alias, "약칭 입력");
        aliasRow.Children.Add(_txbMeasAlias);
        _measEditForm.Children.Add(aliasRow);

        // 추천 약칭 버튼
        var suggestions = SuggestAliases(company);
        if (suggestions.Count > 0)
        {
            var sugRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            sugRow.Children.Add(new TextBlock
            {
                Text = "추천:", FontSize = 9, FontFamily = Font,
                Foreground = Brush.Parse("#666"), VerticalAlignment = VerticalAlignment.Center,
            });
            foreach (var sug in suggestions)
            {
                var s = sug;
                var btn = new Button
                {
                    Content = s, Height = 20, FontSize = 9, FontFamily = Font,
                    Background = Brush.Parse("#2a2a4a"), Foreground = Brush.Parse("#99aacc"),
                    BorderThickness = new Thickness(1), BorderBrush = Brush.Parse("#444"),
                    CornerRadius = new CornerRadius(3), Padding = new Thickness(6, 0),
                };
                btn.Click += (_, _) => { if (_txbMeasAlias != null) _txbMeasAlias.Text = s; };
                sugRow.Children.Add(btn);
            }
            _measEditForm.Children.Add(sugRow);
        }

        // 계약금액
        var amtRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        amtRow.Children.Add(MakeLabel("계약금액(원)"));
        _txbMeasAmount = MakeTextBox(amount, "예) 12000000");
        amtRow.Children.Add(_txbMeasAmount);
        _measEditForm.Children.Add(amtRow);

        // 견적구분
        var qtRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
        qtRow.Children.Add(MakeLabel("견적구분"));
        _cmbMeasQuotType = new ComboBox
        {
            Width = 200, FontSize = 12, FontFamily = Font,
            Background = Brush.Parse("#3a3a4a"), Foreground = AppRes("AppFg"),
            BorderBrush = Brush.Parse("#555577"), BorderThickness = new Thickness(1),
            Padding = new Thickness(8, 4),
        };
        try
        {
            foreach (var col in QuotationService.GetContractTypes())
                _cmbMeasQuotType.Items.Add(col);
            if (!string.IsNullOrEmpty(quotType) && _cmbMeasQuotType.Items.Contains(quotType))
                _cmbMeasQuotType.SelectedItem = quotType;
            else if (_cmbMeasQuotType.Items.Count > 0)
                _cmbMeasQuotType.SelectedIndex = 0;
        }
        catch { }
        qtRow.Children.Add(_cmbMeasQuotType);
        _measEditForm.Children.Add(qtRow);

        // 저장 버튼
        var btnSave = new Button
        {
            Content = "💾  저장", Height = 30, FontSize = 12, FontFamily = Font,
            Background = Brush.Parse("#1a3a2a"), Foreground = Brush.Parse("#88ee88"),
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding = new Thickness(16, 0), Margin = new Thickness(0, 6, 0, 0),
        };
        btnSave.Click += (_, _) =>
        {
            if (_selectedMeasCompany == null) return;
            MeasurerService.UpdateCompanyInfo(
                _selectedMeasCompany,
                _txbMeasAlias?.Text?.Trim()   ?? "",
                _txbMeasAmount?.Text?.Trim()   ?? "",
                _cmbMeasQuotType?.SelectedItem?.ToString() ?? "");
            Log($"측정인 업체 저장: {_selectedMeasCompany}");
        };
        _measEditForm.Children.Add(btnSave);
    }

    // 약칭 추천 생성
    private static List<string> SuggestAliases(string companyName)
    {
        var name = Regex.Replace(companyName, @"\(주\)|\(유\)|주식회사|유한회사", "").Trim();
        int parenIdx = name.IndexOf('(');
        string main = parenIdx > 0 ? name[..parenIdx].Trim() : name.Trim();
        string sub  = "";
        if (parenIdx > 0)
        {
            int close = name.IndexOf(')', parenIdx);
            if (close > parenIdx) sub = name[(parenIdx + 1)..close].Trim();
        }
        // 공백 뒤 지명 제거
        int spaceIdx = main.IndexOf(' ');
        string mainCore = spaceIdx > 0 ? main[..spaceIdx].Trim() : main;

        var set = new List<string>();
        if (mainCore.Length >= 2) set.Add(mainCore[..Math.Min(4, mainCore.Length)]);
        if (mainCore.Length >= 3) set.Add(mainCore[..Math.Min(3, mainCore.Length)]);
        if (!string.IsNullOrEmpty(sub) && mainCore.Length >= 2)
            set.Add(mainCore[..Math.Min(2, mainCore.Length)] + sub[..Math.Min(2, sub.Length)]);

        return set.Distinct().Where(s => s.Length >= 2).ToList();
    }

    private static TextBlock MakeLabel(string text) => new()
    {
        Text = text + " :", Width = 110, FontSize = 11, FontFamily = Font,
        Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
    };

    private static TextBox MakeTextBox(string value, string hint = "") => new()
    {
        Text = value, Width = 200, FontSize = 12, FontFamily = Font, Watermark = hint,
        Background = Brush.Parse("#3a3a4a"), Foreground = AppRes("AppFg"),
        BorderBrush = Brush.Parse("#555577"), BorderThickness = new Thickness(1),
        CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 4),
    };

    // =========================================================================
    // Show3 — 단가 항목 편집 폼 (단일 항목)
    // =========================================================================

    /// <summary>Show2 단가 테이블에서 항목 클릭 시 Show3에 편집 폼 표시</summary>
    private void ShowPriceEditor(string companyName, string analyte)
    {
        // 현재 값: pending이 있으면 pending 값, 없으면 표시 블록의 텍스트
        string currentPrice = _pendingPrices.TryGetValue(analyte, out var p)
                                ? p
                                : (_priceDisplayBlocks.TryGetValue(analyte, out var tb)
                                    ? (tb.Text == "—" ? "" : tb.Text ?? "")
                                    : "");

        var root = new StackPanel { Spacing = 12, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text       = "✏️  단가 편집",
            FontSize   = 13, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = new SolidColorBrush(Color.Parse("#8899bb")),
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = new SolidColorBrush(Color.Parse("#444444")),
        });

        // 항목명
        root.Children.Add(new TextBlock
        {
            Text       = analyte,
            FontSize   = 12, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg"),
        });

        // 단가 입력
        var priceBox = new TextBox
        {
            Text            = currentPrice,
            FontSize        = 14, FontFamily = Font,
            Watermark       = "단가 입력 (원)",
            Background      = new SolidColorBrush(Color.Parse("#3a3a4a")),
            Foreground      = AppRes("AppFg"),
            BorderBrush     = new SolidColorBrush(Color.Parse("#555577")),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(10, 6),
        };
        root.Children.Add(priceBox);

        // 적용 버튼 (Show2 표 업데이트 + pending 기록)
        var btnApply = new Button
        {
            Content         = "✅  적용 (Show2 반영)",
            Height          = 34, FontSize = 12, FontFamily = Font,
            Background      = new SolidColorBrush(Color.Parse("#1a3a5a")),
            Foreground      = new SolidColorBrush(Color.Parse("#88aaee")),
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(16, 0),
        };
        btnApply.Click += (_, _) =>
        {
            var newPrice = priceBox.Text?.Trim() ?? "";
            _pendingPrices[analyte] = newPrice;

            // Show2 단가 표시 즉시 업데이트
            if (_priceDisplayBlocks.TryGetValue(analyte, out var displayBlock))
            {
                displayBlock.Text       = string.IsNullOrEmpty(newPrice) ? "—" : newPrice;
                displayBlock.Foreground = string.IsNullOrEmpty(newPrice)
                                            ? new SolidColorBrush(Color.Parse("#555555"))
                                            : new SolidColorBrush(Color.Parse("#ffdd88")); // 변경됨 표시
            }
            Log($"단가 적용(pending): {analyte} = {newPrice}");
        };
        root.Children.Add(btnApply);

        root.Children.Add(new TextBlock
        {
            Text       = "※ 서브메뉴 [저장]으로 서버에 최종 반영됩니다.",
            FontSize   = 10, FontFamily = Font,
            Foreground = new SolidColorBrush(Color.Parse("#666666")),
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
        });

        PricePanelChanged?.Invoke(root);
    }

    /// <summary>서브메뉴 [저장] 호출 시 pending 단가를 DB에 반영</summary>
    public void SavePendingPrices()
    {
        if (_selectedContract == null || _pendingPrices.Count == 0)
        {
            Log("단가 저장 스킵: 선택 없음 또는 변경 없음");
            return;
        }
        var priceList = _pendingPrices.Select(kv => (kv.Key, kv.Value)).ToList();
        bool ok = ContractService.UpdateContractPrices(_selectedContract.C_CompanyName, priceList);
        if (ok)
        {
            // 저장 완료 후 표시 색 초기화 (노란색 → 초록색)
            foreach (var (analyte, _) in priceList)
            {
                if (_priceDisplayBlocks.TryGetValue(analyte, out var tb))
                    tb.Foreground = new SolidColorBrush(Color.Parse("#aaddaa"));
            }
            _pendingPrices.Clear();
            Log($"✅ 단가 최종 저장: {_selectedContract.C_CompanyName} ({priceList.Count}건)");
        }
        else
        {
            Log($"❌ 단가 저장 실패: {_selectedContract.C_CompanyName}");
        }
    }

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [Contract] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("Logs/ContractDebug.log", line + Environment.NewLine); } catch { }
    }
}
