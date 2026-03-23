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
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services;

namespace ETA.Views.Pages;

public partial class ContractPage : UserControl
{
    // ── 외부(MainPage) 연결 ──────────────────────────────────────────────────
    public event Action<Control?>? DetailPanelChanged;

    // ── 상태 ────────────────────────────────────────────────────────────────
    private Contract?   _selectedContract;
    private StackPanel? _detailPanel;
    private bool        _isAddMode = false;

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

        try
        {
            var items = ContractService.GetAllContracts()
                            .OrderBy(c => c.C_CompanyName)
                            .ToList();

            foreach (var item in items)
                ContractTreeView.Items.Add(CreateTreeItem(item));

            Log($"로드 완료 → {items.Count}건");
        }
        catch (Exception ex) { Log("★ 크래시 ★ " + ex.Message); }
    }

    // =========================================================================
    // TreeViewItem 생성
    // =========================================================================
    private static TreeViewItem CreateTreeItem(Contract contract)
    {
        // 계약 타입별 아이콘
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
                                Foreground = Brushes.WhiteSmoke,
                            },
                            new TextBlock
                            {
                                Text       = string.IsNullOrEmpty(contract.C_Abbreviation)
                                                 ? contract.C_ContractType
                                                 : $"{contract.C_Abbreviation} · {contract.C_ContractType}",
                                FontSize   = 10,
                                FontFamily = Font,
                                Foreground = new SolidColorBrush(Color.Parse("#888888")),
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
        _detailPanel      = BuildEditPanel(contract);
        DetailPanelChanged?.Invoke(_detailPanel);
        Log($"선택: {contract.C_CompanyName}");
    }

    // =========================================================================
    // 계약업체 추가 패널  (MainPage BT3)
    // =========================================================================
    public void ShowAddPanel()
    {
        _selectedContract              = null;
        _isAddMode                     = true;
        ContractTreeView.SelectedItem  = null;
        _detailPanel                   = BuildAddPanel();
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

    // =========================================================================
    // 패널 — 수정 모드 (업체명 🔒)
    // =========================================================================
    private StackPanel BuildEditPanel(Contract c)
    {
        var root = MakeRootPanel($"🏢  {c.C_CompanyName} — 계약 정보");

        // 2열 그리드 레이아웃
        var grid = MakeTwoColumnGrid();

        AddGridRow(grid, 0, "업체명",           c.C_CompanyName,    isReadOnly: true,  isLocked: true);
        AddGridRow(grid, 1, "약칭",             c.C_Abbreviation);
        AddGridRow(grid, 2, "계약구분",         c.C_ContractType);
        AddGridRow(grid, 3, "계약시작",         c.C_ContractStartStr, isReadOnly: true);
        AddGridRow(grid, 4, "계약종료",         c.C_ContractEndStr,   isReadOnly: true);
        AddGridRow(grid, 5, "계약일수",         c.C_ContractDays?.ToString() ?? "");
        AddGridRow(grid, 6, "계약금액(VAT별도)", c.C_ContractAmountStr);
        AddGridRow(grid, 7, "주소",             c.C_Address);
        AddGridRow(grid, 8, "대표자",           c.C_Representative);
        AddGridRow(grid, 9, "시설별",           c.C_FacilityType);
        AddGridRow(grid,10, "종류별",           c.C_CategoryType);
        AddGridRow(grid,11, "주생산품",         c.C_MainProduct);
        AddGridRow(grid,12, "담당자",           c.C_ContactPerson);
        AddGridRow(grid,13, "연락처",           c.C_PhoneNumber);
        AddGridRow(grid,14, "이메일",           c.C_Email);

        root.Children.Add(grid);
        return root;
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
        AddGridRow(grid, 2, "계약구분",         "");
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
    private static void SyncPanelToContract(StackPanel panel, Contract c, bool includeReadOnly)
    {
        // 패널 안의 Grid 찾기
        var grid = panel.Children.OfType<Grid>().FirstOrDefault();
        if (grid == null) return;

        foreach (var child in grid.Children.OfType<StackPanel>())
        {
            if (child.Children.Count < 2) continue;
            var tb = child.Children[1] as TextBox;
            if (tb == null) continue;
            if (tb.IsReadOnly && !includeReadOnly) continue;

            var label = (child.Children[0] as TextBlock)?.Text ?? "";
            label = label.Replace("🔒 ", "").Replace("    ", "").Replace(" :", "").Trim();

            switch (label)
            {
                case "업체명":           c.C_CompanyName    = tb.Text ?? ""; break;
                case "약칭":             c.C_Abbreviation   = tb.Text ?? ""; break;
                case "계약구분":         c.C_ContractType   = tb.Text ?? ""; break;
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

    // 2열 Grid (필드가 많아서 2열로 배치)
    private static Grid MakeTwoColumnGrid()
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,*"),
            ColumnSpacing     = 16,
            RowSpacing        = 8,
        };
        // 15행 추가
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
                                    : Brushes.LightGray,
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
                                  : Brushes.WhiteSmoke,
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#333333"))
                                  : new SolidColorBrush(Color.Parse("#555577")),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 4),
        });

        // 홀수 행은 왼쪽, 짝수 행은 오른쪽 열
        int col = row % 2;
        int gridRow = row / 2;
        Grid.SetColumn(panel, col);
        Grid.SetRow(panel, gridRow);
        grid.Children.Add(panel);
    }

    // 확인 다이얼로그
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
                new TextBlock { Text = message, Foreground = Brushes.WhiteSmoke,
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

    private void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [Contract] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText("ContractDebug.log", line + Environment.NewLine); } catch { }
    }
}