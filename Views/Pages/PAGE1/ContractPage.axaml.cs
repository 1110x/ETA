using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
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
using ETA.Views;
using ETA.Views.Controls;

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
    public event Action<Control?>? EditPanelChanged;   // Show3 (진행상황 프로그래스바)
    public event Action<Control?>? PricePanelChanged;  // Show2 (단가 테이블)
    public event Action<Control?>? StatsPanelChanged;  // Show4 (단가 정보)

    // ── 상태 ────────────────────────────────────────────────────────────────
    private bool _suppressTotalUpdate;  // Show4 빌드 중 TextChanged 억제
    private Contract?   _selectedContract;
    private StackPanel? _detailPanel;        // 내부 StackPanel (SyncPanelToContract용)
    private TextBox?    _contractAmountTextBox;
    private bool        _isAddMode = false;
    private bool        _showActive = true;   // true=현행계약, false=종료계약
    private bool        _trashMode = false;   // true=휴지통, false=일반
    public MainPage? ParentMainPage { get; set; }
    private TextBlock? _txbTotalAmount;      // Show1 합계 표시용

    // 단가 in-memory 편집본 (서브메뉴 저장 시 DB 반영)
    private readonly Dictionary<string, string>    _pendingPrices       = new();
    // Show2 단가 표시용 TextBlock 참조 (항목명 → 표시 TextBlock)
    private readonly Dictionary<string, TextBlock> _priceDisplayBlocks  = new();

    // Show4 수량 편집 모드
    private bool                                   _show4EditMode       = false;
    private readonly Dictionary<string, TextBox>   _show4QtyTextBoxes   = new();
    private readonly Dictionary<string, TextBox>   _show4PriceTextBoxes = new();
    private readonly Dictionary<string, CheckBox>  _show4CheckBoxes     = new();
    private readonly Dictionary<string, string>    _pendingQuantities   = new();

    // Show4 패널 재사용 — 업체 변경 시 값만 갱신 (구조 재빌드 회피)
    private Control?    _show4CachedRoot;
    private TextBlock?  _show4HeaderCompanyTb;
    private Border?     _show4HeaderAbbrBorder;
    private TextBlock?  _show4HeaderAbbrTb;
    private TextBlock?  _show4ItemCountTb;
    private TextBlock?  _show4TotalSummaryTb;
    private ComboBox?   _show4CompareCb;
    private CheckBox?   _show4SelectAllChk;
    private readonly Dictionary<string, TextBlock>  _show4SubtotalTbs = new();
    private readonly Dictionary<string, decimal>   _show4Subtotals   = new();
    private readonly List<CheckBox>                _show4AllCheckBoxes = new();
    private string?     _show4CurrentCompany;
    private StackPanel? _show4CurrentShow2Panel;
    private HashSet<string>? _show4KnownAnalytes;
    private List<Contract>? _show4CompareSnapshot;

    // Show3 placeholder 캐시 — 버튼 한 개라 비용은 작지만 재사용
    private Control?  _show3CachedPlaceholder;
    private Button?   _show3CalcBtn;
    private TextBlock? _show3HintTb;
    private Contract? _show3PlaceholderContract;
    private Dictionary<string, string>? _show3PlaceholderQtys;

    // Show2 (계약 정보) 캐시 — 구조 1회 빌드, 업체 변경 시 값만 갱신
    private bool          _suppressShow2Sync;
    private ScrollViewer? _show2CachedScroll;
    private StackPanel?   _show2CachedRoot;
    private StackPanel?   _show2CachedPriceContainer;
    private TextBlock?    _show2TitleTb;
    private TextBlock?    _show2PriceSectionTitleTb;
    private TextBox?      _show2TbCompanyName, _show2TbAbbreviation;
    private TextBox?      _show2TbStart, _show2TbEnd, _show2TbRemain, _show2TbAmount;
    private TextBox?      _show2TbAddress, _show2TbRepresentative, _show2TbFacilityType;
    private TextBox?      _show2TbCategoryType, _show2TbMainProduct;
    private TextBox?      _show2TbContactPerson, _show2TbPhoneNumber, _show2TbEmail;
    private ComboBox?     _show2BasisCombo, _show2PlaceCombo;
    private Action<string, string?>? _show2RefreshPlaces;

    // 신규 추가 시 선택한 템플릿 업체명 (단가 복사용)
    private string? _templateCompanyName;

    // Show4 메타데이터 캐시 (LoadData/Save* 시 무효화) — 선택 변경 시 DB 재조회 회피
    private List<Contract>? _cachedAllContracts;
    private List<ETA.Models.AnalysisItem>? _cachedAllItems;
    private Dictionary<string, string>? _cachedAnalyteAliasMap;
    private List<MeasurerService.MeasurerContract>? _cachedMeasurerContracts;

    private void InvalidateMetadataCache()
    {
        _cachedAllContracts      = null;
        _cachedAllItems          = null;
        _cachedAnalyteAliasMap   = null;
        _cachedMeasurerContracts = null;

        // 메타데이터가 바뀌면 캐시된 패널 구조도 폐기
        InvalidateShow4Cache();
        InvalidateShow3Cache();
        InvalidateShow2Cache();
    }

    private void InvalidateShow2Cache()
    {
        _show2CachedScroll         = null;
        _show2CachedRoot           = null;
        _show2CachedPriceContainer = null;
        _show2TitleTb              = null;
        _show2PriceSectionTitleTb  = null;
        _show2TbCompanyName = _show2TbAbbreviation = null;
        _show2TbStart = _show2TbEnd = _show2TbRemain = _show2TbAmount = null;
        _show2TbAddress = _show2TbRepresentative = _show2TbFacilityType = null;
        _show2TbCategoryType = _show2TbMainProduct = null;
        _show2TbContactPerson = _show2TbPhoneNumber = _show2TbEmail = null;
        _show2BasisCombo = _show2PlaceCombo = null;
        _show2RefreshPlaces = null;
    }

    private void InvalidateShow4Cache()
    {
        _show4CachedRoot         = null;
        _show4HeaderCompanyTb    = null;
        _show4HeaderAbbrBorder   = null;
        _show4HeaderAbbrTb       = null;
        _show4ItemCountTb        = null;
        _show4TotalSummaryTb     = null;
        _show4CompareCb          = null;
        _show4SelectAllChk       = null;
        _show4SubtotalTbs.Clear();
        _show4Subtotals.Clear();
        _show4AllCheckBoxes.Clear();
        _show4CurrentCompany     = null;
        _show4CurrentShow2Panel  = null;
        _show4KnownAnalytes      = null;
        _show4CompareSnapshot    = null;
    }

    private void InvalidateShow3Cache()
    {
        _show3CachedPlaceholder  = null;
        _show3CalcBtn            = null;
        _show3HintTb             = null;
        _show3PlaceholderContract = null;
        _show3PlaceholderQtys    = null;
    }

    // 측정인 관리 패널 상태
    private string?    _selectedMeasCompany;
    private TextBox?   _txbMeasAlias;
    private TextBox?   _txbMeasAmount;
    private ComboBox?  _cmbMeasQuotType;
    private StackPanel? _measEditForm;

    public ContractPage()
    {
        InitializeComponent();
        this.AttachedToVisualTree += (s, e) =>
        {
            var window = this.FindAncestorOfType<Window>();
            if (window != null)
                window.KeyDown += OnWindowKeyDown;
        };
    }

    private void OnWindowKeyDown(object? sender, KeyEventArgs e)
    {
        // Shift+2: Show4 수량 편집 모드 토글 (TextBox 포커스 중엔 무시)
        if (e.KeyModifiers.HasFlag(KeyModifiers.Shift) && e.Key == Key.D2 &&
            TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() is not TextBox)
        {
            _show4EditMode = !_show4EditMode;

            foreach (var (analyte, textBox) in _show4QtyTextBoxes)
            {
                var display = textBox.Parent as StackPanel;
                if (display?.Children.Count > 1)
                {
                    var displayBlock = display.Children[1];
                    textBox.IsVisible = _show4EditMode;
                    if (displayBlock is TextBlock tb)
                        tb.IsVisible = !_show4EditMode;

                    if (_show4EditMode)
                    {
                        textBox.Text = _pendingQuantities.TryGetValue(analyte, out var q) ? q : "";
                    }
                }
            }

            e.Handled = true;
        }
    }

    // =========================================================================
    // 데이터 로드
    // =========================================================================
    public void LoadData()
    {
        Log("LoadData() 시작");
        InvalidateMetadataCache();
        ContractTreeView.Items.Clear();
        _selectedContract = null;
        _isAddMode        = false;
        DetailPanelChanged?.Invoke(null);
        PricePanelChanged?.Invoke(null);

        try
        {
            List<Contract> items;

            if (_trashMode)
            {
                // 휴지통: 삭제된 계약 목록
                items = ContractService.GetDeletedContracts()
                            .OrderBy(c => c.C_CompanyName)
                            .ToList();
                Log($"휴지통 로드 → {items.Count}건");
            }
            else
            {
                // 일반: 현행/종료 계약
                var today = DateTime.Today;
                items = ContractService.GetAllContracts()
                                .Where(c =>
                                {
                                    bool active = c.C_ContractStart.HasValue && c.C_ContractEnd.HasValue
                                                  && c.C_ContractStart.Value.Date <= today
                                                  && c.C_ContractEnd.Value.Date   >= today;
                                    return _showActive ? active : !active;
                                })
                                .OrderBy(c => c.C_CompanyName)
                                .ToList();
                Log($"일반 로드 → {items.Count}건 ({(_showActive ? "현행" : "종료")})");
            }

            foreach (var item in items)
                ContractTreeView.Items.Add(CreateTreeItem(item));

            // 계약금액 합계 계산 및 표시 (휴지통 제외)
            if (_txbTotalAmount == null)
                _txbTotalAmount = this.FindControl<TextBlock>("txbTotalAmount");

            if (_txbTotalAmount != null)
            {
                if (_trashMode)
                {
                    _txbTotalAmount.Text = "";
                }
                else
                {
                    decimal totalAmount = items
                        .Where(c => c.C_ContractAmountVATExcluded.HasValue)
                        .Sum(c => c.C_ContractAmountVATExcluded.Value);

                    _txbTotalAmount.Text = totalAmount > 0
                        ? $"계약금액 합계: {totalAmount:N0} 원"
                        : "계약금액 합계: -";
                }
            }
        }
        catch (Exception ex) { Log("★ 크래시 ★ " + ex.Message); }
    }

    private void CmbContractView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is not ComboBox cb) return;

        int index = cb.SelectedIndex;
        _trashMode = (index == 2);  // 2 = 휴지통
        _showActive = (index == 0); // 0 = 현행, 1 = 종료

        LoadData();
    }

    private void TglContractStatus_Changed(object? sender, RoutedEventArgs e)
    {
        if (sender is not ToggleSwitch ts) return;
        _showActive = ts.IsChecked == true;
        _trashMode = false;
        LoadData();
    }

    // =========================================================================
    // TreeViewItem 생성
    // =========================================================================
    private static TreeViewItem CreateTreeItem(Contract contract)
    {
        string icon = contract.C_ContractType switch
        {
            "용역" => "📋",
            "구매" => "🛒",
            _     => "🏢"
        };

        var tvi = new TreeViewItem
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
                        Orientation = Orientation.Horizontal,
                        Spacing     = 6,
                        VerticalAlignment = VerticalAlignment.Center,
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
                                        FontSize   = AppTheme.FontXS, FontFamily = Font,
                                        Foreground = Brush.Parse(BadgeColorHelper.GetBadgeColor(contract.C_Abbreviation).Fg),
                                    }
                                },
                            new TextBlock
                            {
                                Text       = contract.C_CompanyName,
                                FontSize   = AppTheme.FontLG,
                                FontFamily = Font,
                                Foreground = AppRes("AppFg"),
                            }
                        }
                    }
                }
            }
        };
        TextShimmer.AttachHover(tvi);
        return tvi;
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

        _contractAmountTextBox = null;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        _selectedContract = contract;
        _isAddMode        = false;
        _pendingPrices.Clear();
        _priceDisplayBlocks.Clear();
        _show4EditMode    = false;
        // 주의: _show4QtyTextBoxes/_show4PriceTextBoxes/_show4CheckBoxes 는 캐시된 Show4 행 위젯 참조이므로 Clear 금지
        _pendingQuantities.Clear();
        Log($"[타이밍] 초기화: {sw.ElapsedMilliseconds}ms");

        // 휴지통 모드: 복구/삭제 버튼 표시
        if (_trashMode)
        {
            var (root, scroll, _) = BuildInfoPanel(contract);
            _detailPanel = root;
            DetailPanelChanged?.Invoke(scroll);

            var trashPanel = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

            var (rBg, rFg, rBd) = StatusBadge.GetBrushes(BadgeStatus.Info);
            var restoreBtn = new Button
            {
                Content         = "♻️  복구",
                FontSize        = AppTheme.FontMD,
                FontFamily      = Font, FontWeight = FontWeight.SemiBold,
                Background      = rBg, Foreground = rFg, BorderBrush = rBd,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(999),
                Padding         = new Thickness(14, 6),
                Cursor          = new Cursor(StandardCursorType.Hand),
            };
            restoreBtn.Click += (_, _) =>
            {
                if (ContractService.Restore(contract.C_CompanyName))
                {
                    Log($"✅ {contract.C_CompanyName} 복구 완료");
                    LoadData();
                }
                else
                {
                    Log($"❌ {contract.C_CompanyName} 복구 실패");
                }
            };

            var (dBg, dFg, dBd) = StatusBadge.GetBrushes(BadgeStatus.Bad);
            var deleteBtn = new Button
            {
                Content         = "🗑️  영구 삭제",
                FontSize        = AppTheme.FontMD,
                FontFamily      = Font, FontWeight = FontWeight.SemiBold,
                Background      = dBg, Foreground = dFg, BorderBrush = dBd,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(999),
                Padding         = new Thickness(14, 6),
                Cursor          = new Cursor(StandardCursorType.Hand),
            };
            deleteBtn.Click += (_, _) =>
            {
                if (ContractService.PermanentDelete(contract.C_CompanyName))
                {
                    Log($"✅ {contract.C_CompanyName} 영구 삭제 완료");
                    LoadData();
                }
                else
                {
                    Log($"❌ {contract.C_CompanyName} 영구 삭제 실패");
                }
            };

            trashPanel.Children.Add(new TextBlock
            {
                Text       = "삭제된 계약",
                FontSize   = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgMuted,
                Margin     = new Thickness(0, 0, 0, 8),
            });
            trashPanel.Children.Add(restoreBtn);
            trashPanel.Children.Add(deleteBtn);

            var scroll2 = new ScrollViewer
            {
                Content = trashPanel,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            };
            EditPanelChanged?.Invoke(scroll2);
            StatsPanelChanged?.Invoke(null);
            return;
        }

        // ① 계약정보 패널 즉시 표시 (DB 없음 — 빠름)
        var (root2, scroll3, priceContainer) = BuildInfoPanel(contract);
        Log($"[타이밍] BuildInfoPanel: {sw.ElapsedMilliseconds}ms");
        _detailPanel = root2;
        DetailPanelChanged?.Invoke(scroll3);  // Show2에 업체정보 표시
        Log($"[타이밍] Show2 표시: {sw.ElapsedMilliseconds}ms");

        // ② Show4에 로딩 프로그레스바 표시 (최초 빌드 시에만 — 이후는 캐시된 패널 값만 갱신되므로 깜빡임 없음)
        var captured = contract;
        Avalonia.Controls.ProgressBar? loadingProgressBar = null;
        TextBlock? loadingStatusText = null;
        if (_show4CachedRoot == null)
        {
            loadingProgressBar = new Avalonia.Controls.ProgressBar
            {
                Minimum         = 0,
                Maximum         = 100,
                Value           = 0,
                IsIndeterminate = false,
                Height          = 4,
                Margin          = new Thickness(0, 0, 0, 8),
            };
            loadingStatusText = new TextBlock
            {
                Text       = "단가 정보 로딩 중...",
                FontSize   = AppTheme.FontSM, FontFamily = Font,
                Foreground = AppTheme.FgMuted,
            };
            var loadingPanel = new StackPanel { Margin = new Thickness(12) };
            loadingPanel.Children.Add(loadingProgressBar);
            loadingPanel.Children.Add(loadingStatusText);
            StatsPanelChanged?.Invoke(loadingPanel);
        }

        // ③ 단가/계약수량/처리수량 백그라운드 로드
        _ = LoadContractDataAsync(captured, loadingProgressBar, loadingStatusText);
    }

    private async Task LoadContractDataAsync(Contract captured,
        Avalonia.Controls.ProgressBar? progressBar, TextBlock? statusText)
    {
        try
        {
            Log($"[로드 시작] {captured.C_CompanyName}");
            var sw2 = System.Diagnostics.Stopwatch.StartNew();

            Log($"  → GetContractPrices 시작...");
            var pricesSw = System.Diagnostics.Stopwatch.StartNew();
            var prices       = await Task.Run(() => ContractService.GetContractPrices(captured.C_CompanyName));
            Log($"  → GetContractPrices 완료: {pricesSw.ElapsedMilliseconds}ms, {prices.Count}개");
            if (progressBar != null) progressBar.Value = 33;
            if (statusText  != null) statusText.Text   = "계약수량 로딩 중...";

            Log($"  → GetContractQuantities 시작...");
            var qtySw = System.Diagnostics.Stopwatch.StartNew();
            var contractQtys = await Task.Run(() => ContractService.GetContractQuantities(captured.C_CompanyName));
            Log($"  → GetContractQuantities 완료: {qtySw.ElapsedMilliseconds}ms, {contractQtys.Count}개");
            if (progressBar != null) progressBar.Value = 66;
            if (statusText  != null) statusText.Text   = "처리수량 로딩 중...";

            if (progressBar != null) progressBar.Value = 100;
            if (statusText  != null) statusText.Text   = "완료";
            Log($"[데이터 로드] {captured.C_CompanyName}: 단가={prices.Count}개, 계약수량={contractQtys.Count}개, 누적={sw2.ElapsedMilliseconds}ms");

            if (_selectedContract?.C_CompanyName != captured.C_CompanyName)
            {
                Log($"[로드 취소] {captured.C_CompanyName} (선택 변경됨)");
                return;
            }

            // Show3: 진행상황 — 무거운 계산은 버튼 클릭 시점까지 지연 (캐시 재사용됨)
            var swUi = System.Diagnostics.Stopwatch.StartNew();
            var show3Placeholder = BuildProgressPlaceholder(captured, contractQtys);
            EditPanelChanged?.Invoke(show3Placeholder);
            Log($"[타이밍] Show3 placeholder 렌더링: {swUi.ElapsedMilliseconds}ms");

            // Show4: 단가 정보 + 계약수량 (캐시 재사용 시 값만 갱신 — 깜빡임 없음)
            await Task.Delay(1);  // UI 스레드 양보, Show3가 먼저 렌더링되도록
            if (_selectedContract?.C_CompanyName != captured.C_CompanyName) return;

            swUi.Restart();
            var priceInfoPanel = BuildPriceInfoPanel(captured.C_CompanyName, captured.C_Abbreviation, prices, contractQtys, _detailPanel);
            Log($"[타이밍] BuildPriceInfoPanel: {swUi.ElapsedMilliseconds}ms");
            StatsPanelChanged?.Invoke(priceInfoPanel);
            Log($"[타이밍] Show4 렌더링: {swUi.ElapsedMilliseconds}ms");
            await Task.Yield();

            Log($"[타이밍] 전체 완료: {sw2.ElapsedMilliseconds}ms / [Show4 표시 완료] 단가={prices.Count}개");
        }
        catch (Exception ex)
        {
            Log($"데이터 로드 오류: {ex.Message}");
        }
    }


    // =========================================================================
    // 계약업체 추가 패널  (MainPage BT3)
    // =========================================================================
    public void ShowAddPanel()
    {
        _selectedContract              = null;
        _isAddMode                     = true;
        _templateCompanyName           = null;
        _pendingPrices.Clear();
        _priceDisplayBlocks.Clear();
        _contractAmountTextBox         = null;
        ContractTreeView.SelectedItem  = null;
        _detailPanel                   = BuildAddPanel();
        DetailPanelChanged?.Invoke(_detailPanel);
        ShowAddModeTemplateSelection();  // Show3 템플릿 선택
        StatsPanelChanged?.Invoke(null);  // Show4 초기화
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
            // Show4(단가/수량) 현재 UI 상태를 pending으로 스냅샷 — 확정 버튼 없이도 저장되게.
            CommitShow4Changes();
            SaveEdit();
            SavePendingPrices();
            SavePendingQuantities();
        }
    }

    /// <summary>Show4 현재 체크박스/단가/수량 UI 상태를 pending dict 에 반영.
    /// 체크된 항목 → 현재 textbox 값, 체크해제 → 빈 문자열(DBNull로 저장). 저장 직전 호출.</summary>
    private void CommitShow4Changes()
    {
        foreach (var analyte in _show4CheckBoxes.Keys)
        {
            bool isChecked = _show4CheckBoxes[analyte].IsChecked == true;
            if (isChecked)
            {
                if (_show4PriceTextBoxes.TryGetValue(analyte, out var pb))
                    _pendingPrices[analyte] = pb.Text ?? "";
                if (_show4QtyTextBoxes.TryGetValue(analyte, out var qb))
                    _pendingQuantities[analyte] = qb.Text ?? "";
            }
            else
            {
                // 체크해제 → NULL 로 저장 (빈 문자열은 ContractService 에서 DBNull 처리됨)
                _pendingPrices[analyte] = "";
                _pendingQuantities[analyte] = "";
            }
        }
        Log($"[CommitShow4] 체크박스 {_show4CheckBoxes.Count}개 → pending prices={_pendingPrices.Count}, qtys={_pendingQuantities.Count}");
    }

    private void SavePendingQuantities()
    {
        if (_selectedContract == null || _pendingQuantities.Count == 0) return;

        try
        {
            var quantities = _pendingQuantities.Select(kv => (Analyte: kv.Key, Qty: kv.Value)).ToList();
            bool ok = ContractService.UpdateContractQuantities(_selectedContract.C_CompanyName, quantities);
            if (ok)
            {
                Log($"✅ 계약수량 저장: {_selectedContract.C_CompanyName} ({_pendingQuantities.Count}항목)");
                _pendingQuantities.Clear();
            }
            else
                Log($"❌ 계약수량 저장 실패: {_selectedContract.C_CompanyName}");
        }
        catch (Exception ex) { Log($"❌ 계약수량 저장 오류: {ex.Message}"); }
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
            _contractAmountTextBox = null;
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
    //   구조는 최초 1회만 빌드하고 업체 변경 시엔 값만 갱신
    // =========================================================================
    /// <returns>(innerStackPanel, scrollWrapper, priceContainer)</returns>
    private (StackPanel root, ScrollViewer scroll, StackPanel priceContainer) BuildInfoPanel(Contract c)
    {
        if (_show2CachedScroll != null && _show2CachedRoot != null && _show2CachedPriceContainer != null)
        {
            PopulateShow2Values(c);
            _contractAmountTextBox = _show2TbAmount;
            return (_show2CachedRoot, _show2CachedScroll, _show2CachedPriceContainer);
        }
        return BuildShow2Structure(c);
    }

    private (StackPanel root, ScrollViewer scroll, StackPanel priceContainer) BuildShow2Structure(Contract c)
    {
        var root = MakeRootPanel($"🏢  {c.C_CompanyName} — 계약 정보");
        // MakeRootPanel adds the title TextBlock as the first child
        _show2TitleTb = root.Children.OfType<TextBlock>().FirstOrDefault();

        var grid = MakeTwoColumnGrid();

        _show2TbCompanyName   = AddGridRow(grid, 0, "업체명",           c.C_CompanyName,      isReadOnly: true, isLocked: true);
        _show2TbAbbreviation  = AddGridRow(grid, 1, "약칭",             c.C_Abbreviation);

        string remainDays = c.C_ContractEnd.HasValue
            ? (c.C_ContractEnd.Value.Date - DateTime.Today).Days.ToString()
            : "";
        string amountStr = c.C_ContractAmountVATExcluded.HasValue
            ? c.C_ContractAmountVATExcluded.Value.ToString("N0") + " 원"
            : "";

        _show2TbStart   = AddGridRow(grid, 2, "계약시작", c.C_ContractStartStr, hint: "예) 20240101");
        _show2TbEnd     = AddGridRow(grid, 3, "계약종료", c.C_ContractEndStr,   hint: "예) 20241231");
        _show2TbRemain  = AddGridRow(grid, 4, "잔여일수", remainDays, isReadOnly: true);
        _show2TbAmount  = AddGridRow(grid, 5, "계약금액(VAT별도)", amountStr);
        _contractAmountTextBox = _show2TbAmount;

        var tbStart  = _show2TbStart;
        var tbEnd    = _show2TbEnd;
        var tbRemain = _show2TbRemain;

        static string FormatDate(string raw)
        {
            var digits = new string(raw.Where(char.IsDigit).ToArray());
            if (digits.Length == 8) return $"{digits[..4]}-{digits[4..6]}-{digits[6..8]}";
            return raw;
        }
        void RecalcRemain()
        {
            tbStart.Text = FormatDate(tbStart.Text ?? "");
            tbEnd.Text   = FormatDate(tbEnd.Text   ?? "");
            tbRemain.Text = DateTime.TryParse(tbEnd.Text ?? "", out var ed)
                ? (ed.Date - DateTime.Today).Days.ToString()
                : "";
        }
        tbStart.LostFocus += (_, _) => { if (!_suppressShow2Sync) RecalcRemain(); };
        tbEnd.LostFocus   += (_, _) => { if (!_suppressShow2Sync) RecalcRemain(); };

        AddBasisContractComboBox(grid, 6, c.C_ContractType, c.C_PlaceName);

        _show2TbAddress        = AddGridRow(grid, 8,  "주소",     c.C_Address);
        _show2TbRepresentative = AddGridRow(grid, 9,  "대표자",   c.C_Representative);
        _show2TbFacilityType   = AddGridRow(grid, 10, "시설별",   c.C_FacilityType);
        _show2TbCategoryType   = AddGridRow(grid, 11, "종류별",   c.C_CategoryType);
        _show2TbMainProduct    = AddGridRow(grid, 12, "주생산품", c.C_MainProduct);
        _show2TbContactPerson  = AddGridRow(grid, 13, "담당자",   c.C_ContactPerson);
        _show2TbPhoneNumber    = AddGridRow(grid, 14, "연락처",   c.C_PhoneNumber);
        _show2TbEmail          = AddGridRow(grid, 15, "이메일",   c.C_Email);
        root.Children.Add(grid);

        // ── 단가 섹션 컨테이너 (백그라운드 로드 후 채워짐) ───────────────────
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = AppTheme.BorderMuted,
            Margin     = new Thickness(0, 16, 0, 8),
        });

        var currentUser = ETA.Services.Common.CurrentUserManager.Instance.CurrentUserId;
        var priceContainer = new StackPanel { Spacing = 4 };

        if (_priceAllowedUsers.Contains(currentUser))
        {
            _show2PriceSectionTitleTb = new TextBlock
            {
                Text       = $"📊  분석 단가 — {c.C_CompanyName}",
                FontSize   = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgMuted,
                Margin     = new Thickness(0, 0, 0, 4),
            };
            root.Children.Add(_show2PriceSectionTitleTb);
            root.Children.Add(new TextBlock
            {
                Text       = "항목 클릭 → 오른쪽에서 편집  |  서브메뉴 [저장] → 서버 반영",
                FontSize   = AppTheme.FontSM, FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Thickness(0, 0, 0, 6),
            });
            priceContainer.Children.Add(new TextBlock
            {
                Text       = "⏳ 단가 로드 중...",
                FontSize   = AppTheme.FontSM, FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
            });
        }

        root.Children.Add(priceContainer);

        var scroll = new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };

        _show2CachedRoot           = root;
        _show2CachedScroll         = scroll;
        _show2CachedPriceContainer = priceContainer;
        return (root, scroll, priceContainer);
    }

    /// <summary>캐시된 Show2 위젯의 값만 갱신.</summary>
    private void PopulateShow2Values(Contract c)
    {
        _suppressShow2Sync = true;
        try
        {
            if (_show2TitleTb != null) _show2TitleTb.Text = $"🏢  {c.C_CompanyName} — 계약 정보";
            if (_show2PriceSectionTitleTb != null) _show2PriceSectionTitleTb.Text = $"📊  분석 단가 — {c.C_CompanyName}";

            if (_show2TbCompanyName   != null) _show2TbCompanyName.Text   = c.C_CompanyName     ?? "";
            if (_show2TbAbbreviation  != null) _show2TbAbbreviation.Text  = c.C_Abbreviation    ?? "";
            if (_show2TbStart         != null) _show2TbStart.Text         = c.C_ContractStartStr ?? "";
            if (_show2TbEnd           != null) _show2TbEnd.Text           = c.C_ContractEndStr   ?? "";
            if (_show2TbRemain        != null)
            {
                _show2TbRemain.Text = c.C_ContractEnd.HasValue
                    ? (c.C_ContractEnd.Value.Date - DateTime.Today).Days.ToString()
                    : "";
            }
            if (_show2TbAmount != null)
            {
                _show2TbAmount.Text = c.C_ContractAmountVATExcluded.HasValue
                    ? c.C_ContractAmountVATExcluded.Value.ToString("N0") + " 원"
                    : "";
            }
            if (_show2TbAddress        != null) _show2TbAddress.Text        = c.C_Address        ?? "";
            if (_show2TbRepresentative != null) _show2TbRepresentative.Text = c.C_Representative ?? "";
            if (_show2TbFacilityType   != null) _show2TbFacilityType.Text   = c.C_FacilityType   ?? "";
            if (_show2TbCategoryType   != null) _show2TbCategoryType.Text   = c.C_CategoryType   ?? "";
            if (_show2TbMainProduct    != null) _show2TbMainProduct.Text    = c.C_MainProduct    ?? "";
            if (_show2TbContactPerson  != null) _show2TbContactPerson.Text  = c.C_ContactPerson  ?? "";
            if (_show2TbPhoneNumber    != null) _show2TbPhoneNumber.Text    = c.C_PhoneNumber    ?? "";
            if (_show2TbEmail          != null) _show2TbEmail.Text          = c.C_Email          ?? "";

            if (_show2BasisCombo != null)
            {
                string? matchLabel = null;
                var contracts = _cachedMeasurerContracts;
                if (!string.IsNullOrEmpty(c.C_ContractType) && contracts != null)
                {
                    var match = contracts.FirstOrDefault(x => x.계약번호 == c.C_ContractType);
                    if (!string.IsNullOrEmpty(match.계약번호))
                    {
                        matchLabel = string.IsNullOrWhiteSpace(match.전체텍스트)
                            ? $"{match.계약번호} / {match.업체명}"
                            : match.전체텍스트;
                    }
                }
                _show2BasisCombo.SelectedItem = matchLabel;
            }

            _show2RefreshPlaces?.Invoke(c.C_ContractType ?? "", c.C_PlaceName);
        }
        finally
        {
            _suppressShow2Sync = false;
        }
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
                FontSize   = AppTheme.FontBase, FontFamily = Font,
                Foreground = AppTheme.FgMuted,
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
                Width             = 60, FontSize = AppTheme.FontSM, FontFamily = Font,
                Foreground        = string.IsNullOrEmpty(price)
                                        ? AppTheme.BorderDefault
                                        : AppTheme.FgSuccess,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment     = Avalonia.Media.TextAlignment.Right,
            };
            _priceDisplayBlocks[analyte] = priceBlock;

            var cell = new Border
            {
                Background      = AppTheme.BgPrimary,
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
                            FontSize          = AppTheme.FontXS, FontFamily = Font,
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
    // Show4 — 단가 정보 + 계약수량 + 소계 (업체추가 Show4와 동일 레이아웃)
    //   구조는 최초 1회만 빌드하고, 업체 변경 시엔 값만 갱신 (느린 UI 재구성 회피)
    // =========================================================================
    private Control BuildPriceInfoPanel(string companyName,
                                        string abbr,
                                        List<(string Analyte, string Price)> prices,
                                        Dictionary<string, string> contractQtys,
                                        StackPanel? show2Panel = null,
                                        bool uncheckedByDefault = false)
    {
        if (_show4CachedRoot != null && !Show4RequiresRebuild(prices))
        {
            var swFast = System.Diagnostics.Stopwatch.StartNew();
            PopulateShow4Values(companyName, abbr, prices, contractQtys, show2Panel, uncheckedByDefault);
            Log($"[UI빌드] Show4 값 갱신 완료: {swFast.ElapsedMilliseconds}ms");
            return _show4CachedRoot;
        }

        InvalidateShow4Cache();
        var swBuild = System.Diagnostics.Stopwatch.StartNew();
        Log($"[UI빌드] BuildShow4Structure({companyName}) 전체 빌드");
        var root = BuildShow4Structure(companyName, abbr, prices, contractQtys, show2Panel, uncheckedByDefault);
        _show4CachedRoot = root;
        Log($"[UI빌드] BuildShow4Structure 완료: {swBuild.ElapsedMilliseconds}ms");
        return root;
    }

    /// <summary>현재 캐시된 Show4 구조가 유효한지 판단. (unknown analyte 포함 등)</summary>
    private bool Show4RequiresRebuild(List<(string Analyte, string Price)> prices)
    {
        if (_show4KnownAnalytes == null) return true;

        // 캐시가 allItems 범위를 초과하는 "extras" 를 포함한다면 다음 업체 선택 시 재빌드
        var allItemsAnalytes = (_cachedAllItems ?? new List<AnalysisItem>())
            .Select(a => a.Analyte)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var a in _show4KnownAnalytes)
            if (!allItemsAnalytes.Contains(a)) return true;

        foreach (var (analyte, price) in prices)
        {
            if (string.IsNullOrWhiteSpace(price)) continue;
            if (!_show4KnownAnalytes.Contains(analyte)) return true;
        }
        return false;
    }

    /// <summary>캐시된 Show4 행 위젯의 값만 갱신 (구조 재빌드 없이).</summary>
    private void PopulateShow4Values(string companyName,
                                     string abbr,
                                     List<(string Analyte, string Price)> prices,
                                     Dictionary<string, string> contractQtys,
                                     StackPanel? show2Panel,
                                     bool uncheckedByDefault)
    {
        _show4CurrentShow2Panel = show2Panel;
        _show4CurrentCompany    = companyName;

        _suppressTotalUpdate = true;
        try
        {
            if (_show4HeaderCompanyTb != null) _show4HeaderCompanyTb.Text = companyName;
            UpdateShow4AbbrBadge(abbr);
            if (_show4ItemCountTb != null) _show4ItemCountTb.Text = $"({prices.Count}개 항목)";

            RefreshShow4CompareDropdown();

            var priceDict = prices.ToDictionary(p => p.Analyte, p => p.Price, StringComparer.OrdinalIgnoreCase);
            _show4Subtotals.Clear();
            decimal grandTotal = 0m;

            foreach (var analyte in _show4CheckBoxes.Keys)
            {
                bool hasRealPrice = priceDict.TryGetValue(analyte, out var priceStr)
                    && !string.IsNullOrWhiteSpace(priceStr);
                bool hasRealQuantity = contractQtys.TryGetValue(analyte, out var rawQtyStr)
                    && !string.IsNullOrWhiteSpace(rawQtyStr);

                decimal priceValue = 0m;
                if (hasRealPrice) decimal.TryParse(priceStr, out priceValue);

                int contractQty;
                if (hasRealQuantity && int.TryParse(rawQtyStr, out int cqInt) && cqInt > 0)
                    contractQty = cqInt;
                else if (hasRealPrice)
                    contractQty = 1;
                else
                    contractQty = 0;

                decimal subtotal = hasRealPrice ? priceValue * contractQty : 0m;
                if (subtotal > 0) _show4Subtotals[analyte] = subtotal;
                grandTotal += subtotal;

                if (_show4CheckBoxes.TryGetValue(analyte, out var chk))
                    chk.IsChecked = !uncheckedByDefault && hasRealPrice;
                if (_show4PriceTextBoxes.TryGetValue(analyte, out var priceBox))
                    priceBox.Text = priceValue > 0 ? ((long)priceValue).ToString() : "";
                if (_show4QtyTextBoxes.TryGetValue(analyte, out var qtyBox))
                    qtyBox.Text = contractQty > 0 ? contractQty.ToString() : "";
                if (_show4SubtotalTbs.TryGetValue(analyte, out var stb))
                    stb.Text = subtotal > 0 ? subtotal.ToString("N0") : "—";
            }

            if (_show4TotalSummaryTb != null)
                _show4TotalSummaryTb.Text = $"💵 합계: {grandTotal.ToString("N0")} 원";
            // NOTE: 선택 전환 시 Show2 계약금액 자동 덮어쓰기 제거 (#6) — DB에 저장된 원본 값 보존.
            // 사용자가 Show4에서 가격/수량 편집할 때만 RecomputeShow4Total가 Show2 금액을 갱신함.
        }
        finally
        {
            _suppressTotalUpdate = false;
        }
    }

    private void UpdateShow4AbbrBadge(string abbr)
    {
        if (_show4HeaderAbbrBorder == null || _show4HeaderAbbrTb == null) return;
        if (string.IsNullOrEmpty(abbr))
        {
            _show4HeaderAbbrBorder.IsVisible = false;
            return;
        }
        var (bg, fg) = BadgeColorHelper.GetBadgeColor(abbr);
        _show4HeaderAbbrBorder.Background = Brush.Parse(bg);
        _show4HeaderAbbrTb.Text           = abbr;
        _show4HeaderAbbrTb.Foreground     = Brush.Parse(fg);
        _show4HeaderAbbrBorder.IsVisible  = true;
    }

    private void RefreshShow4CompareDropdown()
    {
        if (_show4CompareCb == null) return;
        var allContracts = _cachedAllContracts ??= ContractService.GetAllContracts();
        if (!ReferenceEquals(_show4CompareSnapshot, allContracts))
        {
            _show4CompareCb.Items.Clear();
            foreach (var c in allContracts.OrderBy(x => x.C_CompanyName))
            {
                var displayText = string.IsNullOrEmpty(c.C_Abbreviation)
                    ? c.C_CompanyName
                    : $"{c.C_CompanyName} ({c.C_Abbreviation})";
                _show4CompareCb.Items.Add(displayText);
            }
            _show4CompareSnapshot = allContracts;
        }
        _show4CompareCb.SelectedIndex = -1;
    }

    private void RecomputeShow4Total()
    {
        if (_suppressTotalUpdate) return;
        if (_show4TotalSummaryTb == null) return;
        decimal grand = 0m;
        foreach (var s in _show4Subtotals.Values) grand += s;
        _show4TotalSummaryTb.Text = $"💵 합계: {grand.ToString("N0")} 원";
        if (_show4CurrentShow2Panel != null)
            UpdateContractAmountInShow2(_show4CurrentShow2Panel, grand);
    }

    private Control BuildShow4Structure(string companyName,
                                        string abbr,
                                        List<(string Analyte, string Price)> prices,
                                        Dictionary<string, string> contractQtys,
                                        StackPanel? show2Panel,
                                        bool uncheckedByDefault)
    {
        _show4CurrentShow2Panel = show2Panel;
        _show4CurrentCompany    = companyName;
        _show4Subtotals.Clear();
        _show4AllCheckBoxes.Clear();
        _show4SubtotalTbs.Clear();
        _show4CheckBoxes.Clear();
        _show4PriceTextBoxes.Clear();
        _show4QtyTextBoxes.Clear();

        var root = new StackPanel { Spacing = 0, Margin = new Thickness(8) };

        // ─── 비교 계약 드롭다운 ─────────────────────────────────────────────────
        var allContracts = _cachedAllContracts ??= ContractService.GetAllContracts();
        _show4CompareSnapshot = allContracts;

        var lblCompare = new TextBlock
        {
            Text              = "비교 계약:",
            FontSize          = AppTheme.FontMD, FontFamily = Font,
            Foreground        = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var cbContract = new ComboBox
        {
            Width             = 200, Height = 34,
            FontSize          = AppTheme.FontMD, FontFamily = Font,
            Background        = AppTheme.BorderSeparator, Foreground = AppRes("AppFg"),
            BorderThickness   = new Thickness(1), BorderBrush = AppTheme.BorderDefault,
            Padding           = new Thickness(8, 4),
            VerticalAlignment = VerticalAlignment.Center,
        };
        _show4CompareCb = cbContract;

        foreach (var c in allContracts.OrderBy(x => x.C_CompanyName))
        {
            var displayText = string.IsNullOrEmpty(c.C_Abbreviation)
                ? c.C_CompanyName
                : $"{c.C_CompanyName} ({c.C_Abbreviation})";
            cbContract.Items.Add(displayText);
        }

        cbContract.SelectionChanged += async (_, _) =>
        {
            if (cbContract.SelectedIndex < 0) return;

            var selectedText = cbContract.SelectedItem as string;
            if (string.IsNullOrEmpty(selectedText)) return;

            var srcCompany = selectedText.Contains("(")
                ? selectedText[..selectedText.LastIndexOf("(")].Trim()
                : selectedText;

            var currentCompany = _show4CurrentCompany ?? _selectedContract?.C_CompanyName;
            if (string.IsNullOrEmpty(currentCompany)) { cbContract.SelectedIndex = -1; return; }

            var srcContract = (_cachedAllContracts ?? ContractService.GetAllContracts())
                .FirstOrDefault(c => c.C_CompanyName == srcCompany);
            if (srcContract == null) return;

            var selectedPrices = ContractService.GetContractPrices(srcCompany);
            if (selectedPrices.Count == 0)
            {
                Log($"⚠️ {srcCompany}에 단가가 없습니다");
                cbContract.SelectedIndex = -1;
                return;
            }

            try
            {
                Log($"💾 {srcCompany}의 단가를 {currentCompany}에 적용 중...");
                bool success = await Task.Run(() => ContractService.CopyContractPrices(srcCompany, currentCompany));

                if (success)
                {
                    Log($"✅ 단가 적용 완료: {selectedPrices.Count}개, Show4 값 갱신...");
                    cbContract.SelectedIndex = -1;

                    if (_selectedContract != null)
                    {
                        var updatedPrices = await Task.Run(() => ContractService.GetContractPrices(_selectedContract.C_CompanyName));
                        var updatedQtys   = await Task.Run(() => ContractService.GetContractQuantities(_selectedContract.C_CompanyName));
                        var refreshedPanel = BuildPriceInfoPanel(
                            _selectedContract.C_CompanyName,
                            _selectedContract.C_Abbreviation,
                            updatedPrices, updatedQtys, _detailPanel, uncheckedByDefault: true);
                        StatsPanelChanged?.Invoke(refreshedPanel);
                        Log($"✅ Show4 갱신 완료 (체크박스 해제)");
                    }
                }
                else
                {
                    Log($"❌ 단가 적용 실패");
                    cbContract.SelectedIndex = -1;
                }
            }
            catch (Exception ex)
            {
                Log($"❌ 오류: {ex.Message}");
                cbContract.SelectedIndex = -1;
            }
        };

        // ─── 항목명 → 약칭 맵 ───────────────────────────────────────────────────
        var aliasMap = _cachedAnalyteAliasMap ??= ContractService.GetAnalyteAliasMap();

        // ─── 헤더: 업체명 + 약칭 뱃지 ───────────────────────────────────────────
        var headerLine = new StackPanel
        {
            Orientation       = Orientation.Horizontal,
            Spacing           = 6,
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Thickness(0, 0, 0, 2),
        };
        _show4HeaderCompanyTb = new TextBlock
        {
            Text              = companyName,
            FontSize          = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground        = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        };
        headerLine.Children.Add(_show4HeaderCompanyTb);

        _show4HeaderAbbrTb = new TextBlock
        {
            FontSize = AppTheme.FontXS, FontFamily = Font,
            VerticalAlignment = VerticalAlignment.Center,
        };
        _show4HeaderAbbrBorder = new Border
        {
            CornerRadius      = new CornerRadius(3),
            Padding           = new Thickness(5, 1),
            VerticalAlignment = VerticalAlignment.Center,
            Child             = _show4HeaderAbbrTb,
            IsVisible         = false,
        };
        headerLine.Children.Add(_show4HeaderAbbrBorder);
        UpdateShow4AbbrBadge(abbr);
        root.Children.Add(headerLine);

        _show4ItemCountTb = new TextBlock
        {
            Text       = $"({prices.Count}개 항목)",
            FontSize   = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            Margin     = new Thickness(0, 0, 0, 8),
        };
        root.Children.Add(_show4ItemCountTb);

        // ─── 상단 컨트롤 (모두 선택 | 전체 해제 | 비교 계약) ───────────────────
        var topControlPanel = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto,Auto"),
            ColumnSpacing     = 8,
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Thickness(0, 0, 0, 8),
        };

        var leftGroup = new StackPanel
        {
            Orientation       = Orientation.Horizontal,
            Spacing           = 16,
            VerticalAlignment = VerticalAlignment.Center,
        };
        _show4SelectAllChk = new CheckBox
        {
            Content   = "모두 선택",
            FontSize  = AppTheme.FontMD, FontFamily = Font,
            IsChecked = !uncheckedByDefault,
        };
        leftGroup.Children.Add(_show4SelectAllChk);

        var btnClearAll = new Button
        {
            Content         = "❌  전체 해제",
            Height          = 34,
            FontSize        = AppTheme.FontSM, FontFamily = Font,
            Background      = AppTheme.BgDanger, Foreground = AppTheme.FgDanger,
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(10, 0),
        };
        leftGroup.Children.Add(btnClearAll);

        topControlPanel.Children.Add(leftGroup);  Grid.SetColumn(leftGroup, 0);
        topControlPanel.Children.Add(lblCompare); Grid.SetColumn(lblCompare, 2);
        topControlPanel.Children.Add(cbContract); Grid.SetColumn(cbContract, 3);
        root.Children.Add(topControlPanel);

        // ─── 컬럼 헤더 ──────────────────────────────────────────────────────────
        var headerGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,*,90,70,90"),
            ColumnSpacing     = 8,
        };
        var hdrDefs = new[] { ("", 0), ("항목명", 1), ("단가", 2), ("수량", 3), ("소계", 4) };
        foreach (var (txt, col) in hdrDefs)
        {
            var align = col <= 1 ? Avalonia.Media.TextAlignment.Left : Avalonia.Media.TextAlignment.Right;
            var hdr = new TextBlock { Text = txt, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppTheme.FgDimmed, TextAlignment = align };
            Grid.SetColumn(hdr, col);
            headerGrid.Children.Add(hdr);
        }
        root.Children.Add(new Border
        {
            Background = AppTheme.BgSecondary ?? AppTheme.BgPrimary,
            Padding    = new Thickness(8, 4),
            Child      = headerGrid,
        });

        // ─── 합계 ──────────────────────────────────────────────────────────────
        _show4TotalSummaryTb = new TextBlock
        {
            FontSize   = AppTheme.FontMD, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgSuccess,
            Margin     = new Thickness(0, 4, 0, 4),
        };
        root.Children.Add(_show4TotalSummaryTb);

        // ─── 카테고리/항목 행 ──────────────────────────────────────────────────
        var allItems  = _cachedAllItems ??= AnalysisService.GetAllItems();
        var priceDict = prices.ToDictionary(p => p.Analyte, p => p.Price, StringComparer.OrdinalIgnoreCase);
        var catOrder  = allItems
            .Select(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .Distinct().ToList();
        if (catOrder.Remove("일반항목")) catOrder.Insert(0, "일반항목");

        var itemMeta = allItems
            .GroupBy(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
            .ToDictionary(g => g.Key, g => g.ToList());

        var knownAnalytes = allItems.Select(a => a.Analyte).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var unknownItems  = prices
            .Where(p => !string.IsNullOrWhiteSpace(p.Price) && !knownAnalytes.Contains(p.Analyte))
            .Select(p => new AnalysisItem { Analyte = p.Analyte, Category = "기타", 약칭 = "", ES = "zzz" })
            .ToList();
        if (unknownItems.Count > 0 && !itemMeta.ContainsKey("기타"))
        {
            itemMeta["기타"] = new List<AnalysisItem>();
            if (!catOrder.Contains("기타")) catOrder.Add("기타");
        }

        var builtAnalytes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var itemsPanel    = new StackPanel { Spacing = 4 };
        decimal grandTotal = 0m;

        foreach (var catKey in catOrder.Where(k => itemMeta.ContainsKey(k)))
        {
            var groupItems = itemMeta[catKey].OrderBy(a => a.ES).ToList();
            if (catKey == "기타") groupItems.AddRange(unknownItems);
            if (groupItems.Count == 0) continue;

            // ── 카테고리 헤더 ──────────────────────────────────────────────
            var catChk = new CheckBox
            {
                Content    = $"전체 ({groupItems.Count})",
                IsChecked  = !uncheckedByDefault,
                FontSize   = AppTheme.FontXS, FontFamily = Font,
                Foreground = Brush.Parse("#88bb88"),
                Padding    = new Thickness(2, 0), Margin = new Thickness(0),
                MinHeight  = 18,
            };
            var chevron = new TextBlock
            {
                Text              = "▾", FontSize = 10,
                Foreground        = AppTheme.FgMuted,
                VerticalAlignment = VerticalAlignment.Center,
                Margin            = new Thickness(0, 0, 4, 0),
            };
            var headerRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
            headerRow.Children.Add(chevron); Grid.SetColumn(chevron, 0);
            var catLabel = new TextBlock
            {
                Text              = catKey,
                FontSize          = AppTheme.FontSM, FontFamily = Font,
                FontWeight        = FontWeight.SemiBold,
                Foreground        = AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
            };
            headerRow.Children.Add(catLabel); Grid.SetColumn(catLabel, 1);
            headerRow.Children.Add(catChk);   Grid.SetColumn(catChk, 2);

            var headerBorder = new Border
            {
                Background   = AppTheme.BgSecondary ?? AppTheme.BgPrimary,
                Padding      = new Thickness(8, 5),
                Margin       = new Thickness(0, 0, 0, 1),
                CornerRadius = new CornerRadius(4),
                Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child        = headerRow,
            };

            // ── 항목 목록 패널 ─────────────────────────────────────────────
            var itemsBorder = new Border
            {
                Background   = AppTheme.BgPrimary,
                Padding      = new Thickness(8, 4, 4, 4),
                Margin       = new Thickness(0, 0, 0, 4),
                CornerRadius = new CornerRadius(0, 0, 4, 4),
                IsVisible    = true,
            };
            var itemsStack = new StackPanel { Spacing = 0 };
            itemsBorder.Child = itemsStack;

            var catCheckBoxes = new List<CheckBox>();

            foreach (var item in groupItems)
            {
                var analyte = item.Analyte;
                builtAnalytes.Add(analyte);

                bool hasRealPrice = priceDict.TryGetValue(analyte, out var priceStr)
                    && !string.IsNullOrWhiteSpace(priceStr);
                bool hasRealQuantity = contractQtys.TryGetValue(analyte, out var rawQtyStr)
                    && !string.IsNullOrWhiteSpace(rawQtyStr);

                decimal priceValue = 0m;
                if (hasRealPrice) decimal.TryParse(priceStr, out priceValue);

                int contractQty;
                if (hasRealQuantity && int.TryParse(rawQtyStr, out int cqInt) && cqInt > 0)
                    contractQty = cqInt;
                else if (hasRealPrice)
                    contractQty = 1;
                else
                    contractQty = 0;

                decimal initSubtotal = hasRealPrice ? priceValue * contractQty : 0m;
                if (initSubtotal > 0) _show4Subtotals[analyte] = initSubtotal;
                grandTotal += initSubtotal;

                var rowGrid = new Grid
                {
                    ColumnDefinitions = new ColumnDefinitions("Auto,*,90,70,90"),
                    ColumnSpacing     = 8,
                };

                var chk = new CheckBox
                {
                    IsChecked         = !uncheckedByDefault && hasRealPrice,
                    VerticalAlignment = VerticalAlignment.Center,
                };
                Grid.SetColumn(chk, 0);
                rowGrid.Children.Add(chk);
                _show4AllCheckBoxes.Add(chk);
                catCheckBoxes.Add(chk);
                _show4CheckBoxes[analyte] = chk;

                var analyteAlias = string.IsNullOrEmpty(item.약칭) ? (aliasMap.TryGetValue(analyte, out var al) ? al : "") : item.약칭;
                var (abgColor, afgColor) = BadgeColorHelper.GetBadgeColor(analyteAlias);
                var analyteBadge = new Border
                {
                    Background        = Brush.Parse(abgColor),
                    BorderBrush       = Brush.Parse(afgColor),
                    BorderThickness   = new Thickness(1),
                    CornerRadius      = new CornerRadius(10),
                    Padding           = new Thickness(6, 1, 8, 1),
                    Margin            = new Thickness(0, 0, 4, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text              = string.IsNullOrEmpty(analyteAlias) ? "✓" : analyteAlias,
                        FontSize          = AppTheme.FontSM, FontFamily = Font,
                        FontWeight        = FontWeight.Medium,
                        Foreground        = Brush.Parse(afgColor),
                        VerticalAlignment = VerticalAlignment.Center,
                    }
                };
                var namePanel = new StackPanel
                {
                    Orientation       = Orientation.Horizontal,
                    Spacing           = 0,
                    VerticalAlignment = VerticalAlignment.Center,
                    Children = { analyteBadge, new TextBlock
                    {
                        Text              = analyte,
                        FontSize          = AppTheme.FontSM, FontFamily = Font,
                        Foreground        = AppTheme.FgPrimary,
                        VerticalAlignment = VerticalAlignment.Center,
                        TextTrimming      = TextTrimming.CharacterEllipsis,
                    }}
                };
                Grid.SetColumn(namePanel, 1);
                rowGrid.Children.Add(namePanel);

                var priceBox = new TextBox
                {
                    Text              = priceValue > 0 ? ((long)priceValue).ToString() : "",
                    FontSize          = AppTheme.FontSM, FontFamily = Font,
                    Foreground        = AppTheme.FgSuccess,
                    FontWeight        = FontWeight.SemiBold,
                    TextAlignment     = Avalonia.Media.TextAlignment.Right,
                    Height            = 28,
                    Padding           = new Thickness(4, 2),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                Grid.SetColumn(priceBox, 2);
                rowGrid.Children.Add(priceBox);
                _show4PriceTextBoxes[analyte] = priceBox;

                var qtyBox = new TextBox
                {
                    Text              = contractQty > 0 ? contractQty.ToString() : "",
                    FontSize          = AppTheme.FontSM, FontFamily = Font,
                    TextAlignment     = Avalonia.Media.TextAlignment.Right,
                    Height            = 28,
                    Padding           = new Thickness(4, 2),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                _show4QtyTextBoxes[analyte] = qtyBox;
                Grid.SetColumn(qtyBox, 3);
                rowGrid.Children.Add(qtyBox);

                var subtotalTb = new TextBlock
                {
                    Text              = initSubtotal > 0 ? initSubtotal.ToString("N0") : "—",
                    FontSize          = AppTheme.FontSM, FontFamily = Font,
                    Foreground        = AppTheme.FgSuccess,
                    FontWeight        = FontWeight.SemiBold,
                    TextAlignment     = Avalonia.Media.TextAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                };
                Grid.SetColumn(subtotalTb, 4);
                rowGrid.Children.Add(subtotalTb);
                _show4SubtotalTbs[analyte] = subtotalTb;

                priceBox.TextChanged += (_, _) =>
                {
                    if (_suppressTotalUpdate) return;
                    decimal pv = 0m;
                    if (decimal.TryParse(priceBox.Text, out decimal np))
                    {
                        pv = np;
                        _pendingPrices[analyte] = np.ToString();
                        if (np > 0) chk.IsChecked = true;
                    }
                    if (int.TryParse(qtyBox.Text, out int nq) && nq > 0)
                    {
                        var sub = pv * nq;
                        _show4Subtotals[analyte] = sub;
                        subtotalTb.Text = sub.ToString("N0");
                    }
                    else
                    {
                        _show4Subtotals.Remove(analyte);
                        subtotalTb.Text = "—";
                    }
                    RecomputeShow4Total();
                };
                qtyBox.TextChanged += (_, _) =>
                {
                    if (_suppressTotalUpdate) return;
                    _pendingQuantities[analyte] = qtyBox.Text ?? "";
                    decimal pv = decimal.TryParse(priceBox.Text, out var p) ? p : 0m;
                    if (int.TryParse(qtyBox.Text, out int nq) && nq > 0)
                    {
                        var sub = pv * nq;
                        _show4Subtotals[analyte] = sub;
                        subtotalTb.Text = sub.ToString("N0");
                        chk.IsChecked = true;
                    }
                    else
                    {
                        _show4Subtotals.Remove(analyte);
                        subtotalTb.Text = "—";
                    }
                    RecomputeShow4Total();
                };
                chk.IsCheckedChanged += (_, _) =>
                {
                    if (_suppressTotalUpdate) return;
                    if (chk.IsChecked == false)
                    {
                        priceBox.Text = "";
                        qtyBox.Text = "";
                        _show4Subtotals.Remove(analyte);
                        subtotalTb.Text = "—";
                        _pendingPrices[analyte] = "";
                        _pendingQuantities[analyte] = "";
                        RecomputeShow4Total();
                    }
                };

                itemsStack.Children.Add(new Border
                {
                    Padding         = new Thickness(0, 4),
                    BorderBrush     = AppTheme.BorderSubtle,
                    BorderThickness = new Thickness(0, 0, 0, 1),
                    Child           = rowGrid,
                });
            }

            var capturedCatChk = catChk;
            var capturedCatCbs = catCheckBoxes;
            catChk.IsCheckedChanged += (_, _) =>
            {
                if (_suppressTotalUpdate) return;
                var v = capturedCatChk.IsChecked ?? false;
                foreach (var cb in capturedCatCbs) cb.IsChecked = v;
            };

            var capturedItems   = itemsBorder;
            var capturedChevron = chevron;
            headerBorder.PointerPressed += (_, _) =>
            {
                capturedItems.IsVisible = !capturedItems.IsVisible;
                capturedChevron.Text    = capturedItems.IsVisible ? "▾" : "▸";
            };

            itemsPanel.Children.Add(headerBorder);
            itemsPanel.Children.Add(itemsBorder);
        }

        _show4SelectAllChk.IsCheckedChanged += (_, _) =>
        {
            if (_suppressTotalUpdate) return;
            var v = _show4SelectAllChk.IsChecked ?? false;
            foreach (var chk in _show4AllCheckBoxes) chk.IsChecked = v;
        };

        btnClearAll.Click += (_, _) =>
        {
            foreach (var chk in _show4AllCheckBoxes) chk.IsChecked = false;
            if (_show4SelectAllChk != null) _show4SelectAllChk.IsChecked = false;
        };

        _show4TotalSummaryTb.Text = $"💵 합계: {grandTotal.ToString("N0")} 원";
        // NOTE: 최초 구조 빌드 시 Show2 계약금액 자동 덮어쓰기 제거 (#6) — DB 원본 유지.
        //       RecomputeShow4Total (사용자 편집)만 Show2에 반영한다.

        _show4KnownAnalytes = builtAnalytes;

        var scroll = new ScrollViewer
        {
            Content = itemsPanel,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Margin = new Thickness(0, 0, 0, 8),
        };
        root.Children.Add(scroll);

        return new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // =========================================================================
    // Show3 — 진행률 계산 전 placeholder (버튼 + 단가복사)
    //   구조는 최초 1회만 빌드하고 업체 변경 시엔 캡처된 필드(_show3PlaceholderContract/Qtys)만 갱신
    // =========================================================================
    private Control BuildProgressPlaceholder(Contract captured,
                                             Dictionary<string, string> contractQtys)
    {
        _show3PlaceholderContract = captured;
        _show3PlaceholderQtys     = contractQtys;

        if (_show3CachedPlaceholder != null && _show3CalcBtn != null)
        {
            // 버튼 상태 원복 (이전 업체 계산이 끝난 경우 대비)
            _show3CalcBtn.IsEnabled = true;
            _show3CalcBtn.Content   = "📊  진행률 계산";
            return _show3CachedPlaceholder;
        }

        var root = new StackPanel { Spacing = 12, Margin = new Thickness(8) };

        _show3CalcBtn = new Button
        {
            Content         = "📊  진행률 계산",
            FontSize        = AppTheme.FontSM,
            FontFamily      = Font, FontWeight = FontWeight.SemiBold,
            Background      = AppTheme.BgActiveGreen,
            Foreground      = AppTheme.FgSuccess,
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(12, 6),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        _show3HintTb = new TextBlock
        {
            Text       = "버튼을 누르면 항목별 처리수량을 조회하여 진행률을 표시합니다.",
            FontSize   = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
        };

        _show3CalcBtn.Click += async (_, _) =>
        {
            var curCaptured = _show3PlaceholderContract;
            var curQtys     = _show3PlaceholderQtys;
            if (curCaptured == null || curQtys == null) return;

            _show3CalcBtn.IsEnabled = false;
            _show3CalcBtn.Content   = "⏳  계산 중...";
            try
            {
                var procSw = System.Diagnostics.Stopwatch.StartNew();
                var processedQtys = await Task.Run(() => ContractService.GetProcessedQuantities(
                    curCaptured.C_CompanyName, curCaptured.C_ContractStart, curCaptured.C_ContractEnd));
                Log($"  → GetProcessedQuantities(클릭): {procSw.ElapsedMilliseconds}ms, {processedQtys.Count}개");

                if (_selectedContract?.C_CompanyName != curCaptured.C_CompanyName) return;

                var swBuild = System.Diagnostics.Stopwatch.StartNew();
                var progressPanel = BuildProgressPanel(curCaptured.C_CompanyName, curQtys, processedQtys);
                Log($"[타이밍] BuildProgressPanel(클릭): {swBuild.ElapsedMilliseconds}ms");

                EditPanelChanged?.Invoke(progressPanel);
            }
            catch (Exception ex)
            {
                Log($"[Show3] 진행률 계산 오류: {ex.Message}");
                _show3CalcBtn.IsEnabled = true;
                _show3CalcBtn.Content   = "📊  진행률 계산 (재시도)";
            }
        };

        root.Children.Add(_show3CalcBtn);
        root.Children.Add(_show3HintTb);

        _show3CachedPlaceholder = new ScrollViewer { Content = root };
        return _show3CachedPlaceholder;
    }

    // =========================================================================
    // Show3 — 항목별 계약수량 vs 처리수량 진행상황 (휠스크롤 가능)
    // =========================================================================
    private Control BuildProgressPanel(string companyName,
                                       Dictionary<string, string> contractQtys,
                                       Dictionary<string, int> processedQtys)
    {
        Log($"[UI빌드] BuildProgressPanel({companyName}) 시작, 계약수량={contractQtys.Count}, 처리수량={processedQtys.Count}");
        var swProg = System.Diagnostics.Stopwatch.StartNew();

        var scrollViewer = new ScrollViewer();

        var root = new StackPanel { Spacing = 8, Margin = new Thickness(8) };
        scrollViewer.Content = root;

        var items = ContractService.GetAnalysisItems();
        if (items.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text       = "분석 항목이 없습니다.",
                FontSize   = AppTheme.FontBase, FontFamily = Font,
                Foreground = AppTheme.FgMuted,
            });
            return scrollViewer;
        }

        // ── 헤더: 총계값 표시 ──
        var headerPanel = new StackPanel { Spacing = 8, Margin = new Thickness(0, 0, 0, 8) };

        // 계약수량 합계, 처리수량 합계, 진행률 합계
        int totalContractQty = 0;
        int totalProcessedQty = 0;
        foreach (var item in items)
        {
            string contractQtyStr = (contractQtys.TryGetValue(item, out var cq) ? cq : "0") ?? "0";
            int contractQty = int.TryParse(contractQtyStr, out int c) ? c : 0;
            totalContractQty += contractQty;

            int processedQty = (processedQtys.TryGetValue(item, out var pq) ? pq : 0);
            totalProcessedQty += processedQty;
        }

        double totalProgress = totalContractQty > 0 ? (double)totalProcessedQty / totalContractQty * 100 : 0;

        var summaryLine = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
        summaryLine.Children.Add(new TextBlock
        {
            Text       = "📊  총계",
            FontSize   = AppTheme.FontMD, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
            VerticalAlignment = VerticalAlignment.Center,
        });
        summaryLine.Children.Add(new TextBlock
        {
            Text       = $"계약수량: {totalContractQty}",
            FontSize   = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        });
        summaryLine.Children.Add(new TextBlock
        {
            Text       = $"처리수량: {totalProcessedQty}",
            FontSize   = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgSuccess,
            VerticalAlignment = VerticalAlignment.Center,
        });
        summaryLine.Children.Add(new TextBlock
        {
            Text       = $"진행률: {totalProgress:F1}%",
            FontSize   = AppTheme.FontBase, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = totalProgress >= 100 ? AppTheme.FgSuccess : AppTheme.FgWarn,
            VerticalAlignment = VerticalAlignment.Center,
        });
        headerPanel.Children.Add(summaryLine);
        headerPanel.Children.Add(new Border { Height = 1, Background = AppTheme.BorderMuted });

        root.Children.Add(headerPanel);

        // ── 항목별 진행상황 그리드 (4컬럼: 진행률·바 통합) ──
        const int COLS = 4;
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("2*,1*,1*,3*"),  // 항목명, 계약수량, 처리수량, 진행률(바)
            ColumnSpacing     = 6,
            RowSpacing        = 4,
        };

        // 헤더 행
        var headerCols = new[] { "항목명", "계약수량", "처리수량", "진행률" };
        for (int i = 0; i < COLS; i++)
        {
            var header = new TextBlock
            {
                Text              = headerCols[i],
                FontSize          = AppTheme.FontXS, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground        = AppTheme.FgMuted,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            Grid.SetColumn(header, i);
            Grid.SetRow(header, 0);
            grid.Children.Add(header);
        }

        // 행 정의 (헤더 1 + 항목들)
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        for (int i = 0; i < items.Count; i++)
            grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        // 데이터 행들
        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];
            string contractQtyStr = (contractQtys.TryGetValue(item, out var cq) ? cq : "0") ?? "0";
            int contractQty = int.TryParse(contractQtyStr, out int c) ? c : 0;

            int processedQty = (processedQtys.TryGetValue(item, out var pq) ? pq : 0);
            double progress = contractQty > 0 ? (double)processedQty / contractQty * 100 : 0;

            int row = i + 1;

            // 항목명
            var itemName = new TextBlock
            {
                Text              = item,
                FontSize          = AppTheme.FontXS, FontFamily = Font,
                Foreground        = AppTheme.FgPrimary,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
            };
            Grid.SetColumn(itemName, 0);
            Grid.SetRow(itemName, row);
            grid.Children.Add(itemName);

            // 계약수량
            var contractQtyBlock = new TextBlock
            {
                Text              = contractQty.ToString(),
                FontSize          = AppTheme.FontXS, FontFamily = Font,
                Foreground        = AppTheme.FgInfo,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            Grid.SetColumn(contractQtyBlock, 1);
            Grid.SetRow(contractQtyBlock, row);
            grid.Children.Add(contractQtyBlock);

            // 처리수량
            var processedQtyBlock = new TextBlock
            {
                Text              = processedQty.ToString(),
                FontSize          = AppTheme.FontXS, FontFamily = Font,
                Foreground        = AppTheme.FgSuccess,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
            };
            Grid.SetColumn(processedQtyBlock, 2);
            Grid.SetRow(processedQtyBlock, row);
            grid.Children.Add(processedQtyBlock);

            // 진행률 — 바 위에 % 텍스트 오버레이 (한 컬럼으로 통합)
            var progressBar = new ProgressBar
            {
                Value   = progress,
                Minimum = 0,
                Maximum = 100,
                Height  = 18,
                Foreground = progress >= 100 ? new SolidColorBrush(Color.Parse("#22C55E"))
                           : progress >= 50  ? new SolidColorBrush(Color.Parse("#F59E0B"))
                                             : new SolidColorBrush(Color.Parse("#EF4444")),
                Background = AppTheme.BgSecondary,
            };
            var percentLabel = new TextBlock
            {
                Text              = $"{progress:F1}%",
                FontSize          = AppTheme.FontXS, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground        = AppTheme.FgPrimary,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            };
            var progressCell = new Grid
            {
                Height = 24,
                Background = AppTheme.BgSecondary,
                VerticalAlignment = VerticalAlignment.Center,
            };
            progressCell.Children.Add(progressBar);
            progressCell.Children.Add(percentLabel);
            var progressBarContainer = new Border
            {
                Child = progressCell,
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(2),
            };
            Grid.SetColumn(progressBarContainer, 3);
            Grid.SetRow(progressBarContainer, row);
            grid.Children.Add(progressBarContainer);
        }

        root.Children.Add(grid);
        Log($"[UI빌드] BuildProgressPanel 완료: {swProg.ElapsedMilliseconds}ms");
        return scrollViewer;
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
        var txbStart  = AddGridRow(grid, 2, "계약시작",         "", hint: "예) 20240101");
        var txbEnd    = AddGridRow(grid, 3, "계약종료",         "", hint: "예) 20241231");
        var txbDays   = AddGridRow(grid, 4, "계약일수",         "", isReadOnly: true);
        AddGridRow(grid, 5, "계약금액(VAT별도)", "");
        AddGridRow(grid, 6, "주소",             "");
        AddGridRow(grid, 7, "대표자",           "");
        AddGridRow(grid, 8, "시설별",           "");
        AddGridRow(grid, 9, "종류별",           "");
        AddGridRow(grid,10, "주생산품",         "");
        AddGridRow(grid,11, "담당자",           "");
        AddGridRow(grid,12, "연락처",           "");
        AddGridRow(grid,13, "이메일",           "");

        // [5][6] 날짜 자동 포맷 + 계약일수 자동 계산
        AttachDateAutoFormat(txbStart, txbEnd, txbDays);

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
        if (ok) _cachedAllContracts = null;
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
               : $"❌ ��가 실패: {newContract.C_CompanyName}");

        if (ok)
        {
            _cachedAllContracts = null;
            // pending 단가/수량 저장 (Show4에서 선택한 항목)
            if (_pendingPrices.Count > 0)
            {
                try
                {
                    var priceList = _pendingPrices.Select(kv => (kv.Key, kv.Value)).ToList();
                    ContractService.UpdateContractPrices(newContract.C_CompanyName, priceList);
                    Log($"✅ 단가 저장: {newContract.C_CompanyName} ({priceList.Count}항목)");
                }
                catch (Exception ex) { Log($"❌ 단가 저장 실패: {ex.Message}"); }
            }

            if (_pendingQuantities.Count > 0)
            {
                try
                {
                    var qtyList = _pendingQuantities.Select(kv => (kv.Key, kv.Value)).ToList();
                    ContractService.UpdateContractQuantities(newContract.C_CompanyName, qtyList);
                    Log($"✅ 수량 저장: {newContract.C_CompanyName} ({qtyList.Count}항목)");
                }
                catch (Exception ex) { Log($"❌ 수량 저장 실패: {ex.Message}"); }
            }

            ContractTreeView.Items.Add(CreateTreeItem(newContract));
            _isAddMode           = false;
            _templateCompanyName = null;
            _pendingPrices.Clear();
            _pendingQuantities.Clear();
            _detailPanel         = null;
            DetailPanelChanged?.Invoke(null);
            StatsPanelChanged?.Invoke(null);

            // Show1 즉시 리프레시 (트리뷰 선택 해제 → 새 업체 자동 표시)
            ContractTreeView.SelectedItem = ContractTreeView.Items.Cast<TreeViewItem>().LastOrDefault();
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

            // ComboBox 처리 (계약근거 / 처리시설)
            var cb = child.Children[1] as ComboBox;
            if (cb != null)
            {
                if (label == "계약근거" && cb.SelectedItem is string selectedBasis)
                {
                    // 전체텍스트 포맷: "계약번호 / 업체명 / 대표사업장 / 기간"
                    // 구 포맷 호환: "계약번호_업체명" (언더바 분리)
                    string no;
                    if (selectedBasis.Contains(" / "))
                        no = selectedBasis.Split('/')[0].Trim();
                    else
                        no = selectedBasis.Split('_')[0].Trim();
                    c.C_ContractType = no;
                }
                else if (label == "측정대상 사업장" && cb.SelectedItem is string selectedPlace)
                {
                    c.C_PlaceName = selectedPlace;
                }
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

    /// <summary>Show2 패널에서 계약금액 TextBox를 찾아 업데이트</summary>
    private void UpdateContractAmountInShow2(StackPanel show2Panel, decimal amount)
    {
        if (_contractAmountTextBox != null)
        {
            var newVal = amount > 0 ? amount.ToString("N0") : "";
            if (_contractAmountTextBox.Text == newVal) return;
            _contractAmountTextBox.Text = newVal;
            Log($"[Show2금액] 직접 업데이트 (저장된 TextBox) 적용금액={amount:N0}");
            return;
        }

        var grid = show2Panel.Children.OfType<Grid>().FirstOrDefault();
        if (grid == null)
        {
            Log("[Show2금액] Grid 없음 — show2Panel에 Grid가 없습니다");
            return;
        }

        var stackPanels = grid.Children.OfType<StackPanel>().ToList();
        Log($"[Show2금액] StackPanel 개수={stackPanels.Count}, 적용금액={amount:N0}");

        foreach (var child in stackPanels)
        {
            if (child.Children.Count < 2) continue;

            var labelBlock = child.Children[0] as TextBlock;
            var rawLabel = labelBlock?.Text ?? "";
            var label = rawLabel.Replace("🔒 ", "").Replace("    ", "").Replace(" :", "").Trim();

            Log($"[Show2금액] 라벨확인: '{label}' (원문='{rawLabel}')");

            if (label == "계약금액(VAT별도)")
            {
                if (child.Children[1] is TextBox txb)
                {
                    var newVal = amount > 0 ? amount.ToString("N0") : "";
                    if (txb.Text == newVal) return;  // 동일값이면 스킵 (무한루프 방지)
                    txb.Text = newVal;
                    return;
                }
                else
                {
                    Log($"[Show2금액] ❌ 계약금액 라벨은 찾았으나 TextBox가 아님: {child.Children[1]?.GetType().Name}");
                    return;
                }
            }
        }

        Log("[Show2금액] ❌ '계약금액(VAT별도)' 라벨을 가진 StackPanel을 찾지 못했습니다");
    }

    /// <summary>[5][6] 날짜 TextBox에 자동 포맷(20260504→2026-05-04) + 계약일수 자동 계산 연결</summary>
    private static void AttachDateAutoFormat(TextBox txbStart, TextBox txbEnd, TextBox txbDays)
    {
        static string FormatDateInput(string raw)
        {
            var digits = new string(raw.Where(char.IsDigit).ToArray());
            if (digits.Length == 8)
                return $"{digits[..4]}-{digits[4..6]}-{digits[6..8]}";
            return raw;
        }

        static void RecalcDays(TextBox s, TextBox e, TextBox d)
        {
            var sf = FormatDateInput(s.Text ?? "");
            var ef = FormatDateInput(e.Text ?? "");
            if (DateTime.TryParse(sf, out var sd) && DateTime.TryParse(ef, out var ed))
                d.Text = (ed.Date - sd.Date).Days.ToString();
            else
                d.Text = "";
        }

        txbStart.LostFocus += (_, _) =>
        {
            txbStart.Text = FormatDateInput(txbStart.Text ?? "");
            RecalcDays(txbStart, txbEnd, txbDays);
        };

        txbEnd.LostFocus += (_, _) =>
        {
            txbEnd.Text = FormatDateInput(txbEnd.Text ?? "");
            RecalcDays(txbStart, txbEnd, txbDays);
        };
    }

    /// <summary>[1] 기존 계약 선택 상태에서 단가 복사 → Add 모드로 진입</summary>
    private void ShowAddModeTemplateSelectionWithCompany(string companyName)
    {
        // Add 모드가 아니면 진입
        if (!_isAddMode)
        {
            ShowAddPanel();  // Show2에 신규 추가 패널
        }
        ShowAddModeAnalyteSelection(companyName);
    }

    // =========================================================================
    // UI 헬퍼
    // =========================================================================
    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    private static StackPanel MakeRootPanel(string title)
    {
        var root = new StackPanel { Spacing = 10, Margin = new Thickness(4) };
        root.Children.Add(new TextBlock
        {
            Text       = title,
            FontSize   = AppTheme.FontXL,
            FontFamily = Font,
            FontWeight = FontWeight.SemiBold,
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

    private static TextBox AddGridRow(Grid grid, int row, string label, string value,
                                   bool isReadOnly = false, bool isLocked = false,
                                   string hint = "")
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };

        panel.Children.Add(new TextBlock
        {
            Text              = (isLocked ? "🔒 " : "") + label + " :",
            Width             = 90,
            FontSize          = AppTheme.FontMD,
            FontFamily        = Font,
            Foreground        = isLocked
                                    ? AppTheme.FgMuted
                                    : AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            TextAlignment     = Avalonia.Media.TextAlignment.Right,
        });

        var textBox = new TextBox
        {
            Text            = value ?? "",
            Width           = 200,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            IsReadOnly      = isReadOnly,
            Watermark       = hint,
            Background      = isReadOnly
                                  ? new SolidColorBrush(Color.Parse("#252525"))
                                  : AppTheme.BorderSeparator,
            Foreground      = isReadOnly
                                  ? AppTheme.BorderMuted
                                  : AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = isReadOnly
                                  ? AppTheme.BorderSubtle
                                  : AppTheme.BorderDefault,
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(8, 4),
        };
        panel.Children.Add(textBox);

        int col     = row % 2;
        int gridRow = row / 2;
        Grid.SetColumn(panel, col);
        Grid.SetRow(panel, gridRow);
        grid.Children.Add(panel);
        return textBox;
    }

    // =========================================================================
    // ★ 계약구분 ComboBox 행 추가
    //   - 분석단가 테이블의 FS100 이후 컬럼을 항목으로 채움
    //   - 기본 선택값 : FS25 (없으면 첫 번째 항목)
    //   - 선택 변경 시 : Show3 에 해당 컬럼 데이터 테이블 표시
    // =========================================================================
    private void AddGridRowComboBox(Grid grid, int row, string label, string currentValue)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };

        panel.Children.Add(new TextBlock
        {
            Text              = "    " + label + " :",
            Width             = 140,
            FontSize          = AppTheme.FontMD,
            FontFamily        = Font,
            Foreground        = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
            TextAlignment     = Avalonia.Media.TextAlignment.Right,
        });

        // 계약구분 목록 (계약 DB C_ContractType 고유값, 없으면 기본값)
        List<string> columns;
        try
        {
            columns = QuotationService.GetContractTypes();
        }
        catch
        {
            columns = new List<string> { "용역", "구매", "기타" };
        }

        var comboBox = new ComboBox
        {
            Width           = 200,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            Background      = AppTheme.BorderSeparator,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderDefault,
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

        panel.Children.Add(comboBox);

        int col2    = row % 2;
        int gridRow = row / 2;
        Grid.SetColumn(panel, col2);
        Grid.SetRow(panel, gridRow);
        grid.Children.Add(panel);

    }

    // 계약근거 + 처리시설 ComboBox 추가 (측정인계약/측정인처리시설 DB 기준)
    // - 계약근거 → 스크랩된 add_meas_cont_no 전체텍스트 (ex: "1000497193 / 보임.../ .../ 2026-...")
    // - 처리시설 → 선택된 계약번호에 속한 cmb_emis_cmpy_plc_no 옵션 텍스트
    // - row: 계약근거 행, row+1: 처리시설 행
    private void AddBasisContractComboBox(Grid grid, int row, string currentContractNo, string currentPlaceName)
    {
        StackPanel MakeRow(string label)
        {
            var p = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            p.Children.Add(new TextBlock
            {
                Text              = label + " :",
                Width             = 90,
                FontSize          = AppTheme.FontMD,
                FontFamily        = Font,
                Foreground        = AppRes("FgMuted"),
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment     = Avalonia.Media.TextAlignment.Right,
            });
            return p;
        }

        ComboBox MakeCombo() => new ComboBox
        {
            Width           = 200,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            Background      = AppTheme.BorderSeparator,
            Foreground      = AppRes("AppFg"),
            BorderThickness = new Thickness(1),
            BorderBrush     = AppTheme.BorderDefault,
            Padding         = new Thickness(8, 4),
        };

        var basisPanel = MakeRow("계약근거");
        var placePanel = MakeRow("측정대상 사업장");
        var basisCombo = MakeCombo();
        var placeCombo = MakeCombo();
        basisPanel.Children.Add(basisCombo);
        placePanel.Children.Add(placeCombo);
        _show2BasisCombo = basisCombo;
        _show2PlaceCombo = placeCombo;

        // 계약근거: 측정인계약 전체텍스트 옵션
        List<MeasurerService.MeasurerContract> contracts = new();
        try
        {
            bool measCacheMiss = _cachedMeasurerContracts == null;
            contracts = _cachedMeasurerContracts ??= MeasurerService.GetAllMeasurerContracts();
            Log($"[ComboBox] GetAllMeasurerContracts ({(measCacheMiss ? "miss" : "hit")}): {contracts.Count}건");
            foreach (var c in contracts)
            {
                string label = string.IsNullOrWhiteSpace(c.전체텍스트)
                    ? $"{c.계약번호} / {c.업체명}"
                    : c.전체텍스트;
                basisCombo.Items.Add(label);
            }
            if (!string.IsNullOrEmpty(currentContractNo))
            {
                var match = contracts.FirstOrDefault(x => x.계약번호 == currentContractNo);
                if (!string.IsNullOrEmpty(match.계약번호))
                {
                    string lbl = string.IsNullOrWhiteSpace(match.전체텍스트)
                        ? $"{match.계약번호} / {match.업체명}" : match.전체텍스트;
                    basisCombo.SelectedItem = lbl;
                }
            }
        }
        catch (Exception ex) { Log($"[ComboBox] 계약근거 로드 오류: {ex.Message}"); }

        // 처리시설 옵션 갱신 (계약번호 기반)
        void RefreshPlaces(string 계약번호, string? preserveName)
        {
            placeCombo.Items.Clear();
            if (string.IsNullOrWhiteSpace(계약번호)) return;
            try
            {
                var places = MeasurerService.GetMeasurerPlaces(계약번호);
                foreach (var p in places) placeCombo.Items.Add(p.처리시설명);
                if (!string.IsNullOrEmpty(preserveName) &&
                    places.Any(p => p.처리시설명 == preserveName))
                    placeCombo.SelectedItem = preserveName;
            }
            catch (Exception ex) { Log($"[ComboBox] 처리시설 로드 오류: {ex.Message}"); }
        }
        _show2RefreshPlaces = RefreshPlaces;

        RefreshPlaces(currentContractNo, currentPlaceName);

        basisCombo.SelectionChanged += (_, _) =>
        {
            if (_suppressShow2Sync) return;
            if (basisCombo.SelectedItem is not string sel) return;
            var no = sel.Split('/')[0].Trim();
            if (_selectedContract != null) _selectedContract.C_ContractType = no;
            RefreshPlaces(no, null);
        };

        placeCombo.SelectionChanged += (_, _) =>
        {
            if (_suppressShow2Sync) return;
            if (placeCombo.SelectedItem is string sp && _selectedContract != null)
                _selectedContract.C_PlaceName = sp;
        };

        int col1     = row % 2;
        int gridRow1 = row / 2;
        Grid.SetColumn(basisPanel, col1);
        Grid.SetRow(basisPanel, gridRow1);
        grid.Children.Add(basisPanel);

        int col2     = (row + 1) % 2;
        int gridRow2 = (row + 1) / 2;
        Grid.SetColumn(placePanel, col2);
        Grid.SetRow(placePanel, gridRow2);
        grid.Children.Add(placePanel);
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
            Background = AppTheme.BgSecondary,
        };

        bool result = false;
        var (yBg, yFg, yBd) = StatusBadge.GetBrushes(BadgeStatus.Bad);
        var (nBg, nFg, nBd) = StatusBadge.GetBrushes(BadgeStatus.Muted);
        var yes = new Button { Content = "삭제", Width = 88, Height = 30,
                               Background = yBg, Foreground = yFg, BorderBrush = yBd,
                               BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999),
                               FontWeight = FontWeight.SemiBold, Cursor = new Cursor(StandardCursorType.Hand) };
        var no  = new Button { Content = "취소", Width = 88, Height = 30,
                               Background = nBg, Foreground = nFg, BorderBrush = nBd,
                               BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999),
                               Cursor = new Cursor(StandardCursorType.Hand) };

        yes.Click += (_, _) => { result = true;  dlg.Close(); };
        no.Click  += (_, _) => { result = false; dlg.Close(); };

        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock { Text = message, Foreground = AppRes("AppFg"),
                                FontSize = AppTheme.FontLG, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
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
                FontSize = AppTheme.FontBase, FontFamily = Font,
                Foreground = AppTheme.FgWarn,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            });
            return root;
        }

        root.Children.Add(new TextBlock
        {
            Text = $"업체 {companies.Count}개 — 선택하여 정보를 편집하세요.",
            FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppTheme.FgMuted,
            Margin = new Thickness(0, 0, 0, 4),
        });

        var listBox = new ListBox
        {
            MaxHeight = 220,
            Background = AppTheme.BgPrimary,
            BorderBrush = AppTheme.BorderSubtle, BorderThickness = new Thickness(1),
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
                        new TextBlock { Text = c, FontSize = AppTheme.FontBase, FontFamily = Font,
                                        Foreground = AppTheme.FgPrimary },
                        string.IsNullOrEmpty(info.약칭)
                            ? (Control)new TextBlock()
                            : new Border
                            {
                                Background = AppTheme.BgActiveBlue, CornerRadius = new CornerRadius(3),
                                Padding = new Thickness(4, 1),
                                Child = new TextBlock { Text = info.약칭, FontSize = AppTheme.FontXS,
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
            Height = 1, Background = AppTheme.BorderMuted, Margin = new Thickness(0, 0, 0, 4)
        });
        _measEditForm.Children.Add(new TextBlock
        {
            Text = company, FontSize = AppTheme.FontMD, FontFamily = Font,
            Foreground = AppTheme.FgInfo, FontWeight = FontWeight.SemiBold,
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
                Text = "추천:", FontSize = AppTheme.FontXS, FontFamily = Font,
                Foreground = AppTheme.FgDimmed, VerticalAlignment = VerticalAlignment.Center,
            });
            var (sBg, sFg, sBd) = StatusBadge.GetBrushes(BadgeStatus.Info);
            foreach (var sug in suggestions)
            {
                var s = sug;
                var btn = new Button
                {
                    Content = s, Height = 22, FontSize = AppTheme.FontXS, FontFamily = Font,
                    Background = sBg, Foreground = sFg, BorderBrush = sBd,
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(999),
                    Padding = new Thickness(8, 0),
                    Cursor = new Cursor(StandardCursorType.Hand),
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
            Width = 200, FontSize = AppTheme.FontMD, FontFamily = Font,
            Background = AppTheme.BorderSeparator, Foreground = AppRes("AppFg"),
            BorderBrush = AppTheme.BorderDefault, BorderThickness = new Thickness(1),
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
            Content = "💾  저장", Height = 30, FontSize = AppTheme.FontMD, FontFamily = Font,
            Background = AppTheme.BgActiveGreen, Foreground = AppTheme.FgSuccess,
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
        Text = text + " :", Width = 110, FontSize = AppTheme.FontBase, FontFamily = Font,
        Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center,
    };

    private static TextBox MakeTextBox(string value, string hint = "") => new()
    {
        Text = value, Width = 200, FontSize = AppTheme.FontMD, FontFamily = Font, Watermark = hint,
        Background = AppTheme.BorderSeparator, Foreground = AppRes("AppFg"),
        BorderBrush = AppTheme.BorderDefault, BorderThickness = new Thickness(1),
        CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 4),
    };

    // =========================================================================
    // Show3 — 신규 추가 시 템플릿 업체 선택
    // =========================================================================
    private Control BuildTemplateSelector()
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(8) };
        root.Children.Add(new TextBlock
        {
            Text       = "📋  단가 테이블 템플릿 선택",
            FontSize   = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
        });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderMuted });
        root.Children.Add(new TextBlock
        {
            Text         = "기존 계약업체의 단가 테이블을 선택하면 신규 업체에 동일 단가가 복사됩니다.",
            FontSize     = AppTheme.FontSM, FontFamily = Font,
            Foreground   = AppTheme.FgMuted,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin       = new Thickness(0, 0, 0, 4),
        });

        // 선택 상태 표시 라벨
        var selLabel = new TextBlock
        {
            Text       = "선택 안됨 — 빈 단가로 생성됩니다.",
            FontSize   = AppTheme.FontBase, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
            Margin     = new Thickness(0, 0, 0, 6),
        };
        root.Children.Add(selLabel);

        // 업체 목록 로드
        var contracts = ContractService.GetAllContracts()
            .OrderBy(c => c.C_CompanyName).ToList();

        Border? prevSelected = null;

        foreach (var c in contracts)
        {
            var companyName = c.C_CompanyName;
            var abbr = c.C_Abbreviation;

            // 해당 업체의 단가 요약 (값이 있는 항목 수)
            var prices = ContractService.GetContractPrices(companyName);
            var filledCount = prices.Count(p => !string.IsNullOrWhiteSpace(p.Price));

            var card = new Border
            {
                Background      = AppTheme.BgPrimary,
                BorderBrush     = AppTheme.BorderSubtle,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(5),
                Padding         = new Thickness(10, 6),
                Margin          = new Thickness(0, 1),
                Cursor          = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            };

            var cardContent = new StackPanel { Spacing = 3 };

            // 1줄: 업체명 + 약칭
            var nameLine = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
            nameLine.Children.Add(new TextBlock
            {
                Text       = companyName,
                FontSize   = AppTheme.FontMD, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("AppFg"),
            });
            if (!string.IsNullOrEmpty(abbr))
            {
                var (bg, fg) = BadgeColorHelper.GetBadgeColor(abbr);
                nameLine.Children.Add(new Border
                {
                    Background        = Brush.Parse(bg),
                    CornerRadius      = new CornerRadius(3),
                    Padding           = new Thickness(5, 1),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text       = abbr,
                        FontSize   = AppTheme.FontXS, FontFamily = Font,
                        Foreground = Brush.Parse(fg),
                    }
                });
            }
            cardContent.Children.Add(nameLine);

            // 2줄: 단가 요약 (채워진 항목 / 전체)
            var priceSummary = new WrapPanel { Orientation = Orientation.Horizontal };
            priceSummary.Children.Add(new TextBlock
            {
                Text       = $"단가: {filledCount}/{prices.Count}항목",
                FontSize   = AppTheme.FontSM, FontFamily = Font,
                Foreground = filledCount > 0
                    ? AppTheme.FgSuccess
                    : AppTheme.BorderMuted,
            });

            // 대표 단가 몇 개 미리보기
            var preview = prices.Where(p => !string.IsNullOrWhiteSpace(p.Price)).Take(5).ToList();
            if (preview.Count > 0)
            {
                priceSummary.Children.Add(new TextBlock
                {
                    Text   = "  │  ",
                    FontSize = AppTheme.FontXS, FontFamily = Font,
                    Foreground = AppTheme.FgDimmed,
                });
                priceSummary.Children.Add(new TextBlock
                {
                    Text       = string.Join(", ", preview.Select(p => $"{p.Analyte}:{FormatPrice(p.Price)}")),
                    FontSize   = AppTheme.FontXS, FontFamily = Font,
                    Foreground = AppTheme.FgMuted,
                    TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                });
            }
            cardContent.Children.Add(priceSummary);

            card.Child = cardContent;

            var capturedCard = card;
            var capturedName = companyName;
            card.PointerPressed += (_, _) =>
            {
                // 이전 선택 해제
                if (prevSelected != null)
                {
                    prevSelected.Background  = AppTheme.BgPrimary;
                    prevSelected.BorderBrush = AppTheme.BorderSubtle;
                }

                // 현재 선택
                capturedCard.Background  = AppTheme.BgActiveBlue;
                capturedCard.BorderBrush = AppTheme.BorderInfo;
                prevSelected = capturedCard;

                _templateCompanyName = capturedName;
                selLabel.Text       = $"✅ 선택됨: {capturedName}";
                selLabel.Foreground = AppTheme.FgInfo;
                Log($"[Template] 선택: {capturedName}");
            };

            root.Children.Add(card);
        }

        return root;
    }

    // =========================================================================
    // Show3 — 단가 항목 편집 폼 (단일 항목)
    // =========================================================================

    /// <summary>신규 추가 모드에서 Show3에 템플릿 선택 UI 표시 및 Show4 아이템 선택</summary>
    private void ShowAddModeTemplateSelection()
    {
        var root = new StackPanel { Spacing = 12, Margin = new Thickness(12) };

        root.Children.Add(new TextBlock
        {
            Text       = "💾  기존 계약 단가 복사",
            FontSize   = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
        });
        root.Children.Add(new Border { Height = 1, Background = AppTheme.BorderMuted });

        root.Children.Add(new TextBlock
        {
            Text       = "기존 업체 선택 후 [적용] → Show4에서 항목 선택",
            FontSize   = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
            Margin     = new Thickness(0, 0, 0, 8),
        });

        // 드롭다운 + 적용 버튼 수평 배치
        var controlPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 8,
            VerticalAlignment = VerticalAlignment.Center,
        };

        // 기존 계약 업체 목록 (ComboBox)
        var cmbTemplate = new ComboBox
        {
            Width           = 250,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            BorderThickness = new Thickness(1),
            Background      = AppTheme.BgInput,
            CornerRadius    = new CornerRadius(3),
            Padding         = new Thickness(6, 4),
        };

        try
        {
            var allContracts = ContractService.GetAllContracts();
            foreach (var c in allContracts.OrderBy(x => x.C_CompanyName))
                cmbTemplate.Items.Add(c.C_CompanyName);
        }
        catch { }

        controlPanel.Children.Add(cmbTemplate);

        // 적용 버튼
        var btnApply = new Button
        {
            Content         = "✅  적용",
            Height          = 34,
            FontSize        = AppTheme.FontMD,
            FontFamily      = Font,
            Background      = AppTheme.BgActiveGreen,
            Foreground      = AppTheme.FgSuccess,
            BorderThickness = new Thickness(0),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(16, 0),
        };
        btnApply.Click += (_, _) =>
        {
            var selectedCompany = cmbTemplate.SelectedItem?.ToString();
            if (!string.IsNullOrEmpty(selectedCompany))
            {
                _templateCompanyName = selectedCompany;
                ShowAddModeAnalyteSelection(selectedCompany);
                Log($"✅ 템플릿 선택: {selectedCompany}");
            }
        };
        controlPanel.Children.Add(btnApply);

        root.Children.Add(controlPanel);

        EditPanelChanged?.Invoke(root);
    }

    /// <summary>신규 추가 모드에서 Show4에 선택된 템플릿의 분석항목 + 수량 입력 표시 (카테고리 그룹)</summary>
    private void ShowAddModeAnalyteSelection(string templateCompanyName)
    {
        try
        {
            var templatePrices = ContractService.GetContractPrices(templateCompanyName);
            var priceDict      = templatePrices
                .Where(p => !string.IsNullOrWhiteSpace(p.Price))
                .ToDictionary(p => p.Analyte, p => p.Price, StringComparer.OrdinalIgnoreCase);

            var aliasMap = ContractService.GetAnalyteAliasMap();
            var allItems = AnalysisService.GetAllItems();

            var root = new StackPanel { Spacing = 0, Margin = new Thickness(12) };

            // 헤더
            root.Children.Add(new TextBlock
            {
                Text       = $"📋  {templateCompanyName}",
                FontSize   = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgMuted,
                Margin     = new Thickness(0, 0, 0, 2),
            });
            root.Children.Add(new TextBlock
            {
                Text       = $"({priceDict.Count}개 항목)",
                FontSize   = AppTheme.FontSM, FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Thickness(0, 0, 0, 8),
            });

            // 모두 선택 + 확정 버튼
            var topControlPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing     = 16,
                VerticalAlignment = VerticalAlignment.Center,
                Margin     = new Thickness(0, 0, 0, 8),
            };
            var chkSelectAll = new CheckBox { Content = "모두 선택", FontSize = AppTheme.FontMD, FontFamily = Font, IsChecked = true };
            topControlPanel.Children.Add(chkSelectAll);
            var btnConfirm = new Button
            {
                Content         = "✅  확정",
                Height          = 34,
                FontSize        = AppTheme.FontMD, FontFamily = Font,
                Background      = AppTheme.BgActiveGreen,
                Foreground      = AppTheme.FgSuccess,
                BorderThickness = new Thickness(0),
                CornerRadius    = new CornerRadius(4),
                Padding         = new Thickness(16, 0),
            };
            topControlPanel.Children.Add(btnConfirm);
            root.Children.Add(topControlPanel);

            // 컬럼 헤더
            var headerGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,90,70,90"), ColumnSpacing = 8 };
            foreach (var (txt, col, align) in new[] {
                ("",      0, Avalonia.Media.TextAlignment.Left),
                ("항목명", 1, Avalonia.Media.TextAlignment.Left),
                ("단가",  2, Avalonia.Media.TextAlignment.Right),
                ("수량",  3, Avalonia.Media.TextAlignment.Right),
                ("소계",  4, Avalonia.Media.TextAlignment.Right) })
            {
                var h = new TextBlock { Text = txt, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppTheme.FgDimmed, TextAlignment = align };
                Grid.SetColumn(h, col); headerGrid.Children.Add(h);
            }
            root.Children.Add(new Border { Background = AppTheme.BgSecondary ?? AppTheme.BgPrimary, Padding = new Thickness(8, 4), Child = headerGrid, Margin = new Thickness(0, 0, 0, 4) });

            // 합계 블록
            var addModeTotalBlock = new TextBlock
            {
                FontSize = AppTheme.FontMD, FontFamily = Font, FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgSuccess, Margin = new Thickness(0, 0, 0, 6),
            };
            root.Children.Add(addModeTotalBlock);

            // 카테고리 그룹핑
            var catOrder = allItems
                .Select(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
                .Distinct().ToList();
            if (catOrder.Remove("일반항목")) catOrder.Insert(0, "일반항목");

            var knownAnalytes = allItems.Select(a => a.Analyte).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var unknownItems  = priceDict.Keys
                .Where(k => !knownAnalytes.Contains(k))
                .Select(k => new AnalysisItem { Analyte = k, Category = "기타", 약칭 = "", ES = "zzz" })
                .ToList();

            var itemMeta = allItems
                .Where(a => priceDict.ContainsKey(a.Analyte))
                .GroupBy(a => string.IsNullOrEmpty(a.Category) ? "기타" : a.Category)
                .ToDictionary(g => g.Key);
            if (unknownItems.Count > 0 && !itemMeta.ContainsKey("기타"))
                catOrder.Add("기타");

            // analyte → (CheckBox, TextBox) 맵 (확정 시 사용)
            var chkMap = new Dictionary<string, CheckBox>(StringComparer.OrdinalIgnoreCase);
            var qtyMap = new Dictionary<string, TextBox>(StringComparer.OrdinalIgnoreCase);
            var allCheckBoxes = new List<CheckBox>();

            void RecalcTotal()
            {
                decimal total = 0m;
                foreach (var (an, cb) in chkMap)
                {
                    if (cb.IsChecked != true) continue;
                    if (!priceDict.TryGetValue(an, out var pr)) continue;
                    decimal.TryParse(pr, out decimal pv);
                    int qv = qtyMap.TryGetValue(an, out var qb) && int.TryParse(qb.Text, out int qi) && qi > 0 ? qi : 1;
                    total += pv * qv;
                }
                addModeTotalBlock.Text = $"💵 합계: {total:N0} 원";
            }

            var itemsPanel = new StackPanel { Spacing = 4 };

            foreach (var catKey in catOrder)
            {
                var groupItems = (itemMeta.TryGetValue(catKey, out var grp) ? grp.OrderBy(a => a.ES).ToList() : new List<AnalysisItem>());
                if (catKey == "기타") groupItems.AddRange(unknownItems);
                groupItems = groupItems.Where(a => priceDict.ContainsKey(a.Analyte)).ToList();
                if (groupItems.Count == 0) continue;

                // 카테고리 헤더
                var catChk = new CheckBox
                {
                    Content = $"전체 ({groupItems.Count})", IsChecked = true,
                    FontSize = AppTheme.FontXS, FontFamily = Font,
                    Foreground = Brush.Parse("#88bb88"),
                    Padding = new Thickness(2, 0), Margin = new Thickness(0), MinHeight = 18,
                };
                var chevron = new TextBlock { Text = "▾", FontSize = 10, Foreground = AppTheme.FgMuted, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) };
                var hdrRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto") };
                hdrRow.Children.Add(chevron); Grid.SetColumn(chevron, 0);
                var catLbl = new TextBlock { Text = catKey, FontSize = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.SemiBold, Foreground = AppTheme.FgInfo, VerticalAlignment = VerticalAlignment.Center };
                hdrRow.Children.Add(catLbl); Grid.SetColumn(catLbl, 1);
                hdrRow.Children.Add(catChk); Grid.SetColumn(catChk, 2);
                var hdrBorder = new Border
                {
                    Background = AppTheme.BgSecondary ?? AppTheme.BgPrimary,
                    Padding = new Thickness(8, 5), Margin = new Thickness(0, 0, 0, 1),
                    CornerRadius = new CornerRadius(4),
                    Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                    Child = hdrRow,
                };

                var catItemsBorder = new Border
                {
                    Background = AppTheme.BgPrimary,
                    Padding = new Thickness(8, 4, 4, 4), Margin = new Thickness(0, 0, 0, 4),
                    CornerRadius = new CornerRadius(0, 0, 4, 4), IsVisible = true,
                };
                var catItemsStack = new StackPanel { Spacing = 0 };
                catItemsBorder.Child = catItemsStack;

                var catCbs = new List<CheckBox>();

                foreach (var item in groupItems)
                {
                    var analyte = item.Analyte;
                    if (!priceDict.TryGetValue(analyte, out var price)) continue;
                    decimal.TryParse(price, out decimal priceValue);

                    var chk    = new CheckBox { IsChecked = true, VerticalAlignment = VerticalAlignment.Center };
                    var qtyBox = new TextBox
                    {
                        Text = "1", FontSize = AppTheme.FontSM, FontFamily = Font,
                        TextAlignment = Avalonia.Media.TextAlignment.Right,
                        Height = 28, Padding = new Thickness(4, 2),
                        VerticalAlignment = VerticalAlignment.Center,
                    };
                    var subtotalTb = new TextBlock
                    {
                        Text = priceValue > 0 ? priceValue.ToString("N0") : "—",
                        FontSize = AppTheme.FontSM, FontFamily = Font,
                        Foreground = AppTheme.FgSuccess, FontWeight = FontWeight.SemiBold,
                        TextAlignment = Avalonia.Media.TextAlignment.Right,
                        VerticalAlignment = VerticalAlignment.Center,
                    };

                    qtyBox.TextChanged += (_, _) =>
                    {
                        subtotalTb.Text = int.TryParse(qtyBox.Text, out int q) && q > 0
                            ? (priceValue * q).ToString("N0") : "—";
                        RecalcTotal();
                    };
                    chk.IsCheckedChanged += (_, _) => RecalcTotal();

                    // 뱃지 + 항목명
                    var analyteAlias = string.IsNullOrEmpty(item.약칭) ? (aliasMap.TryGetValue(analyte, out var al) ? al : "") : item.약칭;
                    var (abg, afg) = BadgeColorHelper.GetBadgeColor(analyteAlias);
                    var badge = new Border
                    {
                        Background = Brush.Parse(abg), BorderBrush = Brush.Parse(afg),
                        BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10),
                        Padding = new Thickness(6, 1, 8, 1), Margin = new Thickness(0, 0, 4, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                        Child = new TextBlock
                        {
                            Text = string.IsNullOrEmpty(analyteAlias) ? "✓" : analyteAlias,
                            FontSize = AppTheme.FontSM, FontFamily = Font, FontWeight = FontWeight.Medium,
                            Foreground = Brush.Parse(afg), VerticalAlignment = VerticalAlignment.Center,
                        }
                    };
                    var namePanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal, Spacing = 0,
                        VerticalAlignment = VerticalAlignment.Center,
                        Children = { badge, new TextBlock { Text = analyte, FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppTheme.FgPrimary, VerticalAlignment = VerticalAlignment.Center, TextTrimming = TextTrimming.CharacterEllipsis } }
                    };

                    var itemGrid = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,90,70,90"), ColumnSpacing = 8 };
                    Grid.SetColumn(chk,         0); itemGrid.Children.Add(chk);
                    Grid.SetColumn(namePanel,   1); itemGrid.Children.Add(namePanel);
                    var priceBlock = new TextBlock { Text = FormatPrice(price), FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = AppTheme.FgSuccess, TextAlignment = Avalonia.Media.TextAlignment.Right, VerticalAlignment = VerticalAlignment.Center };
                    Grid.SetColumn(priceBlock,  2); itemGrid.Children.Add(priceBlock);
                    Grid.SetColumn(qtyBox,      3); itemGrid.Children.Add(qtyBox);
                    Grid.SetColumn(subtotalTb,  4); itemGrid.Children.Add(subtotalTb);

                    catItemsStack.Children.Add(new Border
                    {
                        Padding = new Thickness(0, 4), BorderBrush = AppTheme.BorderSubtle,
                        BorderThickness = new Thickness(0, 0, 0, 1), Child = itemGrid,
                    });

                    chkMap[analyte] = chk;
                    qtyMap[analyte] = qtyBox;
                    allCheckBoxes.Add(chk);
                    catCbs.Add(chk);
                }

                // 카테고리 전체선택
                var capturedCatChk = catChk;
                var capturedCatCbs = catCbs;
                catChk.IsCheckedChanged += (_, _) =>
                {
                    foreach (var cb in capturedCatCbs) cb.IsChecked = capturedCatChk.IsChecked ?? false;
                };

                // 접힘/펼침
                var capturedCatItems = catItemsBorder;
                hdrBorder.PointerPressed += (_, _) =>
                {
                    capturedCatItems.IsVisible = !capturedCatItems.IsVisible;
                    chevron.Text = capturedCatItems.IsVisible ? "▾" : "▸";
                };

                itemsPanel.Children.Add(hdrBorder);
                itemsPanel.Children.Add(catItemsBorder);
            }

            // 초기 합계
            decimal initTotal = 0m;
            foreach (var (an, pr) in priceDict) { decimal.TryParse(pr, out decimal p); initTotal += p; }
            addModeTotalBlock.Text = $"💵 합계: {initTotal:N0} 원";

            // 모두 선택/해제
            chkSelectAll.IsCheckedChanged += (_, _) =>
            {
                var v = chkSelectAll.IsChecked ?? false;
                foreach (var cb in allCheckBoxes) cb.IsChecked = v;
            };

            var scroll = new ScrollViewer
            {
                Content = itemsPanel,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                Margin = new Thickness(0, 0, 0, 8),
            };
            root.Children.Add(scroll);

            // 확정 버튼
            btnConfirm.Click += (_, _) =>
            {
                Log("[확정] 버튼 클릭됨");
                var selectedItems = new List<(string Analyte, string Price, string Qty)>();
                foreach (var (analyte, cb) in chkMap)
                {
                    if (cb.IsChecked != true) continue;
                    if (!priceDict.TryGetValue(analyte, out var price)) continue;
                    var qty = qtyMap.TryGetValue(analyte, out var qb) ? qb.Text?.Trim() ?? "1" : "1";
                    selectedItems.Add((analyte, price, qty));
                }

                if (selectedItems.Count > 0)
                {
                    decimal confirmTotal = 0m;
                    Log($"[확정] 선택항목 {selectedItems.Count}개 처리 시작");
                    foreach (var (analyte, price, qty) in selectedItems)
                    {
                        _pendingPrices[analyte] = price;
                        if (!string.IsNullOrWhiteSpace(qty)) _pendingQuantities[analyte] = qty;
                        decimal.TryParse(price, out decimal p); int.TryParse(qty, out int q);
                        confirmTotal += p * q;
                        Log($"[확정]  항목={analyte}, price='{price}'({p}), qty='{qty}'({q}), 소계={p*q:N0}");
                    }
                    Log($"[확정] 합계={confirmTotal:N0}, pendingPrices={_pendingPrices.Count}, pendingQtys={_pendingQuantities.Count}");
                    if (_detailPanel != null && confirmTotal > 0)
                        UpdateContractAmountInShow2(_detailPanel, confirmTotal);
                    SaveAdd();
                }
                else
                {
                    Log("❌ 확정 실패: 선택된 항목 없음");
                }
            };

            var outerScroll = new ScrollViewer
            {
                Content = root,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            };
            StatsPanelChanged?.Invoke(outerScroll);
        }
        catch (Exception ex)
        {
            Log($"❌ 항목 선택 오류: {ex.Message}");
        }
    }

    /// <summary>Show4 행 선택 업데이트 및 수량 인라인 편집 활성화</summary>
    private void UpdateAddModeRowSelection(List<Border> itemRows, List<TextBox> qtyTextBoxes, int selectedIndex)
    {
        for (int i = 0; i < itemRows.Count; i++)
        {
            var row = itemRows[i];
            var grid = row.Child as Grid;
            if (grid == null) continue;

            if (i == selectedIndex)
            {
                // 선택 상태: 밝은 배경 강조
                row.Background = new SolidColorBrush(Color.Parse("#334455"));
                row.BorderThickness = new Thickness(2);
                row.BorderBrush = new SolidColorBrush(Color.Parse("#88ccff"));

                if (i < qtyTextBoxes.Count)
                {
                    qtyTextBoxes[i].Focus();
                    qtyTextBoxes[i].SelectAll();
                }
            }
            else
            {
                // 비선택 상태: 원래 배경
                row.Background = i % 2 == 0
                    ? AppTheme.BgPrimary
                    : (AppTheme.BgSecondary ?? AppTheme.BgPrimary);
                row.BorderThickness = new Thickness(0);
            }
        }
    }

    /// <summary>detailPanel에서 업체명 TextBox 찾기</summary>
    private TextBox? FindCompanyNameTextBox(StackPanel? panel)
    {
        if (panel == null) return null;
        var grid = panel.Children.OfType<Grid>().FirstOrDefault();
        if (grid == null) return null;
        foreach (var child in grid.Children)
        {
            if (child is StackPanel sp && sp.Children.Count > 1)
            {
                if (sp.Children[0] is TextBlock tb && tb.Text?.Contains("업체명") == true)
                {
                    if (sp.Children[1] is TextBox txb)
                        return txb;
                }
            }
        }
        return null;
    }

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
            FontSize   = AppTheme.FontLG, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgMuted,
        });
        root.Children.Add(new Border
        {
            Height     = 1,
            Background = AppTheme.BorderMuted,
        });

        // 항목명
        root.Children.Add(new TextBlock
        {
            Text       = analyte,
            FontSize   = AppTheme.FontMD, FontFamily = Font, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg"),
        });

        // 단가 입력
        var priceBox = new TextBox
        {
            Text            = currentPrice,
            FontSize        = AppTheme.FontXL, FontFamily = Font,
            Watermark       = "단가 입력 (원)",
            Background      = AppTheme.BorderSeparator,
            Foreground      = AppRes("AppFg"),
            BorderBrush     = AppTheme.BorderDefault,
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(10, 6),
        };
        root.Children.Add(priceBox);

        // 적용 버튼 (Show2 표 업데이트 + pending 기록)
        var btnApply = new Button
        {
            Content         = "✅  적용 (Show2 반영)",
            Height          = 34, FontSize = AppTheme.FontMD, FontFamily = Font,
            Background      = AppTheme.BgActiveBlue,
            Foreground      = AppTheme.FgInfo,
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
                                            ? AppTheme.BorderDefault
                                            : new SolidColorBrush(Color.Parse("#ffdd88")); // 변경됨 표시
            }
            Log($"단가 적용(pending): {analyte} = {newPrice}");
        };
        root.Children.Add(btnApply);

        root.Children.Add(new TextBlock
        {
            Text       = "※ 서브메뉴 [저장]으로 서버에 최종 반영됩니다.",
            FontSize   = AppTheme.FontSM, FontFamily = Font,
            Foreground = AppTheme.FgDimmed,
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
                    tb.Foreground = AppTheme.FgSuccess;
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
        if (App.EnableLogging)
        {
            try { File.AppendAllText("Logs/ContractDebug.log", line + Environment.NewLine); } catch { }
        }
    }
}
