using Avalonia;
using ETA.Views;
using Avalonia.Input;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using Avalonia.VisualTree;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

namespace ETA.Views.Pages.PAGE1;

public partial class QuotationNewPanel : UserControl
{
    // Window Tunnel KeyDown 핸들러 등록용
    private Window? _attachedWindow;
    // Show2(항목 목록) 네비/인라인 편집 상태 (WasteAnalysisInputPage 참고)
    private bool _keyNavShow2 = false; // Shift+2 네비 모드
    private int _keyNavShow2Index = -1; // 현재 포커스 행
    private int _selectedRowIndex = -1; // 실제 선택된 행 인덱스
    private List<Button?> _rowQtyButtons = new(); // 각 행의 수량 버튼
    private List<TextBox?> _rowQtyInputs = new(); // 각 행의 수량 TextBox
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 상태 ──────────────────────────────────────────────────────────────
    private Contract?       _company;
    private QuotationIssue? _editingIssue;
    private string?         _editingAnalyte;

    // 당근(재활용) 모드에서 업체 정보 보존용
    private string _carrotCompanyName = "";
    private string _carrotAbbr        = "";

    private List<string> _knownManagers = new();

    // 항목명 → { qty, unitPrice }
    private readonly Dictionary<string, (int Qty, decimal Price)> _itemData = new();
    private readonly Dictionary<string, AnalysisItem>             _analyteMap = new();

    private Dictionary<string, decimal> _priceMap = new();

    // 경고 상태
    private bool _sampleDuplicated = false;

    // 최근 발행건 (중복/적용구분 비교용)
    private List<QuotationIssue> _allIssues = [];

    public QuotationNewPanel()
    {
        InitializeComponent();
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbQuotationNo.Text = GenerateNo();
        _allIssues = QuotationService.GetAllIssues();

        // ESC 키 → 당근/오작성 수정 모드일 때 취소
        KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape &&
                (_editingIssue != null || _carrotCompanyName.Length > 0))
            {
                Clear();
                EscapeCancelled?.Invoke();
            }
        };

        // Enter 키 네비게이션
        txbSampleName.KeyDown  += (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; acbManager.Focus(); } };
        acbManager.KeyDown     += (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; txbSampleName.Focus(); } };
        txbIssueDate.KeyDown   += (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; acbManager.Focus(); } };
        txbQuotationNo.KeyDown += (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; txbSampleName.Focus(); } };
        txbBulkQty.KeyDown     += (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; BtnBulkQty_Click(null, null!); } };
        txbQty.KeyDown         += (_, e) => { if (e.Key == Key.Enter) { e.Handled = true; BtnApplyQty_Click(null, null!); } };

        acbManager.SelectionChanged += (_, _) => OnManagerChanged();
        acbManager.KeyUp += (_, _) => OnManagerChanged();

        void OnManagerChanged()
        {
            var name = (acbManager.SelectedItem as string)
                       ?? acbManager.Text?.Trim() ?? "";
            bool isKnown = !string.IsNullOrWhiteSpace(name)
                        && _knownManagers.Contains(name, StringComparer.OrdinalIgnoreCase);
            bool isNew   = !string.IsNullOrWhiteSpace(name) && !isKnown;

            if (isKnown)
            {
                var companyName = _company?.C_CompanyName ?? _carrotCompanyName;
                if (!string.IsNullOrEmpty(companyName))
                {
                    var (phone, email) = QuotationService.GetManagerContactInfo(companyName, name);
                    txbManagerPhone.Text = phone;
                    txbManagerEmail.Text = email;
                }
                pnlContactInfo.IsVisible = true;
            }
            else if (isNew)
            {
                pnlContactInfo.IsVisible = true;
            }
            else
            {
                txbManagerPhone.Text = "";
                txbManagerEmail.Text = "";
                pnlContactInfo.IsVisible = false;
            }
        }

        // 한글 IME 따라오기 방지: GotFocus 시 글자 수 저장, 이후 Background 우선순위로 초과분 제거
        AttachImeFix(txbSampleName);
        AttachImeFix(txbIssueDate);
        AttachImeFix(txbQuotationNo);
        AttachImeFix(txbBulkQty);
        AttachImeFix(txbQty);

        // Window Tunnel KeyDown 핸들러 등록 (WasteAnalysisInputPage 패턴)
        AttachedToVisualTree += (_, _) =>
        {
            // Control.Parent를 타고 올라가 Window를 찾음
            Control? v = this;
            Window? win = null;
            while (v != null)
            {
                if (v is Window w)
                {
                    win = w;
                    break;
                }
                v = v.Parent as Control;
            }
            if (win != null && win != _attachedWindow)
            {
                if (_attachedWindow != null)
                {
                    _attachedWindow.RemoveHandler(InputElement.KeyDownEvent, OnNavKeyDownTunnel);
                }
                _attachedWindow = win;
                win.AddHandler(InputElement.KeyDownEvent, OnNavKeyDownTunnel, Avalonia.Interactivity.RoutingStrategies.Tunnel);
            }
        };
        DetachedFromVisualTree += (_, _) =>
        {
            if (_attachedWindow != null)
            {
                _attachedWindow.RemoveHandler(InputElement.KeyDownEvent, OnNavKeyDownTunnel);
                _attachedWindow = null;
            }
        };
    }

    // ── 한글 IME 따라오기 방지 ───────────────────────────────────────────
    private static void AttachImeFix(TextBox tb)
    {
        tb.GotFocus += (s, _) =>
        {
            var box      = (TextBox)s!;
            var savedLen = (box.Text ?? "").Length;
            Avalonia.Threading.Dispatcher.UIThread.Post(() =>
            {
                var cur = box.Text ?? "";
                if (cur.Length > savedLen)
                {
                    box.Text       = cur[..savedLen];
                    box.CaretIndex = savedLen;
                }
            }, Avalonia.Threading.DispatcherPriority.Background);
        };
    }

    // ══════════════════════════════════════════════════════════════════════
    //  외부 API
    // ══════════════════════════════════════════════════════════════════════
    // ── 업체 블록 표시 헬퍼 ──────────────────────────────────────────────────
    private void ShowCompanyBadge(string companyName, string abbr)
    {
        txbCompany.Text = companyName;
        if (!string.IsNullOrWhiteSpace(abbr))
        {
            var (bg, fg) = BadgeColorHelper.GetBadgeColor(abbr);
            bdgAbbr.Background = Avalonia.Media.Brush.Parse(bg);
            txbAbbr.Foreground = Avalonia.Media.Brush.Parse(fg);
            txbAbbr.Text       = abbr;
            bdgAbbr.IsVisible  = true;
        }
        else
        {
            bdgAbbr.IsVisible = false;
        }
    }

    public void SetCompany(Contract company)
    {
        _company = company;
        ShowCompanyBadge(company.C_CompanyName, company.C_Abbreviation);

        var latest = _allIssues
            .Where(i => i.업체명 == company.C_CompanyName)
            .OrderByDescending(i => i.발행일)
            .FirstOrDefault();

        // 업체 단가 로드
        RefreshPrices();

        // 담당자 목록 로드
        _knownManagers = QuotationService.GetDistinctManagersForCompany(company.C_CompanyName);
        acbManager.ItemsSource = _knownManagers;

        // 최근 담당자 자동 설정
        if (latest != null && !string.IsNullOrEmpty(latest.담당자))
        {
            acbManager.Text = latest.담당자;
            Log($"업체 최근 담당자 자동설정: {latest.담당자}");
        }
    }

    public void SetSelectedAnalytes(List<AnalysisItem> items)
    {
        // 항상 단가를 먼저 최신화 (이전 정상 동작 방식)
        RefreshPrices();

        var newNames = items.Select(a => a.Analyte).ToHashSet();

        foreach (var key in _analyteMap.Keys.ToList())
            if (!newNames.Contains(key))
            {
                _analyteMap.Remove(key);
                _itemData.Remove(key);
            }

        foreach (var item in items)
        {
            _analyteMap[item.Analyte] = item;
            if (!_itemData.ContainsKey(item.Analyte))
            {
                // 단가를 즉시 _priceMap에서 조회하여 할당
                var price = _priceMap.TryGetValue(item.Analyte, out var p) ? p : 0m;
                _itemData[item.Analyte] = (1, price);
            }
        }

        RebuildItemList();
    }

    // 🥕 당근: 이 건 재활용 — 항목 복사, 번호·날짜는 신규
    public void LoadFromIssue(QuotationIssue issue)
    {
        _editingIssue      = null;   // 재활용은 신규 저장
        _company           = null;
        _carrotCompanyName = issue.업체명;
        _carrotAbbr        = issue.약칭;

        // ── 네비게이션 상태 초기화 ──────────────────────────────────────────
        _keyNavShow2      = false;
        _keyNavShow2Index = -1;
        _selectedRowIndex = -1;
        _rowQtyButtons.Clear();
        _rowQtyInputs.Clear();

        txbTitle.Text       = "🔄  ReCyle (재활용)";
        txbMode.Text        = "재활용 모드";
        ShowCompanyBadge(issue.업체명, issue.약칭);
        txbQuotationNo.Text = GenerateNo();
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbSampleName.Text  = "";   // 시료명은 새로 입력
        acbManager.Text     = issue.담당자;   // 담당자 유지
        pnlContactInfo.IsVisible = false;
        txbManagerPhone.Text = "";
        txbManagerEmail.Text = "";
        pnlQtyEdit.IsVisible = false;
        _editingAnalyte      = null;

        var row = QuotationService.GetIssueRow(issue.Id);
        LoadItemsFromRow(row);

        CheckSampleDuplicate();
    }

    // ✏️ 오작성 수정: 같은 Id 덮어쓰기 — 시료명·번호·날짜·업체명 수정 가능
    public void LoadFromIssueCorrect(QuotationIssue issue)
    {
        _editingIssue      = issue;
        _company           = null;
        _carrotCompanyName = "";
        _carrotAbbr        = "";

        // ── 네비게이션 상태 초기화 ──────────────────────────────────────────
        _keyNavShow2      = false;
        _keyNavShow2Index = -1;
        _selectedRowIndex = -1;
        _rowQtyButtons.Clear();
        _rowQtyInputs.Clear();

        // _allIssues 갱신 후 자기 자신 제외하므로 중복 오류 안 남
        _allIssues = QuotationService.GetAllIssues();

        txbTitle.Text       = "✏️  오작성 수정";
        txbMode.Text        = "수정 모드";
        ShowCompanyBadge(issue.업체명, issue.약칭);
        txbQuotationNo.Text = issue.견적번호;   // 기존 번호 유지 (수정 가능)
        txbIssueDate.Text   = issue.발행일;     // 기존 날짜 유지 (수정 가능)
        txbSampleName.Text  = issue.시료명;     // 기존 시료명 유지 (수정 가능)
        acbManager.Text     = issue.담당자;     // 담당자 유지

        // 기존 연락처/이메일 — DB 최신값 우선, 없으면 issue에 저장된 값 사용
        var (dbPhone, dbEmail) = !string.IsNullOrEmpty(issue.담당자)
            ? QuotationService.GetManagerContactInfo(issue.업체명, issue.담당자)
            : ("", "");
        txbManagerPhone.Text = !string.IsNullOrEmpty(dbPhone) ? dbPhone : issue.담당자연락처;
        txbManagerEmail.Text = !string.IsNullOrEmpty(dbEmail) ? dbEmail : issue.담당자이메일;
        pnlContactInfo.IsVisible = !string.IsNullOrEmpty(issue.담당자);
        pnlQtyEdit.IsVisible = false;
        _editingAnalyte      = null;

        var row = QuotationService.GetIssueRow(issue.Id);
        LoadItemsFromRow(row);

        // 오작성 수정 모드: 동일 시료명 중복 경고 제외 (자기 자신이므로)
        CheckSampleDuplicate();
    }

    public void Clear()
    {
        _editingIssue      = null;
        _company           = null;
        _editingAnalyte    = null;
        _carrotCompanyName = "";
        _carrotAbbr        = "";
        _itemData.Clear();
        _analyteMap.Clear();
        _priceMap.Clear();
        _sampleDuplicated = false;

        // ── 네비게이션 상태 초기화 ──────────────────────────────────────────
        _keyNavShow2      = false;
        _keyNavShow2Index = -1;
        _selectedRowIndex = -1;
        _rowQtyButtons.Clear();
        _rowQtyInputs.Clear();

        txbTitle.Text       = "📝  신규 견적 작성";
        txbMode.Text        = "";
        txbCompany.Text     = "— 오른쪽에서 업체를 선택하세요 —";
        bdgAbbr.IsVisible   = false;
        txbQuotationNo.Text = GenerateNo();
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbSampleName.Text  = "";
        acbManager.Text     = "";
        txbManagerPhone.Text = "";
        txbManagerEmail.Text = "";
        pnlContactInfo.IsVisible = false;
        spItems.Children.Clear();
        txbNoItems.IsVisible = true;
        pnlQtyEdit.IsVisible = false;
        txbTotal.Text        = "—";
        txbWarning.IsVisible = false;
        txbItemCount.Text    = "";
        pnlQtyEdit.IsVisible = false;
        _editingAnalyte      = null;
        UpdateSaveButton();

        _allIssues = QuotationService.GetAllIssues();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  단가 로드 (업체별 계약 DB 기준)
    // ══════════════════════════════════════════════════════════════════════
    private string GetCurrentCompanyName() =>
        _company?.C_CompanyName ??
        (_carrotCompanyName.Length > 0 ? _carrotCompanyName : "");

    private void RefreshPrices()
    {
        var company = GetCurrentCompanyName();
        if (string.IsNullOrEmpty(company)) return;

        _priceMap = QuotationService.GetPricesByCompany(company);
        Log($"단가 로드: {company} → {_priceMap.Count}개");

        foreach (var key in _itemData.Keys.ToList())
        {
            var price = _priceMap.TryGetValue(key, out var p) ? p : 0;
            _itemData[key] = (_itemData[key].Qty, price);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  항목 목록 빌드
    // ══════════════════════════════════════════════════════════════════════
    /// <summary>Show2(항목 목록) 빌드 후 외부 패널 할당용 이벤트 (분석결과입력과 동일)</summary>
    public event Action<Control?>? ListPanelChanged;

    // Show2(항목 목록) 네비/인라인 편집 복구
    private void RebuildItemList()
    {
        spItems.Children.Clear();
        _rowQtyButtons = new List<Button?>();
        _rowQtyInputs = new List<TextBox?>();
        txbNoItems.IsVisible = _analyteMap.Count == 0;
        txbItemCount.Text    = $"{_analyteMap.Count}개 항목";

        // 빌드 후 ListPanelChanged 이벤트 호출 (자기 자신 전달)
        ListPanelChanged?.Invoke(this);

        var aliasMap = ContractService.GetAnalyteAliasMap();
        decimal totalAmount = 0;
        bool odd = false;
        int idx = 0;
        foreach (var (name, meta) in _analyteMap)
        {
            var (qty, price) = _itemData.TryGetValue(name, out var d) ? d : (1, 0m);
            decimal sub = qty * price;
            totalAmount += sub;

            // 배지
            var alias = aliasMap.TryGetValue(name, out var a) ? a : name;
            var (bgColor, fgColor) = BadgeColorHelper.GetBadgeColor(alias);
            var badge = new Border
            {
                Background = new SolidColorBrush(Color.Parse(bgColor)),
                BorderBrush = new SolidColorBrush(Color.Parse(fgColor)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(6, 1, 8, 1),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = alias,
                    FontSize = AppFonts.Base,
                    FontWeight = FontWeight.Medium,
                    FontFamily = Font,
                    Foreground = new SolidColorBrush(Color.Parse(fgColor)),
                    VerticalAlignment = VerticalAlignment.Center,
                }
            };

            var contentGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,70,70,70"),
                VerticalAlignment = VerticalAlignment.Center,
            };

            // Column 0: 배지
            contentGrid.Children.Add(badge);
            Grid.SetColumn(badge, 0);

            // Column 1: 항목명
            var nameBlock = new TextBlock
            {
                Text      = name,
                FontSize  = AppFonts.MD, FontFamily = Font,
                Foreground = Brush.Parse("#cccccc"),
                TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
                Margin    = new Avalonia.Thickness(8, 0, 0, 0),
            };
            contentGrid.Children.Add(nameBlock);
            Grid.SetColumn(nameBlock, 1);

            // Column 2: 수량 — 항상 편집 가능한 TextBox로 노출 (직접 입력)
            var qtyInput = new TextBox
            {
                Text = qty.ToString(),
                FontSize = AppFonts.MD,
                FontFamily = Font,
                Foreground = qty > 1
                    ? Brush.Parse("#88d888")
                    : AppTheme.FgSecondary,
                Background = AppRes("InputBg"),
                BorderBrush = AppRes("InputBorder"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                HorizontalContentAlignment = HorizontalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 56,
                Padding = new Thickness(2, 1),
            };
            Grid.SetColumn(qtyInput, 2);
            contentGrid.Children.Add(qtyInput);

            while (_rowQtyButtons.Count <= idx) _rowQtyButtons.Add(null);
            while (_rowQtyInputs.Count <= idx) _rowQtyInputs.Add(null);
            // 호환을 위해 버튼 슬롯은 남겨두되 실제로는 TextBox만 사용
            _rowQtyButtons[idx] = null;
            _rowQtyInputs[idx]  = qtyInput;

            qtyInput.GotFocus += (_, _) => qtyInput.SelectAll();
            qtyInput.LostFocus += (_, _) => CommitQty(idx);
            qtyInput.AddHandler(TextBox.KeyDownEvent, (object? _, KeyEventArgs ke) =>
            {
                if (ke.Key == Key.Enter || ke.Key == Key.Down)
                {
                    ke.Handled = true;
                    CommitQty(idx);
                    if (idx < _rowQtyInputs.Count - 1) MoveQtyFocus(idx, +1);
                }
                else if (ke.Key == Key.Up)     { ke.Handled = true; CommitQty(idx); MoveQtyFocus(idx, -1); }
                else if (ke.Key == Key.Escape) { ke.Handled = true; qtyInput.Text = qty.ToString(); }
            }, Avalonia.Interactivity.RoutingStrategies.Tunnel);

            // Column 3: 단가
            var priceLabel = new TextBlock
            {
                Text = price > 0 ? $"{price:#,0}" : "—",
                FontSize = AppFonts.MD,
                Foreground = Brush.Parse("#888888"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            contentGrid.Children.Add(priceLabel);
            Grid.SetColumn(priceLabel, 3);

            // Column 4: 소계
            var subtotalLabel = new TextBlock
            {
                Text = sub > 0 ? $"{sub:#,0}" : "—",
                FontSize = AppFonts.MD,
                Foreground = Brush.Parse("#88cc88"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            contentGrid.Children.Add(subtotalLabel);
            Grid.SetColumn(subtotalLabel, 4);

            // Border 감싸기
            var itemBorder = new Border
            {
                Background = Brush.Parse(odd ? "#1a1a28" : "#1e1e30"),
                Padding = new Thickness(12, 4),
                Margin = new Thickness(0, 1),
                CornerRadius = new CornerRadius(4),
                Cursor = new Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child = contentGrid,
            };

            itemBorder.PointerPressed += (_, pe) => { SelectGridRow(idx); };

            spItems.Children.Add(itemBorder);
            odd = !odd;
            idx++;
        }

        txbTotal.Text = _analyteMap.Count > 0
            ? $"{totalAmount:#,0} 원  ({_analyteMap.Count}항목)"
            : "—";
    }
    // 네비/인라인 편집 로직 — TextBox만 사용 (버튼 토글 제거됨)
    private void OpenQtyInput(int idx)
    {
        if (idx < 0 || idx >= _rowQtyInputs.Count || _rowQtyInputs[idx] == null) return;
        _selectedRowIndex = idx;
        var tb = _rowQtyInputs[idx]!;
        tb.Focus();
        tb.SelectAll();
    }

    private void CommitQty(int idx)
    {
        if (idx < 0 || idx >= _rowQtyInputs.Count || _rowQtyInputs[idx] == null) return;
        var tb = _rowQtyInputs[idx]!;
        if (!int.TryParse(tb.Text, out int qty) || qty < 1) qty = 1;
        var key = _analyteMap.Keys.ElementAtOrDefault(idx);
        // 변경 없으면 재빌드 회피 (포커스 빈번한 LostFocus에서 깜빡임 방지)
        if (key != null && _itemData.TryGetValue(key, out var cur) && cur.Qty == qty)
        {
            tb.Text = qty.ToString();
            return;
        }
        tb.Text = qty.ToString();
        if (key != null && _itemData.ContainsKey(key))
            _itemData[key] = (qty, _itemData[key].Price);
        RebuildItemList();
        _selectedRowIndex = idx;
    }

    private void MoveQtyFocus(int idx, int delta)
    {
        int next = idx + delta;
        if (next >= 0 && next < _rowQtyInputs.Count && _rowQtyInputs[next] != null)
            OpenQtyInput(next);
    }

    private void SelectGridRow(int idx)
    {
        _selectedRowIndex = idx;
        _keyNavShow2Index = idx;
    }

    // WasteAnalysisInputPage와 동일한 Tunnel KeyDown 핸들러 이식
    private void OnNavKeyDownTunnel(object? sender, KeyEventArgs e)
    {
        if (e.Handled) return;
        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

        // Shift+2: 네비게이션 모드 토글
        if (shift && e.Key == Key.D2)
        {
            if (_analyteMap.Count > 0)
            {
                _keyNavShow2 = !_keyNavShow2;
                if (_keyNavShow2)
                {
                    _keyNavShow2Index = 0;
                    _selectedRowIndex = 0;
                    OpenQtyInput(0);
                }
                else
                {
                    _keyNavShow2Index = -1;
                    _selectedRowIndex = -1;
                }
                e.Handled = true;
            }
            return;
        }

        // ── _keyNavShow2 모드: 방향키만 = 이동, Enter = 편집 ──
        if (!shift && _keyNavShow2)
        {
            // Up/Down → 현재 셀 저장 후 행 이동 + 자동으로 선택된 셀 열기
            if (e.Key == Key.Up || e.Key == Key.Down)
            {
                e.Handled = true;
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    // 현재 셀 편집 내용 저장
                    if (_selectedRowIndex >= 0)
                        CommitQty(_selectedRowIndex);

                    bool up = e.Key == Key.Up;
                    int newIdx = _selectedRowIndex + (up ? -1 : 1);
                    SelectGridRow(newIdx);
                    OpenQtyInput(newIdx);
                });
                return;
            }
            // Enter → TextBox 편집 중이면 내부 핸들러에 위임, 아니면 현재 셀 열기
            if (e.Key == Key.Enter)
            {
                if (e.Source is TextBox) return;
                e.Handled = true;
                OpenQtyInput(_selectedRowIndex);
                return;
            }
            // 숫자 입력 → 먼저 현재 셀 저장 후 즉시 편집 시작 + 해당 키를 TextBox에 전달
            if ((e.Key >= Key.D0 && e.Key <= Key.D9) || (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9))
            {
                e.Handled = true;
                // 현재 셀 편집 내용 저장
                if (_selectedRowIndex >= 0 && _rowQtyInputs.ElementAtOrDefault(_selectedRowIndex)?.IsVisible == true)
                    CommitQty(_selectedRowIndex);

                OpenQtyInput(_selectedRowIndex);
                var tb = _rowQtyInputs.ElementAtOrDefault(_selectedRowIndex);
                if (tb != null)
                {
                    char ch = e.Key switch
                    {
                        Key.D0 or Key.NumPad0 => '0', Key.D1 or Key.NumPad1 => '1',
                        Key.D2 or Key.NumPad2 => '2', Key.D3 or Key.NumPad3 => '3',
                        Key.D4 or Key.NumPad4 => '4', Key.D5 or Key.NumPad5 => '5',
                        Key.D6 or Key.NumPad6 => '6', Key.D7 or Key.NumPad7 => '7',
                        Key.D8 or Key.NumPad8 => '8', Key.D9 or Key.NumPad9 => '9',
                        _ => '\0'
                    };
                    if (ch != '\0')
                    {
                        tb.Text = ch.ToString();
                        tb.CaretIndex = 1;
                    }
                }
                return;
            }
            // Escape → 네비 모드 종료
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                _keyNavShow2 = false;
                _keyNavShow2Index = -1;
                _selectedRowIndex = -1;
                RebuildItemList();
                return;
            }
        }
    }

    // Show3(하단) 패널: 네비/인라인 편집 관련 코드 완전 제거



    private static TextBlock Cell(string text, double size, string color,
                                  int col, HorizontalAlignment ha = HorizontalAlignment.Left)
    {
        var tb = new TextBlock
        {
            Text = text, FontSize = size, FontFamily = Font,
            Foreground = Brush.Parse(color),
            HorizontalAlignment = ha, VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(6, 2),
            [Grid.ColumnProperty] = col,
        };
        return tb;
    }

    // ── DB row 에서 항목 로드 ─────────────────────────────────────────────
    // 단가/소계 버그 보완: 오작성 수정 시에도 항상 단가 정상 로드
    private void LoadItemsFromRow(Dictionary<string, string> row)
    {
        _itemData.Clear();
        _analyteMap.Clear();

        var meta = TestReportService.GetAnalyteMeta();
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "_id","rowid",
            "견적발행일자","업체명","약칭","대표자","견적요청담당",
            "담당자","담당자연락처","담당자 e-Mail","시료명","견적번호",
            "적용구분","견적작성","합계 금액",
            "수량","단가","소계","수량2","단가3","소계4",
        };

        // 업체명 우선순위: _company, _editingIssue, _carrotCompanyName
        var companyForPrices = _company?.C_CompanyName
            ?? _editingIssue?.업체명
            ?? (_carrotCompanyName.Length > 0 ? _carrotCompanyName : "");
        var prices = string.IsNullOrEmpty(companyForPrices)
            ? new Dictionary<string, decimal>()
            : QuotationService.GetPricesByCompany(companyForPrices);

        foreach (var kv in row)
        {
            var col = kv.Key;
            if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
            if (fixedCols.Contains(col)) continue;
            if (string.IsNullOrWhiteSpace(kv.Value)) continue;
            if (decimal.TryParse(kv.Value, NumberStyles.Any,
                CultureInfo.InvariantCulture, out var dv) && dv == 0) continue;

            int qty = 1;
            if (int.TryParse(kv.Value, out var qi)) qty = Math.Max(1, qi);

            var price = prices.TryGetValue(col, out var p) ? p : 0m;
            _itemData[col]  = (qty, price);
            _analyteMap[col] = meta.TryGetValue(col, out var m)
                ? m : new AnalysisItem { Analyte = col };
        }

        RebuildItemList();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  경고 체크
    // ══════════════════════════════════════════════════════════════════════
    private void CheckSampleDuplicate()
    {
        var sample = txbSampleName.Text?.Trim() ?? "";
        _sampleDuplicated = !string.IsNullOrEmpty(sample) &&
            _allIssues.Any(i => i.시료명.Equals(sample, StringComparison.OrdinalIgnoreCase)
                             && i.Id != (_editingIssue?.Id ?? 0));

        // 노란색 테두리
        txbSampleName.BorderBrush = _sampleDuplicated
            ? AppTheme.FgWarn
            : Brush.Parse("#444");

        UpdateWarning();
        UpdateSaveButton();
    }

    private void UpdateWarning()
    {
        txbWarning.Text      = "⚠️ 동일한 시료명의 발행내역이 이미 존재합니다.";
        txbWarning.IsVisible = _sampleDuplicated;
    }

    private void UpdateSaveButton()
    {
        btnSave.IsEnabled = !_sampleDuplicated;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  이벤트 핸들러
    // ══════════════════════════════════════════════════════════════════════

    // 견적번호 새로 생성
    private void BtnGenNo_Click(object? sender, RoutedEventArgs e)
        => txbQuotationNo.Text = GenerateNo();

    // 시료명 변경 → 중복 체크
    private void TxbSampleName_Changed(object? sender, TextChangedEventArgs e)
        => CheckSampleDuplicate();

    // 일괄 수량 적용
    private void BtnBulkQty_Click(object? sender, RoutedEventArgs e)
    {
        if (!int.TryParse(txbBulkQty.Text?.Trim(), out int qty) || qty < 0)
            qty = 1;

        foreach (var key in _itemData.Keys.ToList())
            _itemData[key] = (qty, _itemData[key].Price);

        RebuildItemList();
        Log($"일괄 수량 {qty} 적용");
    }

    // 항목 클릭 → 수량 편집 시작
    private void StartQtyEdit(string name, int currentQty)
    {
        _editingAnalyte      = name;
        txbEditingItem.Text  = name;
        txbQty.Text          = currentQty.ToString();
        pnlQtyEdit.IsVisible = true;
        txbQty.Focus();
        txbQty.SelectAll();
    }

    // 수량 적용
    private void BtnApplyQty_Click(object? sender, RoutedEventArgs e)
    {
        if (_editingAnalyte == null) return;
        if (!int.TryParse(txbQty.Text, out int qty) || qty < 0) qty = 1;

        if (qty == 0)
        {
            _itemData.Remove(_editingAnalyte);
            _analyteMap.Remove(_editingAnalyte);
        }
        else
        {
            var price = _itemData.TryGetValue(_editingAnalyte, out var d) ? d.Price : 0m;
            _itemData[_editingAnalyte] = (qty, price);
        }

        pnlQtyEdit.IsVisible = false;
        _editingAnalyte       = null;
        RebuildItemList();
    }

    // 초기화
    private void BtnClear_Click(object? sender, RoutedEventArgs e) => Clear();

    /// <summary>저장 완료 시 발생 — HistoryPanel 자동 새로고침용 (저장된 issue 전달)</summary>
    public event Action<ETA.Models.QuotationIssue>? SaveCompleted;

    /// <summary>ESC 취소 시 발생 — MainPage가 DetailPanel로 복귀하도록</summary>
    public event Action? EscapeCancelled;

    // 저장 (async — 프로그레스바 표시 후 DB 저장, 완료 후 목록 갱신)
    private async void BtnSave_Click(object? sender, RoutedEventArgs e)
    {
        Log("==== 저장 버튼 클릭 ====");
        Log($"  _sampleDuplicated={_sampleDuplicated}");

        if (_sampleDuplicated) { Log("  → 중복 시료 차단으로 중단"); return; }

        var companyName = _company?.C_CompanyName
                       ?? _editingIssue?.업체명
                       ?? (_carrotCompanyName.Length > 0 ? _carrotCompanyName : null)
                       ?? "";
        var abbr        = _company?.C_Abbreviation
                       ?? _editingIssue?.약칭
                       ?? (_carrotAbbr.Length > 0 ? _carrotAbbr : null)
                       ?? "";

        Log($"  companyName='{companyName}'  abbr='{abbr}'");
        Log($"  _itemData.Count={_itemData.Count}");
        Log($"  발행일={txbIssueDate.Text}  시료명={txbSampleName.Text}  번호={txbQuotationNo.Text}");
        Log($"  담당자={acbManager.Text}");

        if (string.IsNullOrEmpty(companyName)) { Log("  → 업체 미선택으로 중단"); return; }

        decimal total = _itemData.Values.Sum(d => d.Qty * d.Price);
        Log($"  total={total:N0}  editingIssue={_editingIssue?.Id.ToString() ?? "null"}");

        // ── 담당자 정보 확인창 ───────────────────────────────────────────
        var owner = TopLevel.GetTopLevel(this) as Window;
        var mgr = await ShowManagerConfirmAsync(
            owner,
            acbManager.Text     ?? "",
            txbManagerPhone.Text ?? "",
            txbManagerEmail.Text ?? "");
        if (mgr == null) { Log("  → 담당자 확인 취소"); return; }

        // 확인창에서 수정된 값으로 패널 필드도 갱신
        acbManager.Text      = mgr.Value.name;
        txbManagerPhone.Text = mgr.Value.phone;
        txbManagerEmail.Text = mgr.Value.email;

        var issue = new QuotationIssue
        {
            발행일   = txbIssueDate.Text   ?? DateTime.Today.ToString("yyyy-MM-dd"),
            업체명   = companyName,
            약칭     = abbr,
            시료명   = txbSampleName.Text  ?? "",
            견적번호 = txbQuotationNo.Text ?? GenerateNo(),
            견적구분 = "",
            담당자       = mgr.Value.name,
            담당자연락처 = mgr.Value.phone,
            담당자이메일 = mgr.Value.email,
            총금액   = total,
        };

        // ── 프로그레스바 표시 ────────────────────────────────────────────
        btnSave.IsEnabled       = false;
        prgSave.IsVisible       = true;
        prgSave.IsIndeterminate = true;

        // UI 갱신 프레임 양보
        await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(
            () => { }, Avalonia.Threading.DispatcherPriority.Render);

        bool ok = false;
        try
        {
            Log("  DB 저장 시작 (Task.Run)");
            // DB 작업을 백그라운드 스레드에서 실행
            ok = await System.Threading.Tasks.Task.Run(() =>
            {
                if (_editingIssue != null)
                {
                    Log($"  기존 건 삭제: id={_editingIssue.Id}");
                    QuotationService.Delete(_editingIssue.Id);
                }
                return QuotationService.Insert(issue, _itemData);
            });
        }
        catch (Exception ex)
        {
            Log($"  저장 예외: {ex.GetType().Name}: {ex.Message}");
            Log($"  StackTrace: {ex.StackTrace}");
        }
        finally
        {
            prgSave.IsIndeterminate = false;
            prgSave.IsVisible       = false;
            btnSave.IsEnabled       = !_sampleDuplicated;
        }

        Log(ok ? $"  → 저장 성공: {issue.견적번호}  항목{_itemData.Count}개" : "  → 저장 실패 (Insert 반환 false)");

        if (ok)
        {
            _editingIssue       = null;
            _carrotCompanyName  = "";
            _carrotAbbr         = "";
            txbMode.Text        = "";
            txbTitle.Text       = "📝  신규 견적 작성";
            txbQuotationNo.Text = GenerateNo();

            // 최신 발행 목록 갱신 (백그라운드에서 읽어와 UI 스레드에 반영)
            _allIssues = await System.Threading.Tasks.Task.Run(
                () => QuotationService.GetAllIssues());

            // HistoryPanel 자동 새로고침 트리거 (저장된 issue 전달)
            SaveCompleted?.Invoke(issue);
        }
    }

    private void BtnPreview_Click(object? sender, RoutedEventArgs e)
        => Log("미리보기 — 추후 연동");

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private static string GenerateNo()
        => DateTime.Now.ToString("yyyyMMdd-HHmmss");

    private static readonly string LogPath =
        Path.Combine(ETA.Services.Common.AppPaths.LogsDir, "Quotation.log");

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [NewPanel] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    // ── 담당자 확인창 ─────────────────────────────────────────────────────
    /// <summary>
    /// 저장 전 담당자/연락처/이메일 확인창.
    /// 확인 → (name, phone, email) 반환, 취소 → null 반환
    /// </summary>
    private static System.Threading.Tasks.Task<(string name, string phone, string email)?> ShowManagerConfirmAsync(
        Window? owner, string initName, string initPhone, string initEmail)
    {
        var tcs = new System.Threading.Tasks.TaskCompletionSource<(string, string, string)?>();

        var font = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        // 입력 필드
        var txName  = new TextBox { Text = initName,  Watermark = "담당자 이름",  FontFamily = font, FontSize = AppTheme.FontLG };
        var txPhone = new TextBox { Text = initPhone, Watermark = "연락처",        FontFamily = font, FontSize = AppTheme.FontLG };
        var txEmail = new TextBox { Text = initEmail, Watermark = "이메일",        FontFamily = font, FontSize = AppTheme.FontLG };

        // 경고 레이블
        var txWarn = new TextBlock
        {
            FontFamily = font, FontSize = AppTheme.FontBase,
            Foreground = Brushes.OrangeRed,
            IsVisible  = false,
            Text       = "담당자 이름과 연락처는 필수입니다.",
            Margin     = new Thickness(0, 4, 0, 0),
        };

        static Border MakeRow(string label, TextBox box, FontFamily f) => new()
        {
            Margin = new Thickness(0, 6, 0, 0),
            Child  = new StackPanel
            {
                Spacing = 4,
                Children =
                {
                    new TextBlock { Text = label, FontFamily = f, FontSize = AppTheme.FontBase, Foreground = AppRes("FgMuted") },
                    box,
                }
            }
        };

        var confirmBtn = new Button
        {
            Content     = "저장 확인",
            FontFamily  = font,
            FontSize    = AppTheme.FontLG,
            Padding     = new Thickness(20, 6),
            Background  = new SolidColorBrush(Color.Parse("#3a6ea5")),
            Foreground  = Brushes.White,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        var cancelBtn = new Button
        {
            Content    = "취소",
            FontFamily = font,
            FontSize   = AppTheme.FontLG,
            Padding    = new Thickness(12, 6),
            Background = AppTheme.BorderSubtle,
            Foreground = AppRes("FgMuted"),
            HorizontalAlignment = HorizontalAlignment.Right,
        };

        var dlg = new Window
        {
            Title                 = "담당자 정보 확인",
            Width                 = 340,
            SizeToContent         = SizeToContent.Height,
            CanResize             = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background            = AppTheme.BgPrimary,
            Content               = new StackPanel
            {
                Margin  = new Thickness(20, 16),
                Spacing = 0,
                Children =
                {
                    new TextBlock
                    {
                        Text       = "견적 발행 전 담당자 정보를 확인하세요.",
                        FontFamily = font,
                        FontSize   = AppTheme.FontMD,
                        Foreground = AppRes("AppFg"),
                        Margin     = new Thickness(0, 0, 0, 8),
                    },
                    MakeRow("견적요청담당자 *", txName,  font),
                    MakeRow("연락처 *",         txPhone, font),
                    MakeRow("이메일",            txEmail, font),
                    txWarn,
                    new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        HorizontalAlignment = HorizontalAlignment.Right,
                        Spacing = 8,
                        Margin  = new Thickness(0, 16, 0, 0),
                        Children = { cancelBtn, confirmBtn },
                    }
                }
            }
        };

        confirmBtn.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(txName.Text) || string.IsNullOrWhiteSpace(txPhone.Text))
            {
                txWarn.IsVisible = true;
                return;
            }
            tcs.TrySetResult((txName.Text.Trim(), txPhone.Text.Trim(), txEmail.Text?.Trim() ?? ""));
            dlg.Close();
        };
        cancelBtn.Click += (_, _) => { tcs.TrySetResult(null); dlg.Close(); };
        dlg.Closed      += (_, _) => tcs.TrySetResult(null);

        if (owner != null) dlg.ShowDialog(owner);
        else               dlg.Show();

        return tcs.Task;
    }
}
