using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace ETA.Views.Pages;

public partial class QuotationNewPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // ── 상태 ──────────────────────────────────────────────────────────────
    private Contract?       _company;
    private QuotationIssue? _editingIssue;
    private string?         _editingAnalyte;

    // 항목명 → { qty, unitPrice }
    private readonly Dictionary<string, (int Qty, decimal Price)> _itemData = new();
    private readonly Dictionary<string, AnalysisItem>             _analyteMap = new();

    // 단가표: 적용구분 컬럼 → { Analyte → 단가 }
    private Dictionary<string, decimal> _priceMap = new();

    // 경고 상태
    private bool _sampleDuplicated = false;
    private bool _typeWarning      = false;

    // 최근 발행건 (중복/적용구분 비교용)
    private List<QuotationIssue> _allIssues = [];

    public QuotationNewPanel()
    {
        InitializeComponent();
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbQuotationNo.Text = GenerateNo();
        LoadContractTypes();
        _allIssues = QuotationService.GetAllIssues();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  적용구분 콤보박스 — DB C_ContractType 로드
    // ══════════════════════════════════════════════════════════════════════
    private void LoadContractTypes()
    {
        cmbType.Items.Clear();
        // DB 에서 C_ContractType 컬럼 고유값 가져오기
        var types = QuotationService.GetContractTypes();

        if (types.Count == 0)
        {
            // DB 업데이트 전이면 기본값 사용
            types = new List<string>
            {
                "FS100","FS100+","FS56","NFS56","FS55","FS52",
                "FSHN52","NFS50","NFS45","NFS39","NFS36","NFS36RE","FS25"
            };
        }

        foreach (var t in types)
            cmbType.Items.Add(new ComboBoxItem { Content = t });

        if (cmbType.Items.Count > 0)
            cmbType.SelectedIndex = 0;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  외부 API
    // ══════════════════════════════════════════════════════════════════════
    public void SetCompany(Contract company)
    {
        _company = company;
        txbCompany.Text =
            $"{company.C_CompanyName}  [{company.C_Abbreviation}]";

        // 해당 업체의 최근 발행건 확인해서 적용구분 경고 체크
        CheckTypeWarning();
    }

    public void SetSelectedAnalytes(List<AnalysisItem> items)
    {
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
                _itemData[item.Analyte] = (1, 0);
        }

        // 현재 선택된 적용구분으로 단가 갱신
        RefreshPrices();
        RebuildItemList();
    }

    public void LoadFromIssue(QuotationIssue issue)
    {
        _editingIssue = issue;
        _company      = null;

        txbTitle.Text       = "✏️  발행건 수정";
        txbMode.Text        = "수정 모드";
        txbCompany.Text     = $"{issue.업체명}  [{issue.약칭}]";
        txbQuotationNo.Text = GenerateNo();   // 신규 번호 자동 생성
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbSampleName.Text  = issue.시료명;

        // 적용구분 선택
        SelectCombo(cmbType, issue.견적구분);

        // 기존 항목 로드
        var row = QuotationService.GetIssueRow(issue.Id);
        LoadItemsFromRow(row);

        CheckSampleDuplicate();
        CheckTypeWarning();
    }

    public void Clear()
    {
        _editingIssue   = null;
        _company        = null;
        _editingAnalyte = null;
        _itemData.Clear();
        _analyteMap.Clear();
        _priceMap.Clear();
        _sampleDuplicated = false;
        _typeWarning      = false;

        txbTitle.Text       = "📝  신규 견적 작성";
        txbMode.Text        = "";
        txbCompany.Text     = "— 오른쪽에서 업체를 선택하세요 —";
        txbQuotationNo.Text = GenerateNo();
        txbIssueDate.Text   = DateTime.Today.ToString("yyyy-MM-dd");
        txbSampleName.Text  = "";
        if (cmbType.Items.Count > 0) cmbType.SelectedIndex = 0;

        spItems.Children.Clear();
        txbNoItems.IsVisible = true;
        pnlQtyEdit.IsVisible = false;
        txbTotal.Text        = "—";
        txbWarning.IsVisible = false;
        txbItemCount.Text    = "";
        UpdateSaveButton();

        _allIssues = QuotationService.GetAllIssues();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  단가 로드 (적용구분 컬럼 기준)
    // ══════════════════════════════════════════════════════════════════════
    private void RefreshPrices()
    {
        var colName = (cmbType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
        if (string.IsNullOrEmpty(colName)) return;

        _priceMap = QuotationService.GetPricesByColumn(colName);
        Log($"단가 로드: {colName} → {_priceMap.Count}개");

        // 기존 itemData 의 단가 갱신
        foreach (var key in _itemData.Keys.ToList())
        {
            var price = _priceMap.TryGetValue(key, out var p) ? p : 0;
            _itemData[key] = (_itemData[key].Qty, price);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  항목 목록 빌드
    // ══════════════════════════════════════════════════════════════════════
    private void RebuildItemList()
    {
        spItems.Children.Clear();
        txbNoItems.IsVisible = _analyteMap.Count == 0;
        txbItemCount.Text    = $"{_analyteMap.Count}개 항목";

        decimal totalAmount = 0;
        bool odd = false;

        foreach (var (name, meta) in _analyteMap)
        {
            var (qty, price) = _itemData.TryGetValue(name, out var d) ? d : (1, 0m);
            decimal sub = qty * price;
            totalAmount += sub;

            spItems.Children.Add(MakeItemRow(name, meta, qty, price, sub, odd));
            odd = !odd;
        }

        txbTotal.Text = _analyteMap.Count > 0
            ? $"{totalAmount:#,0} 원  ({_analyteMap.Count}항목)"
            : "—";
    }

    private Border MakeItemRow(string name, AnalysisItem meta,
                               int qty, decimal price, decimal sub, bool odd)
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,55,75,80"),
            Background = Brush.Parse(odd ? "#1a1a28" : "#1e1e30"),
        };

        // 항목명 — 클릭 시 수량 편집
        var nameBlock = new TextBlock
        {
            Text      = name,
            FontSize  = 11, FontFamily = Font,
            Foreground = Brush.Parse("#cccccc"),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            Margin    = new Avalonia.Thickness(8, 3),
            Cursor    = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            [Grid.ColumnProperty] = 0,
        };
        Avalonia.Controls.ToolTip.SetTip(nameBlock,
            $"클릭 → 수량 변경\nES: {meta.ES}  단위: {meta.unit}");
        nameBlock.PointerPressed += (_, _) => StartQtyEdit(name, qty);
        grid.Children.Add(nameBlock);

        // 수량
        grid.Children.Add(Cell(qty.ToString(), 11,
            qty > 1 ? "#88d888" : "#888888", 1, HorizontalAlignment.Right));
        // 단가
        grid.Children.Add(Cell(price > 0 ? $"{price:#,0}" : "—", 10,
            "#888888", 2, HorizontalAlignment.Right));
        // 소계
        grid.Children.Add(Cell(sub > 0 ? $"{sub:#,0}" : "—", 10,
            "#88cc88", 3, HorizontalAlignment.Right));

        return new Border { Child = grid };
    }

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
    private void LoadItemsFromRow(Dictionary<string, string> row)
    {
        _itemData.Clear();
        _analyteMap.Clear();

        var meta = TestReportService.GetAnalyteMeta();
        var fixedCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "견적발행일자","업체명","약칭","대표자","견적요청담당",
            "담당자연락처","담당자 e-Mail","시료명","견적번호",
            "적용구분","견적작성","합계 금액",
            "수량","단가","소계","수량2","단가3","소계4",
        };

        // 현재 적용구분으로 단가 로드
        var colName = (cmbType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
        var prices  = string.IsNullOrEmpty(colName)
            ? new Dictionary<string, decimal>()
            : QuotationService.GetPricesByColumn(colName);

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
            ? Brush.Parse("#f0c040")
            : Brush.Parse("#444");

        UpdateWarning();
        UpdateSaveButton();
    }

    private void CheckTypeWarning()
    {
        if (_company == null && _editingIssue == null) return;

        var companyName = _company?.C_CompanyName ?? _editingIssue?.업체명 ?? "";
        var latest = _allIssues
            .Where(i => i.업체명 == companyName)
            .OrderByDescending(i => i.발행일)
            .FirstOrDefault();

        var curType = (cmbType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "";
        _typeWarning = latest != null &&
                       !string.IsNullOrEmpty(latest.견적구분) &&
                       latest.견적구분 != curType;

        // 노란색 테두리
        cmbType.BorderBrush = _typeWarning
            ? Brush.Parse("#f0c040")
            : Brush.Parse("#444");

        UpdateWarning();
    }

    private void UpdateWarning()
    {
        var msgs = new List<string>();
        if (_sampleDuplicated)
            msgs.Add("⚠️ 동일한 시료명의 발행내역이 이미 존재합니다.");
        if (_typeWarning)
            msgs.Add("⚠️ 최근 발행건의 적용구분과 다릅니다.");

        if (msgs.Count > 0)
        {
            txbWarning.Text      = string.Join("\n", msgs);
            txbWarning.IsVisible = true;
        }
        else
        {
            txbWarning.IsVisible = false;
        }
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

    // 적용구분 변경 → 단가 갱신
    private void CmbType_Changed(object? sender, SelectionChangedEventArgs e)
    {
        RefreshPrices();
        RebuildItemList();
        CheckTypeWarning();
    }

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

    // 저장
    private void BtnSave_Click(object? sender, RoutedEventArgs e)
    {
        if (_sampleDuplicated) return;

        var companyName = _company?.C_CompanyName ?? _editingIssue?.업체명 ?? "";
        var abbr        = _company?.C_Abbreviation ?? _editingIssue?.약칭 ?? "";
        if (string.IsNullOrEmpty(companyName)) { Log("업체 미선택"); return; }

        decimal total = _itemData.Values.Sum(d => d.Qty * d.Price);

        var issue = new QuotationIssue
        {
            발행일   = txbIssueDate.Text   ?? DateTime.Today.ToString("yyyy-MM-dd"),
            업체명   = companyName,
            약칭     = abbr,
            시료명   = txbSampleName.Text  ?? "",
            견적번호 = txbQuotationNo.Text ?? GenerateNo(),
            견적구분 = (cmbType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "",
            총금액   = total,
        };

        if (_editingIssue != null)
            QuotationService.Delete(_editingIssue.Id);

        bool ok = QuotationService.Insert(issue);
        Log(ok ? $"저장 완료 → {issue.견적번호}" : "저장 실패");

        if (ok)
        {
            _editingIssue     = null;
            txbMode.Text      = "";
            txbTitle.Text     = "📝  신규 견적 작성";
            txbQuotationNo.Text = GenerateNo();
            _allIssues = QuotationService.GetAllIssues();
        }
    }

    private void BtnPreview_Click(object? sender, RoutedEventArgs e)
        => Log("미리보기 — 추후 연동");

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private static string GenerateNo()
        => DateTime.Now.ToString("yyyyMMdd-HHmmss");

    private static void SelectCombo(ComboBox cb, string value)
    {
        for (int i = 0; i < cb.Items.Count; i++)
            if ((cb.Items[i] as ComboBoxItem)?.Content?.ToString() == value)
            { cb.SelectedIndex = i; return; }
        if (cb.Items.Count > 0) cb.SelectedIndex = 0;
    }

    private void Log(string msg) => Debug.WriteLine($"[NewPanel] {msg}");
}
