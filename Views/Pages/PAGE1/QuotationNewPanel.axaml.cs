using Avalonia;
using ETA.Views;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
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
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
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

        acbManager.TextChanged += (_, _) =>
        {
            var name = acbManager.Text?.Trim() ?? "";
            bool isKnown = !string.IsNullOrWhiteSpace(name)
                        && _knownManagers.Contains(name, StringComparer.OrdinalIgnoreCase);
            bool isNew   = !string.IsNullOrWhiteSpace(name) && !isKnown;

            if (isKnown)
            {
                // 기존 담당자 → DB에서 연락처/이메일 자동 조회
                var companyName = _company?.C_CompanyName ?? _carrotCompanyName;
                if (!string.IsNullOrEmpty(companyName))
                {
                    var (phone, email) = QuotationService.GetManagerContactInfo(companyName, name);
                    txbManagerPhone.Text = phone;
                    txbManagerEmail.Text = email;
                }
                pnlContactInfo.IsVisible = true;   // 확인용으로 표시 (읽기전용 느낌)
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
        };

        // 한글 IME 따라오기 방지: GotFocus 시 글자 수 저장, 이후 Background 우선순위로 초과분 제거
        AttachImeFix(txbSampleName);
        AttachImeFix(txbIssueDate);
        AttachImeFix(txbQuotationNo);
        AttachImeFix(txbBulkQty);
        AttachImeFix(txbQty);
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

    // 🥕 당근: 이 건 재활용 — 항목 복사, 번호·날짜는 신규
    public void LoadFromIssue(QuotationIssue issue)
    {
        _editingIssue      = null;   // 재활용은 신규 저장
        _company           = null;
        _carrotCompanyName = issue.업체명;
        _carrotAbbr        = issue.약칭;

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
            FontSize  = AppFonts.Base, FontFamily = Font,
            Foreground = AppTheme.FgSecondary,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            Margin    = new Avalonia.Thickness(8, 3),
            Cursor    = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
            [Grid.ColumnProperty] = 0,
        };
        Avalonia.Controls.ToolTip.SetTip(nameBlock,
            new TextBlock
            {
                Text = $"클릭 → 수량 변경\nES: {meta.ES}  단위: {meta.unit}",
                FontSize = AppFonts.Base,
                FontFamily = Font,
            });
        nameBlock.PointerPressed += (_, _) => StartQtyEdit(name, qty);
        grid.Children.Add(nameBlock);

        // 수량
        grid.Children.Add(Cell(qty.ToString(), AppFonts.Base,
            qty > 1 ? "#88d888" : "#888888", 1, HorizontalAlignment.Right));
        // 단가
        grid.Children.Add(Cell(price > 0 ? $"{price:#,0}" : "—", AppFonts.SM,
            "#888888", 2, HorizontalAlignment.Right));
        // 소계
        grid.Children.Add(Cell(sub > 0 ? $"{sub:#,0}" : "—", AppFonts.SM,
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
            "_id","rowid",
            "견적발행일자","업체명","약칭","대표자","견적요청담당",
            "담당자","담당자연락처","담당자 e-Mail","시료명","견적번호",
            "적용구분","견적작성","합계 금액",
            "수량","단가","소계","수량2","단가3","소계4",
        };

        // 업체별 계약 DB 단가 로드
        var companyForPrices = GetCurrentCompanyName();
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

    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Quotation.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] [NewPanel] {msg}";
        Debug.WriteLine(line);
        try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
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
