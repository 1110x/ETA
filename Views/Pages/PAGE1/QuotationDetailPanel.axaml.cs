using Avalonia;
using ETA.Views;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Input;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE1;

public partial class QuotationDetailPanel : UserControl
{
    private QuotationIssue?             _current;
    private Dictionary<string, string> _cachedRow = new(StringComparer.OrdinalIgnoreCase);
    public  QuotationCheckPanel?  CheckPanel    { get; set; }
    /// <summary>🥕 당근 — 이 건을 재활용해서 신규 작성 (항목 복사, 번호·날짜 신규)</summary>
    public event Action<QuotationIssue>? CarrotRequested;
    /// <summary>✏️ 오작성 수정 — 시료명·견적번호·발행일·적용구분·업체명 등 메타 수정 (더 이상 사용 안 함)</summary>
    public event Action<QuotationIssue>? CorrectRequested;

    private bool _isEditMode = false;
    private string _currentCompany = "";  // 현재 선택된 업체명
    private List<string> _availableManagers = new();  // 업체별 담당자 목록
    private bool _settingManagerName = false;  // 프로그래매틱 설정 중 이벤트 방지
    private bool _isInitializing = false;  // 초기 로드 중인지 플래그

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 로그 ─────────────────────────────────────────────────────────────
    private static readonly string LogPath = Path.GetFullPath(
        Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Logs", "Quotation.log"));

    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText(LogPath, line + Environment.NewLine); } catch { }
        }
    }

    // ── 고정 컬럼 (항목 순회 제외) — Trim 된 실제 컬럼명 기준 ────────────
    private static readonly HashSet<string> FixedCols =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "_id","rowid",
            "견적발행일자","업체명","약칭","대표자","견적요청담당",
            "담당자","담당자연락처","담당자 e-Mail","시료명","견적번호",
            "적용구분","견적작성","합계 금액",   // ← 공백 없이 Trim된 값
            "수량","단가","소계","수량2","단가3","소계4",
        };

    public QuotationDetailPanel() => InitializeComponent();

    // ══════════════════════════════════════════════════════════════════════
    //  외부 연동
    // ══════════════════════════════════════════════════════════════════════
    public void ShowIssue(QuotationIssue issue)
    {
        _isInitializing = true;  // 초기 로드 플래그 설정
        _current = issue;
        _currentCompany = issue.업체명;
        txbEmpty.IsVisible  = false;
        spContent.IsVisible = true;
        spButtons.IsVisible = true;

        // 업체명/약칭은 읽기 전용 (선택된 업체 정보 반영)
        if (txbCompany is TextBlock tb1) tb1.Text = issue.업체명;
        if (txbAbbr is TextBlock tb2)
        {
            tb2.Text = issue.약칭;
            // 약칭 뱃지 배경색을 초성헬퍼로 설정
            try
            {
                var badgeColor = BadgeColorHelper.GetBadgeColor(issue.약칭);
                if (badgeAbbr is Border b)
                    b.Background = new SolidColorBrush(Color.Parse(badgeColor.Bg));
                if (tb2 is TextBlock)
                    tb2.Foreground = new SolidColorBrush(Color.Parse(badgeColor.Fg));
            }
            catch { }
        }
        txbSampleName.Text = issue.시료명;
        txbNo.Text         = issue.견적번호;
        txbDate.Text       = issue.발행일;

        // 적용구분은 더 이상 사용하지 않음

        txbAmount.Text     = issue.총금액 > 0 ? $"{issue.총금액:#,0} 원" : "—";

        // 읽기 전용 모드로 설정 (편집 불가)
        _isEditMode = false;
        SetEditMode(false);

        // 업체별 담당자 목록 로드
        LoadCompanyManagers(issue.업체명);

        Log($"=== ShowIssue rowid={issue.Id}  {issue.업체명}  {issue.시료명} ===");

        Dictionary<string, string> row;
        try
        {
            row = QuotationService.GetIssueRow(issue.Id);
            Log($"row 컬럼수={row.Count}");

            // 고정 컬럼 제외한 항목 컬럼 목록만 로그
            var itemCols = row.Keys
                .Where(k => !k.EndsWith("단가") && !k.EndsWith("소계") && !FixedCols.Contains(k))
                .OrderBy(k => k)
                .ToList();
            Log($"항목 컬럼 후보({itemCols.Count}개): {string.Join(", ", itemCols.Take(20))}");
        }
        catch (Exception ex)
        {
            Log($"GetIssueRow 오류: {ex.Message}");
            return;
        }

        _cachedRow = row;

        // 담당자 정보 설정
        // (LoadCompanyManagers에서 자동 설정됨)
        txbManagerPhone.Text = row.TryGetValue("담당자연락처",  out var ph) ? ph : issue.담당자연락처;
        txbManagerEmail.Text = row.TryGetValue("담당자 e-Mail", out var em) ? em : issue.담당자이메일;

        BuildItemLines(row);

        // CheckPanel 동기화: ShowIssue 호출 시 CurrentIssue 자동 설정
        if (CheckPanel != null)
        {
            CheckPanel.CurrentIssue = issue;
            ApplyCheckStates(row);
        }
        else
            Log("WARNING: CheckPanel == null");

        _isInitializing = false;  // 초기 로드 완료
    }

    public void Clear()
    {
        _current            = null;
        _cachedRow          = new(StringComparer.OrdinalIgnoreCase);
        txbEmpty.IsVisible  = true;
        spContent.IsVisible = false;
        spButtons.IsVisible = false;
        spItems.Children.Clear();
        _currentCompany = "";
        _availableManagers.Clear();
        cmbManagerName.Items.Clear();
        cmbManagerName.Text = "";
        txbManagerPhone.Text = "";
        txbManagerEmail.Text = "";
    }

    // ══════════════════════════════════════════════════════════════════════
    //  업체별 담당자 목록 로드 (견적요청담당)
    // ══════════════════════════════════════════════════════════════════════
    private async void LoadCompanyManagers(string companyName)
    {
        try
        {
            // 비동기로 담당자 목록 로드
            var managers = await Task.Run(() => QuotationService.GetDistinctManagersForCompany(companyName));
            _availableManagers = managers;

            // ComboBox에 항목 설정
            cmbManagerName.Items.Clear();
            foreach (var manager in managers)
                cmbManagerName.Items.Add(manager);

            // 현재 담당자 설정
            _settingManagerName = true;
            cmbManagerName.Text = _current?.담당자 ?? "";
            _settingManagerName = false;

            Log($"담당자 로드: {companyName} ({managers.Count}명)");
        }
        catch (Exception ex)
        {
            Log($"담당자 목록 로드 오류: {ex.Message}");
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  담당자 선택 (ComboBox SelectionChanged)
    // ══════════════════════════════════════════════════════════════════════
    private async void CmbManagerName_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_settingManagerName) return;

        var selectedName = cmbManagerName.SelectedItem as string ?? cmbManagerName.Text;
        if (string.IsNullOrEmpty(selectedName)) return;

        await OnManagerSelected(selectedName);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  담당자 선택 시 연락처/이메일 로드
    // ══════════════════════════════════════════════════════════════════════
    private async Task OnManagerSelected(string selectedName)
    {
        if (string.IsNullOrEmpty(selectedName) || string.IsNullOrEmpty(_currentCompany)) return;

        try
        {
            // DB에서 업체+담당자로 연락처/이메일 조회
            var (phone, email) = await Task.Run(() =>
                QuotationService.GetManagerContactInfo(_currentCompany, selectedName));

            txbManagerPhone.Text = phone ?? "";
            txbManagerEmail.Text = email ?? "";
            Log($"담당자 선택: {selectedName} ({_currentCompany})");
        }
        catch (Exception ex)
        {
            Log($"담당자 정보 로드 오류: {ex.Message}");
        }
    }

    /// <summary>체크박스 변경 실시간 반영 — 캐시된 행 기반으로 체크된 항목만 표시</summary>
    public void PreviewCheckedItems(IEnumerable<string> checkedAnalyteNames)
    {
        if (_cachedRow.Count == 0 || _isInitializing) return;  // 초기화 중에는 무시

        var checkedList = checkedAnalyteNames.ToList();
        Log($"🔍 PreviewCheckedItems: {checkedList.Count}개 항목 체크됨");

        var checkedSet = new HashSet<string>(checkedList, StringComparer.OrdinalIgnoreCase);
        // 캐시 복사 후 체크 해제된 메인 항목 컬럼값을 빈값으로 설정
        var preview = new Dictionary<string, string>(_cachedRow, StringComparer.OrdinalIgnoreCase);

        int hiddenCount = 0;
        foreach (var key in preview.Keys.ToList())
        {
            if (FixedCols.Contains(key)) continue;
            if (key.EndsWith("단가") || key.EndsWith("소계")) continue;
            if (!checkedSet.Contains(key))
            {
                preview[key] = "";  // HasNonZeroStr → false → 행 미표시
                hiddenCount++;
            }
        }

        Log($"   → {hiddenCount}개 항목 숨김");
        BuildItemLines(preview);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  항목 라인 빌드
    // ══════════════════════════════════════════════════════════════════════
    private int _selectedItemIndex = -1;
    private List<(Grid Grid, TextBox QuantityBox)> _itemRows = new();
    private List<TextBlock> _subtotalBlocks = new();  // 모든 소계 블록 추적

    private void BuildItemLines(Dictionary<string, string> row)
    {
        spItems.Children.Clear();
        _itemRows.Clear();
        _subtotalBlocks.Clear();
        _selectedItemIndex = -1;
        bool odd = false;

        // 항목명 → 약칭 매핑 로드
        var aliasMap = FacilityResultService.GetAnalyteAliasMap();

        foreach (var kv in row)
        {
            var col = kv.Key;   // 이미 Trim된 상태 (GetIssueRow에서 Trim)

            if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
            if (FixedCols.Contains(col)) continue;
            if (!HasNonZeroStr(kv.Value)) continue;

            row.TryGetValue(col + "단가", out var priceStr);
            row.TryGetValue(col + "소계", out var subStr);

            var grid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,60,80,80"),  // XAML 헤더와 동일하게
                Background        = Brush.Parse(odd ? "#1a1a28" : "#1e1e30"),
                Cursor            = new Cursor(Avalonia.Input.StandardCursorType.Hand),
                RowDefinitions    = new RowDefinitions("Auto"),
                Margin            = new Avalonia.Thickness(12, 2, 12, 2),
            };
            odd = !odd;

            // 컬럼 0: 항목명 + 약칭 뱃지 (수평 배치)
            var itemNamePanel = BuildItemNameWithBadge(col, aliasMap);
            Grid.SetColumn(itemNamePanel, 0);
            grid.Children.Add(itemNamePanel);

            // 컬럼 1: 수량 (TextBox, 편집 가능)
            var quantityBox = new TextBox
            {
                Text                = kv.Value ?? "",
                FontSize            = AppFonts.Base,
                FontFamily          = Font,
                Foreground          = Brush.Parse("#aaaaaa"),
                Background          = Brush.Parse("Transparent"),
                BorderThickness     = new Thickness(0),
                Margin              = new Avalonia.Thickness(0, 3),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment   = VerticalAlignment.Center,
                Padding             = new Thickness(4, 0),
                TextAlignment       = Avalonia.Media.TextAlignment.Right,
            };

            int rowIdx = _itemRows.Count;

            // 컬럼 3: 소계 (동적으로 업데이트되는 TextBlock)
            var subtotalBlock = new TextBlock
            {
                Text                = FmtNum(subStr),
                FontSize            = AppFonts.Base,
                FontFamily          = Font,
                Foreground          = Brush.Parse("#88cc88"),
                Margin              = new Avalonia.Thickness(0, 3),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center,
                TextTrimming        = Avalonia.Media.TextTrimming.CharacterEllipsis,
            };
            Grid.SetColumn(subtotalBlock, 3);

            // 수량 변경 이벤트: 수량 × 단가 = 소계 자동 계산
            quantityBox.TextChanged += (_, _) =>
            {
                if (int.TryParse(quantityBox.Text, out int qty) &&
                    decimal.TryParse((priceStr ?? "0").Replace(",", ""), out decimal price))
                {
                    // 소계 계산
                    decimal subtotal = qty * price;
                    // UI 업데이트
                    subtotalBlock.Text = subtotal == 0 ? "" : $"{subtotal:#,0}";
                    // 캐시 업데이트 (나중에 저장할 때 반영)
                    _cachedRow[col] = qty.ToString();
                    _cachedRow[col + "소계"] = subtotal.ToString("F0");
                }
                else
                {
                    subtotalBlock.Text = "";
                }
                // 합계 금액 다시 계산
                UpdateTotalAmount();
            };

            quantityBox.KeyDown += (_, e) =>
            {
                if (e.Key == Avalonia.Input.Key.Up)
                {
                    e.Handled = true;
                    if (rowIdx > 0) SelectItemRow(rowIdx - 1);
                }
                else if (e.Key == Avalonia.Input.Key.Down || e.Key == Avalonia.Input.Key.Enter)
                {
                    e.Handled = true;
                    if (rowIdx < _itemRows.Count - 1) SelectItemRow(rowIdx + 1);
                }
            };

            Grid.SetColumn(quantityBox, 1);
            grid.Children.Add(quantityBox);

            // 컬럼 2: 단가
            grid.Children.Add(Cell(FmtNum(priceStr), AppFonts.Base, "#aaaaaa", 2, HorizontalAlignment.Right));

            // 컬럼 3: 소계 (동적)
            grid.Children.Add(subtotalBlock);
            _subtotalBlocks.Add(subtotalBlock);

            int capturedIdx = rowIdx;
            grid.PointerPressed += (_, _) => SelectItemRow(capturedIdx);

            spItems.Children.Add(grid);
            _itemRows.Add((grid, quantityBox));
        }

        // 초기 합계 금액 계산
        UpdateTotalAmount();

        int cnt = spItems.Children.Count;
        Log($"BuildItemLines → {cnt}행");

        if (cnt == 0)
        {
            spItems.Children.Add(new TextBlock
            {
                Text       = "항목 데이터 없음",
                FontSize   = AppFonts.Base,
                FontFamily = Font,
                Foreground = AppTheme.FgDimmed,
                Margin     = new Avalonia.Thickness(12, 4),
            });
        }
    }

    private void UpdateTotalAmount()
    {
        decimal total = 0;
        foreach (var block in _subtotalBlocks)
        {
            var text = block.Text ?? "";
            text = text.Replace("₩", "").Replace(",", "").Trim();
            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var amt))
                total += amt;
        }
        // txbAmount는 XAML의 합계 금액 TextBlock
        txbAmount.Text = total == 0 ? "" : $"₩{total:#,0}";
    }

    private void SelectItemRow(int index)
    {
        if (index < 0 || index >= _itemRows.Count) return;

        // 이전 선택 해제
        if (_selectedItemIndex >= 0 && _selectedItemIndex < _itemRows.Count)
        {
            var (prevGrid, _) = _itemRows[_selectedItemIndex];
            prevGrid.Background = Brush.Parse(_selectedItemIndex % 2 == 0 ? "#1a1a28" : "#1e1e30");
        }

        _selectedItemIndex = index;
        var (grid, quantityBox) = _itemRows[index];

        // 새 선택 하이라이트
        grid.Background = Brush.Parse("#2a4a3a");
        quantityBox.Focus();
        quantityBox.SelectAll();
    }

    private StackPanel BuildItemNameWithBadge(string itemName, Dictionary<string, string> aliasMap)
    {
        var panel = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center,
        };

        // 약칭 뱃지
        var alias = aliasMap.TryGetValue(itemName, out var a) ? a : itemName;
        if (!string.IsNullOrEmpty(alias))
        {
            try
            {
                var badgeColor = BadgeColorHelper.GetBadgeColor(alias);
                var badge = new Border
                {
                    Height = 20,
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(5, 2),
                    Background = new SolidColorBrush(Color.Parse(badgeColor.Bg)),
                    Child = new TextBlock
                    {
                        Text = alias,
                        FontSize = AppFonts.XS,
                        FontWeight = FontWeight.Bold,
                        Foreground = new SolidColorBrush(Color.Parse(badgeColor.Fg)),
                        FontFamily = Font,
                        VerticalAlignment = VerticalAlignment.Center,
                    }
                };
                panel.Children.Add(badge);
            }
            catch { }
        }

        // 항목명
        var itemLabel = new TextBlock
        {
            Text = itemName,
            FontSize = AppFonts.MD,
            Foreground = Brush.Parse("#cccccc"),
            FontFamily = Font,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
        };
        panel.Children.Add(itemLabel);

        return panel;
    }

    private TextBlock Cell(string text, double size, string color,
                           int col, HorizontalAlignment ha = HorizontalAlignment.Left)
    {
        var tb = new TextBlock
        {
            Text                = text ?? "",
            FontSize            = size,
            FontFamily          = Font,
            Foreground          = Brush.Parse(color),
            Margin              = new Avalonia.Thickness(0, 3),
            HorizontalAlignment = ha,
            VerticalAlignment   = VerticalAlignment.Center,
            TextTrimming        = Avalonia.Media.TextTrimming.CharacterEllipsis,
        };
        Grid.SetColumn(tb, col);
        return tb;
    }

    // ══════════════════════════════════════════════════════════════════════
    //  체크박스 동기화
    // ══════════════════════════════════════════════════════════════════════
    private void ApplyCheckStates(Dictionary<string, string> row)
    {
        var names = CheckPanel!.GetAllAnalyteNames();
        Log($"ApplyCheckStates: CheckPanel항목={names.Count}개");

        int matched = 0, checkedCnt = 0;
        foreach (var name in names)
        {
            // row 키도 Trim됐고, name도 DB에서 온 값이므로 OrdinalIgnoreCase 비교로 충분
            bool has = HasNonZeroStr(GetVal(row, name))
                    || HasNonZeroStr(GetVal(row, name + "단가"))
                    || HasNonZeroStr(GetVal(row, name + "소계"));

            if (row.ContainsKey(name) || row.ContainsKey(name + "단가") || row.ContainsKey(name + "소계"))
                matched++;

            if (has) checkedCnt++;
            CheckPanel.SetChecked(name, has);
        }

        CheckPanel.SyncAllCategories();
        Log($"ApplyCheckStates 완료: 컬럼매칭={matched}  체크됨={checkedCnt}");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private static string GetVal(Dictionary<string, string> row, string key)
    {
        if (row.TryGetValue(key, out var v)) return v ?? "";
        // Trim 후 재시도
        var found = row.FirstOrDefault(kv =>
            string.Equals(kv.Key.Trim(), key.Trim(), StringComparison.OrdinalIgnoreCase));
        return found.Value ?? "";
    }

    private static bool HasNonZeroStr(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return false;
        // ₩ 기호, 쉼표 제거 후 파싱
        var clean = val.Replace("₩","").Replace(",","").Trim();
        if (decimal.TryParse(clean, NumberStyles.Any,
                             CultureInfo.InvariantCulture, out var d))
            return d != 0;
        return true;
    }

    private static string FmtNum(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return "";
        var clean = val.Replace("₩","").Replace(",","").Trim();
        if (decimal.TryParse(clean, NumberStyles.Any,
                             CultureInfo.InvariantCulture, out var d))
            return d == 0 ? "" : $"{d:#,0}";
        return val;
    }

    // ── 버튼 ─────────────────────────────────────────────────────────────
    // 🥕 당근: 이 건 재활용 (항목 복사, 번호·날짜는 신규 생성)
    private void BtnCarrot_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;
        CarrotRequested?.Invoke(_current);
    }

    // ✏️ 오작성 수정: 직접 편집 or 저장
    private void BtnCorrect_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;

        if (!_isEditMode)
        {
            // 편집 모드 활성화
            _isEditMode = true;
            SetEditMode(true);
        }
        else
        {
            // 편집 완료 → DB 저장
            try
            {
                // 업체명/약칭은 업체 정보에서 자동 설정되므로 수정하지 않음
                _current.시료명 = txbSampleName.Text ?? "";
                _current.견적번호 = txbNo.Text ?? "";
                _current.발행일 = txbDate.Text ?? "";

                // DB 업데이트
                bool ok = QuotationService.UpdateIssueMetadata(_current.Id, new Dictionary<string, object>
                {
                    { "시료명", _current.시료명 },
                    { "견적번호", _current.견적번호 },
                    { "견적발행일자", _current.발행일 },  // DB 컬럼명
                });

                if (ok)
                {
                    Log($"✅ 오작성 수정: {_current.견적번호}");
                    _isEditMode = false;
                    SetEditMode(false);
                }
                else
                {
                    Log($"❌ 오작성 수정 실패: {_current.견적번호}");
                }
            }
            catch (Exception ex)
            {
                Log($"❌ 오작성 수정 오류: {ex.Message}");
            }
        }
    }

    private void SetEditMode(bool enable)
    {
        // 업체명/약칭은 TextBlock이므로 항상 읽기 전용
        txbSampleName.IsReadOnly = !enable;
        txbNo.IsReadOnly = !enable;
        txbDate.IsReadOnly = !enable;
    }

    /// <summary>의뢰서 편집 패널 전환 요청 — MainPage가 구독</summary>
    public event Action<QuotationIssue, List<string>, HashSet<string>>? OrderRequestEditRequested;

    // 📋 의뢰서 작성
    private async void BtnOrderRequest_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null) return;

        // 견적서 분석항목 추출
        var quotedItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var row = QuotationService.GetIssueRow(_current.Id);
            foreach (var kv in row)
            {
                var col = kv.Key;
                if (col.EndsWith("단가") || col.EndsWith("소계")) continue;
                if (FixedCols.Contains(col)) continue;
                if (HasNonZeroStr(kv.Value)) quotedItems.Add(col);
            }
            Log($"[의뢰서] 견적항목 {quotedItems.Count}개");
        }
        catch (Exception ex) { Log($"[의뢰서] 항목 조회 오류: {ex.Message}"); }

        // 시료명 선택 팝업 (이동 가능, 최소/최대/닫기)
        var owner = TopLevel.GetTopLevel(this) as Window;
        var popup = new ETA.Views.OrderRequestWindow(_current, quotedItems);
        if (owner != null) await popup.ShowDialog(owner);
        else popup.Show();

        if (!popup.Confirmed || popup.SelectedSamples.Count == 0) return;

        // MainPage에 편집 패널 전환 요청
        OrderRequestEditRequested?.Invoke(_current, popup.SelectedSamples, quotedItems);
    }

    // 💾 저장: 수량 및 메타 정보 변경사항 저장
    private async void BtnSave_Click(object? sender, RoutedEventArgs e)
    {
        if (_current == null || _cachedRow.Count == 0) return;

        try
        {
            // 현재 rows와 XAML 필드에서 변경사항 수집
            var updates = new Dictionary<string, object>();

            // 편집된 시료명, 견적번호, 발행일 (필요시)
            if (_current.시료명 != txbSampleName.Text)
                updates["시료명"] = txbSampleName.Text ?? "";
            if (_current.견적번호 != txbNo.Text)
                updates["견적번호"] = txbNo.Text ?? "";
            if (_current.발행일 != txbDate.Text)
                updates["견적발행일자"] = txbDate.Text ?? "";

            // 항목별 수량 저장
            foreach (var kv in _cachedRow)
            {
                var col = kv.Key;
                if (FixedCols.Contains(col) || col.EndsWith("단가") || col.EndsWith("소계")) continue;
                updates[col] = kv.Value ?? "";
                if (_cachedRow.TryGetValue(col + "소계", out var subtotal))
                    updates[col + "소계"] = subtotal ?? "";
            }

            if (updates.Count == 0)
            {
                Log("💾 변경 사항이 없습니다.");
                return;
            }

            bool ok = QuotationService.UpdateIssueMetadata(_current.Id, updates);
            if (ok)
            {
                Log($"✅ 저장 완료: {_current.견적번호} ({updates.Count}개 필드)");
                // UI 업데이트
                _current.시료명 = txbSampleName.Text ?? "";
                _current.견적번호 = txbNo.Text ?? "";
                _current.발행일 = txbDate.Text ?? "";
            }
            else
            {
                Log($"❌ 저장 실패: {_current.견적번호}");
            }
        }
        catch (Exception ex)
        {
            Log($"❌ 저장 오류: {ex.Message}");
        }
    }

    private void BtnPrint_Click(object? sender, RoutedEventArgs e)
        => Log($"인쇄 → {_current?.견적번호}");

}
