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
    /// <summary>💾 저장 완료 후 Show1 트리뷰 리프레시 요청</summary>
    public event Action? SaveCompleted;

    private bool _isEditMode;
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

        // 업체별 단가 로드 및 _cachedRow에 추가 (Show4에서 단가 표시용)
        var priceMap = QuotationService.GetPricesByCompany(issue.업체명);
        Log($"업체 단가 로드: {issue.업체명} ({priceMap.Count}개)");
        foreach (var kv in priceMap)
        {
            var key = kv.Key;
            var price = kv.Value;
            // 해당 항목이 존재하고 수량이 있으면 단가/소계 추가
            if (_cachedRow.ContainsKey(key) && !string.IsNullOrWhiteSpace(_cachedRow[key]))
            {
                if (!int.TryParse(_cachedRow[key], out var qty)) qty = 1;
                _cachedRow[key + "단가"] = price.ToString("F0");
                _cachedRow[key + "소계"] = (qty * price).ToString("F0");
            }
        }

        // 담당자 정보 설정
        // (LoadCompanyManagers에서 자동 설정됨)
        txbManagerPhone.Text = row.TryGetValue("담당자연락처",  out var ph) ? ph : issue.담당자연락처;
        txbManagerEmail.Text = row.TryGetValue("담당자 e-Mail", out var em) ? em : issue.담당자이메일;

        BuildItemLines(_cachedRow);

        // CheckPanel 동기화: ShowIssue 호출 시 CurrentIssue 자동 설정
        if (CheckPanel != null)
        {
            CheckPanel.CurrentIssue = issue;
            ApplyCheckStates(row);
        }
        else
            Log("WARNING: CheckPanel == null");

        _isInitializing = false;  // 초기 로드 완료

        // PreviewCheckedItems는 _isInitializing=false 이후에 호출
        if (CheckPanel != null)
        {
            var checkedAnalysisItems = CheckPanel.GetSelected();
            if (checkedAnalysisItems.Any())
            {
                var checkedNames = checkedAnalysisItems.Select(ai => ai.Analyte).ToList();
                PreviewCheckedItems(checkedNames);
            }
        }
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

            Log($"담당자 로드: {companyName} ({managers.Count}명): {string.Join(", ", managers)}");

            // ComboBox에 항목 설정 (전체 목록 초기 표시)
            cmbManagerName.Items.Clear();
            foreach (var manager in managers)
                cmbManagerName.Items.Add(manager);

            // 현재 담당자 설정
            _settingManagerName = true;
            cmbManagerName.Text = _current?.담당자 ?? "";
            cmbManagerName.IsDropDownOpen = false;  // 초기에는 드롭다운 닫기
            _settingManagerName = false;
        }
        catch (Exception ex)
        {
            Log($"담당자 목록 로드 오류: {ex.Message}");
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  담당자 텍스트 입력 후 KeyUp (검색)
    // ══════════════════════════════════════════════════════════════════════
    private void CmbManagerName_KeyUp(object? sender, Avalonia.Input.KeyEventArgs e)
    {
        if (_settingManagerName || string.IsNullOrEmpty(_currentCompany)) return;

        // ArrowDown/Up은 자체 드롭다운 네비게이션이므로 무시
        if (e.Key == Avalonia.Input.Key.Up || e.Key == Avalonia.Input.Key.Down) return;

        var searchText = cmbManagerName.Text ?? "";

        if (string.IsNullOrEmpty(searchText))
        {
            // 검색 텍스트 없으면 전체 담당자 목록 표시
            cmbManagerName.Items.Clear();
            foreach (var manager in _availableManagers)
                cmbManagerName.Items.Add(manager);
            Log($"담당자 목록: 전체 {_availableManagers.Count}명");
            return;
        }

        // 입력된 텍스트로 필터링
        var filtered = _availableManagers
            .Where(m => m.Contains(searchText, StringComparison.OrdinalIgnoreCase))
            .ToList();

        cmbManagerName.Items.Clear();
        foreach (var manager in filtered)
            cmbManagerName.Items.Add(manager);

        Log($"담당자 검색: '{searchText}' → {filtered.Count}명: {string.Join(", ", filtered)}");

        // 필터 결과가 있으면 드롭다운 열기
        if (filtered.Count > 0)
            cmbManagerName.IsDropDownOpen = true;
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
        Log($"\n📊 === PreviewCheckedItems 호출 시작 ===");
        Log($"   _cachedRow.Count={_cachedRow.Count}, _isInitializing={_isInitializing}");

        if (_cachedRow.Count == 0 || _isInitializing)
        {
            Log($"   ❌ 반환: {(_isInitializing ? "초기화 중" : "_cachedRow가 비어있음")}");
            return;  // 초기화 중에는 무시
        }

        var checkedList = checkedAnalyteNames.ToList();
        Log($"   체크된 항목 ({checkedList.Count}개): {string.Join(", ", checkedList)}");

        var checkedSet = new HashSet<string>(checkedList, StringComparer.OrdinalIgnoreCase);
        // 캐시 복사 후 체크 해제된 메인 항목 컬럼값을 빈값으로 설정
        var preview = new Dictionary<string, string>(_cachedRow, StringComparer.OrdinalIgnoreCase);

        int hiddenCount = 0;
        var hiddenItems = new List<string>();
        int addedCount = 0;
        var addedItems = new List<string>();

        // 업체별 단가 로드
        var priceMap = QuotationService.GetPricesByCompany(_currentCompany);
        Log($"   업체 단가 로드: {_currentCompany} ({priceMap.Count}개 항목)");

        foreach (var key in preview.Keys.ToList())
        {
            if (FixedCols.Contains(key)) continue;
            if (key.EndsWith("단가") || key.EndsWith("소계")) continue;

            if (!checkedSet.Contains(key))
            {
                preview[key] = "";  // HasNonZeroStr → false → 행 미표시
                hiddenCount++;
                hiddenItems.Add(key);
            }
            else
            {
                // 체크된 항목인데 수량이 비어있으면 1로 설정하고, 단가도 추가
                if (string.IsNullOrWhiteSpace(preview[key]))
                {
                    preview[key] = "1";
                    addedCount++;
                    addedItems.Add(key);

                    // 단가 로드
                    if (priceMap.TryGetValue(key, out var price))
                    {
                        preview[key + "단가"] = price.ToString("F0");
                        preview[key + "소계"] = (1 * price).ToString("F0");
                        Log($"   ⚙️ '{key}' 추가: 수량=1, 단가={price:#,0}, 소계={price:#,0}");
                    }
                    else
                    {
                        Log($"   ⚙️ '{key}' 추가: 수량=1 (단가 없음)");
                    }
                }
            }
        }

        Log($"   숨길 항목 ({hiddenCount}개): {string.Join(", ", hiddenItems)}");
        Log($"   표시할 항목 ({checkedList.Count}개), 수량 추가 ({addedCount}개): {string.Join(", ", addedItems)}");

        // preview 변경사항을 _cachedRow에도 반영 (Show3에서 다시 선택해도 데이터 유지)
        foreach (var kv in preview)
        {
            _cachedRow[kv.Key] = kv.Value;
        }

        BuildItemLines(preview);
        Log($"📊 === PreviewCheckedItems 완료 ===\n");
    }

    // ══════════════════════════════════════════════════════════════════════
    //  항목 라인 빌드
    // ══════════════════════════════════════════════════════════════════════
    private int _selectedItemIndex = -1;
    private List<(Grid Grid, TextBox QuantityBox)> _itemRows = new();
    private List<TextBlock> _subtotalBlocks = new();  // 모든 소계 블록 추적

    private void BuildItemLines(Dictionary<string, string> row)
    {
        Log($"\n🔨 === BuildItemLines 시작 ===");
        Log($"   전체 행 데이터: {row.Count}개");

        // 단가/소계 컬럼 여부 확인
        var priceColsInRow = row.Keys.Where(k => k.EndsWith("단가")).ToList();
        var subtotalColsInRow = row.Keys.Where(k => k.EndsWith("소계")).ToList();
        Log($"   단가 컬럼: {priceColsInRow.Count}개, 소계 컬럼: {subtotalColsInRow.Count}개");

        spItems.Children.Clear();
        _itemRows.Clear();
        _subtotalBlocks.Clear();
        _selectedItemIndex = -1;
        bool odd = false;

        // 항목명 → 약칭 매핑 로드 (분석정보 테이블 기준)
        var aliasMap = ContractService.GetAnalyteAliasMap();

        int displayCount = 0;
        int skipCount = 0;
        var displayedItems = new List<string>();

        foreach (var kv in row)
        {
            var col = kv.Key;   // 이미 Trim된 상태 (GetIssueRow에서 Trim)

            if (col.EndsWith("단가") || col.EndsWith("소계"))
            {
                Log($"   ⊘ '{col}' 숨김 (단가/소계 컬럼)");
                skipCount++;
                continue;
            }
            if (FixedCols.Contains(col))
            {
                Log($"   ⊘ '{col}' 숨김 (고정컬럼)");
                skipCount++;
                continue;
            }
            if (!HasNonZeroStr(kv.Value))
            {
                Log($"   ⊘ '{col}' 숨김 (빈값)");
                skipCount++;
                continue;
            }

            displayCount++;
            displayedItems.Add(col);

            row.TryGetValue(col + "단가", out var priceStr);
            row.TryGetValue(col + "소계", out var subStr);

            Log($"   ✓ '{col}' 표시 (수량: {kv.Value}, 단가: {FmtNum(priceStr)}, 소계: {FmtNum(subStr)})");

            // ── 배지 기반 레이아웃 (Show4 스타일) ──────────────────────────
            int rowIdx = _itemRows.Count;
            int capturedIdx = rowIdx;
            string capturedColName = col;

            // 약칭 배지
            var alias = aliasMap.TryGetValue(col, out var a) ? a : col;
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

            // 수량 입력 박스
            var quantityBox = new TextBox
            {
                Text = kv.Value ?? "",
                Width = 50,
                FontSize = AppFonts.Base,
                FontFamily = Font,
                Foreground = Brush.Parse("#aaaaaa"),
                Background = Brush.Parse("Transparent"),
                BorderThickness = new Thickness(0),
                Padding = new Thickness(2, 0),
                TextAlignment = Avalonia.Media.TextAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                VerticalContentAlignment = VerticalAlignment.Center,
            };

            // 소계 표시 (동적 업데이트)
            var subtotalBlock = new TextBlock
            {
                Text = FmtNum(subStr),
                FontSize = AppFonts.Base,
                FontFamily = Font,
                Foreground = Brush.Parse("#88cc88"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                MinWidth = 60,
                TextAlignment = Avalonia.Media.TextAlignment.Right,
                Margin = new Thickness(0, 0, 8, 0),
            };
            _subtotalBlocks.Add(subtotalBlock);

            // 수량 변경 이벤트
            quantityBox.TextChanged += (_, _) =>
            {
                if (int.TryParse(quantityBox.Text, out int qty) &&
                    decimal.TryParse((priceStr ?? "0").Replace(",", ""), out decimal price))
                {
                    decimal subtotal = qty * price;
                    subtotalBlock.Text = subtotal == 0 ? "" : $"{subtotal:#,0}";
                    _cachedRow[col] = qty.ToString();
                    _cachedRow[col + "소계"] = subtotal.ToString("F0");
                }
                else
                {
                    subtotalBlock.Text = "";
                }
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

            // 항목 정보 Grid: [배지] [항목명] [수량] [단가] [소계]
            var contentGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*,70,70,70"),
                Margin = new Thickness(0),
                VerticalAlignment = VerticalAlignment.Center,
            };

            // 배지
            contentGrid.Children.Add(badge);
            Grid.SetColumn(badge, 0);

            // 항목명
            var itemNameTb = new TextBlock
            {
                Text = col,
                FontSize = AppFonts.MD,
                Foreground = Brush.Parse("#cccccc"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 0, 0),
                TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            };
            contentGrid.Children.Add(itemNameTb);
            Grid.SetColumn(itemNameTb, 1);

            // 수량 (중앙 정렬)
            quantityBox.HorizontalAlignment = HorizontalAlignment.Center;
            quantityBox.VerticalAlignment = VerticalAlignment.Center;
            quantityBox.Width = double.NaN;
            quantityBox.FontSize = AppFonts.MD;
            contentGrid.Children.Add(quantityBox);
            Grid.SetColumn(quantityBox, 2);

            // 단가
            var priceLabel = new TextBlock
            {
                Text = FmtNum(priceStr),
                FontSize = AppFonts.MD,
                Foreground = Brush.Parse("#888888"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            contentGrid.Children.Add(priceLabel);
            Grid.SetColumn(priceLabel, 3);

            // 소계
            subtotalBlock.FontSize = AppFonts.MD;
            subtotalBlock.HorizontalAlignment = HorizontalAlignment.Center;
            contentGrid.Children.Add(subtotalBlock);
            Grid.SetColumn(subtotalBlock, 4);

            // 감싸기 Border
            var itemBorder = new Border
            {
                Background = Brush.Parse(odd ? "#1a1a28" : "#1e1e30"),
                Padding = new Thickness(12, 2),
                Margin = new Thickness(0, 2),
                CornerRadius = new CornerRadius(4),
                Cursor = new Cursor(Avalonia.Input.StandardCursorType.Hand),
                Child = contentGrid,
            };
            odd = !odd;

            // 이벤트 처리 (더미 grid 대신 border에 적용)
            itemBorder.PointerPressed += (_, _) => SelectItemRow(capturedIdx);
            itemBorder.DoubleTapped += (_, _) =>
            {
                if (CheckPanel != null)
                {
                    CheckPanel.SetChecked(capturedColName, false);
                    Log($"   더블클릭: '{capturedColName}' 체크 해제");
                }
            };

            spItems.Children.Add(itemBorder);
            _itemRows.Add((contentGrid, quantityBox));
        }

        // 초기 합계 금액 계산
        UpdateTotalAmount();

        int cnt = spItems.Children.Count;
        Log($"🔨 === BuildItemLines 완료 ===");
        Log($"   표시된 항목: {displayCount}개, 숨김 항목: {skipCount}개");
        Log($"   표시 항목 목록: [{string.Join(", ", displayedItems)}]");
        Log($"   UI 라인: {cnt}개\n");

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

    private void SetEditMode(bool enable)
    {
        // 시료명/견적번호/발행일 항상 편집 가능
        txbSampleName.IsReadOnly = false;
        txbNo.IsReadOnly = false;
        txbDate.IsReadOnly = false;
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

            // DB 실제 컬럼 목록 조회
            var dbCols = QuotationService.GetIssueTableColumns();

            // 단가 없는 항목 수집 (수량이 있는데 단가가 0이거나 없는 항목)
            var priceMap = QuotationService.GetPricesByCompany(_currentCompany);
            var noPriceItems = new List<string>();
            // 단가 없어도 알림 제외 항목 (현장측정 항목 등 단가 없이 포함되는 항목)
            var noPriceExempt = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "온도", "전기전도도"
            };

            // 항목별 수량/소계 저장 (DB에 컬럼이 존재하는 것만)
            foreach (var kv in _cachedRow)
            {
                var col = kv.Key;
                if (FixedCols.Contains(col) || col.EndsWith("단가") || col.EndsWith("소계")) continue;
                if (!dbCols.Contains(col)) continue;

                var qtyVal = kv.Value;
                if (string.IsNullOrWhiteSpace(qtyVal)) continue; // 수량 없으면 skip

                // 수량은 있는데 단가가 없는 항목 체크 (예외 항목 제외)
                if (!noPriceExempt.Contains(col) && (!priceMap.ContainsKey(col) || priceMap[col] == 0))
                    noPriceItems.Add(col);

                updates[col] = int.TryParse(qtyVal, out var q) ? q : (object)DBNull.Value;

                // 소계: DB에 컬럼이 있고 값이 있을 때만
                var subCol = col + "소계";
                if (dbCols.Contains(subCol) && _cachedRow.TryGetValue(subCol, out var subtotal))
                    updates[subCol] = string.IsNullOrWhiteSpace(subtotal) ? (object)DBNull.Value
                                      : decimal.TryParse(subtotal, out var s) ? s : (object)DBNull.Value;
            }

            // 단가 없는 항목이 있으면 알림
            if (noPriceItems.Count > 0)
            {
                var owner = TopLevel.GetTopLevel(this) as Window;
                var dlg = new Window
                {
                    Title           = "단가 미설정 항목",
                    Width           = 420,
                    Height          = 240,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Background      = Brush.Parse("#1e1e2e"),
                    CanResize       = false,
                };

                var msg = $"아래 항목은 계약 단가가 설정되지 않았습니다.\n단가를 먼저 설정하거나, 그냥 저장하시겠습니까?\n\n{string.Join("\n", noPriceItems.Select(x => $"  • {x}"))}";

                var btnSave   = new Button { Content = "그냥 저장", Margin = new Thickness(0,0,8,0), Padding = new Thickness(16,6), Background = Brush.Parse("#2a4a2a"), Foreground = Brush.Parse("#aef0ae"), BorderThickness = new Thickness(0) };
                var btnCancel = new Button { Content = "취소 (단가 먼저 설정)", Padding = new Thickness(16,6), Background = Brush.Parse("#4a2a2a"), Foreground = Brush.Parse("#f0aeae"), BorderThickness = new Thickness(0) };

                bool confirmed = false;
                btnSave.Click   += (_, _) => { confirmed = true;  dlg.Close(); };
                btnCancel.Click += (_, _) => { confirmed = false; dlg.Close(); };

                dlg.Content = new StackPanel
                {
                    Margin   = new Thickness(20),
                    Spacing  = 16,
                    Children =
                    {
                        new TextBlock { Text = msg, FontSize = AppFonts.Base, FontFamily = Font,
                                        Foreground = Brush.Parse("#dddddd"), TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                        new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right,
                                         Children = { btnSave, btnCancel } }
                    }
                };

                if (owner != null) await dlg.ShowDialog(owner);
                else dlg.Show();

                if (!confirmed)
                {
                    Log($"💾 저장 취소 — 단가 없는 항목 {noPriceItems.Count}개");
                    return;
                }
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
                _current.시료명 = txbSampleName.Text ?? "";
                _current.견적번호 = txbNo.Text ?? "";
                _current.발행일 = txbDate.Text ?? "";
                SaveCompleted?.Invoke();  // Show1 트리뷰 리프레시
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
