using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// Content2: 분석의뢰 레코드 세부 표시 패널
/// 성능 개선:
///   - BuildItemLines: 컨트롤 풀 재사용 (Clear+재생성 대신 기존 행 갱신)
///   - ApplyCheckStates: UI 스레드 배치 처리 (Dispatcher.UIThread.InvokeAsync)
///   - SyncAllCategories: ShowRecord 완료 후 1회만 호출
/// </summary>
public class AnalysisRequestDetailPanel : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // 미리 파싱해 둔 브러시 (매 행마다 Brush.Parse 호출 비용 제거)
    private static readonly IBrush BrushOdd      = Brush.Parse("#1a1a28");
    private static readonly IBrush BrushEven     = Brush.Parse("#1e1e30");
    private static readonly IBrush BrushItemName = Brush.Parse("#cccccc");
    private static readonly IBrush BrushAnalyzing= Brush.Parse("#ff6666");
    private static readonly IBrush BrushResult   = Brush.Parse("#88cc88");
    private static readonly IBrush BrushEmpty    = Brush.Parse("#555");

    public QuotationCheckPanel? CheckPanel { get; set; }

    private static readonly HashSet<string> FixedCols =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // 내부 키 컬럼 (MariaDB: _id, SQLite: rowid)
            "_id","rowid",
            "약칭","시료명","접수번호","의뢰일","업체명","대표자",
            "담당자","연락처","이메일","견적번호","비고",
            // 분석의뢰및결과 고정 컬럼 추가
            "채취일자","채취시간","의뢰사업장","입회자",
            "시료채취자-1","시료채취자-2","방류허용기준 적용유무",
            "정도보증유무","분석완료일자","견적구분",
        };

    private TextBlock  _txbAbbr    = new();
    private TextBlock  _txbSample  = new();
    private TextBlock  _txbAccNo   = new();
    private TextBlock  _txbDate    = new();
    private StackPanel _spItems    = new();
    private TextBlock  _txbEmpty   = new();
    private StackPanel _spContent  = new();

    // ── 행 재사용 풀 (배지 추가 후 미사용 — 호환성 유지) ────────────────
    private readonly List<(Grid Row, TextBlock Name, TextBlock Result)> _rowPool = new();

    // 분석항목 약칭 (전체명 → 약칭): "생물화학적 산소요구량" → "BOD"
    private Dictionary<string, string> _shortNames = new(StringComparer.OrdinalIgnoreCase);

    // -- 방류기준 콤보박스 -----------------------------------------------
    private ComboBox _cmbDischarge = new();
    private int      _currentRecId = -1;
    private bool     _suppressDischargeEvent = false;

    public AnalysisRequestDetailPanel()
    {
        Content = BuildUI();
        LoadShortNames();
    }

    private void LoadShortNames()
    {
        try
        {
            // 분장표준처리 에서 약칭 행 직접 조회: 컬럼헤드 → 약칭
            _shortNames = AnalysisRequestService.GetShortNames();
        }
        catch { /* 실패해도 배지 기본값 사용 */ }
    }

    private Control BuildUI()
    {
        var header = new TextBlock
        {
            Text = "🧪  분석의뢰 상세",
            FontSize = 13, FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = Brush.Parse("#e0e0e0"),
            Margin = new Avalonia.Thickness(0, 0, 0, 4),
        };

        _txbAbbr   = MetaVal(); _txbSample = MetaVal();
        _txbAccNo  = MetaVal(); _txbDate   = MetaVal();

        var metaGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,*,16,Auto,*"),
            RowDefinitions    = new RowDefinitions("Auto,Auto"),
            Margin = new Avalonia.Thickness(0, 0, 0, 8),
        };
        AddMeta(metaGrid, "약칭",     _txbAbbr,   0, 0);
        AddMeta(metaGrid, "시료명",   _txbSample, 0, 3);
        AddMeta(metaGrid, "접수번호", _txbAccNo,  1, 0);
        AddMeta(metaGrid, "의뢰일",   _txbDate,   1, 3);

        var colHeader = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("64,*,80"),
            Margin = new Avalonia.Thickness(0, 0, 0, 2),
        };
        colHeader.Children.Add(ColLbl("",        0));
        colHeader.Children.Add(ColLbl("항목명",   1));
        colHeader.Children.Add(ColLbl("결과/상태",2, HorizontalAlignment.Right));

        // ── 방류기준 콤보박스 ────────────────────────────────────────────
        _cmbDischarge = new ComboBox
        {
            FontFamily            = Font,
            FontSize              = 12,
            HorizontalAlignment   = HorizontalAlignment.Stretch,
            PlaceholderText       = "방류기준 선택...",
            Margin                = new Avalonia.Thickness(0, 0, 0, 6),
        };
        _cmbDischarge.SelectionChanged += OnDischargeSelectionChanged;

        var dischargeRow = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,*"),
            Margin = new Avalonia.Thickness(0, 0, 0, 0),
        };
        dischargeRow.Children.Add(new TextBlock
        {
            Text              = "법정방류기준",
            FontFamily        = Font,
            FontSize          = 11,
            Foreground        = Brush.Parse("#aaaacc"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Avalonia.Thickness(0, 0, 8, 0),
            [Grid.ColumnProperty] = 0,
        });
        dischargeRow.Children.Add(new Border
        {
            [Grid.ColumnProperty] = 1,
            Child = _cmbDischarge,
        });

        _spItems   = new StackPanel { Spacing = 0 };
        _spContent = new StackPanel
        {
            Spacing = 0,
            Children =
            {
                metaGrid,
                dischargeRow,
                new Border { Height=1, Background=Brush.Parse("#333"),
                             Margin=new Avalonia.Thickness(0,6,0,6) },
                colHeader,
                _spItems,
            },
        };

        _txbEmpty = new TextBlock
        {
            Text       = "좌측에서 분석의뢰 항목을 선택하세요.",
            FontSize   = 11, FontFamily = Font,
            Foreground = Brush.Parse("#555"),
            Margin     = new Avalonia.Thickness(12, 20),
        };

        return new Border
        {
            Padding = new Avalonia.Thickness(10),
            Child   = new Grid
            {
                RowDefinitions = new RowDefinitions("Auto,Auto,*"),
                Children =
                {
                    header,
                    new Border { [Grid.RowProperty]=1, Height=1,
                                 Background=Brush.Parse("#333"),
                                 Margin=new Avalonia.Thickness(0,0,0,6) },
                    new ScrollViewer
                    {
                        [Grid.RowProperty] = 2,
                        VerticalScrollBarVisibility =
                            Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                        Content = new StackPanel
                        {
                            Children = { _txbEmpty, _spContent }
                        }
                    }
                }
            }
        };
    }

    // ══════════════════════════════════════════════════════════════════════
    //  외부 API — async로 DB + UI 분리
    // ══════════════════════════════════════════════════════════════════════
    public async void ShowRecord(AnalysisRequestRecord rec)
    {
        _currentRecId = rec.Id;
        _txbEmpty.IsVisible  = false;
        _spContent.IsVisible = true;

        _txbAbbr.Text   = rec.약칭;
        _txbSample.Text = rec.시료명;
        _txbAccNo.Text  = rec.접수번호;
        _txbDate.Text   = rec.의뢰일;

        // [1] DB 조회를 백그라운드 스레드에서 실행
        Dictionary<string, string> row;
        List<string> dischargeNames;
        try
        {
            (row, dischargeNames) = await Task.Run(() =>
            (
                AnalysisRequestService.GetRecordRow(rec.Id),
                AnalysisRequestService.GetDischargeStandardNames()
            ));
        }
        catch
        {
            row           = new Dictionary<string, string>();
            dischargeNames = new List<string>();
        }

        // [1-1] 방류기준 콤보박스 채우기
        _suppressDischargeEvent = true;
        _cmbDischarge.Items.Clear();
        foreach (var name in dischargeNames)
            _cmbDischarge.Items.Add(name);
        row.TryGetValue("방류허용기준 적용유무", out var currentStd);
        _cmbDischarge.SelectedItem = dischargeNames.FirstOrDefault(
            n => string.Equals(n, currentStd, StringComparison.OrdinalIgnoreCase));
        _suppressDischargeEvent = false;

        // [2] 항목 라인 갱신 (행 풀 재사용)
        BuildItemLines(row);

        // [3] CheckPanel 동기화 — 배치로 처리 후 1회만 SyncAllCategories
        if (CheckPanel != null)
            await ApplyCheckStatesAsync(row);
    }

    private void OnDischargeSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_suppressDischargeEvent) return;
        if (_currentRecId < 0) return;
        var selected = _cmbDischarge.SelectedItem as string;
        if (selected == null) return;
        _ = Task.Run(() => AnalysisRequestService.UpdateDischargeStandard(_currentRecId, selected));
    }

    public void Clear()
    {
        _currentRecId        = -1;
        _txbEmpty.IsVisible  = true;
        _spContent.IsVisible = false;
        _spItems.Children.Clear();
        _rowPool.Clear();
    }

    /// <summary>체크박스 변경 실시간 반영 — DB 재조회 없이 체크된 항목만 '분석중'으로 표시</summary>
    public void PreviewCheckedItems(IEnumerable<string> checkedAnalyteNames)
    {
        _txbEmpty.IsVisible  = false;
        _spContent.IsVisible = true;
        var dict = checkedAnalyteNames.ToDictionary(
            n => n, _ => "O", StringComparer.OrdinalIgnoreCase);
        BuildItemLines(dict);
    }

    // ══════════════════════════════════════════════════════════════════════
    //  항목 라인 빌드 — 카테고리 배지 포함
    // ══════════════════════════════════════════════════════════════════════
    private void BuildItemLines(Dictionary<string, string> row)
    {
        var items = row
            .Where(kv => !FixedCols.Contains(kv.Key.Trim()) &&
                         !string.IsNullOrWhiteSpace(kv.Value))
            .ToList();

        _spItems.Children.Clear();

        if (items.Count == 0)
        {
            _spItems.Children.Add(new TextBlock
            {
                Text       = "항목 데이터 없음",
                FontSize   = 11, FontFamily = Font,
                Foreground = BrushEmpty,
                Margin     = new Avalonia.Thickness(12, 4),
            });
            return;
        }

        bool odd = false;
        foreach (var kv in items)
        {
            var col       = kv.Key.Trim();
            var val       = kv.Value;
            bool analyzing = string.Equals(val.Trim(), "O",
                StringComparison.OrdinalIgnoreCase);

            // 배지 색상 — Category 기반(있으면), 없으면 항목명 기반
            var meta     = CheckPanel?.GetItem(col);
            var badgeKey = meta?.Category is { Length: > 0 } cat ? cat : col;
            var (bg, fg) = BadgeColorHelper.GetBadgeColor(badgeKey);
            // 배지 텍스트 — 분장표준처리 약칭 우선, 없으면 항목명 앞 2자
            var badgeText = _shortNames.TryGetValue(col, out var sn) && sn.Length > 0
                ? sn
                : (col.Length <= 3 ? col : col[..2]);

            var grid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("64,*,80"),
                Background        = odd ? BrushOdd : BrushEven,
            };

            // Col 0 — 카테고리 배지
            grid.Children.Add(new Border
            {
                Background        = Brush.Parse(bg),
                CornerRadius      = new Avalonia.CornerRadius(3),
                Padding           = new Avalonia.Thickness(3, 1),
                Margin            = new Avalonia.Thickness(4, 2),
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 0,
                Child = new TextBlock
                {
                    Text                = badgeText,
                    FontSize            = 9, FontFamily = Font,
                    Foreground          = Brush.Parse(fg),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment   = VerticalAlignment.Center,
                },
            });

            // Col 1 — 항목명
            grid.Children.Add(new TextBlock
            {
                Text              = col,
                FontSize          = 11, FontFamily = Font,
                Foreground        = BrushItemName,
                Margin            = new Avalonia.Thickness(4, 3),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
                [Grid.ColumnProperty] = 1,
            });

            // Col 2 — 결과값
            grid.Children.Add(new TextBlock
            {
                Text                = analyzing ? "🔴 분석중" : val,
                FontSize            = 10, FontFamily = Font,
                Foreground          = analyzing ? BrushAnalyzing : BrushResult,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin              = new Avalonia.Thickness(6, 3),
                [Grid.ColumnProperty] = 2,
            });

            _spItems.Children.Add(grid);
            odd = !odd;
        }
    }


    // ══════════════════════════════════════════════════════════════════════
    //  [개선] CheckPanel 동기화 — 배치 처리 후 SyncAllCategories 1회
    // ══════════════════════════════════════════════════════════════════════
    private async Task ApplyCheckStatesAsync(Dictionary<string, string> row)
    {
        if (CheckPanel == null) return;

        // UI 스레드 블로킹 최소화:
        // 변경이 필요한 항목만 골라낸 뒤 한 번에 적용
        var names = CheckPanel.GetAllAnalyteNames();

        // 백그라운드에서 체크 여부 계산
        var changes = await Task.Run(() =>
            names.Select(name =>
            {
                bool has = row.TryGetValue(name, out var v) &&
                           !string.IsNullOrWhiteSpace(v);
                return (name, has);
            }).ToList()
        );

        // UI 스레드에서 일괄 적용 (_suspendEvents 이미 있으므로 이벤트 미발생)
        foreach (var (name, has) in changes)
            CheckPanel.SetChecked(name, has);

        // SyncAllCategories 는 전체 완료 후 딱 1회
        CheckPanel.SyncAllCategories();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════════════
    private static TextBlock MetaVal() => new()
    {
        FontSize   = 11,
        FontFamily = new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
        Foreground = Brush.Parse("#dddddd"),
        VerticalAlignment = VerticalAlignment.Center,
    };

    private static void AddMeta(Grid g, string label, TextBlock val, int row, int valCol)
    {
        var lbl = new TextBlock
        {
            Text      = label + " : ",
            FontSize  = 10,
            FontFamily = new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = Brush.Parse("#888"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Avalonia.Thickness(0, 2, 0, 2),
        };
        Grid.SetRow(lbl, row); Grid.SetColumn(lbl, valCol);
        Grid.SetRow(val, row); Grid.SetColumn(val, valCol + 1);
        g.Children.Add(lbl);
        g.Children.Add(val);
    }

    private static TextBlock ColLbl(string t, int col,
        HorizontalAlignment ha = HorizontalAlignment.Left)
    {
        var tb = new TextBlock
        {
            Text      = t, FontSize = 10,
            FontFamily = new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
            Foreground = Brush.Parse("#666"),
            HorizontalAlignment = ha,
            Margin = new Avalonia.Thickness(8, 0),
            [Grid.ColumnProperty] = col,
        };
        return tb;
    }
}
