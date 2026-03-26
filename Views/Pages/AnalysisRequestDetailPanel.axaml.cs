using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views.Pages;

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

    // ── 행 재사용 풀 ─────────────────────────────────────────────────────
    private readonly List<(Grid Row, TextBlock Name, TextBlock Result)> _rowPool = new();

    public AnalysisRequestDetailPanel()
    {
        Content = BuildUI();
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
            ColumnDefinitions = new ColumnDefinitions("*,80"),
            Margin = new Avalonia.Thickness(0, 0, 0, 2),
        };
        colHeader.Children.Add(ColLbl("항목명",   0));
        colHeader.Children.Add(ColLbl("결과/상태",1, HorizontalAlignment.Right));

        _spItems   = new StackPanel { Spacing = 0 };
        _spContent = new StackPanel
        {
            Spacing = 0,
            Children =
            {
                metaGrid,
                new Border { Height=1, Background=Brush.Parse("#333"),
                             Margin=new Avalonia.Thickness(0,0,0,6) },
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
        _txbEmpty.IsVisible  = false;
        _spContent.IsVisible = true;

        _txbAbbr.Text   = rec.약칭;
        _txbSample.Text = rec.시료명;
        _txbAccNo.Text  = rec.접수번호;
        _txbDate.Text   = rec.의뢰일;

        // [1] DB 조회를 백그라운드 스레드에서 실행
        Dictionary<string, string> row;
        try
        {
            row = await Task.Run(() => AnalysisRequestService.GetRecordRow(rec.Id));
        }
        catch
        {
            row = new Dictionary<string, string>();
        }

        // [2] 항목 라인 갱신 (행 풀 재사용)
        BuildItemLines(row);

        // [3] CheckPanel 동기화 — 배치로 처리 후 1회만 SyncAllCategories
        if (CheckPanel != null)
            await ApplyCheckStatesAsync(row);
    }

    public void Clear()
    {
        _txbEmpty.IsVisible  = true;
        _spContent.IsVisible = false;
        _spItems.Children.Clear();
        _rowPool.Clear();
    }

    // ══════════════════════════════════════════════════════════════════════
    //  [개선] 항목 라인 — 행 풀 재사용
    // ══════════════════════════════════════════════════════════════════════
    private void BuildItemLines(Dictionary<string, string> row)
    {
        // 표시할 데이터 먼저 필터링 (UI 전에 계산)
        var items = row
            .Where(kv =>
            {
                var col = kv.Key.Trim();
                return !FixedCols.Contains(col) && !string.IsNullOrWhiteSpace(kv.Value);
            })
            .ToList();

        int needed = items.Count;

        // 풀 부족분 생성
        while (_rowPool.Count < needed)
        {
            var g = new Grid { ColumnDefinitions = new ColumnDefinitions("*,80") };
            var n = new TextBlock
            {
                FontSize  = 11, FontFamily = Font,
                Foreground = BrushItemName,
                Margin    = new Avalonia.Thickness(8, 3),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                [Grid.ColumnProperty] = 0,
            };
            var r2 = new TextBlock
            {
                FontSize  = 10, FontFamily = Font,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(6, 3),
                [Grid.ColumnProperty] = 1,
            };
            g.Children.Add(n);
            g.Children.Add(r2);
            _rowPool.Add((g, n, r2));
        }

        // 기존 자식 제거 후 필요한 수만큼 재사용 행 추가
        _spItems.Children.Clear();

        if (needed == 0)
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
        for (int i = 0; i < needed; i++)
        {
            var (grid, nameBlock, resultBlock) = _rowPool[i];
            var (col, val) = (items[i].Key.Trim(), items[i].Value);

            bool analyzing = string.Equals(val.Trim(), "O",
                StringComparison.OrdinalIgnoreCase);

            grid.Background     = odd ? BrushOdd : BrushEven;
            nameBlock.Text      = col;
            resultBlock.Text    = analyzing ? "🔴 분석중" : val;
            resultBlock.Foreground = analyzing ? BrushAnalyzing : BrushResult;

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
