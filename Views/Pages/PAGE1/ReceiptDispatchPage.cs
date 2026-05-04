using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SS = DocumentFormat.OpenXml.Spreadsheet;
using ETA.Models;
using ETA.Services.Common;

namespace ETA.Views.Pages.PAGE1;

/// <summary>측정대행 접수/발송 대장 — 엑셀 `접수발송대장` 시트 1:1 매핑 (독립 테이블, 직접 입력)</summary>
public class ReceiptDispatchPage : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static readonly string ColDefs = "180,90,*,180,*,90,40";

    private readonly StackPanel _list  = new() { Spacing = 0 };
    private readonly TextBox    _search;
    private readonly TextBlock  _count = new()
    {
        FontSize = AppTheme.FontSM, FontFamily = Font,
        Foreground = AppTheme.FgDimmed,
        Margin = new Thickness(0, 4, 0, 4),
    };

    private List<ReceiptDispatchEntry> _cache = new();
    private const int MaxRender = 500;

    // ── Show1 의뢰목록 (좌측 패널) — 외부에서 가져다가 Show1 에 배치 ────────
    public Control RequestListControl { get; private set; } = null!;
    private readonly StackPanel _reqList = new() { Spacing = 2 };
    private readonly TextBox    _reqSearch;
    private readonly TextBlock  _reqCount = new()
    {
        FontSize = AppTheme.FontSM, FontFamily = Font,
        Foreground = AppTheme.FgDimmed,
        Margin = new Thickness(0, 4, 0, 4),
    };
    private List<AnalysisRequestRecord> _reqCache = new();

    public ReceiptDispatchPage()
    {
        _reqSearch = new TextBox
        {
            Watermark = "🔍  의뢰 검색 (약칭/시료명/접수번호)...",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Background = AppTheme.BgPrimary, Foreground = AppTheme.FgPrimary,
            BorderBrush = AppTheme.BorderSubtle, BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 4),
        };
        _reqSearch.TextChanged += (_, _) => RenderReqList();
        RequestListControl = BuildRequestListControl();
        _search = new TextBox
        {
            Watermark = "🔍  접수번호 / 시료명 / 업체명 / 분석항목 검색...",
            FontFamily = Font, FontSize = AppTheme.FontMD,
            Background = AppTheme.BgPrimary, Foreground = AppTheme.FgPrimary,
            BorderBrush = AppTheme.BorderSubtle, BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(8, 5),
        };
        _search.TextChanged += (_, _) => Render();
        Content = BuildUI();
        _ = LoadAsync();
        _ = LoadReqListAsync();
    }

    // ── Show1 의뢰목록 ──────────────────────────────────────────────────────
    private Control BuildRequestListControl()
    {
        var title = new TextBlock
        {
            Text = "📋  의뢰목록 (클릭하여 대장에 추가)",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            FontFamily = Font, Foreground = AppTheme.FgPrimary,
            Margin = new Thickness(0, 0, 0, 6),
        };
        var sv = new ScrollViewer
        {
            Content = _reqList,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
        };
        var root = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*"),
            Margin = new Thickness(8),
        };
        Grid.SetRow(title,      0); root.Children.Add(title);
        Grid.SetRow(_reqSearch, 1); root.Children.Add(_reqSearch);
        Grid.SetRow(_reqCount,  2); root.Children.Add(_reqCount);
        Grid.SetRow(sv,         3); root.Children.Add(sv);
        return root;
    }

    private async Task LoadReqListAsync()
    {
        _reqCount.Text = "📥  로딩 중...";
        try
        {
            _reqCache = await Task.Run(() =>
                ETA.Services.SERVICE1.AnalysisRequestService.GetAllRecords());
            RenderReqList();
        }
        catch (Exception ex)
        {
            _reqCount.Text = $"❌ {ex.Message}";
        }
    }

    private void RenderReqList()
    {
        _reqList.Children.Clear();
        var q = _reqSearch.Text?.Trim().ToLowerInvariant() ?? "";
        var filtered = string.IsNullOrEmpty(q) ? _reqCache : _reqCache.Where(r =>
            r.약칭.ToLower().Contains(q) ||
            r.시료명.ToLower().Contains(q) ||
            r.접수번호.ToLower().Contains(q)).ToList();
        _reqCount.Text = $"의뢰 {_reqCache.Count}건  /  표시 {filtered.Count}건";

        // 채취일자별 그룹핑 (DESC) — 그룹 헤더 클릭 시 일괄 추가
        var groups = filtered
            .GroupBy(r => (r.채취일자 ?? "").Length >= 10 ? r.채취일자[..10] : r.채취일자 ?? "")
            .OrderByDescending(g => g.Key)
            .Take(60);  // 너무 많으면 최근 60일 분만

        foreach (var g in groups)
        {
            _reqList.Children.Add(BuildDateHeader(g.Key, g.Count()));
            foreach (var r in g.Take(50))
                _reqList.Children.Add(BuildReqCard(r));
        }
    }

    private Control BuildDateHeader(string date, int count)
    {
        var capDate = date;
        var btnBatch = new Button
        {
            Content = $"📥 {count}건 일괄추가",
            FontFamily = Font, FontSize = AppTheme.FontXS,
            Padding = new Thickness(6, 1),
            Background = new SolidColorBrush(Color.Parse("#1a3a1a")),
            Foreground = new SolidColorBrush(Color.Parse("#88dd88")),
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(3),
            Cursor = new Cursor(StandardCursorType.Hand),
            VerticalAlignment = VerticalAlignment.Center,
        };
        btnBatch.Click += async (_, _) => await BatchAddForDate(capDate);

        var lbl = new TextBlock
        {
            Text = $"📅  {date}  ({count}건)",
            FontFamily = Font, FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        };
        var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        Grid.SetColumn(lbl,      0); grid.Children.Add(lbl);
        Grid.SetColumn(btnBatch, 1); grid.Children.Add(btnBatch);

        return new Border
        {
            Background = new SolidColorBrush(Color.Parse("#0e1a2a")),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(6, 4),
            Margin = new Thickness(0, 6, 0, 2),
            Child = grid,
        };
    }

    private async Task BatchAddForDate(string date)
    {
        _count.Text = $"📥  {date} 의뢰 일괄추가 중...";
        var added = await Task.Run(() => ReceiptDispatchService.GenerateForDate(date));
        // 발송일 자동계산도 함께 — 분석종료일 MAX → 발송일
        var dispUpdated = await Task.Run(() => ReceiptDispatchService.RecomputeDispatchDates(date));
        _cache = await Task.Run(() => ReceiptDispatchService.GetAll());
        Render();
        RenderReqList();
        _count.Text = $"✅ {added}건 추가, {dispUpdated}건 발송일 자동계산 — 총 {_cache.Count}건";
    }

    private Control BuildReqCard(AnalysisRequestRecord r)
    {
        // 이미 대장에 추가된 의뢰? (접수번호+시료명 키로 판정)
        bool added = IsAlreadyAdded(r);

        var sp = new StackPanel { Spacing = 1 };
        var top = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };

        // 체크 마크 (이미 추가된 경우) — 잘림 방지: 명시적 너비/오른쪽 여백 부여
        if (added)
        {
            top.Children.Add(new TextBlock
            {
                Text = "✅",
                FontSize = AppTheme.FontSM, FontFamily = Font,
                Width = 18,
                Margin = new Thickness(0, 0, 2, 0),
                TextAlignment = TextAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
            });
        }

        if (!string.IsNullOrEmpty(r.약칭))
        {
            var (bg, fg) = ETA.Services.Common.BadgeColorHelper.GetBadgeColor(r.약칭);
            top.Children.Add(new Border
            {
                Background = Brush.Parse(bg), CornerRadius = new CornerRadius(3),
                Padding = new Thickness(4, 1),
                Child = new TextBlock { Text = r.약칭, FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = Brush.Parse(fg) },
            });
        }
        top.Children.Add(new TextBlock
        {
            Text = r.시료명, FontFamily = Font, FontSize = AppTheme.FontSM,
            Foreground = added ? AppTheme.FgMuted : AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
        });
        sp.Children.Add(top);
        sp.Children.Add(new TextBlock
        {
            Text = $"{r.채취일자}  {r.접수번호}".Trim(),
            FontSize = AppTheme.FontXS, FontFamily = Font,
            Foreground = AppTheme.FgMuted,
        });

        var card = new Border
        {
            Background = added
                ? new SolidColorBrush(Color.Parse("#0e2014"))
                : AppTheme.BgPrimary,
            BorderBrush = added
                ? new SolidColorBrush(Color.Parse("#3a7a3a"))
                : AppTheme.BorderSubtle,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(6, 4),
            Margin = new Thickness(0, 1),
            Cursor = new Cursor(StandardCursorType.Hand),
            Opacity = added ? 0.7 : 1.0,
            Child = sp,
        };
        if (!added)
            card.PointerPressed += async (_, _) => await AddFromRequest(r);
        return card;
    }

    /// <summary>의뢰가 이미 대장에 추가됐는지 확인 — 접수번호+시료명+접수일 조합 키</summary>
    private bool IsAlreadyAdded(AnalysisRequestRecord r)
    {
        string key = $"{r.접수번호 ?? ""}|{r.시료명 ?? ""}|{r.채취일자 ?? ""}";
        return _cache.Any(e =>
            $"{e.접수번호}|{e.시료명}|{e.접수일}" == key ||
            // 접수번호 자동생성된 경우는 시료명+접수일로만 판정
            (e.시료명 == r.시료명 && e.접수일 == r.채취일자 && !string.IsNullOrEmpty(r.시료명)));
    }

    private async Task AddFromRequest(AnalysisRequestRecord r)
    {
        // 중복 방지
        if (IsAlreadyAdded(r))
        {
            _count.Text = $"⚠️  이미 추가됨: {r.시료명}";
            return;
        }
        // 의뢰 1건을 대장 행으로 추가 — 분석항목/의뢰인및업체명 자동 계산
        var (summary, requester) = await Task.Run(() => (
            ReceiptDispatchService.ComputeAnalyteSummary(r.Id),
            ReceiptDispatchService.ComputeRequesterLabel(r.Id)
        ));
        var ent = new ReceiptDispatchEntry
        {
            접수번호 = r.접수번호 ?? "",
            접수일   = r.채취일자 ?? "",
            시료명   = r.시료명   ?? "",
            업체명   = string.IsNullOrEmpty(requester) ? (r.약칭 ?? "") : requester,
            분석항목 = summary,
            발송일   = "",
        };
        var newId = await Task.Run(() => ReceiptDispatchService.Insert(ent));
        ent.Id = newId;
        _cache.Insert(0, ent);
        Render();
        RenderReqList();   // 체크마크 갱신
    }

    private Control BuildUI()
    {
        var title = new TextBlock
        {
            Text = "📋  측정대행 접수/발송 대장",
            FontSize = AppTheme.FontXL, FontWeight = FontWeight.Bold,
            FontFamily = Font, Foreground = AppTheme.FgPrimary,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var btnAdd     = MkBtn("➕  행 추가",        "#1a3a1a", "#7cd87c", async (_, _) => await AddRow());
        var btnClear   = MkBtn("🗑  전체 비우기",    "#3a1a1a", "#ee8888", async (_, _) => await ClearAll());
        var btnImport  = MkBtn("📥  엑셀 가져오기",  "#3a3a1a", "#dddd66", async (_, _) => await ImportExcel());
        var btnXlsx    = MkBtn("💾  Excel 내보내기", "#1a3a5a", "#88ccff", async (_, _) => await ExportExcel());
        var btnDocx    = MkBtn("📝  Word 내보내기",  "#3a1a3a", "#dd99ff", async (_, _) => await ExportWord());

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center,
        };
        btnRow.Children.Add(btnAdd);
        btnRow.Children.Add(btnClear);
        btnRow.Children.Add(btnImport);
        btnRow.Children.Add(btnXlsx);
        btnRow.Children.Add(btnDocx);

        var titleRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        Grid.SetColumn(title, 0);  titleRow.Children.Add(title);
        Grid.SetColumn(btnRow, 1); titleRow.Children.Add(btnRow);

        var hdr = BuildHeader();
        var scroll = new ScrollViewer
        {
            Content = _list,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
        };

        var root = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,Auto,*"),
            Margin = new Thickness(12, 12, 12, 8),
        };
        Grid.SetRow(titleRow, 0); root.Children.Add(titleRow);
        Grid.SetRow(_search,  1); root.Children.Add(_search);
        Grid.SetRow(_count,   2); root.Children.Add(_count);
        Grid.SetRow(hdr,      3); root.Children.Add(hdr);
        Grid.SetRow(scroll,   4); root.Children.Add(scroll);
        return root;
    }

    private static Button MkBtn(string text, string bg, string fg, EventHandler<Avalonia.Interactivity.RoutedEventArgs> onClick)
    {
        var b = new Button
        {
            Content = text,
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Padding = new Thickness(10, 4),
            Background = new SolidColorBrush(Color.Parse(bg)),
            Foreground = new SolidColorBrush(Color.Parse(fg)),
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(4),
            Cursor = new Cursor(StandardCursorType.Hand),
        };
        b.Click += onClick;
        return b;
    }

    private static Grid BuildHeader()
    {
        var g = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(ColDefs),
            Margin = new Thickness(0, 6, 0, 4),
        };
        AddHdr(g, "접수번호",         0);
        AddHdr(g, "접수일",           1);
        AddHdr(g, "시료명",           2);
        AddHdr(g, "의뢰인 및 업체명", 3);
        AddHdr(g, "분석항목",         4);
        AddHdr(g, "발송일",           5);
        AddHdr(g, "",                 6);
        return g;
    }

    private static void AddHdr(Grid g, string text, int col)
    {
        var b = new Border
        {
            Background = AppTheme.BgPrimary,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Padding = new Thickness(6, 4),
            Child = new TextBlock
            {
                Text = text, FontFamily = Font, FontSize = AppTheme.FontSM,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppTheme.FgInfo,
                TextAlignment = TextAlignment.Center,
            },
        };
        Grid.SetColumn(b, col);
        g.Children.Add(b);
    }

    private async Task LoadAsync()
    {
        _count.Text = "📥  로딩 중...";
        _list.Children.Clear();
        try
        {
            // 이전 import 잔존 푸터 행 자동 정리
            await Task.Run(() => ReceiptDispatchService.PurgeFooterRows());
            _cache = await Task.Run(() => ReceiptDispatchService.GetAll());
            Render();
        }
        catch (Exception ex)
        {
            _list.Children.Clear();
            _count.Text = "❌ 로드 실패";
            _list.Children.Add(new TextBlock
            {
                Text = $"로드 오류: {ex.Message}",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgDanger,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(8, 12),
            });
        }
    }

    private void Render()
    {
        _list.Children.Clear();
        var q = _search.Text?.Trim().ToLowerInvariant() ?? "";
        // 접수일 ASC 정렬 (엑셀 양식 기준)
        var sorted = _cache.OrderBy(e => e.접수일 ?? "").ThenBy(e => e.접수번호 ?? "").ToList();
        var filtered = string.IsNullOrEmpty(q) ? sorted : sorted.Where(e =>
            e.접수번호.ToLower().Contains(q) ||
            e.시료명.ToLower().Contains(q) ||
            e.업체명.ToLower().Contains(q) ||
            e.분석항목.ToLower().Contains(q)).ToList();

        var visible = filtered.Take(MaxRender).ToList();
        _count.Text = filtered.Count > MaxRender
            ? $"총 {_cache.Count}건  /  표시 {visible.Count} / {filtered.Count}건  (검색어로 더 좁혀주세요)"
            : $"총 {_cache.Count}건  /  표시 {filtered.Count}건";

        int idx = 0;
        foreach (var e in visible)
        {
            _list.Children.Add(BuildRow(e, idx % 2 == 0));
            idx++;
        }
        if (filtered.Count == 0)
            _list.Children.Add(new TextBlock
            {
                Text = "표시할 항목이 없습니다. — [➕ 행 추가] 또는 [📥 엑셀 가져오기]",
                FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgMuted,
                Margin = new Thickness(8, 12),
            });
    }

    private Control BuildRow(ReceiptDispatchEntry e, bool altBg)
    {
        var bg = altBg ? "#14141e" : "#18182a";
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(ColDefs) };

        var tbNo   = MakeEditCell(e.접수번호,           bg, 0);
        var tbDate = MakeEditCell(e.접수일,             bg, 1);
        var tbName = MakeEditCell(e.시료명,             bg, 2);
        var tbComp = MakeEditCell(e.업체명,             bg, 3);
        var tbItem = MakeEditCell(e.분석항목,           bg, 4);
        var tbSend = MakeEditCell(e.발송일,             bg, 5);

        // 자동 저장 (LostFocus / Enter)
        var capId = e.Id;
        void SaveOnEdit(object? _, Avalonia.Interactivity.RoutedEventArgs __)
        {
            var upd = new ReceiptDispatchEntry
            {
                Id = capId,
                접수번호 = tbNo.Text   ?? "",
                접수일   = tbDate.Text ?? "",
                시료명   = tbName.Text ?? "",
                업체명   = tbComp.Text ?? "",
                분석항목 = tbItem.Text ?? "",
                발송일   = tbSend.Text ?? "",
            };
            Task.Run(() => ReceiptDispatchService.Update(upd));
            // 캐시 갱신
            var idxC = _cache.FindIndex(x => x.Id == capId);
            if (idxC >= 0) _cache[idxC] = upd;
        }
        foreach (var tb in new[] { tbNo, tbDate, tbName, tbComp, tbItem, tbSend })
        {
            tb.LostFocus += SaveOnEdit;
            tb.KeyDown += (_, ke) =>
            {
                if (ke.Key == Key.Enter) SaveOnEdit(null, null!);
            };
        }
        g.Children.Add(tbNo.Parent as Control ?? tbNo);
        g.Children.Add(tbDate.Parent as Control ?? tbDate);
        g.Children.Add(tbName.Parent as Control ?? tbName);
        g.Children.Add(tbComp.Parent as Control ?? tbComp);
        g.Children.Add(tbItem.Parent as Control ?? tbItem);
        g.Children.Add(tbSend.Parent as Control ?? tbSend);

        // 삭제 버튼
        var btnDel = new Button
        {
            Content = "🗑", FontSize = AppTheme.FontSM,
            Padding = new Thickness(2),
            Background = Brushes.Transparent,
            Foreground = new SolidColorBrush(Color.Parse("#cc6666")),
            BorderThickness = new Thickness(0),
            Cursor = new Cursor(StandardCursorType.Hand),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
        };
        btnDel.Click += async (_, _) => await DeleteRow(capId);
        var delCell = new Border
        {
            Background = Brush.Parse(bg),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Child = btnDel,
        };
        Grid.SetColumn(delCell, 6);
        g.Children.Add(delCell);
        return g;
    }

    private static TextBox MakeEditCell(string text, string bg, int col)
    {
        var tb = new TextBox
        {
            Text = text,
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Padding = new Thickness(4, 2),
            Background = Brushes.Transparent,
            Foreground = AppTheme.FgPrimary,
            BorderThickness = new Thickness(0),
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalContentAlignment = VerticalAlignment.Center,
        };
        var cell = new Border
        {
            Background = Brush.Parse(bg),
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Padding = new Thickness(2),
            Child = tb,
        };
        Grid.SetColumn(cell, col);
        return tb;
    }

    private async Task AddRow()
    {
        var ent = new ReceiptDispatchEntry { 접수일 = DateTime.Today.ToString("yyyy-MM-dd") };
        var newId = await Task.Run(() => ReceiptDispatchService.Insert(ent));
        ent.Id = newId;
        _cache.Insert(0, ent);
        Render();
    }

    private async Task DeleteRow(int id)
    {
        await Task.Run(() => ReceiptDispatchService.Delete(id));
        _cache.RemoveAll(e => e.Id == id);
        Render();
        RenderReqList();
    }

    private async Task ClearAll()
    {
        if (_cache.Count == 0) return;
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel is Window owner)
        {
            bool ok = await ShowConfirm(owner, $"대장 {_cache.Count}건을 모두 삭제할까요?\n(엑셀 파일에는 영향 없음 — DB만 비웁니다)");
            if (!ok) return;
        }
        await Task.Run(() => ReceiptDispatchService.DeleteAll());
        _cache.Clear();
        Render();
        RenderReqList();
        _count.Text = "🗑 전체 비움";
    }

    private static async Task<bool> ShowConfirm(Window owner, string msg)
    {
        bool result = false;
        var dlg = new Window
        {
            Title = "확인", Width = 360, Height = 180,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = AppTheme.BgPrimary,
        };
        var yes = new Button
        {
            Content = "삭제", Width = 80, Height = 28,
            Background = AppTheme.BgDanger, Foreground = AppTheme.FgDanger,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontBase,
        };
        var no = new Button
        {
            Content = "취소", Width = 70, Height = 28,
            Background = AppTheme.BorderSubtle, Foreground = AppTheme.FgMuted,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            FontFamily = Font, FontSize = AppTheme.FontBase,
        };
        yes.Click += (_, _) => { result = true;  dlg.Close(); };
        no.Click  += (_, _) => { result = false; dlg.Close(); };
        dlg.Content = new StackPanel
        {
            Margin = new Thickness(20), Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text = msg, FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = AppTheme.FgPrimary,
                    TextWrapping = TextWrapping.Wrap,
                },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 8,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Children = { yes, no },
                },
            },
        };
        await dlg.ShowDialog(owner);
        return result;
    }

    private async Task ImportExcel()
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;
        var pick = await topLevel.StorageProvider.OpenFilePickerAsync(
            new Avalonia.Platform.Storage.FilePickerOpenOptions
            {
                Title = "접수발송대장 엑셀 선택",
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel")
                    {
                        Patterns = new[] { "*.xlsm", "*.xlsx" },
                    },
                },
            });
        if (pick.Count == 0) return;
        var path = pick[0].Path.LocalPath;

        _count.Text = "📥  엑셀 가져오는 중...";
        var n = await Task.Run(() => ReceiptDispatchService.ImportFromExcel(path));
        _cache = await Task.Run(() => ReceiptDispatchService.GetAll());
        Render();
        _count.Text = $"✅ {n}건 가져옴 — 총 {_cache.Count}건";
    }

    private async Task ExportExcel()
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;
        var defaultName = $"접수발송대장_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
        var save = await topLevel.StorageProvider.SaveFilePickerAsync(
            new Avalonia.Platform.Storage.FilePickerSaveOptions
            {
                SuggestedFileName = defaultName,
                DefaultExtension = "xlsx",
                FileTypeChoices = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel")
                    {
                        Patterns = new[] { "*.xlsx" },
                    },
                },
            });
        if (save == null) return;
        var path = save.Path.LocalPath;
        // 접수일 ASC 정렬 후 export
        var sorted = _cache.OrderBy(e => e.접수일 ?? "").ThenBy(e => e.접수번호 ?? "").ToList();
        WriteWorkbook(path, sorted);
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = path, UseShellExecute = true,
            });
        }
        catch { }
    }

    private async Task ExportWord()
    {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;
        var defaultName = $"접수발송대장_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
        var save = await topLevel.StorageProvider.SaveFilePickerAsync(
            new Avalonia.Platform.Storage.FilePickerSaveOptions
            {
                SuggestedFileName = defaultName,
                DefaultExtension = "docx",
                FileTypeChoices = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Word")
                    {
                        Patterns = new[] { "*.docx" },
                    },
                },
            });
        if (save == null) return;
        var path = save.Path.LocalPath;
        var sorted = _cache.OrderBy(e => e.접수일 ?? "").ThenBy(e => e.접수번호 ?? "").ToList();
        try
        {
            ETA.Services.Common.ReceiptDispatchWordExporter.Export(path, sorted);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = path, UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            _count.Text = $"❌ Word 출력 실패: {ex.Message}";
        }
    }

    /// <summary>로고 PNG 의 절대경로 (없으면 null)</summary>
    private static string? GetLogoPath()
    {
        try
        {
            var p = System.IO.Path.Combine(
                ETA.Services.Common.AppPaths.RootPath,
                "Assets", "icons", "renewus_vertical_black.png");
            return File.Exists(p) ? p : null;
        }
        catch { return null; }
    }

    /// <summary>엑셀 `접수발송대장` 시트 양식 — ClosedXML 사용. 다중 페이지 자동 분할.
    /// - 1행: 제목, 2행: 헤더 (모든 페이지에 반복 노출)
    /// - 3행~: 데이터 (연속)
    /// - 마지막 페이지 하단: 서명 푸터 3행 + 회사명 1행 (항상 마지막 페이지 끝에 위치)
    /// - 행 수에 맞춰 빈 행 패딩으로 서명을 마지막 페이지 바닥으로 밀어냄.</summary>
    private static void WriteWorkbook(string path, List<ReceiptDispatchEntry> rows)
    {
        if (File.Exists(path)) File.Delete(path);
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add("접수발송대장");

        var year = DateTime.Now.Year;

        // ── 페이지 메트릭 (A4 가로, 좌우상하 마진 0.5") ─────────────────────
        // A4 가로 = 11.69" × 8.27" (842 × 595 pt). 마진 합 1" 빼면 데이터 영역 = ~770pt × ~525pt
        const double pageHeightPt = 770;
        const double titleRowPt   = 32;
        const double headerRowPt  = 21;
        const double dataRowPt    = 18;
        const double sigRowPt     = 28;
        const double footerRowPt  = 40;
        const double sigBlockPt   = sigRowPt * 3 + footerRowPt;   // 124pt

        // 모든 페이지에서 1~2행 (제목+헤더) 반복 → 데이터 영역 = pageHeight - title - header
        double availPerPage  = pageHeightPt - titleRowPt - headerRowPt;             // 717
        double availLastPage = availPerPage - sigBlockPt;                            // 593
        int rowsPerPage      = Math.Max(1, (int)(availPerPage / dataRowPt));         // ~39
        int rowsLastPage     = Math.Max(1, (int)(availLastPage / dataRowPt));        // ~32

        const int MaxData = 1000;
        var dataRows = rows.Take(MaxData).ToList();
        int dataN = dataRows.Count;

        int totalPages, dataOnLast, paddingBeforeSig;
        if (dataN <= rowsLastPage)
        {
            totalPages       = 1;
            dataOnLast       = dataN;
            paddingBeforeSig = rowsLastPage - dataN;
        }
        else
        {
            int fullPages = (int)Math.Ceiling((dataN - rowsLastPage) / (double)rowsPerPage);
            totalPages       = fullPages + 1;
            dataOnLast       = dataN - fullPages * rowsPerPage;
            paddingBeforeSig = rowsLastPage - dataOnLast;
            if (paddingBeforeSig < 0) paddingBeforeSig = 0;
        }

        // ── 1) 제목 (A1:F1 병합) — 진한 블루 배경 + 흰색 굵은 글씨 ─────────
        ws.Cell(1, 1).Value = $"{year}년 측정대행기록부 접수/발송 대장";
        ws.Range(1, 1, 1, 6).Merge();
        ws.Cell(1, 1).Style
            .Font.SetBold().Font.SetFontSize(16).Font.SetFontColor(ClosedXML.Excel.XLColor.FromHtml("#1F4E78"))
            .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);
        ws.Row(1).Height = titleRowPt;

        // 회사 로고 워터마크 (제목 행 좌측에 작게)
        var logoPath = GetLogoPath();
        if (logoPath != null)
        {
            try
            {
                var img = ws.AddPicture(logoPath)
                    .MoveTo(ws.Cell(1, 1), 6, 4)
                    .WithSize(36, 25);
                img.Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.Move;
            }
            catch { }
        }

        // ── 2) 헤더 (반복) — 옅은 블루 배경 + 진한 글씨 ────────────────────
        var hdrs = new[] { "접수번호", "접수일", "시료명", "의뢰인\n및 업체명", "분석항목", "발송일" };
        for (int c = 0; c < hdrs.Length; c++)
        {
            var cell = ws.Cell(2, c + 1);
            cell.Value = hdrs[c];
            cell.Style
                .Font.SetBold().Font.SetFontColor(ClosedXML.Excel.XLColor.FromHtml("#1F4E78"))
                .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center)
                .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center)
                .Alignment.SetWrapText(true)
                .Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.FromHtml("#D9E2F3"))
                .Border.SetOutsideBorder(ClosedXML.Excel.XLBorderStyleValues.Medium)
                .Border.SetOutsideBorderColor(ClosedXML.Excel.XLColor.FromHtml("#1F4E78"));
        }
        ws.Row(2).Height = headerRowPt;

        // ── 3) 데이터 행 ────────────────────────────────────────────────────
        const int dataStart = 3;
        for (int i = 0; i < dataN; i++)
        {
            int r = dataStart + i;
            var e = dataRows[i];
            ws.Cell(r, 1).Value = e.접수번호 ?? "";
            ws.Cell(r, 2).Value = e.접수일 ?? "";
            ws.Cell(r, 3).Value = e.시료명 ?? "";
            ws.Cell(r, 4).Value = e.업체명 ?? "";
            ws.Cell(r, 5).Value = e.분석항목 ?? "";
            ws.Cell(r, 6).Value = e.발송일 ?? "";
            ws.Row(r).Height = dataRowPt;

            // 모든 데이터 셀 통일된 글자 크기 — 7pt
            for (int c = 1; c <= 6; c++)
                ws.Cell(r, c).Style.Font.SetFontSize(7);
        }

        // 데이터 영역 + 패딩 영역 모두 테두리 (빈 행도 그리드 유지)
        int dataAreaEnd = dataStart + dataN + paddingBeforeSig - 1;
        if (dataAreaEnd >= dataStart)
        {
            var dataRange = ws.Range(dataStart, 1, dataAreaEnd, 6);
            dataRange.Style
                .Border.SetOutsideBorder(ClosedXML.Excel.XLBorderStyleValues.Thin)
                .Border.SetInsideBorder(ClosedXML.Excel.XLBorderStyleValues.Hair)
                .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);

            // 컬럼별 정렬 (시료명/업체명/분석항목은 좌측, 나머지는 가운데)
            ws.Range(dataStart, 1, dataAreaEnd, 1).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center);
            ws.Range(dataStart, 2, dataAreaEnd, 2).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center);
            ws.Range(dataStart, 3, dataAreaEnd, 3).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Left);
            ws.Range(dataStart, 4, dataAreaEnd, 4).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Left);
            ws.Range(dataStart, 5, dataAreaEnd, 5).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Left);
            ws.Range(dataStart, 6, dataAreaEnd, 6).Style.Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center);

            // 텍스트 줄바꿈 (긴 분석항목 처리)
            ws.Range(dataStart, 5, dataAreaEnd, 5).Style.Alignment.SetWrapText(true);
        }

        // 패딩 빈 행도 데이터 행 높이 적용
        for (int r = dataStart + dataN; r <= dataAreaEnd; r++)
            ws.Row(r).Height = dataRowPt;

        // ── 4) 서명 푸터 (마지막 페이지 하단) — 옅은 컬러 강조 ──────────────
        int sigStart = dataAreaEnd + 1;
        // A:B 병합 = 시료접수 (옅은 그린)
        ws.Range(sigStart,     1, sigStart + 2, 2).Merge();
        ws.Cell(sigStart, 1).Value = "시료접수";
        ws.Cell(sigStart, 1).Style
            .Font.SetBold().Font.SetFontSize(9).Font.SetFontColor(ClosedXML.Excel.XLColor.FromHtml("#1F5E1F"))
            .Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.FromHtml("#E8F4E8"))
            .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);

        // D 병합 = 시험성적서 발송 (옅은 블루)
        ws.Range(sigStart, 4, sigStart + 2, 4).Merge();
        ws.Cell(sigStart, 4).Value = "시험성적서 발송";
        ws.Cell(sigStart, 4).Style
            .Font.SetBold().Font.SetFontSize(9).Font.SetFontColor(ClosedXML.Excel.XLColor.FromHtml("#1F4E78"))
            .Fill.SetBackgroundColor(ClosedXML.Excel.XLColor.FromHtml("#E2EAF6"))
            .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);

        // C, E:F 병합 = 담당자/기술책임자/품질책임자 라인
        var sigLines = new[]
        {
            "담  당  자 :                                  (서명)",
            "기술책임자 :                                  (서명)",
            "품질책임자 :                                  (서명)",
        };
        for (int i = 0; i < 3; i++)
        {
            int r = sigStart + i;
            ws.Cell(r, 3).Value = sigLines[i];
            ws.Range(r, 5, r, 6).Merge();
            ws.Cell(r, 5).Value = sigLines[i];
            ws.Row(r).Height = sigRowPt;
            ws.Cell(r, 3).Style
                .Font.SetFontSize(8)
                .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Left)
                .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);
            ws.Cell(r, 5).Style
                .Font.SetFontSize(8)
                .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Left)
                .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);
        }
        // 서명 블록 외곽 테두리
        ws.Range(sigStart, 1, sigStart + 2, 6).Style
            .Border.SetOutsideBorder(ClosedXML.Excel.XLBorderStyleValues.Thin)
            .Border.SetInsideBorder(ClosedXML.Excel.XLBorderStyleValues.Hair);

        // ── 5) 회사명 (A:F 병합, 마지막 행) — 진한 컬러 + 로고 ──────────────
        int 회사명행 = sigStart + 3;
        ws.Range(회사명행, 1, 회사명행, 6).Merge();
        ws.Cell(회사명행, 1).Value = "리뉴어스 수질분석센터";
        ws.Cell(회사명행, 1).Style
            .Font.SetBold().Font.SetFontSize(18).Font.SetFontColor(ClosedXML.Excel.XLColor.FromHtml("#1F4E78"))
            .Alignment.SetHorizontal(ClosedXML.Excel.XLAlignmentHorizontalValues.Center)
            .Alignment.SetVertical(ClosedXML.Excel.XLAlignmentVerticalValues.Center);
        ws.Row(회사명행).Height = footerRowPt;

        // ── 6) 컬럼 너비 (A4 세로 페이지 너비 기준) ────────────────────────
        // 합계 = 79 (세로 양식 — 너비 좁아져 분석항목/시료명 줄임)
        ws.Column(1).Width = 13;     // 접수번호
        ws.Column(2).Width = 9;      // 접수일
        ws.Column(3).Width = 17;     // 시료명
        ws.Column(4).Width = 16;     // 의뢰인및업체명
        ws.Column(5).Width = 17;     // 분석항목 (줄바꿈)
        ws.Column(6).Width = 9;      // 발송일

        // ── 7) 페이지 설정 — A4 세로, 너비 1페이지 고정·세로 자동 ──────────
        ws.PageSetup.PageOrientation = ClosedXML.Excel.XLPageOrientation.Portrait;
        ws.PageSetup.PaperSize       = ClosedXML.Excel.XLPaperSize.A4Paper;
        ws.PageSetup.FitToPages(1, totalPages);
        ws.PageSetup.Margins.Left   = 0.4;
        ws.PageSetup.Margins.Right  = 0.4;
        ws.PageSetup.Margins.Top    = 0.5;
        ws.PageSetup.Margins.Bottom = 0.5;
        ws.PageSetup.CenterHorizontally = true;
        // 1~2행 (제목+헤더) 모든 페이지 상단 반복
        ws.PageSetup.SetRowsToRepeatAtTop(1, 2);
        // 인쇄 영역
        ws.PageSetup.PrintAreas.Add($"A1:F{회사명행}");

        wb.SaveAs(path);
    }

    private static SS.Row MakeRow(int rowIdx, params string[] vals)
    {
        var row = new SS.Row { RowIndex = (uint)rowIdx };
        for (int i = 0; i < vals.Length; i++)
        {
            row.Append(new SS.Cell
            {
                CellReference = $"{ColLetter(i + 1)}{rowIdx}",
                DataType = SS.CellValues.InlineString,
                InlineString = new SS.InlineString(new SS.Text(vals[i] ?? "")),
            });
        }
        return row;
    }

    private static string ColLetter(int n)
    {
        string s = "";
        while (n > 0) { n--; s = (char)('A' + n % 26) + s; n /= 26; }
        return s;
    }
}
