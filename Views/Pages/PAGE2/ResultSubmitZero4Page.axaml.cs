using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.SERVICE2;

namespace ETA.Views.Pages.PAGE2;

public partial class ResultSubmitZero4Page : UserControl
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static readonly string[] TrackedItems =
        { "BOD", "TOC", "SS", "T-N", "T-P", "총대장균군" };

    private static readonly string[] Groups =
        { "중흥+월내", "4단계", "율촌+해룡", "세풍" };

    // 선택된 그룹 필터 (null = 전체)
    private string? _selectedGroup = null;

    // 주 단위 버킷: (월요일, 일요일) → 해당 주의 진행도 아이템 리스트 (로드 후 캐시)
    private class WeekBucket
    {
        public DateTime WeekStart;   // 월요일
        public DateTime WeekEnd;     // 일요일
        public List<string> Dates = new();                   // 이 주에 속한 채취일자들
        public List<FacilityResultService.WaybleDateItem>? Loaded; // 지연 로드
    }

    private List<WeekBucket> _weeks = new();

    public event Action<string /*date*/, string /*group*/>? DateGroupSelected;

    public ResultSubmitZero4Page()
    {
        InitializeComponent();
        BuildGroupBadges();
        Reload();
    }

    public void Reload()
    {
        _weeks = LoadWeekBuckets();
        RenderList();
    }

    // ─────────────────────────────────────────────────────────────────
    // 주 단위 버킷 구성 (금주는 즉시 로드, 나머지는 지연)
    // ─────────────────────────────────────────────────────────────────
    private static List<WeekBucket> LoadWeekBuckets()
    {
        var buckets = new List<WeekBucket>();
        List<string> dates;
        try { dates = FacilityResultService.GetWaybleDistinctDates(); }
        catch { return buckets; }

        var map = new Dictionary<DateTime, WeekBucket>();
        foreach (var ds in dates)
        {
            if (!DateTime.TryParse(ds, out var d)) continue;
            var mon = MondayOf(d);
            if (!map.TryGetValue(mon, out var bucket))
            {
                bucket = new WeekBucket { WeekStart = mon, WeekEnd = mon.AddDays(6) };
                map[mon] = bucket;
            }
            bucket.Dates.Add(ds);
        }

        foreach (var kv in map.OrderByDescending(x => x.Key))
            buckets.Add(kv.Value);
        return buckets;
    }

    private static DateTime MondayOf(DateTime d)
    {
        int diff = ((int)d.DayOfWeek + 6) % 7; // Mon=0, Sun=6
        return d.Date.AddDays(-diff);
    }

    private static bool IsCurrentWeek(WeekBucket b)
        => b.WeekStart == MondayOf(DateTime.Today);

    // ─────────────────────────────────────────────────────────────────
    // 상단 시설그룹 뱃지 (뱃지 초성컬러헬퍼 적용)
    // ─────────────────────────────────────────────────────────────────
    private void BuildGroupBadges()
    {
        GroupBadgePanel.Children.Clear();
        GroupBadgePanel.Children.Add(BuildGroupBadge("전체", null));
        foreach (var g in Groups)
            GroupBadgePanel.Children.Add(BuildGroupBadge(g, g));
    }

    private Control BuildGroupBadge(string label, string? groupKey)
    {
        var (bg, fg, bd) = WasteCompanyPage.GetChosungBadgeColorPublic(label);
        bool selected = _selectedGroup == groupKey;

        var border = new Border
        {
            Background = Brush.Parse(bg),
            BorderBrush = Brush.Parse(bd),
            BorderThickness = new Thickness(selected ? 2 : 1),
            CornerRadius = new CornerRadius(999),
            Padding = new Thickness(12, 4),
            Margin = new Thickness(0, 0, 6, 0),
            Cursor = new Cursor(StandardCursorType.Hand),
            Opacity = selected ? 1.0 : 0.75,
            Child = new TextBlock
            {
                Text = label,
                FontFamily = Font,
                FontSize = 12,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse(fg),
                VerticalAlignment = VerticalAlignment.Center,
            },
        };

        border.PointerPressed += (_, _) =>
        {
            _selectedGroup = groupKey;
            BuildGroupBadges();
            RenderList();
        };

        return border;
    }

    // ─────────────────────────────────────────────────────────────────
    // 리스트: 금주 섹션(즉시) + 이전 주 Expander(지연 로드)
    // ─────────────────────────────────────────────────────────────────
    private void RenderList()
    {
        DateListPanel.Children.Clear();

        if (_weeks.Count == 0)
        {
            DateListPanel.Children.Add(new TextBlock
            {
                Text = "표시할 일자가 없습니다.",
                FontFamily = Font,
                FontSize = 12,
                Foreground = Brush.Parse("#888899"),
                Margin = new Thickness(8, 12),
            });
            return;
        }

        foreach (var bucket in _weeks)
        {
            if (IsCurrentWeek(bucket))
            {
                // 금주: 즉시 로드 및 인라인 렌더
                EnsureLoaded(bucket);
                DateListPanel.Children.Add(BuildWeekHeader(bucket, isCurrent: true));
                RenderBucketItems(bucket, DateListPanel);
            }
            else
            {
                // 이전 주: Expander (처음 펼칠 때 로드)
                DateListPanel.Children.Add(BuildWeekExpander(bucket));
            }
        }
    }

    private Control BuildWeekHeader(WeekBucket bucket, bool isCurrent)
    {
        var range = $"{bucket.WeekStart:yyyy-MM-dd} ~ {bucket.WeekEnd:MM-dd}";
        var label = isCurrent ? $"📅 금주  {range}" : range;

        return new Border
        {
            Background = Brush.Parse("#1e3a5a"),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(10, 5),
            Margin = new Thickness(0, 2, 0, 4),
            Child = new TextBlock
            {
                Text = label,
                FontFamily = Font,
                FontSize = 12,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse("#cce0ff"),
            }
        };
    }

    private Control BuildWeekExpander(WeekBucket bucket)
    {
        var range = $"{bucket.WeekStart:yyyy-MM-dd} ~ {bucket.WeekEnd:MM-dd}";
        var headerTb = new TextBlock
        {
            Text = $"{range}   ({bucket.Dates.Count}일)",
            FontFamily = Font,
            FontSize = 12,
            FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#aaaaaa"),
        };

        var inner = new StackPanel { Spacing = 4, Margin = new Thickness(0, 4, 0, 0) };

        var expander = new Expander
        {
            Header = headerTb,
            IsExpanded = false,
            Padding = new Thickness(4, 4),
            Margin = new Thickness(0, 2, 0, 2),
            Background = Brush.Parse("#252535"),
            BorderBrush = Brush.Parse("#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Content = inner,
        };

        // 처음 펼칠 때 로드
        expander.PropertyChanged += (s, e) =>
        {
            if (e.Property != Expander.IsExpandedProperty) return;
            if (!(bool)(e.NewValue ?? false)) return;
            if (bucket.Loaded != null) return; // 이미 로드됨

            EnsureLoaded(bucket);
            inner.Children.Clear();
            RenderBucketItems(bucket, inner);
        };

        return expander;
    }

    private void EnsureLoaded(WeekBucket bucket)
    {
        if (bucket.Loaded != null) return;
        try
        {
            bucket.Loaded = FacilityResultService.GetWaybleDateItemsInRange(
                bucket.WeekStart.ToString("yyyy-MM-dd"),
                bucket.WeekEnd.ToString("yyyy-MM-dd"));
        }
        catch
        {
            bucket.Loaded = new List<FacilityResultService.WaybleDateItem>();
        }
    }

    private void RenderBucketItems(WeekBucket bucket, Panel target)
    {
        if (bucket.Loaded == null) return;

        var items = bucket.Loaded
            .Where(x => _selectedGroup == null || x.Group == _selectedGroup)
            .ToList();

        if (items.Count == 0)
        {
            target.Children.Add(new TextBlock
            {
                Text = "· 선택한 그룹의 자료가 없습니다.",
                FontFamily = Font,
                FontSize = 11,
                Foreground = Brush.Parse("#666677"),
                Margin = new Thickness(8, 2),
            });
            return;
        }

        // 날짜별로 그룹핑 → 날짜가 top-level 아이템, 처리시설 그룹은 하위 행으로
        var byDate = items
            .GroupBy(x => x.Date)
            .OrderByDescending(g => g.Key);
        foreach (var grp in byDate)
            target.Children.Add(BuildDateCard(grp.Key, grp.ToList()));
    }

    private Control BuildDateCard(string date, List<FacilityResultService.WaybleDateItem> groupItems)
    {
        // 헤더: 날짜
        var dateTb = new TextBlock
        {
            Text       = date,
            FontFamily = Font,
            FontSize   = 14,
            FontWeight = FontWeight.SemiBold,
            Foreground = Brush.Parse("#d8d8e0"),
            Margin     = new Thickness(0, 0, 0, 6),
        };

        // 본문: 처리시설(그룹)별 행
        var body = new StackPanel { Spacing = 6 };
        foreach (var item in groupItems.OrderBy(x => Array.IndexOf(Groups, x.Group)))
            body.Children.Add(BuildFacilityRow(date, item));

        var content = new StackPanel { Spacing = 0 };
        content.Children.Add(dateTb);
        content.Children.Add(body);

        return new Border
        {
            Background = Brush.Parse("#2d2d35"),
            BorderBrush = Brush.Parse("#404050"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(10, 8),
            Margin = new Thickness(0, 0, 0, 6),
            Child = content,
        };
    }

    private Control BuildFacilityRow(string date, FacilityResultService.WaybleDateItem item)
    {
        var grp = item.Group;
        var (bg, fg, bd) = WasteCompanyPage.GetChosungBadgeColorPublic(grp);

        // 처리시설 뱃지
        var badge = new Border
        {
            Background = Brush.Parse(bg),
            BorderBrush = Brush.Parse(bd),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(8, 2),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = grp,
                FontFamily = Font,
                FontSize = 11,
                FontWeight = FontWeight.SemiBold,
                Foreground = Brush.Parse(fg),
            },
        };

        var itemBars = new StackPanel { Spacing = 2, Margin = new Thickness(8, 4, 0, 0) };
        foreach (var it in TrackedItems)
        {
            var (filled, total) = item.ItemProgress.TryGetValue(it, out var p) ? p : (0, 0);
            itemBars.Children.Add(BuildItemProgressRow(it, filled, total));
        }

        var stack = new StackPanel { Spacing = 0 };
        stack.Children.Add(badge);
        stack.Children.Add(itemBars);

        var row = new Border
        {
            Background = Brush.Parse("#252530"),
            BorderBrush = Brush.Parse("#383848"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(8, 6),
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = stack,
        };

        row.PointerPressed += (_, _) =>
        {
            DateGroupSelected?.Invoke(date, item.Group);
        };

        return row;
    }

    private Control BuildItemProgressRow(string itemName, int filled, int total)
    {
        double ratio = total > 0 ? (double)filled / total : 0;
        bool done = total > 0 && filled >= total;

        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("64,*,Auto"),
            Margin = new Thickness(0, 1),
        };

        var nameTb = new TextBlock
        {
            Text = itemName,
            FontFamily = Font,
            FontSize = 11,
            Foreground = Brush.Parse("#aaaaaa"),
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(nameTb, 0);
        grid.Children.Add(nameTb);

        var trackBg = new Border
        {
            Height = 10,
            Background = Brush.Parse("#1f1f28"),
            BorderBrush = Brush.Parse("#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(5),
            VerticalAlignment = VerticalAlignment.Center,
            ClipToBounds = true,
        };
        var fill = new Border
        {
            Height = 10,
            Background = Brush.Parse(done ? "#2fa36b" : "#3b6bd8"),
            CornerRadius = new CornerRadius(5),
            HorizontalAlignment = HorizontalAlignment.Left,
        };
        trackBg.Child = fill;
        trackBg.SizeChanged += (_, e) =>
        {
            fill.Width = Math.Max(0, e.NewSize.Width * ratio);
        };
        Grid.SetColumn(trackBg, 1);
        grid.Children.Add(trackBg);

        var right = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 4,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 0, 0),
        };

        if (done)
        {
            right.Children.Add(BuildCheckBadge());
        }
        else if (total == 0)
        {
            right.Children.Add(new TextBlock
            {
                Text = "—",
                FontFamily = Font,
                FontSize = 10,
                Foreground = Brush.Parse("#666677"),
                VerticalAlignment = VerticalAlignment.Center,
            });
        }
        else
        {
            right.Children.Add(new TextBlock
            {
                Text = $"{filled}/{total}",
                FontFamily = Font,
                FontSize = 10,
                Foreground = Brush.Parse("#aaaaaa"),
                VerticalAlignment = VerticalAlignment.Center,
                MinWidth = 30,
                TextAlignment = TextAlignment.Right,
            });
        }
        Grid.SetColumn(right, 2);
        grid.Children.Add(right);

        return grid;
    }

    private static Control BuildCheckBadge()
    {
        var ellipse = new Ellipse
        {
            Width = 16,
            Height = 16,
            Fill = Brush.Parse("#1a3a2a"),
            Stroke = Brush.Parse("#2fa36b"),
            StrokeThickness = 1.5,
        };
        var check = new TextBlock
        {
            Text = "✓",
            FontSize = 11,
            FontWeight = FontWeight.Bold,
            Foreground = Brush.Parse("#88ddaa"),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };
        return new Grid { Children = { ellipse, check } };
    }
}
