using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Views.Pages.PAGE1;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views;

/// <summary>미매칭 시료 수동 매칭 팝업 — 의뢰시료 / 처리장 탭에서 후보 선택</summary>
public class ManualMatchWindow : Window
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static IBrush Res(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is IBrush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    // ── 선택 완료 이벤트 ──────────────────────────────────────────────────────
    public event Action<ManualMatchWindow>? MatchConfirmed;

    public WasteSample?                                  SelectedWaste     { get; private set; }
    public AnalysisRequestRecord?                        SelectedAnalysis  { get; private set; }
    public (string 시설명, string 시료명, int 마스터Id)? SelectedFacility  { get; private set; }

    // ── 데이터 ───────────────────────────────────────────────────────────────
    private readonly List<AnalysisRequestRecord>                        _analysisRecords;
    private readonly List<(string 시설명, string 시료명, int 마스터Id)> _facilityMasters;
    private readonly List<WasteSample>                                  _wasteSamples;
    private readonly string                                              _targetName;
    private readonly Dictionary<int, string>                            _masterNotes;

    // ── 리스트 패널 (검색 시 재구성) ─────────────────────────────────────────
    private StackPanel _analysisList  = new() { Spacing = 0 };
    private StackPanel _facilityList  = new() { Spacing = 0 }; // 검색 시 단일 컬럼
    private StackPanel _wasteList     = new() { Spacing = 0 };
    // 컬럼 레이아웃용 (검색 없을 때)
    private Grid?      _facilityColGrid;
    private Border?    _facilityColContainer; // scroll content 교체용
    private TextBox    _analysisBox   = new();
    private TextBox    _facilityBox   = new();
    private TextBox    _wasteBox      = new();

    public ManualMatchWindow(
        string 시료명, string result,
        List<AnalysisRequestRecord> analysisRecords,
        List<(string 시설명, string 시료명, int 마스터Id)> facilityMasters,
        List<WasteSample> wasteSamples)
    {
        _analysisRecords = analysisRecords;
        _facilityMasters = facilityMasters;
        _wasteSamples    = wasteSamples;
        _targetName      = 시료명;
        System.Diagnostics.Debug.WriteLine($"[ManualMatch] targetName='{시료명}', facilityMasters={facilityMasters.Count}개");
        try { _masterNotes = FacilityResultService.GetMasterNotes(); }
        catch { _masterNotes = new(); }

        Title                 = "수동 매칭";
        Width                 = 900;
        Height                = 600;
        MinWidth              = 600;
        MinHeight             = 400;
        CanResize             = true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        SystemDecorations     = SystemDecorations.Full;
        Background            = AppTheme.BgPrimary;

        // 파비콘
        try
        {
            var uri    = new Uri("avares://ETA/Assets/icons/ETA.ico");
            var stream = Avalonia.Platform.AssetLoader.Open(uri);
            Icon = new WindowIcon(stream);
        }
        catch { }
        FontFamily            = Font;

        // ── 헤더: 미매칭 시료 정보 ────────────────────────────────────────
        var infoBox = new Border
        {
            Background    = Res("GridHeaderBg", "#252535"),
            CornerRadius  = new CornerRadius(6),
            Padding       = new Thickness(10, 8),
            Margin        = new Thickness(0, 0, 0, 8),
            Child = new StackPanel
            {
                Spacing = 2,
                Children =
                {
                    new TextBlock
                    {
                        Text       = $"🔍 미매칭: {시료명}",
                        FontFamily = Font, FontSize = AppTheme.FontLG,
                        FontWeight = FontWeight.Bold,
                        Foreground = AppTheme.FgPrimary,
                    },
                    new TextBlock
                    {
                        Text       = $"결과값: {result}  →  아래에서 매칭할 항목을 선택하세요.",
                        FontFamily = Font, FontSize = AppTheme.FontSM,
                        Foreground = AppTheme.FgMuted,
                    },
                }
            }
        };

        // ── 탭: 의뢰시료 / 폐수배출업소 / 처리장 ────────────────────────
        var tabCtrl = new TabControl { FontFamily = Font, Margin = new Thickness(0, 0, 0, 8) };

        var tabAnalysis  = new TabItem { Header = "의뢰시료",   Content = BuildAnalysisTab() };
        var tabWaste     = _wasteSamples.Count > 0 ? new TabItem { Header = "폐수배출업소", Content = BuildWasteTab() } : null;
        var tabFacility  = new TabItem { Header = "처리장",     Content = BuildFacilityTab() };
        tabCtrl.Items.Add(tabAnalysis);
        if (tabWaste != null) tabCtrl.Items.Add(tabWaste);
        tabCtrl.Items.Add(tabFacility);

        // 시료명에 처리시설 키워드가 있으면 처리장 탭 우선 선택
        bool isFacilityTarget = _facilityMasters.Count > 0 && IsFacilityKeyword(시료명);
        tabCtrl.SelectedItem = isFacilityTarget ? tabFacility : tabAnalysis;

        // ── 닫기 버튼 ─────────────────────────────────────────────────────
        var cancelBtn = new Button
        {
            Content                    = "닫기",
            FontFamily                 = Font,
            FontSize                   = AppTheme.FontBase,
            HorizontalAlignment        = HorizontalAlignment.Right,
            Background                 = Res("BtnBg", "#2a2a3a"),
            Foreground                 = Res("BtnFg", "#aaaacc"),
            BorderBrush                = Res("BtnBorder", "#444466"),
            BorderThickness            = new Thickness(1),
            Padding                    = new Thickness(16, 6),
            CornerRadius               = new CornerRadius(4),
        };
        cancelBtn.Click += (_, _) => Close();

        var root = new DockPanel { Margin = new Thickness(12) };
        DockPanel.SetDock(infoBox,   Dock.Top);
        DockPanel.SetDock(cancelBtn, Dock.Bottom);
        root.Children.Add(infoBox);
        root.Children.Add(cancelBtn);
        root.Children.Add(tabCtrl); // Last = Fill

        Content = root;

        RefreshAnalysisList("");
        RefreshFacilityList("");
        RefreshWasteList("");
    }

    // =========================================================================
    // 탭 빌더
    // =========================================================================
    private Control BuildAnalysisTab()
    {
        _analysisBox = new TextBox
        {
            Watermark      = "🔎 약칭 / 시료명 / 채취일 검색...",
            FontFamily     = Font, FontSize = AppTheme.FontBase,
            Margin         = new Thickness(0, 0, 0, 6),
            Background     = Res("BgInput", "#22223a"),
            Foreground     = AppTheme.FgPrimary,
            BorderBrush    = Res("ThemeBorderSubtle", "#333344"),
            BorderThickness= new Thickness(1),
            CornerRadius   = new CornerRadius(4),
            Padding        = new Thickness(6, 4),
        };
        _analysisBox.TextChanged += (_, _) => RefreshAnalysisList(_analysisBox.Text ?? "");

        var scroll = new ScrollViewer
        {
            Content = _analysisList,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };
        var panel = new DockPanel { Margin = new Thickness(4) };
        DockPanel.SetDock(_analysisBox, Dock.Top);
        panel.Children.Add(_analysisBox);
        panel.Children.Add(scroll);
        return panel;
    }

    private Control BuildFacilityTab()
    {
        _facilityBox = new TextBox
        {
            Watermark       = "🔎 시설명 / 시료명 검색...",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Margin          = new Thickness(0, 0, 0, 6),
            Background      = Res("BgInput", "#22223a"),
            Foreground      = AppTheme.FgPrimary,
            BorderBrush     = Res("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(6, 4),
        };
        _facilityBox.TextChanged += (_, _) => RefreshFacilityList(_facilityBox.Text ?? "");

        var scroll = new ScrollViewer
        {
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
        _facilityColContainer = new Border { Child = _facilityList };
        scroll.Content = _facilityColContainer;

        var panel = new DockPanel { Margin = new Thickness(4) };
        DockPanel.SetDock(_facilityBox, Dock.Top);
        panel.Children.Add(_facilityBox);
        panel.Children.Add(scroll);
        return panel;
    }

    private Control BuildWasteTab()
    {
        _wasteBox = new TextBox
        {
            Watermark       = "🔎 업체명 / SN 검색...",
            FontFamily      = Font, FontSize = AppTheme.FontBase,
            Margin          = new Thickness(0, 0, 0, 6),
            Background      = Res("BgInput", "#22223a"),
            Foreground      = AppTheme.FgPrimary,
            BorderBrush     = Res("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(4),
            Padding         = new Thickness(6, 4),
        };
        _wasteBox.TextChanged += (_, _) => RefreshWasteList(_wasteBox.Text ?? "");

        var scroll = new ScrollViewer
        {
            Content = _wasteList,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };
        var panel = new DockPanel { Margin = new Thickness(4) };
        DockPanel.SetDock(_wasteBox, Dock.Top);
        panel.Children.Add(_wasteBox);
        panel.Children.Add(scroll);
        return panel;
    }

    // =========================================================================
    // 리스트 새로고침
    // =========================================================================
    // ── 추천 점수 (높을수록 좋은 매칭) ─────────────────────────────────────
    private int ScoreAnalysis(AnalysisRequestRecord r)
    {
        if (string.IsNullOrEmpty(_targetName)) return 0;
        var t = _targetName;
        if (r.시료명.Equals(t, StringComparison.OrdinalIgnoreCase)) return 100;
        if (r.약칭.Equals(t, StringComparison.OrdinalIgnoreCase)) return 90;
        if (r.시료명.Contains(t, StringComparison.OrdinalIgnoreCase) ||
            t.Contains(r.시료명, StringComparison.OrdinalIgnoreCase)) return 60;
        if (r.약칭.Contains(t, StringComparison.OrdinalIgnoreCase) ||
            t.Contains(r.약칭, StringComparison.OrdinalIgnoreCase)) return 50;
        return 0;
    }

    // 시료명에 처리시설 키워드 포함 여부 (명시적 키워드 + 마스터 접두어)
    private bool IsFacilityKeyword(string name)
    {
        if (string.IsNullOrEmpty(name)) return false;
        // 명시적 키워드 우선 체크
        string[] keywords = { "중흥", "월내", "율촌", "4단계", "세풍", "해룡", "폐홀" };
        foreach (var kw in keywords)
            if (name.Contains(kw, StringComparison.OrdinalIgnoreCase)) return true;
        // 마스터 접두어 체크
        foreach (var m in _facilityMasters)
        {
            var prefix = FacilityPrefix(m.시설명);
            if (!string.IsNullOrEmpty(prefix) &&
                name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    // 시설명에서 "처리장/처리/사업장/산단" 접미어 제거 → 핵심 이름 추출
    private static string FacilityPrefix(string name)
    {
        foreach (var sfx in new[] { "처리장", "처리", "사업장", "산단", "공단" })
            if (name.EndsWith(sfx) && name.Length > sfx.Length)
                return name[..^sfx.Length];
        return name;
    }

    private int ScoreFacility((string 시설명, string 시료명, int 마스터Id) m)
    {
        if (string.IsNullOrEmpty(_targetName)) return 0;

        var t         = _targetName.Replace(" ", "");
        if (t.Length == 0) return 0;
        var sampleNsp = m.시료명.Replace(" ", "");
        var prefix    = FacilityPrefix(m.시설명);
        var prefixNsp = prefix.Replace(" ", "");
        var combined  = prefixNsp + sampleNsp; // "중흥" + "유입수" = "중흥유입수"

        // ── 1) 완전 일치 ─────────────────────────────────────────────────────
        if (combined.Equals(t, StringComparison.OrdinalIgnoreCase)) return 100;
        if (sampleNsp.Equals(t, StringComparison.OrdinalIgnoreCase)) return 98;

        // ── 2) 포함 관계 ─────────────────────────────────────────────────────
        if (!string.IsNullOrEmpty(combined) &&
            (t.Contains(combined, StringComparison.OrdinalIgnoreCase) ||
             combined.Contains(t, StringComparison.OrdinalIgnoreCase))) return 95;
        if (!string.IsNullOrEmpty(sampleNsp) &&
            (t.Contains(sampleNsp, StringComparison.OrdinalIgnoreCase) ||
             sampleNsp.Contains(t, StringComparison.OrdinalIgnoreCase))) return 90;

        // ── 3) 시설접두어 + 나머지 분리 ──────────────────────────────────────
        bool facilityHit = false;
        if (!string.IsNullOrEmpty(prefixNsp))
        {
            string remainder = "";
            if (t.StartsWith(prefixNsp, StringComparison.OrdinalIgnoreCase))
            { remainder = t[prefixNsp.Length..]; facilityHit = true; }
            else if (t.EndsWith(prefixNsp, StringComparison.OrdinalIgnoreCase))
            { remainder = t[..^prefixNsp.Length]; facilityHit = true; }
            else if (t.Contains(prefixNsp, StringComparison.OrdinalIgnoreCase))
            { facilityHit = true; }

            if (facilityHit && !string.IsNullOrEmpty(remainder) && !string.IsNullOrEmpty(sampleNsp))
            {
                if (sampleNsp.Contains(remainder, StringComparison.OrdinalIgnoreCase) ||
                    remainder.Contains(sampleNsp, StringComparison.OrdinalIgnoreCase))
                    return 88;
            }
        }

        // ── 4) 글자 단위 매칭: 타깃을 1글자씩 쪼개서 combined와 비교 ─────────
        int matchCount = 0;
        var pool = combined.ToLower().ToList();
        foreach (var ch in t.ToLower())
        {
            int idx = pool.IndexOf(ch);
            if (idx >= 0) { matchCount++; pool.RemoveAt(idx); }
        }
        double ratio = (double)matchCount / t.Length;

        // 시설 접두어 타깃 포함 보너스
        int bonus = facilityHit ? 8 : 0;

        // 비율 → 점수 (최대 85)
        int charScore = (int)(ratio * 85) + bonus;
        return Math.Clamp(charScore, 0, 87);
    }

    private int ScoreWaste(WasteSample s)
    {
        if (string.IsNullOrEmpty(_targetName)) return 0;
        var t = _targetName.Replace(" ", "");
        var name = s.업체명.Replace(" ", "");
        if (t.Length == 0 || name.Length == 0) return 0;

        // 1) 완전 일치
        if (name.Equals(t, StringComparison.OrdinalIgnoreCase)) return 100;

        // 2) 접두사(숫자.숫자_ 등) 제거 후 매칭
        string stripped = StripNumericPrefix(t);
        if (!string.IsNullOrEmpty(stripped))
        {
            if (name.Equals(stripped, StringComparison.OrdinalIgnoreCase)) return 98;
            if (name.Contains(stripped, StringComparison.OrdinalIgnoreCase) ||
                stripped.Contains(name, StringComparison.OrdinalIgnoreCase)) return 85;
        }

        // 3) 포함 관계
        if (name.Contains(t, StringComparison.OrdinalIgnoreCase) ||
            t.Contains(name, StringComparison.OrdinalIgnoreCase)) return 70;

        // 4) 글자 단위 유사도
        int matchCount = 0;
        foreach (char c in name)
            if (t.Contains(c, StringComparison.OrdinalIgnoreCase)) matchCount++;
        double ratio = name.Length > 0 ? (double)matchCount / name.Length : 0;
        if (ratio >= 0.6) return (int)(ratio * 60);
        return 0;
    }

    /// <summary>숫자.숫자_ 등 SN 접두사 패턴 제거 (예: "3.11_4" → "4", "03-11-04_금호" → "금호")</summary>
    private static string StripNumericPrefix(string s)
    {
        // "숫자.숫자_나머지" 또는 "숫자-숫자-숫자_나머지" 패턴
        int idx = s.IndexOf('_');
        if (idx >= 0 && idx < s.Length - 1)
        {
            string after = s[(idx + 1)..];
            // 숫자만 남은 경우는 의미없으므로 원본 반환
            if (!after.All(c => char.IsDigit(c) || c == '.' || c == '-'))
                return after;
        }
        return "";
    }

    private void RefreshAnalysisList(string filter)
    {
        _analysisList.Children.Clear();
        var filtered = string.IsNullOrWhiteSpace(filter)
            ? _analysisRecords
            : _analysisRecords.Where(r =>
                r.약칭.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                r.시료명.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                r.채취일자.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();

        if (filtered.Count == 0)
        {
            _analysisList.Children.Add(new TextBlock
            {
                Text = "결과 없음", FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgMuted, Margin = new Thickness(8, 10),
            });
            return;
        }

        // 추천 항목 상단 고정
        var recommended = filtered.Where(r => ScoreAnalysis(r) > 0)
                                  .OrderByDescending(ScoreAnalysis).ToList();
        var rest        = filtered.Where(r => ScoreAnalysis(r) == 0).ToList();

        if (recommended.Count > 0 && string.IsNullOrWhiteSpace(filter))
        {
            _analysisList.Children.Add(new Border
            {
                Background = Res("BtnPrimaryBg", "#1a2a3a"), Padding = new Thickness(8, 3),
                Child = new TextBlock { Text = "⭐ 추천", FontFamily = Font,
                    FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
                    Foreground = AppTheme.FgInfo }
            });
            int ri = 0;
            foreach (var rec in recommended)
            {
                AddAnalysisRow(_analysisList, rec, ri++, highlight: true);
            }
            if (rest.Count > 0)
            {
                _analysisList.Children.Add(new Border
                {
                    Background = Res("GridHeaderBg", "#252535"), Padding = new Thickness(8, 3),
                    Margin = new Thickness(0, 4, 0, 0),
                    Child = new TextBlock { Text = "전체 목록", FontFamily = Font,
                        FontSize = AppTheme.FontSM, Foreground = AppTheme.FgMuted }
                });
            }
            filtered = rest;
        }

        for (int i = 0; i < filtered.Count; i++)
            AddAnalysisRow(_analysisList, filtered[i], i, highlight: false);
    }

    private void AddAnalysisRow(StackPanel panel, AnalysisRequestRecord rec, int idx, bool highlight)
    {
        var bg = highlight
            ? Res("BtnPrimaryBg", "#0d2030")
            : (idx % 2 == 0 ? Res("GridRowBg", "#1e1e28") : Res("GridRowAltBg", "#23233a"));
        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("80,*,100"),
            MinHeight = 32, Background = bg,
        };
        row.Children.Add(MakeTb(rec.채취일자, 0, AppTheme.FgMuted));
        row.Children.Add(MakeTb(
            string.IsNullOrEmpty(rec.약칭) ? rec.시료명 : $"{rec.약칭} ({rec.시료명})",
            1, highlight ? AppTheme.FgSuccess : AppTheme.FgPrimary, leftPad: 6));
        row.Children.Add(MakeTb(rec.접수번호, 2, AppTheme.FgMuted));

        var bdr = new Border
        {
            Child = row, Cursor = new Cursor(StandardCursorType.Hand),
            BorderBrush = Res("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };
        var captured = rec;
        bdr.PointerPressed += (_, _) => ConfirmAnalysis(captured);
        TextShimmer.AttachHover(bdr);
        panel.Children.Add(bdr);
    }

    private void RefreshFacilityList(string filter)
    {
        if (_facilityColContainer == null) return;

        var filtered = string.IsNullOrWhiteSpace(filter)
            ? _facilityMasters
            : _facilityMasters.Where(m =>
                m.시설명.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                m.시료명.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();

        // ── 검색 중: 단순 리스트 ─────────────────────────────────────────────
        if (!string.IsNullOrWhiteSpace(filter))
        {
            _facilityList.Children.Clear();
            if (filtered.Count == 0)
            {
                _facilityList.Children.Add(new TextBlock
                {
                    Text = "결과 없음", FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = AppTheme.FgMuted, Margin = new Thickness(8, 10),
                });
            }
            else
            {
                foreach (var grp in filtered.GroupBy(m => m.시설명))
                    AddFacilityGroup(_facilityList, grp, highlight: false);
            }
            _facilityColContainer.Child = _facilityList;
            return;
        }

        // ── 검색 없음: 시설별 컬럼 레이아웃 ─────────────────────────────────
        int GroupScore(IGrouping<string, (string 시설명, string 시료명, int 마스터Id)> g)
            => g.Max(m => ScoreFacility(m));

        // 시설 순서: 엑셀 기반 동적 순서
        static int FacilityOrder(string name) => FacilityResultService.FacilityOrder(name);

        var byFacility = filtered.GroupBy(m => m.시설명)
            .OrderBy(g => FacilityOrder(g.Key))
            .ThenBy(g => g.Key)
            .ToList();
        foreach (var g in byFacility)
            System.Diagnostics.Debug.WriteLine($"[ManualMatch] col: {g.Key} → order={FacilityOrder(g.Key)}, groupScore={GroupScore(g)}");

        if (byFacility.Count == 0)
        {
            _facilityList.Children.Clear();
            _facilityColContainer.Child = _facilityList;
            return;
        }

        // 시설 수만큼 컬럼 정의
        var colGrid = new Grid { HorizontalAlignment = HorizontalAlignment.Stretch };
        for (int ci = 0; ci < byFacility.Count; ci++)
            colGrid.ColumnDefinitions.Add(new ColumnDefinition(130, GridUnitType.Pixel));
        colGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto)); // 헤더
        colGrid.RowDefinitions.Add(new RowDefinition(GridLength.Star)); // 아이템

        for (int ci = 0; ci < byFacility.Count; ci++)
        {
            var grp       = byFacility[ci];
            int grpScore  = GroupScore(grp);
            // 50 이상이면 그룹 강조 (글자 단위 매칭 반영)
            bool hlGroup  = grpScore >= 50;
            var hdrBg = hlGroup
                ? new SolidColorBrush(Color.Parse("#1e3a10"))
                : Res("GridHeaderBg", "#252535");
            var hdrFg = hlGroup ? new SolidColorBrush(Color.Parse("#88ff88")) : AppTheme.FgInfo;

            // 시설명 헤더 (고정) — ⭐는 헤더가 아니라 아이템에 붙임
            var hdrText = grp.Key;
            var hdr = new Border
            {
                Background = hdrBg,
                BorderBrush = hlGroup
                    ? new SolidColorBrush(Color.Parse("#44aa44"))
                    : Res("ThemeBorderSubtle", "#333344"),
                BorderThickness = new Thickness(0, 0, 1, 2),
                Padding = new Thickness(6, 6),
                Child = new TextBlock
                {
                    Text = hdrText, FontFamily = Font, FontSize = AppTheme.FontSM,
                    FontWeight = FontWeight.Bold, Foreground = hdrFg,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                }
            };
            Grid.SetColumn(hdr, ci); Grid.SetRow(hdr, 0);
            colGrid.Children.Add(hdr);

            // 시료명 목록 (점수순 정렬)
            var items = grp.OrderByDescending(ScoreFacility).ToList();
            // 1순위 점수 (그룹 내 최고점, 50 이상일 때만 별표)
            int topScore = items.Count > 0 ? ScoreFacility(items[0]) : 0;

            var colPanel = new StackPanel { Spacing = 0 };
            int rowIdx = 0;
            foreach (var m in items)
            {
                int sc    = ScoreFacility(m);
                bool rowHl = hlGroup && sc >= 50 && sc == topScore;
                var bg    = rowHl
                    ? new SolidColorBrush(Color.Parse("#2a2200"))
                    : (rowIdx % 2 == 0 ? Res("GridRowBg", "#1e1e28") : Res("GridRowAltBg", "#23233a"));

                _masterNotes.TryGetValue(m.마스터Id, out var note);
                bool hasNote = !string.IsNullOrWhiteSpace(note);

                var inner = new StackPanel { Spacing = 0, Margin = new Thickness(0, 2, 0, 2) };

                // 시료명 행
                var nameRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 2 };
                if (rowHl) nameRow.Children.Add(new TextBlock
                {
                    Text = "⭐", FontSize = AppTheme.FontBase,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 2, 0),
                });
                nameRow.Children.Add(new TextBlock
                {
                    Text = m.시료명, FontFamily = Font, FontSize = AppTheme.FontBase,
                    Foreground = rowHl ? new SolidColorBrush(Color.Parse("#ffe066")) : AppTheme.FgPrimary,
                    FontWeight = rowHl ? FontWeight.SemiBold : FontWeight.Normal,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    Margin = new Thickness(rowHl ? 2 : 8, 0, 4, 0),
                });
                inner.Children.Add(nameRow);

                // 비고 행
                if (hasNote)
                    inner.Children.Add(new TextBlock
                    {
                        Text = note, FontFamily = Font, FontSize = AppTheme.FontXS,
                        Foreground = AppTheme.FgMuted,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        Margin = new Thickness(8, 0, 4, 0),
                    });

                var cellBdr = new Border
                {
                    Background = bg, MinHeight = 28,
                    BorderBrush = Res("ThemeBorderSubtle", "#333344"),
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Child = inner, Cursor = new Cursor(StandardCursorType.Hand),
                    Padding = new Thickness(0, 2),
                };
                var captured = m;
                cellBdr.PointerPressed += (_, _) => ConfirmFacility(captured);
                if (rowHl) TextShimmer.AttachHover(cellBdr);
                colPanel.Children.Add(cellBdr);
                rowIdx++;
            }

            var colScroll = new ScrollViewer
            {
                Content = colPanel,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            };
            Grid.SetColumn(colScroll, ci); Grid.SetRow(colScroll, 1);
            colGrid.Children.Add(colScroll);
        }

        _facilityColContainer.Child = colGrid;
    }

    private void AddFacilityGroup(StackPanel panel,
        IGrouping<string, (string 시설명, string 시료명, int 마스터Id)> grp, bool highlight)
    {
        var hdrBg = highlight ? Res("BtnPrimaryBg", "#0d2030") : Res("GridHeaderBg", "#252535");
        panel.Children.Add(new Border
        {
            Background = hdrBg, Padding = new Thickness(8, 4),
            Child = new TextBlock
            {
                Text = $"🏭 {grp.Key}", FontFamily = Font, FontSize = AppTheme.FontSM,
                FontWeight = FontWeight.SemiBold,
                Foreground = highlight ? AppTheme.FgSuccess : AppTheme.FgInfo,
            }
        });

        // 추천 그룹 내에서 점수 높은 항목 먼저 (침전지 등이 타깃에 포함되면 상단)
        var items = highlight
            ? grp.OrderByDescending(ScoreFacility).ToList()
            : grp.ToList();

        int rowIdx = 0;
        foreach (var m in items)
        {
            int score = ScoreFacility(m);
            bool rowHl = highlight && score >= 95; // 시설접두어+시료명 일치 이상만 강조
            var row = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*"), MinHeight = 32,
                Background = rowHl
                    ? new SolidColorBrush(Color.Parse("#2a2200")) // 노란색 배경
                    : (rowIdx % 2 == 0 ? Res("GridRowBg", "#1e1e28") : Res("GridRowAltBg", "#23233a")),
            };

            if (rowHl)
            {
                // ⭐ 별표 아이콘 추가
                var star = new TextBlock
                {
                    Text = "⭐", FontSize = AppTheme.FontBase,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(14, 0, 2, 0),
                };
                Grid.SetColumn(star, 0);
                row.Children.Add(star);
            }

            var nameTb = MakeTb(m.시료명, 1,
                rowHl ? new SolidColorBrush(Color.Parse("#ffe066")) : AppTheme.FgPrimary,
                leftPad: rowHl ? 4 : 16);
            if (rowHl) nameTb.FontWeight = FontWeight.SemiBold;
            row.Children.Add(nameTb);

            var bdr = new Border
            {
                Child = row, Cursor = new Cursor(StandardCursorType.Hand),
                BorderBrush = rowHl ? new SolidColorBrush(Color.Parse("#665500")) : Res("ThemeBorderSubtle", "#333344"),
                BorderThickness = new Thickness(0, rowHl ? 1 : 0, 0, 1),
            };
            var captured = m;
            bdr.PointerPressed += (_, _) => ConfirmFacility(captured);
            if (rowHl)
                TextShimmer.AttachHover(bdr); // shimmer는 추천 항목만
            panel.Children.Add(bdr);
            rowIdx++;
        }
    }

    private void RefreshWasteList(string filter)
    {
        _wasteList.Children.Clear();
        var filtered = string.IsNullOrWhiteSpace(filter)
            ? _wasteSamples
            : _wasteSamples.Where(s =>
                s.업체명.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                s.SN.Contains(filter, StringComparison.OrdinalIgnoreCase)).ToList();

        if (filtered.Count == 0)
        {
            _wasteList.Children.Add(new TextBlock
            {
                Text = "결과 없음", FontFamily = Font, FontSize = AppTheme.FontSM,
                Foreground = AppTheme.FgMuted, Margin = new Thickness(8, 10),
            });
            return;
        }

        // 추천 항목 상단 고정 (접두사 제외 매칭)
        var recommended = filtered.Where(s => ScoreWaste(s) > 0)
                                  .OrderByDescending(ScoreWaste).ToList();
        var rest = filtered.Where(s => ScoreWaste(s) == 0).ToList();

        if (recommended.Count > 0 && string.IsNullOrWhiteSpace(filter))
        {
            _wasteList.Children.Add(new Border
            {
                Background = Res("BtnPrimaryBg", "#1a2a3a"), Padding = new Thickness(8, 3),
                Child = new TextBlock { Text = "⭐ 추천", FontFamily = Font,
                    FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
                    Foreground = AppTheme.FgInfo }
            });
            int ri = 0;
            foreach (var s in recommended)
                AddWasteRow(_wasteList, s, ri++, highlight: true);

            if (rest.Count > 0)
            {
                _wasteList.Children.Add(new Border
                {
                    Background = Res("GridHeaderBg", "#252535"), Padding = new Thickness(8, 3),
                    Margin = new Thickness(0, 4, 0, 0),
                    Child = new TextBlock { Text = "전체 목록", FontFamily = Font,
                        FontSize = AppTheme.FontSM, Foreground = AppTheme.FgMuted }
                });
            }
            filtered = rest;
        }

        for (int i = 0; i < filtered.Count; i++)
            AddWasteRow(_wasteList, filtered[i], i, highlight: false);
    }

    private void AddWasteRow(StackPanel panel, WasteSample s, int idx, bool highlight)
    {
        int score = ScoreWaste(s);
        bool starred = highlight && score >= 70;
        var bg = starred
            ? new SolidColorBrush(Color.Parse("#2a2200"))
            : (idx % 2 == 0 ? Res("GridRowBg", "#1e1e28") : Res("GridRowAltBg", "#23233a"));

        var row = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("120,Auto,*,60"),
            MinHeight = 32, Background = bg,
        };
        row.Children.Add(MakeTb(s.SN, 0, AppTheme.FgInfo));

        if (starred)
        {
            var star = new TextBlock
            {
                Text = "⭐", FontSize = AppTheme.FontBase,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(2, 0),
            };
            Grid.SetColumn(star, 1); row.Children.Add(star);
        }

        var nameTb = MakeTb(s.업체명, 2,
            starred ? new SolidColorBrush(Color.Parse("#ffe066")) : AppTheme.FgPrimary,
            leftPad: starred ? 2 : 6);
        if (starred) nameTb.FontWeight = FontWeight.SemiBold;
        row.Children.Add(nameTb);
        row.Children.Add(MakeTb(s.구분, 3, AppTheme.FgMuted));

        var bdr = new Border
        {
            Child = row, Cursor = new Cursor(StandardCursorType.Hand),
            BorderBrush = Res("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };
        var captured = s;
        bdr.PointerPressed += (_, _) => ConfirmWaste(captured);
        TextShimmer.AttachHover(bdr);
        panel.Children.Add(bdr);
    }

    // =========================================================================
    // 선택 확인
    // =========================================================================
    private void ConfirmAnalysis(AnalysisRequestRecord rec)
    {
        SelectedAnalysis = rec;
        MatchConfirmed?.Invoke(this);
        Close();
    }

    private void ConfirmFacility((string 시설명, string 시료명, int 마스터Id) m)
    {
        SelectedFacility = m;
        MatchConfirmed?.Invoke(this);
        Close();
    }

    private void ConfirmWaste(WasteSample s)
    {
        SelectedWaste = s;
        MatchConfirmed?.Invoke(this);
        Close();
    }

    // =========================================================================
    // 헬퍼
    // =========================================================================
    private TextBlock MakeTb(string text, int col, IBrush fg, double leftPad = 4)
    {
        var tb = new TextBlock
        {
            Text                = text,
            FontFamily          = Font,
            FontSize            = AppTheme.FontBase,
            Foreground          = fg,
            VerticalAlignment   = VerticalAlignment.Center,
            TextTrimming        = TextTrimming.CharacterEllipsis,
            Margin              = new Thickness(leftPad, 0, 4, 0),
        };
        Grid.SetColumn(tb, col);
        return tb;
    }
}
