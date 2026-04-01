using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ETA.Models;
using ETA.Services.SERVICE1;

namespace ETA.Views.Pages.PAGE1;

/// <summary>시험성적서 미리보기 + 인쇄/저장 윈도우</summary>
public sealed class TestReportPrintWindow : Window
{
    // ── 폰트 ───────────────────────────────────────────────────────────────
    private static readonly FontFamily FM = new("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M");

    // ── 색상 (템플릿: 배경색 없음) ──────────────────────────────────────────
    private static readonly IBrush BgPage  = new SolidColorBrush(Color.Parse("#ffffff"));
    private static readonly IBrush BgLabel = new SolidColorBrush(Color.Parse("#efefef")); // 연회색
    private static readonly IBrush BgHdr   = new SolidColorBrush(Color.Parse("#d8d8d8")); // 컬럼헤더
    private static readonly IBrush FgDark  = new SolidColorBrush(Color.Parse("#111111"));
    private static readonly IBrush FgGray  = new SolidColorBrush(Color.Parse("#555555"));
    private static readonly IBrush BrdClr  = new SolidColorBrush(Color.Parse("#aaaaaa"));
    private static readonly IBrush ToolBg  = new SolidColorBrush(Color.Parse("#1a2a1a"));

    // ── 컬럼 비율 (템플릿 문자폭 기준) ────────────────────────────────────
    // 헤더 구역:   A+B(19.375) | C+D(48.5) | E(21.25) | F+G+H(47)
    private const string InfoCols = "19.375*,48.5*,21.25*,47*";
    // 데이터 구역: A(8) | B+C(22.75) | D(37.125) | E(21.25) | F(13.25) | G(9.625) | H(24.125)
    private const string DataCols = "8*,22.75*,37.125*,21.25*,13.25*,9.625*,24.125*";

    private const double RowH = 22;   // 화면 픽셀 행 높이

    // ── 데이터 ─────────────────────────────────────────────────────────────
    private readonly SampleRequest           _sample;
    private readonly List<AnalysisResultRow> _rows;
    private readonly string                  _reportNo;
    private readonly string                  _qualityMgr;
    private TextBlock?                       _statusTb;

    // ═══════════════════════════════════════════════════════════════════════
    public TestReportPrintWindow(
        SampleRequest           sample,
        List<AnalysisResultRow> rows,
        string                  reportNo   = "",
        string                  qualityMgr = "")
    {
        _sample     = sample;
        _rows       = rows;
        _reportNo   = reportNo;
        _qualityMgr = qualityMgr;

        Title                 = $"시험성적서 미리보기 — {sample.시료명}";
        Width                 = 960;
        Height                = 780;
        MinWidth              = 700;
        MinHeight             = 500;
        Background            = new SolidColorBrush(Color.Parse("#222222"));
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        CanResize             = true;
        BuildLayout();
    }

    // ═══════════════════════════════════════════════════════════════════════
    private void BuildLayout()
    {
        var root = new Grid { RowDefinitions = new RowDefinitions("Auto,*,Auto") };
        root.Children.Add(BuildToolbar());

        var scroll = new ScrollViewer
        {
            [Grid.RowProperty]            = 1,
            Background                    = new SolidColorBrush(Color.Parse("#333333")),
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            Padding                       = new Thickness(24),
            Content                       = BuildAllPages(),
        };
        root.Children.Add(scroll);

        _statusTb = new TextBlock
        {
            [Grid.RowProperty] = 2,
            FontFamily = FM, FontSize = 11,
            Foreground = new SolidColorBrush(Color.Parse("#88cc88")),
            Margin     = new Thickness(12, 4),
        };
        root.Children.Add(_statusTb);
        Content = root;
    }

    private Border BuildToolbar()
    {
        var btnPrint = Btn("🖨  인쇄",      "#1a3a1a", "#7cd87c");
        var btnPdf   = Btn("💾  PDF 저장",  "#1a2a3a", "#7aaae8");
        var btnXlsx  = Btn("📄  Excel 저장","#2a2a1a", "#d8c87c");
        var btnClose = Btn("✕  닫기",       "#2a1a1a", "#cc7777");

        btnPrint.Click += async (_, _) => await DoPrintAsync();
        btnPdf.Click   += async (_, _) => await DoSaveAsync(toPdf: true);
        btnXlsx.Click  += async (_, _) => await DoSaveAsync(toPdf: false);
        btnClose.Click += (_, _) => Close();

        return new Border
        {
            [Grid.RowProperty] = 0,
            Background = ToolBg,
            Padding    = new Thickness(12, 8),
            Child      = new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 8,
                Children    = { btnPrint, btnPdf, btnXlsx, btnClose },
            },
        };
    }

    private static Button Btn(string t, string bg, string fg) => new()
    {
        Content = t, FontFamily = new FontFamily("avares://ETA/Assets/Fonts#KBIZ한마음고딕 M"),
        FontSize = 12, Padding = new Thickness(14, 6), CornerRadius = new CornerRadius(4),
        BorderThickness = new Thickness(0),
        Background = new SolidColorBrush(Color.Parse(bg)),
        Foreground = new SolidColorBrush(Color.Parse(fg)),
    };

    // ═══════════════════════════════════════════════════════════════════════
    // 전체 페이지 빌드
    // ═══════════════════════════════════════════════════════════════════════
    private StackPanel BuildAllPages()
    {
        var outer  = new StackPanel { Spacing = 28 };
        int total  = Math.Max(1, (int)Math.Ceiling(_rows.Count / 32.0));

        var stdMap              = BuildStandardMap();
        var (company, repres)   = GetContractInfo();
        bool isQC               = IsQC();
        string no               = string.IsNullOrEmpty(_reportNo)
            ? $"WAC-{DateTime.Now:yyyyMMdd}-{_sample.약칭}" : _reportNo;

        for (int p = 0; p < total; p++)
        {
            var pageRows = _rows.Skip(p * 32).Take(32).ToList();
            outer.Children.Add(BuildPage(pageRows, no, company, repres, isQC, stdMap, p, total));
        }
        return outer;
    }

    // ─── 1개 페이지 ────────────────────────────────────────────────────────
    private Border BuildPage(
        List<AnalysisResultRow> rows,
        string no, string company, string repres, bool isQC,
        Dictionary<string, string> stdMap, int idx, int total)
    {
        string suffix = idx == 0 ? "-A" : $"-A-{idx + 1}";
        string sampler = $"{_sample.시료채취자1} {_sample.시료채취자2}".Trim();

        var sp = new StackPanel
        {
            Background = BgPage,
            Margin     = new Thickness(0),
        };

        // ── 페이지 헤더: ANALYSIS REPORT ──────────────────────────────
        sp.Children.Add(new Border
        {
            Padding    = new Thickness(8, 12, 8, 4),
            Background = BgPage,
            Child      = new TextBlock
            {
                Text                = total > 1 ? $"ANALYSIS REPORT  ({idx+1}/{total})" : "ANALYSIS REPORT",
                FontSize            = 26,
                FontWeight          = FontWeight.Bold,
                FontFamily          = FM,
                Foreground          = FgDark,
                HorizontalAlignment = HorizontalAlignment.Center,
            },
        });

        // ── Row 1: 성적서 번호 ─────────────────────────────────────────
        sp.Children.Add(MakeRow1(no + suffix));

        // ── Row 2: 빈 행 ───────────────────────────────────────────────
        sp.Children.Add(new Border { Height = RowH, Background = BgPage });

        // ── Rows 3~5: 업체/채수/입회 정보 ─────────────────────────────
        sp.Children.Add(MakeInfoRow("업 체 명",
            string.IsNullOrEmpty(company) ? _sample.의뢰사업장 : company,
            "채 수 일 자", FormatDate(_sample.채취일자)));
        sp.Children.Add(MakeInfoRow("대 표 자", repres,
            "채수담당자", sampler));
        sp.Children.Add(MakeInfoRow("채수입회자", _sample.입회자,
            "분석완료일", FormatDate(_sample.분석종료일)));

        // ── Row 6: 빈 행 ───────────────────────────────────────────────
        sp.Children.Add(new Border { Height = RowH, Background = BgPage });

        // ── Row 7: 시료명 / 의뢰정보 ─────────────────────────────────
        sp.Children.Add(MakeInfoRow("시 료 명", _sample.시료명,
            "의뢰정보(용도)", isQC ? "정도보증 적용" : "참고용"));

        // ── Row 8: 시험결과 배너 ───────────────────────────────────────
        sp.Children.Add(MakeRow8());

        // ── Row 9: 컬럼 헤더 ──────────────────────────────────────────
        sp.Children.Add(MakeColHeaders());

        // ── Rows 10~41: 데이터 32행 ───────────────────────────────────
        for (int i = 0; i < 32; i++)
        {
            if (i < rows.Count)
            {
                var r = rows[i];
                sp.Children.Add(MakeDataRow(idx * 32 + i + 1,
                    r.Category ?? "", r.항목명 ?? "", r.ES ?? "",
                    FormatResult(r.결과값), r.단위 ?? "",
                    stdMap.GetValueOrDefault(r.항목명 ?? "", "")));
            }
            else
            {
                sp.Children.Add(MakeEmptyDataRow());
            }
        }

        // ── Row 42: 마감선 ────────────────────────────────────────────
        sp.Children.Add(new Border
        {
            BorderBrush     = new SolidColorBrush(Color.Parse("#000000")),
            BorderThickness = new Thickness(0, 2, 0, 0),
            Margin          = new Thickness(4, 0),
        });

        // ── Row 43: 서명 ──────────────────────────────────────────────
        if (!string.IsNullOrEmpty(_qualityMgr))
        {
            sp.Children.Add(new Border
            {
                Height     = RowH,
                Background = BgPage,
                Padding    = new Thickness(8, 0),
                Child      = new TextBlock
                {
                    Text = $"품질책임 수질분야 환경측정분析사       {_qualityMgr}       (서명)",
                    FontSize = 9, FontFamily = FM, Foreground = FgDark,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                },
            });
        }

        // ── Row 44: 면책고지 ──────────────────────────────────────────
        sp.Children.Add(new Border
        {
            MinHeight  = 32,
            Background = BgPage,
            Padding    = new Thickness(8, 4),
            Child      = new TextBlock
            {
                Text = isQC
                    ? "▩ 이 시험성적서는 ES 04001.b(정도보증/관리) 등 국립환경과학원고시 『수질오염공정시험기준』을 적용한 분析결과 입니다."
                    : "▩ 이 시험성적서는 ES 04001.b/04130.1e 등 일부가 적용되지 않는 참고용 분析결과입니다.",
                FontSize = 8, FontFamily = FM, Foreground = FgGray,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            },
        });

        // ── 페이지 푸터 ───────────────────────────────────────────────
        sp.Children.Add(new Border
        {
            Padding    = new Thickness(8, 4, 8, 12),
            Background = BgPage,
            Child      = new TextBlock
            {
                Text = "리뉴어스주식회사 - 수질분析센터",
                FontSize = 14, FontWeight = FontWeight.Bold,
                FontFamily = FM, Foreground = FgDark,
                HorizontalAlignment = HorizontalAlignment.Center,
            },
        });

        return new Border
        {
            Background      = BgPage,
            BorderBrush     = new SolidColorBrush(Color.Parse("#888888")),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(2),
            Child           = sp,
            MaxWidth        = 900,
        };
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 행 빌더
    // ═══════════════════════════════════════════════════════════════════════

    // Row 1: 성적서번호
    private Grid MakeRow1(string reportNo)
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(InfoCols), Height = RowH };
        g.Children.Add(Cell("성 적 서  번 호", 0, isLbl: true, align: TA.Center));
        g.Children.Add(Cell(reportNo,          1, isLbl: false, align: TA.Left));
        // E1 column: empty (address - 템플릿 고정값)
        g.Children.Add(Cell("",                2, isLbl: false, align: TA.Left));
        // F+G+H: 빈 칸
        g.Children.Add(Cell("",                3, isLbl: false, align: TA.Left));
        return g;
    }

    // Rows 3/4/5/7: 정보 행
    private Grid MakeInfoRow(string lbl1, string val1, string lbl2, string val2)
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(InfoCols), Height = RowH };
        g.Children.Add(Cell(lbl1, 0, isLbl: true,  align: TA.Center));
        g.Children.Add(Cell(val1, 1, isLbl: false,  align: TA.Left));
        g.Children.Add(Cell(lbl2, 2, isLbl: true,  align: TA.Center));
        g.Children.Add(Cell(val2, 3, isLbl: false,  align: TA.Left));
        return g;
    }

    // Row 8: 시험결과 배너
    private Grid MakeRow8()
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(InfoCols), Height = RowH };
        g.Children.Add(Cell("시험결과", 0, isLbl: true, align: TA.Center));
        // 나머지 열: 하단 선만 (빈 셀 + 하단 보더)
        var rest = new Border
        {
            [Grid.ColumnProperty] = 1, [Grid.ColumnSpanProperty] = 3,
            BorderBrush = BrdClr, BorderThickness = new Thickness(0, 0, 0, 1),
        };
        g.Children.Add(rest);
        return g;
    }

    // Row 9: 컬럼 헤더
    private Grid MakeColHeaders()
    {
        var g = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(DataCols),
            Height = RowH,
            // 상/하 굵은 선
            Margin = new Thickness(0),
        };
        string[] labels = { "번호", "항목 구분", "시험 항목", "적용 시험방법", "결과", "단위", "비고" };
        for (int i = 0; i < 7; i++)
            g.Children.Add(Cell(labels[i], i, isLbl: true, align: TA.Center, bg: BgHdr));

        return new Grid
        {
            RowDefinitions = new RowDefinitions("Auto"),
            Children =
            {
                new Border
                {
                    BorderBrush     = new SolidColorBrush(Color.Parse("#000000")),
                    BorderThickness = new Thickness(0, 2, 0, 2),
                    Child           = g,
                    Margin          = new Thickness(4, 0),
                },
            },
        };
    }

    // Rows 10~41: 데이터
    private Grid MakeDataRow(int no, string cat, string item, string method,
        string result, string unit, string standard)
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(DataCols), Height = RowH };
        g.Children.Add(Cell(no.ToString(), 0, false, TA.Center));
        g.Children.Add(Cell(cat,           1, false, TA.Left));
        g.Children.Add(Cell(item,          2, false, TA.Left));
        g.Children.Add(Cell(method,        3, false, TA.Center));
        g.Children.Add(Cell(result,        4, false, TA.Right));
        g.Children.Add(Cell(unit,          5, false, TA.Center));
        g.Children.Add(Cell(standard,      6, false, TA.Center));
        return g;
    }

    private Grid MakeEmptyDataRow()
    {
        var g = new Grid { ColumnDefinitions = new ColumnDefinitions(DataCols), Height = RowH };
        for (int i = 0; i < 7; i++) g.Children.Add(Cell("", i, false, TA.Left));
        return g;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 셀 빌더
    // ═══════════════════════════════════════════════════════════════════════
    private enum TA { Left, Center, Right }

    private Border Cell(string text, int col, bool isLbl, TA align, IBrush? bg = null)
    {
        var b = new Border
        {
            [Grid.ColumnProperty] = col,
            Background            = bg ?? (isLbl ? BgLabel : BgPage),
            BorderBrush           = BrdClr,
            BorderThickness       = new Thickness(0, 0, 1, 1),
            Padding               = new Thickness(3, 0),
        };

        if (!string.IsNullOrEmpty(text))
        {
            b.Child = new TextBlock
            {
                Text              = text,
                FontSize          = isLbl ? 9 : 8.5,
                FontFamily        = FM,
                Foreground        = FgDark,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = align == TA.Center ? HorizontalAlignment.Center
                                    : align == TA.Right  ? HorizontalAlignment.Right
                                    :                      HorizontalAlignment.Left,
                TextTrimming      = Avalonia.Media.TextTrimming.CharacterEllipsis,
            };
        }
        return b;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 인쇄 / 저장
    // ═══════════════════════════════════════════════════════════════════════
    private async Task DoPrintAsync()
    {
        SetStatus("⏳ 인쇄 준비 중...");
        try
        {
            var stdMap            = BuildStandardMap();
            var (company, repres) = GetContractInfo();
            bool isQC             = IsQC();
            var sample            = _sample;
            var rows              = _rows;
            var reportNo          = _reportNo;
            var qualMgr           = _qualityMgr;

            await Task.Run(() =>
                TestReportGdiPrinter.Print(sample, rows, stdMap,
                    reportNo, qualMgr, company, repres, isQC));

            SetStatus("✅ 인쇄 완료");
        }
        catch (Exception ex) { SetStatus($"❌ 오류: {ex.Message}"); }
    }

    private async Task DoSaveAsync(bool toPdf)
    {
        SetStatus(toPdf ? "⏳ PDF 저장 중..." : "⏳ Excel 저장 중...");
        try
        {
            var path = await Task.Run(() =>
                TestReportPrintService.FillAndSave(
                    _sample, _rows,
                    meta:       new Dictionary<string, AnalysisItem>(),
                    reportNo:   _reportNo,
                    qualityMgr: _qualityMgr,
                    toPdf:      toPdf,
                    openAfter:  true));
            SetStatus($"✅ 저장 완료: {Path.GetFileName(path)}");
        }
        catch (Exception ex) { SetStatus($"❌ 오류: {ex.Message}"); }
    }

    private void SetStatus(string msg) =>
        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        { if (_statusTb != null) _statusTb.Text = msg; });

    // ═══════════════════════════════════════════════════════════════════════
    // 데이터 헬퍼
    // ═══════════════════════════════════════════════════════════════════════
    private Dictionary<string, string> BuildStandardMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(_sample.방류허용기준)) return map;
        foreach (var r in _rows)
        {
            if (string.IsNullOrEmpty(r.항목명) || map.ContainsKey(r.항목명)) continue;
            var val = TestReportService.GetStandardValue(r.항목명, _sample.방류허용기준);
            if (!string.IsNullOrEmpty(val)) map[r.항목명] = val;
        }
        return map;
    }

    private (string company, string repres) GetContractInfo()
    {
        try { return (_sample.의뢰사업장, ContractService.GetRepresentativeByCompany(_sample.의뢰사업장)); }
        catch { return (_sample.의뢰사업장, ""); }
    }

    private bool IsQC() =>
        _sample.정도보증.Trim().ToUpper() is "O" or "Y" ||
        _sample.정도보증.Trim() == "정도보증 적용";

    private static string FormatDate(string raw) =>
        string.IsNullOrWhiteSpace(raw) ? "" :
        DateTime.TryParse(raw, out var dt) ? dt.ToString("yyyy-MM-dd") : raw;

    private static string FormatResult(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        raw = raw.Trim();
        if (raw.StartsWith('<') || raw.StartsWith('>')) return raw;
        return decimal.TryParse(raw,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var d)
            ? d.ToString("G6") : raw;
    }
}
