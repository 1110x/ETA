using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// 시험기록부 뷰어 Show2 — 분석결과입력(WasteAnalysisInputPage) 의 파서 결과 뷰와
/// 동일한 외형으로 저장된 *_시험기록부 행을 표시하는 정적 빌더.
/// 분석결과입력의 LoadVerifiedGrid + BuildDocRowUnified 패턴을 그대로 따름.
/// </summary>
public static class TestRecordBookParsedView
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    public class Model
    {
        public string AnalysisDate   = "";
        public string CategoryKey    = "";
        public string FileLabel      = "";
        public string TargetTable    = "";
        public string AnalysisMethod = "";
        public string Memo           = "";
        public string ES             = "";   // 분석정보.ES (시험방법코드)
        public string Instrument     = "";   // 분석정보.instrument (분석장비)
        public string Method         = "";   // 분석정보.Method (시험법명)

        public List<int>    StandardKeys = new();
        public List<string> StandardConc = new();
        public List<string> StandardAbs  = new();
        public List<string> StandardIstd = new();   // VOC 등 ISTD 응답값 (없으면 빈 리스트)
        public string SlopeText = "";
        public string R2Text    = "";

        // TOC TCIC 전용 — TC/IC 두 세트
        public bool         IsTcic         = false;
        public List<int>    TcStandardKeys = new();
        public List<string> TcStandardConc = new();
        public List<string> TcStandardAbs  = new();
        public string       TcSlopeText    = "";
        public string       TcR2Text       = "";
        public List<int>    IcStandardKeys = new();
        public List<string> IcStandardConc = new();
        public List<string> IcStandardAbs  = new();
        public string       IcSlopeText    = "";
        public string       IcR2Text       = "";

        public List<string>       SampleHeaders    = new();
        public List<List<string>> SampleRows       = new();
        public List<string>       SampleClassByRow = new();

        // 검정곡선의 보증 (MBK/FBK/CCV 등 정도관리 시료)
        public List<List<string>> QcRows           = new();
        public List<string>       QcClassByRow    = new();

        // BOD (생물화학적 산소요구량) 전용 — 식종 정보 영역 (검정곡선 자리)
        public bool IsBod = false;
        public List<string>       SeedHeaders = new();   // 시료명 | V | D1(B1) | D2(B2) | f | P | Result | Remark
        public List<List<string>> SeedRows    = new();   // 식종수의 BOD / SCF(식종희석수)
    }

    // ── 인쇄용 흰 바탕 팔레트 (현재 모드 색상 무시하고 고정) ──
    private const bool PRINT = true;

    private static readonly IBrush PaperBg     = new SolidColorBrush(Color.Parse("#ffffff"));
    private static readonly IBrush RowAltBg    = new SolidColorBrush(Color.Parse("#f5f5f5"));
    private static readonly IBrush HeaderBg    = new SolidColorBrush(Color.Parse("#e8e8e8"));
    private static readonly IBrush BorderClr   = new SolidColorBrush(Color.Parse("#888888"));
    private static readonly IBrush BorderSub   = new SolidColorBrush(Color.Parse("#c0c0c0"));
    private static readonly IBrush FgBlack     = new SolidColorBrush(Color.Parse("#000000"));
    private static readonly IBrush FgDark      = new SolidColorBrush(Color.Parse("#222222"));
    private static readonly IBrush FgMuted2    = new SolidColorBrush(Color.Parse("#555555"));
    private static readonly IBrush FgDim       = new SolidColorBrush(Color.Parse("#999999"));
    private static readonly IBrush ResultGreen = new SolidColorBrush(Color.Parse("#1a7e3a"));
    private static readonly IBrush ResultRed   = new SolidColorBrush(Color.Parse("#c0392b"));
    private static readonly IBrush AccentBar   = new SolidColorBrush(Color.Parse("#ff6a3d"));

    private static IBrush AppRes(string key, string fallback = "#888888")
    {
        if (PRINT)
        {
            return key switch
            {
                "GridHeaderBg"        => HeaderBg,
                "GridRowBg"           => PaperBg,
                "GridRowAltBg"        => RowAltBg,
                "ThemeBorderSubtle"   => BorderSub,
                "AppFg"               => FgBlack,
                "FgMuted"             => FgMuted2,
                "FgDimmed"            => FgDim,
                "StatusInfoBg"        => new SolidColorBrush(Color.Parse("#e3eef9")),
                "StatusInfoFg"        => new SolidColorBrush(Color.Parse("#1f4e79")),
                "StatusAccentFg"      => AccentBar,
                _                     => new SolidColorBrush(Color.Parse(fallback)),
            };
        }
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is IBrush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    public static Control Build(Model m)
    {
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(8) };

        // 1) 타이틀: "{Analyte} 시험기록부"
        root.Children.Add(BuildTitle(m));

        // 2) 분석정보 (분석일·표준·기기·방법·비고) + 분석조건 토글
        bool hasMethodLine = !string.IsNullOrWhiteSpace(m.AnalysisMethod) || !string.IsNullOrWhiteSpace(m.Memo)
                          || !string.IsNullOrWhiteSpace(m.ES) || !string.IsNullOrWhiteSpace(m.Instrument)
                          || !string.IsNullOrWhiteSpace(m.Method);
        var methodHost = new ContentControl();
        root.Children.Add(BuildSectionTitleWithToggle("분석정보", _showMethod, isOn =>
        {
            _showMethod = isOn;
            methodHost.Content = (_showMethod && hasMethodLine) ? BuildMethodLine(m) : null;
        }));
        root.Children.Add(BuildHeader(m));
        if (hasMethodLine)
        {
            methodHost.Content = _showMethod ? BuildMethodLine(m) : null;
            root.Children.Add(methodHost);
        }

        // 3) 검정곡선 (단일 또는 TCIC 듀얼)
        if (m.IsTcic || m.StandardKeys.Count > 0)
            root.Children.Add(BuildSectionTitle("검정곡선"));
        if (m.IsTcic)
        {
            if (m.TcStandardKeys.Count > 0)
                root.Children.Add(BuildCalibrationTableTcic(m, "TC", m.TcStandardKeys, m.TcStandardConc, m.TcStandardAbs, m.TcSlopeText, m.TcR2Text));
            if (m.IcStandardKeys.Count > 0)
                root.Children.Add(BuildCalibrationTableTcic(m, "IC", m.IcStandardKeys, m.IcStandardConc, m.IcStandardAbs, m.IcSlopeText, m.IcR2Text));
        }
        else if (m.StandardKeys.Count > 0)
            root.Children.Add(BuildCalibrationTable(m));

        // 4) 검정곡선의 보증 (또는 BOD 식종 정보) — QC 호스트 (계산식 토글로 재빌드)
        var qcHost = new ContentControl();
        if (m.IsBod)
        {
            if (m.SeedRows.Count > 0)
                root.Children.Add(BuildSeedSection(m));
        }
        else
        {
            qcHost.Content = BuildQcSection(m);
            root.Children.Add(qcHost);
        }

        // 5) 시료분석결과 (+ 계산식 표시 토글) — 토글 변경 시 QC 와 시료 둘 다 재빌드
        if (m.SampleRows.Count > 0)
        {
            var sampleHost = new ContentControl();
            root.Children.Add(BuildSectionTitleWithToggle("시료분석결과", _showFormula, isOn =>
            {
                _showFormula = isOn;
                sampleHost.Content = BuildSampleGrid(m);
                if (!m.IsBod) qcHost.Content = BuildQcSection(m);
            }));
            sampleHost.Content = BuildSampleGrid(m);
            root.Children.Add(sampleHost);
        }

        // 6) 서명란
        root.Children.Add(BuildSignatureBox());

        // 인쇄 모드: ScrollViewer 래핑 없이 root 그대로 (전체 높이 캡처)
        if (IsPrintMode) return root;

        return new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    // ── 토글 상태 (정적 — 페이지 단위 유지) ───────────────────
    private static bool _showFormula = true;   // 계산식 컬럼
    private static bool _showMethod  = true;   // 분석조건 라인
    public  static bool IsPrintMode  = false;  // true 이면 토글 UI 빼고 plain 타이틀만

    // ── 섹션 타이틀 + 우측 토글 스위치 ───────────────────────
    private static Control BuildSectionTitleWithToggle(string text, bool initial, System.Action<bool> onChanged)
    {
        // 인쇄 모드에서는 토글 제거 — 종이에 UI 컨트롤 안 찍히게
        if (IsPrintMode) return BuildSectionTitle(text);

        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto"),
            Margin = new Thickness(0, 8, 0, 4),
        };
        var title = (StackPanel)BuildSectionTitle(text);
        title.Margin = new Thickness(0);
        Grid.SetColumn(title, 0);
        grid.Children.Add(title);

        var toggle = new ToggleSwitch
        {
            IsChecked = initial,
            OnContent = "표시", OffContent = "숨김",
            FontFamily = Font, FontSize = 11,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Right,
            MinHeight = 22,
        };
        toggle.IsCheckedChanged += (_, _) => onChanged(toggle.IsChecked == true);
        Grid.SetColumn(toggle, 1);
        grid.Children.Add(toggle);
        return grid;
    }

    // ── 시험기록부 타이틀 ─────────────────────────────────────
    private static Control BuildTitle(Model m)
    {
        // 테이블명 → "Analyte 시험기록부" 라벨 (예: "Bromoform 시험기록부")
        string analyte = ETA.Services.SERVICE1.TestRecordBookViewerService.PrettyName(m.TargetTable);
        if (string.IsNullOrWhiteSpace(analyte)) analyte = m.TargetTable;
        return new TextBlock
        {
            Text = $"{analyte} 시험기록부",
            FontFamily = Font, FontSize = 18, FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg", "#000000"),
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 4, 0, 8),
        };
    }

    // ── 섹션 타이틀 (좌측 액센트 바 + 제목) ────────────────────
    private static Control BuildSectionTitle(string text)
    {
        var sp = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 6,
            Margin = new Thickness(0, 8, 0, 4),
            VerticalAlignment = VerticalAlignment.Center,
        };
        sp.Children.Add(new Border
        {
            Width = 4, Height = 14,
            Background = AppRes("StatusAccentFg", "#ff6a3d"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(new TextBlock
        {
            Text = text, FontFamily = Font, FontSize = 13, FontWeight = FontWeight.SemiBold,
            Foreground = AppRes("AppFg", "#000000"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        return sp;
    }

    // ── 서명란 (담당자 / 기술책임자) ──────────────────────────
    private static Control BuildSignatureBox()
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,*"),
            MinHeight = 60, Margin = new Thickness(0, 16, 0, 4),
        };
        Border SigCell(string label)
        {
            var inner = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
            inner.Children.Add(new TextBlock
            {
                Text = label, FontFamily = Font, FontSize = 11, FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted", "#888899"),
                Margin = new Thickness(8, 4),
            });
            // 서명 자리 (빈 박스)
            var signPad = new Border
            {
                Margin = new Thickness(8, 0, 8, 8),
                MinHeight = 36,
                BorderBrush = AppRes("ThemeBorderSubtle", "#888888"),
                BorderThickness = new Thickness(0, 1, 0, 0),
            };
            Grid.SetRow(signPad, 1);
            inner.Children.Add(signPad);
            return new Border
            {
                Child = inner,
                BorderBrush = AppRes("ThemeBorderSubtle", "#888888"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 4, 0),
            };
        }
        var c1 = SigCell("담당자");
        var c2 = SigCell("기술책임자");
        Grid.SetColumn(c1, 0); Grid.SetColumn(c2, 1);
        c2.Margin = new Thickness(4, 0, 0, 0);
        grid.Children.Add(c1); grid.Children.Add(c2);
        return grid;
    }

    // ── 헤더 바 ───────────────────────────────────────────────
    private static Control BuildHeader(Model m)
    {
        var sp = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 8,
            VerticalAlignment = VerticalAlignment.Center,
        };
        sp.Children.Add(new TextBlock
        {
            Text = "📄", FontFamily = Font, FontSize = 14,
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(new TextBlock
        {
            Text = $"분석일: {m.AnalysisDate}",
            FontFamily = Font, FontSize = 13, FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg", "#ffffff"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        if (!string.IsNullOrWhiteSpace(m.CategoryKey))
            sp.Children.Add(MakeChip($"[{m.CategoryKey}]", "StatusInfoBg", "StatusInfoFg", "#1e2a3e", "#88aadd"));
        if (!string.IsNullOrWhiteSpace(m.FileLabel))
            sp.Children.Add(new TextBlock
            {
                Text = m.FileLabel, FontFamily = Font, FontSize = 11,
                Foreground = AppRes("FgMuted", "#888899"),
                VerticalAlignment = VerticalAlignment.Center,
            });
        sp.Children.Add(new TextBlock
        {
            Text = "→", FontFamily = Font, FontSize = 11,
            Foreground = AppRes("FgDimmed", "#555566"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(new TextBlock
        {
            Text = m.TargetTable, FontFamily = Font, FontSize = 12,
            Foreground = AppRes("AppFg", "#ffffff"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        return sp;
    }

    private static Control MakeChip(string text, string bgKey, string fgKey, string bgFb, string fgFb)
        => new Border
        {
            Background = AppRes(bgKey, bgFb),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(6, 1),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = text, FontFamily = Font, FontSize = 11, FontWeight = FontWeight.SemiBold,
                Foreground = AppRes(fgKey, fgFb),
            }
        };

    private static Control BuildMethodLine(Model m)
    {
        var sp = new StackPanel { Spacing = 2, Margin = new Thickness(20, 4, 0, 6) };

        // 한 줄에 ES / 시험방법 / 장비 칩 가로 배치
        var line1 = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 8,
            VerticalAlignment = VerticalAlignment.Center,
        };
        void AddLabelValue(string label, string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return;
            line1.Children.Add(new TextBlock
            {
                Text = label, FontFamily = Font, FontSize = 11,
                Foreground = AppRes("FgDimmed", "#555566"),
                VerticalAlignment = VerticalAlignment.Center,
            });
            line1.Children.Add(new TextBlock
            {
                Text = val, FontFamily = Font, FontSize = 11,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("AppFg", "#ffffff"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0),
            });
        }
        AddLabelValue("ES",     m.ES);
        AddLabelValue("시험방법", m.Method);
        AddLabelValue("장비",    m.Instrument);
        AddLabelValue("분석방법", m.AnalysisMethod);
        if (line1.Children.Count > 0) sp.Children.Add(line1);

        if (!string.IsNullOrWhiteSpace(m.Memo))
            sp.Children.Add(new TextBlock
            {
                Text = m.Memo, FontFamily = Font, FontSize = 11,
                Foreground = AppRes("FgDimmed", "#555566"),
                TextWrapping = TextWrapping.Wrap,
            });
        return sp;
    }

    // ── 검정계수 박스 (검정곡선 상단에 a/b/R² 표시) ──
    private static Control BuildCoefBox(string slopeText, string r2Text, string? label = null)
    {
        var sp = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 12,
            Margin = new Thickness(8, 4),
            VerticalAlignment = VerticalAlignment.Center,
        };
        if (!string.IsNullOrWhiteSpace(label))
            sp.Children.Add(new TextBlock
            {
                Text = $"검량계수 ({label})",
                FontFamily = Font, FontSize = 11, FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted", "#888899"),
                VerticalAlignment = VerticalAlignment.Center,
            });
        else
            sp.Children.Add(new TextBlock
            {
                Text = "검량계수",
                FontFamily = Font, FontSize = 11, FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted", "#888899"),
                VerticalAlignment = VerticalAlignment.Center,
            });
        if (!string.IsNullOrWhiteSpace(slopeText))
            sp.Children.Add(new TextBlock
            {
                Text = slopeText, FontFamily = Font, FontSize = 12, FontWeight = FontWeight.Medium,
                Foreground = AppRes("AppFg", "#ffffff"),
                VerticalAlignment = VerticalAlignment.Center,
            });
        if (!string.IsNullOrWhiteSpace(r2Text))
            sp.Children.Add(new TextBlock
            {
                Text = r2Text, FontFamily = Font, FontSize = 12, FontWeight = FontWeight.Medium,
                Foreground = AppRes("AppFg", "#ffffff"),
                VerticalAlignment = VerticalAlignment.Center,
            });

        return new Border
        {
            Child = sp,
            Background = AppRes("GridHeaderBg", "#2a2a3a"),
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
            CornerRadius = new CornerRadius(4, 4, 0, 0),
        };
    }

    // ── 검정곡선 표 (분석결과입력 BuildDocRowUnified 패턴) ──
    private static Control BuildCalibrationTable(Model m)
    {
        // colDefs: 60(구분 4span) + STx(75 each) + 우측 여백(*)
        var colDefs = new ColumnDefinitions();
        colDefs.Add(new ColumnDefinition(15, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(15, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(15, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(60, GridUnitType.Pixel));
        for (int i = 0; i < m.StandardKeys.Count; i++)
            colDefs.Add(new ColumnDefinition(75, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(1, GridUnitType.Star));

        Grid Row(double minH, IBrush bg)
            => new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDefs.ToString()),
                MinHeight = minH, Background = bg,
            };

        // 헤더 (검량계수 컬럼은 분리 — 상단 박스로 이동)
        var hdr = Row(26, AppRes("GridHeaderBg", "#2a2a3a"));
        var hl = HCell("구분", true);
        Grid.SetColumn(hl, 0); Grid.SetColumnSpan(hl, 4);
        hdr.Children.Add(hl);
        for (int i = 0; i < m.StandardKeys.Count; i++)
        {
            var c = HCell($"ST-{m.StandardKeys[i]}", true);
            Grid.SetColumn(c, 4 + i);
            hdr.Children.Add(c);
        }

        var headerBorder = new Border
        {
            Child = hdr,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };

        var stack = new StackPanel { Spacing = 0 };
        // 검량계수 박스 — 검정곡선 표 상단
        stack.Children.Add(BuildCoefBox(m.SlopeText, m.R2Text));
        stack.Children.Add(headerBorder);

        // STANDARD (coefText 빈값 — 인라인 표시 안 함)
        var stdVals = new List<string>();
        for (int i = 0; i < m.StandardKeys.Count; i++)
            stdVals.Add(i < m.StandardConc.Count ? m.StandardConc[i] : "");
        stack.Children.Add(BuildDocRowUnified(colDefs, "STANDARD", stdVals, "", 0));

        // Resp/Area (응답값/흡광도) — ISTD 행 존재 = VOC/GC 형식이라 "Area" 로 라벨
        bool isVocCal = m.StandardIstd.Count > 0
                       && m.StandardIstd.Any(v => !string.IsNullOrWhiteSpace(v));
        var absVals = new List<string>();
        for (int i = 0; i < m.StandardKeys.Count; i++)
            absVals.Add(i < m.StandardAbs.Count ? m.StandardAbs[i] : "");
        stack.Children.Add(BuildDocRowUnified(colDefs, isVocCal ? "Area" : "Resp.", absVals, "", 1));

        // ISTD Resp. (휘발성/VOC 등) — 데이터 있을 때만
        if (isVocCal)
        {
            var istdVals = new List<string>();
            for (int i = 0; i < m.StandardKeys.Count; i++)
                istdVals.Add(i < m.StandardIstd.Count ? m.StandardIstd[i] : "");
            stack.Children.Add(BuildDocRowUnified(colDefs, "ISTD Area", istdVals, "", 2));
        }

        return new Border
        {
            Child = stack,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
        };
    }

    // TOC TCIC 전용 — TC/IC 한 세트의 검정곡선 표시
    private static Control BuildCalibrationTableTcic(Model m, string label,
        List<int> stKeys, List<string> stdConc, List<string> stdAbs, string slopeText, string r2Text)
    {
        var colDefs = new ColumnDefinitions();
        colDefs.Add(new ColumnDefinition(15, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(15, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(15, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(75, GridUnitType.Pixel));
        for (int i = 0; i < stKeys.Count; i++)
            colDefs.Add(new ColumnDefinition(75, GridUnitType.Pixel));
        colDefs.Add(new ColumnDefinition(1, GridUnitType.Star));

        Grid Row(double minH, IBrush bg)
            => new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDefs.ToString()),
                MinHeight = minH, Background = bg,
            };

        var hdr = Row(26, AppRes("GridHeaderBg", "#2a2a3a"));
        var hl = HCell($"구분 ({label})", true);
        Grid.SetColumn(hl, 0); Grid.SetColumnSpan(hl, 4);
        hdr.Children.Add(hl);
        for (int i = 0; i < stKeys.Count; i++)
        {
            var c = HCell($"ST-{stKeys[i]}", true);
            Grid.SetColumn(c, 4 + i);
            hdr.Children.Add(c);
        }

        var headerBorder = new Border
        {
            Child = hdr,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };

        var stack = new StackPanel { Spacing = 0 };
        // 검량계수 박스 — 상단 (TC/IC 라벨 포함)
        stack.Children.Add(BuildCoefBox(slopeText, r2Text, label));
        stack.Children.Add(headerBorder);

        var concVals = new List<string>();
        for (int i = 0; i < stKeys.Count; i++)
            concVals.Add(i < stdConc.Count ? stdConc[i] : "");
        stack.Children.Add(BuildDocRowUnified(colDefs, "STANDARD", concVals, "", 0));

        var absVals = new List<string>();
        for (int i = 0; i < stKeys.Count; i++)
            absVals.Add(i < stdAbs.Count ? stdAbs[i] : "");
        stack.Children.Add(BuildDocRowUnified(colDefs, "Absorbance", absVals, "", 1));

        return new Border
        {
            Child = stack,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 0, 0, 4),
        };
    }

    private static Border BuildDocRowUnified(ColumnDefinitions colDefs, string label,
        List<string> vals, string coefText, int rowIdx)
    {
        var bg = rowIdx % 2 == 0
            ? AppRes("GridRowBg", "#202028")
            : AppRes("GridRowAltBg", "#26262e");

        var g = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(colDefs.ToString()),
            MinHeight = 26, Background = bg,
        };
        var labTb = new TextBlock
        {
            Text = label, FontFamily = Font, FontSize = 12,
            FontWeight = FontWeight.Medium,
            Foreground = AppRes("AppFg", "#ffffff"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 2),
        };
        Grid.SetColumn(labTb, 0); Grid.SetColumnSpan(labTb, 4);
        g.Children.Add(labTb);

        for (int i = 0; i < vals.Count; i++)
        {
            var t = new TextBlock
            {
                Text = string.IsNullOrWhiteSpace(vals[i]) ? "—" : vals[i],
                FontFamily = Font, FontSize = 12,
                Foreground = string.IsNullOrWhiteSpace(vals[i])
                    ? AppRes("FgDimmed", "#555566")
                    : AppRes("AppFg", "#ffffff"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(2),
            };
            Grid.SetColumn(t, 4 + i);
            g.Children.Add(t);
        }

        if (!string.IsNullOrWhiteSpace(coefText))
        {
            var coef = new TextBlock
            {
                Text = coefText, FontFamily = Font, FontSize = 12,
                FontWeight = FontWeight.Medium,
                Foreground = AppRes("AppFg", "#ffffff"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 2),
            };
            Grid.SetColumn(coef, 4 + vals.Count);
            g.Children.Add(coef);
        }

        return new Border
        {
            Child = g,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        };
    }

    private static TextBlock HCell(string text, bool isHeader)
        => new TextBlock
        {
            Text = text, FontFamily = Font,
            FontSize = isHeader ? 11 : 12,
            FontWeight = isHeader ? FontWeight.SemiBold : FontWeight.Regular,
            Foreground = isHeader ? AppRes("FgMuted", "#888899") : AppRes("AppFg", "#ffffff"),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(2),
        };

    // ── 시료 그리드 ──────────────────────────────────────────
    private static readonly System.Collections.Generic.HashSet<string> NumericHeaders =
        new() { "시료량", "흡광도", "AU", "TCAU", "TCcon", "ICAU", "ICcon",
                "A", "B", "ISTD", "Resp", "Resp.",
                "전무게", "후무게", "무게차",
                "희석배수", "원시료희석배수", "농도", "계산농도", "결과", "결과값" };

    private static string ColWidthFor(string h)
    {
        // 계산식만 star (*), 나머지는 고정폭 — 행마다 너비가 어긋나는 것 방지
        return h switch
        {
            "SN"        => "100",
            "시료구분"  => "100",
            "시료명"    => "170",   // 실 데이터 95% 이내 수용, 초과는 ...로 트림
            "계산식"    => "*",
            "시료량"    => "65",
            "흡광도"    => "75",
            "AU"        => "75",
            "TCAU"      => "75",
            "TCcon"     => "75",
            "ICAU"      => "75",
            "ICcon"     => "75",
            // SS 중량법 — 전무게/후무게/무게차 (g 단위, 4자리 소수)
            "전무게"    => "85",
            "후무게"    => "85",
            "무게차"    => "85",
            // 총대장균군(평판집락법) — A/B 평판 colony 수
            "A"         => "55",
            "B"         => "55",
            // VOC/GC 행단위 ISTD 응답
            "ISTD"      => "75",
            "Resp"      => "75",
            "Resp."     => "75",
            "희석배수"  => "70",
            "원시료희석배수" => "90",
            "농도"      => "75",
            "계산농도"  => "75",
            "결과"      => "85",
            "결과값"    => "85",
            // ── BOD 전용 (시료 그리드) ──
            "D1"        => "70",
            "D2"        => "70",
            "F_xy"      => "85",   // 식종액 함유율 (x/y)
            "식종여부"  => "70",
            // ── BOD 식종 정보 영역 (검정곡선 자리) ──
            "구분"      => "150",  // "식종수의 BOD" / "SCF (식종희석수)" / "GGA 1/2/3"
            "시료량(V)" => "75",
            "Result(mg/L)" => "100",
            "비고"      => "*",
            _           => "Auto",
        };
    }

    // ── 식종 정보 영역 (BOD 전용 — 검정곡선 자리) ─────────────
    private static Control BuildSeedSection(Model m)
    {
        var sp = new StackPanel { Spacing = 4, Margin = new Thickness(0, 8, 0, 0) };
        var title = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center,
        };
        title.Children.Add(new Border
        {
            Width = 4, Height = 12,
            Background = AppRes("StatusAccentFg", "#ffaa66"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        title.Children.Add(new TextBlock
        {
            Text = "식종 정보 (식종수의 BOD / 식종희석수)",
            FontFamily = Font, FontSize = 12, FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg", "#ffffff"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(title);

        // 행/클래스 빌드 (뱃지 색상은 시료명 기준으로 자동)
        var classes = new List<string>();
        foreach (var _ in m.SeedRows) classes.Add("");
        sp.Children.Add(BuildRowsBlock(m.SeedHeaders, m.SeedRows, classes));
        return sp;
    }

    // ── 검정곡선의 보증 (MBK/FBK/CCV 등 QC 행) ───────────────────
    private static Control BuildQcSection(Model m)
    {
        var sp = new StackPanel { Spacing = 4, Margin = new Thickness(0, 8, 0, 0) };
        var title = new StackPanel
        {
            Orientation = Orientation.Horizontal, Spacing = 6,
            VerticalAlignment = VerticalAlignment.Center,
        };
        title.Children.Add(new Border
        {
            Width = 4, Height = 12,
            Background = AppRes("StatusAccentFg", "#ffaa66"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        title.Children.Add(new TextBlock
        {
            Text = "검정곡선의 보증",
            FontFamily = Font, FontSize = 12, FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg", "#ffffff"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        sp.Children.Add(title);

        if (m.QcRows.Count == 0)
        {
            sp.Children.Add(new Border
            {
                BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
                BorderThickness = new Thickness(1),
                Background = AppRes("GridRowBg", "#202028"),
                Padding = new Thickness(12, 8),
                Child = new TextBlock
                {
                    Text = "자료없음",
                    FontFamily = Font, FontSize = 12,
                    Foreground = AppRes("FgDimmed", "#555566"),
                    HorizontalAlignment = HorizontalAlignment.Center,
                }
            });
            return sp;
        }

        // 헤더와 데이터 그리드를 시료 그리드와 동일 스타일로 구성
        // 계산식 토글 OFF 면 QC 도 동일하게 헤더/셀 제외
        var qcHeaders = _showFormula
            ? m.SampleHeaders
            : m.SampleHeaders.Where(h => h != "계산식").ToList();
        var qcRows = _showFormula ? m.QcRows : m.QcRows.ConvertAll(r =>
        {
            int idx = m.SampleHeaders.IndexOf("계산식");
            if (idx < 0 || idx >= r.Count) return r;
            var copy = new List<string>(r); copy.RemoveAt(idx); return copy;
        });
        sp.Children.Add(BuildRowsBlock(qcHeaders, qcRows, m.QcClassByRow));
        return sp;
    }

    // 시료 그리드와 QC 그리드 모두에서 사용하는 행렬 빌더
    private static Control BuildRowsBlock(List<string> headers, List<List<string>> rows, List<string> rowClasses)
    {
        var sp = new StackPanel { Spacing = 0 };

        var widths = headers.ConvertAll(ColWidthFor);
        string colDef = string.Join(",", widths);
        if (!widths.Contains("*")) colDef += ",*";

        var hdrGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(colDef),
            MinHeight = 28, Background = AppRes("GridHeaderBg", "#2a2a3a"),
        };
        // VOC/GC: 흡광도 → "Area" 라벨 swap (ISTD 컬럼 존재가 식별 기준)
        bool isVocLabel = headers.Contains("ISTD");
        for (int i = 0; i < headers.Count; i++)
        {
            var h = headers[i];
            string label = (isVocLabel && h == "흡광도") ? "Area" : h;
            HorizontalAlignment ha = HorizontalAlignment.Center;
            if (h == "시료명" || h == "계산식") ha = HorizontalAlignment.Left;
            else if (NumericHeaders.Contains(h))  ha = HorizontalAlignment.Right;
            var t = new TextBlock
            {
                Text = label, FontFamily = Font, FontSize = 11,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted", "#888899"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = ha,
                Margin = new Thickness(8, 3),
            };
            Grid.SetColumn(t, i); hdrGrid.Children.Add(t);
        }
        sp.Children.Add(new Border
        {
            Child = hdrGrid,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        });

        for (int r = 0; r < rows.Count; r++)
        {
            var bg = r % 2 == 0
                ? AppRes("GridRowBg", "#202028")
                : AppRes("GridRowAltBg", "#26262e");
            var rg = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDef),
                MinHeight = 26, Background = bg,
            };
            for (int c = 0; c < headers.Count; c++)
            {
                var hdr = headers[c];
                var val = c < rows[r].Count ? rows[r][c] : "";
                rg.Children.Add(BuildSampleCell(hdr, val, c, r < rowClasses.Count ? rowClasses[r] : ""));
            }
            sp.Children.Add(new Border
            {
                Child = rg,
                BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
                BorderThickness = new Thickness(0, 0, 0, 1),
            });
        }

        return new Border
        {
            Child = sp,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
        };
    }

    private static Control BuildSampleGrid(Model m)
    {
        var sp = new StackPanel { Spacing = 0, Margin = new Thickness(0, 4, 0, 0) };

        // 계산식 토글 OFF 면 해당 헤더 제외해서 BuildSampleGrid 동작
        var headers = _showFormula
            ? m.SampleHeaders
            : m.SampleHeaders.Where(h => h != "계산식").ToList();
        var rows = _showFormula ? m.SampleRows : m.SampleRows.ConvertAll(r =>
        {
            int idx = m.SampleHeaders.IndexOf("계산식");
            if (idx < 0 || idx >= r.Count) return r;
            var copy = new List<string>(r); copy.RemoveAt(idx); return copy;
        });
        // 헤더에 계산식이 이미 *인 경우 추가 스페이서 불필요. 없으면 우측 여백용 * 추가.
        var widths = headers.ConvertAll(ColWidthFor);
        string colDef = string.Join(",", widths);
        if (!widths.Contains("*")) colDef += ",*";

        var hdrGrid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions(colDef),
            MinHeight = 28, Background = AppRes("GridHeaderBg", "#2a2a3a"),
        };
        // VOC/GC: 흡광도 컬럼은 실제로 chromatogram Area 값. 라벨만 "Area" 로 표시
        bool isVocLabel = headers.Contains("ISTD");
        for (int i = 0; i < headers.Count; i++)
        {
            var h = headers[i];
            string label = (isVocLabel && h == "흡광도") ? "Area" : h;
            HorizontalAlignment ha = HorizontalAlignment.Center;
            if (h == "시료명" || h == "계산식") ha = HorizontalAlignment.Left;
            else if (NumericHeaders.Contains(h))  ha = HorizontalAlignment.Right;

            var t = new TextBlock
            {
                Text = label, FontFamily = Font, FontSize = 11,
                FontWeight = FontWeight.SemiBold,
                Foreground = AppRes("FgMuted", "#888899"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = ha,
                Margin = new Thickness(8, 3),
            };
            Grid.SetColumn(t, i); hdrGrid.Children.Add(t);
        }
        sp.Children.Add(new Border
        {
            Child = hdrGrid,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(0, 0, 0, 1),
        });

        for (int r = 0; r < rows.Count; r++)
        {
            var bg = r % 2 == 0
                ? AppRes("GridRowBg", "#202028")
                : AppRes("GridRowAltBg", "#26262e");
            var rg = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions(colDef),
                MinHeight = 26, Background = bg,
            };
            for (int c = 0; c < headers.Count; c++)
            {
                var hdr = headers[c];
                var val = c < rows[r].Count ? rows[r][c] : "";
                rg.Children.Add(BuildSampleCell(hdr, val, c, r < m.SampleClassByRow.Count ? m.SampleClassByRow[r] : ""));
            }
            sp.Children.Add(new Border
            {
                Child = rg,
                BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
                BorderThickness = new Thickness(0, 0, 0, 1),
            });
        }

        return new Border
        {
            Child = sp,
            BorderBrush = AppRes("ThemeBorderSubtle", "#333344"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
        };
    }

    private static Control BuildSampleCell(string header, string val, int col, string sampleClass)
    {
        bool empty = string.IsNullOrWhiteSpace(val);

        Control inner;
        if (header == "SN" && !empty)
        {
            inner = new Border
            {
                Background = AppRes("StatusInfoBg", "#1e2a3e"),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 1),
                Child = new TextBlock
                {
                    Text = val, FontFamily = Font, FontSize = 10,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = AppRes("StatusInfoFg", "#88aadd"),
                },
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
            };
        }
        else if (header == "시료구분" && !empty)
        {
            string bg = "#e3eef9", fg = "#1f4e79";
            if (val.Contains("정도")) { bg = "#f3e8fc"; fg = "#6a3a8c"; }
            else if (val.Contains("4단계")) { bg = "#fff3d6"; fg = "#8a6a1c"; }
            else if (val.Contains("율촌")) { bg = "#e8edf9"; fg = "#2c4a8a"; }
            else if (val.Contains("처리")) { bg = "#e0f2e6"; fg = "#1a7e3a"; }
            inner = new Border
            {
                Background = new SolidColorBrush(Color.Parse(bg)),
                CornerRadius = new CornerRadius(3),
                Padding = new Thickness(5, 1),
                Child = new TextBlock
                {
                    Text = val, FontFamily = Font, FontSize = 10,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = new SolidColorBrush(Color.Parse(fg)),
                },
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
            };
        }
        else if (header == "결과값" || header == "결과")
        {
            IBrush color;
            if (empty) color = FgDim;
            else if (decimal.TryParse(val, out var d) && d < 0) color = ResultRed;
            else color = ResultGreen;
            inner = new TextBlock
            {
                Text = empty ? "—" : val, FontFamily = Font, FontSize = 12,
                FontWeight = FontWeight.SemiBold,
                Foreground = color,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(8, 3),
            };
        }
        else
        {
            HorizontalAlignment ha;
            if (header == "시료명" || header == "계산식") ha = HorizontalAlignment.Left;
            else if (NumericHeaders.Contains(header))      ha = HorizontalAlignment.Right;
            else                                           ha = HorizontalAlignment.Center;

            inner = new TextBlock
            {
                Text = empty ? "—" : val, FontFamily = Font, FontSize = 12,
                Foreground = empty
                    ? AppRes("FgDimmed", "#555566")
                    : AppRes("AppFg", "#ffffff"),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = ha,
                Margin = new Thickness(8, 3),
                // 계산식은 길어질 수 있어 줄바꿈 허용, 다른 컬럼은 ... 으로 트림
                TextWrapping = header == "계산식" ? TextWrapping.Wrap : TextWrapping.NoWrap,
                TextTrimming = header == "계산식"
                    ? TextTrimming.None
                    : TextTrimming.CharacterEllipsis,
            };
        }

        Grid.SetColumn(inner, col);
        return inner;
    }
}
