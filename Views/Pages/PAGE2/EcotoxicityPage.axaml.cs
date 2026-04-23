using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using ClosedXML.Excel;
using ETA.Services.Common;
using ETA.Services.SERVICE2;
using ETA.Views.Controls;
using static ETA.Services.Common.AppFonts;

namespace ETA.Views.Pages.PAGE2;

/// <summary>
/// 생태독성 TSK/Probit 통계분석 전용 페이지
/// DOS용 통계프로그램(TSK/Probit)의 단계별 입력을 재현
/// </summary>
public partial class EcotoxicityPage : UserControl
{
    public event Action<Control?>? StatsPanelChanged;   // Show1
    public event Action<Control?>? ListPanelChanged;   // Show2
    public event Action<Control?>? EditPanelChanged;   // Show3

    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }
    private static T FontBind<T>(T ctrl, string resKey) where T : Control
    {
        var prop = ctrl switch
        {
            TextBlock => TextBlock.FontSizeProperty,
            TextBox   => TextBox.FontSizeProperty,
            _         => Avalonia.Controls.Primitives.TemplatedControl.FontSizeProperty,
        };
        ctrl.Bind(prop, AppFonts.Obs(resKey));
        return ctrl;
    }
    private static T FsSM<T>(T c) where T : Control => FontBind(c, "FontSizeSM");
    private static T FsXS<T>(T c) where T : Control => FontBind(c, "FontSizeXS");
    private static T FsLG<T>(T c) where T : Control => FontBind(c, "FontSizeLG");
    private static T FsBase<T>(T c) where T : Control => FontBind(c, "FontSizeBase");

    private static readonly FontFamily Font =
        new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

    // ── 입력 데이터 ─────────────────────────────────────────────────────────
    private string _testDate = DateTime.Today.ToString("yyyy-MM-dd");
    private string _testNumber = "1";
    private string _endpoint = "EC50";    // EC50 (기본값)
    private string _species = "물벼룩(Daphnia magna)";
    private string _toxicant = "방류수";
    private string _concUnit = "%";
    private string _durUnit = "H";
    private string _duration = "24";
    private int _controlOrganisms = 20;
    private int _controlMortalities = 0;
    private int _numConcentrations = 5;
    private double[] _concentrations = Array.Empty<double>();
    private int[] _organisms = Array.Empty<int>();
    private int[] _mortalities = Array.Empty<int>();
    private bool _equalOrganisms = true;
    private int _organismsPerConc = 20;
    private TextBlock? _selectedTreeNameTb;  // Show1 선택된 시료명 TextBlock
    private string _selectedSn = "";          // Show1 선택된 접수번호(또는 약칭) — DB SN 식별자로 사용
    private string _selectedAlias = "";       // Show1 선택된 약칭

    // ── 추가 시험조건 데이터 (ES 04704.1c 기준) ────────────────────────────
    private double _testTemperature = 20.0;  // (20±2)°C
    private double _testPH = 7.8;            // 7.6~8.0
    private double _sampleTemperature = 20.0;
    private double _samplePH = 7.0;
    private double _sampleDO = 8.0;          // 용존산소 (mg/L)
    private string? _testOrganism;
    private string? _culledOrganism;
    private string _ecCalculationMethod = "Probit";  // Probit 또는 Trimmed Spearman-Karber
    private string _analysisObservations = "";
    private string _analystName = "";

    // QA/QC 참고물질 시험 결과
    private double? _qcReferenceEC50;        // 표준물질(K₂Cr₂O₇) EC50 값

    // ── 결과 ────────────────────────────────────────────────────────────────
    private EcotoxicityService.EcotoxResult? _tskResult;
    private EcotoxicityService.EcotoxResult? _probitResult;

    // ── 시험 이력 ───────────────────────────────────────────────────────────
    private readonly List<TestRecord> _records = new();

    private sealed record TestRecord(
        string Date, string TestNo, string Species, string Toxicant,
        string SampleName, EcotoxicityService.EcotoxResult TskResult,
        EcotoxicityService.EcotoxResult? ProbitResult,
        double[] Conc, int[] Org, int[] Mort, int CtrlOrg, int CtrlMort,
        double TestTemperature, double TestPH, double SampleTemperature, double SamplePH, double SampleDO,
        string Duration, string EcCalculationMethod, string Observations, string AnalystName);

    public EcotoxicityPage()
    {
        InitializeComponent();
    }

    public void LoadData()
    {
        BuildRequestTree();
        ShowInputForm();
        ShowHistoryPanel();
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Show1: 의뢰시료 트리 (생태독성 포함 의뢰만)
    // ══════════════════════════════════════════════════════════════════════════
    private void BuildRequestTree()
    {
        var root = new StackPanel { Spacing = 4, Margin = new Thickness(8) };

        root.Children.Add(FsSM(new TextBlock
        {
            Text = "🐟 생태독성 의뢰시료",
            FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = AppRes("AppFg"),
        }));
        root.Children.Add(FsXS(new TextBlock
        {
            Text = "클릭하면 시료명이 자동 입력됩니다",
            FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 0, 4),
        }));

        List<(string 채취일자, int Id, string 약칭, string 시료명, string 접수번호, string 결과)> records;
        try { records = ETA.Services.SERVICE1.AnalysisRequestService.GetEcotoxRecords(6); }
        catch { records = new(); }

        if (records.Count == 0)
        {
            root.Children.Add(FsSM(new TextBlock
            {
                Text = "생태독성 의뢰가 없습니다.",
                FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 8),
            }));
            StatsPanelChanged?.Invoke(new ScrollViewer { Content = root });
            return;
        }

        // 날짜별 그룹
        var groups = records.GroupBy(r => r.채취일자).OrderByDescending(g => g.Key);

        foreach (var grp in groups)
        {
            // 날짜 헤더
            var dateHeader = new Border
            {
                Background = AppRes("GridRowAltBg"),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(8, 4),
                Margin = new Thickness(0, 4, 0, 0),
            };
            dateHeader.Child = FsSM(new TextBlock
            {
                Text = $"📅 {grp.Key}  ({grp.Count()}건)",
                FontFamily = Font, FontWeight = FontWeight.SemiBold, Foreground = AppRes("ThemeFgInfo"),
            });
            root.Children.Add(dateHeader);

            foreach (var rec in grp)
            {
                var item = new Border
                {
                    Padding = new Thickness(12, 6),
                    Margin = new Thickness(0, 1),
                    CornerRadius = new CornerRadius(3),
                    Cursor = new Cursor(StandardCursorType.Hand),
                };
                TextShimmer.AttachHover(item);

                var sp = new StackPanel { Spacing = 1 };
                var nameTb = FsBase(new TextBlock
                {
                    Text = $"{rec.약칭}  {rec.시료명}",
                    FontFamily = Font, Foreground = AppRes("AppFg"), FontWeight = FontWeight.SemiBold,
                });
                sp.Children.Add(nameTb);

                bool hasResult = !string.IsNullOrEmpty(rec.결과) && rec.결과 != "O";
                sp.Children.Add(FsXS(new TextBlock
                {
                    Text = hasResult ? $"결과: {rec.결과} TU" : "대기 (O)",
                    FontFamily = Font,
                    Foreground = hasResult ? AppRes("ThemeFgSuccess") : AppRes("ThemeFgWarn"),
                }));

                item.Child = sp;

                // 클릭 시 시료명/독성물질/시험일자 자동 입력 + 금색 표시
                var capturedRec = rec;
                var capturedNameTb = nameTb;
                item.PointerPressed += (_, _) =>
                {
                    // 이전 선택 해제
                    if (_selectedTreeNameTb != null)
                        _selectedTreeNameTb.Foreground = AppRes("AppFg");
                    // 현재 선택 금색
                    capturedNameTb.Foreground = new SolidColorBrush(Color.Parse("#FFD700"));
                    _selectedTreeNameTb = capturedNameTb;

                    _toxicant = capturedRec.시료명;
                    _testDate = capturedRec.채취일자;
                    _species = "물벼룩";
                    _selectedSn = !string.IsNullOrWhiteSpace(capturedRec.접수번호)
                        ? capturedRec.접수번호
                        : capturedRec.약칭;
                    _selectedAlias = capturedRec.약칭;
                    ShowInputForm();
                };

                root.Children.Add(item);
            }
        }

        StatsPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Show2: 단계별 입력 폼 (DOS TSK 프로그램 재현)
    // ══════════════════════════════════════════════════════════════════════════
    private void ShowInputForm()
    {
        var root = new StackPanel { Spacing = 12, Margin = new Thickness(14) };

        // ── 헤더 바 ─────────────────────────────────────────────────────────
        var titleBar = new Border
        {
            Background = AppRes("PanelInnerBg"),
            BorderBrush = AppRes("BtnPrimaryBorder"),
            BorderThickness = new Thickness(0, 0, 0, 2),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(16, 12),
        };
        var titleSp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, VerticalAlignment = VerticalAlignment.Center };
        titleSp.Children.Add(new TextBlock { Text = "🐟", FontSize = 26, VerticalAlignment = VerticalAlignment.Center });
        var titleStack = new StackPanel { Spacing = 2, VerticalAlignment = VerticalAlignment.Center };
        titleStack.Children.Add(FsLG(new TextBlock
        {
            Text = "생태독성 통계분석",
            FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = AppRes("AppFg"),
        }));
        titleStack.Children.Add(FsXS(new TextBlock
        {
            Text = "TSK / Probit · ES 04704.1c (수질오염공정시험기준)",
            FontFamily = Font, Foreground = AppRes("FgMuted"),
        }));
        titleSp.Children.Add(titleStack);
        // 선택된 시료 배지 (wire-v01)
        if (!string.IsNullOrWhiteSpace(_selectedSn))
        {
            var badge = StatusBadge.Info($"선택: {_selectedAlias} · {_selectedSn}", withIcon: false);
            badge.Margin = new Thickness(12, 0, 0, 0);
            titleSp.Children.Add(badge);
        }
        titleBar.Child = titleSp;
        root.Children.Add(titleBar);

        // ── 2열 카드 그리드 ─────────────────────────────────────────────────
        var twoCol = new Grid { ColumnSpacing = 12 };
        twoCol.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        twoCol.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));

        var leftCol = new StackPanel { Spacing = 12 };
        var rightCol = new StackPanel { Spacing = 12 };
        Grid.SetColumn(leftCol, 0); Grid.SetColumn(rightCol, 1);
        twoCol.Children.Add(leftCol); twoCol.Children.Add(rightCol);
        root.Children.Add(twoCol);

        // ── 카드 1: 시험 정보 (좌상) ────────────────────────────────────────
        var infoStack = new StackPanel { Spacing = 8 };

        // 상단 read-only 상수 스트립 (시험종/농도단위/시험시간)
        var constStrip = new Border
        {
            Background = AppRes("Panel4Bg"),
            CornerRadius = new CornerRadius(4),
            Padding = new Thickness(10, 6),
        };
        var constSp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 14 };
        constSp.Children.Add(MakeReadOnlyItem("시험종", _species));
        constSp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
        constSp.Children.Add(MakeReadOnlyItem("농도단위", _concUnit));
        constSp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
        constSp.Children.Add(MakeReadOnlyItem("시험시간", $"{_duration} {_durUnit}"));
        constStrip.Child = constSp;
        infoStack.Children.Add(constStrip);

        // 편집 필드: 단일 컬럼 (라벨 좁게, 입력 넓게 stretch)
        var fieldGrid = new Grid { ColumnSpacing = 10, RowSpacing = 8 };
        fieldGrid.ColumnDefinitions.Add(new ColumnDefinition(72, GridUnitType.Pixel));
        fieldGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        for (int i = 0; i < 5; i++) fieldGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

        // Row 0: 시험일자 (CalendarDatePicker, 기본값 = 채수일)
        AddFieldLabel(fieldGrid, 0, "시험일자");
        DateTime.TryParse(_testDate, out var parsedDate);
        var datePicker = new CalendarDatePicker
        {
            SelectedDate = parsedDate == DateTime.MinValue ? DateTime.Today : parsedDate,
            FontFamily = Font,
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        datePicker.Bind(CalendarDatePicker.FontSizeProperty, AppFonts.Obs("FontSizeBase"));
        datePicker.SelectedDateChanged += (_, _) =>
        {
            if (datePicker.SelectedDate.HasValue)
                _testDate = datePicker.SelectedDate.Value.ToString("yyyy-MM-dd");
        };
        Grid.SetRow(datePicker, 0); Grid.SetColumn(datePicker, 1);
        fieldGrid.Children.Add(datePicker);

        // Row 1: 시험번호
        AddFieldLabel(fieldGrid, 1, "시험번호");
        var noBox = MakeInput(_testNumber, v => _testNumber = v);
        noBox.HorizontalAlignment = HorizontalAlignment.Stretch;
        Grid.SetRow(noBox, 1); Grid.SetColumn(noBox, 1);
        fieldGrid.Children.Add(noBox);

        // Row 2: 시료명 (긴 이름 대비, 가장 넓게)
        AddFieldLabel(fieldGrid, 2, "시료명");
        var nameBox = MakeInput(_toxicant, v => _toxicant = v);
        nameBox.HorizontalAlignment = HorizontalAlignment.Stretch;
        nameBox.TextWrapping = TextWrapping.NoWrap;
        Grid.SetRow(nameBox, 2); Grid.SetColumn(nameBox, 1);
        fieldGrid.Children.Add(nameBox);

        // Row 3: 분석유형 (LC50/EC50)
        AddFieldLabel(fieldGrid, 3, "분석유형");
        var endpointPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 16, VerticalAlignment = VerticalAlignment.Center };
        var rbLC = new RadioButton { Content = "LC50", IsChecked = _endpoint == "LC50", FontFamily = Font, Foreground = AppRes("AppFg"), GroupName = "EP" };
        var rbEC = new RadioButton { Content = "EC50", IsChecked = _endpoint == "EC50", FontFamily = Font, Foreground = AppRes("AppFg"), GroupName = "EP" };
        rbLC.IsCheckedChanged += (_, _) => { if (rbLC.IsChecked == true) _endpoint = "LC50"; };
        rbEC.IsCheckedChanged += (_, _) => { if (rbEC.IsChecked == true) _endpoint = "EC50"; };
        endpointPanel.Children.Add(rbLC);
        endpointPanel.Children.Add(rbEC);
        Grid.SetRow(endpointPanel, 3); Grid.SetColumn(endpointPanel, 1);
        fieldGrid.Children.Add(endpointPanel);

        // Row 4: 단위 (시간 단위 H/M)
        AddFieldLabel(fieldGrid, 4, "단위");
        var durBox = MakeInput(_durUnit, v => _durUnit = v);
        durBox.HorizontalAlignment = HorizontalAlignment.Left;
        durBox.MinWidth = 80;
        Grid.SetRow(durBox, 4); Grid.SetColumn(durBox, 1);
        fieldGrid.Children.Add(durBox);

        infoStack.Children.Add(fieldGrid);
        leftCol.Children.Add(MakeCard("📋", "시험 정보", infoStack));

        // ── 카드 2: 대조군 (좌하) ──────────────────────────────────────────
        var ctrlGrid = new Grid { ColumnSpacing = 16, RowSpacing = 4 };
        ctrlGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        ctrlGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        ctrlGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        AddStretchedField(ctrlGrid, 0, 0, "생물수", _controlOrganisms.ToString(),
            v => int.TryParse(v, out _controlOrganisms));
        AddStretchedField(ctrlGrid, 0, 1, "사망수", _controlMortalities.ToString(),
            v => int.TryParse(v, out _controlMortalities));
        leftCol.Children.Add(MakeCard("🧪", "대조군 (Control)", ctrlGrid));

        // ── 카드 3: 농도 설정 (우상) ────────────────────────────────────────
        var concSetStack = new StackPanel { Spacing = 8 };
        var concSetGrid = MakeGrid(1, 2);
        AddLabelInput(concSetGrid, 0, 0, "농도 수 (대조군 제외)", _numConcentrations.ToString(), v =>
        {
            if (int.TryParse(v, out var nc) && nc >= 2 && nc <= 8) _numConcentrations = nc;
        });
        var eqPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        var cbEqual = new CheckBox { IsChecked = _equalOrganisms, FontFamily = Font, Foreground = AppRes("AppFg") };
        cbEqual.IsCheckedChanged += (_, _) => _equalOrganisms = cbEqual.IsChecked == true;
        eqPanel.Children.Add(cbEqual);
        eqPanel.Children.Add(FsSM(new TextBlock { Text = "각 농도 생물수 동일", FontFamily = Font, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center }));
        Grid.SetRow(eqPanel, 0); Grid.SetColumn(eqPanel, 1); concSetGrid.Children.Add(eqPanel);
        concSetStack.Children.Add(concSetGrid);

        var stdBtn = new Button
        {
            Content = "⚡ 표준농도 자동입력 (6.25 / 12.5 / 25 / 50 / 100 %)",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Background = AppRes("BtnPrimaryBg"), Foreground = AppRes("BtnPrimaryFg"),
            BorderBrush = AppRes("BtnPrimaryBorder"), BorderThickness = new Thickness(1),
            Padding = new Thickness(10, 4), CornerRadius = new CornerRadius(6),
            HorizontalContentAlignment = HorizontalAlignment.Center,
        };
        stdBtn.Click += (_, _) =>
        {
            _numConcentrations = 5;
            _concentrations = (double[])EcotoxicityService.StandardConcentrations.Clone();
            _organisms = new int[5];
            _mortalities = new int[5];
            for (int si = 0; si < 5; si++) _organisms[si] = EcotoxicityService.StandardOrganismsPerConc;
            _equalOrganisms = true;
            _organismsPerConc = EcotoxicityService.StandardOrganismsPerConc;
            ShowInputForm();
        };
        rightCol.Children.Add(MakeCard("⚗️", "농도 설정", concSetStack));

        // ── 카드 4: 농도별 데이터 입력 (우하) ───────────────────────────────
        // 기존 값 복원
        if (_concentrations.Length != _numConcentrations)
        {
            var old = _concentrations;
            _concentrations = new double[_numConcentrations];
            Array.Copy(old, _concentrations, Math.Min(old.Length, _numConcentrations));
            var oldO = _organisms;
            _organisms = new int[_numConcentrations];
            Array.Copy(oldO, _organisms, Math.Min(oldO.Length, _numConcentrations));
            var oldM = _mortalities;
            _mortalities = new int[_numConcentrations];
            Array.Copy(oldM, _mortalities, Math.Min(oldM.Length, _numConcentrations));
            if (_organisms.All(x => x == 0))
                for (int i = 0; i < _numConcentrations; i++) _organisms[i] = _organismsPerConc;
        }

        var dataGrid = new Grid { Margin = new Thickness(0) };
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(36, GridUnitType.Pixel));
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));

        // 헤더 행 (배경색 강조)
        dataGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        var headerBg = new Border
        {
            Background = AppRes("GridHeaderBg"),
            CornerRadius = new CornerRadius(4, 4, 0, 0),
        };
        Grid.SetRow(headerBg, 0); Grid.SetColumn(headerBg, 0); Grid.SetColumnSpan(headerBg, 4);
        dataGrid.Children.Add(headerBg);

        var hNum = FsXS(new TextBlock { Text = "#", FontFamily = Font, Foreground = AppRes("FgMuted"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6) });
        var hConc = FsSM(new TextBlock { Text = $"농도 ({_concUnit})", FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6) });
        var hOrg = FsSM(new TextBlock { Text = "생물수", FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6) });
        var hMort = FsSM(new TextBlock { Text = "💀 사망수", FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 6) });
        Grid.SetRow(hNum, 0); Grid.SetColumn(hNum, 0);
        Grid.SetRow(hConc, 0); Grid.SetColumn(hConc, 1);
        Grid.SetRow(hOrg, 0); Grid.SetColumn(hOrg, 2);
        Grid.SetRow(hMort, 0); Grid.SetColumn(hMort, 3);
        dataGrid.Children.Add(hNum); dataGrid.Children.Add(hConc); dataGrid.Children.Add(hOrg); dataGrid.Children.Add(hMort);

        var concInputs = new TextBox[_numConcentrations];
        var orgInputs = new TextBox[_numConcentrations];
        var mortInputs = new TextBox[_numConcentrations];

        for (int ci = 0; ci < _numConcentrations; ci++)
        {
            dataGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
            int row = ci + 1;
            int idx = ci;

            var num = FsXS(new TextBlock { Text = $"{ci + 1}", FontFamily = Font, Foreground = AppRes("FgMuted"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center });
            Grid.SetRow(num, row); Grid.SetColumn(num, 0); dataGrid.Children.Add(num);

            concInputs[ci] = MakeInput(_concentrations[ci] > 0 ? _concentrations[ci].ToString("G") : "", v =>
            { if (double.TryParse(v, out var c)) _concentrations[idx] = c; });
            Grid.SetRow(concInputs[ci], row); Grid.SetColumn(concInputs[ci], 1); dataGrid.Children.Add(concInputs[ci]);

            orgInputs[ci] = MakeInput(_organisms[ci] > 0 ? _organisms[ci].ToString() : "20", v =>
            { if (int.TryParse(v, out var o)) _organisms[idx] = o; });
            orgInputs[ci].IsEnabled = !_equalOrganisms;
            Grid.SetRow(orgInputs[ci], row); Grid.SetColumn(orgInputs[ci], 2); dataGrid.Children.Add(orgInputs[ci]);

            mortInputs[ci] = MakeInput(_mortalities[ci] > 0 ? _mortalities[ci].ToString() : "", v =>
            { if (int.TryParse(v, out var m)) _mortalities[idx] = m; });
            Grid.SetRow(mortInputs[ci], row); Grid.SetColumn(mortInputs[ci], 3); dataGrid.Children.Add(mortInputs[ci]);
        }

        // 사망수 KeyDown: Enter→다음행, W/Up→위, S/Down→아래
        for (int ci = 0; ci < _numConcentrations; ci++)
        {
            int idx2 = ci;
            mortInputs[ci].KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter || e.Key == Key.Down || e.Key == Key.S)
                {
                    e.Handled = true;
                    // 현재값 저장
                    if (int.TryParse(mortInputs[idx2].Text, out var mv)) _mortalities[idx2] = mv;
                    // 다음 행 포커스
                    int next = idx2 + 1;
                    if (next < _numConcentrations)
                    {
                        mortInputs[next].Focus();
                        mortInputs[next].SelectAll();
                    }
                }
                else if (e.Key == Key.Up || e.Key == Key.W)
                {
                    e.Handled = true;
                    if (int.TryParse(mortInputs[idx2].Text, out var mv)) _mortalities[idx2] = mv;
                    int prev = idx2 - 1;
                    if (prev >= 0)
                    {
                        mortInputs[prev].Focus();
                        mortInputs[prev].SelectAll();
                    }
                }
            };
        }

        rightCol.Children.Add(MakeCard("📊", "농도별 데이터 입력", dataGrid, headerRight: stdBtn));

        // ── 결과 카드 (하단 강조) ──────────────────────────────────────────
        var resultTb = FsLG(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });
        var detailTb = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap });
        var compareTb = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0) });
        var resultPlaceholder = FsSM(new TextBlock { Text = "계산 버튼을 눌러 결과를 확인하세요.", FontFamily = Font, Foreground = AppRes("FgMuted"), FontStyle = FontStyle.Italic });

        void CollectInputs()
        {
            for (int ci = 0; ci < _numConcentrations; ci++)
            {
                double.TryParse(concInputs[ci].Text, out var c); _concentrations[ci] = c;
                int.TryParse(orgInputs[ci].Text, out var o); _organisms[ci] = o > 0 ? o : 20;
                int.TryParse(mortInputs[ci].Text, out var m); _mortalities[ci] = m;
            }
            if (_equalOrganisms && _organisms.Length > 0)
            {
                _organismsPerConc = _organisms[0] > 0 ? _organisms[0] : 20;
                for (int i = 0; i < _organisms.Length; i++) _organisms[i] = _organismsPerConc;
            }
        }

        (List<double> vc, List<int> vo, List<int> vm) CollectValid()
        {
            CollectInputs();
            var vc2 = new List<double>(); var vo2 = new List<int>(); var vm2 = new List<int>();
            for (int i = 0; i < _numConcentrations; i++)
                if (_concentrations[i] > 0) { vc2.Add(_concentrations[i]); vo2.Add(_organisms[i]); vm2.Add(_mortalities[i]); }
            return (vc2, vo2, vm2);
        }

        void DoCalcTSK()
        {
            var (vc, vo, vm) = CollectValid();
            if (vc.Count < 2) { resultTb.Text = "최소 2개 유효 농도가 필요합니다."; return; }
            try
            {
                _tskResult = EcotoxicityService.CalculateTSK(vc.ToArray(), vo.ToArray(), vm.ToArray(), _controlOrganisms, _controlMortalities);
                resultTb.Text = $"TSK  {_endpoint}:  {_tskResult.EC50}     TU = {_tskResult.TU}  (100/{_tskResult.EC50})";
                detailTb.Text = $"95% CI: {_tskResult.LowerCI} ~ {_tskResult.UpperCI}  |  Trim: {_tskResult.TrimPercent}%"
                    + (_tskResult.Smoothed ? "  |  단조보정 적용" : "")
                    + (string.IsNullOrEmpty(_tskResult.Warning) ? "" : $"\n⚠ {_tskResult.Warning}");
                compareTb.Text = "";
            }
            catch (Exception ex) { resultTb.Text = $"TSK 오류: {ex.Message}"; }
        }

        void DoCalcProbit()
        {
            var (vc, vo, vm) = CollectValid();
            if (vc.Count < 2) { resultTb.Text = "최소 2개 유효 농도가 필요합니다."; return; }
            try
            {
                _probitResult = EcotoxicityService.CalculateProbit(vc.ToArray(), vo.ToArray(), vm.ToArray(), _controlOrganisms, _controlMortalities);
                resultTb.Text = $"Probit  {_endpoint}:  {_probitResult.EC50}     TU = {_probitResult.TU}  (100/{_probitResult.EC50})";
                detailTb.Text = $"95% CI: {_probitResult.LowerCI} ~ {_probitResult.UpperCI}"
                    + (string.IsNullOrEmpty(_probitResult.Warning) ? "" : $"\n⚠ {_probitResult.Warning}");
                compareTb.Text = "";
            }
            catch (Exception ex) { resultTb.Text = $"Probit 오류: {ex.Message}"; }
        }

        void DoCalcBoth()
        {
            var (vc, vo, vm) = CollectValid();
            if (vc.Count < 2) { resultTb.Text = "최소 2개 유효 농도가 필요합니다."; return; }
            try
            {
                _tskResult = EcotoxicityService.CalculateTSK(vc.ToArray(), vo.ToArray(), vm.ToArray(), _controlOrganisms, _controlMortalities);
                resultTb.Text = $"TSK  {_endpoint}:  {_tskResult.EC50}     TU = {_tskResult.TU}  (100/{_tskResult.EC50})";
                detailTb.Text = $"95% CI: {_tskResult.LowerCI} ~ {_tskResult.UpperCI}  |  Trim: {_tskResult.TrimPercent}%"
                    + (_tskResult.Smoothed ? "  |  단조보정 적용" : "")
                    + (string.IsNullOrEmpty(_tskResult.Warning) ? "" : $"\n⚠ {_tskResult.Warning}");
            }
            catch (Exception ex) { resultTb.Text = $"TSK 오류: {ex.Message}"; }

            try
            {
                _probitResult = EcotoxicityService.CalculateProbit(vc.ToArray(), vo.ToArray(), vm.ToArray(), _controlOrganisms, _controlMortalities);
                compareTb.Text = $"Probit  {_endpoint}:  {_probitResult.EC50}     TU = {_probitResult.TU}  (100/{_probitResult.EC50})  |  95% CI: {_probitResult.LowerCI} ~ {_probitResult.UpperCI}";
            }
            catch (Exception ex) { compareTb.Text = $"Probit 오류: {ex.Message}"; }
        }

        Button MakeActionBtn(string text, BadgeStatus status)
        {
            var (bg, fg, bd) = StatusBadge.GetBrushes(status);
            return new Button
            {
                Content = text, FontFamily = Font,
                Background = bg, Foreground = fg,
                BorderBrush = bd, BorderThickness = new Thickness(1),
                Padding = new Thickness(18, 8), CornerRadius = new CornerRadius(999),
                FontWeight = FontWeight.SemiBold,
                Cursor = new Cursor(StandardCursorType.Hand),
            };
        }

        var tskBtn    = MakeActionBtn("TSK 계산",     BadgeStatus.Info);
        var probitBtn = MakeActionBtn("Probit 계산",  BadgeStatus.Info);
        var bothBtn   = MakeActionBtn("TSK + Probit", BadgeStatus.Accent);
        var saveBtn   = MakeActionBtn("💾 DB 저장",   BadgeStatus.Ok);
        tskBtn.Click    += (_, _) => DoCalcTSK();
        probitBtn.Click += (_, _) => DoCalcProbit();
        bothBtn.Click   += (_, _) => DoCalcBoth();
        saveBtn.Click += (_, _) =>
        {
            Log("───── DB 저장 버튼 클릭 ─────");
            Log($"입력 상태: 시험일자={_testDate}, 시험번호={_testNumber}, 시험종={_species}, 시료명(_toxicant)={_toxicant}, endpoint={_endpoint}, 농도단위={_concUnit}, 시험시간={_duration}{_durUnit}");
            Log($"트리 선택: SN={_selectedSn}, 약칭={_selectedAlias}");
            Log($"대조군: 생물수={_controlOrganisms}, 사망수={_controlMortalities}");
            Log($"농도배열: [{string.Join(",", _concentrations)}], 생물수배열: [{string.Join(",", _organisms)}], 사망수배열: [{string.Join(",", _mortalities)}]");

            if (_tskResult == null)
            {
                Log("_tskResult == null → DoCalcTSK() 호출");
                DoCalcTSK();
            }
            if (_tskResult == null)
            {
                Log("⚠ DoCalcTSK 후에도 _tskResult == null. 저장 중단.");
                detailTb.Text = "⚠ TSK 계산 결과가 없어 저장 중단됨 (농도/사망수를 입력하세요)";
                return;
            }
            CollectInputs();
            Log($"CollectInputs 후 농도배열: [{string.Join(",", _concentrations)}], 사망수배열: [{string.Join(",", _mortalities)}]");
            Log($"TSK 결과: EC50={_tskResult.EC50}, TU={_tskResult.TU}, CI=[{_tskResult.LowerCI},{_tskResult.UpperCI}]");

            try
            {
                int validCnt = _concentrations.Count(c => c > 0);
                var snForDb = !string.IsNullOrWhiteSpace(_selectedSn) ? _selectedSn : _toxicant;
                Log($"UpsertEcotoxData 호출 시작 → SN={snForDb}, 분석일={_testDate}, 시료명={_toxicant}, 시험종={_species}, validCnt={validCnt}");

                bool ok = WasteSampleService.UpsertEcotoxData(
                    _testDate, snForDb, "", "", _toxicant, "생태독성",
                    _species, _duration, _durUnit,
                    _controlOrganisms, _controlMortalities,
                    _concentrations.Where(c => c > 0).ToArray(),
                    _organisms.Take(validCnt).ToArray(),
                    _mortalities.Take(validCnt).ToArray(),
                    result: _tskResult,
                    비고: "",
                    시험번호: _testNumber,
                    endpoint: _endpoint,
                    농도단위: _concUnit,
                    probitResult: _probitResult);

                Log($"UpsertEcotoxData 반환값: {ok}");

                // 수질분석센터_결과.생태독성 컨럼 업데이트 (트리 Show1 새로고침 대비)
                try
                {
                    ETA.Services.SERVICE1.AnalysisRequestService.UpdateEcotoxResult(
                        _testDate, _toxicant, _tskResult.TU.ToString("F1"));
                    Log($"UpdateEcotoxResult 호출 완료 (수질분석센터_결과 갱신)");
                }
                catch (Exception ex) { Log($"UpdateEcotoxResult 실패: {ex.Message}"); }

                // 이력 추가
                _records.Insert(0, new TestRecord(
                    _testDate, _testNumber, _species, _toxicant, _toxicant,
                    _tskResult, _probitResult,
                    _concentrations.Where(c => c > 0).ToArray(),
                    _organisms.Take(validCnt).ToArray(),
                    _mortalities.Take(validCnt).ToArray(),
                    _controlOrganisms, _controlMortalities,
                    _testTemperature, _testPH, _sampleTemperature, _samplePH, _sampleDO,
                    $"{_duration} {_durUnit}", _ecCalculationMethod, _analysisObservations, _analystName));
                ShowHistoryPanel();
                detailTb.Text += ok
                    ? $"  ✅ 시험기록부 저장 완료 (SN={snForDb})"
                    : "  ⚠ 시험기록부 저장 실패 (Logs/EcotoxDebug.log 확인)";

                // Show1 트리 새로고침
                BuildRequestTree();
                Log("───── DB 저장 버튼 처리 완료 ─────");
            }
            catch (Exception ex)
            {
                Log($"❌ 저장 중 예외: {ex.Message}\n{ex.StackTrace}");
                detailTb.Text = $"저장 오류: {ex.Message}";
            }
        };

        var clearBtn = new Button
        {
            Content = "초기화", FontFamily = Font,
            Background = AppRes("BtnDangerBg"), Foreground = AppRes("BtnDangerFg"),
            BorderBrush = AppRes("BtnDangerBorder"), BorderThickness = new Thickness(1),
            Padding = new Thickness(16, 8), CornerRadius = new CornerRadius(6),
        };
        clearBtn.Click += (_, _) =>
        {
            _concentrations = new double[_numConcentrations];
            _organisms = new int[_numConcentrations];
            _mortalities = new int[_numConcentrations];
            for (int i = 0; i < _numConcentrations; i++) _organisms[i] = 20;
            _controlMortalities = 0;
            _tskResult = null; _probitResult = null;
            _testNumber = (_records.Count + 1).ToString();
            ShowInputForm();
        };

        // 액션바: 좌측(분석실행) + 우측(저장/초기화)
        var actionBar = new Border
        {
            Background = AppRes("PanelInnerBg"),
            BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(14, 10),
        };
        var actionGrid = new Grid { ColumnSpacing = 8 };
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        actionGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        var leftActions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        leftActions.Children.Add(FsSM(new TextBlock { Text = "▶ 분석", FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center }));
        leftActions.Children.Add(tskBtn);
        leftActions.Children.Add(probitBtn);
        leftActions.Children.Add(bothBtn);
        Grid.SetColumn(leftActions, 0); actionGrid.Children.Add(leftActions);

        var rightActions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        rightActions.Children.Add(saveBtn);
        rightActions.Children.Add(clearBtn);
        Grid.SetColumn(rightActions, 2); actionGrid.Children.Add(rightActions);

        actionBar.Child = actionGrid;
        root.Children.Add(actionBar);

        // 결과 카드
        var resultStack = new StackPanel { Spacing = 4 };
        resultStack.Children.Add(resultPlaceholder);
        resultStack.Children.Add(resultTb);
        resultStack.Children.Add(detailTb);
        resultStack.Children.Add(compareTb);
        // 결과 텍스트가 채워지면 placeholder 숨김
        resultTb.PropertyChanged += (_, e) =>
        {
            if (e.Property == TextBlock.TextProperty)
                resultPlaceholder.IsVisible = string.IsNullOrEmpty(resultTb.Text);
        };

        var resultCard = new Border
        {
            Background = AppRes("PanelInnerBg"),
            BorderBrush = AppRes("ThemeFgSuccess"),
            BorderThickness = new Thickness(0, 0, 0, 3),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(16, 12),
        };
        var resultRoot = new StackPanel { Spacing = 6 };
        var resultHeader = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        resultHeader.Children.Add(new TextBlock { Text = "✅", FontSize = 18, VerticalAlignment = VerticalAlignment.Center });
        resultHeader.Children.Add(FsBase(new TextBlock { Text = "분석 결과", FontFamily = Font, FontWeight = FontWeight.Bold, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center }));
        resultRoot.Children.Add(resultHeader);
        resultRoot.Children.Add(resultStack);
        resultCard.Child = resultRoot;
        root.Children.Add(resultCard);

        ListPanelChanged?.Invoke(new ScrollViewer { Content = root, Padding = new Thickness(0, 0, 0, 40) });
    }

    // ── 카드 컨테이너 헬퍼 ────────────────────────────────────────────────────
    private Border MakeCard(string icon, string title, Control content, Control? headerRight = null)
    {
        var card = new Border
        {
            Background = AppRes("PanelInnerBg"),
            BorderBrush = AppRes("ThemeBorderSubtle"),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(14, 10, 14, 12),
        };
        var stack = new StackPanel { Spacing = 8 };

        var headerGrid = new Grid();
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));

        var titleRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center };
        titleRow.Children.Add(new TextBlock { Text = icon, FontSize = 16, VerticalAlignment = VerticalAlignment.Center });
        titleRow.Children.Add(FsBase(new TextBlock
        {
            Text = title, FontFamily = Font, FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center,
        }));
        Grid.SetColumn(titleRow, 0);
        headerGrid.Children.Add(titleRow);

        if (headerRight != null)
        {
            headerRight.HorizontalAlignment = HorizontalAlignment.Right;
            headerRight.VerticalAlignment = VerticalAlignment.Center;
            Grid.SetColumn(headerRight, 1);
            headerGrid.Children.Add(headerRight);
        }
        stack.Children.Add(headerGrid);
        stack.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 0, 0, 2) });
        stack.Children.Add(content);

        card.Child = stack;
        return card;
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Show3: 시험 이력
    // ══════════════════════════════════════════════════════════════════════════
    private void ShowHistoryPanel()
    {
        var root = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        root.Children.Add(FsBase(new TextBlock
        {
            Text = "시험 이력", FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = AppRes("AppFg"),
        }));

        if (_records.Count == 0)
        {
            root.Children.Add(FsSM(new TextBlock
            {
                Text = "아직 계산된 시험이 없습니다.", FontFamily = Font, Foreground = AppRes("FgMuted"),
            }));
        }

        foreach (var rec in _records)
        {
            var card = new Border
            {
                Background = AppRes("GridRowBg"),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 8),
                Margin = new Thickness(0, 2),
            };

            var sp = new StackPanel { Spacing = 2 };
            sp.Children.Add(FsBase(new TextBlock
            {
                Text = $"{rec.Date}  #{rec.TestNo}  {rec.Species}  {rec.Toxicant}",
                FontFamily = Font, Foreground = AppRes("AppFg"), FontWeight = FontWeight.SemiBold,
            }));
            sp.Children.Add(FsSM(new TextBlock
            {
                Text = $"TSK: {rec.TskResult.EC50} ({rec.TskResult.LowerCI}~{rec.TskResult.UpperCI})  TU={rec.TskResult.TU}",
                FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"),
            }));
            if (rec.ProbitResult != null)
            {
                sp.Children.Add(FsSM(new TextBlock
                {
                    Text = $"Probit: {rec.ProbitResult.EC50} ({rec.ProbitResult.LowerCI}~{rec.ProbitResult.UpperCI})  TU={rec.ProbitResult.TU}",
                    FontFamily = Font, Foreground = AppRes("ThemeFgInfo"),
                }));
            }
            var rawData = string.Join(", ", rec.Conc.Zip(rec.Mort, (c, m) => $"{c}%:{m}마리"));
            sp.Children.Add(FsXS(new TextBlock
            {
                Text = $"농도별 사망: {rawData}  (대조군 {rec.CtrlOrg}중 {rec.CtrlMort})",
                FontFamily = Font, Foreground = AppRes("FgMuted"),
            }));

            card.Child = sp;
            root.Children.Add(card);
        }

        EditPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  UI 헬퍼
    // ══════════════════════════════════════════════════════════════════════════
    private TextBlock SectionHeader(string text)
    {
        return FsBase(new TextBlock
        {
            Text = text, FontFamily = Font, FontWeight = FontWeight.Bold,
            Foreground = AppRes("ThemeFgWarn"),
            Margin = new Thickness(0, 8, 0, 2),
        });
    }

    private static Grid MakeGrid(int rows, int cols)
    {
        var g = new Grid { Margin = new Thickness(0, 2), ColumnSpacing = 12, RowSpacing = 4 };
        for (int r = 0; r < rows; r++) g.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        for (int c = 0; c < cols; c++) g.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));
        return g;
    }

    private void AddLabelInput(Grid grid, int row, int col, string label, string value, Action<string> onChange)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
        panel.Children.Add(FsXS(new TextBlock { Text = label, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center, MinWidth = 60 }));
        panel.Children.Add(MakeInput(value, onChange));
        Grid.SetRow(panel, row);
        Grid.SetColumn(panel, col);
        grid.Children.Add(panel);
    }

    // 파일 로그 (Logs/EcotoxDebug.log)
    private static void Log(string msg)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [Ecotox] {msg}";
        if (App.EnableLogging)
        {
            try { File.AppendAllText("Logs/EcotoxDebug.log", line + Environment.NewLine); } catch { }
        }
        System.Diagnostics.Debug.WriteLine(line);
    }

    private TextBox MakeInput(string value, Action<string>? onChange = null)
    {
        var tb = FsBase(new TextBox
        {
            Text = value, FontFamily = Font,
            Foreground = AppRes("InputFg"), Background = AppRes("InputBg"),
            BorderBrush = AppRes("InputBorder"), BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4), Padding = new Thickness(6, 3),
            MinWidth = 20,
        });
        if (onChange != null)
            tb.LostFocus += (_, _) => onChange(tb.Text ?? "");
        return tb;
    }

    // 읽기 전용 상수 표시 ("라벨: 값" 형태)
    private Control MakeReadOnlyItem(string label, string value)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6, VerticalAlignment = VerticalAlignment.Center };
        sp.Children.Add(FsXS(new TextBlock { Text = label, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center }));
        sp.Children.Add(FsSM(new TextBlock { Text = value, FontFamily = Font, Foreground = AppRes("AppFg"), FontWeight = FontWeight.SemiBold, VerticalAlignment = VerticalAlignment.Center }));
        return sp;
    }

    // 좌측 라벨만 추가 (입력은 별도)
    private void AddFieldLabel(Grid grid, int row, string label)
    {
        var tb = FsXS(new TextBlock
        {
            Text = label, FontFamily = Font, Foreground = AppRes("FgMuted"),
            VerticalAlignment = VerticalAlignment.Center,
        });
        Grid.SetRow(tb, row); Grid.SetColumn(tb, 0);
        grid.Children.Add(tb);
    }

    // 라벨 + 입력창 (입력창이 셀 가득 stretch, MinWidth 보장)
    private void AddStretchedField(Grid grid, int row, int col, string label, string value, Action<string> onChange, double inputMinWidth = 110)
    {
        var inner = new Grid { ColumnSpacing = 8 };
        inner.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        inner.ColumnDefinitions.Add(new ColumnDefinition(1, GridUnitType.Star));

        var lbl = FsXS(new TextBlock { Text = label, FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center });
        Grid.SetColumn(lbl, 0); inner.Children.Add(lbl);

        var input = MakeInput(value, onChange);
        input.HorizontalAlignment = HorizontalAlignment.Stretch;
        input.MinWidth = inputMinWidth;
        Grid.SetColumn(input, 1); inner.Children.Add(input);

        Grid.SetRow(inner, row); Grid.SetColumn(inner, col);
        grid.Children.Add(inner);
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  시험기록부 출력 (Excel)
    // ══════════════════════════════════════════════════════════════════════════
    public async Task ExportTestReportAsync()
    {
        // 저장된 기록이 없으면 현재 입력된 데이터로 임시 기록 생성
        if (!_records.Any())
        {
            if (_tskResult == null && _probitResult == null)
            {
                System.Diagnostics.Debug.WriteLine("[시험기록부 출력] 저장된 시험 기록이 없습니다. 먼저 계산을 수행하세요.");
                return;
            }

            // 현재 입력된 데이터를 임시 기록으로 생성
            var validCnt = _concentrations.Count(c => c > 0);
            _records.Insert(0, new TestRecord(
                _testDate, _testNumber, _species, _toxicant, _selectedTreeNameTb?.Text ?? "미지정",
                _tskResult ?? new EcotoxicityService.EcotoxResult(0, 0, 0, 0, "TSK", -1, false, "계산 결과 없음"),
                _probitResult,
                _concentrations.Where(c => c > 0).ToArray(),
                _organisms.Take(validCnt).ToArray(),
                _mortalities.Take(validCnt).ToArray(),
                _controlOrganisms, _controlMortalities,
                _testTemperature, _testPH, _sampleTemperature, _samplePH, _sampleDO,
                $"{_duration} {_durUnit}", _ecCalculationMethod, _analysisObservations, _analystName));
        }

        try
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("시험기록부");

            ws.Columns().Width = 14;
            int row = 1;

            // ── 제목 ────────────────────────────────────────────────────────────
            var titleCell = ws.Cell(row, 1);
            titleCell.Value = "물벼룩을 이용한 급성 독성 시험 기록부";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 14;
            ws.Range(row, 1, row, 8).Merge();
            ws.Row(row).Height = 24;
            row++;

            var stdCell = ws.Cell(row, 1);
            stdCell.Value = "ES 04704.1c (수질오염공정시험기준, 2023)";
            stdCell.Style.Font.Italic = true;
            ws.Range(row, 1, row, 8).Merge();
            row += 2;

            // ── 각 시험 기록 ─────────────────────────────────────────────────────
            foreach (var rec in _records)
            {
                // ─ 시험 기본 정보 ─
                ws.Cell(row, 1).Value = "시험기본정보";
                ws.Cell(row, 1).Style.Font.Bold = true;
                row++;

                ws.Cell(row, 1).Value = "시험일자";
                ws.Cell(row, 2).Value = rec.Date;
                ws.Cell(row, 3).Value = "시험번호";
                ws.Cell(row, 4).Value = rec.TestNo;
                ws.Cell(row, 5).Value = "시험생물";
                ws.Cell(row, 6).Value = "Daphnia magna Straus";
                ws.Cell(row, 7).Value = "시험기간";
                ws.Cell(row, 8).Value = rec.Duration;
                row++;

                ws.Cell(row, 1).Value = "시료명";
                ws.Cell(row, 2).Value = rec.SampleName;
                ws.Cell(row, 3).Value = "오염물질";
                ws.Cell(row, 4).Value = rec.Toxicant;
                row++;

                // ─ 시험 조건 (ES 04704.1c 기준) ─
                row++;
                ws.Cell(row, 1).Value = "시험조건";
                ws.Cell(row, 1).Style.Font.Bold = true;
                row++;

                ws.Cell(row, 1).Value = "시험온도(°C)";
                ws.Cell(row, 2).Value = rec.TestTemperature;
                ws.Cell(row, 3).Value = "(기준: 20±2)";
                ws.Cell(row, 4).Value = "시험 pH";
                ws.Cell(row, 5).Value = rec.TestPH;
                ws.Cell(row, 6).Value = "(기준: 7.6~8.0)";
                row++;

                ws.Cell(row, 1).Value = "시료온도(°C)";
                ws.Cell(row, 2).Value = rec.SampleTemperature;
                ws.Cell(row, 3).Value = "시료 pH";
                ws.Cell(row, 4).Value = rec.SamplePH;
                ws.Cell(row, 5).Value = "용존산소(mg/L)";
                ws.Cell(row, 6).Value = rec.SampleDO;
                row++;

                ws.Cell(row, 1).Value = "EC50 계산방법";
                ws.Cell(row, 2).Value = rec.EcCalculationMethod;
                row++;

                // ─ 대조군 정보 ─
                row++;
                ws.Cell(row, 1).Value = "대조군 정보";
                ws.Cell(row, 1).Style.Font.Bold = true;
                row++;

                ws.Cell(row, 1).Value = "대조생물수";
                ws.Cell(row, 2).Value = rec.CtrlOrg;
                ws.Cell(row, 3).Value = "대조 치사수";
                ws.Cell(row, 4).Value = rec.CtrlMort;
                var ctrlMortalityRate = rec.CtrlOrg > 0 ? (rec.CtrlMort * 100.0 / rec.CtrlOrg) : 0;
                ws.Cell(row, 5).Value = "치사율(%)";
                ws.Cell(row, 6).Value = ctrlMortalityRate.ToString("F1");
                if (ctrlMortalityRate > 15) ws.Cell(row, 6).Style.Fill.BackgroundColor = XLColor.LightCoral;
                row++;

                // ─ 농도별 데이터 ─
                row++;
                ws.Cell(row, 1).Value = "농도별 독성 시험 데이터";
                ws.Cell(row, 1).Style.Font.Bold = true;
                row++;

                const int dataColStart = 1;
                ws.Cell(row, dataColStart).Value = "농도";
                ws.Cell(row, dataColStart + 1).Value = "단위";
                ws.Cell(row, dataColStart + 2).Value = "생물수";
                ws.Cell(row, dataColStart + 3).Value = "사망수";
                ws.Cell(row, dataColStart + 4).Value = "치사율(%)";
                for (int c = dataColStart; c <= dataColStart + 4; c++)
                    ws.Cell(row, c).Style.Font.Bold = true;
                row++;

                for (int i = 0; i < rec.Conc.Length; i++)
                {
                    var conc = rec.Conc[i];
                    var org = i < rec.Org.Length ? rec.Org[i] : 0;
                    var mort = i < rec.Mort.Length ? rec.Mort[i] : 0;
                    var mortRate = org > 0 ? (mort * 100.0 / org) : 0;

                    ws.Cell(row, dataColStart).Value = conc;
                    ws.Cell(row, dataColStart + 1).Value = "%";
                    ws.Cell(row, dataColStart + 2).Value = org;
                    ws.Cell(row, dataColStart + 3).Value = mort;
                    ws.Cell(row, dataColStart + 4).Value = mortRate.ToString("F1");
                    row++;
                }
                row++;

                // ─ 분석 결과 ─
                ws.Cell(row, 1).Value = "독성 분석 결과";
                ws.Cell(row, 1).Style.Font.Bold = true;
                row++;

                if (rec.TskResult != null)
                {
                    var tsk = rec.TskResult;
                    ws.Cell(row, 1).Value = "TSK 분석";
                    ws.Cell(row, 1).Style.Font.Bold = true;
                    row++;

                    ws.Cell(row, 1).Value = "EC50";
                    ws.Cell(row, 2).Value = tsk.EC50;
                    ws.Cell(row, 3).Value = "TU (Toxic Units)";
                    ws.Cell(row, 4).Value = tsk.TU;
                    row++;

                    ws.Cell(row, 1).Value = "95% 신뢰구간 (하한)";
                    ws.Cell(row, 2).Value = tsk.LowerCI;
                    ws.Cell(row, 3).Value = "95% 신뢰구간 (상한)";
                    ws.Cell(row, 4).Value = tsk.UpperCI;
                    row++;

                    if (tsk.TrimPercent >= 0)
                    {
                        ws.Cell(row, 1).Value = "Trim %";
                        ws.Cell(row, 2).Value = tsk.TrimPercent;
                        row++;
                    }

                    if (!string.IsNullOrEmpty(tsk.Warning))
                    {
                        ws.Cell(row, 1).Value = "⚠ 주의사항";
                        ws.Cell(row, 2).Value = tsk.Warning;
                        ws.Row(row).Style.Fill.BackgroundColor = XLColor.LightYellow;
                        row++;
                    }
                    row++;
                }

                if (rec.ProbitResult != null)
                {
                    var probit = rec.ProbitResult;
                    ws.Cell(row, 1).Value = "Probit 분석";
                    ws.Cell(row, 1).Style.Font.Bold = true;
                    row++;

                    ws.Cell(row, 1).Value = "EC50";
                    ws.Cell(row, 2).Value = probit.EC50;
                    ws.Cell(row, 3).Value = "TU (Toxic Units)";
                    ws.Cell(row, 4).Value = probit.TU;
                    row++;

                    ws.Cell(row, 1).Value = "95% 신뢰구간 (하한)";
                    ws.Cell(row, 2).Value = probit.LowerCI;
                    ws.Cell(row, 3).Value = "95% 신뢰구간 (상한)";
                    ws.Cell(row, 4).Value = probit.UpperCI;
                    row++;

                    if (!string.IsNullOrEmpty(probit.Warning))
                    {
                        ws.Cell(row, 1).Value = "⚠ 주의사항";
                        ws.Cell(row, 2).Value = probit.Warning;
                        ws.Row(row).Style.Fill.BackgroundColor = XLColor.LightYellow;
                        row++;
                    }
                    row++;
                }

                // ─ 독성 분류 ─
                if (rec.TskResult != null)
                {
                    var tu = rec.TskResult.TU;
                    string classification = tu > 16 ? "높은 독성" : tu > 8 ? "중간 독성" : tu > 1 ? "낮은 독성" : "매우 낮은 독성";

                    ws.Cell(row, 1).Value = "독성 분류 (TU 기준)";
                    ws.Cell(row, 2).Value = classification;
                    if (tu > 16) ws.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightCoral;
                    else if (tu > 8) ws.Cell(row, 2).Style.Fill.BackgroundColor = XLColor.LightYellow;
                    row += 2;
                }

                // ─ 비고 ─
                if (!string.IsNullOrEmpty(rec.Observations))
                {
                    ws.Cell(row, 1).Value = "시험 중 관찰사항";
                    ws.Cell(row, 2).Value = rec.Observations;
                    row++;
                }

                if (!string.IsNullOrEmpty(rec.AnalystName))
                {
                    ws.Cell(row, 1).Value = "분석자";
                    ws.Cell(row, 2).Value = rec.AnalystName;
                    row++;
                }

                // ─ 용량-반응 곡선 PNG 삽입 ─
                if (rec.Conc != null && rec.Conc.Length >= 1 && rec.TskResult != null)
                {
                    try
                    {
                        var chartBytes = ETA.Services.SERVICE2.EcotoxicityChartGenerator.Generate(
                            rec.Conc, rec.Org, rec.Mort,
                            rec.TskResult.EC50, rec.TskResult.LowerCI, rec.TskResult.UpperCI,
                            rec.TskResult.TU, rec.TskResult.Method);

                        using var chartStream = new System.IO.MemoryStream(chartBytes);
                        var pic = ws.AddPicture(chartStream)
                                    .MoveTo(ws.Cell(row + 1, 1))
                                    .WithSize(640, 400);
                        row += 22;  // 차트 공간 확보 (행 높이 고려)
                    }
                    catch (Exception chartEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[차트 삽입 실패] {chartEx.Message}");
                    }
                }

                row += 3;
            }

            // 저장
            string filename = $"생태독성_시험기록부_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            string savePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                filename);

            wb.SaveAs(savePath);
            System.Diagnostics.Debug.WriteLine($"시험기록부 저장: {savePath}");
            await Task.CompletedTask;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"시험기록부 출력 오류: {ex.Message}");
        }
    }
}
