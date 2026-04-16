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
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(16) };

        root.Children.Add(FsLG(new TextBlock
        {
            Text = "🐟 생태독성 통계분석 (TSK / Probit)",
            FontWeight = FontWeight.Bold, FontFamily = Font, Foreground = AppRes("AppFg"),
        }));

        // ── 1. 시험 정보 ────────────────────────────────────────────────────
        root.Children.Add(SectionHeader("1. 시험 정보"));

        var infoGrid = MakeGrid(4, 2);
        AddLabelInput(infoGrid, 0, 0, "시험일자", _testDate, v => _testDate = v);
        AddLabelInput(infoGrid, 0, 1, "시험번호", _testNumber, v => _testNumber = v);
        AddLabelInput(infoGrid, 1, 0, "시험종", _species, v => _species = v);
        AddLabelInput(infoGrid, 1, 1, "독성물질", _toxicant, v => _toxicant = v);
        AddLabelInput(infoGrid, 2, 0, "농도단위", _concUnit, v => _concUnit = v);

        // LC50/EC50 선택
        var endpointPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, VerticalAlignment = VerticalAlignment.Center };
        var rbLC = new RadioButton { Content = "LC50", IsChecked = _endpoint == "LC50", FontFamily = Font, Foreground = AppRes("AppFg"), GroupName = "EP" };
        var rbEC = new RadioButton { Content = "EC50", IsChecked = _endpoint == "EC50", FontFamily = Font, Foreground = AppRes("AppFg"), GroupName = "EP" };
        rbLC.IsCheckedChanged += (_, _) => { if (rbLC.IsChecked == true) _endpoint = "LC50"; };
        rbEC.IsCheckedChanged += (_, _) => { if (rbEC.IsChecked == true) _endpoint = "EC50"; };
        endpointPanel.Children.Add(rbLC);
        endpointPanel.Children.Add(rbEC);
        var epLabel = FsXS(new TextBlock { Text = "분석유형", FontFamily = Font, Foreground = AppRes("FgMuted"), Margin = new Thickness(0, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center });
        var epRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
        epRow.Children.Add(epLabel);
        epRow.Children.Add(endpointPanel);
        Grid.SetRow(epRow, 2); Grid.SetColumn(epRow, 1); infoGrid.Children.Add(epRow);

        AddLabelInput(infoGrid, 3, 0, "시험시간", _duration, v => _duration = v);
        var durUnitPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
        durUnitPanel.Children.Add(FsXS(new TextBlock { Text = "단위", FontFamily = Font, Foreground = AppRes("FgMuted"), VerticalAlignment = VerticalAlignment.Center }));
        var durBox = MakeInput(_durUnit, v => _durUnit = v);
        durBox.Width = 40;
        durUnitPanel.Children.Add(durBox);
        Grid.SetRow(durUnitPanel, 3); Grid.SetColumn(durUnitPanel, 1); infoGrid.Children.Add(durUnitPanel);
        root.Children.Add(infoGrid);

        // ── 2. 대조군 ──────────────────────────────────────────────────────
        root.Children.Add(SectionHeader("2. 대조군 (Control)"));
        var ctrlGrid = MakeGrid(1, 2);
        AddLabelInput(ctrlGrid, 0, 0, "생물수", _controlOrganisms.ToString(), v => int.TryParse(v, out _controlOrganisms));
        AddLabelInput(ctrlGrid, 0, 1, "사망수", _controlMortalities.ToString(), v => int.TryParse(v, out _controlMortalities));
        root.Children.Add(ctrlGrid);

        // ── 3. 농도 수 ─────────────────────────────────────────────────────
        root.Children.Add(SectionHeader("3. 농도 설정"));
        var concSetGrid = MakeGrid(1, 2);
        AddLabelInput(concSetGrid, 0, 0, "농도 수 (대조군 제외)", _numConcentrations.ToString(), v =>
        {
            if (int.TryParse(v, out var nc) && nc >= 2 && nc <= 8) _numConcentrations = nc;
        });
        var eqPanel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, VerticalAlignment = VerticalAlignment.Center };
        var cbEqual = new CheckBox { IsChecked = _equalOrganisms, FontFamily = Font, Foreground = AppRes("AppFg") };
        cbEqual.IsCheckedChanged += (_, _) => _equalOrganisms = cbEqual.IsChecked == true;
        eqPanel.Children.Add(cbEqual);
        eqPanel.Children.Add(FsSM(new TextBlock { Text = "각 농도 생물수 동일", FontFamily = Font, Foreground = AppRes("AppFg"), VerticalAlignment = VerticalAlignment.Center }));
        Grid.SetRow(eqPanel, 0); Grid.SetColumn(eqPanel, 1); concSetGrid.Children.Add(eqPanel);
        // 표준농도 입력 버튼 (ES 04704.1c: 6.25, 12.5, 25, 50, 100%)
        var stdBtn = new Button
        {
            Content = "표준농도 입력 (6.25/12.5/25/50/100%)", FontFamily = Font,
            Background = AppRes("ThemeFgInfo"), Foreground = Brushes.White,
            Padding = new Thickness(12, 4), CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 4),
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
            ShowInputForm(); // UI 새로고침
        };
        Grid.SetRow(stdBtn, 0); Grid.SetColumn(stdBtn, 1);
        // concSetGrid의 기존 eqPanel 대신 stdBtn + eqPanel을 세로로
        var concRightPanel = new StackPanel { Spacing = 4 };
        // eqPanel을 concSetGrid에서 제거 후 concRightPanel에 추가
        concSetGrid.Children.Remove(eqPanel);
        concRightPanel.Children.Add(stdBtn);
        concRightPanel.Children.Add(eqPanel);
        Grid.SetRow(concRightPanel, 0); Grid.SetColumn(concRightPanel, 1);
        concSetGrid.Children.Add(concRightPanel);

        root.Children.Add(concSetGrid);

        // ── 4. 농도별 데이터 입력 ─────────────────────────────────────────
        root.Children.Add(SectionHeader("4. 농도별 데이터 입력"));

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

        var dataGrid = new Grid { Margin = new Thickness(0, 4) };
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(30, GridUnitType.Pixel));
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(70, GridUnitType.Pixel));
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(70, GridUnitType.Pixel));
        dataGrid.ColumnDefinitions.Add(new ColumnDefinition(70, GridUnitType.Pixel));

        // 헤더
        dataGrid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        var hNum = FsXS(new TextBlock { Text = "#", FontFamily = Font, Foreground = AppRes("FgMuted"), HorizontalAlignment = HorizontalAlignment.Center });
        var hConc = FsSM(new TextBlock { Text = $"농도({_concUnit})", FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), FontWeight = FontWeight.SemiBold });
        var hOrg = FsSM(new TextBlock { Text = "생물수", FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center });
        var hMort = FsSM(new TextBlock { Text = "💀 사망수", FontFamily = Font, Foreground = AppRes("ThemeFgWarn"), FontWeight = FontWeight.SemiBold, HorizontalAlignment = HorizontalAlignment.Center });
        Grid.SetColumn(hNum, 0); Grid.SetColumn(hConc, 1); Grid.SetColumn(hOrg, 2); Grid.SetColumn(hMort, 3);
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

        root.Children.Add(dataGrid);

        // ── 5. 계산 버튼 ───────────────────────────────────────────────────
        root.Children.Add(new Border { Height = 1, Background = AppRes("ThemeBorderSubtle"), Margin = new Thickness(0, 8) });

        var resultTb = FsLG(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgSuccess"), FontWeight = FontWeight.Bold, TextWrapping = TextWrapping.Wrap });
        var detailTb = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("FgMuted"), TextWrapping = TextWrapping.Wrap });
        var compareTb = FsSM(new TextBlock { FontFamily = Font, Foreground = AppRes("ThemeFgInfo"), TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 4, 0, 0) });

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

        var btnPanel = new WrapPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 4) };

        var tskBtn = new Button
        {
            Content = "TSK 계산", FontFamily = Font,
            Background = AppRes("BtnPrimaryBg"), Foreground = Brushes.White,
            Padding = new Thickness(16, 8), CornerRadius = new CornerRadius(4),
            FontWeight = FontWeight.Bold, Margin = new Thickness(0, 0, 8, 4),
        };
        tskBtn.Click += (_, _) => DoCalcTSK();

        var probitBtn = new Button
        {
            Content = "Probit 계산", FontFamily = Font,
            Background = AppRes("ThemeFgInfo"), Foreground = Brushes.White,
            Padding = new Thickness(16, 8), CornerRadius = new CornerRadius(4),
            FontWeight = FontWeight.Bold, Margin = new Thickness(0, 0, 8, 4),
        };
        probitBtn.Click += (_, _) => DoCalcProbit();

        var bothBtn = new Button
        {
            Content = "TSK + Probit", FontFamily = Font,
            Background = new SolidColorBrush(Color.Parse("#2E7D32")), Foreground = Brushes.White,
            Padding = new Thickness(16, 8), CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 0, 8, 4),
        };
        bothBtn.Click += (_, _) => DoCalcBoth();

        var saveBtn = new Button
        {
            Content = "DB 저장", FontFamily = Font,
            Background = AppRes("ThemeFgSuccess"), Foreground = Brushes.White,
            Padding = new Thickness(16, 8), CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 0, 8, 4),
        };
        saveBtn.Click += (_, _) =>
        {
            if (_tskResult == null) { DoCalcTSK(); }
            if (_tskResult == null) return;
            CollectInputs();
            try
            {
                int validCnt = _concentrations.Count(c => c > 0);
                WasteSampleService.UpsertEcotoxData(
                    _testDate, _toxicant, "", "", _species, "생태독성",
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

                // 수질분석센터_결과.생태독성 컨럼 업데이트 (트리 Show1 새로고침 대비)
                try
                {
                    ETA.Services.SERVICE1.AnalysisRequestService.UpdateEcotoxResult(
                        _testDate, _toxicant, _tskResult.TU.ToString("F1"));
                }
                catch { }

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
                detailTb.Text += "  ✅ DB 저장 완료";

                // Show1 트리 새로고침
                BuildRequestTree();
            }
            catch (Exception ex) { detailTb.Text = $"저장 오류: {ex.Message}"; }
        };

        var clearBtn = new Button
        {
            Content = "초기화", FontFamily = Font,
            Background = Brushes.Transparent, Foreground = AppRes("ThemeFgDanger"),
            BorderBrush = AppRes("ThemeFgDanger"), BorderThickness = new Thickness(1),
            Padding = new Thickness(16, 8), CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 0, 0, 4),
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

        btnPanel.Children.Add(tskBtn);
        btnPanel.Children.Add(probitBtn);
        btnPanel.Children.Add(bothBtn);
        btnPanel.Children.Add(saveBtn);
        btnPanel.Children.Add(clearBtn);
        root.Children.Add(btnPanel);
        root.Children.Add(resultTb);
        root.Children.Add(detailTb);
        root.Children.Add(compareTb);

        ListPanelChanged?.Invoke(new ScrollViewer { Content = root, Padding = new Thickness(0, 0, 0, 40) });
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
