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
    public event Action<Control?>? RecordsPanelChanged; // Show4 — 시험기록부 저장 항목 리스트

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
    private int    _selectedRecId = -1;       // Show1 선택된 수질분석센터_결과 rowid — 현장측정 조회용

    // 생태독성 담당자 입력 3개 (염분/암모니아/경도) TextBox 참조
    private TextBox? _tbSalinity, _tbAmmonia, _tbHardness;

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

    // 마지막 시험기록부 출력 상태 — Show4 헤더에 표시 (성공/실패 사유 가시화).
    private string _lastExportStatus = "";
    private bool   _lastExportOk     = true;

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
        RefreshRecordsPanel();
    }

    /// <summary>생태독성 전용 출력 폴더 — 시험기록부/법정양식 docx 모두 여기에 저장</summary>
    private static string EcotoxOutputDir =>
        System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "ETA 생태독성 시험기록부");

    /// <summary>Show4 — 생태독성 출력 폴더의 docx 파일 리스트. 클릭 시 Word 로 열림.</summary>
    public void RefreshRecordsPanel()
    {
        var root = new StackPanel { Spacing = 4, Margin = new Thickness(8) };

        // 헤더 행 — 좌측 라벨 + 우측 "저장 폴더 열기" 버튼
        var headerRow = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("*,Auto"),
            Margin = new Thickness(0, 0, 0, 6),
        };
        var headerLabel = new TextBlock
        {
            Text = "📋 생태독성 시험기록부",
            FontSize = AppTheme.FontMD, FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgInfo,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(headerLabel, 0);
        headerRow.Children.Add(headerLabel);

        var openFolderBtn = new Button
        {
            Content = "📂 폴더 열기",
            FontSize = AppTheme.FontXS,
            Padding = new Thickness(8, 3),
            VerticalAlignment = VerticalAlignment.Center,
        };
        openFolderBtn.Click += (_, _) => OpenSavedRecordsFolder();
        Grid.SetColumn(openFolderBtn, 1);
        headerRow.Children.Add(openFolderBtn);
        root.Children.Add(headerRow);

        // 폴더 파일 스캔 (mtime 내림차순)
        var files = new List<System.IO.FileInfo>();
        try
        {
            System.IO.Directory.CreateDirectory(EcotoxOutputDir);
            files = new System.IO.DirectoryInfo(EcotoxOutputDir)
                .GetFiles("*.docx")
                .Where(fi => !fi.Name.StartsWith("~$"))   // Word 임시파일 제외
                .OrderByDescending(fi => fi.LastWriteTime)
                .ToList();
        }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[Ecotox.Show4] {ex.Message}"); }

        if (files.Count == 0)
        {
            root.Children.Add(new TextBlock
            {
                Text = "출력된 시험기록부 / 법정양식이 없습니다.",
                FontSize = AppTheme.FontSM, Foreground = AppTheme.FgDimmed,
            });
        }
        else
        {
            foreach (var fi in files)
            {
                bool isLegal = fi.Name.Contains("법정양식", StringComparison.OrdinalIgnoreCase);
                string icon  = isLegal ? "📜" : "🧪";

                var stack = new StackPanel { Spacing = 2 };
                stack.Children.Add(new TextBlock
                {
                    Text = $"{icon}  {System.IO.Path.GetFileNameWithoutExtension(fi.Name)}",
                    FontSize = AppTheme.FontSM, FontWeight = FontWeight.Medium,
                    TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
                });
                stack.Children.Add(new TextBlock
                {
                    Text = $"   {fi.LastWriteTime:yyyy-MM-dd HH:mm}    ({fi.Length / 1024.0:N1} KB)",
                    FontSize = AppTheme.FontXS, Foreground = AppTheme.FgDimmed,
                });

                var border = new Border
                {
                    Padding = new Thickness(8, 6),
                    Margin = new Thickness(0, 0, 0, 4),
                    Background = AppRes("PanelBg", "#1f1f28"),
                    CornerRadius = new CornerRadius(4),
                    Cursor = new Cursor(StandardCursorType.Hand),
                    Child = stack,
                };
                string capturedPath = fi.FullName;
                border.PointerEntered += (_, _) => border.Background = AppRes("BgHover", "#2a2a35");
                border.PointerExited  += (_, _) => border.Background = AppRes("PanelBg", "#1f1f28");
                border.PointerPressed += (_, _) =>
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = capturedPath,
                            UseShellExecute = true,
                        });
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[Ecotox.OpenFile] {ex.Message}"); }
                };
                root.Children.Add(border);
            }
        }

        RecordsPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    /// <summary>DB 행 → 폼 입력 필드 복원</summary>
    /// <summary>Windows/Unix 양쪽 금지문자 제거 — 파일명 안전화</summary>
    private static string SanitizeFileName(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "시료";
        var invalid = System.IO.Path.GetInvalidFileNameChars();
        var s = raw.Trim();
        foreach (var c in invalid) s = s.Replace(c, '_');
        // 공백 정리 — 연속 공백 1개로
        while (s.Contains("  ")) s = s.Replace("  ", " ");
        return s.Length == 0 ? "시료" : s;
    }

    /// <summary>SN 으로 생태독성_시험기록부에서 가장 최근 기록 id 검색 (없으면 -1)</summary>
    private int FindEcotoxRecordIdBySn(string sn)
    {
        if (string.IsNullOrWhiteSpace(sn)) return -1;
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            if (!DbConnectionFactory.TableExists(conn, "생태독성_시험기록부")) return -1;
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT id FROM `생태독성_시험기록부` WHERE SN=@sn ORDER BY 분석일 DESC, id DESC LIMIT 1";
            cmd.Parameters.AddWithValue("@sn", sn);
            var v = cmd.ExecuteScalar();
            if (v != null && int.TryParse(v.ToString(), out var id)) return id;
        }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[FindEcotoxRecordIdBySn] {ex.Message}"); }
        return -1;
    }

    /// <summary>새 시료 선택 시 입력값 초기화 — 이전 시료 분석 데이터 잔존 방지</summary>
    private void ResetAnalysisState()
    {
        _controlOrganisms   = 20;
        _controlMortalities = 0;
        _tskResult    = null;
        _probitResult = null;
        _records.Clear();
        _testNumber = "1";
        _endpoint = "EC50";
        _concUnit = "%";
        _duration = "24";
        _durUnit  = "H";
        // 표준농도(6.25/12.5/25/50/100)를 디폴트로 자동 적용
        ApplyStandardConcentrations();
    }

    /// <summary>표준농도 자동 입력 — 6.25/12.5/25/50/100% × 생물수 20 균등</summary>
    private void ApplyStandardConcentrations()
    {
        _numConcentrations = 5;
        _concentrations = (double[])EcotoxicityService.StandardConcentrations.Clone();
        _organisms      = new int[5];
        _mortalities    = new int[5];
        for (int i = 0; i < 5; i++) _organisms[i] = EcotoxicityService.StandardOrganismsPerConc;
        _equalOrganisms   = true;
        _organismsPerConc = EcotoxicityService.StandardOrganismsPerConc;
    }

    private void LoadRecordToForm(int id)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT * FROM `생태독성_시험기록부` WHERE id=@id LIMIT 1";
            cmd.Parameters.AddWithValue("@id", id);
            using var r = cmd.ExecuteReader();
            if (!r.Read()) return;

            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < r.FieldCount; i++)
                dict[r.GetName(i)] = r.IsDBNull(i) ? "" : r.GetValue(i)?.ToString() ?? "";

            string Get(string key) => dict.TryGetValue(key, out var v) ? v : "";

            _testDate    = Get("분석일");
            _testNumber  = Get("시험번호");
            _species     = Get("시험종");
            _toxicant    = Get("시료명");
            _duration    = Get("시험시간");
            _durUnit     = Get("시험시간단위");

            int.TryParse(Get("대조군_생물수"), out _controlOrganisms);
            int.TryParse(Get("대조군_사망수"), out _controlMortalities);

            for (int i = 0; i < 8 && i < _concentrations.Length; i++)
            {
                double.TryParse(Get($"농도_{i+1}"), out var conc);
                int.TryParse(Get($"생물수_{i+1}"), out var org);
                int.TryParse(Get($"사망수_{i+1}"), out var mort);
                _concentrations[i] = conc;
                _organisms[i] = org;
                _mortalities[i] = mort;
            }

            // EC50/TU 결과 — 시험기록부 출력 시 필요 (이전엔 누락 → 출력 무반응의 원인)
            double.TryParse(Get("LC50"), out var ec50);
            double.TryParse(Get("TU"),  out var tu);
            double.TryParse(Get("LC50_하한"), out var ciLow);
            double.TryParse(Get("LC50_상한"), out var ciHigh);
            // DB 컬럼은 `분석방법` — 옛버전 호환: `분석법` 도 fallback 으로 시도
            string method = !string.IsNullOrWhiteSpace(Get("분석방법")) ? Get("분석방법")
                          : !string.IsNullOrWhiteSpace(Get("분석법"))   ? Get("분석법")
                          : "TSK";
            double.TryParse(Get("Trim_percent"), out var trim);
            string warn = Get("비고");
            var loadedResult = new EcotoxicityService.EcotoxResult(
                ec50, ciLow, ciHigh, tu, method, trim < 0 ? 0 : trim, false,
                string.IsNullOrWhiteSpace(warn) ? null : warn);
            if (method.Contains("Probit", StringComparison.OrdinalIgnoreCase))
            {
                _probitResult = loadedResult;
                _tskResult = null;
            }
            else
            {
                _tskResult = loadedResult;
                _probitResult = null;
            }

            // _records 에도 추가 — 시험기록부 출력이 _records.Any() 분기 통과하도록
            _records.Clear();
            int validCnt = _concentrations.Count(c => c > 0);
            _records.Add(new TestRecord(
                _testDate,
                !string.IsNullOrWhiteSpace(_selectedSn) ? _selectedSn : _testNumber,
                _species, _toxicant,
                Get("시료명"),
                _tskResult ?? _probitResult ?? loadedResult,
                _probitResult,
                _concentrations.Where(c => c > 0).ToArray(),
                _organisms.Take(validCnt).ToArray(),
                _mortalities.Take(validCnt).ToArray(),
                _controlOrganisms, _controlMortalities,
                _testTemperature, _testPH, _sampleTemperature, _samplePH, _sampleDO,
                $"{_duration} {_durUnit}", _ecCalculationMethod,
                _analysisObservations, _analystName));

            ShowInputForm();
            ShowHistoryPanel();
        }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[Ecotox.LoadRecord] {ex.Message}"); }
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
                    _selectedRecId = capturedRec.Id;

                    // SN 기준 생태독성_시험기록부에서 기존 분석 기록 조회 — 있으면 폼에 로드, 없으면 입력값 초기화
                    int existingRecId = FindEcotoxRecordIdBySn(_selectedSn);
                    if (existingRecId > 0)
                    {
                        LoadRecordToForm(existingRecId);
                    }
                    else
                    {
                        ResetAnalysisState();
                        ShowInputForm();
                    }
                };

                root.Children.Add(item);
            }
        }

        StatsPanelChanged?.Invoke(new ScrollViewer { Content = root });
    }

    // 생태독성 담당자 입력(염분/암모니아/경도) 저장
    private void SaveEcotoxFields()
    {
        if (_selectedRecId < 0) { Log("생태측정 저장 스킵: 선택된 시료 없음"); return; }
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["생태_염분"]     = _tbSalinity?.Text ?? "",
            ["생태_암모니아"] = _tbAmmonia?.Text  ?? "",
            ["생태_경도"]     = _tbHardness?.Text ?? "",
        };
        int rowId = _selectedRecId;
        _ = System.Threading.Tasks.Task.Run(() =>
            ETA.Services.SERVICE1.AnalysisRequestService.UpdateEcotoxFields(rowId, values));
        Log($"생태측정 저장: rowId={rowId} 염분={values["생태_염분"]} 암모니아={values["생태_암모니아"]} 경도={values["생태_경도"]}");
    }

    // ══════════════════════════════════════════════════════════════════════════
    //  Show2: 단계별 입력 폼 (DOS TSK 프로그램 재현)
    // ══════════════════════════════════════════════════════════════════════════
    private void ShowInputForm()
    {
        // 농도가 비어있거나 모두 0이면 표준농도 자동 적용 (사용자가 버튼 안 눌러도)
        if (_concentrations == null || _concentrations.Length == 0 ||
            _concentrations.All(c => c <= 0))
        {
            ApplyStandardConcentrations();
        }

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

        // 현장측정항목 read-only 스트립 — 분석의뢰내역에서 입력된 값을 그대로 표시
        // 그리고 생태독성 담당자가 입력하는 3개 항목(염분/암모니아/경도) 입력란 추가
        if (_selectedRecId >= 0)
        {
            try
            {
                var row = ETA.Services.SERVICE1.AnalysisRequestService.GetRecordRow(_selectedRecId);
                string V(string col) => row.TryGetValue(col, out var s) && !string.IsNullOrWhiteSpace(s) ? s : "—";

                var fieldStrip = new Border
                {
                    Background = AppRes("Panel4Bg"),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(10, 6),
                };
                var fsp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 14 };
                fsp.Children.Add(new TextBlock
                {
                    Text = "📍 현장측정",
                    FontSize = AppTheme.FontSM, FontFamily = Font,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                });
                fsp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
                fsp.Children.Add(MakeReadOnlyItem("온도",       V("현장_온도")       + " ℃"));
                fsp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
                fsp.Children.Add(MakeReadOnlyItem("pH",         V("현장_pH")));
                fsp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
                fsp.Children.Add(MakeReadOnlyItem("용존산소",   V("현장_용존산소")   + " mg/L"));
                fsp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
                fsp.Children.Add(MakeReadOnlyItem("전기전도도", V("현장_전기전도도") + " μS/cm"));
                fsp.Children.Add(new Border { Width = 1, Background = AppRes("ThemeBorderSubtle") });
                fsp.Children.Add(MakeReadOnlyItem("잔류염소",   V("현장_잔류염소")   + " mg/L"));
                fieldStrip.Child = fsp;
                // root 레벨에 추가 — twoCol(시험정보·대조군 좌, 농도설정·데이터 우) 다음에 위치 = 대조군 아래 풀 폭
                fieldStrip.Margin = new Thickness(0, 8, 0, 0);
                root.Children.Add(fieldStrip);

                // 생태독성 담당자 입력 — 염분 / 암모니아 / 경도 + 저장 버튼
                var ecoStrip = new Border
                {
                    Background = AppRes("Panel4Bg"),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(10, 6),
                };
                TextBox MkInput(string initial)
                {
                    return new TextBox
                    {
                        Text       = initial,
                        FontFamily = Font, FontSize = AppTheme.FontSM,
                        Height     = 26, Padding = new Thickness(4, 2),
                        Width      = 72,                // 명시 Width — 라벨/단위와 겹치지 않도록
                        VerticalAlignment = VerticalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center,
                        TextAlignment = Avalonia.Media.TextAlignment.Right,
                    };
                }
                TextBlock MkLbl(string text) => new TextBlock
                {
                    Text       = text,
                    FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                TextBlock MkUnit(string text) => new TextBlock
                {
                    Text       = text,
                    FontFamily = Font, FontSize = AppTheme.FontXS,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                };
                Border MkDivider() => new Border
                {
                    Width = 1, Background = AppRes("ThemeBorderSubtle"),
                    Margin = new Thickness(0, 2),
                };

                _tbSalinity = MkInput(V("생태_염분") == "—" ? "" : V("생태_염분"));
                _tbAmmonia  = MkInput(V("생태_암모니아") == "—" ? "" : V("생태_암모니아"));
                _tbHardness = MkInput(V("생태_경도") == "—" ? "" : V("생태_경도"));

                var btnSaveEco = new Button
                {
                    Content    = "💾  저장",
                    FontFamily = Font, FontSize = AppTheme.FontSM,
                    Padding    = new Thickness(10, 4),
                    Background = AppRes("BtnPrimaryBg"),
                    Foreground = AppRes("BtnPrimaryFg"),
                    BorderBrush = AppRes("BtnPrimaryBorder"),
                    BorderThickness = new Thickness(1),
                    CornerRadius    = new CornerRadius(4),
                    Margin     = new Thickness(8, 0, 0, 0),
                    Cursor     = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
                };
                btnSaveEco.Click += (_, _) => SaveEcotoxFields();

                // 가로 StackPanel — 라벨/입력란/단위가 자연스럽게 쌓이도록
                var ecoSp = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing     = 6,
                    VerticalAlignment = VerticalAlignment.Center,
                };
                ecoSp.Children.Add(new TextBlock
                {
                    Text = "🐟 생태측정",
                    FontFamily = Font, FontSize = AppTheme.FontSM,
                    Foreground = AppRes("FgMuted"),
                    VerticalAlignment = VerticalAlignment.Center,
                });
                ecoSp.Children.Add(MkDivider());
                ecoSp.Children.Add(MkLbl("염분"));    ecoSp.Children.Add(_tbSalinity); ecoSp.Children.Add(MkUnit("‰"));
                ecoSp.Children.Add(MkDivider());
                ecoSp.Children.Add(MkLbl("암모니아")); ecoSp.Children.Add(_tbAmmonia);  ecoSp.Children.Add(MkUnit("mg/L"));
                ecoSp.Children.Add(MkDivider());
                ecoSp.Children.Add(MkLbl("경도"));    ecoSp.Children.Add(_tbHardness); ecoSp.Children.Add(MkUnit("mg/L"));
                ecoSp.Children.Add(btnSaveEco);

                // 좁은 폭에서도 가로 스크롤로 잘리지 않도록
                var ecoScroll = new ScrollViewer
                {
                    Content = ecoSp,
                    HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                    VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                };
                ecoStrip.Child = ecoScroll;
                // 현장측정 strip 바로 아래 풀 폭 — 라벨/입력란 겹침 방지
                ecoStrip.Margin = new Thickness(0, 4, 0, 0);
                root.Children.Add(ecoStrip);
            }
            catch { }
        }

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
            ApplyStandardConcentrations();
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

        var tskBtn     = MakeActionBtn("TSK 계산",     BadgeStatus.Info);
        var probitBtn  = MakeActionBtn("Probit 계산",  BadgeStatus.Info);
        var bothBtn    = MakeActionBtn("TSK + Probit", BadgeStatus.Accent);
        var saveBtn    = MakeActionBtn("💾 DB 저장",   BadgeStatus.Ok);
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
    //  시험기록부 출력 (Word .docx) — 다른 시험기록부와 동일한 시각언어
    // ══════════════════════════════════════════════════════════════════════════

    /// <summary>EcotoxExport.log 에 한 줄 추가 — 출력 단계 추적용. AppPaths.LogsDir 기준.</summary>
    private static void ExportLog(string msg)
    {
        var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [EcotoxExport] {msg}";
        try
        {
            string path = Path.Combine(AppPaths.LogsDir, "EcotoxExport.log");
            File.AppendAllText(path, line + Environment.NewLine);
        }
        catch { }
        System.Diagnostics.Debug.WriteLine(line);
    }

    /// <summary>법정양식(원수 시료 결과 기록부 — ES 04704.1c 별지 1) Word 출력</summary>
    public async Task ExportLegalReportAsync()
    {
        try
        {
            ExportLog("══════ ExportLegalReportAsync 진입 ══════");
            if (_selectedRecId < 0)
            {
                SetExportStatus("출력 실패: Show1에서 시료를 먼저 선택하세요.", ok: false);
                return;
            }

            // 시료 row + 현장/생태측정값
            var row = ETA.Services.SERVICE1.AnalysisRequestService.GetRecordRow(_selectedRecId);
            string V(string col) => row.TryGetValue(col, out var s) && !string.IsNullOrWhiteSpace(s) ? s : "";

            // 분장표준처리에서 오늘 자(=작성일) 생태독성 담당자 조회
            string today = DateTime.Today.ToString("yyyy-MM-dd");
            string 기입자 = "";
            try
            {
                var managers = ETA.Services.SERVICE1.AnalysisRequestService.GetManagersByDate(today);
                foreach (var kv in managers)
                {
                    if (kv.Key.Contains("생태독성", StringComparison.OrdinalIgnoreCase) ||
                        kv.Key.Contains("물벼룩",   StringComparison.OrdinalIgnoreCase))
                    {
                        기입자 = kv.Value;
                        break;
                    }
                }
            }
            catch { }

            // 독성결과 6×5 매트릭스 — Control 행은 _controlOrganisms / _controlMortalities 1열만, 나머지 5개 농도는
            // _concentrations[i] 와 _mortalities[i] 사용 (반복 1회만 입력되는 구조라 1번 열에만 채움)
            string[][] tox = new string[6][];
            // Row 0: Control
            tox[0] = new[]
            {
                _controlOrganisms > 0 ? _controlMortalities.ToString() : "",
                "", "", "", _controlOrganisms > 0
                    ? (_controlMortalities * 100.0 / Math.Max(1, _controlOrganisms)).ToString("0.#")
                    : ""
            };
            // Row 1~5: 6.25, 12.5, 25, 50, 100
            double[] stdConc = { 6.25, 12.5, 25, 50, 100 };
            for (int r = 0; r < 5; r++)
            {
                int idx = -1;
                for (int i = 0; i < _concentrations.Length; i++)
                    if (Math.Abs(_concentrations[i] - stdConc[r]) < 0.001) { idx = i; break; }

                if (idx < 0)
                {
                    tox[r + 1] = new[] { "", "", "", "", "" };
                    continue;
                }
                int org = _organisms[idx];
                int mort = _mortalities[idx];
                double rate = org > 0 ? mort * 100.0 / org : 0;
                tox[r + 1] = new[] { mort.ToString(), "", "", "", org > 0 ? rate.ToString("0.#") : "" };
            }

            // 수질 측정항목 / 독성시험결과 / 독성시험항목은 인쇄 후 수기 작성하도록 공란 처리
            var data = new ETA.Services.SERVICE2.EcotoxicityLegalReportExporter.FormData(
                배출시설:    row.TryGetValue("업체명", out var c) ? c : (row.TryGetValue("약칭", out var ab) ? ab : ""),
                시료채취장소: row.TryGetValue("시료명", out var s2) ? s2 : "",
                시료번호:    row.TryGetValue("접수번호", out var sn) ? sn : "",
                채취일자:    row.TryGetValue("채취일자", out var cd) ? cd : "",
                // 현장 측정항목 — 자동 채움 유지
                온도:        V("현장_온도"),
                pH:          V("현장_pH"),
                용존산소:    V("현장_용존산소"),
                유입수량:    "",
                전기전도도:  V("현장_전기전도도"),
                염분:        V("생태_염분"),
                // 수질 측정항목 — 공란
                잔류염소:    "",
                암모니아:    "",
                경도:        "",
                // 독성시험결과 매트릭스 — 공란
                독성결과:    tox,
                // 독성시험항목 — 공란
                EC50:        "",
                TU:          "",
                통계분석법:  "",
                작성일:      today,
                기입자성명:  기입자);

            // 데스크톱\ETA 생태독성 시험기록부\ (시험기록부/법정양식 통합 폴더)
            string saveDir = EcotoxOutputDir;
            System.IO.Directory.CreateDirectory(saveDir);
            // 파일명: MM-DD {시료명} - 법정양식.docx (Windows 금지문자 제거)
            string mmdd = DateTime.Now.ToString("MM-dd");
            string sampleSafe = SanitizeFileName(data.시료채취장소);
            string filename = $"{mmdd} {sampleSafe} - 법정양식.docx";
            string savePath = System.IO.Path.Combine(saveDir, filename);

            var (ok, msg) = await Task.Run(() =>
                ETA.Services.SERVICE2.EcotoxicityLegalReportExporter.Export(data, savePath));

            if (!ok)
            {
                SetExportStatus($"법정기록부 출력 실패: {msg}", ok: false);
                return;
            }

            SetExportStatus($"✅ 법정기록부 출력 완료 → {savePath}", ok: true);

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName        = savePath,
                    UseShellExecute = true,
                });
            }
            catch (Exception ex) { ExportLog($"  자동 열기 실패: {ex.Message}"); }
        }
        catch (Exception ex)
        {
            ExportLog($"법정기록부 예외: {ex.Message}");
            SetExportStatus($"법정기록부 출력 예외: {ex.Message}", ok: false);
        }
    }

    public async Task ExportTestReportAsync()
    {
        ExportLog("══════ ExportTestReportAsync 진입 ══════");
        ExportLog($"_records.Count = {_records.Count}");
        SetExportStatus("출력 중…", ok: true);

        // 저장된 기록이 없으면 현재 입력된 데이터로 임시 기록 생성
        if (!_records.Any())
        {
            ExportLog($"기록 없음 → 임시생성 분기. _tskResult={(_tskResult==null?"null":"OK")}, _probitResult={(_probitResult==null?"null":"OK")}");
            if (_tskResult == null && _probitResult == null)
            {
                ExportLog("⚠ 계산 결과 없음 — 출력 종료");
                SetExportStatus("출력 실패: 계산 결과 없음. 먼저 EC50 계산을 수행하거나 저장된 기록을 선택하세요.", ok: false);
                return;
            }

            // 현재 입력된 데이터를 임시 기록으로 생성
            var validCnt = _concentrations.Count(c => c > 0);
            ExportLog($"임시 record 생성: validCnt={validCnt}, sampleName={_selectedTreeNameTb?.Text ?? "미지정"}");
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

        string? savePath = null;
        try
        {
            ExportLog($"DTO 변환 시작 — records={_records.Count}");
            var dtos = _records.Select(r => new EcotoxicityWordExporter.Record(
                r.Date, r.TestNo, r.Species, r.Toxicant, r.SampleName,
                r.TskResult, r.ProbitResult,
                r.Conc, r.Org, r.Mort, r.CtrlOrg, r.CtrlMort,
                r.TestTemperature, r.TestPH, r.SampleTemperature, r.SamplePH, r.SampleDO,
                r.Duration, r.EcCalculationMethod, r.Observations, r.AnalystName)).ToList();
            ExportLog($"DTO 변환 완료 — {dtos.Count}건");

            ExportLog("EcotoxicityWordExporter.Export 호출");
            string tmpPath = EcotoxicityWordExporter.Export(dtos);
            ExportLog($"Export 완료 — tmpPath={tmpPath}, exists={File.Exists(tmpPath)}, size={(File.Exists(tmpPath)?new FileInfo(tmpPath).Length:0)}");

            // 데스크톱\ETA 생태독성 시험기록부\ (Show4 와 동일 폴더)
            string saveDir = EcotoxOutputDir;
            Directory.CreateDirectory(saveDir);
            // 파일명: MM-DD {시료명} - 시험기록부.docx (Windows 금지문자 제거)
            string mmddT = DateTime.Now.ToString("MM-dd");
            string sampleT = SanitizeFileName(_toxicant ?? "시료");
            string filename = $"{mmddT} {sampleT} - 시험기록부.docx";
            savePath = Path.Combine(saveDir, filename);
            ExportLog($"복사 시작 → {savePath}");
            File.Copy(tmpPath, savePath, overwrite: true);
            ExportLog($"복사 완료 — exists={File.Exists(savePath)}, size={new FileInfo(savePath).Length}");

            // 자동 열기
            try
            {
                ExportLog($"Process.Start (UseShellExecute) 호출 — {savePath}");
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = savePath,
                    UseShellExecute = true,
                });
                ExportLog("Process.Start 완료 — 시스템 기본앱이 docx 를 열어야 함");
            }
            catch (Exception openEx)
            {
                ExportLog($"❌ 자동 열기 실패: {openEx.GetType().Name}: {openEx.Message}");
            }

            SetExportStatus($"✅ 출력 완료: {filename}", ok: true);
            RefreshRecordsPanel();
            ExportLog("══════ ExportTestReportAsync 정상 종료 ══════");
        }
        catch (Exception wex)
        {
            ExportLog($"❌ EXPORT 예외: {wex.GetType().FullName}: {wex.Message}");
            ExportLog($"   StackTrace: {wex.StackTrace}");
            if (wex.InnerException != null)
                ExportLog($"   Inner: {wex.InnerException.GetType().FullName}: {wex.InnerException.Message}");
            SetExportStatus($"❌ 출력 실패: {wex.GetType().Name}: {wex.Message}", ok: false);
        }
        await Task.CompletedTask;
    }

    private void SetExportStatus(string msg, bool ok)
    {
        _lastExportStatus = msg;
        _lastExportOk     = ok;
        try { Avalonia.Threading.Dispatcher.UIThread.Post(() => RefreshRecordsPanel()); } catch { }
    }

    /// <summary>출력된 docx 가 모이는 "Desktop\ETA 생태독성 시험기록부" 폴더를 OS 파일탐색기로 연다.</summary>
    public void OpenSavedRecordsFolder()
    {
        try
        {
            string folder = EcotoxOutputDir;
            Directory.CreateDirectory(folder);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = folder,
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[생태독성 시험기록부 폴더 열기 오류] {ex.Message}");
        }
    }
}
