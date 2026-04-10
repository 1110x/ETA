using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ETA.Services.Common;
using ETA.Views;

namespace ETA.Views.Pages.PAGE5;

public partial class AiDocClassificationPage : UserControl
{
    // ── 외부(MainPage) 연결 ────────────────────────────────────────────
    public event Action<Control?>? ListPanelChanged;    // Show2
    public event Action<Control?>? DetailPanelChanged;  // Show3
    public event Action<Control?>? StatsPanelChanged;   // Show4

    // ── 파서 항목 목록 ─────────────────────────────────────────────────
    private static readonly List<(string Name, string ParserKey)> ParserItems =
    [
        ("BOD",                   "BOD"),
        ("SS",                    "SS"),
        ("N-Hexan",               "NHex"),
        ("UVVIS",                 "UVVIS"),
        ("TOC-시마즈 (CSV/TXT)",   "TOC_Shimadzu"),
        ("TOC-시마즈 (PDF)",       "TOC_Shimadzu_PDF"),
        ("TOC-예나 (PDF)",         "TOC_NPOC"),
        ("TOC-스칼라 NPOC",        "TOC_Scalar_NPOC"),
        ("TOC-스칼라 TCIC",        "TOC_Scalar_TCIC"),
        ("GC",                    "GC"),
        ("UV-Shimadzu (PDF)",     "UV_Shimadzu_PDF"),
        ("UV-Shimadzu (ASCII)",   "UV_Shimadzu_ASCII"),
        ("UV-Cary (PDF)",         "UV_Cary_PDF"),
        ("UV-Cary (CSV)",         "UV_Cary_CSV"),
        ("ICP",                   "ICP"),
        ("LCMS",                  "LCMS"),
    ];

    private string? _selectedParserKey;
    private ListBox? _fileListBox;

    // ── 학습 진행 UI ──────────────────────────────────────────────────────
    private ProgressBar?  _progressBar;
    private TextBlock?    _progressStatus;
    private TextBlock?    _progressPercent;
    private TextBlock?    _progressLog;
    private int           _progressMax;
    private readonly StringBuilder _logBuf = new();

    private const string LogFile      = "Logs/AiDocClassificationDebug.log";
    private const string AddFileLog   = "Logs/파일추가.log";
    private static string TrainDataPath => Path.Combine(AppContext.BaseDirectory, "Data", "ai_training_data.json");
    private static string OnnxPath      => Path.Combine(AppContext.BaseDirectory, "Data", "분석항목_분류.onnx");
    private static string ScriptPath    => Path.Combine(AppContext.BaseDirectory, "Scripts", "ai_train.py");

    public AiDocClassificationPage()
    {
        InitializeComponent();
        EnsureTables();
        BuildParserTree();
        // Show4는 구독 후 MainPage가 직접 RefreshShow4() 호출
    }

    public void RefreshShow4() => StatsPanelChanged?.Invoke(BuildOnnxInfoPanel());

    // ── DB 초기화 ──────────────────────────────────────────────────────
    private static void EnsureTables()
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            AddLog($"[EnsureTables] DB연결 OK. 타입: {conn.GetType().Name}");

            bool exists = DbConnectionFactory.TableExists(conn, "AI_문서분류_학습데이터");
            AddLog($"[EnsureTables] 테이블 존재: {exists}");

            if (exists)
            {
                // 기존 테이블이 잘못된 DEFAULT를 가질 수 있으므로 컬럼 확인 후 재생성
                try
                {
                    using var testCmd = conn.CreateCommand();
                    testCmd.CommandText = "INSERT INTO `AI_문서분류_학습데이터` (`파서타입`,`파일명`,`파일형식`,`파일텍스트`,`등록일`) VALUES ('__test__','__test__','test','','2000-01-01')";
                    testCmd.ExecuteNonQuery();
                    // 테스트 행 삭제
                    using var delCmd = conn.CreateCommand();
                    delCmd.CommandText = "DELETE FROM `AI_문서분류_학습데이터` WHERE `파서타입`='__test__'";
                    delCmd.ExecuteNonQuery();
                    AddLog("[EnsureTables] 테이블 정상 확인");
                }
                catch
                {
                    // 잘못된 테이블 → 재생성
                    AddLog("[EnsureTables] 기존 테이블 불량 → DROP 후 재생성");
                    using var dropCmd = conn.CreateCommand();
                    dropCmd.CommandText = "DROP TABLE `AI_문서분류_학습데이터`";
                    dropCmd.ExecuteNonQuery();
                    exists = false;
                }
            }

            if (!exists)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = $@"
                    CREATE TABLE `AI_문서분류_학습데이터` (
                        `_id`       INTEGER PRIMARY KEY {DbConnectionFactory.AutoIncrement},
                        `파서타입`   TEXT NOT NULL,
                        `파일명`     TEXT NOT NULL,
                        `파일형식`   TEXT,
                        `파일텍스트` LONGTEXT,
                        `등록일`     TEXT
                    )";
                cmd.ExecuteNonQuery();
                AddLog("[EnsureTables] 테이블 생성 완료");
            }
        }
        catch (Exception ex)
        {
            AddLog($"[EnsureTables] 오류: {ex.Message}");
            Log($"EnsureTables 오류: {ex.Message}");
        }
    }

    // ── 트리뷰 구성 (Show1) ────────────────────────────────────────────
    private void BuildParserTree()
    {
        bool modelReady = File.Exists(OnnxPath);
        ParserTreeView.Items.Clear();
        foreach (var (name, key) in ParserItems)
        {
            int count  = GetFileCount(key);
            var header = BuildTreeHeader(name, count, modelReady);
            var item   = new TreeViewItem { Header = header, Tag = key };
            ParserTreeView.Items.Add(item);
        }
    }

    private static Control BuildTreeHeader(string name, int count, bool modelReady)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

        sp.Children.Add(new TextBlock
        {
            Text      = name,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontBase,
            VerticalAlignment = VerticalAlignment.Center,
        });

        sp.Children.Add(new Border
        {
            CornerRadius = new CornerRadius(9),
            Padding      = new Thickness(6, 1),
            Background   = count > 0
                           ? new SolidColorBrush(Color.Parse("#2563EB"))
                           : new SolidColorBrush(Color.Parse("#D1D5DB")),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text       = count.ToString(),
                FontSize   = AppTheme.FontXS,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = count > 0 ? Brushes.White : new SolidColorBrush(Color.Parse("#6B7280")),
            },
        });

        if (modelReady)
            sp.Children.Add(new TextBlock
            {
                Text      = "ONNX",
                FontSize   = AppTheme.FontXS,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = new SolidColorBrush(Color.Parse("#16A34A")),
                VerticalAlignment = VerticalAlignment.Center,
            });

        return sp;
    }

    // ── 트리뷰 선택 ────────────────────────────────────────────────────
    private void ParserTreeView_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (e.AddedItems.Count == 0) return;
        if (e.AddedItems[0] is not TreeViewItem tvi || tvi.Tag is not string key) return;

        _selectedParserKey = key;
        ListPanelChanged?.Invoke(BuildListPanel(key));
        DetailPanelChanged?.Invoke(null);
    }

    // ── Show2: 파일 목록 + 파일 추가 버튼 ─────────────────────────────
    private Control BuildListPanel(string parserKey)
    {
        var rows  = LoadFiles(parserKey);
        var label = ParserItems.FirstOrDefault(p => p.ParserKey == parserKey).Name ?? parserKey;

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));
        grid.RowDefinitions.Add(new RowDefinition(GridLength.Star));

        // 상단: 제목 + 파일 추가 버튼
        var topBar = new Grid { Margin = new Thickness(10, 8, 10, 4) };
        topBar.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        topBar.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        var title = new TextBlock
        {
            Text      = $"{label} — 학습 파일 ({rows.Count}건)",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontSM,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetColumn(title, 0);

        var addBtn = new Button
        {
            Content    = "+ 파일 추가",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontSM,
            Padding    = new Thickness(10, 4),
        };
        addBtn.Click += async (_, _) => await AddFilesAsync(parserKey);
        Grid.SetColumn(addBtn, 1);

        topBar.Children.Add(title);
        topBar.Children.Add(addBtn);
        Grid.SetRow(topBar, 0);
        grid.Children.Add(topBar);

        // 파일 목록
        var listBox = new ListBox { Margin = new Thickness(8, 0, 8, 8) };
        _fileListBox = listBox;
        foreach (var row in rows)
            listBox.Items.Add(BuildFileItem(row, id =>
            {
                DeleteFile(id);
                BuildParserTree();
                DetailPanelChanged?.Invoke(null);
                ListPanelChanged?.Invoke(BuildListPanel(parserKey));
            }));

        listBox.SelectionChanged += (_, e2) =>
        {
            if (e2.AddedItems.Count > 0 && e2.AddedItems[0] is ListBoxItem lbi && lbi.Tag is AiTrainingFile f)
                DetailPanelChanged?.Invoke(BuildPreviewPanel(f));
        };

        Grid.SetRow(listBox, 1);
        grid.Children.Add(listBox);
        return grid;
    }

    private static ListBoxItem BuildFileItem(AiTrainingFile row, Action<int> onDelete)
    {
        var ext      = (row.Format ?? "").ToUpper();
        var extColor = ext switch
        {
            "PDF"          => "#DC2626",
            "XLSX"         => "#16A34A",
            "CSV" or "TXT" => "#2563EB",
            _              => "#6B7280",
        };

        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(4, 3) };

        panel.Children.Add(new Border
        {
            Width        = 36,
            CornerRadius = new CornerRadius(3),
            Padding      = new Thickness(3, 1),
            Background   = new SolidColorBrush(Color.Parse(extColor)),
            Child = new TextBlock
            {
                Text       = ext.Length > 0 ? ext : "?",
                FontSize   = AppTheme.FontXS,
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                Foreground = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
            },
        });
        panel.Children.Add(new TextBlock
        {
            Text      = row.FileName,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontSM,
            VerticalAlignment = VerticalAlignment.Center,
            Width     = 200,
        });
        panel.Children.Add(new TextBlock
        {
            Text      = row.RegDate,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
            VerticalAlignment = VerticalAlignment.Center,
        });

        var delBtn = new Button
        {
            Content           = "×",
            FontFamily        = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize          = AppTheme.FontBase,
            Padding           = new Thickness(6, 1),
            Margin            = new Thickness(4, 0, 0, 0),
            Background        = Brushes.Transparent,
            Foreground        = new SolidColorBrush(Color.Parse("#DC2626")),
            VerticalAlignment = VerticalAlignment.Center,
        };
        var capturedId = row.Id;
        delBtn.Click += (_, e) =>
        {
            e.Handled = true;   // ListBoxItem 선택 막기
            onDelete(capturedId);
        };
        panel.Children.Add(delBtn);

        return new ListBoxItem { Content = panel, Tag = row };
    }

    // ── 파일 추가 ──────────────────────────────────────────────────────
    private async Task AddFilesAsync(string parserKey)
    {
        AddLog($"=== 파일추가 시작 | 파서: {parserKey} ===");
        AddLog($"ListPanelChanged 구독 여부: {ListPanelChanged != null}");

        var topLevel = TopLevel.GetTopLevel(this);
        AddLog($"TopLevel: {(topLevel == null ? "NULL ← 문제!" : "OK")}");
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title         = "RawDATA 파일 선택",
            AllowMultiple = true,
            FileTypeFilter =
            [
                new FilePickerFileType("분석 데이터 파일")
                {
                    Patterns = ["*.xlsx", "*.csv", "*.txt", "*.pdf"],
                },
            ],
        });
        AddLog($"파일 선택 개수: {files.Count}");
        if (files.Count == 0) { AddLog("선택된 파일 없음 → 종료"); return; }

        int saved = 0;
        foreach (var file in files)
        {
            AddLog($"  처리 중: {file.Name}");
            try
            {
                await using var stream = await file.OpenReadAsync();
                using var ms = new MemoryStream();
                await stream.CopyToAsync(ms);
                AddLog($"  읽기 완료: {ms.Length:N0} bytes");

                var ext  = Path.GetExtension(file.Name).TrimStart('.').ToLower();
                var text = await Task.Run(() => ExtractTextFromBytes(ms.ToArray(), ext));
                AddLog($"  텍스트 추출: {text.Length:N0}자");

                SaveFile(parserKey, file.Name, ext, text);

                int countAfter = GetFileCount(parserKey);
                AddLog($"  DB저장 후 {parserKey} 파일 수: {countAfter}");
                Log($"파일 저장: {file.Name} → {parserKey}");
                saved++;
            }
            catch (Exception ex)
            {
                AddLog($"  오류: {ex.Message}");
                Log($"파일 추가 오류 ({file.Name}): {ex.Message}");
            }
        }

        AddLog($"저장 완료: {saved}/{files.Count}개");
        AddLog("BuildParserTree() 호출");
        BuildParserTree();

        AddLog($"ListPanelChanged?.Invoke() 호출 (null={ListPanelChanged == null})");
        var panel = BuildListPanel(parserKey);
        AddLog($"BuildListPanel 완료: {panel.GetType().Name}");
        ListPanelChanged?.Invoke(panel);
        AddLog("=== 파일추가 완료 ===\n");
    }

    private static void AddLog(string msg)
    {
        try { File.AppendAllText(AddFileLog, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n"); }
        catch { }
    }

    // ── AI 학습 (서브메뉴 "학습" → MainPage에서 호출) ──────────────────
    public async Task LearnAsync()
    {
        // ── 총 파일 수 계산 후 진행 패널 표시 ─────────────────────────
        int totalFileCount = ParserItems.Sum(p => GetFileCount(p.ParserKey));
        ShowProgressPanel(totalFileCount);
        await Task.Yield();

        // ── 파일별 텍스트 추출 ─────────────────────────────────────────
        var samples   = new List<object>();
        int processed = 0;

        foreach (var (name, key) in ParserItems)
        {
            var files = LoadFiles(key);
            if (files.Count == 0) continue;

            AppendLog($"▶ {name}  ({files.Count}개 파일)");
            await Task.Yield();

            int okCount = 0;
            foreach (var f in files)
            {
                SetProgress(processed, $"{name}  ›  {f.FileName}");
                await Task.Yield();

                var text = await Task.Run(() => ExtractText(f));
                if (!string.IsNullOrWhiteSpace(text))
                {
                    samples.Add(new { text, label = key });
                    okCount++;
                }
                processed++;
                SetProgress(processed);
                await Task.Yield();
            }
            AppendLog($"   ✓ {okCount}/{files.Count}개 추출 완료");
            await Task.Yield();
        }

        if (samples.Count == 0)
        {
            ShowResult("학습할 파일이 없습니다.\n각 분석항목에 파일을 추가해주세요.");
            return;
        }

        var labelCounts = samples
            .GroupBy(s => ((dynamic)s).label.ToString())
            .ToDictionary(g => g.Key, g => g.Count());
        int classCount = labelCounts.Count;

        if (classCount < 2)
        {
            ShowResult($"최소 2종 이상의 분석항목에 파일이 필요합니다.\n현재: {classCount}종");
            return;
        }

        // ── 훈련 데이터 JSON 저장 ──────────────────────────────────────
        var trainData = new { samples, labels = ParserItems.Select(p => p.ParserKey).ToArray() };
        var json      = JsonSerializer.Serialize(trainData, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(TrainDataPath, json, new System.Text.UTF8Encoding(false));
        AppendLog($"\n총 {samples.Count}건  {classCount}종 → ai_training_data.json 저장 완료");
        Log($"훈련 데이터 저장: {samples.Count}건, {classCount}종");
        await Task.Yield();

        // ── Python 훈련 (Indeterminate 모드) ──────────────────────────
        if (_progressBar     != null) _progressBar.IsIndeterminate = true;
        if (_progressPercent != null) _progressPercent.Text = "";
        if (_progressStatus  != null) _progressStatus.Text  = "Python 모델 훈련 중...";
        AppendLog("\n▶ python Scripts/ai_train.py 실행 중...");
        await Task.Yield();

        try
        {
            var (success, _) = await RunPythonAsync(line => AppendLog($"   {line}"));

            if (_progressBar != null) _progressBar.IsIndeterminate = false;

            if (success)
            {
                SetProgress(_progressMax, "완료!");
                AppendLog("\n✅ ONNX 모델 생성 완료!");
                ETA.Services.SERVICE4.AiParserClassifier.Reload();

                // 시그니처 DB 자동 빌드
                try
                {
                    var (built, skipped) = await Task.Run(() =>
                        ETA.Services.SERVICE4.SignatureClassifier.BuildFromTrainingData());
                    AppendLog($"📌 시그니처 DB 갱신: {built}개 저장, {skipped}개 건너뜀");
                    Log($"시그니처 DB 빌드 완료: {built}건");
                }
                catch (Exception sigEx)
                {
                    AppendLog($"⚠ 시그니처 빌드 실패: {sigEx.Message}");
                    Log($"시그니처 빌드 오류: {sigEx.Message}");
                }

                BuildParserTree();
                StatsPanelChanged?.Invoke(BuildOnnxInfoPanel());
                Log("ONNX 학습 완료");
            }
            else
            {
                if (_progressStatus != null) _progressStatus.Text = "오류 발생";
                AppendLog("\n❌ 훈련 오류. 위 로그를 확인하세요.");
                Log("Python 훈련 오류");
            }
        }
        catch (Exception ex)
        {
            if (_progressBar    != null) _progressBar.IsIndeterminate = false;
            if (_progressStatus != null) _progressStatus.Text = "실행 오류";
            AppendLog($"\n❌ {ex.Message}");
            AppendLog("▶ 패키지 미설치 시:  pip install -r Scripts/requirements.txt");
            AppendLog("▶ Python 미설치 시:  python.org 에서 설치 (PATH 포함 체크)");
            Log($"LearnAsync 예외: {ex.Message}");
        }
    }

    /// <summary>windows: py → python → python3 순으로 실행 가능한 명령어를 찾아 반환</summary>
    private static string FindPythonExe()
    {
        foreach (var cmd in new[] { "py", "python", "python3" })
        {
            try
            {
                var test = Process.Start(new ProcessStartInfo
                {
                    FileName               = cmd,
                    Arguments              = "--version",
                    UseShellExecute        = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError  = true,
                    CreateNoWindow         = true,
                });
                test?.WaitForExit(3000);
                if (test?.ExitCode == 0) return cmd;
            }
            catch { }
        }
        throw new InvalidOperationException(
            "Python을 찾을 수 없습니다.\n" +
            "python.org 에서 설치 후 PATH에 추가하거나\n" +
            "Windows Python Launcher(py)를 설치해주세요.");
    }

    private static async Task<(bool Success, string Output)> RunPythonAsync(Action<string>? onLine = null)
    {
        var pythonExe = FindPythonExe();
        onLine?.Invoke($"python 명령어: {pythonExe}");

        var psi = new ProcessStartInfo
        {
            FileName               = pythonExe,
            Arguments              = ScriptPath,
            UseShellExecute        = false,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            CreateNoWindow         = true,
            WorkingDirectory       = AppContext.BaseDirectory,
            StandardOutputEncoding = System.Text.Encoding.UTF8,
            StandardErrorEncoding  = System.Text.Encoding.UTF8,
        };
        psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

        using var proc = Process.Start(psi)
            ?? throw new InvalidOperationException("python 프로세스를 시작할 수 없습니다.");

        var sb     = new StringBuilder();
        var errSb  = new StringBuilder();

        // stdout + stderr 병렬 스트리밍 (UI 업데이트는 UI 스레드에서)
        void PostLine(string msg) =>
            Avalonia.Threading.Dispatcher.UIThread.Post(() => onLine?.Invoke(msg));

        var stdoutTask = Task.Run(async () =>
        {
            while (!proc.StandardOutput.EndOfStream)
            {
                var line = await proc.StandardOutput.ReadLineAsync();
                if (line == null) continue;
                sb.AppendLine(line);
                PostLine(line);
            }
        });
        var stderrTask = Task.Run(async () =>
        {
            while (!proc.StandardError.EndOfStream)
            {
                var line = await proc.StandardError.ReadLineAsync();
                if (line == null) continue;
                errSb.AppendLine(line);
                PostLine($"[오류] {line}");
            }
        });

        await Task.WhenAll(stdoutTask, stderrTask);
        await proc.WaitForExitAsync();

        bool success = proc.ExitCode == 0 && File.Exists(OnnxPath);
        return (success, sb.ToString().Trim());
    }

    // ── 텍스트 추출 (bytes + ext → string) ───────────────────────────
    private static string ExtractText(AiTrainingFile file) => file.Text ?? "";

    private static string ExtractTextFromBytes(byte[] data, string ext)
    {
        if (data.Length == 0) return "";
        try
        {
            return ext switch
            {
                "csv" or "txt" => Encoding.UTF8.GetString(data),
                "xlsx"         => ExtractXlsx(data),
                "pdf"          => ExtractPdf(data),
                _              => Encoding.UTF8.GetString(data),
            };
        }
        catch { return ""; }
    }

    private static string ExtractXlsx(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var wb = new XLWorkbook(ms);
        var sb = new StringBuilder();
        foreach (var ws in wb.Worksheets)
        {
            foreach (var row in ws.RowsUsed().Take(50))
            {
                var cells = row.CellsUsed().Select(c => c.GetString()).Where(s => !string.IsNullOrWhiteSpace(s));
                sb.AppendLine(string.Join("\t", cells));
            }
        }
        return sb.ToString();
    }

    private static string ExtractPdf(byte[] data)
    {
        using var ms  = new MemoryStream(data);
        using var doc = UglyToad.PdfPig.PdfDocument.Open(ms);
        var sb = new StringBuilder();
        foreach (var page in doc.GetPages().Take(3))
        {
            foreach (var word in page.GetWords())
                sb.Append(word.Text).Append(' ');
            sb.AppendLine();
        }
        return sb.ToString();
    }

    // ── 미리보기 (Show3) ───────────────────────────────────────────────
    private static Control BuildPreviewPanel(AiTrainingFile row)
    {
        var scroll = new ScrollViewer { Margin = new Thickness(8) };
        var sp = new StackPanel { Spacing = 4, Margin = new Thickness(8) };

        sp.Children.Add(new TextBlock
        {
            Text       = row.FileName,
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontSM,
            FontWeight = FontWeight.SemiBold,
        });
        sp.Children.Add(new TextBlock
        {
            Text       = $"{row.ParserType}  |  {row.Format}  |  {row.RegDate}",
            FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize   = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
        });

        var text = row.Text ?? "";
        if (!string.IsNullOrWhiteSpace(text))
        {
            if (text.Length > 5000) text = text[..5000] + "\n...(이하 생략)";
            sp.Children.Add(new TextBlock
            {
                Text         = text,
                FontFamily   = new FontFamily("Consolas, Courier New"),
                FontSize     = AppTheme.FontXS,
                TextWrapping = TextWrapping.Wrap,
                Margin       = new Thickness(0, 6, 0, 0),
            });
        }
        else
        {
            sp.Children.Add(new TextBlock
            {
                Text       = $"[{row.Format?.ToUpper()} — 텍스트 없음]",
                FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
                FontSize   = AppTheme.FontSM,
                Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
                Margin     = new Thickness(0, 6, 0, 0),
            });
        }

        scroll.Content = sp;
        return scroll;
    }

    // ── 학습 진행 패널 (Show3) ────────────────────────────────────────────
    private void ShowProgressPanel(int total)
    {
        _progressMax = Math.Max(total, 1);
        _logBuf.Clear();
        var ff = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");

        var root = new StackPanel { Spacing = 6, Margin = new Thickness(12, 10) };

        root.Children.Add(new TextBlock
        {
            Text       = "AI 모델 학습",
            FontFamily = ff, FontSize = AppTheme.FontBase, FontWeight = FontWeight.SemiBold,
            Margin     = new Thickness(0, 0, 0, 4),
        });

        _progressStatus = new TextBlock
        {
            Text         = "준비 중...",
            FontFamily   = ff, FontSize = AppTheme.FontSM,
            Foreground   = new SolidColorBrush(Color.Parse("#374151")),
            TextTrimming = TextTrimming.CharacterEllipsis,
        };
        root.Children.Add(_progressStatus);

        // 진행바 + 퍼센트
        var pbRow = new Grid { Margin = new Thickness(0, 2, 0, 2) };
        pbRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
        pbRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        _progressBar = new ProgressBar { Minimum = 0, Maximum = _progressMax, Value = 0, Height = 16 };
        Grid.SetColumn(_progressBar, 0);
        pbRow.Children.Add(_progressBar);

        _progressPercent = new TextBlock
        {
            Text      = "0%",
            FontFamily = ff, FontSize = AppTheme.FontXS,
            Margin    = new Thickness(8, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Width     = 38, TextAlignment = TextAlignment.Right,
        };
        Grid.SetColumn(_progressPercent, 1);
        pbRow.Children.Add(_progressPercent);
        root.Children.Add(pbRow);

        root.Children.Add(new Border
        {
            Height = 1, Background = new SolidColorBrush(Color.Parse("#E5E7EB")),
            Margin = new Thickness(0, 4),
        });
        root.Children.Add(new TextBlock
        {
            Text = "진행 로그", FontFamily = ff, FontSize = AppTheme.FontXS,
            Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
        });

        _progressLog = new TextBlock
        {
            Text = "", FontFamily = new FontFamily("Consolas, Courier New"),
            FontSize = AppTheme.FontSM, TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.Parse("#D1FAE5")),
            Margin = new Thickness(8, 6),
        };
        root.Children.Add(new Border
        {
            CornerRadius = new CornerRadius(6),
            Background   = new SolidColorBrush(Color.Parse("#0F172A")),
            MaxHeight    = 380,
            Child        = new ScrollViewer { Content = _progressLog },
        });

        DetailPanelChanged?.Invoke(root);
    }

    private void SetProgress(int current, string? status = null)
    {
        if (_progressBar != null)
        {
            _progressBar.IsIndeterminate = false;
            _progressBar.Value           = current;
        }
        if (_progressPercent != null)
            _progressPercent.Text = $"{(double)current / _progressMax * 100:0}%";
        if (status != null && _progressStatus != null)
            _progressStatus.Text = status;
    }

    private void AppendLog(string line)
    {
        _logBuf.AppendLine(line);
        if (_progressLog != null)
            _progressLog.Text = _logBuf.ToString();
    }

    // ── Show4: ONNX 모델 정보 ─────────────────────────────────────────
    private static Control BuildOnnxInfoPanel()
    {
        var ff   = new FontFamily("avares://ETA/Assets/Fonts#Pretendard");
        var root = new StackPanel { Spacing = 6, Margin = new Thickness(10, 8) };

        root.Children.Add(new TextBlock
        {
            Text = "ONNX 모델", FontFamily = ff,
            FontSize = AppTheme.FontSM, FontWeight = FontWeight.SemiBold,
            Margin = new Thickness(0, 0, 0, 4),
        });

        // Data/ 폴더의 .onnx 파일 목록
        var onnxFiles = new List<FileInfo>();
        try
        {
            var dataDir = Path.Combine(AppContext.BaseDirectory, "Data");
            if (Directory.Exists(dataDir))
                onnxFiles = Directory.GetFiles(dataDir, "*.onnx")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();
        }
        catch { }

        if (onnxFiles.Count == 0)
        {
            root.Children.Add(new Border
            {
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 8),
                Background = new SolidColorBrush(Color.Parse("#FEF3C7")),
                Child = new TextBlock
                {
                    Text = "생성된 모델 없음\n「학습」 버튼으로 모델을 만드세요.",
                    FontFamily = ff, FontSize = AppTheme.FontXS,
                    Foreground = new SolidColorBrush(Color.Parse("#92400E")),
                    TextWrapping = TextWrapping.Wrap,
                },
            });
            return root;
        }

        foreach (var fi in onnxFiles)
        {
            var isActive = fi.Name == "분析항목_분류.onnx";
            var sizeKb   = fi.Length / 1024.0;
            var date     = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm");

            var card = new Border
            {
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 8),
                Margin          = new Thickness(0, 0, 0, 4),
                BorderThickness = new Thickness(isActive ? 2 : 1),
                BorderBrush     = new SolidColorBrush(Color.Parse(isActive ? "#2563EB" : "#D1D5DB")),
                Background      = new SolidColorBrush(Color.Parse(isActive ? "#EFF6FF" : "#F9FAFB")),
            };

            var sp = new StackPanel { Spacing = 3 };

            // 파일명 + 활성 뱃지
            var titleRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            titleRow.Children.Add(new TextBlock
            {
                Text = fi.Name, FontFamily = ff, FontSize = AppTheme.FontXS,
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap,
            });
            if (isActive)
                titleRow.Children.Add(new Border
                {
                    CornerRadius = new CornerRadius(8),
                    Padding = new Thickness(5, 1),
                    Background = new SolidColorBrush(Color.Parse("#2563EB")),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = "현재", FontFamily = ff, FontSize = AppTheme.FontXS,
                        Foreground = Brushes.White,
                    },
                });
            sp.Children.Add(titleRow);

            sp.Children.Add(new TextBlock
            {
                Text = $"{sizeKb:0.0} KB  |  {date}",
                FontFamily = ff, FontSize = AppTheme.FontXS,
                Foreground = new SolidColorBrush(Color.Parse("#6B7280")),
            });

            card.Child = sp;
            root.Children.Add(card);
        }

        return new ScrollViewer { Content = root };
    }

    private void ShowResult(string text)
    {
        var scroll = new ScrollViewer { Margin = new Thickness(8) };
        scroll.Content = new TextBlock
        {
            Text         = text,
            FontFamily   = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
            FontSize     = AppTheme.FontSM,
            TextWrapping = TextWrapping.Wrap,
            Margin       = new Thickness(8),
        };
        DetailPanelChanged?.Invoke(scroll);
    }

    // ── DB CRUD ────────────────────────────────────────────────────────
    private static void SaveFile(string parserKey, string fileName, string format, string text)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "INSERT INTO `AI_문서분류_학습데이터` (`파서타입`,`파일명`,`파일형식`,`파일텍스트`,`등록일`) VALUES (@k,@n,@f,@t,@dt)";
            var p1 = cmd.CreateParameter(); p1.ParameterName = "@k";  p1.Value = parserKey;                          cmd.Parameters.Add(p1);
            var p2 = cmd.CreateParameter(); p2.ParameterName = "@n";  p2.Value = fileName;                           cmd.Parameters.Add(p2);
            var p3 = cmd.CreateParameter(); p3.ParameterName = "@f";  p3.Value = format;                             cmd.Parameters.Add(p3);
            var p4 = cmd.CreateParameter(); p4.ParameterName = "@t";  p4.Value = text;                               cmd.Parameters.Add(p4);
            var p5 = cmd.CreateParameter(); p5.ParameterName = "@dt"; p5.Value = DateTime.Now.ToString("yyyy-MM-dd"); cmd.Parameters.Add(p5);
            int rows = cmd.ExecuteNonQuery();
            AddLog($"  [SaveFile] INSERT rows={rows}  {fileName}");
        }
        catch (Exception ex)
        {
            AddLog($"  [SaveFile] 오류: {ex.GetType().Name}: {ex.Message}");
            Log($"SaveFile 오류: {ex.Message}");
        }
    }

    private static List<AiTrainingFile> LoadFiles(string parserKey)
    {
        var result = new List<AiTrainingFile>();
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT `_id`,`파서타입`,`파일명`,`파일형식`,`파일텍스트`,`등록일` FROM `AI_문서분류_학습데이터` WHERE `파서타입`=@k ORDER BY `_id` DESC";
            var p = cmd.CreateParameter(); p.ParameterName = "@k"; p.Value = parserKey; cmd.Parameters.Add(p);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                result.Add(new AiTrainingFile
                {
                    Id         = r.IsDBNull(0) ? 0  : r.GetInt32(0),
                    ParserType = r.IsDBNull(1) ? "" : r.GetString(1),
                    FileName   = r.IsDBNull(2) ? "" : r.GetString(2),
                    Format     = r.IsDBNull(3) ? "" : r.GetString(3),
                    Text       = r.IsDBNull(4) ? "" : r.GetString(4),
                    RegDate    = r.IsDBNull(5) ? "" : r.GetString(5),
                });
        }
        catch (Exception ex) { Log($"LoadFiles 오류: {ex.Message}"); }
        return result;
    }

    private static void DeleteFile(int id)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "DELETE FROM `AI_문서분류_학습데이터` WHERE `_id`=@id";
            var p = cmd.CreateParameter(); p.ParameterName = "@id"; p.Value = id; cmd.Parameters.Add(p);
            cmd.ExecuteNonQuery();
            AddLog($"[DeleteFile] id={id} 삭제 완료");
        }
        catch (Exception ex)
        {
            AddLog($"[DeleteFile] 오류: {ex.Message}");
            Log($"DeleteFile 오류: {ex.Message}");
        }
    }

    private static int GetFileCount(string parserKey)
    {
        try
        {
            using var conn = DbConnectionFactory.CreateConnection();
            conn.Open();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM `AI_문서분류_학습데이터` WHERE `파서타입`=@k";
            var p = cmd.CreateParameter(); p.ParameterName = "@k"; p.Value = parserKey; cmd.Parameters.Add(p);
            var v = cmd.ExecuteScalar();
            return v == null || v == DBNull.Value ? 0 : Convert.ToInt32(v);
        }
        catch (Exception ex)
        {
            AddLog($"  [GetFileCount] 오류: {ex.GetType().Name}: {ex.Message}");
            return 0;
        }
    }

    private static void Log(string msg)
    {
        try { File.AppendAllText(LogFile, $"[{DateTime.Now:HH:mm:ss}] {msg}\n"); }
        catch { }
    }
}

// ── 모델 ───────────────────────────────────────────────────────────────
public class AiTrainingFile
{
    public int     Id         { get; set; }
    public string  ParserType { get; set; } = "";
    public string  FileName   { get; set; } = "";
    public string? Format     { get; set; }
    public string? Text       { get; set; }   // 저장 시 추출된 텍스트 (BLOB 대신)
    public string  RegDate    { get; set; } = "";
}
