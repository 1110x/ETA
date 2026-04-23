using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;
using ClosedXML.Excel;

namespace ETA.Services.Common;

/// <summary>ERP iU (neoweb.exe) WinForms 자동화 — Windows P/Invoke 기반</summary>
public static class ErpUiAutoService
{
    // ── 파일 경로 ────────────────────────────────────────────────────────────
    private static readonly string LogPath =
        Path.Combine(AppPaths.LogsDir, "ERP.log");

    private static readonly string ExcelPath =
        Path.Combine(AppPaths.RootPath, "Data", "ERP 자료", "ERP 자료입력.xlsx");

    // ── 파일 로그 ────────────────────────────────────────────────────────────
    public static void Log(string msg)
    {
        var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        if (App.EnableLogging)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(LogPath, line + Environment.NewLine);
            }
            catch { }
        }
    }

    // ── Win32 P/Invoke ──────────────────────────────────────────────────────
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsProc fn, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int n);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetWindowText(IntPtr hWnd, StringBuilder sb, int n);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string? lParam);

    [DllImport("user32.dll")]
    static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    static extern uint GetLastError();

    [DllImport("user32.dll")]
    static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("kernel32.dll")]
    static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern IntPtr GetFocus();

    [DllImport("user32.dll")]
    static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP   = 0x0004;

    private const uint KEYEVENTF_KEYUP = 0x0002;

    private struct RECT { public int Left, Top, Right, Bottom; }

    private const int  GWL_STYLE    = -16;
    private const uint ES_READONLY  = 0x0800;

    private const uint WM_KEYDOWN  = 0x0100;
    private const uint WM_KEYUP    = 0x0101;
    private const uint WM_CHAR     = 0x0102;
    private const byte VK_CONTROL  = 0x11;
    private const byte VK_DELETE   = 0x2E;
    private const byte VK_HOME     = 0x24;
    private const byte VK_DOWN     = 0x28;
    private const byte VK_OEM_PERIOD = 0xBE; // '.'
    private const byte VK_OEM_MINUS  = 0xBD; // '-'



    private const uint WM_SETTEXT    = 0x000C;
    private const uint WM_GETTEXT    = 0x000D;
    private const uint WM_PASTE      = 0x0302;
    private const uint EM_SETSEL     = 0x00B1;
    private const uint EM_REPLACESEL = 0x00C2;
    private const uint BM_CLICK      = 0x00F5;
    private const uint WM_LBUTTONDOWN = 0x0201;
    private const uint WM_LBUTTONUP   = 0x0202;

    // ── 내부 컨트롤 정보 ────────────────────────────────────────────────────
    public record ControlInfo(IntPtr Hwnd, string ClassName, string Text, bool Visible);

    // ── ERP iU 프로세스명 ────────────────────────────────────────────────────
    private const string ProcessName = "neoweb";

    // ── 권한 확인 / 재실행 ────────────────────────────────────────────────────
    public static bool IsAdmin()
    {
        if (!OperatingSystem.IsWindows()) return false;
        using var id = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator);
    }

    /// <summary>관리자 권한으로 ETA 재실행.</summary>
    public static bool RestartAsAdmin()
    {
        try
        {
            var exe = Process.GetCurrentProcess().MainModule?.FileName;
            if (exe == null) return false;
            Process.Start(new ProcessStartInfo(exe)
            {
                UseShellExecute = true,
                Verb = "runas",
            });
            return true;
        }
        catch { return false; }
    }

    // =========================================================================
    // Excel 데이터 로드
    // =========================================================================
    // ERP AnalyteLabels 순서: BOD, TOC, SS, NH, PN, TN, TP
    // Excel '자료입력' 열 순서: A=시료명, B=S/N, C=BOD, D=TOC, E=SS, F=T-N, G=T-P, H=NH, I=PN
    // → ERP 순서에 맞춰 [C,D,E,H,I,F,G] 로 읽어야 함

    public record ExcelRow(int 번호, string SN, string 시료명, string[] Values);

    /// <summary>Excel 전체 데이터 로드 (한 번에 열기)</summary>
    public static List<ExcelRow> LoadAllExcelData() => LoadAllExcelData(ExcelPath);

    /// <summary>지정된 경로의 Excel 데이터 로드</summary>
    public static List<ExcelRow> LoadAllExcelData(string path)
    {
        var result = new List<ExcelRow>();
        try
        {
            if (!File.Exists(path)) { Log($"[LoadAllExcelData] 파일 없음: {path}"); return result; }
            using var fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);
            var ws = wb.Worksheet("자료입력");

            for (int r = 2; r <= 300; r++)
            {
                var nameCell = ws.Cell(r, 1);
                if (nameCell.IsEmpty()) break;
                string name = nameCell.GetString();
                if (string.IsNullOrWhiteSpace(name)) break;

                string sn  = ws.Cell(r, 2).GetString();
                string bod = FormatCell(ws.Cell(r, 3));
                string toc = FormatCell(ws.Cell(r, 4));
                string ss  = FormatCell(ws.Cell(r, 5));
                string tn  = FormatCell(ws.Cell(r, 6));
                string tp  = FormatCell(ws.Cell(r, 7));
                string nh  = FormatCell(ws.Cell(r, 8));
                string pn  = FormatCell(ws.Cell(r, 9));

                // ERP 입력 순서: BOD, TOC, SS, NH, PN, TN, TP
                result.Add(new ExcelRow(r - 1, sn, name, [bod, toc, ss, nh, pn, tn, tp]));
            }
            Log($"[LoadAllExcelData] {result.Count}행 로드");
        }
        catch (Exception ex) { Log($"[LoadAllExcelData] 오류: {ex.Message}"); }
        return result;
    }

    /// <summary>
    /// Excel 자료입력 시트의 rowNumber번째 행(1-based)을 읽어
    /// ERP 입력 순서(BOD,TOC,SS,NH,PN,TN,TP)로 반환합니다.
    /// </summary>
    public static (string 시료명, string[] values) LoadExcelRow(int rowNumber)
    {
        try
        {
            if (!File.Exists(ExcelPath))
            {
                Log($"[LoadExcelRow] 파일 없음: {ExcelPath}");
                return ("파일 없음", []);
            }
            using var fs = new System.IO.FileStream(ExcelPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
            using var wb = new XLWorkbook(fs);
            var ws = wb.Worksheet("자료입력");
            int r = rowNumber + 1; // 1행은 헤더

            string name = ws.Cell(r, 1).GetString();           // A = 시료명
            string bod  = FormatCell(ws.Cell(r, 3));            // C = BOD
            string toc  = FormatCell(ws.Cell(r, 4));            // D = TOC
            string ss   = FormatCell(ws.Cell(r, 5));            // E = SS
            string tn   = FormatCell(ws.Cell(r, 6));            // F = T-N
            string tp   = FormatCell(ws.Cell(r, 7));            // G = T-P
            string nh   = FormatCell(ws.Cell(r, 8));            // H = NH
            string pn   = FormatCell(ws.Cell(r, 9));            // I = PN

            // ERP 입력 순서: BOD, TOC, SS, NH, PN, TN, TP
            return (name, [bod, toc, ss, nh, pn, tn, tp]);
        }
        catch (Exception ex)
        {
            Log($"[LoadExcelRow] 오류: {ex.Message}");
            return ("오류", []);
        }
    }

    private static string FormatCell(IXLCell cell)
    {
        if (cell.IsEmpty()) return "";
        if (cell.DataType == XLDataType.Number)
        {
            double v = cell.GetDouble();
            // 불필요한 소수점 제거 (예: 10.0 → "10", 31.5 → "31.5")
            return v == Math.Floor(v) ? ((long)v).ToString() : v.ToString("G");
        }
        return cell.GetString();
    }

    // =========================================================================
    // Probe — 실행 가능 여부 전체 점검
    // =========================================================================
    public static ProbeResult Probe()
    {
        var r = new ProbeResult();

        var procs = Process.GetProcessesByName(ProcessName);
        if (procs.Length == 0)
        {
            r.Add("❌", $"'{ProcessName}.exe' 프로세스를 찾을 수 없습니다.");
            r.Add("   ", "ERP iU를 실행한 후 다시 시도하세요.");
            return r;
        }
        var proc = procs[0];
        r.Add("✅", $"프로세스 발견  PID={proc.Id}  제목='{proc.MainWindowTitle}'");

        var hwnd = proc.MainWindowHandle;
        if (hwnd == IntPtr.Zero)
        {
            r.Add("❌", "메인 윈도우 핸들이 없습니다 (최소화 상태?)");
            return r;
        }
        r.Add("✅", $"메인 윈도우  HWND=0x{hwnd:X8}");

        var all = EnumAllChildren(hwnd);
        r.Add("✅", $"자식 컨트롤 총 {all.Count}개 발견");

        var edits = all.FindAll(c => c.ClassName.StartsWith("WindowsForms10.EDIT", StringComparison.OrdinalIgnoreCase));
        r.Add(edits.Count > 0 ? "✅" : "❌", $"EDIT 컨트롤 {edits.Count}개 발견");
        for (int i = 0; i < Math.Min(edits.Count, 12); i++)
        {
            var e = edits[i];
            r.Add("   ", $"[{i + 1}]  HWND=0x{e.Hwnd:X8}  vis={e.Visible}  val='{e.Text}'");
        }

        var btns = all.FindAll(c => c.ClassName.StartsWith("WindowsForms10.BUTTON", StringComparison.OrdinalIgnoreCase));
        r.Add(btns.Count > 0 ? "✅" : "❌", $"BUTTON 컨트롤 {btns.Count}개 발견");
        foreach (var b in btns)
            r.Add("   ", $"  HWND=0x{b.Hwnd:X8}  제목='{b.Text}'  vis={b.Visible}");

        var doneBtn = btns.Find(b => b.Text.Contains("분석완료") || b.Text.Contains("완료"));
        if (doneBtn != null)
            r.Add("✅", $"'분석완료' 버튼 발견  HWND=0x{doneBtn.Hwnd:X8}  → BM_CLICK 전송 가능");
        else
            r.Add("⚠️", "'분석완료' 버튼 미발견 (화면이 로드된 상태에서 실행하세요)");

        // Excel 파일 확인
        r.Add(File.Exists(ExcelPath) ? "✅" : "⚠️",
              $"Excel 파일: {(File.Exists(ExcelPath) ? "있음" : "없음")}  {ExcelPath}");

        r.Add("──", "─────────────────────────────────────────");
        r.Add(r.CanAutomate ? "🟢" : "🔴",
              r.CanAutomate ? "자동화 가능합니다." : "일부 항목 확인 필요.");
        return r;
    }

    // =========================================================================
    // 모든 자식 윈도우 열거
    // =========================================================================
    public static List<ControlInfo> EnumAllChildren(IntPtr parent)
    {
        var list = new List<ControlInfo>();
        EnumChildWindows(parent, (h, _) =>
        {
            var sb = new StringBuilder(256);
            GetClassName(h, sb, 256);
            string cls = sb.ToString();
            GetWindowText(h, sb, 256);
            string txt = sb.ToString();
            bool vis = IsWindowVisible(h);
            list.Add(new ControlInfo(h, cls, txt, vis));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    // =========================================================================
    // RunWorkflow — 행 선택 → 필드 로드 대기 → 값 입력 → 분석완료 클릭
    // =========================================================================
    // 항목 순서: BOD, TOC, SS, NH, PN, TN, TP
    private static readonly string[] AnalyteLabels =
        ["BOD 측정치", "TOC 측정치", "SS 측정치", "NH 측정치", "PN 측정치", "TN 측정치", "TP 측정치"];

    // ── 캐시: 첫 호출에서 EDIT/버튼 HWND 매핑 후 재사용 ──────────────────
    private static List<(string label, IntPtr editHwnd)>? _cachedMapping;
    private static IntPtr _cachedDoneBtn;
    private static IntPtr _cachedGridHwnd;
    private static IntPtr _cachedMainHwnd;

    /// <summary>캐시 초기화 (새 자동화 세션 시작 시 호출)</summary>
    public static void ResetCache() { _cachedMapping = null; _cachedDoneBtn = IntPtr.Zero; _cachedGridHwnd = IntPtr.Zero; }

    /// <param name="values">입력값 배열 (BOD,TOC,SS,NH,PN,TN,TP 순서)</param>
    /// <param name="rowIndex">클릭할 그리드 행 번호 (1-based)</param>
    /// <param name="expectedSN">Excel S/N — ERP 현재 행과 매칭 확인용 (빈 문자열이면 건너뜀)</param>
    /// <param name="navigateFromStart">true=Ctrl+Home부터 전체 네비게이션, false=현재 위치 유지</param>
    public static List<(string Icon, string Message)> RunWorkflow(string[] values, int rowIndex, string expectedSN = "", bool navigateFromStart = true)
    {
        var lines = new List<(string, string)>();
        void L(string icon, string msg) { lines.Add((icon, msg)); Log($"{icon} {msg}"); }

        Log($"=== RunWorkflow(row={rowIndex}, nav={navigateFromStart}) ===");

        var procs = Process.GetProcessesByName(ProcessName);
        if (procs.Length == 0) { L("❌", "프로세스 없음"); return lines; }

        var hwndMain = procs[0].MainWindowHandle;
        _cachedMainHwnd = hwndMain;
        SetForegroundWindow(hwndMain);

        // ── 첫 호출: 그리드/EDIT/버튼 탐색 + 캐시 ──
        if (_cachedMapping == null || navigateFromStart)
        {
            var all = EnumAllChildren(hwndMain);
            var rects = new Dictionary<IntPtr, RECT>();
            foreach (var c in all)
                if (GetWindowRect(c.Hwnd, out RECT r)) rects[c.Hwnd] = r;

            // 그리드 탐색
            ControlInfo? grid = null;
            long bestArea = 0;
            foreach (var c in all.Where(c => c.Visible &&
                c.ClassName.StartsWith("WindowsForms10.Window.8", StringComparison.OrdinalIgnoreCase)))
            {
                if (!rects.TryGetValue(c.Hwnd, out RECT gr)) continue;
                long area = (long)(gr.Right - gr.Left) * (gr.Bottom - gr.Top);
                if (area > bestArea) { bestArea = area; grid = c; }
            }
            if (grid == null) { L("⚠️", "그리드 컨트롤 미발견"); return lines; }
            _cachedGridHwnd = grid.Hwnd;

            // EDIT 매핑
            var edits = all.FindAll(c =>
                c.ClassName.StartsWith("WindowsForms10.EDIT", StringComparison.OrdinalIgnoreCase) && c.Visible);
            var mapped = new List<(string label, IntPtr editHwnd)>();
            foreach (var label in AnalyteLabels)
            {
                var lbl = all.Find(c =>
                    c.ClassName.StartsWith("WindowsForms10.STATIC", StringComparison.OrdinalIgnoreCase) &&
                    c.Text == label && c.Visible);
                if (lbl == null || !rects.TryGetValue(lbl.Hwnd, out RECT lr))
                { mapped.Add((label, IntPtr.Zero)); continue; }

                int lblCx = (lr.Left + lr.Right) / 2, lblCy = (lr.Top + lr.Bottom) / 2;
                IntPtr best = IntPtr.Zero; double bestD = double.MaxValue;
                foreach (var e in edits)
                {
                    if (!rects.TryGetValue(e.Hwnd, out RECT er)) continue;
                    int ex = (er.Left + er.Right) / 2, ey = (er.Top + er.Bottom) / 2;
                    if (ex < lr.Left - 10 && ey < lr.Top - 10) continue;
                    double d = Math.Sqrt(Math.Pow(ex - lblCx, 2) + Math.Pow(ey - lblCy, 2));
                    if (d < bestD) { bestD = d; best = e.Hwnd; }
                }
                mapped.Add((label, best));
                Log($"  매핑: '{label}' → 0x{best:X8}  거리={bestD:F0}");
            }
            _cachedMapping = mapped;

            // 버튼 캐시
            var btns = all.FindAll(c =>
                c.ClassName.StartsWith("WindowsForms10.BUTTON", StringComparison.OrdinalIgnoreCase) && c.Visible);
            var doneBtn = btns.Find(b => b.Text.Contains("완료") || b.Text.Contains("저장"));
            _cachedDoneBtn = doneBtn?.Hwnd ?? IntPtr.Zero;

            L("✅", $"컨트롤 매핑 완료 (grid=0x{_cachedGridHwnd:X8}, edits={mapped.Count(m => m.editHwnd != IntPtr.Zero)}, btn=0x{_cachedDoneBtn:X8})");
        }

        // ② 그리드 행 선택
        if (navigateFromStart)
        {
            EnsureGridFocus(hwndMain, _cachedGridHwnd);
            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            keybd_event(VK_HOME,    0, 0, UIntPtr.Zero);
            keybd_event(VK_HOME,    0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            Thread.Sleep(30);

            for (int i = 0; i < rowIndex - 1; i++)
            {
                keybd_event(VK_DOWN, 0, 0, UIntPtr.Zero);
                keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                if (i % 20 == 19) Thread.Sleep(15);
            }
            Thread.Sleep(80);
            L("✅", $"행 선택 (처음부터)  row={rowIndex}");
        }
        else
        {
            // 이전 Down으로 이미 다음 행 — 필드 로드 대기만
            Thread.Sleep(100);
        }

        // ④ 값 입력 (캐시된 EDIT HWND 사용)
        int inputN = Math.Min(values.Length, _cachedMapping.Count);
        for (int i = 0; i < inputN; i++)
        {
            var (label, hwnd) = _cachedMapping[i];
            if (hwnd == IntPtr.Zero) continue;
            TypeText(hwnd, values[i]);
        }
        L("✅", $"값 입력 완료 ({inputN}개)");

        // ⑤ 분석완료 버튼 클릭
        if (_cachedDoneBtn != IntPtr.Zero)
        {
            SendMessage(_cachedDoneBtn, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            L("✅", $"분석완료 클릭");

            Thread.Sleep(200);
            EnsureGridFocus(hwndMain, _cachedGridHwnd);
            keybd_event(VK_DOWN, 0, 0, UIntPtr.Zero);
            keybd_event(VK_DOWN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            Thread.Sleep(50);
            L("✅", $"다음 행 이동");
        }
        else
        {
            L("⚠️", "분석완료 버튼 캐시 없음");
        }

        Log("=== RunWorkflow 완료 ===");
        return lines;
    }

    private static IntPtr MakeLParam(int x, int y) =>
        (IntPtr)(((uint)y << 16) | ((uint)x & 0xFFFF));

    /// <summary>SendMessage WM_CHAR로 대상 HWND에 직접 키 입력. 포커스 불필요.</summary>
    private static void TypeText(IntPtr hwnd, string text)
    {
        SendMessage(hwnd, EM_SETSEL, IntPtr.Zero, (IntPtr)(-1));
        SendMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_DELETE, (IntPtr)0x00000001);
        foreach (char c in text)
            SendMessage(hwnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
    }

    /// <summary>
    /// 그리드에 포커스를 확실히 설정 (keybd_event 전 호출).
    /// OS-level mouse_event로 그리드 중앙을 실제 클릭하여 DevExpress 내부 포커스를 강제 설정.
    /// </summary>
    private static void EnsureGridFocus(IntPtr hwndMain, IntPtr gridHwnd)
    {
        SetForegroundWindow(hwndMain);
        Thread.Sleep(20);

        if (GetWindowRect(gridHwnd, out RECT gr))
        {
            int screenX = (gr.Left + gr.Right) / 2;
            int screenY = gr.Top + 40;
            if (screenY > gr.Bottom - 5) screenY = (gr.Top + gr.Bottom) / 2;

            SetCursorPos(screenX, screenY);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(10);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(30);
        }
        else
        {
            Log($"  EnsureGridFocus: GetWindowRect 실패 grid=0x{gridHwnd:X8}");
        }
    }

    // 숫자패드 VK 코드 (0x60~0x6E)
    private static byte CharToVk(char c) => c switch
    {
        '0' => 0x60, '1' => 0x61, '2' => 0x62, '3' => 0x63, '4' => 0x64,
        '5' => 0x65, '6' => 0x66, '7' => 0x67, '8' => 0x68, '9' => 0x69,
        '.' => 0x6E, '-' => 0x6D, _   => 0,
    };

    // WM_CHAR용 ASCII 코드
    private static byte CharToAscii(char c) => (byte)c;  // '0'→0x30, '.'→0x2E, '-'→0x2D

    // =========================================================================
    // 테스트 입력 — 상세 진단
    // =========================================================================
    public static List<(string Icon, string Message)> TestInput(string value)
    {
        var lines = new List<(string, string)>();
        void L(string icon, string msg) { lines.Add((icon, msg)); Log($"{icon} {msg}"); }

        Log($"=== TestInput('{value}') 시작 ===");

        var procs = Process.GetProcessesByName(ProcessName);
        if (procs.Length == 0)
        {
            L("❌", $"'{ProcessName}.exe' 프로세스 없음"); return lines;
        }
        var proc = procs[0];
        var hwndMain = proc.MainWindowHandle;
        L("✅", $"프로세스 PID={proc.Id}  HWND=0x{hwndMain:X8}");

        SetForegroundWindow(hwndMain);
        BringWindowToTop(hwndMain);

        var all = EnumAllChildren(hwndMain);
        L("──", $"자식 컨트롤 총 {all.Count}개");

        var allEdits = all.FindAll(c => c.ClassName.StartsWith("WindowsForms10.EDIT", StringComparison.OrdinalIgnoreCase));
        L("──", $"WindowsForms10.EDIT 총 {allEdits.Count}개 (가시: {allEdits.Count(e => e.Visible)})");

        var visEdits = allEdits.FindAll(e => e.Visible);
        if (visEdits.Count == 0)
        {
            L("⚠️", "보이는 EDIT 없음"); return lines;
        }

        // 첫 번째 가시 EDIT에 keybd_event로 입력
        var rects = new Dictionary<IntPtr, RECT>();
        foreach (var c in all)
            if (GetWindowRect(c.Hwnd, out RECT rc)) rects[c.Hwnd] = rc;

        TypeText(visEdits[0].Hwnd, value);
        L("✅", $"첫 번째 EDIT에 입력 완료  0x{visEdits[0].Hwnd:X8}  값='{value}'");

        Log("=== TestInput 완료 ===");
        return lines;
    }

    // =========================================================================
    // 단순 헬퍼
    // =========================================================================
    public static List<ControlInfo> GetVisibleEditControls()
    {
        var procs = Process.GetProcessesByName(ProcessName);
        if (procs.Length == 0) return [];
        var all = EnumAllChildren(procs[0].MainWindowHandle);
        return all.FindAll(c =>
            c.ClassName.StartsWith("WindowsForms10.EDIT", StringComparison.OrdinalIgnoreCase)
            && c.Visible);
    }

    public static bool SetText(IntPtr hwnd, string value)
    {
        var r = SendMessage(hwnd, WM_SETTEXT, IntPtr.Zero, value);
        return r != IntPtr.Zero || value == "";
    }

    public static void ClickButton(IntPtr hwnd)
    {
        SendMessage(hwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
    }

    public static bool ClickDoneButton()
    {
        var procs = Process.GetProcessesByName(ProcessName);
        if (procs.Length == 0) return false;
        var all = EnumAllChildren(procs[0].MainWindowHandle);
        var btn = all.Find(c =>
            c.ClassName.StartsWith("WindowsForms10.BUTTON", StringComparison.OrdinalIgnoreCase)
            && c.Text.Contains("완료")
            && c.Visible);
        if (btn == null) return false;
        ClickButton(btn.Hwnd);
        return true;
    }

    // =========================================================================
    // ERP 그리드에서 S/N 목록 읽기
    // =========================================================================
    /// <summary>ERP 그리드의 보이는 EDIT/STATIC 컨트롤에서 S/N 패턴 텍스트를 수집</summary>
    /// <summary>
    /// ERP 그리드에서 S/N 목록을 읽는다.
    /// DevExpress XtraGrid는 owner-draw라 GetWindowText로 셀 읽기 불가 →
    /// 그리드에 포커스 후 Ctrl+A → Ctrl+C → 클립보드 텍스트 파싱.
    /// </summary>
    public static List<string> GetGridSnList()
    {
        var snList = new List<string>();
        var procs = Process.GetProcessesByName(ProcessName);
        if (procs.Length == 0) { Log("[GetGridSnList] 프로세스 없음"); return snList; }

        var hwndMain = procs[0].MainWindowHandle;
        if (hwndMain == IntPtr.Zero) { Log("[GetGridSnList] 메인 HWND 없음"); return snList; }

        // 가장 넓은 WindowsForms10.Window.8 = XtraGrid
        var all = EnumAllChildren(hwndMain);
        var rects = new Dictionary<IntPtr, RECT>();
        foreach (var c in all)
            if (GetWindowRect(c.Hwnd, out RECT r)) rects[c.Hwnd] = r;

        ControlInfo? grid = null;
        long bestArea = 0;
        foreach (var c in all.Where(c => c.Visible &&
            c.ClassName.StartsWith("WindowsForms10.Window.8", StringComparison.OrdinalIgnoreCase)))
        {
            if (!rects.TryGetValue(c.Hwnd, out RECT gr)) continue;
            long area = (long)(gr.Right - gr.Left) * (gr.Bottom - gr.Top);
            if (area > bestArea) { bestArea = area; grid = c; }
        }

        if (grid == null) { Log("[GetGridSnList] 그리드 컨트롤 미발견"); return snList; }
        Log($"[GetGridSnList] 그리드 HWND=0x{grid.Hwnd:X8}  area={bestArea}");

        // 그리드 포커스 → Ctrl+A (전체선택) → Ctrl+C (클립보드 복사)
        SetForegroundWindow(hwndMain);
        Thread.Sleep(60);
        EnsureGridFocus(hwndMain, grid.Hwnd);
        Thread.Sleep(80);

        // Ctrl+A
        keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event(0x41, 0, 0, UIntPtr.Zero);
        keybd_event(0x41, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Thread.Sleep(120);

        // Ctrl+C
        keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event(0x43, 0, 0, UIntPtr.Zero);
        keybd_event(0x43, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        Thread.Sleep(300);

        // 클립보드 읽기 (powershell Get-Clipboard 사용)
        string clipText = "";
        try
        {
            var psi = new ProcessStartInfo("powershell", "-NoProfile -Command \"Get-Clipboard\"")
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var proc = Process.Start(psi);
            if (proc != null)
            {
                clipText = proc.StandardOutput.ReadToEnd();
                proc.WaitForExit(3000);
            }
        }
        catch (Exception ex) { Log($"[GetGridSnList] 클립보드 오류: {ex.Message}"); }

        Log($"[GetGridSnList] 클립보드 {clipText.Length}자 수신");
        if (clipText.Length > 0)
            Log($"[GetGridSnList] 첫 줄: '{clipText.Split('\n')[0].Trim()}'");

        // 각 줄에서 S/N 패턴(MM-DD-숫자, 접두사 포함 가능) 추출
        var snRegex = new System.Text.RegularExpressions.Regex(@"\d{2}-\d{2}-\d+");
        foreach (var line in clipText.Split('\n'))
        {
            var m = snRegex.Match(line);
            if (m.Success && !snList.Contains(m.Value))
                snList.Add(m.Value);
        }

        Log($"[GetGridSnList] {snList.Count}개 S/N 발견");
        return snList;
    }

    // =========================================================================
    // Probe 결과 컨테이너
    // =========================================================================
    public class ProbeResult
    {
        private readonly List<(string Icon, string Message)> _lines = [];

        public void Add(string icon, string msg) => _lines.Add((icon, msg));
        public IReadOnlyList<(string Icon, string Message)> Lines => _lines;

        public bool CanAutomate =>
            _lines.Exists(l => l.Icon == "✅" && l.Message.Contains("EDIT"));
    }
}
