using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ClosedXML.Excel;
using ETA.Views;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ETA.Services.Common;

/// <summary>
/// DB 업데이트 직전 스냅샷 저장:
///   1) 파싱 결과를 xlsm 템플릿(Data/Templates/PreDbSnapshot_template.xlsm)의
///      "자료시트"에 덤프 → Data/Exports/PreDbSnapshots/YYYY-MM-DD/ 저장
///   2) "인쇄하시겠습니까?" 모달 표시 (예: 기본앱으로 오픈 / 아니오: 스킵)
///
/// 용도 필터 · 스핀버튼 · 기록부 시트 생성은 xlsm 내부 VBA가 처리.
/// 템플릿 없으면 조용히 스킵 (DB 업데이트 흐름은 절대 막지 않음).
/// </summary>
public static class PreDbSnapshotService
{
    private const string DataSheetName = "자료시트";
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private static string TemplatePath => TemplateConfiguration.Resolve("PreDbSnapshot");

    private static string OutputDir(DateTime d) => Path.Combine(
        Directory.GetCurrentDirectory(), "Data", "Exports", "PreDbSnapshots",
        d.ToString("yyyy-MM-dd"));

    /// <summary>
    /// 스냅샷 저장 + 인쇄여부 모달. 예외·템플릿 부재 시 조용히 false 반환.
    /// </summary>
    /// <param name="parent">모달 부모 창 (null이면 모달 생략)</param>
    /// <param name="context">파일명 접두(예: "폐수_BOD_20260419")</param>
    /// <param name="headers">1행 헤더</param>
    /// <param name="rows">2행부터 기록할 데이터 (각 object[] = 한 행)</param>
    public static async Task<bool> SaveAndAskAsync(
        Window? parent,
        string context,
        IReadOnlyList<string> headers,
        IReadOnlyList<object?[]> rows)
    {
        if (headers.Count == 0 || rows.Count == 0) return false;

        if (!File.Exists(TemplatePath))
        {
            Debug.WriteLine($"[PreDbSnapshot] 템플릿 없음 — 스킵: {TemplatePath}");
            return false;
        }

        string outPath;
        try
        {
            var dir = OutputDir(DateTime.Today);
            Directory.CreateDirectory(dir);
            var safe = SanitizeFileName(context);
            outPath = Path.Combine(dir, $"{safe}_{DateTime.Now:HHmmss}.xlsm");

            File.Copy(TemplatePath, outPath, overwrite: true);
            WriteDataSheet(outPath, headers, rows);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[PreDbSnapshot] 저장 오류: {ex.Message}");
            return false;
        }

        if (parent == null) return true;

        bool print = await AskPrintAsync(parent, outPath, rows.Count);
        if (print)
        {
            try { OpenDefault(outPath); }
            catch (Exception ex) { Debug.WriteLine($"[PreDbSnapshot] 오픈 오류: {ex.Message}"); }
        }
        return true;
    }

    private static void WriteDataSheet(
        string path, IReadOnlyList<string> headers, IReadOnlyList<object?[]> rows)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.Contains(DataSheetName)
            ? wb.Worksheet(DataSheetName)
            : wb.AddWorksheet(DataSheetName);

        ws.Clear();
        for (int c = 0; c < headers.Count; c++)
            ws.Cell(1, c + 1).Value = headers[c];

        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            for (int c = 0; c < row.Length && c < headers.Count; c++)
            {
                var v = row[c];
                if (v == null) continue;
                var cell = ws.Cell(r + 2, c + 1);
                switch (v)
                {
                    case string s: cell.Value = s; break;
                    case double d: cell.Value = d; break;
                    case int    i: cell.Value = i; break;
                    case bool   b: cell.Value = b; break;
                    case DateTime dt: cell.Value = dt; break;
                    default:       cell.Value = v.ToString(); break;
                }
            }
        }
        wb.Save();
    }

    private static async Task<bool> AskPrintAsync(Window parent, string filePath, int rowCount)
    {
        bool result = false;
        var title = new TextBlock
        {
            Text       = "📄 시험기록부 스냅샷 저장 완료",
            FontFamily = Font, FontSize = AppTheme.FontLG,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        };
        var msg1 = new TextBlock
        {
            Text       = $"{rowCount}건 저장 → {Path.GetFileName(filePath)}",
            FontFamily = Font, FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgSecondary,
        };
        var msg2 = new TextBlock
        {
            Text       = "엑셀에서 열어 기록부를 미리보기/인쇄하시겠습니까?\n(선택과 무관하게 DB 업데이트는 계속 진행됩니다)",
            FontFamily = Font, FontSize = AppTheme.FontSM,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
        };

        var noBtn  = new Button { Content = "아니오", Padding = new Thickness(16, 6), Margin = new Thickness(0, 0, 8, 0) };
        var yesBtn = new Button
        {
            Content    = "예 (엑셀로 열기)",
            Padding    = new Thickness(16, 6),
            Background = AppTheme.BgActiveGreen,
            Foreground = AppTheme.FgSuccess,
        };
        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
        };
        btnPanel.Children.Add(noBtn);
        btnPanel.Children.Add(yesBtn);

        var content = new StackPanel { Spacing = 10, Margin = new Thickness(20) };
        content.Children.Add(title);
        content.Children.Add(msg1);
        content.Children.Add(msg2);
        content.Children.Add(btnPanel);

        var dialog = new Window
        {
            Title = "기록부 미리보기",
            Content = content,
            Width = 460, Height = 210,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false,
        };
        noBtn.Click  += (_, _) => { result = false; dialog.Close(); };
        yesBtn.Click += (_, _) => { result = true;  dialog.Close(); };

        await dialog.ShowDialog(parent);
        return result;
    }

    private static void OpenDefault(string path)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            Process.Start("open", $"\"{path}\"");
        else
            Process.Start("xdg-open", $"\"{path}\"");
    }

    private static string SanitizeFileName(string s)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var clean = new string(s.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        return string.IsNullOrWhiteSpace(clean) ? "snapshot" : clean;
    }
}
