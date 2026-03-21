using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Markup.Xaml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.Data.Sqlite;
using ETA.Models;
using ETA.Services;

namespace ETA.Views.Pages;

public partial class WasteCompanyPage : UserControl
{
    public WasteCompanyPage()
    {
        InitializeComponent();
        LoadData();
    }

    public void LoadData()
    {
        UpdateStatus("트리 로드 중...");
        Log("LoadData() 시작");

        WasteCompanyTreeView.Items.Clear();

        try
        {
            //var Wastecompanys = WasteCompanyService.GetAllItems() ?? new List<WasteCompany>();
            var wastecompanys = WasteCompanyService.GetAllItems() ?? new List<WasteCompany>();
            //Log($"DB 로드 완료 → {agents.Count}명");

            if (wastecompanys.Count == 0)
            {
                UpdateStatus("❌ DB에 데이터가 없습니다.");
                return;
            }
            foreach (var WasteCompany in wastecompanys)
            {
                var CompanyItem = new TreeViewItem
                {
                    IsExpanded = false,
                    Tag = WasteCompany,
                    Classes = { "no-expander" }   // ←←← 이 한 줄이 핵심!
                };

                CompanyItem.Header = new TextBlock
                {
                    Text = WasteCompany.업체명 ?? "이름 없음",
                    FontSize = 13,
                    FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
                    Foreground = Brushes.WhiteSmoke,
                    Margin = new Thickness(12, 8, 0, 8),
                    VerticalAlignment = VerticalAlignment.Center
                };

                CompanyItem.Items.Add(CreateLabeledBox("프로젝트", WasteCompany.프로젝트));
                CompanyItem.Items.Add(CreateLabeledBox("프로젝트명", WasteCompany.프로젝트명));
                CompanyItem.Items.Add(CreateLabeledBox("관리번호", WasteCompany.관리번호));
                CompanyItem.Items.Add(CreateLabeledBox("사업자번호", WasteCompany.사업자번호));

                WasteCompanyTreeView.Items.Add(CompanyItem);
            }
            //UpdateStatus($"✅ {agents.Count}명 표시 완료 → [전체 저장] 버튼 누르세요");
        }
        catch (Exception ex)
        {
            Log("★ 크래시 ★ " + ex.Message);
            UpdateStatus($"❌ 오류: {ex.Message}");
        }
    }

    private StackPanel CreateLabeledBox(string label, string value, bool isReadOnly = false)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };

        panel.Children.Add(new TextBlock
        {
            Text = label + ": ",
            Width = 110,
            Foreground = Brushes.LightGray,
            FontSize = 13,
            VerticalAlignment = VerticalAlignment.Center
        });

        var textBox = new TextBox
        {
            Text = value ?? "",
            FontSize = 10,
            FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Background = Brushes.Transparent,

            BorderThickness = new Thickness(0),
            Padding = new Thickness(6, 3),
            Width = 220,
            IsReadOnly = isReadOnly
        };

        panel.Children.Add(textBox);
        return panel;
    }

    private void SaveAllButton_Click(object? sender, RoutedEventArgs e)
    {
        int success = 0;
        Log("=== 전체 저장 시작 (키 = 성명) ===");

        foreach (var item in WasteCompanyTreeView.Items)
        {
            if (item is TreeViewItem CompanyItem && CompanyItem.Tag is WasteCompany wastecompany)
            {
                foreach (var child in CompanyItem.Items)
                {
                    if (child is StackPanel panel && panel.Children.Count == 2)
                    {


                        var labelText = (panel.Children[0] as TextBlock)?.Text.Replace(": ", "").Trim() ?? "";
                        var tb = panel.Children[1] as TextBox;
                        if (tb == null) continue;

                        switch (labelText)
                        {
                            case "프로젝트": wastecompany.프로젝트 = tb.Text ?? ""; break;
                            case "프로젝트명": wastecompany.프로젝트명 = tb.Text ?? ""; break;
                            case "관리번호": wastecompany.관리번호 = tb.Text ?? ""; break;
                            case "사업자번호": wastecompany.사업자번호 = tb.Text ?? ""; break;

                        }
                    }
                }

                if (ExecuteDirectUpdate(wastecompany)) success++;
            }
        }

        UpdateStatus($"✅ 저장 완료! 성공 {success}명");
        Log($"=== 저장 종료 → {success}명 성공 ===");
    }

    private bool ExecuteDirectUpdate(WasteCompany wastecompany)
    {
        string dbPath = WasteCompanyService.GetDatabasePath();           // ← 이게 실제 사용하는 파일 경로
        string fullAbsolutePath = Path.GetFullPath(dbPath);       // ← 절대경로 (C:\Users\... 형태)

        Log($"[DB 위치 확인] 상대경로: {dbPath}");
        Log($"[DB 위치 확인] 절대경로: {fullAbsolutePath}");
        Log($"[DB 위치 확인] 파일 존재? {File.Exists(dbPath)}");
        Log($"[DB 위치 확인] 현재 파일 수정 시간: {File.GetLastWriteTime(dbPath):yyyy-MM-dd HH:mm:ss}");

        string sql = @"
        UPDATE ""폐수배출업소"" SET 
            프로젝트=@프로젝트, 프로젝트명=@프로젝트명, 관리번호=@관리번호, 사업자번호=@사업자번호
        WHERE 업체명 = @OriginalKey";

        Log($"[SQL] 실행 → {sql.Replace("\r\n", " ")}");
        Log($"[파라미터] Original업체명 = '{wastecompany.Original업체명}' | 새 업체 = '{wastecompany.업체명}'");

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;

        cmd.Parameters.AddWithValue("@프로젝트", wastecompany.프로젝트 ?? "");
        cmd.Parameters.AddWithValue("@프로젝트명", wastecompany.프로젝트명 ?? "");
        cmd.Parameters.AddWithValue("@관리번호", wastecompany.관리번호 ?? "");
        cmd.Parameters.AddWithValue("@사업자번호", wastecompany.사업자번호 ?? "");
        int rows = cmd.ExecuteNonQuery();

        string afterTime = File.GetLastWriteTime(dbPath).ToString("yyyy-MM-dd HH:mm:ss");
        Log($"[SQL 결과] {rows}행 업데이트");
        Log($"[DB 위치 확인] 업데이트 후 파일 수정 시간: {afterTime}");

        if (rows > 0)
        {
            wastecompany.Original업체명 = wastecompany.업체명;
            return true;
        }
        return false;
    }

    private void Log(string message)
    {
        string log = $"[{DateTime.Now:HH:mm:ss}] {message}";
        Debug.WriteLine(log);
        try { File.AppendAllText("AgentDebug.log", log + Environment.NewLine); } catch { }
    }

    private void UpdateStatus(string text)
    {
        //if (StatusText != null) StatusText.Text = text;
    }
}