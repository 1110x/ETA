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
using System.Data;
using System.Data.Common;
using ETA.Models;
using ETA.Services;
using System.Linq;

namespace ETA.Views.Pages;

public partial class WasteCompanyPage : UserControl
{
    private const string TableName = "폐수배출업소";  // SQL TABLE 이름
    private const string KeyColumn = "업체명";   // 나중에 키가 바뀌어도 여기만 수정
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
        var items = WasteCompanyService.GetAllItems() ?? new List<WasteCompany>();
        if (items.Count == 0)
        {
            UpdateStatus("❌ 데이터 없음");
            return;
        }
        foreach (var entity in items)
        {
            var node = new TreeViewItem
            {
                IsExpanded = false,
                Tag = entity,
                Classes = { "no-expander" }
            };
            node.Header = new TextBlock
            {
                Text = entity.업체명 ?? "이름 없음",
                FontSize = 13,
                FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
                Foreground = Brushes.WhiteSmoke,
                Margin = new Thickness(12, 8, 0, 8),
                VerticalAlignment = VerticalAlignment.Center
            };
            node.Items.Add(CreateFieldRow("프로젝트",     entity.프로젝트));
            node.Items.Add(CreateFieldRow("프로젝트명",   entity.프로젝트명));
            node.Items.Add(CreateFieldRow("관리번호",     entity.관리번호));
            node.Items.Add(CreateFieldRow("사업자번호",   entity.사업자번호));
            WasteCompanyTreeView.Items.Add(node);
        }

        UpdateStatus($"✅ {items.Count}건 로드 완료");
    }

    // 이름도 좀 더 일반적으로 → CreateFieldRow
    private StackPanel CreateFieldRow(string label, string? value, bool readOnly = false)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };

        panel.Children.Add(new TextBlock
        {
            Text = label + ": ",
            Width = 110,
            Foreground = Brushes.LightGray,
            FontSize = 12,
            FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            VerticalAlignment = VerticalAlignment.Center
        });

        var tb = new TextBox
        {
            Text = value ?? "",
            FontSize = 11,
            FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(6, 3),
            Width = 220,
            IsReadOnly = readOnly
        };

        panel.Children.Add(tb);
        return panel;
    }

    private void SaveAllButton_Click(object? sender, RoutedEventArgs e)
    {
        int successCount = 0;
        Log("=== 전체 저장 시작 ===");

        foreach (var node in WasteCompanyTreeView.Items.OfType<TreeViewItem>())
        {
            if (node.Tag is not WasteCompany item) continue;

            // UI → 모델 동기화
            foreach (var child in node.Items.OfType<StackPanel>())
            {
                if (child.Children.Count < 2) continue;

                var labelTb = child.Children[0] as TextBlock;
                var textBox = child.Children[1] as TextBox;

                if (labelTb is null || textBox is null) continue;

                string label = (labelTb.Text ?? "").Replace(": ", "").Trim();

                switch (label)
                {
                    case "프로젝트":     item.프로젝트   = textBox.Text ?? ""; break;
                    case "프로젝트명":   item.프로젝트명 = textBox.Text ?? ""; break;
                    case "관리번호":     item.관리번호   = textBox.Text ?? ""; break;
                    case "사업자번호":   item.사업자번호 = textBox.Text ?? ""; break;
                }
            }

            if (SaveItem(item))
                successCount++;
        }

        UpdateStatus($"저장 완료 : {successCount} / {WasteCompanyTreeView.Items.Count}");
        Log($"=== 저장 완료 → 성공 {successCount}건 ===");
    }

    private bool SaveItem(WasteCompany item)
    {
        Log($"DB 경로 : {Path.GetFullPath(DbPathHelper.DbPath)}");

        string sql = $@"
            UPDATE `{TableName}`
            SET 프로젝트=@p, 프로젝트명=@pn, 관리번호=@m, 사업자번호=@b
            WHERE {KeyColumn} = @key";

        using var conn = DbConnectionFactory.CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;

        cmd.Parameters.AddWithValue("@p",   item.프로젝트   ?? "");
        cmd.Parameters.AddWithValue("@pn",  item.프로젝트명 ?? "");
        cmd.Parameters.AddWithValue("@m",   item.관리번호   ?? "");
        cmd.Parameters.AddWithValue("@b",   item.사업자번호 ?? "");
        cmd.Parameters.AddWithValue("@key", item.Original업체명 ?? item.업체명 ?? "");

        int affected = cmd.ExecuteNonQuery();

        if (affected > 0)
        {
            item.Original업체명 = item.업체명 ?? "";  // 키 갱신
            return true;
        }

        return false;
    }

    // Log, UpdateStatus는 그대로 유지

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