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
using System.Linq;
using Microsoft.Data.Sqlite;
using ETA.Models;
using ETA.Services;

namespace ETA.Views.Pages;

public partial class AgentTreePage : UserControl
{
    // ====================== ★★★ 여기 4곳만 바꾸세요 ★★★ ======================
    // WasteCompanyPage로 복사할 때 이 4줄만 수정하면 끝!
    private readonly Func<List<Agent>> GetData = AgentService.GetAllItems;   // 1. 서비스
    private readonly string DefaultIcon = "👨‍💼";                            // 2. 기본 아이콘
    private readonly string ItemNameProperty = "성명";                       // 3. 이름 표시할 필드
    private readonly string TreeViewName = "AgentTreeView";                  // 4. XAML에 있는 TreeView x:Name
    // =========================================================================

    public AgentTreePage()
    {
        InitializeComponent();
        LoadData();
    }

    public void LoadData()
    {
        UpdateStatus("트리 로드 중...");
        Log("LoadData() 시작");

        var tree = this.FindControl<TreeView>(TreeViewName);
        if (tree == null) return;
        tree.Items.Clear();

        try
        {
            var items = GetData() ?? new List<Agent>();           // ← 설정한 GetData 사용
            items = items.OrderBy(a => a.입사일).ToList();

            Log($"DB 로드 완료 → {items.Count}명");

            if (items.Count == 0)
            {
                UpdateStatus("❌ DB에 데이터가 없습니다.");
                return;
            }

            foreach (var item in items)
            {
                var personItem = CreateTreeItem(item);
                tree.Items.Add(personItem);
            }

            UpdateStatus($"✅ {items.Count}명 표시 완료 → [전체 저장] 버튼 누르세요");
        }
        catch (Exception ex)
        {
            Log("★ 크래시 ★ " + ex.Message);
            UpdateStatus($"❌ 오류: {ex.Message}");
        }
    }

    // ★★★ TreeItem 생성 (설정값 자동 적용) ★★★
    private TreeViewItem CreateTreeItem(Agent item)
    {
        var personItem = new TreeViewItem { IsExpanded = false };
        personItem.Tag = item;

        // 아이콘 + 이름 (설정값 자동 적용)
        string icon = item.기타 switch
        {
            "0" => "🥷",
            "1" => "👨‍💼",
            "2" => "👩‍⚖️"
            
        };

        personItem.Header = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            VerticalAlignment = VerticalAlignment.Center,
            Children =
            {
                new TextBlock { Text = icon, FontSize = 24, VerticalAlignment = VerticalAlignment.Center },
                new TextBlock
                {
                    Text = GetProperty(item, ItemNameProperty) ?? "이름 없음",
                    FontSize = 13,
                    FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R",
                    Foreground = Brushes.WhiteSmoke,
                    VerticalAlignment = VerticalAlignment.Center
                }
            }
        };

        // 필드 목록 (여기만 필요시 수정)
        personItem.Items.Add(CreateLabeledBox("직급", item.직급));
        personItem.Items.Add(CreateLabeledBox("직무", item.직무));
        personItem.Items.Add(CreateLabeledBox("입사일", item.입사일표시, true));
        personItem.Items.Add(CreateLabeledBox("사번", item.사번));
        personItem.Items.Add(CreateLabeledBox("자격사항", item.자격사항));
        personItem.Items.Add(CreateLabeledBox("Email", item.Email));
        personItem.Items.Add(CreateLabeledBox("측정인고유번호", item.측정인고유번호));

        return personItem;
    }

    // 도우미 (건드릴 필요 없음)
    private string GetProperty(object obj, string propName)
    {
        var prop = obj.GetType().GetProperty(propName);
        return prop?.GetValue(obj)?.ToString() ?? "";
    }

    // CreateLabeledBox, SaveAllButton_Click, ExecuteDirectUpdate, Log, UpdateStatus는 기존 그대로 사용
    private StackPanel CreateLabeledBox(string label, string value, bool isReadOnly = false)
    {
        // (기존 코드 그대로)
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };
        panel.Children.Add(new TextBlock { Text = label + ": ", Width = 110, Foreground = Brushes.LightGray, FontSize = 12, FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R", VerticalAlignment = VerticalAlignment.Center });
        var textBox = new TextBox { Text = value ?? "", FontSize = 11, FontFamily = "avares://ETA/Assets/Fonts#KBIZ한마음고딕 R", Background = Brushes.Transparent, BorderThickness = new Thickness(0), Padding = new Thickness(6, 3), Width = 220, IsReadOnly = isReadOnly };
        panel.Children.Add(textBox);
        return panel;
    }
    private void SaveAllButton_Click(object? sender, RoutedEventArgs e)
    {
        int success = 0;
        Log("=== 전체 저장 시작 (키 = 성명) ===");

        foreach (var item in AgentTreeView.Items)
        {
            if (item is TreeViewItem personItem && personItem.Tag is Agent agent)
            {
                foreach (var child in personItem.Items)
                {
                    if (child is StackPanel panel && panel.Children.Count == 2)
                    {
                        var labelText = (panel.Children[0] as TextBlock)?.Text.Replace(": ", "").Trim() ?? "";
                        var tb = panel.Children[1] as TextBox;
                        if (tb == null) continue;

                        switch (labelText)
                        {
                            case "직급": agent.직급 = tb.Text ?? ""; break;
                            case "직무": agent.직무 = tb.Text ?? ""; break;
                            case "사번": agent.사번 = tb.Text ?? ""; break;
                            case "자격사항": agent.자격사항 = tb.Text ?? ""; break;
                            case "Email": agent.Email = tb.Text ?? ""; break;
                            case "측정인고유번호": agent.측정인고유번호 = tb.Text ?? ""; break;
                            case "기타": agent.기타 = tb.Text ?? ""; break;
                        }
                    }
                }

                if (ExecuteDirectUpdate(agent)) success++;
            }
        }

        UpdateStatus($"✅ 저장 완료! 성공 {success}명");
        Log($"=== 저장 종료 → {success}명 성공 ===");
    }

    private bool ExecuteDirectUpdate(Agent agent)
    {
        string dbPath = AgentService.GetDatabasePath();           // ← 이게 실제 사용하는 파일 경로
        string fullAbsolutePath = Path.GetFullPath(dbPath);       // ← 절대경로 (C:\Users\... 형태)

        Log($"[DB 위치 확인] 상대경로: {dbPath}");
        Log($"[DB 위치 확인] 절대경로: {fullAbsolutePath}");
        Log($"[DB 위치 확인] 파일 존재? {File.Exists(dbPath)}");
        Log($"[DB 위치 확인] 현재 파일 수정 시간: {File.GetLastWriteTime(dbPath):yyyy-MM-dd HH:mm:ss}");

        string sql = @"
        UPDATE ""Agent"" SET 
            성명=@성명, 직급=@직급, 직무=@직무, 입사일=@입사일,
            사번=@사번, 자격사항=@자격사항, Email=@Email,
            기타=@기타, 측정인고유번호=@측정인고유번호
        WHERE 성명 = @OriginalKey";

        Log($"[SQL] 실행 → {sql.Replace("\r\n", " ")}");
        Log($"[파라미터] Original성명 = '{agent.Original성명}' | 새 성명 = '{agent.성명}'");

        using var conn = new SqliteConnection($"Data Source={dbPath}");
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;

        cmd.Parameters.AddWithValue("@성명", agent.성명 ?? "");
        cmd.Parameters.AddWithValue("@직급", agent.직급 ?? "");
        cmd.Parameters.AddWithValue("@직무", agent.직무 ?? "");
        cmd.Parameters.AddWithValue("@입사일", agent.입사일 == DateOnly.MinValue ? DBNull.Value : agent.입사일.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("@사번", agent.사번 ?? "");
        cmd.Parameters.AddWithValue("@자격사항", agent.자격사항 ?? "");
        cmd.Parameters.AddWithValue("@Email", agent.Email ?? "");
        cmd.Parameters.AddWithValue("@기타", agent.기타 ?? "");
        cmd.Parameters.AddWithValue("@측정인고유번호", agent.측정인고유번호 ?? "");
        cmd.Parameters.AddWithValue("@OriginalKey", agent.Original성명);

        int rows = cmd.ExecuteNonQuery();

        string afterTime = File.GetLastWriteTime(dbPath).ToString("yyyy-MM-dd HH:mm:ss");
        Log($"[SQL 결과] {rows}행 업데이트");
        Log($"[DB 위치 확인] 업데이트 후 파일 수정 시간: {afterTime}");

        if (rows > 0)
        {
            agent.Original성명 = agent.성명;
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