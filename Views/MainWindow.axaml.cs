using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Microsoft.Data.Sqlite;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views;

public partial class MainWindow : Window
{
    private static IPublicClientApplication? _pca;
    private static GraphServiceClient? _graphClient;
    private static readonly string[] Scopes = { "https://graph.microsoft.com/Tasks.ReadWrite", "offline_access" };

    private static readonly string CacheFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "ETA",
        "msalcache.bin");

    private static string? _etaListId;
    private static string? _analysisTaskId;

    public MainWindow()
    {
        InitializeComponent();

#if DEBUG
        this.AttachDevTools();
#endif

        _ = InitializeGraphAsync();

        LoadItems();
        Tree.SelectionChanged += Tree_SelectionChanged;
    }

    private async Task InitializeGraphAsync()
    {
        if (_graphClient != null) return;

        _pca = PublicClientApplicationBuilder
            .Create("34db47cd-deb9-4217-8677-15a05af5199e")
            .WithAuthority(AadAuthorityAudience.AzureAdAndPersonalMicrosoftAccount)
            .WithRedirectUri("http://localhost")
            .Build();

        EnableTokenCacheSerialization(_pca.UserTokenCache);

        var accounts = await _pca.GetAccountsAsync();
        AuthenticationResult? authResult = null;

        try
        {
            authResult = await _pca.AcquireTokenSilent(Scopes, accounts.FirstOrDefault())
                .ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            try
            {
                authResult = await _pca.AcquireTokenInteractive(Scopes)
                    .WithPrompt(Microsoft.Identity.Client.Prompt.SelectAccount)
                    .ExecuteAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"로그인 실패: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"토큰 획득 실패: {ex.Message}");
            return;
        }

        _graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
            async requestMessage =>
            {
                requestMessage.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue(
                    "Bearer", authResult.AccessToken);
                await Task.CompletedTask;
            }));
    }

    private static void EnableTokenCacheSerialization(ITokenCache tokenCache)
    {
        tokenCache.SetBeforeAccess(notificationArgs =>
        {
            if (System.IO.File.Exists(CacheFilePath))
            {
                notificationArgs.TokenCache.DeserializeMsalV3(System.IO.File.ReadAllBytes(CacheFilePath));
            }
        });

        tokenCache.SetAfterAccess(notificationArgs =>
        {
            if (notificationArgs.HasStateChanged)
            {
                System.IO.Directory.CreateDirectory(Path.GetDirectoryName(CacheFilePath)!);
                System.IO.File.WriteAllBytes(CacheFilePath, notificationArgs.TokenCache.SerializeMsalV3());
            }
        });
    }

    private async Task<string> GetOrCreateEtaListIdAsync()
    {
        if (!string.IsNullOrEmpty(_etaListId)) return _etaListId!;

        var listsPage = await _graphClient!.Me.Todo.Lists.Request().GetAsync();
        var etaList = listsPage?.CurrentPage?.FirstOrDefault(l => string.Equals(l.DisplayName, "ETA", StringComparison.OrdinalIgnoreCase));

        if (etaList?.Id != null)
        {
            _etaListId = etaList.Id;
            return _etaListId;
        }

        var newList = new TodoTaskList { DisplayName = "ETA" };
        var created = await _graphClient.Me.Todo.Lists.Request().AddAsync(newList);
        _etaListId = created?.Id ?? throw new Exception("ETA 리스트 생성 실패");
        return _etaListId;
    }

    private async Task<string> GetOrCreateAnalysisTaskIdAsync(string listId)
    {
        if (!string.IsNullOrEmpty(_analysisTaskId)) return _analysisTaskId!;

        var tasksPage = await _graphClient!.Me.Todo.Lists[listId].Tasks.Request().GetAsync();
        var analysisTask = tasksPage?.CurrentPage?.FirstOrDefault(t => string.Equals(t.Title, "분석", StringComparison.OrdinalIgnoreCase));

        if (analysisTask?.Id != null)
        {
            _analysisTaskId = analysisTask.Id;
            Console.WriteLine($"분석 태스크 재사용: {_analysisTaskId}");
            return _analysisTaskId;
        }

        var newTask = new TodoTask { Title = "분석" };
        var createdTask = await _graphClient.Me.Todo.Lists[listId].Tasks.Request().AddAsync(newTask);
        _analysisTaskId = createdTask?.Id ?? throw new Exception("분석 태스크 생성 실패");
        Console.WriteLine($"분석 태스크 생성: {_analysisTaskId}");
        return _analysisTaskId;
    }

    private async Task AddAnalysisStepAsync(string stepTitle, string taskId)
    {
        var checklistItem = new ChecklistItem
        {
            DisplayName = stepTitle,  // ← Title → DisplayName으로 수정 (API 스펙)
            IsChecked = false
        };

        await _graphClient!.Me.Todo.Lists[_etaListId!].Tasks[taskId].ChecklistItems.Request().AddAsync(checklistItem);
        Console.WriteLine($"Step 추가: {stepTitle} (분석 아래)");
    }

    private async void BtnCreateTodo_Click(object? sender, RoutedEventArgs e)
    {
        if (_graphClient == null)
        {
            Console.WriteLine("Graph 초기화 대기 중...");
            return;
        }

        string listId;
        string taskId;

        try
        {
            listId = await GetOrCreateEtaListIdAsync();
            taskId = await GetOrCreateAnalysisTaskIdAsync(listId);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ETA/분석 구조 오류: {ex.Message}");
            return;
        }

        foreach (var item in SelectedList.Items)
        {
            var text = item?.ToString();
            if (string.IsNullOrEmpty(text)) continue;

            string stepTitle = $"{text} 분석";
            await AddAnalysisStepAsync(stepTitle, taskId);
        }
    }

    // 누락된 메서드들 전체 추가
    private void Tree_SelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (Tree.SelectedItem is TreeViewItem item)
        {
            if (item.Header is string text)
            {
                // 부모 노드 제외
                if (item.ItemsSource != null && (item.ItemsSource as IEnumerable<object>)?.Any() == true)
                    return;

                if (!SelectedList.Items.Contains(text))
                {
                    SelectedList.Items.Add(text);
                }
            }
        }
    }

    private void LoadItems()
    {
        var rootItems = new List<TreeViewItem>();
        var categories = new Dictionary<string, TreeViewItem>();

        using var conn = new SqliteConnection("Data Source=Data/eta.db");
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Category, Analyte FROM AnalyteInfo ORDER BY Category";

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            string category = reader.GetString(0);
            string analyte = reader.GetString(1);

            if (!categories.TryGetValue(category, out var parent))
            {
                var children = new List<TreeViewItem>();
                parent = new TreeViewItem
                {
                    Header = category,
                    ItemsSource = children,
                    IsExpanded = true
                };
                categories[category] = parent;
                rootItems.Add(parent);
            }

            var child = new TreeViewItem { Header = analyte };
            (parent.ItemsSource as List<TreeViewItem>)?.Add(child);
        }

        Tree.ItemsSource = rootItems;
    }
}