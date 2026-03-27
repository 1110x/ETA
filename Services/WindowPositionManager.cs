using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ETA.Services;

/// <summary>
/// 각 페이지(ActivePage1-4)의 창 위치, 크기, 레이아웃을 저장/복원하는 서비스
/// - 사용자별 설정 분리 (Windows user 이름 사용)
/// - PageHW.log 파일 관리
/// - 모드별 레이아웃 정보 저장 (Grid 비율, ContentControl 크기 등)
/// </summary>
public class WindowPositionManager
{
    private static readonly string LogFileName = "PageHW.log";
    private readonly string _logFilePath;
    private Dictionary<string, PageLayoutInfo> _layouts = new();

    /// <summary>
    /// 기본 생성자: 현재 Windows 사용자 자동 감지
    /// </summary>
    public WindowPositionManager() : this(GetCurrentWindowsUser()) { }

    /// <summary>
    /// 지정된 사용자 ID로 초기화
    /// </summary>
    public WindowPositionManager(string currentUserId)
    {
        if (string.IsNullOrWhiteSpace(currentUserId))
            currentUserId = "DefaultUser";

        // 사용자별 로그 파일 경로: AppData/ETA/Users/{UserId}/PageHW.log
        string appDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
            "ETA");
        string userFolder = Path.Combine(appDataFolder, "Users", currentUserId);
        
        if (!Directory.Exists(userFolder))
            Directory.CreateDirectory(userFolder);
        
        _logFilePath = Path.Combine(userFolder, LogFileName);
        
        System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 사용자: {currentUserId}, 로그 경로: {_logFilePath}");
        LoadLayouts();
    }

    /// <summary>
    /// 현재 Windows 사용자명 조회
    /// </summary>
    private static string GetCurrentWindowsUser()
    {
        try
        {
            return Environment.UserName ?? "DefaultUser";
        }
        catch
        {
            return "DefaultUser";
        }
    }

    /// <summary>
    /// 로그 파일에서 저장된 레이아웃들을 로드
    /// </summary>
    private void LoadLayouts()
    {
        try
        {
            if (!File.Exists(_logFilePath))
            {
                _layouts = new Dictionary<string, PageLayoutInfo>();
                return;
            }

            string json = File.ReadAllText(_logFilePath);
            var data = JsonSerializer.Deserialize<Dictionary<string, PageLayoutInfo>>(json);
            _layouts = data ?? new Dictionary<string, PageLayoutInfo>();
            System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 로드됨: {_layouts.Count}개 레이아웃");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 로드 오류: {ex.Message}");
            _layouts = new Dictionary<string, PageLayoutInfo>();
        }
    }

    /// <summary>
    /// 현재 레이아웃들을 로그 파일에 저장
    /// </summary>
    public void SaveLayouts()
    {
        try
        {
            string? directory = Path.GetDirectoryName(_logFilePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            var options = new JsonSerializerOptions 
            { 
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            };
            string json = JsonSerializer.Serialize(_layouts, options);
            File.WriteAllText(_logFilePath, json);
            
            System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 저장 완료: {_logFilePath}");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 저장 오류: {ex.Message}");
        }
    }

    /// <summary>
    /// 특정 페이지 모드의 레이아웃 정보 저장
    /// </summary>
    public void SavePageLayout(string pageName, PageLayoutInfo layoutInfo)
    {
        _layouts[pageName] = layoutInfo;
        SaveLayouts();
        System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 저장: {pageName} - {layoutInfo}");
    }

    /// <summary>
    /// 특정 페이지 모드의 저장된 레이아웃 조회
    /// </summary>
    public PageLayoutInfo? GetPageLayout(string pageName)
    {
        if (_layouts.TryGetValue(pageName, out var layout))
        {
            System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 조회: {pageName} - {layout}");
            return layout;
        }
        return null;
    }

    /// <summary>
    /// 모든 저장된 레이아웃 조회
    /// </summary>
    public Dictionary<string, PageLayoutInfo> GetAllLayouts() => new Dictionary<string, PageLayoutInfo>(_layouts);

    /// <summary>
    /// 특정 페이지의 레이아웃 정보 삭제
    /// </summary>
    public void RemovePageLayout(string pageName)
    {
        if (_layouts.Remove(pageName))
        {
            SaveLayouts();
            System.Diagnostics.Debug.WriteLine($"[WindowPositionManager] 삭제: {pageName}");
        }
    }

    /// <summary>
    /// 모든 레이아웃 정보 초기화
    /// </summary>
    public void ClearAllLayouts()
    {
        _layouts.Clear();
        SaveLayouts();
        System.Diagnostics.Debug.WriteLine("[WindowPositionManager] 모든 레이아웃 정보 초기화");
    }

    /// <summary>
    /// 로그 파일 경로 조회 (디버깅용)
    /// </summary>
    public string GetLogFilePath() => _logFilePath;
}

/// <summary>
/// 페이지 레이아웃 정보 (위치, 크기, 그리드 비율 등)
/// </summary>
public class PageLayoutInfo
{
    [JsonPropertyName("windowX")]
    public double WindowX { get; set; }

    [JsonPropertyName("windowY")]
    public double WindowY { get; set; }

    [JsonPropertyName("windowWidth")]
    public double WindowWidth { get; set; }

    [JsonPropertyName("windowHeight")]
    public double WindowHeight { get; set; }

    /// <summary>
    /// Grid 열 비율 (예: Content2와 Content4 분할 비율)
    /// Content2Star : Content4Star
    /// </summary>
    [JsonPropertyName("content2Star")]
    public double Content2Star { get; set; } = 1;

    [JsonPropertyName("content4Star")]
    public double Content4Star { get; set; } = 1;

    /// <summary>
    /// Grid 행 비율 (예: 상단과 하단 분할 비율)
    /// UpperStar : LowerStar
    /// </summary>
    [JsonPropertyName("upperStar")]
    public double UpperStar { get; set; } = 4;

    [JsonPropertyName("lowerStar")]
    public double LowerStar { get; set; } = 1;

    /// <summary>
    /// 왼쪽 패널(Content1) 너비
    /// </summary>
    [JsonPropertyName("leftPanelWidth")]
    public double LeftPanelWidth { get; set; } = 260;

    [JsonPropertyName("savedAt")]
    public DateTime SavedAt { get; set; } = DateTime.Now;

    public override string ToString() =>
        $"Window:({WindowX},{WindowY}) {WindowWidth}x{WindowHeight}, " +
        $"Content2Star={Content2Star}, Content4Star={Content4Star}, " +
        $"UpperStar={UpperStar}, LowerStar={LowerStar}, LeftPanelWidth={LeftPanelWidth}";
}
