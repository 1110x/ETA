using Avalonia;
using ETA.Views;
using Avalonia.Controls;
using ETA.Views;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services;
using ETA.Services.SERVICE1;
using ETA.Services.SERVICE2;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ETA.Views;

/// <summary>
/// 최초 실행 시 필요한 패키지 설치를 안내하는 다이얼로그.
/// App.axaml.cs 의 OnFrameworkInitializationCompleted() 에서 호출.
/// </summary>
public class InstallCheckDialog : Window
{
    private static Brush AppRes(string key, string fallback = "#888888")
    {
        if (Application.Current?.Resources.TryGetResource(key, null, out var v) == true && v is Brush b) return b;
        return new SolidColorBrush(Color.Parse(fallback));
    }

    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    private readonly StackPanel _itemPanel  = new() { Spacing = 6 };
    private readonly TextBlock  _statusText = new()
    {
        FontFamily = new FontFamily("avares://ETA/Assets/Fonts#Pretendard"),
        FontSize   = AppTheme.FontBase,
        Foreground = Brushes.Gray,
        Margin     = new Avalonia.Thickness(0, 4, 0, 0),
    };
    private readonly Button _installBtn = new()
    {
        Content    = "설치 시작",
        Width      = 100,
        Height     = 32,
        Background = new SolidColorBrush(Color.Parse("#2a4a2a")),
        Foreground = Brushes.White,
        BorderThickness = new Avalonia.Thickness(0),
    };
    private readonly Button _skipBtn = new()
    {
        Content    = "나중에",
        Width      = 80,
        Height     = 32,
        Background = AppRes("SubBtnBg"),
        Foreground = Brushes.White,
        BorderThickness = new Avalonia.Thickness(0),
    };

    private List<InstallItem> _missing = new();

    public InstallCheckDialog()
    {
        Title  = "ETA — 초기 설정";
        Width  = 460;
        Height = 340;
        CanResize              = false;
        WindowStartupLocation  = WindowStartupLocation.CenterScreen;
        Background             = AppRes("PanelBg");

        BuildUI();
    }

    private void BuildUI()
    {
        _missing = AppInstaller.GetMissingPackages();

        var titleText = new TextBlock
        {
            Text       = "📦  초기 설정",
            FontFamily = Font,
            FontSize   = AppTheme.FontXL,
            FontWeight = FontWeight.Bold,
            Foreground = AppRes("AppFg"),
            Margin     = new Avalonia.Thickness(0, 0, 0, 6),
        };

        var subText = new TextBlock
        {
            Text       = "아래 항목이 설치되어 있지 않습니다.\n지금 설치하시겠습니까?",
            FontFamily = Font,
            FontSize   = AppTheme.FontMD,
            Foreground = AppRes("FgMuted"),
            Margin     = new Avalonia.Thickness(0, 0, 0, 12),
        };

        // 설치 항목 목록
        foreach (var item in _missing)
        {
            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing     = 10,
                Children    =
                {
                    new TextBlock
                    {
                        Text       = item.Required ? "🔴" : "🟡",
                        FontSize   = AppTheme.FontXL,
                        VerticalAlignment = VerticalAlignment.Center,
                    },
                    new StackPanel
                    {
                        Orientation = Orientation.Vertical,
                        Spacing     = 2,
                        Children    =
                        {
                            new TextBlock
                            {
                                Text       = item.Name,
                                FontFamily = Font,
                                FontSize   = AppTheme.FontLG,
                                FontWeight = FontWeight.SemiBold,
                                Foreground = AppRes("AppFg"),
                            },
                            new TextBlock
                            {
                                Text       = item.Description,
                                FontFamily = Font,
                                FontSize   = AppTheme.FontBase,
                                Foreground = AppRes("FgMuted"),
                            },
                        }
                    }
                }
            };
            _itemPanel.Children.Add(new Border
            {
                Background    = AppRes("PanelInnerBg"),
                CornerRadius  = new Avalonia.CornerRadius(4),
                Padding       = new Avalonia.Thickness(10, 8),
                Child         = row,
            });
        }

        // 버튼
        _installBtn.Click += async (_, _) => await StartInstall();
        _skipBtn.Click    += (_, _) =>
        {
            AppInstaller.MarkInstallDone();  // 다시 묻지 않음
            Close();
        };

        var btnRow = new StackPanel
        {
            Orientation         = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing             = 8,
            Margin              = new Avalonia.Thickness(0, 12, 0, 0),
            Children            = { _skipBtn, _installBtn },
        };

        Content = new Border
        {
            Padding = new Avalonia.Thickness(24),
            Child   = new StackPanel
            {
                Spacing  = 0,
                Children =
                {
                    titleText,
                    subText,
                    _itemPanel,
                    _statusText,
                    btnRow,
                }
            }
        };
    }

    private async Task StartInstall()
    {
        _installBtn.IsEnabled = false;
        _skipBtn.IsEnabled    = false;
        _statusText.Foreground = new SolidColorBrush(Color.Parse("#aaaaee"));

        var results = await AppInstaller.InstallAllAsync(
            items:      _missing,
            onProgress: p =>
            {
                // UI 스레드에서 업데이트
                Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                {
                    _statusText.Text =
                        $"[{p.Current}/{p.Total}] {p.CurrentItem} — {p.Status}";
                });
            },
            silent: true);

        // 결과 요약
        int ok   = results.Count(r => r.Success && !r.Skipped);
        int fail = results.Count(r => !r.Success);

        Avalonia.Threading.Dispatcher.UIThread.Post(() =>
        {
            if (fail == 0)
            {
                _statusText.Text       = $"✅ {ok}개 설치 완료";
                _statusText.Foreground = AppTheme.FgSuccess;
            }
            else
            {
                _statusText.Text       = $"✅ {ok}개 완료  ❌ {fail}개 실패 (수동 설치 필요)";
                _statusText.Foreground = new SolidColorBrush(Color.Parse("#cc8844"));
            }

            _skipBtn.Content   = "닫기";
            _skipBtn.IsEnabled = true;
        });
    }

    // ── 정적 헬퍼: 최초 실행 시 필요하면 다이얼로그 표시 ────────────────────
    /// <summary>
    /// App.axaml.cs 의 OnFrameworkInitializationCompleted() 에서 호출.
    /// 설치 필요 항목이 있을 때만 다이얼로그를 표시한다.
    /// </summary>
    public static async Task CheckAndShowAsync(Window? owner = null)
    {
        // Required=false 패키지(LibreOffice 등)는 설치 확인 스킵
        var missing = AppInstaller.GetMissingPackages().Where(p => p.Required).ToList();
        if (missing.Count == 0)
        {
            // 설치 필요 없음 — 마커만 기록
            AppInstaller.MarkInstallDone();
            return;
        }

        var dlg = new InstallCheckDialog();
        if (owner != null)
            await dlg.ShowDialog(owner);
        else
            dlg.Show();
    }
}
