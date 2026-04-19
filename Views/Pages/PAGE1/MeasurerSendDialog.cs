using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using ETA.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.PAGE1;

public sealed class MeasurerSendDialog : Window
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    public bool Confirmed { get; private set; }
    public List<string> PurposeValues { get; private set; } = new();
    public List<List<string>> EmpIdsPerRecord { get; private set; } = new();
    public List<string> SelectedAgentNames { get; private set; } = new();
    public List<(string name, string empId)> SelectedAgents { get; private set; } = new();

    private List<CheckBox> _agentChecks = new();
    private List<(string name, string empId)> _agentInfo = new();
    private int _recordCount;
    private StackPanel _selectionPanel;

    public MeasurerSendDialog(
        IReadOnlyList<(string sample, string analytes, string company)> records,
        List<Agent> agents,
        IReadOnlyList<string>? preSelectedNames = null)
    {
        var preSet = new HashSet<string>(
            (preSelectedNames ?? Array.Empty<string>())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n.Trim()),
            StringComparer.OrdinalIgnoreCase);
        _recordCount = records.Count;
        Title = "측정인 전송 설정";
        Width = 750;
        Height = 650;
        CanResize = true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new LinearGradientBrush
        {
            StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative),
            EndPoint = new Avalonia.RelativePoint(0.5, 1, Avalonia.RelativeUnit.Relative),
            GradientStops = new Avalonia.Media.GradientStops
            {
                new Avalonia.Media.GradientStop(Color.Parse("#0f0f1f"), 0),
                new Avalonia.Media.GradientStop(Color.Parse("#1a1a3a"), 1),
            }
        };
        SystemDecorations = Avalonia.Controls.SystemDecorations.Full;

        var measAgents = agents
            .Where(a => !string.IsNullOrEmpty(a.측정인고유번호))
            .OrderBy(a => string.IsNullOrEmpty(a.사번) ? "zzz" : a.사번, StringComparer.OrdinalIgnoreCase)
            .ThenBy(a => a.성명)
            .ToList();

        var root = new StackPanel { Margin = new Avalonia.Thickness(20), Spacing = 16 };

        // 헤더
        root.Children.Add(new TextBlock
        {
            Text = "🌐 측정인 전송 설정",
            FontSize = 16,
            FontFamily = Font,
            FontWeight = FontWeight.Bold,
            Foreground = Brush.Parse("#aaaaff"),
        });

        // ── 선택 패널 ──
        _selectionPanel = new StackPanel { Spacing = 14 };

        _selectionPanel.Children.Add(new TextBlock
        {
            Text = "측정인력을 선택하세요. (모든 시료에 동일 적용)",
            FontSize = 12,
            FontFamily = Font,
            Foreground = Brush.Parse("#888888"),
        });

        // ── 측정인력 섹션 ──
        var empPanel = new StackPanel { Spacing = 12 };

        empPanel.Children.Add(new TextBlock
        {
            Text = "📋 수질분석센터 기술인력:",
            FontSize = 13,
            FontFamily = Font,
            FontWeight = FontWeight.Bold,
            Foreground = Brush.Parse("#aaaacc"),
        });

        // 측정인력 WrapPanel (뱃지 스타일)
        var empWrap = new WrapPanel
        {
            Orientation = Orientation.Horizontal,
            ItemHeight = 32,
            Margin = new Avalonia.Thickness(0),
            ItemWidth = 220,
        };

        foreach (var ag in measAgents)
        {
            var itemPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 8,
                Margin = new Avalonia.Thickness(0),
                VerticalAlignment = VerticalAlignment.Center,
            };

            // 프로필 사진 (동그란 원)
            var photoPath = !string.IsNullOrEmpty(ag.PhotoPath)
                ? (System.IO.Path.IsPathRooted(ag.PhotoPath) ? ag.PhotoPath : System.IO.Path.Combine("Data/Photos", ag.PhotoPath))
                : null;

            Control photoControl;
            if (!string.IsNullOrEmpty(photoPath) && System.IO.File.Exists(photoPath))
            {
                try
                {
                    var photoImage = new Image
                    {
                        Width = 28,
                        Height = 28,
                        Stretch = Avalonia.Media.Stretch.UniformToFill,
                        Source = new Avalonia.Media.Imaging.Bitmap(photoPath),
                    };

                    photoControl = new Border
                    {
                        Width = 28,
                        Height = 28,
                        CornerRadius = new Avalonia.CornerRadius(14),
                        BorderBrush = Brush.Parse("#2a2a2a"),
                        BorderThickness = new Avalonia.Thickness(1),
                        ClipToBounds = true,
                        Child = photoImage,
                        VerticalAlignment = VerticalAlignment.Center,
                    };
                }
                catch
                {
                    var (initBg, _) = BadgeColorHelper.GetBadgeColor(ag.성명);
                    photoControl = new Border
                    {
                        Width = 28,
                        Height = 28,
                        CornerRadius = new Avalonia.CornerRadius(14),
                        Background = new SolidColorBrush(Color.Parse(initBg)),
                        VerticalAlignment = VerticalAlignment.Center,
                    };
                }
            }
            else
            {
                // 사진 없으면 색상 원
                var (initBg, _) = BadgeColorHelper.GetBadgeColor(ag.성명);
                photoControl = new Border
                {
                    Width = 28,
                    Height = 28,
                    CornerRadius = new Avalonia.CornerRadius(14),
                    Background = new SolidColorBrush(Color.Parse(initBg)),
                    VerticalAlignment = VerticalAlignment.Center,
                };
            }

            itemPanel.Children.Add(photoControl);

            // 직급 뱃지
            var (badgeBg, badgeFg) = BadgeColorHelper.GetBadgeColor(ag.직급);
            var badge = new Border
            {
                Background = new SolidColorBrush(Color.Parse(badgeBg)),
                CornerRadius = new Avalonia.CornerRadius(6),
                Padding = new Avalonia.Thickness(6, 2, 6, 2),
                Height = 18,
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = ag.직급,
                    FontFamily = Font,
                    FontSize = 9,
                    Foreground = new SolidColorBrush(Color.Parse(badgeFg)),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                }
            };
            itemPanel.Children.Add(badge);

            var nameText = new TextBlock
            {
                Text = ag.성명,
                FontFamily = Font,
                FontSize = 11,
                Foreground = Brush.Parse("#c8c8e8"),
                VerticalAlignment = VerticalAlignment.Center,
            };
            itemPanel.Children.Add(nameText);

            bool isPreSelected = preSet.Contains(ag.성명.Trim());
            var cb = new CheckBox
            {
                Tag = ag.측정인고유번호,
                IsChecked = isPreSelected,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(4, 0, 0, 0),
            };
            _agentChecks.Add(cb);
            _agentInfo.Add((ag.성명, ag.측정인고유번호));
            itemPanel.Children.Add(cb);

            empWrap.Children.Add(itemPanel);
        }

        empPanel.Children.Add(empWrap);

        // 측정인력 박스
        var empBorder = new Border
        {
            Background = new LinearGradientBrush
            {
                StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative),
                EndPoint = new Avalonia.RelativePoint(1, 1, Avalonia.RelativeUnit.Relative),
                GradientStops = new Avalonia.Media.GradientStops
                {
                    new Avalonia.Media.GradientStop(Color.Parse("#1a1a3a"), 0),
                    new Avalonia.Media.GradientStop(Color.Parse("#0f0f2a"), 1),
                }
            },
            CornerRadius = new Avalonia.CornerRadius(8),
            Padding = new Avalonia.Thickness(12, 10),
            Child = empPanel,
            BorderBrush = Brush.Parse("#88ffaa"),
            BorderThickness = new Avalonia.Thickness(0, 0, 0, 2),
            Margin = new Avalonia.Thickness(0, 0, 0, 12),
        };
        _selectionPanel.Children.Add(empBorder);

        // ── 의뢰 섹션 ──
        var recordPanel = new StackPanel { Spacing = 12 };

        recordPanel.Children.Add(new TextBlock
        {
            Text = $"📝 의뢰 ({records.Count}건):",
            FontSize = 13,
            FontFamily = Font,
            FontWeight = FontWeight.Bold,
            Foreground = Brush.Parse("#88ccff"),
        });

        var recordsList = new StackPanel { Spacing = 6, Margin = new Avalonia.Thickness(0) };
        for (int i = 0; i < records.Count; i++)
        {
            var (sample, analytes, company) = records[i];

            // 레코드 아이템: 아이콘 + 샘플명 + 업체뱃지
            var recordItem = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 8,
                VerticalAlignment = VerticalAlignment.Center,
            };

            // 아이콘 (샘플색상 배경)
            var (iconBg, _) = BadgeColorHelper.GetBadgeColor(sample ?? "");
            var icon = new Border
            {
                Width = 18,
                Height = 18,
                CornerRadius = new Avalonia.CornerRadius(3),
                Background = new SolidColorBrush(Color.Parse(iconBg)),
                VerticalAlignment = VerticalAlignment.Center,
            };
            recordItem.Children.Add(icon);

            recordItem.Children.Add(new TextBlock
            {
                Text = $"[{i + 1}] {sample}",
                FontSize = 10,
                FontFamily = Font,
                Foreground = Brush.Parse("#c8c8e8"),
                VerticalAlignment = VerticalAlignment.Center,
            });

            // 업체 뱃지
            var (companyBg, companyFg) = BadgeColorHelper.GetBadgeColor(company ?? "");
            var companyBadge = new Border
            {
                Background = new SolidColorBrush(Color.Parse(companyBg)),
                CornerRadius = new Avalonia.CornerRadius(8),
                Padding = new Avalonia.Thickness(8, 0, 8, 0),
                Child = new TextBlock
                {
                    Text = company ?? "-",
                    FontFamily = Font,
                    FontSize = 9,
                    Foreground = new SolidColorBrush(Color.Parse(companyFg)),
                    VerticalAlignment = VerticalAlignment.Center,
                }
            };
            recordItem.Children.Add(companyBadge);

            recordsList.Children.Add(recordItem);
        }

        recordPanel.Children.Add(new Border
        {
            Background = new LinearGradientBrush
            {
                StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative),
                EndPoint = new Avalonia.RelativePoint(0, 1, Avalonia.RelativeUnit.Relative),
                GradientStops = new Avalonia.Media.GradientStops
                {
                    new Avalonia.Media.GradientStop(Color.Parse("#0a1a2a"), 0),
                    new Avalonia.Media.GradientStop(Color.Parse("#0f0f1a"), 1),
                }
            },
            CornerRadius = new Avalonia.CornerRadius(6),
            Padding = new Avalonia.Thickness(12, 8),
            BorderBrush = Brush.Parse("#1a4a6a"),
            BorderThickness = new Avalonia.Thickness(1),
            Child = new ScrollViewer
            {
                Content = recordsList,
                MaxHeight = 110,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            }
        });

        // 의뢰 박스
        var recordBorder = new Border
        {
            Background = new LinearGradientBrush
            {
                StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative),
                EndPoint = new Avalonia.RelativePoint(1, 1, Avalonia.RelativeUnit.Relative),
                GradientStops = new Avalonia.Media.GradientStops
                {
                    new Avalonia.Media.GradientStop(Color.Parse("#1a2a1a"), 0),
                    new Avalonia.Media.GradientStop(Color.Parse("#0f1a0f"), 1),
                }
            },
            CornerRadius = new Avalonia.CornerRadius(8),
            Padding = new Avalonia.Thickness(12, 10),
            Child = recordPanel,
            BorderBrush = Brush.Parse("#88ccff"),
            BorderThickness = new Avalonia.Thickness(0, 0, 0, 2),
        };
        _selectionPanel.Children.Add(recordBorder);

        root.Children.Add(_selectionPanel);

        // 버튼
        var btnConfirm = new Button
        {
            Content = "✅ 전송 시작",
            Height = 32,
            FontSize = 12,
            FontFamily = Font,
            FontWeight = FontWeight.Bold,
            Background = Brush.Parse("#1a3a1a"),
            Foreground = Brush.Parse("#88ffaa"),
            BorderThickness = new Avalonia.Thickness(0),
            CornerRadius = new Avalonia.CornerRadius(4),
            Padding = new Avalonia.Thickness(20, 0),
            HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center,
            VerticalContentAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        var btnCancel = new Button
        {
            Content = "취소",
            Height = 32,
            FontSize = 12,
            FontFamily = Font,
            Background = Brush.Parse("#2a1a1a"),
            Foreground = Brush.Parse("#ff9999"),
            BorderThickness = new Avalonia.Thickness(0),
            CornerRadius = new Avalonia.CornerRadius(4),
            Padding = new Avalonia.Thickness(20, 0),
            HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center,
            VerticalContentAlignment = Avalonia.Layout.VerticalAlignment.Center,
        };

        btnConfirm.Click += (_, _) =>
        {
            var selected = new List<(string name, string empId)>();
            for (int i = 0; i < _agentChecks.Count; i++)
            {
                if (_agentChecks[i].IsChecked == true && !string.IsNullOrEmpty(_agentInfo[i].empId))
                    selected.Add(_agentInfo[i]);
            }

            var selectedEmpIds = selected.Select(x => x.empId).ToList();
            var selectedNames  = selected.Select(x => x.name).ToList();

            string purpose = selectedEmpIds.Count > 0 ? "SELF" : "CF";

            PurposeValues = Enumerable.Repeat(purpose, _recordCount).ToList();
            EmpIdsPerRecord = Enumerable.Range(0, _recordCount)
                .Select(_ => new List<string>(selectedEmpIds))
                .ToList();
            SelectedAgents = selected;
            SelectedAgentNames = selectedNames;

            Confirmed = true;
            Close();
        };

        btnConfirm.Background = new LinearGradientBrush
        {
            StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative),
            EndPoint = new Avalonia.RelativePoint(1, 1, Avalonia.RelativeUnit.Relative),
            GradientStops = new Avalonia.Media.GradientStops
            {
                new Avalonia.Media.GradientStop(Color.Parse("#2a5a3a"), 0),
                new Avalonia.Media.GradientStop(Color.Parse("#0a3a1a"), 1),
            }
        };
        btnConfirm.BorderBrush = Brush.Parse("#88ffaa");
        btnConfirm.BorderThickness = new Avalonia.Thickness(0, 0, 0, 2);

        btnCancel.Background = new LinearGradientBrush
        {
            StartPoint = new Avalonia.RelativePoint(0, 0, Avalonia.RelativeUnit.Relative),
            EndPoint = new Avalonia.RelativePoint(1, 1, Avalonia.RelativeUnit.Relative),
            GradientStops = new Avalonia.Media.GradientStops
            {
                new Avalonia.Media.GradientStop(Color.Parse("#5a3a2a"), 0),
                new Avalonia.Media.GradientStop(Color.Parse("#3a1a1a"), 1),
            }
        };
        btnCancel.BorderBrush = Brush.Parse("#ff9999");
        btnCancel.BorderThickness = new Avalonia.Thickness(0, 0, 0, 2);

        btnCancel.Click += (_, _) => Close();

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Avalonia.Thickness(0, 16, 0, 0),
            Spacing = 10,
        };
        btnPanel.Children.Add(btnCancel);
        btnPanel.Children.Add(btnConfirm);

        root.Children.Add(btnPanel);

        Content = new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }
}
