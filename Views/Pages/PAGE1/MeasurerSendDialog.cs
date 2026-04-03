using Avalonia.Controls;
using ETA.Views;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Models;
using System.Collections.Generic;
using System.Linq;

namespace ETA.Views.Pages.PAGE1;

/// <summary>
/// 측정인 전송 사전 설정 다이얼로그
/// — 부모노드(레코드)별 측정목적(자가측정용/참고용) 라디오버튼
/// — 부모노드별 측정인력 다중 선택
/// </summary>
public sealed class MeasurerSendDialog : Window
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    // ── 결과 프로퍼티 ─────────────────────────────────────────────────
    public bool Confirmed { get; private set; }
    /// <summary>레코드 인덱스 순서대로 측정목적 값 ("SELF" or "CF")</summary>
    public List<string> PurposeValues { get; private set; } = new();
    /// <summary>레코드 인덱스 순서대로 측정인력 고유번호 리스트</summary>
    public List<List<string>> EmpIdsPerRecord { get; private set; } = new();

    // ── 내부 per-record 컨트롤 ───────────────────────────────────────
    private readonly List<(RadioButton rbSelf, RadioButton rbCf)> _purposeRows = new();
    private readonly List<List<CheckBox>> _agentChecksPerRecord = new();

    public MeasurerSendDialog(
        IReadOnlyList<(string sample, string analytes, string company)> records,
        List<Agent> agents)
    {
        Title   = "측정인 전송 설정";
        Width   = 780;
        Height  = 680;
        CanResize = true;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = Brush.Parse("#12121f");
        SystemDecorations = Avalonia.Controls.SystemDecorations.Full;

        var sectionBg = AppTheme.BgPrimary;
        var labelFg   = Brush.Parse("#c8c8e8");
        var valueFg   = AppTheme.FgSuccess;
        var dimFg     = AppTheme.FgMuted;
        var sepBg     = Brush.Parse("#2a2a44");

        // 측정인고유번호 있는 인원만
        var measAgents = agents
            .Where(a => !string.IsNullOrEmpty(a.측정인고유번호))
            .OrderBy(a => a.성명)
            .ToList();

        // ── 헤더 ───────────────────────────────────────────────────
        var headerSp = new StackPanel { Margin = new Avalonia.Thickness(0, 0, 0, 8) };
        headerSp.Children.Add(MakeLabel("🌐 측정인 전송 설정", "#aaaaff", 13, bold: true));
        headerSp.Children.Add(new TextBlock
        {
            Text = "각 의뢰별로 측정목적과 측정인력을 선택하세요.",
            FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = dimFg,
            Margin = new Avalonia.Thickness(0, 2, 0, 0),
        });

        // ── 레코드별 설정 패널 ────────────────────────────────────
        var recordsSp = new StackPanel { Spacing = 8 };

        for (int i = 0; i < records.Count; i++)
        {
            var (sample, analytes, company) = records[i];
            string groupName = $"purpose_{i}";

            // ─ 레코드 카드 ─────────────────────────────────────
            var card = new Border
            {
                Background = sectionBg,
                CornerRadius = new Avalonia.CornerRadius(6),
                Padding = new Avalonia.Thickness(10, 8),
            };

            var cardContent = new StackPanel { Spacing = 6 };

            // 레코드 제목
            var titleRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
            titleRow.Children.Add(new TextBlock
            {
                Text = $"[{i + 1}]  {sample}",
                FontSize = AppTheme.FontBase, FontFamily = Font, FontWeight = FontWeight.Bold,
                Foreground = valueFg,
                VerticalAlignment = VerticalAlignment.Center,
                [Grid.ColumnProperty] = 0,
            });
            titleRow.Children.Add(new TextBlock
            {
                Text = $"  {analytes}  ·  {company}",
                FontSize = AppTheme.FontXS, FontFamily = Font, Foreground = dimFg,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                [Grid.ColumnProperty] = 1,
            });
            cardContent.Children.Add(titleRow);

            // 구분선
            cardContent.Children.Add(new Border
            {
                Height = 1, Background = sepBg,
                Margin = new Avalonia.Thickness(0, 2, 0, 2),
            });

            // 측정목적 라디오
            var rbSelf = new RadioButton
            {
                Content = "자가측정용 (SELF)",
                GroupName = groupName,
                FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = labelFg,
                Margin = new Avalonia.Thickness(0, 0, 16, 0),
            };
            var rbCf = new RadioButton
            {
                Content = "참고용 (CF)",
                GroupName = groupName,
                IsChecked = true,
                FontSize = AppTheme.FontSM, FontFamily = Font, Foreground = labelFg,
            };
            var purposeRow = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 4,
            };
            purposeRow.Children.Add(MakeLabel("측정목적:", "#aaaacc", 10));
            purposeRow.Children.Add(rbSelf);
            purposeRow.Children.Add(rbCf);
            cardContent.Children.Add(purposeRow);
            _purposeRows.Add((rbSelf, rbCf));

            // 측정인력 체크박스 리스트
            cardContent.Children.Add(MakeLabel("측정인력 (체크박스 다중선택):", "#aaaacc", 10));
            var checks = new List<CheckBox>();
            var checksPanel = new StackPanel
            {
                Spacing = 2,
                Margin = new Avalonia.Thickness(4, 2, 0, 0),
            };

            foreach (var ag in measAgents)
            {
                var cb = new CheckBox
                {
                    Content = $"{ag.성명}  ({ag.직급})  [{ag.측정인고유번호}]",
                    Tag = ag.측정인고유번호,
                    FontFamily = Font,
                    FontSize = AppTheme.FontSM,
                    Foreground = labelFg,
                    Margin = new Avalonia.Thickness(0, 0, 0, 1),
                };
                checks.Add(cb);
                checksPanel.Children.Add(cb);
            }

            cardContent.Children.Add(new Border
            {
                Background = Brush.Parse("#16162a"),
                CornerRadius = new Avalonia.CornerRadius(4),
                Padding = new Avalonia.Thickness(6, 4),
                Child = new ScrollViewer
                {
                    Content = checksPanel,
                    MaxHeight = 120,
                    HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
                    VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                }
            });

            _agentChecksPerRecord.Add(checks);

            card.Child = cardContent;
            recordsSp.Children.Add(card);
        }

        // ── 버튼 ────────────────────────────────────────────────
        var btnConfirm = MakeBtn("✅ 전송 시작", "#1a3a1a", "#88ffaa");
        var btnCancel  = MakeBtn("취소", "#2a1a1a", "#ff9999");

        btnConfirm.Click += (_, _) =>
        {
            PurposeValues.Clear();
            EmpIdsPerRecord.Clear();
            for (int i = 0; i < _purposeRows.Count; i++)
            {
                PurposeValues.Add(_purposeRows[i].rbSelf.IsChecked == true ? "SELF" : "CF");
                EmpIdsPerRecord.Add(
                    _agentChecksPerRecord[i]
                        .Where(cb => cb.IsChecked == true)
                        .Select(cb => cb.Tag as string ?? "")
                        .Where(s => !string.IsNullOrEmpty(s))
                        .ToList()
                );
            }
            Confirmed = true;
            Close();
        };
        btnCancel.Click += (_, _) => Close();

        var btnRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Avalonia.Thickness(0, 12, 0, 0),
            Spacing = 8,
        };
        btnRow.Children.Add(btnCancel);
        btnRow.Children.Add(btnConfirm);

        // ── 최종 레이아웃 ────────────────────────────────────────
        var root = new StackPanel { Margin = new Avalonia.Thickness(16) };
        root.Children.Add(headerSp);
        root.Children.Add(recordsSp);
        root.Children.Add(btnRow);

        Content = new ScrollViewer
        {
            Content = root,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
        };
    }

    private static TextBlock MakeLabel(string text, string fg, double size = 11, bool bold = false) =>
        new TextBlock
        {
            Text = text, FontSize = size, FontFamily = Font,
            FontWeight = bold ? FontWeight.Bold : FontWeight.Normal,
            Foreground = Brush.Parse(fg),
            Margin = new Avalonia.Thickness(0, 0, 0, 2),
            VerticalAlignment = VerticalAlignment.Center,
        };

    private static Button MakeBtn(string text, string bg, string fg) => new Button
    {
        Content = text, Height = 30, FontSize = AppTheme.FontBase, FontFamily = Font,
        Background = Brush.Parse(bg), Foreground = Brush.Parse(fg),
        BorderThickness = new Avalonia.Thickness(0),
        CornerRadius = new Avalonia.CornerRadius(4),
        Padding = new Avalonia.Thickness(16, 0),
        VerticalContentAlignment = VerticalAlignment.Center,
    };
}

