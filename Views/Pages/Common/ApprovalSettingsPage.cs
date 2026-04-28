using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using ETA.Services.Common;
using ETA.Services.SERVICE1;

namespace ETA.Views.Pages.Common;

/// <summary>
/// 설정 → 결재정보. 시험성적서 / 시험기록부 서명란에 들어갈 사람 이름을 사용자별로 저장한다.
/// 입력 후 포커스 이탈(LostFocus) 시 즉시 UserPrefsService 에 영속화된다.
/// </summary>
public sealed class ApprovalSettingsPage
{
    private static readonly FontFamily Font = new("avares://ETA/Assets/Fonts#Pretendard");

    private readonly StackPanel _root;

    public Control View => _root;

    public ApprovalSettingsPage()
    {
        _root = BuildRoot();
    }

    private StackPanel BuildRoot()
    {
        var outer = new StackPanel { Spacing = 8, Margin = new Thickness(12) };

        outer.Children.Add(new TextBlock
        {
            Text = "✍️  결재정보 (서명란)",
            FontFamily = Font,
            FontSize = AppTheme.FontXL,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        });

        outer.Children.Add(new TextBlock
        {
            Text = "시험성적서·시험기록부 출력 시 서명란에 들어갈 텍스트를 사용자별로 지정합니다. "
                 + "입력 후 다른 곳을 클릭하면 즉시 저장됩니다. "
                 + "시험성적서 양식 파일을 이전 버전으로 한 번이라도 출력했다면 아래 “양식 재생성” 을 눌러야 변경이 반영됩니다.",
            FontFamily = Font,
            FontSize = AppTheme.FontBase,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
        });

        // ── 시험성적서 ──────────────────────────────────────────────────────
        outer.Children.Add(SectionHeader("📄 시험성적서"));
        outer.Children.Add(BuildRow(
            "서명란 텍스트",
            "성적서 하단 “(서명)” 앞에 붙는 풀텍스트 (예: 품질책임 수질분야 환경측정분석사 박은지)",
            UserPrefsService.TestReportSignerQualityMgr,
            v => UserPrefsService.TestReportSignerQualityMgr = v));

        // 시험성적서 Word 양식 재생성 버튼
        var regenBtn = new Button
        {
            Content = "🔄 시험성적서 Word 양식 재생성",
            FontFamily = Font,
            FontSize = AppTheme.FontBase,
            Background = new SolidColorBrush(Color.Parse("#1a3a1a")),
            Foreground = new SolidColorBrush(Color.Parse("#7cd87c")),
            BorderThickness = new Thickness(0),
            Padding = new Thickness(10, 6),
            CornerRadius = new CornerRadius(4),
            Margin = new Thickness(4, 4, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left,
        };
        var regenStatus = new TextBlock
        {
            FontFamily = Font,
            FontSize = AppTheme.FontXS,
            Foreground = AppTheme.FgMuted,
            Margin = new Thickness(8, 6, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };
        regenBtn.Click += (_, _) =>
        {
            try
            {
                TestReportWordPrintService.GenerateDefaultTemplate();
                regenStatus.Text = "✓ 재생성 완료 — 다음 출력부터 새 결재정보가 적용됩니다.";
                regenStatus.Foreground = AppTheme.FgSuccess;
            }
            catch (System.Exception ex)
            {
                regenStatus.Text = $"✗ 실패: {ex.Message}";
                regenStatus.Foreground = AppTheme.FgDanger;
            }
        };
        var regenRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 4, 0, 4),
            Children = { regenBtn, regenStatus },
        };
        outer.Children.Add(regenRow);

        // ── 시험성적서 비고문구 ─────────────────────────────────────────────
        outer.Children.Add(SectionHeader("📝 시험성적서 비고문구"));
        outer.Children.Add(new TextBlock
        {
            Text = "성적서 하단 결재라인 아래 ▩ 안내 문구. 시료의 정도보증 여부에 따라 둘 중 하나가 출력됩니다. "
                 + "비워두면 기본 문구가 사용됩니다.",
            FontFamily = Font,
            FontSize = AppTheme.FontXS,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(4, 0, 0, 4),
        });
        outer.Children.Add(BuildLongRow(
            "정도보증 적용",
            "용도 = ‘정도보증 적용’ 시 출력될 비고문구",
            UserPrefsService.TestReportRemarkQc,
            v => UserPrefsService.TestReportRemarkQc = v,
            placeholder: UserPrefsService.DefaultTestReportRemarkQc));
        outer.Children.Add(BuildLongRow(
            "참고용",
            "용도 = ‘참고용’ (정도보증 미적용) 시 출력될 비고문구",
            UserPrefsService.TestReportRemarkRef,
            v => UserPrefsService.TestReportRemarkRef = v,
            placeholder: UserPrefsService.DefaultTestReportRemarkRef));

        // ── 시험기록부 ──────────────────────────────────────────────────────
        outer.Children.Add(SectionHeader("📕 시험기록부"));
        outer.Children.Add(BuildRow(
            "작성자",
            "시험기록부 하단 결재 표 좌측 칸",
            UserPrefsService.TestRecordBookSignerAuthor,
            v => UserPrefsService.TestRecordBookSignerAuthor = v));
        outer.Children.Add(BuildRow(
            "검토자",
            "시험기록부 하단 결재 표 가운데 칸",
            UserPrefsService.TestRecordBookSignerReviewer,
            v => UserPrefsService.TestRecordBookSignerReviewer = v));
        outer.Children.Add(BuildRow(
            "승인자",
            "시험기록부 하단 결재 표 우측 칸",
            UserPrefsService.TestRecordBookSignerApprover,
            v => UserPrefsService.TestRecordBookSignerApprover = v));

        return outer;
    }

    private static TextBlock SectionHeader(string text) => new()
    {
        Text = text,
        FontFamily = Font,
        FontSize = AppTheme.FontLG,
        FontWeight = FontWeight.SemiBold,
        Foreground = AppTheme.FgPrimary,
        Margin = new Thickness(0, 12, 0, 4),
    };

    /// <summary>긴 텍스트(비고문구 등) 입력용 — 라벨/힌트를 위에, TextBox 를 아래 전체 폭으로 배치.</summary>
    private static Border BuildLongRow(string label, string hint, string initial,
        System.Action<string> save, string placeholder = "")
    {
        var stack = new StackPanel { Spacing = 2, Margin = new Thickness(4, 2) };

        stack.Children.Add(new TextBlock
        {
            Text = label,
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        });
        stack.Children.Add(new TextBlock
        {
            Text = hint,
            FontFamily = Font,
            FontSize = AppTheme.FontXS,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
        });

        // 사용자 미입력 (= 기본 문구 사용) 상태면 플레이스홀더로 기본 문구를 보여줌.
        // 입력값이 기본과 동일하면 TextBox 내용은 비워둠 — 다음 저장 시 실제 빈 문자열로 저장돼 기본 문구 유지.
        bool isDefault = initial == placeholder;
        var input = new TextBox
        {
            Text = isDefault ? "" : initial,
            FontFamily = Font,
            FontSize = AppTheme.FontSM,
            Watermark = placeholder,
            AcceptsReturn = false,
            TextWrapping = TextWrapping.Wrap,
            MinHeight = 50,
            Margin = new Thickness(0, 4, 0, 0),
        };
        input.LostFocus += (_, _) => save(input.Text ?? "");
        stack.Children.Add(input);

        return new Border
        {
            Background = AppTheme.BgSecondary,
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(6, 4),
            Child = stack,
        };
    }

    private static Border BuildRow(string label, string hint, string initial, System.Action<string> save)
    {
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("160,240,*"),
            Margin = new Thickness(4, 2),
        };

        var labelPanel = new StackPanel { Spacing = 1, VerticalAlignment = VerticalAlignment.Center };
        labelPanel.Children.Add(new TextBlock
        {
            Text = label,
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            FontWeight = FontWeight.SemiBold,
            Foreground = AppTheme.FgPrimary,
        });
        labelPanel.Children.Add(new TextBlock
        {
            Text = hint,
            FontFamily = Font,
            FontSize = AppTheme.FontXS,
            Foreground = AppTheme.FgMuted,
            TextWrapping = TextWrapping.Wrap,
        });
        Grid.SetColumn(labelPanel, 0);
        grid.Children.Add(labelPanel);

        var input = new TextBox
        {
            Text = initial,
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            Watermark = "이름 입력",
            Margin = new Thickness(8, 0),
            VerticalAlignment = VerticalAlignment.Center,
        };
        input.LostFocus += (_, _) => save(input.Text ?? "");
        Grid.SetColumn(input, 1);
        grid.Children.Add(input);

        return new Border
        {
            Background = AppTheme.BgSecondary,
            CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 4),
            Child = grid,
        };
    }
}
