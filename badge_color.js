const fs = require('fs');
const p = 'C:/Users/ironu/Documents/ETA/Views/Pages/TestReportPage.axaml.cs';
let s = fs.readFileSync(p, 'utf8');

const a = s.indexOf('    private TreeViewItem MakeSampleNode');
const b = s.indexOf('    // \uc5c5\uccb4 \ub178\ub4dc \ud558\uc704'); // 업체 노드 하위
if (a===-1||b===-1){console.error('markers not found',a,b);process.exit(1);}

const nm = `    // 초성별 배지 색상 (bg, fg)
    private static (string Bg, string Fg) GetBadgeColor(string \uc57d\uce6d)
    {
        if (string.IsNullOrEmpty(\uc57d\uce6d)) return ("#2a2a2a", "#aaaaaa");
        char c = \uc57d\uce6d[0];
        if (c < '\uac00' || c > '\ud7a3') return ("#2a2a2a", "#aaaaaa");
        int cho = (c - 0xAC00) / (21 * 28);
        return cho switch
        {
            0  => ("#1a3a1a", "#88cc88"),  // ㄱ
            1  => ("#1a2a3a", "#88aacc"),  // ㄲ
            2  => ("#2a1a3a", "#aa88cc"),  // ㄴ
            3  => ("#3a2a1a", "#ccaa88"),  // ㄷ
            4  => ("#1a3a3a", "#88ccbb"),  // ㄸ
            5  => ("#2a1a2a", "#cc88aa"),  // ㄹ
            6  => ("#1a2a2a", "#88cccc"),  // ㅁ
            7  => ("#1a1a3a", "#8888cc"),  // ㅂ
            8  => ("#2a3a1a", "#aacc88"),  // ㅃ
            9  => ("#3a1a1a", "#cc8888"),  // ㅅ
            10 => ("#2a3a3a", "#88ccee"),  // ㅆ
            11 => ("#2a2a1a", "#cccc88"),  // ㅇ
            12 => ("#1a3a2a", "#88ccaa"),  // ㅈ
            13 => ("#3a3a1a", "#ccccaa"),  // ㅉ
            14 => ("#3a1a2a", "#cc88bb"),  // ㅊ
            15 => ("#2a2a3a", "#aaaacc"),  // ㅋ
            16 => ("#1a2a1a", "#88cc99"),  // ㅌ
            17 => ("#3a1a3a", "#cc88cc"),  // ㅍ
            18 => ("#3a2a3a", "#ccaacc"),  // ㅎ
            _  => ("#2a2a2a", "#aaaaaa"),
        };
    }

    private TreeViewItem MakeSampleNode(SampleRequest sample, bool showCompany = false)
    {
        bool incomplete = sample.\ubd84\uc11d\uacb0\uacfc.Values.Any(v =>
            string.Equals(v, "O", StringComparison.OrdinalIgnoreCase));
        string iconColor = incomplete ? "#ee4444" : "#44cc44";

        var chk = new CheckBox
        {
            IsChecked         = _checkedSamples.Contains(sample),
            VerticalAlignment = VerticalAlignment.Center,
            Margin            = new Thickness(0, 0, 4, 0),
        };
        chk.IsCheckedChanged += (_, _) =>
        {
            if (chk.IsChecked == true) _checkedSamples.Add(sample);
            else                       _checkedSamples.Remove(sample);
        };

        string mmdd = "";
        if (sample.\ucc44\ucde8\uc77c\uc790.Length >= 10)
            mmdd = sample.\ucc44\ucde8\uc77c\uc790.Substring(5, 5).Replace("-", "/");
        else if (sample.\ucc44\ucde8\uc77c\uc790.Length >= 7)
            mmdd = sample.\ucc44\ucde8\uc77c\uc790.Substring(5);

        string icon = incomplete ? "\ud83c\udf76" : "\ud83e\uddea";
        string labelText = string.IsNullOrEmpty(mmdd)
            ? sample.\uc2dc\ub8cc\uba85
            : $"{sample.\uc2dc\ub8cc\uba85}  ({mmdd})";

        var (badgeBg, badgeFg) = GetBadgeColor(sample.\uc57d\uce6d);

        var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,Auto,Auto,*"), Margin = new Thickness(2, 2) };
        row.Children.Add(new ContentControl { Content = chk, [Grid.ColumnProperty] = 0 });
        row.Children.Add(new TextBlock
        {
            Text = icon, FontSize = 12,
            Foreground = new SolidColorBrush(Color.Parse(iconColor)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(2, 0, 4, 0),
            [Grid.ColumnProperty] = 1,
        });
        row.Children.Add(new Border
        {
            Background = Brush.Parse(badgeBg), CornerRadius = new CornerRadius(3),
            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 2,
            Child = new TextBlock { Text = sample.\uc57d\uce6d, FontSize = 9, FontFamily = Font,
                                    Foreground = Brush.Parse(badgeFg) },
        });
        row.Children.Add(new TextBlock
        {
            Text = labelText, FontSize = 11, FontFamily = Font,
            Foreground = Brush.Parse("#dddddd"),
            TextTrimming = Avalonia.Media.TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            [Grid.ColumnProperty] = 3,
        });

        return new TreeViewItem { Tag = sample, Header = row };
    }

`;

s = s.slice(0, a) + nm + s.slice(b);
fs.writeFileSync(p, s, 'utf8');
console.log('OK');
