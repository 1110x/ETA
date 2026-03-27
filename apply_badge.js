const fs = require('fs');

// ── 1. TestReportPage: local GetBadgeColor → BadgeColorHelper.GetBadgeColor ─
{
  const p = 'C:/Users/ironu/Documents/ETA/Views/Pages/TestReportPage.axaml.cs';
  let s = fs.readFileSync(p, 'utf8');

  // Remove local GetBadgeColor method
  const start = s.indexOf('    // \ucd08\uc131\ubcc4 \ubc30\uc9c0 \uc0c9\uc0c1'); // 초성별 배지 색상
  const end   = s.indexOf('    private TreeViewItem MakeSampleNode');
  if (start !== -1 && end !== -1)
    s = s.slice(0, start) + s.slice(end);

  // Replace local call with static helper
  s = s.replace(
    'GetBadgeColor(sample.\uc57d\uce6d)',
    'BadgeColorHelper.GetBadgeColor(sample.\uc57d\uce6d)'
  );

  if (!s.includes('using ETA.Services;'))
    s = s.replace('using ETA.Models;', 'using ETA.Models;\nusing ETA.Services;');

  fs.writeFileSync(p, s, 'utf8');
  console.log('1. TestReportPage OK');
}

// ── 2. QuotationHistoryPanel ─────────────────────────────────────────────────
{
  const p = 'C:/Users/ironu/Documents/ETA/Views/Pages/QuotationHistoryPanel.axaml.cs';
  let s = fs.readFileSync(p, 'utf8');

  if (!s.includes('using ETA.Services;'))
    s = s.replace(/^(using [^\n]+\n)/m, '$1using ETA.Services;\n');

  // MakeIssueLeaf
  const issueOld = [
    '    private TreeViewItem MakeIssueLeaf(QuotationIssue issue)',
    '    {',
    '        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };',
    '        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };',
    '        topRow.Children.Add(new Border',
    '        {',
    '            Background = Brush.Parse("#1a3a1a"), CornerRadius = new CornerRadius(3),',
    '            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),',
    '            [Grid.ColumnProperty] = 0,',
    '            Child = new TextBlock { Text = issue.\uc57d\uce6d, FontSize = 9, FontFamily = Font,',
    '                                    Foreground = Brush.Parse("#88cc88") },',
    '        });',
  ].join('\r\n');

  const issueNew = [
    '    private TreeViewItem MakeIssueLeaf(QuotationIssue issue)',
    '    {',
    '        var (ibg, ifg) = BadgeColorHelper.GetBadgeColor(issue.\uc57d\uce6d);',
    '        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };',
    '        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };',
    '        topRow.Children.Add(new Border',
    '        {',
    '            Background = Brush.Parse(ibg), CornerRadius = new CornerRadius(3),',
    '            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),',
    '            [Grid.ColumnProperty] = 0,',
    '            Child = new TextBlock { Text = issue.\uc57d\uce6d, FontSize = 9, FontFamily = Font,',
    '                                    Foreground = Brush.Parse(ifg) },',
    '        });',
  ].join('\r\n');

  // Also try LF line ending
  const issueOldLF = issueOld.replace(/\r\n/g, '\n');
  const issueNewLF = issueNew.replace(/\r\n/g, '\n');

  if (s.includes(issueOld)) {
    s = s.replace(issueOld, issueNew);
    console.log('  issue CRLF match');
  } else if (s.includes(issueOldLF)) {
    s = s.replace(issueOldLF, issueNewLF);
    console.log('  issue LF match');
  } else {
    console.error('  issue: no match found');
  }

  // MakeAnalysisLeaf
  const recOld = [
    '    private TreeViewItem MakeAnalysisLeaf(AnalysisRequestRecord rec)',
    '    {',
    '        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };',
    '        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };',
    '        topRow.Children.Add(new Border',
    '        {',
    '            Background = Brush.Parse("#3a1a1a"), CornerRadius = new CornerRadius(3),',
    '            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),',
    '            [Grid.ColumnProperty] = 0,',
    '            Child = new TextBlock { Text = rec.\uc57d\uce6d, FontSize = 9, FontFamily = Font,',
    '                                    Foreground = Brush.Parse("#cc8888") },',
    '        });',
  ].join('\r\n');

  const recNew = [
    '    private TreeViewItem MakeAnalysisLeaf(AnalysisRequestRecord rec)',
    '    {',
    '        var (rbg, rfg) = BadgeColorHelper.GetBadgeColor(rec.\uc57d\uce6d);',
    '        var sp = new StackPanel { Spacing = 1, Margin = new Thickness(4, 2) };',
    '        var topRow = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };',
    '        topRow.Children.Add(new Border',
    '        {',
    '            Background = Brush.Parse(rbg), CornerRadius = new CornerRadius(3),',
    '            Padding = new Thickness(4, 1), Margin = new Thickness(0, 0, 5, 0),',
    '            [Grid.ColumnProperty] = 0,',
    '            Child = new TextBlock { Text = rec.\uc57d\uce6d, FontSize = 9, FontFamily = Font,',
    '                                    Foreground = Brush.Parse(rfg) },',
    '        });',
  ].join('\r\n');

  const recOldLF = recOld.replace(/\r\n/g, '\n');
  const recNewLF = recNew.replace(/\r\n/g, '\n');

  if (s.includes(recOld)) {
    s = s.replace(recOld, recNew);
    console.log('  rec CRLF match');
  } else if (s.includes(recOldLF)) {
    s = s.replace(recOldLF, recNewLF);
    console.log('  rec LF match');
  } else {
    console.error('  rec: no match found');
  }

  fs.writeFileSync(p, s, 'utf8');
  console.log('2. QuotationHistoryPanel OK');
}

// ── 3. AnalysisRequestListPanel ──────────────────────────────────────────────
{
  const p = 'C:/Users/ironu/Documents/ETA/Views/Pages/AnalysisRequestListPanel.axaml.cs';
  let s = fs.readFileSync(p, 'utf8');

  if (!s.includes('using ETA.Services;'))
    s = s.replace(/^(using [^\n]+\n)/m, '$1using ETA.Services;\n');

  const old3 = [
    '        topRow.Children.Add(new Border',
    '        {',
    '            Background = BrushAbbrBg, CornerRadius = new Avalonia.CornerRadius(3),',
    '            Padding = new Avalonia.Thickness(4, 1), Margin = new Avalonia.Thickness(0, 0, 5, 0),',
    '            [Grid.ColumnProperty] = 0,',
    '            Child = new TextBlock { Text = rec.\uc57d\uce6d, FontSize = 9, FontFamily = Font, Foreground = BrushAbbrFg },',
    '        });',
  ].join('\r\n');

  const new3 = [
    '        var (abg, afg) = BadgeColorHelper.GetBadgeColor(rec.\uc57d\uce6d);',
    '        topRow.Children.Add(new Border',
    '        {',
    '            Background = Brush.Parse(abg), CornerRadius = new Avalonia.CornerRadius(3),',
    '            Padding = new Avalonia.Thickness(4, 1), Margin = new Avalonia.Thickness(0, 0, 5, 0),',
    '            [Grid.ColumnProperty] = 0,',
    '            Child = new TextBlock { Text = rec.\uc57d\uce6d, FontSize = 9, FontFamily = Font, Foreground = Brush.Parse(afg) },',
    '        });',
  ].join('\r\n');

  const old3LF = old3.replace(/\r\n/g, '\n');
  const new3LF = new3.replace(/\r\n/g, '\n');

  if (s.includes(old3)) {
    s = s.replace(old3, new3);
    console.log('  CRLF match');
  } else if (s.includes(old3LF)) {
    s = s.replace(old3LF, new3LF);
    console.log('  LF match');
  } else {
    console.error('  AnalysisRequest: no badge match');
  }

  fs.writeFileSync(p, s, 'utf8');
  console.log('3. AnalysisRequestListPanel OK');
}

// ── 4. ContractPage: add badge to company tree node ──────────────────────────
{
  const p = 'C:/Users/ironu/Documents/ETA/Views/Pages/ContractPage.axaml.cs';
  let s = fs.readFileSync(p, 'utf8');

  if (!s.includes('using ETA.Services;'))
    s = s.replace(/^(using [^\n]+\n)/m, '$1using ETA.Services;\n');

  // Replace the second TextBlock (약칭 · 계약구분) with badge + text
  const old4 = [
    '                            new TextBlock',
    '                            {',
    '                                Text       = string.IsNullOrEmpty(contract.C_Abbreviation)',
    '                                                 ? contract.C_ContractType',
    '                                                 : $"{contract.C_Abbreviation} \u00b7 {contract.C_ContractType}",',
    '                                FontSize   = 10,',
    '                                FontFamily = Font,',
    '                                Foreground = new SolidColorBrush(Color.Parse("#888888")),',
    '                            }',
  ].join('\r\n');

  const new4 = [
    '                            new StackPanel',
    '                            {',
    '                                Orientation = Orientation.Horizontal, Spacing = 4,',
    '                                Children =',
    '                                {',
    '                                    string.IsNullOrEmpty(contract.C_Abbreviation)',
    '                                        ? (Control)new TextBlock()',
    '                                        : new Border',
    '                                        {',
    '                                            Background   = Brush.Parse(BadgeColorHelper.GetBadgeColor(contract.C_Abbreviation).Bg),',
    '                                            CornerRadius = new CornerRadius(3),',
    '                                            Padding      = new Thickness(4, 1),',
    '                                            VerticalAlignment = VerticalAlignment.Center,',
    '                                            Child = new TextBlock',
    '                                            {',
    '                                                Text       = contract.C_Abbreviation,',
    '                                                FontSize   = 9, FontFamily = Font,',
    '                                                Foreground = Brush.Parse(BadgeColorHelper.GetBadgeColor(contract.C_Abbreviation).Fg),',
    '                                            }',
    '                                        },',
    '                                    new TextBlock',
    '                                    {',
    '                                        Text       = contract.C_ContractType,',
    '                                        FontSize   = 10, FontFamily = Font,',
    '                                        Foreground = new SolidColorBrush(Color.Parse("#888888")),',
    '                                        VerticalAlignment = VerticalAlignment.Center,',
    '                                    }',
    '                                }',
    '                            }',
  ].join('\r\n');

  const old4LF = old4.replace(/\r\n/g, '\n');
  const new4LF = new4.replace(/\r\n/g, '\n');

  if (s.includes(old4)) {
    s = s.replace(old4, new4);
    console.log('  CRLF match');
  } else if (s.includes(old4LF)) {
    s = s.replace(old4LF, new4LF);
    console.log('  LF match');
  } else {
    console.error('  ContractPage: no match');
    // Show nearby text for debug
    const idx = s.indexOf('C_ContractType');
    console.log('nearby:', JSON.stringify(s.slice(idx-200, idx+100)));
  }

  fs.writeFileSync(p, s, 'utf8');
  console.log('4. ContractPage OK');
}
