const fs = require('fs');
const p = 'C:/Users/ironu/Documents/ETA/Views/Pages/TestReportPage.axaml.cs';
let s = fs.readFileSync(p, 'utf8');

// ── 1. Fields ────────────────────────────────────────────────────────────
{
  const a = s.indexOf('    private bool _showImportPanel');
  const b = s.indexOf('    // \uc120\ud0dd\ub41c \uacb0\uacfc \ud589'); // 선택된 결과 행
  if (a===-1||b===-1){console.error('Fields markers not found',a,b);process.exit(1);}
  const nf =
`    private bool _showImportPanel = false;
    private Control? _importActionPanel;
    private ToggleSwitch?   _stdToggle;
    private ContentControl? _stdToggleContainer;

    // Excel \ubd88\ub7ec\uc624\uae30 \uacf5\uc720 \uc0c1\ud0dc
    private readonly List<(string A, string Q, string Y, string S, string Ex, string Nv, string Std, string Bg, string Fg)> _importRows = new();
    private readonly Dictionary<int, Dictionary<string, string>> _pendingByRow = new();
    private readonly List<(int RowId, string Analyte, string Existing, string NewVal)> _conflictRows = new();
    private List<string> _importFileNames = new();
    private string _importStatus      = "";
    private string _importStatusHex   = "#888888";
    private bool   _importHasPending  = false;
    private bool   _importHasConflict = false;

`;
  s = s.slice(0,a) + nf + s.slice(b);
  console.log('1. Fields OK');
}

// ── 2. LoadData reset ────────────────────────────────────────────────────
{
  const old2a = '_cachedImportPanel   = null;';
  const idx = s.indexOf(old2a);
  if (idx === -1) { console.error('LoadData marker not found'); process.exit(1); }
  // skip two lines (_cachedImportPanel and _cachedImportWrapper)
  const line1End = s.indexOf('\n', idx) + 1;
  const line2End = s.indexOf('\n', line1End) + 1;
  const newReset =
`        _importRows.Clear();
        _pendingByRow.Clear();
        _conflictRows.Clear();
        _importFileNames.Clear();
        _importStatus      = "";
        _importStatusHex   = "#888888";
        _importHasPending  = false;
        _importHasConflict = false;
        _importActionPanel = null;
`;
  s = s.slice(0, idx) + newReset + s.slice(line2End);
  console.log('2. LoadData OK');
}

// ── 3. BuildListControl ──────────────────────────────────────────────────
{
  const a = s.indexOf('    private Control BuildListControl()');
  const b = s.indexOf('    // \uc2dc\ub8cc \uc815\ubcf4 \ud5e4\ub354 \ud328\ub110'); // 시료 정보 헤더 패널
  if (a===-1||b===-1){console.error('BLC markers not found',a,b);process.exit(1);}
  const nm =
`    private Control BuildListControl()
    {
        if (_stdToggle == null)
        {
            _stdToggle = new ToggleSwitch
            {
                IsChecked  = _showImportPanel,
                OnContent  = "\ud83d\udcce Excel \ubd88\ub7ec\uc624\uae30",
                OffContent = "\ud83d\udccb \uc2dc\ud5d8\uc131\uc801\uc11c",
                FontFamily = Font, FontSize = 11,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new Thickness(8, 2),
            };
            _stdToggle.IsCheckedChanged += (_, _) =>
            {
                _showImportPanel = _stdToggle.IsChecked == true;
                if (_showImportPanel)
                {
                    if (_importActionPanel == null) _importActionPanel = BuildImportActionPanel();
                    EditPanelChanged?.Invoke(_importActionPanel);
                }
                else
                {
                    EditPanelChanged?.Invoke(null);
                }
                ResultListChanged?.Invoke(BuildListControl());
            };
        }
        else { _stdToggle.IsChecked = _showImportPanel; }

        if (_stdToggle.Parent is ContentControl oldC) oldC.Content = null;
        _stdToggleContainer = new ContentControl { Content = _stdToggle };

        if (_showImportPanel)
        {
            // Content3 = \uc561\uc158 \ud328\ub110 (\ud30c\uc77c \uc120\ud0dd, \uc800\uc7a5 \ub4f1)
            if (_importActionPanel == null) _importActionPanel = BuildImportActionPanel();
            EditPanelChanged?.Invoke(_importActionPanel);

            // Content2 = \ud1a0\uae00 + \uacb0\uacfc \ud14c\uc774\ube14
            var tablePanel = new StackPanel { Spacing = 1 };
            if (_importRows.Count > 0)
            {
                tablePanel.Children.Add(MakeImportTableRow(
                    "\ubd84\uc11d\ud56d\ubaa9", "\uacac\uc801\ubc88\ud638", "\uc57d\uce6d", "\uc2dc\ub8cc\uba85",
                    "\uc774\uc804\uacb0\uacfc", "\uc0c8\uacb0\uacfc\uac12", "\ubc29\ub958\uae30\uc900",
                    true, "#2a2a3a", "#aaaacc"));
                foreach (var row in _importRows)
                    tablePanel.Children.Add(MakeImportTableRow(
                        row.A, row.Q, row.Y, row.S, row.Ex, row.Nv, row.Std,
                        false, row.Bg, row.Fg));
            }
            else
            {
                tablePanel.Children.Add(new TextBlock
                {
                    Text = "\uc624\ub978\ucabd\uc5d0\uc11c Excel \ud30c\uc77c\uc744 \uc120\ud0dd\ud558\uba74 \uacb0\uacfc\uac00 \uc5ec\uae30\uc5d0 \ud45c\uc2dc\ub429\ub2c8\ub2e4.",
                    FontFamily = Font, FontSize = 12,
                    Foreground = new SolidColorBrush(Color.Parse("#555566")),
                    Margin     = new Thickness(12, 20),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                });
            }
            var tableScroll = new ScrollViewer
            {
                Content = tablePanel,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            };
            var importOuter = new Grid { RowDefinitions = new RowDefinitions("Auto,*") };
            Grid.SetRow(_stdToggleContainer, 0);
            Grid.SetRow(tableScroll, 1);
            importOuter.Children.Add(_stdToggleContainer);
            importOuter.Children.Add(tableScroll);
            return importOuter;
        }

        // Content2 = \uacb0\uacfc \ub9ac\uc2a4\ud2b8 \ubaa8\ub4dc
        var listPanel = new StackPanel { Spacing = 0 };
        RefreshListPanel(listPanel);

        var outer = new Grid { RowDefinitions = new RowDefinitions("Auto,Auto,Auto,*") };
        var infoPanel = BuildSampleInfoHeader();
        Grid.SetRow(infoPanel, 0);
        Grid.SetRow(_stdToggleContainer, 1);
        var resultHeader = BuildResultListHeader();
        Grid.SetRow(resultHeader, 2);
        var resultScroll = new ScrollViewer
        {
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            Content                       = listPanel
        };
        Grid.SetRow(resultScroll, 3);

        outer.Children.Add(infoPanel);
        outer.Children.Add(_stdToggleContainer);
        outer.Children.Add(resultHeader);
        outer.Children.Add(resultScroll);
        return outer;
    }

`;
  s = s.slice(0,a) + nm + s.slice(b);
  console.log('3. BuildListControl OK');
}

// ── 4. BuildExcelImportPanel → BuildImportActionPanel + BuildConflictUI ──
{
  const a = s.indexOf('    private Control BuildExcelImportPanel()');
  const b = s.indexOf('    private static Border MakeImportTableRow(');
  if (a===-1||b===-1){console.error('BEI markers not found',a,b);process.exit(1);}
  const nm =
`    // =========================================================================
    // Excel \ubd88\ub7ec\uc624\uae30 \uc561\uc158 \ud328\ub110 (ActivePageContent3)
    // =========================================================================
    private Control BuildImportActionPanel()
    {
        var statusTb = new TextBlock
        {
            Text         = _importStatus,
            Foreground   = new SolidColorBrush(Color.Parse(_importStatusHex)),
            FontFamily   = Font, FontSize = 11,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin       = new Thickness(0, 4, 0, 0),
        };

        var fileListPanel = new StackPanel { Spacing = 3 };
        foreach (var fn in _importFileNames)
            fileListPanel.Children.Add(new StackPanel
            {
                Orientation = Orientation.Horizontal, Spacing = 5,
                Children =
                {
                    new TextBlock { Text = "\ud83d\udcc5", FontSize = 13, VerticalAlignment = VerticalAlignment.Center },
                    new TextBlock { Text = fn, FontFamily = Font, FontSize = 11,
                        Foreground = new SolidColorBrush(Color.Parse("#88ccff")),
                        VerticalAlignment = VerticalAlignment.Center },
                }
            });

        var conflictPanel = new StackPanel { Spacing = 6, IsVisible = false };

        var saveBtn = new Button
        {
            Content         = "\ud83d\udcbe \uacb0\uacfc \uc77c\uad04 \uc800\uc7a5",
            FontFamily      = Font, FontSize = 11,
            Background      = new SolidColorBrush(Color.Parse("#2a5a2a")),
            Foreground      = Avalonia.Media.Brushes.WhiteSmoke,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(12, 5),
            IsEnabled       = _importHasPending,
            IsVisible       = _importHasPending,
        };

        if (_importHasConflict && _conflictRows.Count > 0)
            BuildConflictUI(conflictPanel, saveBtn);

        saveBtn.Click += (_, _) =>
        {
            if (_pendingByRow.Count == 0) return;
            int totalSaved = 0, totalFailed = 0;
            foreach (var kv in _pendingByRow)
            {
                try   { totalSaved += TestReportService.BulkUpdateResults(kv.Key, kv.Value); }
                catch { totalFailed++; }
            }
            _importStatus    = $"\u2705 {totalSaved}\uac1c \ud56d\ubaa9 \uc800\uc7a5 \uc644\ub8cc" + (totalFailed > 0 ? $" ({totalFailed}\ud589 \uc2e4\ud328)" : "");
            _importStatusHex = "#44aa44";
            statusTb.Text       = _importStatus;
            statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
            _importHasPending   = false;
            _importHasConflict  = false;
            _pendingByRow.Clear();
            _conflictRows.Clear();
            saveBtn.IsEnabled = false;
            saveBtn.IsVisible = false;
            LoadData();
        };

        var fileBtn = new Button
        {
            Content         = "\ud83d\udcc2  Excel \uacb0\uacfc \ud30c\uc77c \uc120\ud0dd",
            FontFamily      = Font, FontSize = 11,
            Background      = new SolidColorBrush(Color.Parse("#1a3a5a")),
            Foreground      = Avalonia.Media.Brushes.WhiteSmoke,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(12, 6),
        };

        fileBtn.Click += async (_, _) =>
        {
            var dlg = new Avalonia.Platform.Storage.FilePickerOpenOptions
            {
                Title         = "\ubd84\uc11d\uae30\ub85d\ubd80 Excel \ud30c\uc77c \uc120\ud0dd",
                AllowMultiple = true,
                FileTypeFilter = new[]
                {
                    new Avalonia.Platform.Storage.FilePickerFileType("Excel \ud30c\uc77c")
                    {
                        Patterns = new[] { "*.xlsx" }
                    }
                },
            };
            var top = TopLevel.GetTopLevel(this);
            if (top == null) return;
            var files = await top.StorageProvider.OpenFilePickerAsync(dlg);
            if (files.Count == 0) return;

            _importFileNames = files.Select(f => System.IO.Path.GetFileName(f.Path.LocalPath)).ToList();
            fileListPanel.Children.Clear();
            foreach (var fn in _importFileNames)
                fileListPanel.Children.Add(new StackPanel
                {
                    Orientation = Orientation.Horizontal, Spacing = 5,
                    Children =
                    {
                        new TextBlock { Text = "\ud83d\udcc5", FontSize = 13, VerticalAlignment = VerticalAlignment.Center },
                        new TextBlock { Text = fn, FontFamily = Font, FontSize = 11,
                            Foreground = new SolidColorBrush(Color.Parse("#88ccff")),
                            VerticalAlignment = VerticalAlignment.Center },
                    }
                });

            _importStatus    = "\ud30c\uc77c \uc77d\ub294 \uc911...";
            _importStatusHex = "#888888";
            statusTb.Text       = _importStatus;
            statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
            _importRows.Clear();
            _pendingByRow.Clear();
            _conflictRows.Clear();
            _importHasPending   = false;
            _importHasConflict  = false;
            saveBtn.IsEnabled   = false;
            saveBtn.IsVisible   = false;
            conflictPanel.IsVisible = false;
            conflictPanel.Children.Clear();

            var rows = new List<ExcelResultRow>();
            foreach (var f in files)
            {
                try
                {
                    var partial = await Task.Run(() => AnalysisRecordService.ReadResultsFromFile(f.Path.LocalPath));
                    rows.AddRange(partial);
                }
                catch (Exception ex)
                {
                    _importStatus    = $"\u274c \ud30c\uc77c \uc77d\uae30 \uc2e4\ud328: {ex.Message}";
                    _importStatusHex = "#cc4444";
                    statusTb.Text       = _importStatus;
                    statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
                    return;
                }
            }

            if (rows.Count == 0)
            {
                _importStatus    = $"\u26a0 \ub370\uc774\ud130 \uc5c6\uc74c ({files.Count}\uac1c \ud30c\uc77c)";
                _importStatusHex = "#cc8800";
                statusTb.Text       = _importStatus;
                statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
                ResultListChanged?.Invoke(BuildListControl());
                return;
            }

            int matchedCnt = 0, notFoundCnt = 0;
            bool odd = false;
            foreach (var r in rows)
            {
                if (string.IsNullOrEmpty(r.\uacb0\uacfc\uac12)) continue;
                int? rowId = await Task.Run(() =>
                    TestReportService.FindRowId(r.\uacac\uc801\ubc88\ud638, r.\uc57d\uce6d, r.\uc2dc\ub8cc\uba85));

                bool   found       = rowId.HasValue;
                string existingVal = "";
                bool   hasConflict = false;

                if (found)
                {
                    existingVal = await Task.Run(() =>
                        TestReportService.GetAnalyteValue(rowId!.Value, r.AnalyteName)) ?? "";
                    hasConflict = !string.IsNullOrEmpty(existingVal)
                        && !string.Equals(existingVal, "O", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(existingVal, r.\uacb0\uacfc\uac12, StringComparison.Ordinal);
                    if (!_pendingByRow.ContainsKey(rowId!.Value))
                        _pendingByRow[rowId!.Value] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    _pendingByRow[rowId!.Value][r.AnalyteName] = r.\uacb0\uacfc\uac12;
                    if (hasConflict)
                    {
                        _conflictRows.Add((rowId!.Value, r.AnalyteName, existingVal, r.\uacb0\uacfc\uac12));
                        _importHasConflict = true;
                    }
                    matchedCnt++;
                }
                else notFoundCnt++;

                string bg  = hasConflict ? "#3a2a10" : (odd ? "#1e1e2e" : "#252535");
                string fgR = found ? (hasConflict ? "#ffaa44" : "#88ee88") : "#cc6644";
                odd = !odd;
                _importRows.Add((r.AnalyteName, r.\uacac\uc801\ubc88\ud638, r.\uc57d\uce6d, r.\uc2dc\ub8cc\uba85,
                    existingVal, r.\uacb0\uacfc\uac12, r.\ubc29\ub958\uae30\uc900, bg, fgR));
            }

            int withValue = rows.Count(x => !string.IsNullOrEmpty(x.\uacb0\uacfc\uac12));
            if (matchedCnt > 0)
            {
                string conflictNote = _conflictRows.Count > 0 ? $"  (\u26a0 {_conflictRows.Count}\uac1c \ucda9\ub3cc)" : "";
                _importStatus    = $"\u2705 {matchedCnt}\uac1c \ub9e4\uce6d{conflictNote}"
                    + (notFoundCnt > 0 ? $", {notFoundCnt}\uac1c \ubbf8\ub9e4\uce6d" : "")
                    + $"  \u2190 {files.Count}\uac1c \ud30c\uc77c";
                _importStatusHex  = _conflictRows.Count > 0 ? "#ffaa44" : "#44aa44";
                _importHasPending = true;
                saveBtn.IsEnabled = true;
                saveBtn.IsVisible = true;
                if (_importHasConflict)
                {
                    conflictPanel.Children.Clear();
                    BuildConflictUI(conflictPanel, saveBtn);
                    conflictPanel.IsVisible = true;
                }
            }
            else
            {
                _importStatus    = $"\u26a0 \ub9e4\uce6d\ub41c \uc2dc\ub8cc \uc5c6\uc74c ({withValue}\ud589 \uc77d\uc74c)";
                _importStatusHex = "#cc8800";
            }
            statusTb.Text       = _importStatus;
            statusTb.Foreground = new SolidColorBrush(Color.Parse(_importStatusHex));
            ResultListChanged?.Invoke(BuildListControl());
        };

        return new Border
        {
            Padding = new Thickness(10),
            Child   = new StackPanel
            {
                Spacing  = 8,
                Children =
                {
                    new TextBlock
                    {
                        Text       = "\ud83d\udcca  Excel \uacb0\uacfc \ubd88\ub7ec\uc624\uae30",
                        FontFamily = Font, FontSize = 13,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = Avalonia.Media.Brushes.WhiteSmoke,
                    },
                    new Border { Height = 1, Background = new SolidColorBrush(Color.Parse("#444")) },
                    fileBtn,
                    fileListPanel,
                    statusTb,
                    conflictPanel,
                    saveBtn,
                }
            }
        };
    }

    private void BuildConflictUI(StackPanel conflictPanel, Button saveBtn)
    {
        conflictPanel.Children.Add(new TextBlock
        {
            Text         = $"\u26a0  {_conflictRows.Count}\uac1c \ud56d\ubaa9\uc774 \uae30\uc874 \uacb0\uacfc\uc640 \ub2e4\ub985\ub2c8\ub2e4.",
            FontFamily   = Font, FontSize = 11,
            Foreground   = new SolidColorBrush(Color.Parse("#ffaa44")),
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
        });
        var btnRow       = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var skipBtn      = new Button
        {
            Content         = "\ud83d\udeab \ucda9\ub3cc \ud56d\ubaa9 \uac74\ub108\ubc84\ub9bc",
            FontFamily      = Font, FontSize = 11,
            Background      = new SolidColorBrush(Color.Parse("#4a2a10")),
            Foreground      = Avalonia.Media.Brushes.WhiteSmoke,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };
        var overwriteBtn = new Button
        {
            Content         = "\u270f \ub36e\uc5b4\uc4f0\uae30",
            FontFamily      = Font, FontSize = 11,
            Background      = new SolidColorBrush(Color.Parse("#2a3a5a")),
            Foreground      = Avalonia.Media.Brushes.WhiteSmoke,
            BorderThickness = new Thickness(0), CornerRadius = new CornerRadius(4),
            Padding         = new Thickness(10, 4),
        };
        skipBtn.Click += (_, _) =>
        {
            foreach (var c in _conflictRows)
                if (_pendingByRow.TryGetValue(c.RowId, out var d)) d.Remove(c.Analyte);
            _importHasConflict      = false;
            conflictPanel.IsVisible = false;
            saveBtn.Content         = "\ud83d\udcbe \uacb0\uacfc \uc800\uc7a5 (\ucda9\ub3cc \uc81c\uc678)";
        };
        overwriteBtn.Click += (_, _) =>
        {
            _importHasConflict      = false;
            conflictPanel.IsVisible = false;
            saveBtn.Content         = "\ud83d\udcbe \uacb0\uacfc \uc77c\uad04 \uc800\uc7a5";
        };
        btnRow.Children.Add(skipBtn);
        btnRow.Children.Add(overwriteBtn);
        conflictPanel.Children.Add(btnRow);
    }

`;
  s = s.slice(0,a) + nm + s.slice(b);
  console.log('4. BuildImportActionPanel OK');
}

fs.writeFileSync(p, s, 'utf8');
console.log('Done. Length:', s.length);
