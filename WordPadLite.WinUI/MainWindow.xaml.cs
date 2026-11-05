using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Text;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation;
using Windows.Graphics;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.System;
using Windows.UI;
using WinRT.Interop;
using WordPadLite.WinUI.Services;
using IOPath = System.IO.Path;

namespace WordPadLite.WinUI;

public sealed partial class MainWindow : Window
{
    private const double A4Width = 794;
    private const double A4Height = 1123;
    private const double PixelsPerCm = 96 / 2.54;
    private static readonly Color NoHighlightColor = Color.FromArgb(255, 255, 255, 255);

    private readonly ToggleButton[] _tabs;
    private readonly Dictionary<string, StackPanel> _ribbonPanels;
    private readonly Dictionary<string, StackPanel> _compactPanels = new();
    private readonly Dictionary<string, RibbonGroupSpec[]> _tabGroups = new();
    private readonly Dictionary<ToggleButton, ToggleButton> _compactToggleMap = new();
    private readonly Dictionary<string, List<Button>> _ribbonCommandButtons = new(StringComparer.OrdinalIgnoreCase);

    private bool _editorReady;
    private bool _isLoading;
    private bool _isModified;
    private bool _isUpdatingFormatUi;
    private bool _isUpdatingZoom;
    private bool _isClosingConfirmed;
    private bool _readingMode;
    private bool _ribbonCollapsed;
    private bool _useCompactRibbon;
    private string _activeTab = "home";
    private bool _rulerVisible = true;
    private bool _gridLinesVisible;
    private bool _spellCheckEnabled = true;
    private bool _trackChanges;
    private bool _formatPainterArmed;
    private bool _applyingFormatPainter;
    private bool _autoSaveEnabled;
    private bool _isAutoSaving;
    private bool _showToasts;
    private bool _loadingSettings;
    private bool _updatingProofingPane;
    private bool _updatingBackstage;
    private int _autoSaveIntervalSeconds = 30;
    private DateTime? _lastAutoSaveAt;
    private DocumentMode _documentMode = DocumentMode.Edit;
    private string _themePreference = "system";
    private string _defaultFontName = "游明朝";
    private double _defaultFontSize = 10.5;
    private string _accountName = Environment.UserName;
    private string _accountRole = "ローカル アカウント";
    private string _accountImage = string.Empty;

    private ITextCharacterFormat? _formatPainterCharacter;
    private ITextParagraphFormat? _formatPainterParagraph;

    private StorageFile? _currentFile;
    private string _currentFormat = "RTF";
    private string _untitledName = "文書 1";
    private int _untitledNumber = 1;
    private int _footnoteNumber;

    private double _pageWidth = A4Width;
    private double _pageHeight = A4Height;
    private double _pageMargin = 72;
    private string _pageSizeName = "A4";
    private bool _landscape;

    private Microsoft.UI.Dispatching.DispatcherQueueTimer? _wordCountTimer;
    private Microsoft.UI.Dispatching.DispatcherQueueTimer? _ribbonComboTimer;
    private Microsoft.UI.Dispatching.DispatcherQueueTimer? _autoSaveTimer;
    private Microsoft.UI.Dispatching.DispatcherQueueTimer? _toastTimer;
    private Microsoft.UI.Dispatching.DispatcherQueueTimer? _statusMessageTimer;
    private readonly Dictionary<string, Button> _quickAccessButtons = new();
    private List<string> _quickAccessItems = new();
    private FlyoutBase? _activeFlyout;

    private enum DocumentMode
    {
        Edit,
        Review,
        ReadOnly,
    }

    private sealed record QuickAccessCommand(string Id, string Label, string Glyph, string Command, string Tip);

    private static readonly QuickAccessCommand[] QuickAccessCatalog =
    {
        new("save", "保存", "\uE74E", "save", "保存 (Ctrl+S)"),
        new("undo", "元に戻す", "\uE7A7", "undo", "元に戻す (Ctrl+Z)"),
        new("redo", "やり直す", "\uE7A6", "redo", "やり直す (Ctrl+Y)"),
        new("new", "新規作成", "\uE8A5", "new", "新規作成 (Ctrl+N)"),
        new("open", "開く", "\uE838", "open", "開く (Ctrl+O)"),
        new("print", "印刷", "\uE749", "print", "印刷 (Ctrl+P)"),
        new("find", "検索", "\uE721", "find", "検索 (Ctrl+F)"),
        new("replace", "置換", "\uE8AC", "replace", "置換 (Ctrl+H)"),
        new("spell", "スペルチェック", "\uE8AC", "review-spelling", "スペルチェック"),
        new("zoom-in", "拡大", "\uE8A3", "zoom-in", "拡大"),
        new("zoom-out", "縮小", "\uE71F", "zoom-out", "縮小"),
        new("reading", "閲覧モード", "\uED5E", "view-reading", "閲覧モード"),
    };

    private sealed record RibbonItemSpec(
        string Label,
        string Glyph,
        string Command,
        string? Tip = null,
        IReadOnlyList<(string Label, string Command)>? Menu = null);

    private sealed record RibbonGroupSpec(string Label, params RibbonItemSpec[] Items);

    public MainWindow()
    {
        InitializeComponent();

        _tabs = new[] { FileTab, HomeTab, InsertTab, DesignTab, LayoutTab, ReferencesTab, MailingsTab, ReviewTab, ViewTab, HelpTab };
        _ribbonPanels = new Dictionary<string, StackPanel>
        {
            ["home"] = HomeRibbon,
            ["insert"] = InsertRibbon,
            ["design"] = DesignRibbon,
            ["layout"] = LayoutRibbon,
            ["references"] = ReferencesRibbon,
            ["mailings"] = MailingsRibbon,
            ["review"] = ReviewRibbon,
            ["view"] = ViewRibbon,
            ["help"] = HelpRibbon,
        };

        _compactPanels["home"] = HomeCompactRibbon;
        foreach (var tag in _ribbonPanels.Keys)
        {
            if (tag == "home")
            {
                continue;
            }

            var compactPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Height = 44,
                Visibility = Visibility.Collapsed,
            };
            RibbonPanelHost.Children.Add(compactPanel);
            _compactPanels[tag] = compactPanel;
        }

        CopyComboItems(FontCombo, CompactFontCombo);
        CopyComboItems(FontSizeCombo, CompactFontSizeCombo);
        CompactFontColorButton.Flyout = BuildFontColorFlyout();

        _compactToggleMap[BoldToggle] = CompactBoldToggle;
        _compactToggleMap[ItalicToggle] = CompactItalicToggle;
        _compactToggleMap[UnderlineToggle] = CompactUnderlineToggle;
        _compactToggleMap[AlignLeftToggle] = CompactAlignLeftToggle;
        _compactToggleMap[AlignCenterToggle] = CompactAlignCenterToggle;
        _compactToggleMap[AlignRightToggle] = CompactAlignRightToggle;
        _compactToggleMap[AlignJustifyToggle] = CompactAlignJustifyToggle;
        _compactToggleMap[BulletsToggle] = CompactBulletsToggle;
        _compactToggleMap[NumberingToggle] = CompactNumberingToggle;
        _useCompactRibbon = ReadCompactRibbonPreference();

        _loadingSettings = true;
        try
        {
            _themePreference = AppSettings.Get("Theme", "system");
            _autoSaveEnabled = AppSettings.GetBool("AutoSave", false);
            _autoSaveIntervalSeconds = Math.Max(5, (int)AppSettings.GetDouble("AutoSaveInterval", 30));
            _showToasts = AppSettings.GetBool("ShowToasts", false);
            _spellCheckEnabled = AppSettings.GetBool("SpellCheck", true);
            _defaultFontName = AppSettings.Get("DefaultFont", "游明朝");
            _defaultFontSize = Math.Max(6, AppSettings.GetDouble("DefaultFontSize", 10.5));
            _accountName = AppSettings.Get("AccountName", Environment.UserName);
            _accountRole = AppSettings.Get("AccountRole", "ローカル アカウント");
            _accountImage = AppSettings.Get("AccountImage", string.Empty);
            _quickAccessItems = AppSettings.Get("QuickAccess", "save,undo,redo")
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(id => QuickAccessCatalog.Any(spec => spec.Id == id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        finally
        {
            _loadingSettings = false;
        }

        ApplyThemePreference();
        BuildQuickAccess();
        BuildSecondaryRibbons();
        UpdateRibbonPanels();
        BuildKeyboardAccelerators();
        ConfigureWindow();
        WireToggleVisuals(
            BoldToggle,
            ItalicToggle,
            UnderlineToggle,
            StrikeToggle,
            SubscriptToggle,
            SuperscriptToggle,
            BulletsToggle,
            NumberingToggle,
            AlignLeftToggle,
            AlignCenterToggle,
            AlignRightToggle,
            AlignJustifyToggle,
            CompactBoldToggle,
            CompactItalicToggle,
            CompactUnderlineToggle,
            CompactAlignLeftToggle,
            CompactAlignCenterToggle,
            CompactAlignRightToggle,
            CompactAlignJustifyToggle,
            CompactBulletsToggle,
            CompactNumberingToggle);

        _isUpdatingZoom = true;
        try
        {
            ZoomSlider.Maximum = 250;
            ZoomSlider.Minimum = 50;
            ZoomSlider.StepFrequency = 10;
            ZoomSlider.Value = 100;
        }
        finally
        {
            _isUpdatingZoom = false;
        }

        AppWindow.Closing += AppWindow_Closing;
        RootGrid.Loaded += RootGrid_Loaded;
        RibbonScrollViewer.PointerWheelChanged += RibbonScrollViewer_PointerWheelChanged;
    }

    private void RibbonScrollViewer_PointerWheelChanged(object sender, PointerRoutedEventArgs e)
    {
        if (RibbonScrollViewer.ScrollableWidth <= 0)
        {
            return;
        }

        var delta = -e.GetCurrentPoint(RibbonScrollViewer).Properties.MouseWheelDelta;
        var target = Math.Clamp(RibbonScrollViewer.HorizontalOffset + delta, 0, RibbonScrollViewer.ScrollableWidth);
        if (Math.Abs(target - RibbonScrollViewer.HorizontalOffset) < 0.5)
        {
            return;
        }

        RibbonScrollViewer.ChangeView(target, null, null, disableAnimation: true);
        e.Handled = true;
    }

    private void ApplyTitleBarTheme()
    {
        try
        {
            // The title bar is always Word blue, in every theme.
            var titleBar = AppWindow.TitleBar;
            titleBar.ButtonBackgroundColor = Colors.Transparent;
            titleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
            titleBar.ButtonForegroundColor = Colors.White;
            titleBar.ButtonInactiveForegroundColor = Color.FromArgb(165, 255, 255, 255);
            titleBar.ButtonHoverBackgroundColor = Color.FromArgb(48, 255, 255, 255);
            titleBar.ButtonHoverForegroundColor = Colors.White;
            titleBar.ButtonPressedBackgroundColor = Color.FromArgb(74, 255, 255, 255);
            titleBar.ButtonPressedForegroundColor = Colors.White;
        }
        catch
        {
            // Title bar decoration failures must not prevent editing.
        }
    }

    private void ConfigureWindow()
    {
        try
        {
            ExtendsContentIntoTitleBar = true;
            SetTitleBar(TitleBarDragSurface);

            if (MicaController.IsSupported())
            {
                SystemBackdrop = new MicaBackdrop { Kind = MicaKind.BaseAlt };
                RootGrid.Background = new SolidColorBrush(Colors.Transparent);
            }

            ApplyTitleBarTheme();

            var display = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Primary);
            var work = display.WorkArea;
            var width = Math.Max(980, Math.Min(1800, work.Width - 40));
            var height = Math.Max(720, Math.Min(1000, work.Height - 40));
            AppWindow.Resize(new SizeInt32(width, height));
            AppWindow.Move(new PointInt32(
                work.X + Math.Max(0, (work.Width - width) / 2),
                work.Y + Math.Max(0, (work.Height - height) / 2)));

            var iconPath = IOPath.Combine(AppContext.BaseDirectory, "Assets", "WordPadLite.ico");
            if (File.Exists(iconPath))
            {
                AppWindow.SetIcon(iconPath);
            }
        }
        catch
        {
            // Window decoration failures must not prevent editing.
        }
    }

    private void WireToggleVisuals(params ToggleButton[] toggles)
    {
        foreach (var toggle in toggles)
        {
            toggle.Checked += ToggleVisual_Changed;
            toggle.Unchecked += ToggleVisual_Changed;
            ApplyToggleVisual(toggle);
        }
    }

    private void ToggleVisual_Changed(object sender, RoutedEventArgs e)
    {
        if (sender is ToggleButton toggle)
        {
            ApplyToggleVisual(toggle);
        }
    }

    private void ApplyToggleVisual(ToggleButton toggle)
    {
        var active = toggle.IsChecked == true;
        var accent = (Brush)Application.Current.Resources["WordAccentBrush"];
        toggle.Background = active
            ? new SolidColorBrush(IsDarkTheme
                ? Color.FromArgb(255, 24, 62, 94)
                : Color.FromArgb(255, 219, 235, 250))
            : new SolidColorBrush(Colors.Transparent);
        toggle.BorderBrush = active ? accent : new SolidColorBrush(Colors.Transparent);
        toggle.BorderThickness = active ? new Thickness(1) : new Thickness(0);
        if (active)
        {
            toggle.Foreground = accent;
        }
        else
        {
            toggle.ClearValue(Control.ForegroundProperty);
        }
    }

    private bool IsDarkTheme => RootGrid.ActualTheme == ElementTheme.Dark;

    private void ApplyThemePreference()
    {
        RootGrid.RequestedTheme = _themePreference switch
        {
            "light" => ElementTheme.Light,
            "dark" => ElementTheme.Dark,
            _ => ElementTheme.Default,
        };
        ApplyTitleBarTheme();
        DrawRuler();
        RefreshRibbonToggleStates();
        RefreshProofingPane();
    }

    private void SetThemePreference(string preference, bool notify = true)
    {
        _themePreference = preference is "light" or "dark" ? preference : "system";
        AppSettings.Set("Theme", _themePreference);
        ApplyThemePreference();
        if (notify)
        {
            Notify(_themePreference switch
            {
                "light" => "ライト テーマに切り替えました。",
                "dark" => "ダーク テーマに切り替えました。",
                _ => "システムのテーマ設定に合わせます。",
            });
        }
    }

    private void RootGrid_Loaded(object sender, RoutedEventArgs e)
    {
        ApplyTitleBarTheme();
        ApplyRibbonIconColors(HomeRibbon);
        ApplyRibbonIconColors(HomeCompactRibbon);
        DispatcherQueue.TryEnqueue(RefreshRibbonToggleStates);
        RootGrid.ActualThemeChanged += RootGrid_ActualThemeChanged;

        _loadingSettings = true;
        try
        {
            AutoSaveSwitch.IsOn = _autoSaveEnabled;
        }
        finally
        {
            _loadingSettings = false;
        }

        SetDocumentMode(_documentMode, notify: false);
        HookFlyouts(RootGrid);
        DrawRuler();
        UpdateWordCount();
        UpdateTitle();
        UpdateTabSelection("home");
        UpdateFormattingUi();
        UpdateZoomText(PageScrollViewer.ZoomFactor);
    }

    private void BuildKeyboardAccelerators()
    {
        // Accelerators live on the root grid, which spans the whole window. The
        // default Auto placement would surface a "Ctrl+N" tooltip anywhere the
        // pointer rests over the window, so keep those hints out of tooltips.
        RootGrid.KeyboardAcceleratorPlacementMode = KeyboardAcceleratorPlacementMode.Hidden;
        AddAccelerator(VirtualKey.N, VirtualKeyModifiers.Control, "new");
        AddAccelerator(VirtualKey.O, VirtualKeyModifiers.Control, "open");
        AddAccelerator(VirtualKey.S, VirtualKeyModifiers.Control, "save");
        AddAccelerator(VirtualKey.S, VirtualKeyModifiers.Control | VirtualKeyModifiers.Shift, "save-as");
        AddAccelerator(VirtualKey.P, VirtualKeyModifiers.Control, "print");
        AddAccelerator(VirtualKey.B, VirtualKeyModifiers.Control, "bold");
        AddAccelerator(VirtualKey.I, VirtualKeyModifiers.Control, "italic");
        AddAccelerator(VirtualKey.U, VirtualKeyModifiers.Control, "underline");
        AddAccelerator(VirtualKey.Z, VirtualKeyModifiers.Control, "undo");
        AddAccelerator(VirtualKey.Y, VirtualKeyModifiers.Control, "redo");
        AddAccelerator(VirtualKey.F, VirtualKeyModifiers.Control, "find");
        AddAccelerator(VirtualKey.H, VirtualKeyModifiers.Control, "replace");
        AddAccelerator(VirtualKey.A, VirtualKeyModifiers.Control, "select-all");
        AddAccelerator(VirtualKey.Number0, VirtualKeyModifiers.Control, "zoom-reset");
        AddAccelerator(VirtualKey.Add, VirtualKeyModifiers.Control, "zoom-in");
        AddAccelerator(VirtualKey.Subtract, VirtualKeyModifiers.Control, "zoom-out");
    }

    private void AddAccelerator(VirtualKey key, VirtualKeyModifiers modifiers, string command)
    {
        var accelerator = new KeyboardAccelerator
        {
            Key = key,
            Modifiers = modifiers,
        };
        accelerator.Invoked += async (_, args) =>
        {
            args.Handled = true;
            await ExecuteCommandAsync(command);
        };
        RootGrid.KeyboardAccelerators.Add(accelerator);
    }

    private void RibbonTab_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not ToggleButton tab || tab.Tag is not string tag)
        {
            return;
        }

        if (tag == "file")
        {
            ShowBackstage();
            return;
        }

        HideBackstage(false);
        SetActiveTab(tab);

        if (_ribbonCollapsed)
        {
            SetRibbonCollapsed(false);
        }
    }

    private void RibbonToggle_Click(object sender, RoutedEventArgs e)
    {
        SetRibbonCollapsed(!_ribbonCollapsed);
    }

    private void SetRibbonCollapsed(bool collapsed)
    {
        _ribbonCollapsed = collapsed;
        RibbonHost.Visibility = collapsed ? Visibility.Collapsed : Visibility.Visible;
        RibbonToggleButton.Content = collapsed ? "\uE70D" : "\uE70E";
        ToolTipService.SetToolTip(RibbonToggleButton, collapsed ? "リボンを展開" : "リボンを折りたたむ");
    }

    private void SetActiveTab(ToggleButton tab)
    {
        _activeTab = tab.Tag as string ?? "home";
        UpdateTabSelection(_activeTab);
        UpdateRibbonPanels();
        UpdateFormattingUi();
        PlayEntrance(GetActiveRibbonPanel(), 6, 160);

        // The panel becomes visible now, so refresh toggle visuals after layout
        // has realized the ribbon buttons (icons and active marks).
        DispatcherQueue.TryEnqueue(RefreshRibbonToggleStates);

        try
        {
            RibbonScrollViewer.ChangeView(0, null, null, true);
        }
        catch
        {
            // Resetting the scroll offset is cosmetic.
        }
    }

    private FrameworkElement? GetActiveRibbonPanel()
    {
        if (_useCompactRibbon && _compactPanels.TryGetValue(_activeTab, out var compact))
        {
            return compact;
        }

        return _ribbonPanels.TryGetValue(_activeTab, out var classic) ? classic : null;
    }

    private void UpdateRibbonPanels()
    {
        foreach (var (key, panel) in _ribbonPanels)
        {
            panel.Visibility = !_useCompactRibbon && key == _activeTab
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        foreach (var (key, panel) in _compactPanels)
        {
            panel.Visibility = _useCompactRibbon && key == _activeTab
                ? Visibility.Visible
                : Visibility.Collapsed;
        }
    }

    private static string RibbonModeSettingsPath => System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WordPadLite",
        "ribbon-mode.txt");

    private static bool ReadCompactRibbonPreference()
    {
        try
        {
            var path = RibbonModeSettingsPath;
            if (System.IO.File.Exists(path))
            {
                var saved = System.IO.File.ReadAllText(path).Trim();
                return !string.Equals(saved, "classic", StringComparison.OrdinalIgnoreCase);
            }
        }
        catch
        {
            // A missing or unreadable settings file simply means "use the default".
        }

        return false;
    }

    private static void SaveRibbonModePreference(bool compact)
    {
        try
        {
            var path = RibbonModeSettingsPath;
            var folder = System.IO.Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
            }

            System.IO.File.WriteAllText(path, compact ? "compact" : "classic");
        }
        catch
        {
            // Persisting the preference is optional.
        }
    }

    private void RefreshRibbonCombosAfterLayout()
    {
        // Editable ComboBoxes that were collapsed at startup apply their template after being shown,
        // so the text has to be pushed again once layout has finished.
        _ribbonComboTimer ??= DispatcherQueue.CreateTimer();
        _ribbonComboTimer.Tick -= RibbonComboTimer_Tick;
        _ribbonComboTimer.Tick += RibbonComboTimer_Tick;
        _ribbonComboTimer.Interval = TimeSpan.FromMilliseconds(120);
        _ribbonComboTimer.IsRepeating = false;
        _ribbonComboTimer.Stop();
        _ribbonComboTimer.Start();
    }

    private void RibbonComboTimer_Tick(Microsoft.UI.Dispatching.DispatcherQueueTimer sender, object args)
    {
        FontCombo.Text = string.Empty;
        FontSizeCombo.Text = string.Empty;
        CompactFontCombo.Text = string.Empty;
        CompactFontSizeCombo.Text = string.Empty;
        UpdateFormattingUi();
    }

    private void SetRibbonMode(bool compact)
    {
        if (_useCompactRibbon == compact)
        {
            return;
        }

        _useCompactRibbon = compact;
        UpdateRibbonPanels();
        RefreshRibbonCombosAfterLayout();
        PlayEntrance(GetActiveRibbonPanel(), 6, 160);

        SaveRibbonModePreference(compact);
    }

    private void UpdateTabSelection(string tag)
    {
        var accent = (Brush)Application.Current.Resources["WordAccentBrush"];

        foreach (var candidate in _tabs)
        {
            var selected = string.Equals(candidate.Tag as string, tag, StringComparison.Ordinal);
            candidate.IsChecked = selected;

            // The File tab keeps Word's deep-blue look; only the other tabs use the accent text.
            if (ReferenceEquals(candidate, FileTab))
            {
                candidate.Foreground = new SolidColorBrush(Colors.White);
                candidate.BorderBrush = new SolidColorBrush(Colors.Transparent);
                continue;
            }

            if (selected)
            {
                candidate.Foreground = accent;
            }
            else
            {
                candidate.ClearValue(Control.ForegroundProperty);
            }

            candidate.BorderBrush = selected ? accent : new SolidColorBrush(Colors.Transparent);
        }
    }

    private void ShowBackstage()
    {
        UpdateTabSelection("file");
        ApplyBackstageChrome();
        ShowBackstagePage("info");
        FileBackstage.Visibility = Visibility.Visible;
        PlayEntrance(FileBackstage, 10, 180);
    }

    private void HideBackstage(bool restoreHome = true)
    {
        var wasOpen = FileBackstage.Visibility == Visibility.Visible;
        FileBackstage.Visibility = Visibility.Collapsed;
        if (wasOpen)
        {
            ClearBackstageChrome();
        }

        if (restoreHome && wasOpen)
        {
            SetActiveTab(HomeTab);
        }
    }

    private void ShowBackstagePage(string tag)
    {
        HighlightBackstageNav(tag);
        UpdateBackstageInfo();
        BackstageInfoPanel.Visibility = tag == "account" ? Visibility.Collapsed : Visibility.Visible;
        BackstageAccountPanel.Visibility = tag == "account" ? Visibility.Visible : Visibility.Collapsed;
        if (tag == "account")
        {
            UpdateAccountUi();
        }
    }

    private static string AccountImageFolder => IOPath.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WordPadLite");

    private void UpdateAccountUi()
    {
        AccountUserName.Text = string.IsNullOrWhiteSpace(_accountName) ? Environment.UserName : _accountName;
        AccountRoleText.Text = string.IsNullOrWhiteSpace(_accountRole) ? "ローカル アカウント" : _accountRole;

        if (!string.IsNullOrWhiteSpace(_accountImage) && File.Exists(_accountImage))
        {
            try
            {
                var source = new BitmapImage(new Uri(_accountImage));
                var brush = new ImageBrush
                {
                    ImageSource = source,
                    Stretch = Stretch.UniformToFill,
                };
                AccountAvatarEllipse.Fill = brush;
                AccountAvatarIcon.Visibility = Visibility.Collapsed;
                TitleBarAvatarEllipse.Fill = brush;
                TitleBarAvatarIcon.Visibility = Visibility.Collapsed;
                return;
            }
            catch
            {
                // Fall back to the default glyph when the image cannot be decoded.
            }
        }

        AccountAvatarEllipse.Fill = new SolidColorBrush(RootGrid.ActualTheme == ElementTheme.Dark
            ? Color.FromArgb(255, 20, 58, 87)
            : Color.FromArgb(255, 229, 241, 251));
        AccountAvatarIcon.Visibility = Visibility.Visible;
        TitleBarAvatarEllipse.Fill = new SolidColorBrush(Color.FromArgb(64, 255, 255, 255));
        TitleBarAvatarIcon.Visibility = Visibility.Visible;
    }

    private async void EditAccountProfile_Click(object sender, RoutedEventArgs e)
    {
        var nameBox = new TextBox
        {
            Header = "名前",
            Text = _accountName,
            PlaceholderText = "表示名",
            IsSpellCheckEnabled = false,
        };
        var roleBox = new TextBox
        {
            Header = "役職",
            Text = _accountRole,
            PlaceholderText = "例: 個人利用 / 編集担当",
            IsSpellCheckEnabled = false,
        };
        var note = new TextBlock
        {
            Text = "この情報はこの PC にのみ保存され、外部には送信されません。",
            TextWrapping = TextWrapping.Wrap,
            FontSize = 12,
            Foreground = (Brush)Application.Current.Resources["MutedInkBrush"],
        };
        var panel = new StackPanel { Spacing = 12, Width = 380 };
        panel.Children.Add(nameBox);
        panel.Children.Add(roleBox);
        panel.Children.Add(note);

        var dialog = new ContentDialog
        {
            Title = "プロフィールを編集",
            Content = panel,
            PrimaryButtonText = "保存",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        if (await ShowDialogAsync(dialog) != ContentDialogResult.Primary)
        {
            return;
        }

        _accountName = string.IsNullOrWhiteSpace(nameBox.Text) ? Environment.UserName : nameBox.Text.Trim();
        _accountRole = string.IsNullOrWhiteSpace(roleBox.Text) ? "ローカル アカウント" : roleBox.Text.Trim();
        AppSettings.Set("AccountName", _accountName);
        AppSettings.Set("AccountRole", _accountRole);
        UpdateAccountUi();
        Notify("プロフィールを更新しました。");
    }

    private async void ChangeAccountImage_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FileOpenPicker
        {
            ViewMode = PickerViewMode.Thumbnail,
            SuggestedStartLocation = PickerLocationId.PicturesLibrary,
        };
        picker.FileTypeFilter.Add(".png");
        picker.FileTypeFilter.Add(".jpg");
        picker.FileTypeFilter.Add(".jpeg");
        picker.FileTypeFilter.Add(".bmp");
        picker.FileTypeFilter.Add(".gif");
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(AccountImageFolder);
            var target = IOPath.Combine(AccountImageFolder, "avatar" + IOPath.GetExtension(file.Name).ToLowerInvariant());
            foreach (var old in Directory.GetFiles(AccountImageFolder, "avatar.*"))
            {
                try
                {
                    File.Delete(old);
                }
                catch
                {
                    // A locked leftover image should not block choosing a new one.
                }
            }

            File.Copy(file.Path, target, overwrite: true);
            _accountImage = target;
            AppSettings.Set("AccountImage", _accountImage);
            UpdateAccountUi();
            Notify("プロフィール画像を変更しました。");
        }
        catch
        {
            Notify("画像を読み込めませんでした。別のファイルを選んでください。", InfoBarSeverity.Warning);
        }
    }

    private void ResetAccountProfile_Click(object sender, RoutedEventArgs e)
    {
        _accountName = Environment.UserName;
        _accountRole = "ローカル アカウント";
        _accountImage = string.Empty;
        AppSettings.Set("AccountName", _accountName);
        AppSettings.Set("AccountRole", _accountRole);
        AppSettings.Set("AccountImage", string.Empty);

        try
        {
            foreach (var old in Directory.GetFiles(AccountImageFolder, "avatar.*"))
            {
                File.Delete(old);
            }
        }
        catch
        {
            // Best effort cleanup only.
        }

        UpdateAccountUi();
        Notify("プロフィールを既定の状態に戻しました。");
    }

    private void ApplyBackstageChrome()
    {
        // The title bar is always Word blue. Opening the backstage only hides the
        // quick access toolbar and left-aligns the document name.
        WindowTitleText.Foreground = new SolidColorBrush(Colors.White);
        TitleBarTitleHost.HorizontalAlignment = HorizontalAlignment.Left;
        ModifiedDot.Visibility = Visibility.Collapsed;
        TitleBarAutoSavePanel.Visibility = Visibility.Collapsed;
        QuickAccessPanel.Visibility = Visibility.Collapsed;
        QuickAccessButton.Visibility = Visibility.Collapsed;
        ApplyTitleBarTheme();
    }

    private void ClearBackstageChrome()
    {
        WindowTitleText.Foreground = new SolidColorBrush(Colors.White);
        TitleBarTitleHost.HorizontalAlignment = HorizontalAlignment.Center;
        TitleBarAutoSavePanel.Visibility = Visibility.Visible;
        QuickAccessPanel.Visibility = Visibility.Visible;
        QuickAccessButton.Visibility = Visibility.Visible;
        ApplyTitleBarTheme();
        UpdateTitle();
    }

    private void TitleBarAccount_Click(object sender, RoutedEventArgs e)
    {
        ShowBackstage();
        ShowBackstagePage("account");
    }

    private void HighlightBackstageNav(string? tag)
    {
        var activeBrush = new SolidColorBrush(IsDarkTheme
            ? Color.FromArgb(255, 24, 62, 94)
            : Color.FromArgb(255, 226, 240, 253));
        var accent = (Brush)Application.Current.Resources["WordAccentBrush"];

        foreach (var button in BackstageNavPanel.Children.OfType<Button>())
        {
            var selected = tag is not null && string.Equals(button.Tag as string, tag, StringComparison.Ordinal);
            button.Background = selected ? activeBrush : new SolidColorBrush(Colors.Transparent);
            button.FontWeight = selected
                ? Microsoft.UI.Text.FontWeights.SemiBold
                : Microsoft.UI.Text.FontWeights.Normal;
            if (selected)
            {
                button.Foreground = accent;
            }
            else
            {
                button.ClearValue(Control.ForegroundProperty);
            }
        }
    }

    private void BackstageHome_Click(object sender, RoutedEventArgs e)
    {
        HideBackstage(true);
        Editor.Focus(FocusState.Programmatic);
    }

    private async void BackstageCommand_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not string tag)
        {
            return;
        }

        if (tag is "info" or "account")
        {
            ShowBackstagePage(tag);
            FileBackstage.Visibility = Visibility.Visible;
            return;
        }

        if (tag == "exit")
        {
            Close();
            return;
        }

        HideBackstage(true);
        await ExecuteCommandAsync(tag);
    }

    private async void RibbonCommand_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement element && element.Tag is string tag)
        {
            // Flyouts with custom swatch content do not close themselves when an
            // item is picked, so close the open popup before applying the command.
            _activeFlyout?.Hide();
            _activeFlyout = null;

            if (_editorReady)
            {
                // Restore the document selection before applying formatting so the
                // caret keeps the newly selected character formatting.
                Editor.Focus(FocusState.Programmatic);
            }

            await ExecuteCommandAsync(tag);
        }
    }

    private void BuildSecondaryRibbons()
    {
        BuildTab("insert", InsertRibbon,
            new RibbonGroupSpec("ページ",
                Item("表紙", "\uE7C3", "insert-cover", "表紙を挿入"),
                Item("白紙のページ", "\uE7C3", "insert-blank-page", "白紙のページを挿入"),
                Item("ページ区切り", "\uE8A5", "insert-page-break", "改ページを挿入"),
                Item("ページ番号", "\uE8AC", "insert-page-number", "ページ番号を挿入"),
                Item("ヘッダー", "\uE7C3", "insert-header", "ヘッダーを挿入"),
                Item("フッター", "\uE7C3", "insert-footer", "フッターを挿入")),
            new RibbonGroupSpec("表",
                Item("表", "\uE80A", "insert-table", "表を挿入"),
                Item("表の変換", "\uE8AB", "insert-table-convert", "テキストを表に変換")),
            new RibbonGroupSpec("図",
                Item("画像", "\uE8B9", "insert-picture", "画像を挿入"),
                Item("図形", "\uE7C4", "insert-shape", "図形を挿入"),
                Item("アイコン", "\uE8A9", "insert-icon", "アイコンを挿入"),
                Item("SmartArt", "\uE8A9", "insert-smartart", "SmartArt を挿入"),
                Item("グラフ", "\uE9D2", "insert-chart", "グラフを挿入"),
                Item("スクリーンショット", "\uE722", "insert-screenshot", "スクリーンショットを挿入")),
            new RibbonGroupSpec("リンク",
                Item("リンク", "\uE71B", "insert-link", "リンクを挿入"),
                Item("ブックマーク", "\uE8A4", "insert-bookmark", "ブックマークを挿入")),
            new RibbonGroupSpec("テキスト",
                Item("テキストボックス", "\uE8A5", "insert-textbox", "テキストボックスを挿入"),
                Item("クイックパーツ", "\uE8A9", "insert-quickparts", "クイックパーツを挿入"),
                Item("日付と時刻", "\uE787", "insert-date", "日付と時刻を挿入"),
                Item("オブジェクト", "\uE8B7", "insert-object", "オブジェクトを挿入"),
                Item("署名欄", "\uE8AC", "insert-signature", "署名欄を挿入")),
            new RibbonGroupSpec("記号",
                Item("数式", "\uE94D", "insert-equation", "数式を挿入"),
                Item("記号", "\uE8D2", "insert-symbol", "記号を挿入", Menu: new[]
                {
                    ("©", "insert-symbol:©"), ("®", "insert-symbol:®"), ("™", "insert-symbol:™"),
                    ("€", "insert-symbol:€"), ("¥", "insert-symbol:¥"), ("§", "insert-symbol:§"),
                    ("¶", "insert-symbol:¶"), ("•", "insert-symbol:•"), ("→", "insert-symbol:→"),
                    ("←", "insert-symbol:←"), ("↑", "insert-symbol:↑"), ("↓", "insert-symbol:↓"),
                    ("×", "insert-symbol:×"), ("÷", "insert-symbol:÷"), ("±", "insert-symbol:±"),
                    ("≠", "insert-symbol:≠"), ("≤", "insert-symbol:≤"), ("≥", "insert-symbol:≥"),
                })));

        BuildTab("design", DesignRibbon,
            new RibbonGroupSpec("ドキュメントの書式設定",
                Item("テーマ", "\uE790", "design-theme", "テーマ", Menu: new[]
                {
                    ("Office", "design-theme:office"), ("Blue", "design-theme:blue"),
                    ("Green", "design-theme:green"), ("Warm", "design-theme:warm"),
                }),
                Item("色", "\uE790", "design-colors", "配色", Menu: new[]
                {
                    ("青", "design-colors:blue"), ("緑", "design-colors:green"),
                    ("オレンジ", "design-colors:orange"), ("紫", "design-colors:purple"),
                }),
                Item("フォント", "\uE8D2", "design-fonts", "フォントの組み合わせ"),
                Item("段落の間隔", "\uE8AC", "design-spacing", "段落の間隔"),
                Item("エフェクト", "\uE7E6", "design-effects", "テーマのエフェクト")),
            new RibbonGroupSpec("ページの背景",
                Item("透かし", "\uE7E6", "design-watermark", "透かしを挿入"),
                Item("ページの色", "\uE790", "design-page-color", "ページの色", Menu: new[]
                {
                    ("白", "page-color:#FFFFFF"), ("クリーム", "page-color:#FFF8E7"),
                    ("薄い青", "page-color:#EAF3FB"), ("薄い緑", "page-color:#EDF6EC"),
                    ("薄いグレー", "page-color:#F4F4F4"),
                }),
                Item("ページ罫線", "\uE7C3", "design-page-border", "ページ罫線", Menu: new[]
                {
                    ("なし", "page-border:none"), ("青", "page-border:#0F6CBD"),
                    ("グレー", "page-border:#A19F9D"), ("金", "page-border:#C9A227"),
                })));

        BuildTab("layout", LayoutRibbon,
            new RibbonGroupSpec("ページ設定",
                Item("余白", "\uE7C3", "layout-margins", "ページの余白", Menu: new[]
                {
                    ("標準", "margin:72"), ("狭い", "margin:36"),
                    ("やや狭い", "margin:54"), ("広い", "margin:108"),
                }),
                Item("印刷の向き", "\uE7C3", "layout-orientation", "印刷の向き", Menu: new[]
                {
                    ("縦", "orientation:portrait"), ("横", "orientation:landscape"),
                }),
                Item("サイズ", "\uE7C3", "layout-size", "用紙サイズ", Menu: new[]
                {
                    ("A4", "page-size:a4"), ("B5", "page-size:b5"),
                    ("Letter", "page-size:letter"), ("Legal", "page-size:legal"),
                }),
                Item("段組み", "\uE8A9", "layout-columns", "段組み", Menu: new[]
                {
                    ("1 段", "layout-columns:1"), ("2 段", "layout-columns:2"),
                    ("3 段", "layout-columns:3"),
                }),
                Item("改行とページ区切り", "\uE8A5", "layout-breaks", "改行とページ区切り"),
                Item("行番号", "\uE8AC", "layout-line-numbers", "行番号"),
                Item("ハイフネーション", "\uE8AB", "layout-hyphenation", "ハイフネーション")),
            new RibbonGroupSpec("段落",
                Item("インデント", "\uE8E4", "indent-increase", "インデントを増やす"),
                Item("間隔", "\uE8AC", "layout-spacing", "段落の間隔")));

        BuildTab("references", ReferencesRibbon,
            new RibbonGroupSpec("目次",
                Item("目次", "\uE8FD", "ref-toc", "目次を挿入"),
                Item("目次の更新", "\uE8CB", "ref-update-toc", "目次を更新")),
            new RibbonGroupSpec("脚注",
                Item("脚注", "\uE8AC", "ref-footnote", "脚注を挿入"),
                Item("文末脚注", "\uE8AC", "ref-endnote", "文末脚注を挿入"),
                Item("注の管理", "\uE8A9", "ref-notes", "注の管理")),
            new RibbonGroupSpec("引用文献",
                Item("引用文献の挿入", "\uE8AC", "ref-citation", "引用文献を挿入"),
                Item("文献目録", "\uE8FD", "ref-bibliography", "文献目録を挿入")),
            new RibbonGroupSpec("図表番号",
                Item("図表番号の挿入", "\uE8AC", "ref-caption", "図表番号を挿入"),
                Item("図表目次", "\uE8FD", "ref-figures", "図表目次を挿入")),
            new RibbonGroupSpec("索引",
                Item("索引の挿入", "\uE8FD", "ref-index", "索引を挿入"),
                Item("索引の更新", "\uE8CB", "ref-update-index", "索引を更新")));

        BuildTab("mailings", MailingsRibbon,
            new RibbonGroupSpec("作成",
                Item("封筒", "\uE715", "mail-envelopes", "封筒を作成"),
                Item("宛名ラベル", "\uE8EC", "mail-labels", "宛名ラベルを作成")),
            new RibbonGroupSpec("差し込み印刷の開始",
                Item("差し込み印刷を開始", "\uE8A5", "mail-start", "差し込み印刷を開始"),
                Item("差し込みフィールド", "\uE8AC", "mail-field", "差し込みフィールドを挿入", Menu: new[]
                {
                    ("«氏名»", "insert-field:氏名"), ("«会社名»", "insert-field:会社名"),
                    ("«住所»", "insert-field:住所"), ("«部署名»", "insert-field:部署名"),
                    ("«電話番号»", "insert-field:電話番号"),
                }),
                Item("結果のプレビュー", "\uE8A3", "mail-preview", "差し込み結果をプレビュー")),
            new RibbonGroupSpec("完了",
                Item("宛先の選択", "\uE716", "mail-recipients", "宛先を選択"),
                Item("完了と差し込み", "\uE8FB", "mail-finish", "差し込み印刷を完了")));

        BuildTab("review", ReviewRibbon,
            new RibbonGroupSpec("文章校正",
                Item("スペルチェック", "\uE8AC", "review-spelling", "スペルチェック"),
                Item("文章校正", "\uE8A9", "review-proofing", "文章校正"),
                Item("類義語辞典", "\uE8FD", "review-thesaurus", "類義語辞典"),
                Item("翻訳", "\uE8B7", "review-translate", "翻訳"),
                Item("言語", "\uE8D2", "review-language", "言語"),
                Item("読み上げ", "\uE8D6", "review-read-aloud", "読み上げ")),
            new RibbonGroupSpec("コメント",
                Item("コメント", "\uE90A", "insert-comment", "コメントを挿入"),
                Item("コメントの削除", "\uE74D", "review-delete-comment", "コメントを削除")),
            new RibbonGroupSpec("変更履歴",
                Item("変更履歴の記録", "\uE8C6", "review-track", "変更履歴の記録"),
                Item("変更の承諾", "\uE8FB", "review-accept", "変更を承諾"),
                Item("変更の拒否", "\uE711", "review-reject", "変更を拒否")),
            new RibbonGroupSpec("保護",
                Item("文書の保護", "\uE72E", "review-protect", "文書を保護"),
                Item("編集の制限", "\uE72E", "review-restrict", "編集を制限"),
                Item("文字カウント", "\uE8AC", "review-word-count", "文字カウント")));

        BuildTab("view", ViewRibbon,
            new RibbonGroupSpec("表示の切り替え",
                Item("閲覧モード", "\uED5E", "view-reading", "閲覧モード"),
                Item("印刷レイアウト", "\uE7C3", "view-print", "印刷レイアウト"),
                Item("Web レイアウト", "\uE774", "view-web", "Web レイアウト")),
            new RibbonGroupSpec("表示",
                Item("ルーラー", "\uED5E", "view-toggle-ruler", "ルーラーを表示"),
                Item("グリッド線", "\uE80A", "view-toggle-grid", "グリッド線を表示"),
                Item("ナビゲーション", "\uE8A9", "view-navigation", "ナビゲーションウィンドウ"),
                Item("段落記号", "\uE8AC", "view-paragraph-marks", "段落記号を表示")),
            new RibbonGroupSpec("テーマ",
                Item("ライト", "\uE706", "theme:light", "ライト テーマ"),
                Item("ダーク", "\uE708", "theme:dark", "ダーク テーマ"),
                Item("システム", "\uE770", "theme:system", "システムの設定に合わせる")),
            new RibbonGroupSpec("ズーム",
                Item("100%", "\uE8A3", "view-zoom-100", "100% に戻す"),
                Item("ページ幅", "\uE740", "view-zoom-page", "ページ幅に合わせる"),
                Item("拡大", "\uE8A3", "zoom-in", "拡大"),
                Item("縮小", "\uE71F", "zoom-out", "縮小")),
            new RibbonGroupSpec("ウィンドウ",
                Item("新しいウィンドウ", "\uE8A7", "view-new-window", "新しいウィンドウを開く"),
                Item("ウィンドウの整列", "\uE8A9", "view-arrange", "ウィンドウを整列")));

        BuildTab("help", HelpRibbon,
            new RibbonGroupSpec("ヘルプ",
                Item("ヘルプ", "\uE897", "help-open", "WordPad Lite のヘルプ"),
                Item("ショートカット", "\uE765", "help-shortcuts", "キーボードショートカット"),
                Item("フィードバック", "\uE90A", "help-feedback", "フィードバックを送る")),
            new RibbonGroupSpec("バージョン",
                Item("WordPad Lite について", "\uE946", "help-about", "バージョン情報"),
                Item("更新プログラムの確認", "\uE895", "help-updates", "更新プログラムを確認")));
    }

    private static RibbonItemSpec Item(
        string label,
        string glyph,
        string command,
        string? tip = null,
        IReadOnlyList<(string Label, string Command)>? Menu = null)
        => new(label, glyph, command, tip, Menu);

    private void BuildRibbonTab(StackPanel host, params RibbonGroupSpec[] groups)
    {
        host.Children.Clear();
        var groupStyle = (Style)Application.Current.Resources["RibbonGroupStyle"];
        var largeButtonStyle = (Style)Application.Current.Resources["RibbonLargeButtonStyle"];
        var labelStyle = (Style)Application.Current.Resources["RibbonGroupLabelStyle"];

        foreach (var groupSpec in groups)
        {
            var groupBorder = new Border
            {
                Style = groupStyle,
            };

            var layout = new Grid();
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            layout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(15) });

            var content = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 1,
                VerticalAlignment = VerticalAlignment.Top,
            };

            foreach (var itemSpec in groupSpec.Items)
            {
                var button = new Button
                {
                    Style = largeButtonStyle,
                    Tag = itemSpec.Command,
                };
                ToolTipService.SetToolTip(button, itemSpec.Tip ?? itemSpec.Label);
                AutomationProperties.SetName(button, itemSpec.Label);

                var face = new StackPanel { Spacing = 1 };
                face.Children.Add(new FontIcon
                {
                    Glyph = itemSpec.Glyph,
                    FontSize = 25,
                    Foreground = new SolidColorBrush(RibbonCommandColor(itemSpec.Command)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                });
                face.Children.Add(new TextBlock
                {
                    Text = itemSpec.Label,
                    TextAlignment = TextAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 52,
                    FontFamily = new FontFamily("Yu Gothic UI"),
                    FontSize = 10,
                    HorizontalAlignment = HorizontalAlignment.Center,
                });
                button.Content = CreateRibbonFace(face);

                if (itemSpec.Menu is { Count: > 0 })
                {
                    var flyout = new MenuFlyout();
                    foreach (var (menuLabel, menuCommand) in itemSpec.Menu)
                    {
                        var menuItem = new MenuFlyoutItem
                        {
                            Text = menuLabel,
                            Tag = menuCommand,
                        };
                        menuItem.Click += RibbonCommand_Click;
                        flyout.Items.Add(menuItem);
                    }

                    TrackFlyout(flyout);
                    button.Flyout = flyout;
                }
                else
                {
                    button.Click += RibbonCommand_Click;
                }

                RegisterRibbonCommandButton(itemSpec.Command, button);
                content.Children.Add(button);
            }

            layout.Children.Add(content);
            var label = new TextBlock
            {
                Text = groupSpec.Label,
                Style = labelStyle,
            };

            var labelRow = new Grid();
            labelRow.Children.Add(label);
            if (groupSpec.Label is "クリップボード" or "フォント" or "段落" or "スタイル")
            {
                var launcher = new Button
                {
                    Style = (Style)Application.Current.Resources["RibbonGroupLauncherStyle"],
                    Tag = groupSpec.Label switch
                    {
                        "クリップボード" => "clipboard-dialog",
                        "フォント" => "font-dialog",
                        "段落" => "paragraph-dialog",
                        _ => "styles-dialog",
                    },
                };
                ToolTipService.SetToolTip(launcher, $"{groupSpec.Label}の詳細設定");
                AutomationProperties.SetName(launcher, $"{groupSpec.Label}の詳細設定");
                launcher.Click += RibbonCommand_Click;
                labelRow.Children.Add(launcher);
            }

            Grid.SetRow(labelRow, 1);
            layout.Children.Add(labelRow);
            groupBorder.Child = layout;
            host.Children.Add(groupBorder);
        }
    }

    private void BuildTab(string tag, StackPanel classicHost, params RibbonGroupSpec[] groups)
    {
        _tabGroups[tag] = groups;
        BuildRibbonTab(classicHost, groups);
        if (_compactPanels.TryGetValue(tag, out var compactHost))
        {
            BuildCompactRibbon(compactHost, groups);
        }
    }

    private void BuildCompactRibbon(StackPanel host, IReadOnlyList<RibbonGroupSpec> groups)
    {
        host.Children.Clear();
        var compactStyle = (Style)Application.Current.Resources["RibbonCompactButtonStyle"];
        var divider = new SolidColorBrush(RootGrid.ActualTheme == ElementTheme.Dark
            ? Color.FromArgb(255, 61, 61, 61)
            : Color.FromArgb(255, 225, 223, 221));
        var firstGroup = true;

        foreach (var group in groups)
        {
            if (!firstGroup)
            {
                host.Children.Add(new Border
                {
                    Width = 1,
                    Height = 22,
                    Background = divider,
                    Margin = new Thickness(4, 0, 4, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                });
            }

            firstGroup = false;

            foreach (var itemSpec in group.Items)
            {
                var button = new Button
                {
                    Style = compactStyle,
                    Tag = itemSpec.Command,
                };
                ToolTipService.SetToolTip(button, itemSpec.Tip ?? itemSpec.Label);
                AutomationProperties.SetName(button, itemSpec.Label);

                var face = new StackPanel { Spacing = 1 };
                face.Children.Add(new FontIcon
                {
                    Glyph = itemSpec.Glyph,
                    FontSize = 15,
                    Foreground = new SolidColorBrush(RibbonCommandColor(itemSpec.Command)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                });
                face.Children.Add(new TextBlock
                {
                    Text = itemSpec.Label,
                    MaxWidth = 58,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    TextAlignment = TextAlignment.Center,
                    FontFamily = new FontFamily("Yu Gothic UI"),
                    FontSize = 9.5,
                    HorizontalAlignment = HorizontalAlignment.Center,
                });
                button.Content = CreateRibbonFace(face, compact: true);

                if (itemSpec.Menu is { Count: > 0 })
                {
                    var flyout = new MenuFlyout();
                    foreach (var (menuLabel, menuCommand) in itemSpec.Menu)
                    {
                        var menuItem = new MenuFlyoutItem
                        {
                            Text = menuLabel,
                            Tag = menuCommand,
                        };
                        menuItem.Click += RibbonCommand_Click;
                        flyout.Items.Add(menuItem);
                    }

                    TrackFlyout(flyout);
                    button.Flyout = flyout;
                }
                else
                {
                    button.Click += RibbonCommand_Click;
                }

                RegisterRibbonCommandButton(itemSpec.Command, button);
                host.Children.Add(button);
            }
        }
    }

    private static Color RibbonCommandColor(string command)
    {
        if (command.StartsWith("theme:", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 191, 144, 0);
        }

        if (command.StartsWith("page-", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 46, 117, 182);
        }

        if (command.StartsWith("style:", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 47, 85, 151);
        }

        if (command.StartsWith("insert-", StringComparison.Ordinal) ||
            command.StartsWith("insert", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 46, 125, 50);
        }

        if (command.StartsWith("design-", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 123, 63, 158);
        }

        if (command.StartsWith("layout-", StringComparison.Ordinal) ||
            command.StartsWith("margin:", StringComparison.Ordinal) ||
            command.StartsWith("orientation:", StringComparison.Ordinal) ||
            command.StartsWith("page-size:", StringComparison.Ordinal) ||
            command.StartsWith("indent-", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 180, 83, 9);
        }

        if (command.StartsWith("ref-", StringComparison.Ordinal) ||
            command.StartsWith("mail-", StringComparison.Ordinal) ||
            command.StartsWith("help-", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 15, 108, 189);
        }

        if (command.StartsWith("review-", StringComparison.Ordinal) ||
            command.StartsWith("insert-comment", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 196, 43, 28);
        }

        if (command.StartsWith("view-", StringComparison.Ordinal) ||
            command.StartsWith("zoom-", StringComparison.Ordinal))
        {
            return Color.FromArgb(255, 71, 85, 105);
        }

        return command switch
        {
            "paste" or "cut" or "copy" or "format-painter" => Color.FromArgb(255, 139, 94, 52),
            "find" or "replace" or "select" or "select-all" => Color.FromArgb(255, 15, 108, 189),
            "font-color" or "highlight" or "grow-font" or "shrink-font" or
                "clear-format" or "change-case" or "text-effects" => Color.FromArgb(255, 31, 78, 121),
            "bullets" or "numbering" or "multilevel-list" or "align-left" or "align-center" or
                "align-right" or "align-justify" or "line-spacing" or "paragraph-spacing" or
                "decrease-indent" or "increase-indent" or "borders" or "shading" or "sort" => Color.FromArgb(255, 47, 85, 151),
            "undo" or "redo" or "save" or "new" or "open" or "print" => Color.FromArgb(255, 15, 108, 189),
            _ => Color.FromArgb(255, 68, 84, 106),
        };
    }

    private static void ApplyRibbonIconColors(DependencyObject root)
    {
        var count = VisualTreeHelper.GetChildrenCount(root);
        for (var index = 0; index < count; index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is Button button && button.Tag is string command)
            {
                var color = new SolidColorBrush(RibbonCommandColor(command));
                ApplyIconColor(button, color);
            }

            ApplyRibbonIconColors(child);
        }
    }

    private static void ApplyIconColor(DependencyObject root, Brush brush)
    {
        if (root is FontIcon directIcon)
        {
            directIcon.Foreground = brush;
            if (directIcon.FontSize >= 18)
            {
                directIcon.FontSize = 25;
            }

            return;
        }

        // Walk logical content as well: when a ribbon button is first created its
        // visual tree is not realized yet, so a visual-only walk would miss the
        // glyph and leave it in its category color.
        if (root is ContentControl contentControl && contentControl.Content is DependencyObject content)
        {
            ApplyIconColor(content, brush);
        }
        else if (root is Border border && border.Child is DependencyObject borderChild)
        {
            ApplyIconColor(borderChild, brush);
        }

        if (root is Panel panel)
        {
            foreach (var child in panel.Children)
            {
                ApplyIconColor(child, brush);
            }
        }

        var count = VisualTreeHelper.GetChildrenCount(root);
        for (var index = 0; index < count; index++)
        {
            ApplyIconColor(VisualTreeHelper.GetChild(root, index), brush);
        }
    }

    private static Grid CreateRibbonFace(StackPanel face, bool compact = false)
    {
        var host = new Grid();
        host.Children.Add(face);
        host.Children.Add(new Border
        {
            Height = 2,
            Margin = new Thickness(compact ? 5 : 8, 0, compact ? 5 : 8, compact ? 1 : 2),
            CornerRadius = new CornerRadius(1),
            Background = (Brush)Application.Current.Resources["ActiveBorderBrush"],
            VerticalAlignment = VerticalAlignment.Bottom,
            Visibility = Visibility.Collapsed,
            IsHitTestVisible = false,
        });
        return host;
    }

    private static Border? GetActiveMark(Button button) =>
        button.Content is Grid grid && grid.Children.Count > 1 && grid.Children[1] is Border mark
            ? mark
            : null;

    private void RegisterRibbonCommandButton(string command, Button button)
    {
        if (!_ribbonCommandButtons.TryGetValue(command, out var buttons))
        {
            buttons = new List<Button>();
            _ribbonCommandButtons[command] = buttons;
        }

        if (!buttons.Contains(button))
        {
            buttons.Add(button);
        }

        ApplyRibbonToggleVisual(command, button);
    }

    private void RefreshRibbonToggleStates()
    {
        foreach (var (command, buttons) in _ribbonCommandButtons)
        {
            foreach (var button in buttons)
            {
                ApplyRibbonToggleVisual(command, button);
            }
        }
    }

    private bool TryGetRibbonToggleState(string command, out bool active)
    {
        switch (command)
        {
            case "view-toggle-ruler": active = _rulerVisible; return true;
            case "view-toggle-grid": active = _gridLinesVisible; return true;
            case "review-spelling": active = _spellCheckEnabled; return true;
            case "review-track": active = _documentMode == DocumentMode.Review; return true;
            case "review-protect":
            case "protect": active = _documentMode == DocumentMode.ReadOnly; return true;
            case "view-reading": active = _readingMode; return true;
            case "theme:light": active = _themePreference == "light"; return true;
            case "theme:dark": active = _themePreference == "dark"; return true;
            case "theme:system": active = _themePreference == "system"; return true;
            default: active = false; return false;
        }
    }

    private void ApplyRibbonToggleVisual(string command, Button button)
    {
        if (!TryGetRibbonToggleState(command, out var active))
        {
            return;
        }

        var activeBrush = (Brush)Application.Current.Resources["ActiveBorderBrush"];
        if (active)
        {
            button.Background = (Brush)Application.Current.Resources["ActiveTintBrush"];
            button.Foreground = activeBrush;
        }
        else
        {
            button.ClearValue(Control.BackgroundProperty);
            button.ClearValue(Control.ForegroundProperty);
        }

        if (GetActiveMark(button) is { } mark)
        {
            mark.Background = activeBrush;
            mark.Visibility = active ? Visibility.Visible : Visibility.Collapsed;
        }

        ApplyIconColor(
            button,
            active
                ? activeBrush
                : new SolidColorBrush(RibbonCommandColor(command)));
    }

    private void RebuildCompactRibbons()
    {
        foreach (var (tag, groups) in _tabGroups)
        {
            if (_compactPanels.TryGetValue(tag, out var panel))
            {
                BuildCompactRibbon(panel, groups);
            }
        }
    }

    private static void CopyComboItems(ComboBox source, ComboBox target)
    {
        target.Items.Clear();
        foreach (var item in source.Items)
        {
            switch (item)
            {
                case ComboBoxItem comboItem:
                    target.Items.Add(new ComboBoxItem { Content = comboItem.Content, Tag = comboItem.Tag });
                    break;
                case string text:
                    target.Items.Add(text);
                    break;
                default:
                    target.Items.Add(item?.ToString() ?? string.Empty);
                    break;
            }
        }
    }

    private Flyout BuildFontColorFlyout()
    {
        var flyout = new Flyout { Placement = FlyoutPlacementMode.Bottom };
        var content = new StackPanel { Width = 196, Spacing = 5 };
        content.Children.Add(new TextBlock
        {
            Text = "フォントの色",
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 12,
            FontWeight = FontWeights.SemiBold,
        });

        var swatchStyle = (Style)Application.Current.Resources["ColorSwatchButtonStyle"];
        var colors = new (string Hex, string Name)[]
        {
            ("#000000", "黒"), ("#1F3864", "濃い青"), ("#C00000", "赤"),
            ("#375623", "濃い緑"), ("#7030A0", "紫"), ("#C55A11", "オレンジ"),
            ("#2F75B5", "青"), ("#7F7F7F", "グレー"), ("#BFBFBF", "薄いグレー"),
            ("#FFC000", "黄"), ("#00B0F0", "水色"), ("#FFFFFF", "白"),
        };

        for (var index = 0; index < colors.Length; index += 6)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal };
            for (var offset = index; offset < index + 6 && offset < colors.Length; offset++)
            {
                var (hex, name) = colors[offset];
                var swatch = new Button
                {
                    Style = swatchStyle,
                    Background = new SolidColorBrush(ParseColor(hex)),
                    Tag = $"font-color:{hex}",
                };
                ToolTipService.SetToolTip(swatch, name);
                swatch.Click += RibbonCommand_Click;
                row.Children.Add(swatch);
            }

            content.Children.Add(row);
        }

        flyout.Content = content;
        return flyout;
    }

    private async Task ExecuteCommandAsync(string tag)
    {
        try
        {
            if (tag.StartsWith("theme:", StringComparison.Ordinal))
            {
                SetThemePreference(tag["theme:".Length..]);
                return;
            }

            if (tag.StartsWith("font-color:", StringComparison.Ordinal))
            {
                ApplyTextColor(ParseColor(tag["font-color:".Length..]), highlight: false);
                return;
            }

            if (tag.StartsWith("highlight:", StringComparison.Ordinal))
            {
                var value = tag["highlight:".Length..];
                ApplyTextColor(value == "none" ? Color.FromArgb(0, 0, 0, 0) : ParseColor(value), highlight: true);
                return;
            }

            if (tag.StartsWith("shading:", StringComparison.Ordinal))
            {
                var value = tag["shading:".Length..];
                ApplyParagraphBackground(value == "none" ? Color.FromArgb(0, 0, 0, 0) : ParseColor(value));
                return;
            }

            if (tag.StartsWith("page-color:", StringComparison.Ordinal))
            {
                PageBorder.Background = new SolidColorBrush(ParseColor(tag["page-color:".Length..]));
                return;
            }

            if (tag.StartsWith("page-border:", StringComparison.Ordinal))
            {
                var value = tag["page-border:".Length..];
                if (value == "none")
                {
                    PageBorder.BorderThickness = new Thickness(1);
                    PageBorder.BorderBrush = new SolidColorBrush(Color.FromArgb(255, 210, 208, 206));
                }
                else
                {
                    PageBorder.BorderThickness = new Thickness(2);
                    PageBorder.BorderBrush = new SolidColorBrush(ParseColor(value));
                }
                return;
            }

            if (tag.StartsWith("style:", StringComparison.Ordinal))
            {
                ApplyStyle(tag["style:".Length..]);
                return;
            }

            if (tag.StartsWith("margin:", StringComparison.Ordinal))
            {
                _pageMargin = double.Parse(tag["margin:".Length..], CultureInfo.InvariantCulture);
                ApplyPageLayout();
                return;
            }

            if (tag.StartsWith("orientation:", StringComparison.Ordinal))
            {
                _landscape = tag.EndsWith("landscape", StringComparison.Ordinal);
                ApplyPageSize(_pageSizeName, _landscape);
                return;
            }

            if (tag.StartsWith("page-size:", StringComparison.Ordinal))
            {
                ApplyPageSize(tag["page-size:".Length..], _landscape);
                return;
            }

            if (tag.StartsWith("layout-columns:", StringComparison.Ordinal))
            {
                ShowNotAvailable("段組み");
                return;
            }

            if (tag.StartsWith("insert-symbol:", StringComparison.Ordinal))
            {
                InsertText(tag["insert-symbol:".Length..]);
                return;
            }

            if (tag.StartsWith("insert-field:", StringComparison.Ordinal))
            {
                InsertText($"«{tag["insert-field:".Length..]}»");
                return;
            }

            if (tag.StartsWith("line-spacing:", StringComparison.Ordinal))
            {
                var value = float.Parse(tag["line-spacing:".Length..], CultureInfo.InvariantCulture);
                ApplyLineSpacing(value);
                return;
            }

            if (tag.StartsWith("space-before:", StringComparison.Ordinal))
            {
                ApplyParagraphSpace(before: float.Parse(tag["space-before:".Length..], CultureInfo.InvariantCulture), after: null);
                return;
            }

            if (tag.StartsWith("space-after:", StringComparison.Ordinal))
            {
                ApplyParagraphSpace(before: null, after: float.Parse(tag["space-after:".Length..], CultureInfo.InvariantCulture));
                return;
            }

            if (tag.StartsWith("space-both:", StringComparison.Ordinal))
            {
                var value = float.Parse(tag["space-both:".Length..], CultureInfo.InvariantCulture);
                ApplyParagraphSpace(before: value, after: value);
                return;
            }

            switch (tag)
            {
                case "new":
                    await NewDocumentAsync();
                    break;
                case "open":
                    await OpenAsync();
                    break;
                case "save":
                    await SaveAsync(saveAs: false);
                    break;
                case "save-as":
                    await SaveAsync(saveAs: true);
                    break;
                case "export":
                    await ExportAsync();
                    break;
                case "print":
                    await PrintAsync();
                    break;
                case "options":
                    await ShowOptionsAsync();
                    break;
                case "exit":
                    Close();
                    break;
                case "protect":
                case "review-protect":
                    ToggleProtection();
                    break;
                case "qat-more":
                    ShowQuickAccessMenu();
                    break;
                case "clipboard-dialog":
                    Notify("クリップボードの詳細設定は準備中です。");
                    break;
                case "font-dialog":
                    Notify("フォントの詳細設定は準備中です。");
                    break;
                case "paragraph-dialog":
                    Notify("段落の詳細設定は準備中です。");
                    break;
                case "styles-dialog":
                    Notify("スタイルの詳細設定は準備中です。");
                    break;

                case "undo":
                    if (Editor.Document.CanUndo())
                    {
                        Editor.Document.Undo();
                        UpdateTitle();
                        UpdateFormattingUi();
                    }
                    break;
                case "redo":
                    if (Editor.Document.CanRedo())
                    {
                        Editor.Document.Redo();
                        UpdateTitle();
                        UpdateFormattingUi();
                    }
                    break;

                case "paste":
                    if (Editor.Document.CanPaste())
                    {
                        Editor.Document.Selection.Paste(0);
                        MarkModified();
                    }
                    else
                    {
                        Notify("貼り付ける内容がクリップボードにありません。");
                    }
                    break;
                case "cut":
                    if (Editor.Document.Selection.Length > 0)
                    {
                        Editor.Document.Selection.Cut();
                        Editor.Focus(FocusState.Programmatic);
                        MarkModified();
                    }
                    break;
                case "copy":
                    if (Editor.Document.Selection.Length > 0)
                    {
                        Editor.Document.Selection.Copy();
                        Notify("選択した文字列をコピーしました。");
                    }
                    break;
                case "format-painter":
                    CaptureFormatPainter();
                    break;
                case "select-all":
                    Editor.Document.Selection.SetRange(0, GetPlainText().Length);
                    Editor.Focus(FocusState.Programmatic);
                    break;

                case "bold":
                    ToggleBold();
                    break;
                case "italic":
                    ToggleItalic();
                    break;
                case "underline":
                    ToggleUnderline();
                    break;
                case "underline-single":
                    ApplyUnderline(UnderlineType.Single);
                    break;
                case "underline-double":
                    ApplyUnderline(UnderlineType.Double);
                    break;
                case "underline-dotted":
                    ApplyUnderline(UnderlineType.Dotted);
                    break;
                case "underline-wave":
                    ApplyUnderline(UnderlineType.Wave);
                    break;
                case "underline-none":
                    ApplyUnderline(UnderlineType.None);
                    break;
                case "strike":
                    ToggleStrikethrough();
                    break;
                case "subscript":
                    ToggleScript(superscript: false);
                    break;
                case "superscript":
                    ToggleScript(superscript: true);
                    break;
                case "grow-font":
                    ChangeFontSize(1);
                    break;
                case "shrink-font":
                    ChangeFontSize(-1);
                    break;
                case "clear-format":
                    ClearFormatting();
                    break;
                case "case-upper":
                    ChangeTextCase("upper");
                    break;
                case "case-lower":
                    ChangeTextCase("lower");
                    break;
                case "case-title":
                    ChangeTextCase("title");
                    break;
                case "case-toggle":
                    ChangeTextCase("toggle");
                    break;

                case "align-left":
                    ApplyAlignment(ParagraphAlignment.Left);
                    break;
                case "align-center":
                    ApplyAlignment(ParagraphAlignment.Center);
                    break;
                case "align-right":
                    ApplyAlignment(ParagraphAlignment.Right);
                    break;
                case "align-justify":
                    ApplyAlignment(ParagraphAlignment.Justify);
                    break;

                case "bullets":
                    ApplyList(MarkerType.Bullet, MarkerStyle.Plain);
                    break;
                case "bullets-white":
                    ApplyList(MarkerType.WhiteCircleWingding, MarkerStyle.Plain);
                    break;
                case "bullets-square":
                    ApplyList(MarkerType.BlackCircleWingding, MarkerStyle.Plain);
                    break;
                case "numbering":
                    ApplyList(MarkerType.Arabic, MarkerStyle.Period);
                    break;
                case "numbering-alpha":
                    ApplyList(MarkerType.UppercaseEnglishLetter, MarkerStyle.Period);
                    break;
                case "numbering-alpha-lower":
                    ApplyList(MarkerType.LowercaseEnglishLetter, MarkerStyle.Period);
                    break;
                case "numbering-roman":
                    ApplyList(MarkerType.LowercaseRoman, MarkerStyle.Period);
                    break;
                case "multilevel":
                    ApplyList(MarkerType.Arabic, MarkerStyle.Period);
                    break;
                case "list-none":
                    RemoveList();
                    break;
                case "indent-increase":
                    ChangeIndent(18f);
                    break;
                case "indent-decrease":
                    ChangeIndent(-18f);
                    break;
                case "borders":
                case "sort":
                case "paragraph-outline":
                case "distribute":
                case "text-effects":
                    ShowNotAvailable(TagDisplayName(tag));
                    break;

                case "find":
                    await ShowFindReplaceAsync(replaceMode: false);
                    break;
                case "replace":
                    await ShowFindReplaceAsync(replaceMode: true);
                    break;

                case "insert-cover":
                    await InsertCoverAsync();
                    break;
                case "insert-blank-page":
                    InsertRtf("{\\rtf1\\ansi\\page\\par}");
                    break;
                case "insert-page-break":
                    InsertRtf("{\\rtf1\\ansi\\page}");
                    break;
                case "insert-page-number":
                    InsertText("1");
                    Notify("ページ番号を挿入しました。");
                    break;
                case "insert-header":
                    InsertHeaderFooterInline(header: true);
                    break;
                case "insert-footer":
                    InsertHeaderFooterInline(header: false);
                    break;
                case "insert-table":
                    InsertTable();
                    break;
                case "insert-table-convert":
                    ShowNotAvailable("表の変換");
                    break;
                case "insert-picture":
                    await InsertPictureAsync();
                    break;
                case "insert-shape":
                case "insert-icon":
                case "insert-smartart":
                case "insert-chart":
                case "insert-screenshot":
                case "insert-textbox":
                case "insert-quickparts":
                case "insert-object":
                case "insert-signature":
                    ShowNotAvailable(TagDisplayName(tag));
                    break;
                case "insert-link":
                    await InsertLinkAsync();
                    break;
                case "insert-bookmark":
                    await InsertBookmarkAsync();
                    break;
                case "insert-date":
                    InsertText(DateTime.Now.ToString("yyyy年M月d日 HH:mm", CultureInfo.GetCultureInfo("ja-JP")));
                    break;
                case "insert-equation":
                    InsertText("E = mc²");
                    break;
                case "insert-symbol":
                    InsertText("©");
                    break;
                case "insert-comment":
                    await InsertCommentAsync();
                    break;

                case "design-theme":
                    ShowNotAvailable("テーマ");
                    break;
                case "design-watermark":
                    await InsertWatermarkAsync();
                    break;
                case "design-colors":
                case "design-fonts":
                case "design-spacing":
                case "design-effects":
                    ShowNotAvailable(TagDisplayName(tag));
                    break;

                case "layout-margins":
                case "layout-orientation":
                case "layout-size":
                case "layout-columns":
                case "layout-breaks":
                case "layout-line-numbers":
                case "layout-hyphenation":
                case "layout-spacing":
                    ShowNotAvailable(TagDisplayName(tag));
                    break;

                case "ref-toc":
                    InsertTocPlaceholder();
                    break;
                case "ref-update-toc":
                    Notify("目次を更新しました。");
                    break;
                case "ref-footnote":
                    InsertFootnoteInline(endnote: false);
                    break;
                case "ref-endnote":
                    InsertFootnoteInline(endnote: true);
                    break;
                case "ref-notes":
                    Notify("脚注と文末脚注は文書の末尾にまとめて挿入されます。");
                    break;
                case "ref-citation":
                    await InsertCitationAsync();
                    break;
                case "ref-bibliography":
                    InsertText("\r参考文献\r1. \r");
                    break;
                case "ref-caption":
                    await InsertCaptionAsync();
                    break;
                case "ref-figures":
                    InsertText("\r図表目次\r図 1: \r");
                    break;
                case "ref-index":
                    InsertText("\r索引\r");
                    break;
                case "ref-update-index":
                    Notify("索引を更新しました。");
                    break;

                case "mail-envelopes":
                    await InsertEnvelopeAsync();
                    break;
                case "mail-labels":
                    await InsertLabelAsync();
                    break;
                case "mail-start":
                    await ShowMailMergeStartAsync();
                    break;
                case "mail-preview":
                    Notify("差し込み結果のプレビューを更新しました。");
                    break;
                case "mail-recipients":
                    await ShowRecipientsAsync();
                    break;
                case "mail-finish":
                    Notify("差し込み印刷を完了しました。");
                    break;

                case "review-spelling":
                    ToggleSpellCheck();
                    break;
                case "review-word-count":
                    await ShowWordCountAsync();
                    break;
                case "review-track":
                    SetDocumentMode(_documentMode == DocumentMode.Review ? DocumentMode.Edit : DocumentMode.Review);
                    break;
                case "review-proofing":
                    ShowProofingPane();
                    break;
                case "review-delete-comment":
                case "review-accept":
                case "review-reject":
                case "review-thesaurus":
                case "review-translate":
                case "review-read-aloud":
                case "review-restrict":
                    ShowNotAvailable(TagDisplayName(tag));
                    break;
                case "review-language":
                    await ShowLanguageAsync();
                    break;

                case "view-reading":
                case "view-exit-reading":
                    ToggleReadingMode();
                    break;
                case "view-toggle-ruler":
                    _rulerVisible = !_rulerVisible;
                    UpdateRulerVisibility();
                    Notify(_rulerVisible ? "ルーラー (横・縦) を表示しました。" : "ルーラー (横・縦) を非表示にしました。");
                    break;
                case "view-toggle-grid":
                    _gridLinesVisible = !_gridLinesVisible;
                    GridLinesOverlay.Visibility = _gridLinesVisible ? Visibility.Visible : Visibility.Collapsed;
                    DrawGridLines();
                    RefreshRibbonToggleStates();
                    Notify(_gridLinesVisible ? "グリッド線を表示しました。" : "グリッド線を非表示にしました。");
                    break;
                case "view-navigation":
                    Notify("ナビゲーションウィンドウは現在の文書では不要です。見出しはスタイルで管理されています。");
                    break;
                case "view-paragraph-marks":
                    Notify("段落記号の表示は現在のネイティブ版では未対応です。");
                    break;
                case "view-print":
                case "view-web":
                case "view-arrange":
                case "view-new-window":
                    ShowNotAvailable(TagDisplayName(tag));
                    break;
                case "view-zoom-100":
                case "zoom-reset":
                    SetZoom(100);
                    break;
                case "view-zoom-page":
                    SetZoom(100);
                    break;
                case "zoom-in":
                    SetZoom(Math.Min(250, PageScrollViewer.ZoomFactor * 100 + 10));
                    break;
                case "zoom-out":
                    SetZoom(Math.Max(50, PageScrollViewer.ZoomFactor * 100 - 10));
                    break;

                case "help-open":
                    await ShowHelpAsync();
                    break;
                case "help-shortcuts":
                    await ShowShortcutsAsync();
                    break;
                case "help-feedback":
                    await ShowFeedbackAsync();
                    break;
                case "help-about":
                    ShowBackstage();
                    ShowBackstagePage("account");
                    break;
                case "help-updates":
                    Notify("WordPad Lite は最新のバージョンです。");
                    break;

                default:
                    ShowNotAvailable(tag);
                    break;
            }
        }
        catch (UnauthorizedAccessException)
        {
            Notify("読み取り専用モードではこの操作はできません。編集モードに切り替えてから実行してください。", InfoBarSeverity.Warning);
        }
        catch (Exception ex)
        {
            Notify($"操作を完了できませんでした: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private static string TagDisplayName(string tag)
    {
        return tag switch
        {
            "borders" => "段落の罫線",
            "sort" => "並べ替え",
            "paragraph-outline" => "アウトライン",
            "distribute" => "均等割り付け",
            "text-effects" => "拡張書式",
            "insert-shape" => "図形",
            "insert-icon" => "アイコン",
            "insert-smartart" => "SmartArt",
            "insert-chart" => "グラフ",
            "insert-screenshot" => "スクリーンショット",
            "insert-textbox" => "テキストボックス",
            "insert-quickparts" => "クイックパーツ",
            "insert-object" => "オブジェクト",
            "insert-signature" => "署名欄",
            "design-colors" => "配色",
            "design-fonts" => "フォントの組み合わせ",
            "design-spacing" => "段落の間隔",
            "design-effects" => "テーマのエフェクト",
            "layout-breaks" => "改行とページ区切り",
            "layout-line-numbers" => "行番号",
            "layout-hyphenation" => "ハイフネーション",
            "layout-spacing" => "段落の間隔",
            "review-delete-comment" => "コメントの削除",
            "review-accept" => "変更の承諾",
            "review-reject" => "変更の拒否",
            "review-proofing" => "文章校正",
            "review-thesaurus" => "類義語辞典",
            "review-translate" => "翻訳",
            "review-read-aloud" => "読み上げ",
            "review-restrict" => "編集の制限",
            "view-print" => "印刷レイアウト",
            "view-web" => "Web レイアウト",
            "view-arrange" => "ウィンドウの整列",
            "view-new-window" => "新しいウィンドウ",
            _ => tag,
        };
    }

    private void ShowQuickAccessMenu()
    {
        var flyout = new MenuFlyout();
        flyout.Items.Add(new MenuFlyoutItem { Text = "クイック アクセス ツール バーのカスタマイズ", IsEnabled = false });

        foreach (var spec in QuickAccessCatalog)
        {
            var item = new ToggleMenuFlyoutItem
            {
                Text = spec.Label,
                IsChecked = _quickAccessItems.Contains(spec.Id, StringComparer.OrdinalIgnoreCase),
            };
            item.Click += (_, _) =>
            {
                ToggleQuickAccessItem(spec.Id);
                DispatcherQueue.TryEnqueue(ShowQuickAccessMenu);
            };
            flyout.Items.Add(item);
        }

        flyout.Items.Add(new MenuFlyoutSeparator());

        var resetItem = new MenuFlyoutItem { Text = "クイック アクセス ツール バーを既定に戻す" };
        resetItem.Click += (_, _) =>
        {
            _quickAccessItems = new List<string> { "save", "undo", "redo" };
            AppSettings.Set("QuickAccess", string.Join(",", _quickAccessItems));
            BuildQuickAccess();
            DispatcherQueue.TryEnqueue(ShowQuickAccessMenu);
        };
        flyout.Items.Add(resetItem);

        flyout.Items.Add(new MenuFlyoutSeparator());

        var compactItem = new RadioMenuFlyoutItem
        {
            Text = "簡易リボン",
            GroupName = "RibbonMode",
            IsChecked = _useCompactRibbon,
        };
        compactItem.Click += (_, _) => SetRibbonMode(true);
        flyout.Items.Add(compactItem);

        var classicItem = new RadioMenuFlyoutItem
        {
            Text = "従来のリボン",
            GroupName = "RibbonMode",
            IsChecked = !_useCompactRibbon,
        };
        classicItem.Click += (_, _) => SetRibbonMode(false);
        flyout.Items.Add(classicItem);

        var collapseItem = new MenuFlyoutItem
        {
            Text = _ribbonCollapsed ? "リボンを展開" : "リボンを折りたたむ",
        };
        collapseItem.Click += (_, _) => SetRibbonCollapsed(!_ribbonCollapsed);
        flyout.Items.Add(collapseItem);

        flyout.Items.Add(new MenuFlyoutSeparator());
        AddQuickItem(flyout, "オプション", "options");
        flyout.ShowAt(QuickAccessButton);
    }

    private void ToggleQuickAccessItem(string id)
    {
        var removed = _quickAccessItems.RemoveAll(candidate =>
            string.Equals(candidate, id, StringComparison.OrdinalIgnoreCase));
        if (removed == 0)
        {
            _quickAccessItems.Add(id);
        }

        AppSettings.Set("QuickAccess", string.Join(",", _quickAccessItems));
        BuildQuickAccess();
    }

    private void BuildQuickAccess()
    {
        QuickAccessPanel.Children.Clear();
        _quickAccessButtons.Clear();

        var buttonStyle = (Style)Application.Current.Resources["TitleBarButtonStyle"];
        foreach (var id in _quickAccessItems)
        {
            var spec = QuickAccessCatalog.FirstOrDefault(candidate => string.Equals(candidate.Id, id, StringComparison.OrdinalIgnoreCase));
            if (spec is null)
            {
                continue;
            }

            var button = new Button
            {
                Style = buttonStyle,
                Content = spec.Glyph,
                Tag = spec.Command,
                FontSize = 13.5,
            };
            ToolTipService.SetToolTip(button, spec.Tip);
            AutomationProperties.SetName(button, spec.Label);
            button.Click += RibbonCommand_Click;
            QuickAccessPanel.Children.Add(button);
            _quickAccessButtons[spec.Command] = button;
        }

        RefreshQuickAccessState();
    }

    private void RefreshQuickAccessState()
    {
        foreach (var (command, button) in _quickAccessButtons)
        {
            switch (command)
            {
                case "undo":
                    button.IsEnabled = _editorReady && Editor.Document.CanUndo();
                    break;
                case "redo":
                    button.IsEnabled = _editorReady && Editor.Document.CanRedo();
                    break;
            }
        }
    }

    private void AddQuickItem(MenuFlyout flyout, string text, string command)
    {
        var item = new MenuFlyoutItem { Text = text, Tag = command };
        item.Click += RibbonCommand_Click;
        flyout.Items.Add(item);
    }

    private void ToggleProtection()
    {
        SetDocumentMode(_documentMode == DocumentMode.ReadOnly ? DocumentMode.Edit : DocumentMode.ReadOnly);
    }

    private void BackstageReadOnlySwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_updatingBackstage)
        {
            return;
        }

        SetDocumentMode(BackstageReadOnlySwitch.IsOn ? DocumentMode.ReadOnly : DocumentMode.Edit);
    }

    private void DocumentModeMenu_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not FrameworkElement element || element.Tag is not string tag)
        {
            return;
        }

        SetDocumentMode(tag switch
        {
            "review" => DocumentMode.Review,
            "readonly" => DocumentMode.ReadOnly,
            _ => DocumentMode.Edit,
        });
    }

    private void SetDocumentMode(DocumentMode mode, bool notify = true)
    {
        _documentMode = mode;
        _trackChanges = mode == DocumentMode.Review;
        Editor.IsReadOnly = mode == DocumentMode.ReadOnly;

        ModeEditItem.IsChecked = mode == DocumentMode.Edit;
        ModeReviewItem.IsChecked = mode == DocumentMode.Review;
        ModeReadOnlyItem.IsChecked = mode == DocumentMode.ReadOnly;

        var modeLabel = mode switch
        {
            DocumentMode.Review => "校閲",
            DocumentMode.ReadOnly => "読み取り専用",
            _ => "編集",
        };
        DocumentModeText.Text = modeLabel;
        ModeStatusText.Text = modeLabel;
        RefreshRibbonToggleStates();

        _updatingBackstage = true;
        try
        {
            BackstageReadOnlySwitch.IsOn = mode == DocumentMode.ReadOnly;
            BackstageProtectionText.Text = mode switch
            {
                DocumentMode.Review => "現在のモード: 校閲 (変更履歴を記録)",
                DocumentMode.ReadOnly => "現在のモード: 読み取り専用",
                _ => "現在のモード: 編集",
            };
        }
        finally
        {
            _updatingBackstage = false;
        }

        if (notify)
        {
            Notify(mode switch
            {
                DocumentMode.Review => "校閲モードに切り替えました。変更履歴の記録を有効にしています。",
                DocumentMode.ReadOnly => "読み取り専用にしました。本文は保護されています。",
                _ => "編集モードに切り替えました。",
            });
        }
    }

    private async Task NewDocumentAsync()
    {
        if (!await ConfirmDiscardAsync())
        {
            return;
        }

        _isLoading = true;
        try
        {
            Editor.Document.SetText(TextSetOptions.None, string.Empty);
            Editor.Document.ClearUndoRedoHistory();
            InitializeEditorDefaults();
        }
        finally
        {
            _isLoading = false;
        }

        _currentFile = null;
        _currentFormat = "RTF";
        _isModified = false;
        _untitledNumber++;
        _untitledName = $"文書 {_untitledNumber}";
        Editor.Document.Selection.SetRange(0, 0);
        UpdateTitle();
        UpdateWordCount();
        UpdateFormattingUi();
        Editor.Focus(FocusState.Programmatic);
    }

    private async Task OpenAsync()
    {
        if (!await ConfirmDiscardAsync())
        {
            return;
        }

        var picker = new FileOpenPicker
        {
            ViewMode = PickerViewMode.List,
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
        };
        picker.FileTypeFilter.Add(".rtf");
        picker.FileTypeFilter.Add(".docx");
        picker.FileTypeFilter.Add(".txt");
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSingleFileAsync();
        if (file is not null)
        {
            await LoadFileAsync(file);
        }
    }

    private async Task LoadFileAsync(StorageFile file)
    {
        var extension = IOPath.GetExtension(file.Name).ToLowerInvariant();
        if (extension is not (".rtf" or ".docx" or ".txt"))
        {
            Notify("RTF、DOCX、TXT ファイルを開くことができます。", InfoBarSeverity.Warning);
            return;
        }

        _isLoading = true;
        try
        {
            switch (extension)
            {
                case ".rtf":
                {
                    using var stream = await file.OpenAsync(FileAccessMode.Read);
                    Editor.Document.LoadFromStream(TextSetOptions.FormatRtf, stream);
                    break;
                }
                case ".docx":
                {
                    var rtf = DocxService.ExtractRtf(file.Path);
                    Editor.Document.SetText(TextSetOptions.FormatRtf, rtf);
                    break;
                }
                default:
                {
                    var text = await FileIO.ReadTextAsync(file);
                    text = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r");
                    Editor.Document.SetText(TextSetOptions.None, text);
                    break;
                }
            }

            Editor.Document.ClearUndoRedoHistory();
        }
        catch (Exception ex)
        {
            Notify($"ファイルを開けませんでした: {ex.Message}", InfoBarSeverity.Error);
            return;
        }
        finally
        {
            _isLoading = false;
        }

        _currentFile = file;
        _currentFormat = FormatFromExtension(extension);
        _isModified = false;
        Editor.Document.Selection.SetRange(0, 0);
        UpdateTitle();
        UpdateWordCount();
        UpdateFormattingUi();
        Editor.Focus(FocusState.Programmatic);

        // Loading through RichEdit can raise TextChanged after the method returns,
        // which would incorrectly mark the freshly opened document as modified.
        DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            if (!_editorReady || !ReferenceEquals(_currentFile, file))
            {
                return;
            }

            Editor.Document.ClearUndoRedoHistory();
            _isModified = false;
            UpdateTitle();
            UpdateUndoRedo();
        });

        Notify($"{file.Name} を開きました。");
    }

    private async Task<bool> SaveAsync(bool saveAs)
    {
        if (!_editorReady)
        {
            return false;
        }

        if (saveAs || _currentFile is null)
        {
            return await SaveAsAsync();
        }

        return await WriteToFileAsync(_currentFile);
    }

    private async Task<bool> SaveAsAsync()
    {
        var picker = new FileSavePicker
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            SuggestedFileName = _currentFile is null
                ? _untitledName
                : IOPath.GetFileNameWithoutExtension(_currentFile.Name),
            DefaultFileExtension = _currentFormat switch
            {
                "DOCX" => ".docx",
                "TXT" => ".txt",
                _ => ".rtf",
            },
        };
        picker.FileTypeChoices.Add("リッチ テキスト形式", new List<string> { ".rtf" });
        picker.FileTypeChoices.Add("Word 文書", new List<string> { ".docx" });
        picker.FileTypeChoices.Add("テキスト ファイル", new List<string> { ".txt" });
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSaveFileAsync();
        if (file is null)
        {
            return false;
        }

        return await WriteToFileAsync(file);
    }

    private async Task<bool> WriteToFileAsync(StorageFile file, bool silent = false)
    {
        var extension = IOPath.GetExtension(file.Name).ToLowerInvariant();
        try
        {
            switch (extension)
            {
                case ".docx":
                {
                    var text = GetPlainText();
                    var bytes = DocxService.CreateDocx(text, _pageWidth, _pageHeight, _pageMargin);
                    await FileIO.WriteBytesAsync(file, bytes);
                    break;
                }
                case ".txt":
                {
                    var text = GetPlainText().Replace("\r", "\r\n");
                    await FileIO.WriteTextAsync(file, text, Windows.Storage.Streams.UnicodeEncoding.Utf8);
                    break;
                }
                default:
                {
                    extension = ".rtf";
                    using var stream = await file.OpenAsync(FileAccessMode.ReadWrite);
                    stream.Size = 0;
                    Editor.Document.SaveToStream(TextGetOptions.FormatRtf, stream);
                    await stream.FlushAsync();
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            if (!silent)
            {
                Notify($"保存できませんでした: {ex.Message}", InfoBarSeverity.Error);
            }

            return false;
        }

        _currentFile = file;
        _currentFormat = FormatFromExtension(extension);
        _isModified = false;
        if (!silent)
        {
            _lastAutoSaveAt = null;
        }

        UpdateTitle();
        UpdateBackstageInfo();
        if (!silent)
        {
            Notify($"{file.Name} に保存しました。");
        }

        return true;
    }

    private async Task ExportAsync()
    {
        var picker = new FileSavePicker
        {
            SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            SuggestedFileName = _currentFile is null
                ? _untitledName
                : IOPath.GetFileNameWithoutExtension(_currentFile.Name),
            DefaultFileExtension = ".docx",
        };
        picker.FileTypeChoices.Add("Word 文書", new List<string> { ".docx" });
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSaveFileAsync();
        if (file is null)
        {
            return;
        }

        try
        {
            var bytes = DocxService.CreateDocx(GetPlainText(), _pageWidth, _pageHeight, _pageMargin);
            await FileIO.WriteBytesAsync(file, bytes);
            Notify($"{file.Name} をエクスポートしました。");
        }
        catch (Exception ex)
        {
            Notify($"エクスポートできませんでした: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async Task PrintAsync()
    {
        var settings = await ShowPrintSettingsAsync();
        if (settings is null)
        {
            return;
        }

        var (paper, landscape, marginCm) = settings.Value;
        var paperLabel = paper switch
        {
            "b5" => "B5",
            "letter" => "Letter",
            "legal" => "Legal",
            _ => "A4",
        };

        try
        {
            var folder = await StorageFolder.GetFolderFromPathAsync(IOPath.GetTempPath());
            var file = await folder.CreateFileAsync(
                $"WordPadLite-{DateTime.Now:HHmmss}.rtf",
                CreationCollisionOption.ReplaceExisting);

            var rtf = Encoding.Latin1.GetString(await GetRtfBytesAsync());
            rtf = ApplyRtfPageSettings(rtf, paper, landscape, marginCm);
            await FileIO.WriteBytesAsync(file, Encoding.Latin1.GetBytes(rtf));

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = file.Path,
                    Verb = "print",
                    UseShellExecute = true,
                });
                Notify($"印刷を開始しました ({paperLabel} / {(landscape ? "横" : "縦")} / 余白 {marginCm:0.0} cm)。");
            }
            catch
            {
                await Launcher.LaunchFileAsync(file);
                Notify("既定のアプリで文書を開きました。印刷メニューから印刷してください。");
            }
        }
        catch (Exception ex)
        {
            Notify($"印刷を開始できませんでした: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async Task<(string Paper, bool Landscape, double MarginCm)?> ShowPrintSettingsAsync()
    {
        var paperCombo = new ComboBox
        {
            Header = "用紙サイズ",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        paperCombo.Items.Add("A4 (210 × 297 mm)");
        paperCombo.Items.Add("B5 (182 × 257 mm)");
        paperCombo.Items.Add("Letter (8.5 × 11 in)");
        paperCombo.Items.Add("Legal (8.5 × 14 in)");
        paperCombo.SelectedIndex = _pageSizeName switch
        {
            "b5" => 1,
            "letter" => 2,
            "legal" => 3,
            _ => 0,
        };

        var orientationCombo = new ComboBox
        {
            Header = "印刷の向き",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        orientationCombo.Items.Add("縦");
        orientationCombo.Items.Add("横");
        orientationCombo.SelectedIndex = _landscape ? 1 : 0;

        var marginCombo = new ComboBox
        {
            Header = "余白",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        marginCombo.Items.Add(new ComboBoxItem { Content = "標準 (2.54 cm)", Tag = 2.54 });
        marginCombo.Items.Add(new ComboBoxItem { Content = "狭い (1.27 cm)", Tag = 1.27 });
        marginCombo.Items.Add(new ComboBoxItem { Content = "広い (3.81 cm)", Tag = 3.81 });
        marginCombo.SelectedIndex = 0;

        var info = new TextBlock
        {
            Text = "用紙サイズ、向き、余白を指定して印刷します。部数は印刷ダイアログで指定してください。",
            TextWrapping = TextWrapping.Wrap,
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 11.5,
            Foreground = new SolidColorBrush(Color.FromArgb(255, 96, 94, 92)),
        };

        var panel = new StackPanel { Spacing = 12, Width = 360 };
        panel.Children.Add(paperCombo);
        panel.Children.Add(orientationCombo);
        panel.Children.Add(marginCombo);
        panel.Children.Add(info);

        var dialog = new ContentDialog
        {
            Title = "印刷",
            Content = panel,
            PrimaryButtonText = "印刷",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        if (await ShowDialogAsync(dialog) != ContentDialogResult.Primary)
        {
            return null;
        }

        var paper = paperCombo.SelectedIndex switch
        {
            1 => "b5",
            2 => "letter",
            3 => "legal",
            _ => "a4",
        };
        var margin = marginCombo.SelectedItem is ComboBoxItem item && item.Tag is double value ? value : 2.54;
        return (paper, orientationCombo.SelectedIndex == 1, margin);
    }

    private static string ApplyRtfPageSettings(string rtf, string paper, bool landscape, double marginCm)
    {
        var (width, height) = paper switch
        {
            "b5" => (10318, 14570),
            "letter" => (12240, 15840),
            "legal" => (12240, 20160),
            _ => (11906, 16838),
        };

        if (landscape)
        {
            (width, height) = (height, width);
        }

        var margin = (int)Math.Round(marginCm / 2.54 * 1440);
        var cleaned = Regex.Replace(rtf, @"\\paperw\d+|\\paperh\d+|\\marg[lrtb]\d+|\\landscape", string.Empty);
        var header = $@"\paperw{width}\paperh{height}\margl{margin}\margr{margin}\margt{margin}\margb{margin}"
            + (landscape ? @"\landscape" : string.Empty);

        var index = cleaned.IndexOf(@"\rtf1", StringComparison.Ordinal);
        return index >= 0
            ? cleaned.Insert(index + 5, header)
            : header + cleaned;
    }

    private async Task<bool> ConfirmDiscardAsync()
    {
        if (!_isModified)
        {
            return true;
        }

        var name = _currentFile is null ? _untitledName : IOPath.GetFileNameWithoutExtension(_currentFile.Name);
        var dialog = new ContentDialog
        {
            Title = "変更を保存しますか?",
            Content = $"「{name}」への変更を保存しますか?",
            PrimaryButtonText = "保存",
            SecondaryButtonText = "保存しない",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        var result = await ShowDialogAsync(dialog);
        if (result == ContentDialogResult.Primary)
        {
            return await SaveAsync(saveAs: false);
        }

        return result == ContentDialogResult.Secondary;
    }

    private async void AppWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_isClosingConfirmed || !_isModified)
        {
            return;
        }

        args.Cancel = true;
        if (await ConfirmDiscardAsync())
        {
            _isClosingConfirmed = true;
            Close();
        }
    }

    private void RootGrid_DragOver(object sender, DragEventArgs e)
    {
        if (e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
            e.DragUIOverride.Caption = "WordPad Lite で開く";
            e.DragUIOverride.IsCaptionVisible = true;
        }
    }

    private async void RootGrid_Drop(object sender, DragEventArgs e)
    {
        if (!e.DataView.Contains(StandardDataFormats.StorageItems))
        {
            return;
        }

        e.Handled = true;
        var items = await e.DataView.GetStorageItemsAsync();
        if (items.FirstOrDefault() is StorageFile file && await ConfirmDiscardAsync())
        {
            await LoadFileAsync(file);
        }
    }

    private static string FormatFromExtension(string extension)
        => extension switch
        {
            ".docx" => "DOCX",
            ".txt" => "TXT",
            _ => "RTF",
        };

    private void Editor_Loaded(object sender, RoutedEventArgs e)
    {
        if (_editorReady)
        {
            return;
        }

        _isLoading = true;
        try
        {
            InitializeEditorDefaults();
            Editor.Document.Selection.SetRange(0, 0);
            Editor.Document.ClearUndoRedoHistory();
        }
        finally
        {
            _isLoading = false;
        }

        _editorReady = true;
        _spellCheckEnabled = Editor.IsSpellCheckEnabled;
        RefreshRibbonToggleStates();
        DrawRuler();
        UpdateTitle();
        UpdateWordCount();
        UpdateFormattingUi();
        UpdateZoomText(PageScrollViewer.ZoomFactor);
        Editor.Focus(FocusState.Programmatic);

        DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
        {
            if (!_editorReady)
            {
                return;
            }

            Editor.Document.ClearUndoRedoHistory();
            _isModified = false;
            UpdateTitle();
            UpdateUndoRedo();
        });
    }

    private void InitializeEditorDefaults()
    {
        Editor.IsSpellCheckEnabled = _spellCheckEnabled;
        var character = Editor.Document.GetDefaultCharacterFormat();
        character.Name = _defaultFontName;
        character.Size = (float)_defaultFontSize;
        character.Bold = FormatEffect.Off;
        character.Italic = FormatEffect.Off;
        character.Underline = UnderlineType.None;
        character.Strikethrough = FormatEffect.Off;
        character.ForegroundColor = Color.FromArgb(255, 37, 36, 35);
        Editor.Document.SetDefaultCharacterFormat(character);

        var paragraph = Editor.Document.GetDefaultParagraphFormat();
        paragraph.Alignment = ParagraphAlignment.Left;
        paragraph.SetLineSpacing(LineSpacingRule.Multiple, 1.15f);
        paragraph.SpaceBefore = 0;
        paragraph.SpaceAfter = 8;
        paragraph.ListType = MarkerType.None;
        paragraph.ListLevelIndex = 0;
        paragraph.SetIndents(0, 0, 0);
        Editor.Document.SetDefaultParagraphFormat(paragraph);
    }

    private void Editor_TextChanged(object sender, RoutedEventArgs e)
    {
        if (!_editorReady || _isLoading)
        {
            return;
        }

        MarkModified();
        ScheduleWordCountUpdate();
        UpdateUndoRedo();
        ScheduleAutoSave();
    }

    private void Editor_SelectionChanged(object sender, RoutedEventArgs e)
    {
        if (!_editorReady)
        {
            return;
        }

        UpdateFormattingUi();

        if (_formatPainterArmed && !_applyingFormatPainter && Editor.Document.Selection.Length > 0)
        {
            ApplyFormatPainterToSelection();
        }
    }

    private void MarkModified()
    {
        if (_isLoading)
        {
            return;
        }

        _isModified = true;
        UpdateTitle();
    }

    private void AutoSaveSwitch_Toggled(object sender, RoutedEventArgs e)
    {
        if (_loadingSettings)
        {
            return;
        }

        var enabled = AutoSaveSwitch.IsOn;
        if (enabled == _autoSaveEnabled)
        {
            return;
        }

        SetAutoSaveEnabled(enabled);
        Notify(enabled
            ? (_currentFile is null
                ? "自動保存をオンにしました。新規文書は回復用ファイルに保存されます。"
                : "自動保存をオンにしました。変更は自動的に保存されます。")
            : "自動保存をオフにしました。");
    }

    private void SetAutoSaveEnabled(bool enabled)
    {
        _autoSaveEnabled = enabled;
        AppSettings.Set("AutoSave", enabled);

        _loadingSettings = true;
        try
        {
            AutoSaveSwitch.IsOn = enabled;
        }
        finally
        {
            _loadingSettings = false;
        }

        if (enabled)
        {
            ScheduleAutoSave();
        }
        else
        {
            _autoSaveTimer?.Stop();
        }
    }

    private void ScheduleAutoSave()
    {
        if (!_autoSaveEnabled || !_editorReady || _isLoading)
        {
            return;
        }

        _autoSaveTimer ??= DispatcherQueue.CreateTimer();
        _autoSaveTimer.Tick -= AutoSaveTimer_Tick;
        _autoSaveTimer.Tick += AutoSaveTimer_Tick;
        _autoSaveTimer.Interval = TimeSpan.FromSeconds(Math.Max(5, _autoSaveIntervalSeconds));
        _autoSaveTimer.IsRepeating = false;
        _autoSaveTimer.Stop();
        _autoSaveTimer.Start();
    }

    private async void AutoSaveTimer_Tick(Microsoft.UI.Dispatching.DispatcherQueueTimer sender, object args)
    {
        await RunAutoSaveAsync();
    }

    private async Task RunAutoSaveAsync()
    {
        if (!_autoSaveEnabled || !_editorReady || !_isModified || _isAutoSaving || _isLoading)
        {
            return;
        }

        _isAutoSaving = true;
        try
        {
            if (_currentFile is not null)
            {
                if (await WriteToFileAsync(_currentFile, silent: true))
                {
                    _lastAutoSaveAt = DateTime.Now;
                    UpdateTitle();
                }
            }
            else
            {
                var folder = IOPath.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordPadLite",
                    "recovery");
                System.IO.Directory.CreateDirectory(folder);
                var safeName = string.Concat(_untitledName.Select(c => IOPath.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                var path = IOPath.Combine(folder, $"{safeName}.rtf");
                var bytes = await GetRtfBytesAsync();
                await System.IO.File.WriteAllBytesAsync(path, bytes);
                _lastAutoSaveAt = DateTime.Now;
                UpdateTitle();
            }
        }
        catch
        {
            // Autosave failures are deliberately silent; the manual save path still reports errors.
        }
        finally
        {
            _isAutoSaving = false;
        }
    }

    private async Task<byte[]> GetRtfBytesAsync()
    {
        using var stream = new InMemoryRandomAccessStream();
        Editor.Document.SaveToStream(TextGetOptions.FormatRtf, stream);
        stream.Seek(0);
        var bytes = new byte[stream.Size];
        using var reader = new DataReader(stream.GetInputStreamAt(0));
        await reader.LoadAsync((uint)stream.Size);
        reader.ReadBytes(bytes);
        return bytes;
    }

    private void ScheduleWordCountUpdate()
    {
        if (_wordCountTimer is null)
        {
            _wordCountTimer = DispatcherQueue.CreateTimer();
            _wordCountTimer.Interval = TimeSpan.FromMilliseconds(250);
            _wordCountTimer.IsRepeating = false;
            _wordCountTimer.Tick += (_, _) => UpdateWordCount();
        }

        _wordCountTimer.Stop();
        _wordCountTimer.Start();
    }

    private void UpdateWordCount()
    {
        var stats = GetDocumentStatistics(GetPlainText());
        WordCountStatusText.Text = $"文字数: {stats.Characters:N0}";
        PageStatusText.Text = $"{stats.Pages}/{stats.Pages}";
        BackstageWordCount.Text = stats.Characters.ToString("N0", CultureInfo.CurrentCulture);
        UpdateLanguageStatus();

        if (ProofingPane.Visibility == Visibility.Visible)
        {
            RefreshProofingPane();
        }
    }

    private void UpdateLanguageStatus()
    {
        LanguageStatusText.Text = $"校正: {(_spellCheckEnabled ? "オン" : "オフ")}";
    }

    private string GetPlainText()
    {
        Editor.Document.GetText(TextGetOptions.None, out var text);
        return text;
    }

    private void UpdateTitle()
    {
        var name = _currentFile is null
            ? _untitledName
            : IOPath.GetFileNameWithoutExtension(_currentFile.Name);

        var saveState = _isModified
            ? (_lastAutoSaveAt is { } autoSaved
                ? $"自動保存 {autoSaved:HH:mm}"
                : "未保存の変更")
            : (_lastAutoSaveAt is { } saved
                ? $"自動保存 {saved:HH:mm}"
                : (_currentFile is null ? "未保存" : "保存済み"));
        WindowTitleText.Text = $"{name} - WordPad Lite";
        SaveStateText.Text = saveState;
        ToolTipService.SetToolTip(WindowTitleText, saveState);
        ModifiedDot.Visibility = _isModified ? Visibility.Visible : Visibility.Collapsed;
        Title = $"{name} - WordPad Lite";
        RefreshQuickAccessState();

        BackstageDocName.Text = name;
        BackstageDocPath.Text = _currentFile?.Path ?? "この PC に保存されていません";
        BackstageFormat.Text = _currentFormat;
    }

    private void UpdateUndoRedo()
    {
        if (!_editorReady)
        {
            return;
        }

        RefreshQuickAccessState();
    }

    private void UpdateBackstageInfo()
    {
        var name = _currentFile is null
            ? _untitledName
            : IOPath.GetFileNameWithoutExtension(_currentFile.Name);
        BackstageDocName.Text = name;
        BackstageDocPath.Text = _currentFile?.Path ?? "この PC に保存されていません";
        BackstageFormat.Text = _currentFormat;

        var count = 0;
        foreach (var character in GetPlainText())
        {
            if (!char.IsWhiteSpace(character) && character != '\r' && character != '\n')
            {
                count++;
            }
        }

        BackstageWordCount.Text = count.ToString("N0", CultureInfo.CurrentCulture);
        UpdateBackstageProtection();
        _ = UpdateBackstageFileInfoAsync();
    }

    private void UpdateBackstageProtection()
    {
        _updatingBackstage = true;
        try
        {
            BackstageReadOnlySwitch.IsOn = _documentMode == DocumentMode.ReadOnly;
            BackstageProtectionText.Text = _documentMode switch
            {
                DocumentMode.Review => "現在のモード: 校閲 (変更履歴を記録)",
                DocumentMode.ReadOnly => "現在のモード: 読み取り専用",
                _ => "現在のモード: 編集",
            };
        }
        finally
        {
            _updatingBackstage = false;
        }
    }

    private async Task UpdateBackstageFileInfoAsync()
    {
        BackstageSizeText.Text = "-";
        BackstageCreatedText.Text = "-";
        BackstageModifiedText.Text = "-";

        if (_currentFile is null)
        {
            BackstageSizeText.Text = "保存前";
            return;
        }

        try
        {
            var properties = await _currentFile.GetBasicPropertiesAsync();
            BackstageSizeText.Text = FormatByteSize(properties.Size);
            BackstageModifiedText.Text = properties.DateModified.LocalDateTime.ToString("yyyy/MM/dd HH:mm", CultureInfo.CurrentCulture);

            var created = System.IO.File.GetCreationTime(_currentFile.Path);
            BackstageCreatedText.Text = created == default(DateTime)
                ? "-"
                : created.ToString("yyyy/MM/dd HH:mm", CultureInfo.CurrentCulture);
        }
        catch
        {
            BackstageSizeText.Text = "-";
        }
    }

    private static string FormatByteSize(ulong bytes)
    {
        if (bytes < 1024)
        {
            return $"{bytes} B";
        }

        if (bytes < 1024 * 1024)
        {
            return $"{bytes / 1024.0:0.0} KB";
        }

        return $"{bytes / (1024.0 * 1024.0):0.0} MB";
    }

    private void SetToggleChecked(ToggleButton toggle, bool value)
    {
        toggle.IsChecked = value;
        if (_compactToggleMap.TryGetValue(toggle, out var mirror))
        {
            mirror.IsChecked = value;
        }
    }

    private void UpdateFormattingUi()
    {
        if (!_editorReady || _isUpdatingFormatUi)
        {
            return;
        }

        _isUpdatingFormatUi = true;
        try
        {
            var selection = Editor.Document.Selection;
            var character = selection.CharacterFormat;
            var paragraph = selection.ParagraphFormat;

            SetToggleChecked(BoldToggle, character.Bold == FormatEffect.On);
            SetToggleChecked(ItalicToggle, character.Italic == FormatEffect.On);
            SetToggleChecked(UnderlineToggle, character.Underline is not (UnderlineType.None or UnderlineType.Undefined));
            StrikeToggle.IsChecked = character.Strikethrough == FormatEffect.On;
            SubscriptToggle.IsChecked = character.Subscript == FormatEffect.On;
            SuperscriptToggle.IsChecked = character.Superscript == FormatEffect.On;

            var fontText = string.IsNullOrWhiteSpace(character.Name)
                ? "游明朝"
                : DisplayFontName(character.Name);
            var sizeText = FormatSize(character.Size);
            FontCombo.Text = fontText;
            CompactFontCombo.Text = fontText;
            FontSizeCombo.Text = sizeText;
            CompactFontSizeCombo.Text = sizeText;

            SetToggleChecked(AlignLeftToggle, paragraph.Alignment == ParagraphAlignment.Left);
            SetToggleChecked(AlignCenterToggle, paragraph.Alignment == ParagraphAlignment.Center);
            SetToggleChecked(AlignRightToggle, paragraph.Alignment == ParagraphAlignment.Right);
            SetToggleChecked(AlignJustifyToggle, paragraph.Alignment == ParagraphAlignment.Justify);

            SetToggleChecked(BulletsToggle, paragraph.ListType is
                MarkerType.Bullet or MarkerType.WhiteCircleWingding or MarkerType.BlackCircleWingding);
            SetToggleChecked(NumberingToggle, paragraph.ListType is
                MarkerType.Arabic or MarkerType.UppercaseEnglishLetter or
                MarkerType.LowercaseEnglishLetter or MarkerType.LowercaseRoman);

            FontColorBar.Background = new SolidColorBrush(character.ForegroundColor);
            CompactFontColorBar.Background = new SolidColorBrush(character.ForegroundColor);
            var background = character.BackgroundColor;
            HighlightColorBar.Background = new SolidColorBrush(
                background.A == 0 || (background.R == 255 && background.G == 255 && background.B == 255)
                    ? Colors.Transparent
                    : background);

            UpdateStyleSelection(character);
        }
        finally
        {
            _isUpdatingFormatUi = false;
        }
    }

    private void UpdateStyleSelection(ITextCharacterFormat character)
    {
        var size = character.Size;
        var bold = character.Bold == FormatEffect.On;
        string? active = null;

        if (Math.Abs(size - 20) < 0.6 && bold)
        {
            active = "heading1";
        }
        else if (Math.Abs(size - 16) < 0.6 && bold)
        {
            active = "heading2";
        }
        else if (Math.Abs(size - 28) < 0.6)
        {
            active = "title";
        }
        else if (Math.Abs(size - 10.5) < 0.4 && !bold)
        {
            active = Editor.Document.Selection.ParagraphFormat.SpaceAfter < 1
                ? "nospacing"
                : "normal";
        }

        SetStyleButtonState(StyleNormalButton, active == "normal");
        SetStyleButtonState(StyleNoSpacingButton, active == "nospacing");
        SetStyleButtonState(StyleHeading1Button, active == "heading1");
        SetStyleButtonState(CompactStyleNormalButton, active == "normal");
        SetStyleButtonState(CompactStyleHeading1Button, active == "heading1");
        SetStyleButtonState(CompactStyleHeading2Button, active == "heading2");
    }

    private void SetStyleButtonState(Button button, bool active)
    {
        if (active)
        {
            button.Background = new SolidColorBrush(IsDarkTheme
                ? Color.FromArgb(255, 24, 62, 94)
                : Color.FromArgb(255, 219, 235, 250));
            button.BorderBrush = (Brush)Application.Current.Resources["WordAccentBrush"];
            button.BorderThickness = new Thickness(1);
            button.Foreground = (Brush)Application.Current.Resources["WordAccentBrush"];
            button.FontWeight = Microsoft.UI.Text.FontWeights.SemiBold;
        }
        else
        {
            button.Background = new SolidColorBrush(Colors.Transparent);
            button.BorderBrush = new SolidColorBrush(Colors.Transparent);
            button.BorderThickness = new Thickness(0);
            button.ClearValue(Control.ForegroundProperty);
            button.FontWeight = Microsoft.UI.Text.FontWeights.Normal;
        }
    }

    private static string FormatSize(float size)
    {
        if (size <= 0)
        {
            return "10.5";
        }

        return Math.Abs(size - MathF.Round(size)) < 0.01f
            ? size.ToString("0", CultureInfo.CurrentCulture)
            : size.ToString("0.#", CultureInfo.CurrentCulture);
    }

    private void FontCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_editorReady || _isUpdatingFormatUi)
        {
            return;
        }

        if (sender is ComboBox combo && combo.SelectedItem is ComboBoxItem item && item.Tag is string name)
        {
            ApplyFontName(name);
        }
    }

    private void FontCombo_TextSubmitted(ComboBox sender, ComboBoxTextSubmittedEventArgs e)
    {
        if (!_editorReady || _isUpdatingFormatUi || string.IsNullOrWhiteSpace(e.Text))
        {
            return;
        }

        ApplyFontName(NormalizeFontName(e.Text));
    }

    private void FontSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_editorReady || _isUpdatingFormatUi)
        {
            return;
        }

        var value = (sender as ComboBox)?.SelectedItem switch
        {
            string text => text,
            ComboBoxItem item => item.Content?.ToString(),
            _ => null,
        };
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out var size))
        {
            ApplyFontSize((float)size);
        }
    }

    private void FontSizeCombo_TextSubmitted(ComboBox sender, ComboBoxTextSubmittedEventArgs e)
    {
        if (!_editorReady || _isUpdatingFormatUi)
        {
            return;
        }

        if (double.TryParse(e.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out var size))
        {
            ApplyFontSize((float)size);
        }
    }

    private static string NormalizeFontName(string value)
    {
        var suffix = " (本文のフォント)";
        return value.EndsWith(suffix, StringComparison.Ordinal)
            ? value[..^suffix.Length]
            : value.Trim();
    }

    private static string DisplayFontName(string name) => name switch
    {
        "Yu Mincho" => "游明朝",
        "Yu Gothic" => "游ゴシック",
        "Yu Gothic UI" => "游ゴシック UI",
        "MS Mincho" => "MS 明朝",
        "MS Gothic" => "MS ゴシック",
        "Meiryo" => "メイリオ",
        "Meiryo UI" => "メイリオ UI",
        "BIZ UDPGothic" => "BIZ UDPゴシック",
        "BIZ UDPMincho" => "BIZ UDP明朝",
        _ => name,
    };

    private void ApplyFontName(string name)
    {
        try
        {
            Editor.Document.Selection.CharacterFormat.Name = name;
            AfterFormatChange();
        }
        catch (Exception ex)
        {
            Notify($"フォントを変更できませんでした: {ex.Message}", InfoBarSeverity.Warning);
        }
    }

    private void ApplyFontSize(float size)
    {
        if (size is < 1 or > 300)
        {
            return;
        }

        try
        {
            Editor.Document.Selection.CharacterFormat.Size = size;
            AfterFormatChange();
        }
        catch (Exception ex)
        {
            Notify($"フォントサイズを変更できませんでした: {ex.Message}", InfoBarSeverity.Warning);
        }
    }

    private void ChangeFontSize(float delta)
    {
        var current = Editor.Document.Selection.CharacterFormat.Size;
        if (current <= 0)
        {
            current = 10.5f;
        }

        ApplyFontSize(Math.Clamp(current + delta, 1, 300));
    }

    private void ToggleBold()
    {
        var format = Editor.Document.Selection.CharacterFormat;
        format.Bold = format.Bold == FormatEffect.On ? FormatEffect.Off : FormatEffect.On;
        AfterFormatChange();
    }

    private void ToggleItalic()
    {
        var format = Editor.Document.Selection.CharacterFormat;
        format.Italic = format.Italic == FormatEffect.On ? FormatEffect.Off : FormatEffect.On;
        AfterFormatChange();
    }

    private void ToggleUnderline()
    {
        var format = Editor.Document.Selection.CharacterFormat;
        format.Underline = format.Underline is UnderlineType.None or UnderlineType.Undefined
            ? UnderlineType.Single
            : UnderlineType.None;
        AfterFormatChange();
    }

    private void ApplyUnderline(UnderlineType underline)
    {
        Editor.Document.Selection.CharacterFormat.Underline = underline;
        AfterFormatChange();
    }

    private void ToggleStrikethrough()
    {
        var format = Editor.Document.Selection.CharacterFormat;
        format.Strikethrough = format.Strikethrough == FormatEffect.On ? FormatEffect.Off : FormatEffect.On;
        AfterFormatChange();
    }

    private void ToggleScript(bool superscript)
    {
        var format = Editor.Document.Selection.CharacterFormat;
        if (superscript)
        {
            var next = format.Superscript == FormatEffect.On ? FormatEffect.Off : FormatEffect.On;
            format.Superscript = next;
            if (next == FormatEffect.On)
            {
                format.Subscript = FormatEffect.Off;
            }
        }
        else
        {
            var next = format.Subscript == FormatEffect.On ? FormatEffect.Off : FormatEffect.On;
            format.Subscript = next;
            if (next == FormatEffect.On)
            {
                format.Superscript = FormatEffect.Off;
            }
        }

        AfterFormatChange();
    }

    private void ApplyTextColor(Color color, bool highlight)
    {
        var format = Editor.Document.Selection.CharacterFormat;
        if (highlight)
        {
            format.BackgroundColor = color.A == 0 ? NoHighlightColor : color;
        }
        else
        {
            format.ForegroundColor = color;
        }

        if (color.A == 0)
        {
            if (highlight)
            {
                HighlightColorBar.Background = new SolidColorBrush(Colors.Transparent);
            }
        }
        else
        {
            if (highlight)
            {
                HighlightColorBar.Background = new SolidColorBrush(color);
            }
            else
            {
                FontColorBar.Background = new SolidColorBrush(color);
            }
        }

        AfterFormatChange();
    }

    private void ApplyParagraphBackground(Color color)
    {
        var range = Editor.Document.Selection.GetClone();
        range.Expand(TextRangeUnit.Paragraph);
        range.CharacterFormat.BackgroundColor = color.A == 0 ? NoHighlightColor : color;
        AfterFormatChange();
    }

    private void ApplyAlignment(ParagraphAlignment alignment)
    {
        Editor.Document.Selection.ParagraphFormat.Alignment = alignment;
        AfterFormatChange();
    }

    private void ApplyList(MarkerType markerType, MarkerStyle markerStyle)
    {
        var format = Editor.Document.Selection.ParagraphFormat;
        if (format.ListType == markerType)
        {
            format.ListType = MarkerType.None;
            format.ListLevelIndex = 0;
        }
        else
        {
            format.ListType = markerType;
            format.ListStyle = markerStyle;
            format.ListStart = 1;
            format.ListLevelIndex = 1;
            format.ListTab = 18;
            format.SetIndents(36, format.RightIndent, 0);
        }

        AfterFormatChange();
    }

    private void RemoveList()
    {
        var format = Editor.Document.Selection.ParagraphFormat;
        format.ListType = MarkerType.None;
        format.ListLevelIndex = 0;
        AfterFormatChange();
    }

    private void ChangeIndent(float delta)
    {
        var format = Editor.Document.Selection.ParagraphFormat;
        var left = Math.Max(0, format.LeftIndent + delta);
        format.SetIndents(left, format.RightIndent, format.FirstLineIndent);
        AfterFormatChange();
    }

    private void ApplyLineSpacing(float multiple)
    {
        Editor.Document.Selection.ParagraphFormat.SetLineSpacing(LineSpacingRule.Multiple, multiple);
        AfterFormatChange();
    }

    private void ApplyParagraphSpace(float? before, float? after)
    {
        var format = Editor.Document.Selection.ParagraphFormat;
        if (before.HasValue)
        {
            format.SpaceBefore = before.Value;
        }
        if (after.HasValue)
        {
            format.SpaceAfter = after.Value;
        }
        AfterFormatChange();
    }

    private void ChangeTextCase(string mode)
    {
        var selection = Editor.Document.Selection;
        if (selection.Length == 0)
        {
            Notify("大文字/小文字を変換する文字列を選択してください。");
            return;
        }

        var start = selection.StartPosition;
        selection.GetText(TextGetOptions.None, out var text);
        var converted = mode switch
        {
            "upper" => text.ToUpper(CultureInfo.CurrentCulture),
            "lower" => text.ToLower(CultureInfo.CurrentCulture),
            "title" => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLower(CultureInfo.CurrentCulture)),
            _ => ToggleCase(text),
        };

        if (converted == text)
        {
            return;
        }

        selection.SetText(TextSetOptions.None, converted);
        selection.SetRange(start, start + converted.Length);
        AfterFormatChange();
    }

    private static string ToggleCase(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (var character in text)
        {
            if (char.IsUpper(character))
            {
                builder.Append(char.ToLowerInvariant(character));
            }
            else if (char.IsLower(character))
            {
                builder.Append(char.ToUpperInvariant(character));
            }
            else
            {
                builder.Append(character);
            }
        }
        return builder.ToString();
    }

    private void ClearFormatting()
    {
        var selection = Editor.Document.Selection;
        var character = selection.CharacterFormat;
        character.Name = "游明朝";
        character.Size = 10.5f;
        character.Bold = FormatEffect.Off;
        character.Italic = FormatEffect.Off;
        character.Underline = UnderlineType.None;
        character.Strikethrough = FormatEffect.Off;
        character.Subscript = FormatEffect.Off;
        character.Superscript = FormatEffect.Off;
        character.ForegroundColor = Color.FromArgb(255, 37, 36, 35);
        character.BackgroundColor = NoHighlightColor;

        var paragraph = selection.ParagraphFormat;
        paragraph.Alignment = ParagraphAlignment.Left;
        paragraph.ListType = MarkerType.None;
        paragraph.ListLevelIndex = 0;
        paragraph.SetIndents(0, 0, 0);
        paragraph.SetLineSpacing(LineSpacingRule.Multiple, 1.15f);
        paragraph.SpaceBefore = 0;
        paragraph.SpaceAfter = 8;

        AfterFormatChange();
    }

    private void ApplyStyle(string style)
    {
        var selection = Editor.Document.Selection;
        var originalStart = selection.StartPosition;
        var originalEnd = selection.EndPosition;
        var collapsed = selection.Length == 0;

        if (collapsed)
        {
            var paragraph = selection.GetClone();
            paragraph.Expand(TextRangeUnit.Paragraph);
            selection.SetRange(paragraph.StartPosition, paragraph.EndPosition);
        }

        var character = selection.CharacterFormat;
        var paragraphFormat = selection.ParagraphFormat;
        character.BackgroundColor = NoHighlightColor;
        character.Italic = FormatEffect.Off;
        character.Underline = UnderlineType.None;
        paragraphFormat.ListType = MarkerType.None;
        paragraphFormat.ListLevelIndex = 0;

        switch (style)
        {
            case "heading1":
                character.Name = "游ゴシック";
                character.Size = 20;
                character.Bold = FormatEffect.On;
                character.ForegroundColor = Color.FromArgb(255, 31, 56, 100);
                paragraphFormat.SpaceBefore = 12;
                paragraphFormat.SpaceAfter = 6;
                paragraphFormat.SetLineSpacing(LineSpacingRule.Multiple, 1.05f);
                break;
            case "heading2":
                character.Name = "游ゴシック";
                character.Size = 16;
                character.Bold = FormatEffect.On;
                character.ForegroundColor = Color.FromArgb(255, 47, 84, 150);
                paragraphFormat.SpaceBefore = 10;
                paragraphFormat.SpaceAfter = 4;
                paragraphFormat.SetLineSpacing(LineSpacingRule.Multiple, 1.05f);
                break;
            case "nospacing":
                character.Name = "游明朝";
                character.Size = 10.5f;
                character.Bold = FormatEffect.Off;
                character.ForegroundColor = Color.FromArgb(255, 37, 36, 35);
                paragraphFormat.SpaceBefore = 0;
                paragraphFormat.SpaceAfter = 0;
                paragraphFormat.SetLineSpacing(LineSpacingRule.Multiple, 1);
                break;
            case "title":
                character.Name = "游明朝";
                character.Size = 28;
                character.Bold = FormatEffect.Off;
                character.ForegroundColor = Color.FromArgb(255, 23, 54, 93);
                paragraphFormat.SpaceBefore = 0;
                paragraphFormat.SpaceAfter = 12;
                paragraphFormat.SetLineSpacing(LineSpacingRule.Multiple, 1);
                break;
            default:
                character.Name = "游明朝";
                character.Size = 10.5f;
                character.Bold = FormatEffect.Off;
                character.ForegroundColor = Color.FromArgb(255, 37, 36, 35);
                paragraphFormat.SpaceBefore = 0;
                paragraphFormat.SpaceAfter = 8;
                paragraphFormat.SetLineSpacing(LineSpacingRule.Multiple, 1.15f);
                break;
        }

        if (collapsed)
        {
            selection.SetRange(originalStart, originalEnd);
        }

        AfterFormatChange();
    }

    private void CaptureFormatPainter()
    {
        if (_formatPainterArmed)
        {
            _formatPainterArmed = false;
            _formatPainterCharacter = null;
            _formatPainterParagraph = null;
            Notify("書式のコピーを解除しました。");
            return;
        }

        var selection = Editor.Document.Selection;
        _formatPainterCharacter = selection.CharacterFormat.GetClone();
        _formatPainterParagraph = selection.ParagraphFormat.GetClone();
        _formatPainterArmed = true;
        Notify("書式をコピーしました。適用先の文字列を選択してください。");
    }

    private void ApplyFormatPainterToSelection()
    {
        if (_formatPainterCharacter is null || _formatPainterParagraph is null)
        {
            _formatPainterArmed = false;
            return;
        }

        _applyingFormatPainter = true;
        try
        {
            var selection = Editor.Document.Selection;
            selection.CharacterFormat.SetClone(_formatPainterCharacter);
            selection.ParagraphFormat.SetClone(_formatPainterParagraph);
            _formatPainterArmed = false;
            _formatPainterCharacter = null;
            _formatPainterParagraph = null;
            MarkModified();
            Notify("書式を適用しました。");
        }
        finally
        {
            _applyingFormatPainter = false;
        }
    }

    private void AfterFormatChange()
    {
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
        UpdateFormattingUi();
        UpdateUndoRedo();
    }

    private void InsertText(string text)
    {
        var selection = Editor.Document.Selection;
        selection.SetText(TextSetOptions.None, text);
        selection.Collapse(false);
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
    }

    private void InsertRtf(string rtf)
    {
        var selection = Editor.Document.Selection;
        selection.SetText(TextSetOptions.FormatRtf, rtf);
        selection.Collapse(false);
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
    }

    private static Color ParseColor(string value)
    {
        value = value.TrimStart('#');
        if (value.Length == 8)
        {
            return Color.FromArgb(
                byte.Parse(value[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(value[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(value[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(value[6..8], NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }

        if (value.Length == 6)
        {
            return Color.FromArgb(
                255,
                byte.Parse(value[..2], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(value[2..4], NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(value[4..6], NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }

        return Color.FromArgb(255, 37, 36, 35);
    }

    private void ApplyPageSize(string size, bool landscape)
    {
        _pageSizeName = size.ToLowerInvariant();
        _landscape = landscape;

        var (width, height) = _pageSizeName switch
        {
            "b5" => (709d, 1002d),
            "letter" => (816d, 1056d),
            "legal" => (816d, 1344d),
            _ => (A4Width, A4Height),
        };

        _pageWidth = landscape ? height : width;
        _pageHeight = landscape ? width : height;
        ApplyPageLayout();
    }

    private void ApplyPageLayout()
    {
        PageBorder.Width = _pageWidth;
        PageBorder.Height = _pageHeight;
        PageFrame.Width = _pageWidth + 10;
        PageFrame.Height = _pageHeight + 10;
        RulerHost.Width = _pageWidth;
        RulerCanvas.Width = _pageWidth;
        VerticalRulerHost.Height = _pageHeight;
        VerticalRulerCanvas.Height = _pageHeight;
        Editor.Padding = new Thickness(_pageMargin);
        DrawRuler();
        DrawGridLines();
    }

    private void RootGrid_ActualThemeChanged(FrameworkElement sender, object args)
    {
        ApplyTitleBarTheme();
        DrawRuler();
        RebuildCompactRibbons();
    }

    private void DrawRuler()
    {
        if (RulerCanvas is null || VerticalRulerCanvas is null)
        {
            return;
        }

        RulerCanvas.Children.Clear();
        VerticalRulerCanvas.Children.Clear();
        RulerCanvas.Width = _pageWidth;
        VerticalRulerCanvas.Height = _pageHeight;

        var dark = IsDarkTheme;
        var tickBrush = new SolidColorBrush(dark ? Color.FromArgb(255, 150, 148, 146) : Color.FromArgb(255, 136, 134, 132));
        var numberBrush = new SolidColorBrush(dark ? Color.FromArgb(255, 200, 198, 196) : Color.FromArgb(255, 96, 94, 92));
        var bandBrush = new SolidColorBrush(dark
            ? Color.FromArgb(0x45, 0x4F, 0xA3, 0xE3)
            : Color.FromArgb(0x33, 0x0F, 0x6C, 0xBD));
        var outsideBrush = new SolidColorBrush(dark
            ? Color.FromArgb(0x30, 0x00, 0x00, 0x00)
            : Color.FromArgb(0x16, 0x00, 0x00, 0x00));
        var marginBrush = new SolidColorBrush(dark
            ? Color.FromArgb(255, 79, 163, 227)
            : Color.FromArgb(255, 15, 108, 189));

        // The tinted band marks the editable text area; the darker edges mark the
        // page margins and clearly separate the writable range from the rest.
        var horizontalBand = new Rectangle
        {
            Width = Math.Max(0, _pageWidth - (_pageMargin * 2)),
            Height = 21,
            Fill = bandBrush,
        };
        Canvas.SetLeft(horizontalBand, _pageMargin);
        Canvas.SetTop(horizontalBand, 0);
        RulerCanvas.Children.Add(horizontalBand);

        var horizontalLeft = new Rectangle { Width = _pageMargin, Height = 21, Fill = outsideBrush };
        Canvas.SetLeft(horizontalLeft, 0);
        RulerCanvas.Children.Add(horizontalLeft);
        var horizontalRight = new Rectangle { Width = _pageMargin, Height = 21, Fill = outsideBrush };
        Canvas.SetLeft(horizontalRight, _pageWidth - _pageMargin);
        RulerCanvas.Children.Add(horizontalRight);

        var verticalBand = new Rectangle
        {
            Width = 21,
            Height = Math.Max(0, _pageHeight - (_pageMargin * 2)),
            Fill = bandBrush,
        };
        Canvas.SetTop(verticalBand, _pageMargin);
        VerticalRulerCanvas.Children.Add(verticalBand);

        var verticalTop = new Rectangle { Width = 21, Height = _pageMargin, Fill = outsideBrush };
        VerticalRulerCanvas.Children.Add(verticalTop);
        var verticalBottom = new Rectangle { Width = 21, Height = _pageMargin, Fill = outsideBrush };
        Canvas.SetTop(verticalBottom, _pageHeight - _pageMargin);
        VerticalRulerCanvas.Children.Add(verticalBottom);

        foreach (var margin in new[] { _pageMargin, _pageWidth - _pageMargin })
        {
            RulerCanvas.Children.Add(new Line
            {
                X1 = margin,
                X2 = margin,
                Y1 = 0,
                Y2 = 21,
                Stroke = marginBrush,
                StrokeThickness = 1,
            });
        }

        foreach (var margin in new[] { _pageMargin, _pageHeight - _pageMargin })
        {
            VerticalRulerCanvas.Children.Add(new Line
            {
                X1 = 0,
                X2 = 21,
                Y1 = margin,
                Y2 = margin,
                Stroke = marginBrush,
                StrokeThickness = 1,
            });
        }

        var quarterCm = PixelsPerCm / 4;

        var horizontalTicks = (int)Math.Floor(_pageWidth / quarterCm) + 1;
        for (var index = 0; index <= horizontalTicks; index++)
        {
            var x = index * quarterCm;
            var isCentimeter = index % 4 == 0;
            RulerCanvas.Children.Add(new Line
            {
                X1 = x,
                X2 = x,
                Y1 = isCentimeter ? 7 : 14,
                Y2 = 21,
                Stroke = tickBrush,
                StrokeThickness = isCentimeter ? 1 : 0.7,
            });

            if (isCentimeter && index > 0)
            {
                var label = new TextBlock
                {
                    Text = (index / 4).ToString(CultureInfo.CurrentCulture),
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 8,
                    Foreground = numberBrush,
                };
                Canvas.SetLeft(label, x - 4);
                Canvas.SetTop(label, 0);
                RulerCanvas.Children.Add(label);
            }
        }

        var verticalTicks = (int)Math.Floor(_pageHeight / quarterCm) + 1;
        for (var index = 0; index <= verticalTicks; index++)
        {
            var y = index * quarterCm;
            var isCentimeter = index % 4 == 0;
            VerticalRulerCanvas.Children.Add(new Line
            {
                X1 = isCentimeter ? 7 : 14,
                X2 = 21,
                Y1 = y,
                Y2 = y,
                Stroke = tickBrush,
                StrokeThickness = isCentimeter ? 1 : 0.7,
            });

            if (isCentimeter && index > 0)
            {
                var label = new TextBlock
                {
                    Text = (index / 4).ToString(CultureInfo.CurrentCulture),
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 8,
                    Foreground = numberBrush,
                    RenderTransformOrigin = new Point(0.5, 0.5),
                    RenderTransform = new RotateTransform { Angle = -90 },
                };
                Canvas.SetLeft(label, 2);
                Canvas.SetTop(label, y - 4);
                VerticalRulerCanvas.Children.Add(label);
            }
        }

        // Blue margin markers on both rulers.
        foreach (var margin in new[] { _pageMargin, _pageWidth - _pageMargin })
        {
            RulerCanvas.Children.Add(new Line
            {
                X1 = margin,
                X2 = margin,
                Y1 = 0,
                Y2 = 6,
                Stroke = marginBrush,
                StrokeThickness = 2,
            });
        }

        foreach (var margin in new[] { _pageMargin, _pageHeight - _pageMargin })
        {
            VerticalRulerCanvas.Children.Add(new Line
            {
                X1 = 15,
                X2 = 21,
                Y1 = margin,
                Y2 = margin,
                Stroke = marginBrush,
                StrokeThickness = 2,
            });
        }
    }

    private void DrawGridLines()
    {
        if (GridLinesOverlay is null)
        {
            return;
        }

        GridLinesOverlay.Children.Clear();
        GridLinesOverlay.Width = _pageWidth;
        GridLinesOverlay.Height = _pageHeight;
        if (!_gridLinesVisible)
        {
            return;
        }

        var gridBrush = new SolidColorBrush(Color.FromArgb(30, 15, 108, 189));
        for (var x = 0d; x <= _pageWidth; x += 24)
        {
            GridLinesOverlay.Children.Add(new Line
            {
                X1 = x,
                X2 = x,
                Y1 = 0,
                Y2 = _pageHeight,
                Stroke = gridBrush,
                StrokeThickness = 0.75,
            });
        }

        for (var y = 0d; y <= _pageHeight; y += 24)
        {
            GridLinesOverlay.Children.Add(new Line
            {
                X1 = 0,
                X2 = _pageWidth,
                Y1 = y,
                Y2 = y,
                Stroke = gridBrush,
                StrokeThickness = 0.75,
            });
        }
    }

    private void UpdateRulerVisibility()
    {
        var visible = _rulerVisible && !_readingMode ? Visibility.Visible : Visibility.Collapsed;
        RulerHost.Visibility = visible;
        VerticalRulerHost.Visibility = visible;
        RefreshRibbonToggleStates();
    }

    private void ToggleReadingMode()
    {
        _readingMode = !_readingMode;
        RootGrid.RowDefinitions[1].Height = new GridLength(_readingMode ? 0 : 30);
        RootGrid.RowDefinitions[2].Height = _readingMode ? new GridLength(0) : GridLength.Auto;
        ReadingModeExitButton.Visibility = _readingMode ? Visibility.Visible : Visibility.Collapsed;
        UpdateRulerVisibility();
        RefreshRibbonToggleStates();
        Editor.Focus(FocusState.Programmatic);
    }

    private void SetZoom(double percent)
    {
        var value = Math.Clamp(percent, 50, 250);
        _isUpdatingZoom = true;
        try
        {
            PageScrollViewer.ChangeView(null, null, (float)(value / 100), true);
            ZoomSlider.Value = value;
            ZoomText.Text = $"{value:0}%";
        }
        finally
        {
            _isUpdatingZoom = false;
        }
    }

    private void UpdateZoomText(double factor)
    {
        var percent = Math.Round(factor * 100);
        ZoomText.Text = $"{percent:0}%";
        _isUpdatingZoom = true;
        try
        {
            ZoomSlider.Value = Math.Clamp(percent, 50, 250);
        }
        finally
        {
            _isUpdatingZoom = false;
        }
    }

    private void PageScrollViewer_ViewChanged(object sender, ScrollViewerViewChangedEventArgs e)
    {
        UpdateZoomText(PageScrollViewer.ZoomFactor);
    }

    private void ZoomSlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_isUpdatingZoom)
        {
            return;
        }

        SetZoom(e.NewValue);
    }

    private void ToggleSpellCheck()
    {
        SetSpellCheckEnabled(!_spellCheckEnabled);
        Notify(_spellCheckEnabled
            ? "スペルチェックをオンにしました。"
            : "スペルチェックをオフにしました。");
    }

    private void SetSpellCheckEnabled(bool enabled)
    {
        _spellCheckEnabled = enabled;
        Editor.IsSpellCheckEnabled = enabled;
        AppSettings.Set("SpellCheck", enabled);
        UpdateLanguageStatus();
        RefreshRibbonToggleStates();
        if (ProofingPane.Visibility == Visibility.Visible)
        {
            _updatingProofingPane = true;
            try
            {
                ProofingSpellCheck.IsChecked = enabled;
            }
            finally
            {
                _updatingProofingPane = false;
            }

            RefreshProofingPane();
        }
    }

    private async Task InsertCoverAsync()
    {
        var titleBox = new TextBox
        {
            Header = "タイトル",
            Text = "文書タイトル",
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
        };
        var subtitleBox = new TextBox
        {
            Header = "サブタイトル",
            Text = DateTime.Now.ToString("yyyy年M月d日", CultureInfo.GetCultureInfo("ja-JP")),
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
        };
        var panel = new StackPanel { Spacing = 10, Width = 380 };
        panel.Children.Add(titleBox);
        panel.Children.Add(subtitleBox);

        var dialog = new ContentDialog
        {
            Title = "表紙を挿入",
            Content = panel,
            PrimaryButtonText = "挿入",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        if (await ShowDialogAsync(dialog) != ContentDialogResult.Primary)
        {
            return;
        }

        var title = RtfEscape(string.IsNullOrWhiteSpace(titleBox.Text) ? "文書タイトル" : titleBox.Text.Trim());
        var subtitle = RtfEscape(subtitleBox.Text.Trim());
        var rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Yu Mincho;}{\\f1 Yu Gothic;}}" +
                  "\\pard\\qc\\f0\\fs64\\b " + title + "\\b0\\par" +
                  "\\fs28\\f1 " + subtitle + "\\par\\page}";

        var range = Editor.Document.GetRange(0, 0);
        range.SetText(TextSetOptions.FormatRtf, rtf);
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
        Notify("表紙を挿入しました。");
    }

    private void InsertHeaderFooterInline(bool header)
    {
        // Insert a real, selected placeholder at the insertion point so the user
        // types directly in the document instead of a dimmed modal dialog.
        const string placeholder = "ここに入力";
        var label = header ? "ヘッダー" : "フッター";
        var content = header ? $"{label}: {placeholder}\r" : $"\r{label}: {placeholder}";
        var position = header ? 0 : GetPlainText().Length;
        var range = Editor.Document.GetRange(position, position);
        range.SetText(TextSetOptions.None, content);

        var prefixLength = label.Length + 2;
        var placeholderStart = position + prefixLength;
        var inserted = Editor.Document.GetRange(position, position + content.Length);
        inserted.CharacterFormat.Size = 9;
        inserted.CharacterFormat.ForegroundColor = Color.FromArgb(255, 96, 94, 92);
        inserted.CharacterFormat.Name = "Yu Gothic UI";

        Editor.Document.Selection.SetRange(placeholderStart, placeholderStart + placeholder.Length);
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
        Notify($"{label}を本文の{(header ? "先頭" : "末尾")}に挿入しました。選択中の文字を入力してください。");
    }

    private void InsertTable()
    {
        var builder = new StringBuilder();
        builder.Append("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Yu Mincho;}}");
        for (var row = 0; row < 3; row++)
        {
            builder.Append("\\trowd\\trgaph108\\trleft0");
            for (var column = 0; column < 3; column++)
            {
                builder.Append("\\clvertalt\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10");
                builder.Append("\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10");
                builder.Append("\\cellx").Append((column + 1) * 3000);
            }
            for (var column = 0; column < 3; column++)
            {
                builder.Append("\\intbl\\f0\\fs21 ");
                builder.Append(RtfEscape($"セル {row + 1}-{column + 1}"));
                builder.Append("\\cell ");
            }
            builder.Append("\\row");
        }
        builder.Append("\\pard\\par}");

        InsertRtf(builder.ToString());
        Notify("3 × 3 の表を挿入しました。");
    }

    private async Task InsertPictureAsync()
    {
        var picker = new FileOpenPicker
        {
            ViewMode = PickerViewMode.Thumbnail,
            SuggestedStartLocation = PickerLocationId.PicturesLibrary,
        };
        picker.FileTypeFilter.Add(".png");
        picker.FileTypeFilter.Add(".jpg");
        picker.FileTypeFilter.Add(".jpeg");
        picker.FileTypeFilter.Add(".bmp");
        picker.FileTypeFilter.Add(".gif");
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            return;
        }

        try
        {
            using var stream = await file.OpenAsync(FileAccessMode.Read);
            var decoder = await BitmapDecoder.CreateAsync(stream);
            var width = (double)decoder.PixelWidth;
            var height = (double)decoder.PixelHeight;
            var maxWidth = Math.Max(120, _pageWidth - _pageMargin * 2);
            var scale = Math.Min(1, maxWidth / Math.Max(1, width));
            if (height * scale > 520)
            {
                scale = 520 / Math.Max(1, height);
            }

            var insertWidth = Math.Max(16, (int)Math.Round(width * scale));
            var insertHeight = Math.Max(16, (int)Math.Round(height * scale));
            stream.Seek(0);
            Editor.Document.Selection.InsertImage(
                insertWidth,
                insertHeight,
                0,
                VerticalCharacterAlignment.Baseline,
                file.Name,
                stream);

            MarkModified();
            Editor.Focus(FocusState.Programmatic);
            Notify($"{file.Name} を挿入しました。");
        }
        catch (Exception ex)
        {
            Notify($"画像を挿入できませんでした: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async Task InsertLinkAsync()
    {
        var selection = Editor.Document.Selection;
        string? selectedText = null;
        if (selection.Length > 0)
        {
            selection.GetText(TextGetOptions.None, out selectedText);
        }

        var displayBox = new TextBox
        {
            Header = "表示する文字列",
            Text = string.IsNullOrWhiteSpace(selectedText) ? "リンク" : selectedText.Trim(),
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
        };
        var urlBox = new TextBox
        {
            Header = "リンク先 (URL)",
            PlaceholderText = "https://example.com",
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
        };
        var panel = new StackPanel { Spacing = 10, Width = 380 };
        panel.Children.Add(displayBox);
        panel.Children.Add(urlBox);

        var dialog = new ContentDialog
        {
            Title = "リンクの挿入",
            Content = panel,
            PrimaryButtonText = "挿入",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        if (await ShowDialogAsync(dialog) != ContentDialogResult.Primary)
        {
            return;
        }

        var url = urlBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(url))
        {
            Notify("リンク先の URL を入力してください。", InfoBarSeverity.Warning);
            return;
        }

        try
        {
            var display = string.IsNullOrWhiteSpace(displayBox.Text) ? url : displayBox.Text.Trim();
            var start = selection.StartPosition;
            if (selection.Length > 0)
            {
                selection.GetText(TextGetOptions.None, out var current);
                if (current != display)
                {
                    selection.SetText(TextSetOptions.None, display);
                }
            }
            else
            {
                selection.SetText(TextSetOptions.None, display);
            }

            selection.SetRange(start, start + display.Length);
            selection.Link = url;
            selection.CharacterFormat.ForegroundColor = Color.FromArgb(255, 15, 108, 189);
            selection.CharacterFormat.Underline = UnderlineType.Single;
            selection.SetRange(start + display.Length, start + display.Length);
            MarkModified();
            Editor.Focus(FocusState.Programmatic);
            Notify("リンクを挿入しました。");
        }
        catch (Exception ex)
        {
            Notify($"リンクを挿入できませんでした: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async Task InsertBookmarkAsync()
    {
        var name = await PromptTextAsync("ブックマーク", "ブックマーク名", "Bookmark1");
        if (string.IsNullOrWhiteSpace(name))
        {
            return;
        }

        InsertText($"[ブックマーク: {name.Trim()}]");
        Notify("簡易ブックマークを挿入しました。");
    }

    private async Task InsertCommentAsync()
    {
        var text = await PromptTextAsync("コメント", "コメントの内容", string.Empty);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var selection = Editor.Document.Selection;
        var start = selection.StartPosition;
        var commentText = $"【コメント: {text.Trim()}】";
        selection.SetText(TextSetOptions.None, commentText);
        var range = Editor.Document.GetRange(start, start + commentText.Length);
        range.CharacterFormat.BackgroundColor = Color.FromArgb(255, 255, 242, 204);
        range.CharacterFormat.Italic = FormatEffect.On;
        selection.SetRange(start + commentText.Length, start + commentText.Length);
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
        Notify("コメントを挿入しました。");
    }

    private async Task InsertWatermarkAsync()
    {
        var text = await PromptTextAsync("透かし", "透かしの文字列", "DRAFT");
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        var range = Editor.Document.GetRange(0, 0);
        var value = text.Trim() + "\r";
        range.SetText(TextSetOptions.None, value);
        var inserted = Editor.Document.GetRange(0, value.Length);
        inserted.CharacterFormat.Name = "Yu Gothic UI";
        inserted.CharacterFormat.Size = 42;
        inserted.CharacterFormat.Bold = FormatEffect.On;
        inserted.CharacterFormat.ForegroundColor = Color.FromArgb(70, 128, 128, 128);
        MarkModified();
        Notify("透かし文字を本文の先頭に挿入しました。");
    }

    private void InsertTocPlaceholder()
    {
        InsertText("目次\r（見出し 1 / 見出し 2 を適用した段落がここに表示されます）\r\r");
        Notify("目次の簡易版を挿入しました。");
    }

    private void InsertFootnoteInline(bool endnote)
    {
        // Place the reference marker, append the note block, then select the
        // placeholder so the note body is typed directly into the document.
        const string placeholder = "注の内容を入力";
        _footnoteNumber++;
        var label = endnote ? "文末脚注" : "脚注";
        var selection = Editor.Document.Selection;
        var markerStart = selection.StartPosition;
        var marker = $"[{_footnoteNumber}]";
        selection.SetText(TextSetOptions.None, marker);
        var markerRange = Editor.Document.GetRange(markerStart, markerStart + marker.Length);
        markerRange.CharacterFormat.Superscript = FormatEffect.On;

        var length = GetPlainText().Length;
        var prefix = $"\r{label} [{_footnoteNumber}] ";
        var noteText = $"{prefix}{placeholder}\r";
        var noteRange = Editor.Document.GetRange(length, length);
        noteRange.SetText(TextSetOptions.None, noteText);
        var styledNote = Editor.Document.GetRange(length, length + noteText.Length);
        styledNote.CharacterFormat.Size = 9;
        styledNote.CharacterFormat.ForegroundColor = Color.FromArgb(255, 96, 94, 92);
        styledNote.CharacterFormat.Name = "Yu Gothic UI";

        var placeholderStart = length + prefix.Length;
        Editor.Document.Selection.SetRange(placeholderStart, placeholderStart + placeholder.Length);
        MarkModified();
        Editor.Focus(FocusState.Programmatic);
        Notify($"{label}を挿入しました。選択中の文字を入力してください。");
    }

    private async Task InsertCitationAsync()
    {
        var text = await PromptTextAsync("引用文献の挿入", "著者名、年、タイトルなど", string.Empty);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        InsertText($"({text.Trim()})");
    }

    private async Task InsertCaptionAsync()
    {
        var text = await PromptTextAsync("図表番号の挿入", "図表の説明", string.Empty);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        InsertText($"\r図 1: {text.Trim()}\r");
    }

    private async Task InsertEnvelopeAsync()
    {
        var recipient = await PromptTextAsync("封筒", "宛先", string.Empty);
        if (string.IsNullOrWhiteSpace(recipient))
        {
            return;
        }

        var sender = await PromptTextAsync("封筒", "差出人", string.Empty);
        var text = $"宛先: {recipient.Trim()}\r差出人: {sender?.Trim()}\r\r";
        var range = Editor.Document.GetRange(0, 0);
        range.SetText(TextSetOptions.None, text);
        MarkModified();
        Notify("宛名を本文の先頭に挿入しました。");
    }

    private async Task InsertLabelAsync()
    {
        var text = await PromptTextAsync("宛名ラベル", "ラベルに印刷する文字列", string.Empty);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        InsertText(text.Trim() + "\r");
        Notify("宛名ラベル用の文字列を挿入しました。");
    }

    private async Task ShowMailMergeStartAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "差し込み印刷の開始",
            Content = "差し込みフィールドを本文に挿入して、宛先情報を差し込む準備をします。\n\n「差し込みフィールド」から «氏名» などを選んでください。",
            PrimaryButtonText = "OK",
            CloseButtonText = "キャンセル",
        };
        await ShowDialogAsync(dialog);
    }

    private async Task ShowRecipientsAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "宛先の選択",
            Content = "ネイティブ簡易版では、外部の連絡先データベースとの接続は行いません。\n«氏名»、«会社名»、«住所» などの差し込みフィールドを本文に挿入できます。",
            PrimaryButtonText = "OK",
            CloseButtonText = "閉じる",
        };
        await ShowDialogAsync(dialog);
    }

    private async Task ShowFindReplaceAsync(bool replaceMode)
    {
        var findBox = new TextBox
        {
            Header = "検索する文字列",
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
        };
        var replaceBox = new TextBox
        {
            Header = "置換後の文字列",
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
            Visibility = replaceMode ? Visibility.Visible : Visibility.Collapsed,
        };
        var matchCase = new CheckBox { Content = "大文字と小文字を区別する" };
        var wholeWord = new CheckBox { Content = "単語単位で検索する" };
        var panel = new StackPanel { Spacing = 10, Width = 390 };
        panel.Children.Add(findBox);
        panel.Children.Add(replaceBox);
        panel.Children.Add(matchCase);
        panel.Children.Add(wholeWord);

        var dialog = new ContentDialog
        {
            Title = replaceMode ? "置換" : "検索",
            Content = panel,
            PrimaryButtonText = "次を検索",
            SecondaryButtonText = replaceMode ? "置換" : string.Empty,
            CloseButtonText = "閉じる",
            DefaultButton = ContentDialogButton.Primary,
        };

        dialog.Opened += (_, _) =>
        {
            findBox.Focus(FocusState.Programmatic);
            findBox.SelectAll();
        };

        while (true)
        {
            var result = await ShowDialogAsync(dialog);
            if (result == ContentDialogResult.None)
            {
                break;
            }

            var query = findBox.Text;
            if (string.IsNullOrWhiteSpace(query))
            {
                Notify("検索する文字列を入力してください。", InfoBarSeverity.Warning);
                continue;
            }

            if (result == ContentDialogResult.Secondary)
            {
                ReplaceNext(query, replaceBox.Text, matchCase.IsChecked == true, wholeWord.IsChecked == true);
            }
            else if (!FindNext(query, matchCase.IsChecked == true, wholeWord.IsChecked == true))
            {
                Notify($"「{query}」は見つかりませんでした。", InfoBarSeverity.Warning);
            }
        }

        Editor.Focus(FocusState.Programmatic);
    }

    private bool FindNext(string query, bool matchCase, bool wholeWord)
    {
        if (string.IsNullOrEmpty(query))
        {
            return false;
        }

        var text = GetPlainText();
        if (text.Length == 0)
        {
            return false;
        }

        var selection = Editor.Document.Selection;
        var anchor = selection.EndPosition;
        if (anchor < 0 || anchor > text.Length)
        {
            anchor = text.Length;
        }

        var options = FindOptions.None;
        if (matchCase)
        {
            options |= FindOptions.Case;
        }
        if (wholeWord)
        {
            options |= FindOptions.Word;
        }

        for (var pass = 0; pass < 2; pass++)
        {
            var start = pass == 0 ? anchor : 0;
            var end = pass == 0 ? text.Length : anchor;
            if (end <= start)
            {
                continue;
            }

            var range = Editor.Document.GetRange(start, end);
            var found = range.FindText(query, end - start, options);
            if (found > 0)
            {
                selection.SetRange(range.StartPosition, range.StartPosition + found);
                selection.ScrollIntoView(PointOptions.Start);
                Editor.Focus(FocusState.Programmatic);
                return true;
            }
        }

        return false;
    }

    private void ReplaceNext(string query, string replacement, bool matchCase, bool wholeWord)
    {
        if (!FindNext(query, matchCase, wholeWord))
        {
            Notify($"「{query}」は見つかりませんでした。", InfoBarSeverity.Warning);
            return;
        }

        var selection = Editor.Document.Selection;
        selection.SetText(TextSetOptions.None, replacement);
        selection.Collapse(false);
        MarkModified();
        UpdateFormattingUi();
        Editor.Focus(FocusState.Programmatic);
    }

    private sealed record DocumentStatistics(
        int Characters,
        int CharactersWithSpaces,
        int Words,
        int Paragraphs,
        int Lines,
        int Sentences,
        int Pages);

    private static DocumentStatistics GetDocumentStatistics(string text)
    {
        var characters = 0;
        var charactersWithSpaces = 0;
        foreach (var character in text)
        {
            if (character is '\r' or '\n' || character == '\u0007')
            {
                continue;
            }

            charactersWithSpaces++;
            if (!char.IsWhiteSpace(character))
            {
                characters++;
            }
        }

        var words = Regex.Matches(text, @"[\p{L}\p{N}]+", RegexOptions.CultureInvariant).Count;
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Length == 0 ? 0 : normalized.Split('\n').Length;
        var paragraphs = normalized
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Count(line => line.Trim().Length > 0);
        var sentences = Regex
            .Split(normalized, @"(?<=[。．！？!?])")
            .Count(part => part.Trim().Length > 0);
        var pages = Math.Max(1, (int)Math.Ceiling(Math.Max(1, lines) / 42d));

        return new DocumentStatistics(characters, charactersWithSpaces, words, paragraphs, lines, sentences, pages);
    }

    private async Task ShowWordCountAsync()
    {
        var stats = GetDocumentStatistics(GetPlainText());

        var panel = new StackPanel { Spacing = 14, Width = 390 };
        panel.Children.Add(new TextBlock
        {
            Text = "文書の統計",
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 13,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        });

        var grid = new Grid
        {
            ColumnSpacing = 24,
            RowSpacing = 8,
        };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var rows = new (string Label, int Value)[]
        {
            ("ページ数", stats.Pages),
            ("単語数", stats.Words),
            ("文字数 (スペースなし)", stats.Characters),
            ("文字数 (スペースあり)", stats.CharactersWithSpaces),
            ("段落数", stats.Paragraphs),
            ("行数", stats.Lines),
        };

        for (var index = 0; index < rows.Length; index++)
        {
            grid.RowDefinitions.Add(new RowDefinition());
            AddStatRow(grid, index, rows[index].Label, rows[index].Value.ToString("N0"));
        }

        panel.Children.Add(grid);
        panel.Children.Add(new Border
        {
            Height = 1,
            Background = new SolidColorBrush(Color.FromArgb(255, 225, 223, 221)),
        });
        panel.Children.Add(new TextBlock
        {
            Text = "読みやすさの評価",
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 13,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        });

        var readability = new Grid
        {
            ColumnSpacing = 24,
            RowSpacing = 8,
        };
        readability.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        readability.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        readability.RowDefinitions.Add(new RowDefinition());
        readability.RowDefinitions.Add(new RowDefinition());

        var averageSentence = stats.Sentences == 0
            ? 0
            : (int)Math.Round(stats.Characters / (double)stats.Sentences);
        AddStatRow(readability, 0, "文の数", stats.Sentences.ToString("N0"));
        AddStatRow(readability, 1, "1 文あたりの平均文字数", averageSentence.ToString("N0"));
        panel.Children.Add(readability);

        var dialog = new ContentDialog
        {
            Title = "文字カウント",
            Content = panel,
            CloseButtonText = "閉じる",
        };
        await ShowDialogAsync(dialog);
    }

    private void ShowProofingPane()
    {
        ProofingPane.Visibility = Visibility.Visible;
        RefreshProofingPane();
    }

    private void CloseProofing_Click(object sender, RoutedEventArgs e)
    {
        ProofingPane.Visibility = Visibility.Collapsed;
        Editor.Focus(FocusState.Programmatic);
    }

    private void ProofingSpellCheck_Changed(object sender, RoutedEventArgs e)
    {
        if (_updatingProofingPane)
        {
            return;
        }

        SetSpellCheckEnabled(ProofingSpellCheck.IsChecked == true);
    }

    private void RefreshProofingPane()
    {
        if (ProofingPane.Visibility != Visibility.Visible)
        {
            return;
        }

        _updatingProofingPane = true;
        try
        {
            ProofingSpellCheck.IsChecked = _spellCheckEnabled;
        }
        finally
        {
            _updatingProofingPane = false;
        }

        var text = GetPlainText();
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var issues = new List<string>();
        var tips = new List<string>();

        var sentences = Regex
            .Split(normalized, @"(?<=[。．！？!?])")
            .Select(part => part.Trim())
            .Where(part => part.Length > 0)
            .ToList();
        var longSentences = sentences.Count(sentence => sentence.Length > 60);
        if (longSentences > 0)
        {
            issues.Add($"長い文が {longSentences} 件あります。1 文を 60 文字以内にすると読みやすくなります。");
        }

        var doubleSpaces = Regex.Matches(normalized, @"[ \u3000]{2,}").Count;
        if (doubleSpaces > 0)
        {
            issues.Add($"余分な空白が {doubleSpaces} か所あります。");
        }

        var repeatedPunctuation = Regex.Matches(normalized, @"([、。，．！？])\1+").Count;
        if (repeatedPunctuation > 0)
        {
            issues.Add($"句読点の重複が {repeatedPunctuation} か所あります。");
        }

        var halfWidthKana = Regex.Matches(normalized, @"[\uFF61-\uFF9F]").Count;
        if (halfWidthKana > 0)
        {
            issues.Add($"半角カナが {halfWidthKana} 文字あります。全角カナを推奨します。");
        }

        if (issues.Count == 0)
        {
            issues.Add("大きな問題は見つかりませんでした。");
        }

        tips.Add("一文を短くし、同じ語尾の連続を避けると、より自然な文章になります。");
        tips.Add("日本語と英数字の間には半角スペースを入れると読みやすくなります。");
        if (sentences.Count > 0)
        {
            var average = (int)Math.Round(sentences.Average(sentence => sentence.Length));
            tips.Add($"現在の 1 文あたりの平均文字数は約 {average} 文字です。");
        }

        tips.Add(_spellCheckEnabled
            ? "スペルチェックはオンです。赤い波線の付いた語を右クリックして候補を選べます。"
            : "スペルチェックはオフです。下のチェックボックスで有効にできます。");

        var issueCount = issues.Count(issue => !issue.StartsWith("大きな問題", StringComparison.Ordinal));
        var score = Math.Clamp(100 - (issueCount * 6) - (longSentences > 3 ? 4 : 0), 40, 100);
        ProofingScoreText.Text = score.ToString(CultureInfo.CurrentCulture);
        ProofingIssueCountText.Text = $"{issueCount} 件";
        ProofingSummaryText.Text = issueCount == 0
            ? "問題は見つかりませんでした。"
            : $"{issueCount} 件の項目を確認してください。";
        ProofingLanguageText.Text = "校正言語: 日本語 (日本)";

        ProofingIssuesPanel.Children.Clear();
        var isDark = ProofingPane.ActualTheme == ElementTheme.Dark;
        var accentBrush = new SolidColorBrush(isDark
            ? Color.FromArgb(255, 79, 163, 227)
            : Color.FromArgb(255, 15, 108, 189));
        var errorBrush = new SolidColorBrush(isDark
            ? Color.FromArgb(255, 241, 112, 122)
            : Color.FromArgb(255, 196, 43, 28));
        foreach (var issue in issues)
        {
            var isNote = issue.StartsWith("大きな問題", StringComparison.Ordinal);
            ProofingIssuesPanel.Children.Add(CreateProofingLine(
                issue,
                isNote ? "\uE930" : "\uE7BA",
                isNote ? accentBrush : errorBrush));
        }

        ProofingTipsPanel.Children.Clear();
        var tipBrush = new SolidColorBrush(isDark
            ? Color.FromArgb(255, 200, 198, 196)
            : Color.FromArgb(255, 96, 94, 92));
        foreach (var tip in tips)
        {
            ProofingTipsPanel.Children.Add(CreateProofingLine(
                tip,
                "\uE946",
                tipBrush));
        }
    }

    private static Grid CreateProofingLine(string text, string glyph, Brush brush)
    {
        var grid = new Grid { ColumnSpacing = 7 };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        var icon = new FontIcon
        {
            Glyph = glyph,
            FontSize = 11,
            Foreground = brush,
            VerticalAlignment = VerticalAlignment.Top,
            Margin = new Thickness(0, 2, 0, 0),
        };
        var label = new TextBlock
        {
            Text = text,
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 11.5,
            TextWrapping = TextWrapping.Wrap,
            Foreground = brush,
        };
        Grid.SetColumn(label, 1);
        grid.Children.Add(icon);
        grid.Children.Add(label);
        return grid;
    }

    private static void AddStatRow(Grid grid, int row, string label, string value)
    {
        var labelBlock = new TextBlock
        {
            Text = label,
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 12,
            Foreground = new SolidColorBrush(Color.FromArgb(255, 96, 94, 92)),
            VerticalAlignment = VerticalAlignment.Center,
        };
        var valueBlock = new TextBlock
        {
            Text = value,
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 13,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center,
        };
        Grid.SetRow(labelBlock, row);
        Grid.SetRow(valueBlock, row);
        Grid.SetColumn(valueBlock, 1);
        grid.Children.Add(labelBlock);
        grid.Children.Add(valueBlock);
    }

    private async Task ShowLanguageAsync()
    {
        var combo = new ComboBox
        {
            Width = 300,
            SelectedIndex = 0,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        combo.Items.Add("日本語 (日本)");
        combo.Items.Add("English (United States)");
        var panel = new StackPanel { Spacing = 8, Width = 330 };
        panel.Children.Add(new TextBlock
        {
            Text = "選択範囲の校正言語を指定します。",
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 12,
            Foreground = new SolidColorBrush(Color.FromArgb(255, 96, 94, 92)),
        });
        panel.Children.Add(combo);

        var dialog = new ContentDialog
        {
            Title = "言語",
            Content = panel,
            PrimaryButtonText = "OK",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        if (await ShowDialogAsync(dialog) != ContentDialogResult.Primary)
        {
            return;
        }

        try
        {
            Editor.Document.Selection.CharacterFormat.LanguageTag = combo.SelectedIndex == 0 ? "ja-JP" : "en-US";
            MarkModified();
            Notify("校正言語を変更しました。");
        }
        catch (Exception ex)
        {
            Notify($"言語を変更できませんでした: {ex.Message}", InfoBarSeverity.Warning);
        }
    }

    private async Task ShowOptionsAsync()
    {
        var panel = new StackPanel { Spacing = 10, Width = 420 };

        AddOptionSection(panel, "外観");

        var themeCombo = new ComboBox
        {
            Header = "テーマ",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
            SelectedIndex = _themePreference switch
            {
                "light" => 1,
                "dark" => 2,
                _ => 0,
            },
        };
        themeCombo.Items.Add("システムの設定に合わせる");
        themeCombo.Items.Add("ライト");
        themeCombo.Items.Add("ダーク");
        panel.Children.Add(themeCombo);

        var ribbonCombo = new ComboBox
        {
            Header = "リボンの表示",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
            SelectedIndex = _useCompactRibbon ? 0 : 1,
        };
        ribbonCombo.Items.Add("簡易リボン (1 行)");
        ribbonCombo.Items.Add("従来のリボン (グループ表示)");
        panel.Children.Add(ribbonCombo);

        AddOptionSection(panel, "編集");

        var spell = new CheckBox
        {
            Content = "スペルチェックを有効にする",
            IsChecked = _spellCheckEnabled,
        };
        panel.Children.Add(spell);

        var fontCombo = new ComboBox
        {
            Header = "新規文書の既定フォント",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        foreach (var (label, name) in new (string Label, string Name)[]
        {
            ("游明朝", "游明朝"),
            ("游ゴシック", "游ゴシック"),
            ("MS ゴシック", "MS ゴシック"),
            ("MS 明朝", "MS 明朝"),
            ("メイリオ", "メイリオ"),
            ("BIZ UDPゴシック", "BIZ UDPゴシック"),
            ("Arial", "Arial"),
            ("Calibri", "Calibri"),
            ("Times New Roman", "Times New Roman"),
        })
        {
            fontCombo.Items.Add(new ComboBoxItem { Content = label, Tag = name });
        }

        foreach (var item in fontCombo.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag as string, _defaultFontName, StringComparison.OrdinalIgnoreCase))
            {
                fontCombo.SelectedItem = item;
                break;
            }
        }

        if (fontCombo.SelectedItem is null)
        {
            fontCombo.SelectedIndex = 0;
        }

        panel.Children.Add(fontCombo);

        var sizeCombo = new ComboBox
        {
            Header = "新規文書の既定フォント サイズ",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            IsEditable = true,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        foreach (var size in new[] { "8", "9", "10.5", "11", "12", "14", "16", "18", "20", "24", "28", "36", "48", "72" })
        {
            sizeCombo.Items.Add(size);
        }

        // The editable combo drops text assigned before its template is applied, so
        // select the matching item and re-apply the text once the control is loaded.
        var defaultSizeText = FormatSize((float)_defaultFontSize);
        var matchingSize = sizeCombo.Items.OfType<string>()
            .FirstOrDefault(size => size == defaultSizeText);
        if (matchingSize is not null)
        {
            sizeCombo.SelectedItem = matchingSize;
        }

        sizeCombo.Text = defaultSizeText;
        sizeCombo.Loaded += (_, _) => sizeCombo.Text = defaultSizeText;
        panel.Children.Add(sizeCombo);

        AddOptionSection(panel, "自動保存");
        var autoSave = new CheckBox
        {
            Content = "変更を自動保存する",
            IsChecked = _autoSaveEnabled,
        };
        panel.Children.Add(autoSave);

        var intervalCombo = new ComboBox
        {
            Header = "自動保存の間隔",
            HorizontalAlignment = HorizontalAlignment.Stretch,
            FontFamily = new FontFamily("Yu Gothic UI"),
        };
        foreach (var (label, seconds) in new (string Label, int Seconds)[]
        {
            ("5 秒", 5),
            ("15 秒", 15),
            ("30 秒", 30),
            ("1 分", 60),
            ("5 分", 300),
        })
        {
            intervalCombo.Items.Add(new ComboBoxItem { Content = label, Tag = seconds });
        }

        foreach (var item in intervalCombo.Items.OfType<ComboBoxItem>())
        {
            if (item.Tag is int intervalSeconds && intervalSeconds == _autoSaveIntervalSeconds)
            {
                intervalCombo.SelectedItem = item;
                break;
            }
        }

        intervalCombo.SelectedItem ??= intervalCombo.Items[2];
        panel.Children.Add(intervalCombo);

        AddOptionSection(panel, "通知");
        var toastToggle = new CheckBox
        {
            Content = "状態メッセージをポップアップでも表示する",
            IsChecked = _showToasts,
        };
        panel.Children.Add(toastToggle);
        panel.Children.Add(new TextBlock
        {
            Text = "オフの場合、操作の結果はステータス バーに表示されます。エラーは常にポップアップで通知します。",
            TextWrapping = TextWrapping.Wrap,
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 11.5,
            Foreground = new SolidColorBrush(Color.FromArgb(255, 96, 94, 92)),
        });

        var scroll = new ScrollViewer
        {
            Content = panel,
            MaxHeight = 430,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };

        var dialog = new ContentDialog
        {
            Title = "WordPad Lite のオプション",
            Content = scroll,
            PrimaryButtonText = "OK",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        if (await ShowDialogAsync(dialog) != ContentDialogResult.Primary)
        {
            return;
        }

        var theme = themeCombo.SelectedIndex switch
        {
            1 => "light",
            2 => "dark",
            _ => "system",
        };
        if (theme != _themePreference)
        {
            SetThemePreference(theme, notify: false);
        }

        SetRibbonMode(ribbonCombo.SelectedIndex == 0);

        SetSpellCheckEnabled(spell.IsChecked == true);

        if (fontCombo.SelectedItem is ComboBoxItem fontItem && fontItem.Tag is string fontName)
        {
            _defaultFontName = fontName;
        }

        var sizeText = sizeCombo.Text.Replace("pt", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        if (double.TryParse(sizeText, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsedSize))
        {
            _defaultFontSize = Math.Clamp(parsedSize, 6, 200);
        }

        AppSettings.Set("DefaultFont", _defaultFontName);
        AppSettings.Set("DefaultFontSize", _defaultFontSize);

        if (intervalCombo.SelectedItem is ComboBoxItem savedIntervalItem && savedIntervalItem.Tag is int savedSeconds)
        {
            _autoSaveIntervalSeconds = Math.Max(5, savedSeconds);
            AppSettings.Set("AutoSaveInterval", _autoSaveIntervalSeconds);
        }

        SetAutoSaveEnabled(autoSave.IsChecked == true);

        _showToasts = toastToggle.IsChecked == true;
        AppSettings.Set("ShowToasts", _showToasts);

        Notify("オプションを更新しました。");
    }

    private static void AddOptionSection(StackPanel panel, string title)
    {
        if (panel.Children.Count > 0)
        {
            panel.Children.Add(new Border
            {
                Height = 1,
                Background = new SolidColorBrush(Color.FromArgb(255, 225, 223, 221)),
                Margin = new Thickness(0, 6, 0, 2),
            });
        }

        panel.Children.Add(new TextBlock
        {
            Text = title,
            FontFamily = new FontFamily("Yu Gothic UI"),
            FontSize = 12.5,
            FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        });
    }

    private async Task ShowHelpAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "WordPad Lite のヘルプ",
            Content = new TextBlock
            {
                Text = "WordPad Lite は Word に似たリボンで文書を作成できるネイティブ エディターです。\n\n" +
                       "・ホームタブで文字や段落の書式を設定できます。\n" +
                       "・挿入タブから画像、表、リンク、日付、コメントを挿入できます。\n" +
                       "・レイアウトタブで用紙サイズ、向き、余白を変更できます。\n" +
                       "・Ctrl+S で保存、Ctrl+O で開く、Ctrl+F で検索できます。\n" +
                       "・RTF は書式を保持して保存され、DOCX と TXT はテキストを入出力します。\n\n" +
                       "ファイルタブから保存形式を選べます。",
                TextWrapping = TextWrapping.Wrap,
                FontFamily = new FontFamily("Yu Gothic UI"),
                FontSize = 12.5,
                LineHeight = 22,
                MaxWidth = 480,
            },
            CloseButtonText = "閉じる",
        };
        await ShowDialogAsync(dialog);
    }

    private async Task ShowShortcutsAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "キーボード ショートカット",
            Content = new TextBlock
            {
                Text = "Ctrl+N    新規作成\n" +
                       "Ctrl+O    開く\n" +
                       "Ctrl+S    保存\n" +
                       "Ctrl+Shift+S    名前を付けて保存\n" +
                       "Ctrl+P    印刷\n" +
                       "Ctrl+Z    元に戻す\n" +
                       "Ctrl+Y    やり直す\n" +
                       "Ctrl+B    太字\n" +
                       "Ctrl+I    斜体\n" +
                       "Ctrl+U    下線\n" +
                       "Ctrl+F    検索\n" +
                       "Ctrl+H    置換\n" +
                       "Ctrl+A    すべて選択\n" +
                       "Ctrl+0    ズームを 100% に戻す",
                FontFamily = new FontFamily("Yu Gothic UI"),
                FontSize = 12.5,
                LineHeight = 22,
                TextWrapping = TextWrapping.Wrap,
            },
            CloseButtonText = "閉じる",
        };
        await ShowDialogAsync(dialog);
    }

    private async Task ShowFeedbackAsync()
    {
        var dialog = new ContentDialog
        {
            Title = "フィードバック",
            Content = "WordPad Lite へのご意見ありがとうございます。\n\n" +
                      "ネイティブ版は RTF を中心とした軽量エディターです。表計算や差し込み印刷の高度な機能は、Web 版または Microsoft Word をご利用ください。",
            PrimaryButtonText = "OK",
            CloseButtonText = "閉じる",
        };
        await ShowDialogAsync(dialog);
    }

    private async Task<string?> PromptTextAsync(string title, string label, string initialValue)
    {
        var box = new TextBox
        {
            Header = label,
            Text = initialValue,
            Style = (Style)Application.Current.Resources["DialogTextBoxStyle"],
            Width = 360,
        };
        var dialog = new ContentDialog
        {
            Title = title,
            Content = box,
            PrimaryButtonText = "OK",
            CloseButtonText = "キャンセル",
            DefaultButton = ContentDialogButton.Primary,
        };

        return await ShowDialogAsync(dialog) == ContentDialogResult.Primary ? box.Text : null;
    }

    private async Task<ContentDialogResult> ShowDialogAsync(ContentDialog dialog)
    {
        dialog.XamlRoot = RootGrid.XamlRoot;
        var primaryText = dialog.PrimaryButtonText;
        if (!string.IsNullOrEmpty(primaryText))
        {
            var accentStyle = (Style)Application.Current.Resources["WordAccentButtonStyle"];
            dialog.PrimaryButtonStyle = accentStyle;
            dialog.Opened += (_, _) =>
            {
                // ContentDialog's "PrimaryAsDefaultButton" visual state re-applies the
                // system AccentButtonStyle, so override the live button after that runs.
                var timer = DispatcherQueue.CreateTimer();
                timer.Interval = TimeSpan.FromMilliseconds(180);
                timer.IsRepeating = false;
                timer.Tick += (_, _) =>
                {
                    foreach (var button in EnumerateButtons(dialog))
                    {
                        if (button.Content is string text && text == primaryText)
                        {
                            button.Style = accentStyle;
                            break;
                        }
                    }
                };
                timer.Start();
            };
        }

        return await dialog.ShowAsync();
    }

    private static IEnumerable<Button> EnumerateButtons(DependencyObject root)
    {
        var count = VisualTreeHelper.GetChildrenCount(root);
        for (var index = 0; index < count; index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is Button button)
            {
                yield return button;
            }

            foreach (var nested in EnumerateButtons(child))
            {
                yield return nested;
            }
        }
    }

    private void ShowNotAvailable(string feature)
    {
        Notify($"{feature} はネイティブ簡易版では準備中です。Web 版の WordPad Lite では利用できます。");
    }

    private void Notify(string message, InfoBarSeverity severity = InfoBarSeverity.Informational)
    {
        // Results always land in the status bar. Popups are opt-in, except for
        // errors, which must stay prominent.
        StatusMessageText.Text = message;
        StatusMessageText.Visibility = Visibility.Visible;
        _statusMessageTimer ??= DispatcherQueue.CreateTimer();
        _statusMessageTimer.Tick -= StatusMessageTimer_Tick;
        _statusMessageTimer.Tick += StatusMessageTimer_Tick;
        _statusMessageTimer.Interval = TimeSpan.FromSeconds(severity == InfoBarSeverity.Error ? 10 : 7);
        _statusMessageTimer.IsRepeating = false;
        _statusMessageTimer.Stop();
        _statusMessageTimer.Start();

        if (!_showToasts && severity != InfoBarSeverity.Error)
        {
            return;
        }

        ToastBar.Title = "WordPad Lite";
        ToastBar.Message = message;
        ToastBar.Severity = severity;
        ToastBar.IsOpen = true;
        PlayEntrance(ToastBar, -8, 140);

        _toastTimer ??= DispatcherQueue.CreateTimer();
        _toastTimer.Tick -= ToastTimer_Tick;
        _toastTimer.Tick += ToastTimer_Tick;
        _toastTimer.Interval = TimeSpan.FromSeconds(severity == InfoBarSeverity.Error ? 7 : 5);
        _toastTimer.IsRepeating = false;
        _toastTimer.Stop();
        _toastTimer.Start();
    }

    private void StatusMessageTimer_Tick(Microsoft.UI.Dispatching.DispatcherQueueTimer sender, object args)
    {
        StatusMessageText.Visibility = Visibility.Collapsed;
    }

    private void ToastTimer_Tick(Microsoft.UI.Dispatching.DispatcherQueueTimer sender, object args)
    {
        ToastBar.IsOpen = false;
    }

    private static void PlayEntrance(FrameworkElement? element, double offsetY = 6, int durationMs = 160)
    {
        if (element is null)
        {
            return;
        }

        var transform = new TranslateTransform { Y = offsetY };
        element.RenderTransform = transform;
        element.Opacity = 0;

        var duration = new Duration(TimeSpan.FromMilliseconds(durationMs));
        var ease = new Microsoft.UI.Xaml.Media.Animation.CubicEase
        {
            EasingMode = Microsoft.UI.Xaml.Media.Animation.EasingMode.EaseOut,
        };

        var fade = new Microsoft.UI.Xaml.Media.Animation.DoubleAnimation
        {
            To = 1,
            Duration = duration,
            EasingFunction = ease,
        };
        Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTarget(fade, element);
        Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTargetProperty(fade, "Opacity");

        var slide = new Microsoft.UI.Xaml.Media.Animation.DoubleAnimation
        {
            To = 0,
            Duration = duration,
            EasingFunction = ease,
        };
        Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTarget(slide, transform);
        Microsoft.UI.Xaml.Media.Animation.Storyboard.SetTargetProperty(slide, "Y");

        var storyboard = new Microsoft.UI.Xaml.Media.Animation.Storyboard();
        storyboard.Children.Add(fade);
        storyboard.Children.Add(slide);
        storyboard.Begin();
    }

    private void HookFlyouts(DependencyObject root)
    {
        var count = VisualTreeHelper.GetChildrenCount(root);
        for (var index = 0; index < count; index++)
        {
            var child = VisualTreeHelper.GetChild(root, index);
            if (child is Button { Flyout: not null } button)
            {
                TrackFlyout(button.Flyout);
            }

            HookFlyouts(child);
        }
    }

    private void TrackFlyout(FlyoutBase flyout)
    {
        flyout.Opened -= Flyout_Opened;
        flyout.Opened += Flyout_Opened;
        flyout.Closed -= Flyout_Closed;
        flyout.Closed += Flyout_Closed;
    }

    private void Flyout_Opened(object? sender, object e)
    {
        _activeFlyout = sender as FlyoutBase;
    }

    private void Flyout_Closed(object? sender, object e)
    {
        if (ReferenceEquals(_activeFlyout, sender))
        {
            _activeFlyout = null;
        }
    }

    private static string RtfEscape(string value)
    {
        var builder = new StringBuilder(value.Length + 16);
        foreach (var character in value)
        {
            switch (character)
            {
                case '\\':
                    builder.Append("\\\\");
                    break;
                case '{':
                    builder.Append("\\{");
                    break;
                case '}':
                    builder.Append("\\}");
                    break;
                case '\r':
                case '\n':
                    builder.Append("\\par ");
                    break;
                case '\t':
                    builder.Append("\\tab ");
                    break;
                default:
                    if (character < 128)
                    {
                        builder.Append(character);
                    }
                    else
                    {
                        var code = (int)character;
                        if (code > 32767)
                        {
                            code -= 65536;
                        }
                        builder.Append("\\u").Append(code).Append('?');
                    }
                    break;
            }
        }
        return builder.ToString();
    }

}
