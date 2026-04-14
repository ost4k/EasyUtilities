using System.Runtime.InteropServices;
using System.Text.Json;
using System.Diagnostics;
using EasyUtilities.Services;
using EasyUtilities.TrayHost;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinForms = System.Windows.Forms;
using Windows.ApplicationModel.DataTransfer;
using Windows.Graphics;
using WinRT.Interop;

namespace EasyUtilities;

public sealed partial class MainWindow : Window
{
    private const double TwoColumnCardsMinWidth = 1280;
    private const int HotkeyTopMostId = 1;
    private const int HotkeyOcrId = 2;
    private const int HotkeyShowHideAppId = 3;
    private const int HotkeyRestartExplorerId = 4;
    private const int HotkeyFileListId = 5;
    private const string SettingsBackupType = "easy-utilities.settings";
    private const int SettingsBackupVersion = 1;

    private readonly SettingsService _settingsService = new();
    private readonly StartupService _startupService = new();
    private readonly ShellTweaksService _shellTweaksService = new();
    private readonly DesktopShortcutNameService _shortcutNameService = new();
    private readonly DesktopIconLayoutService _desktopIconLayoutService = new();
    private readonly OcrService _ocrService = new();
    private readonly ExplorerFileListExportService _explorerFileListExportService = new();

    private AppSettings _settings = new();
    private AppWindow? _appWindow;
    private nint _hwnd;
    private nint _oldWndProc;
    private WndProcDelegate? _wndProcDelegate;
    private TrayIconHost? _trayIcon;
    private bool _allowClose;
    private bool _isLoading;

    private nint _mouseHookHandle;
    private NativeMethods.LowLevelMouseProc? _mouseHookProc;
    private CancellationTokenSource? _statusCts;
    private bool _hotkeyOcrRegistered;
    private readonly Task _initializeTask;

    private delegate nint WndProcDelegate(nint hWnd, uint msg, nint wParam, nint lParam);

    public MainWindow()
    {
        InitializeComponent();
        Title = "EasyUtilities";
        ExtendsContentIntoTitleBar = false;

        _initializeTask = InitializeAsync();
    }

    public Task EnsureInitializedAsync() => _initializeTask;

    public async Task HandleStartupShortcutArrowActionAsync(bool hideArrows)
    {
        await EnsureInitializedAsync();

        try
        {
            _shellTweaksService.SetShortcutArrowsHidden(hideArrows);
            _settings.HideShortcutArrows = hideArrows;
            await _settingsService.SaveAsync(_settings);

            _isLoading = true;
            try
            {
                HideShortcutArrowsToggle.IsOn = hideArrows;
            }
            finally
            {
                _isLoading = false;
            }

            SetStatus(Localize("arrows.updated"), InfoBarSeverity.Success);

            var restartNow = await ConfirmAsync(
                Localize("explorer.title"),
                Localize("startup.arrowsAppliedRestart"),
                Localize("explorer.restart"));

            if (restartNow)
            {
                _shellTweaksService.RestartExplorer();
                SetStatus(Localize("explorer.restarted"), InfoBarSeverity.Success);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("arrows.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    public async Task HandleStartupDesktopShortcutArrowActionAsync(bool hideArrows)
    {
        await EnsureInitializedAsync();

        try
        {
            _shellTweaksService.SetDesktopShortcutArrowsHidden(hideArrows);
            _settings.HideDesktopShortcutArrows = hideArrows;
            await _settingsService.SaveAsync(_settings);

            _isLoading = true;
            try
            {
                HideDesktopShortcutArrowsToggle.IsOn = hideArrows;
            }
            finally
            {
                _isLoading = false;
            }

            if (hideArrows && _shellTweaksService.IsDesktopShortcutOverlayPotentiallyUnsupported())
            {
                SetStatus(Localize("desktopArrows.compat.warning"), InfoBarSeverity.Warning);
            }
            else
            {
                SetStatus(Localize("arrows.updated"), InfoBarSeverity.Success);
            }
            await ShowDesktopArrowsRestartDialogAndRestartAsync(hideArrows);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("arrows.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    public async Task HandleStartupLegacyShortcutArrowCleanupAsync(bool showRestartPrompt = true)
    {
        await EnsureInitializedAsync();

        try
        {
            if (!_shellTweaksService.IsAdministrator())
            {
                SetStatus(Localize("legacyArrows.migration.pending"), InfoBarSeverity.Warning);
                return;
            }

            _shellTweaksService.CleanupLegacyMachineShortcutArrowState();
            SetStatus(Localize("arrows.updated"), InfoBarSeverity.Success);

            if (showRestartPrompt)
            {
                var restartNow = await ConfirmAsync(
                    Localize("explorer.title"),
                    Localize("startup.arrowsAppliedRestart"),
                    Localize("explorer.restart"));

                if (restartNow)
                {
                    _shellTweaksService.RestartExplorer();
                    SetStatus(Localize("explorer.restarted"), InfoBarSeverity.Success);
                }
            }
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("arrows.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    public async Task HandleStartupLegacyDesktopShortcutArrowCleanupAsync(bool showRestartPrompt = true)
    {
        await EnsureInitializedAsync();

        try
        {
            if (!_shellTweaksService.IsAdministrator())
            {
                SetStatus(Localize("legacyArrows.migration.pending"), InfoBarSeverity.Warning);
                return;
            }

            _shellTweaksService.CleanupLegacyMachineDesktopArrowState();
            SetStatus(Localize("arrows.updated"), InfoBarSeverity.Success);
            if (showRestartPrompt)
            {
                await ShowDesktopArrowsRestartDialogAndRestartAsync(false);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("arrows.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    public void MinimizeToTray()
    {
        if (_hwnd == nint.Zero)
        {
            _hwnd = WindowNative.GetWindowHandle(this);
        }

        NativeMethods.ShowWindow(_hwnd, NativeMethods.SwHide);
        SetStatus(Localize("app.minimized"));
    }

    private async Task InitializeAsync()
    {
        try
        {
            _settings = await _settingsService.LoadAsync();
            try
            {
                _shellTweaksService.RepairUserShortcutClassAssociations();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to repair shortcut class associations: {ex.Message}");
            }

            _hwnd = WindowNative.GetWindowHandle(this);
            ConfigureWindow();
            InitializeTrayIcon();
            InitializeWindowHook();
            RegisterHotkeys();

            _isLoading = true;
            StartupToggle.IsOn = _settings.StartWithWindows || await _startupService.IsEnabledAsync();
            MiddleClickToggle.IsOn = _settings.MinimizeOnMiddleClickTitle;
            HideShortcutArrowsToggle.IsOn = _settings.HideShortcutArrows;
            HideDesktopShortcutArrowsToggle.IsOn = _settings.HideDesktopShortcutArrows;
            HideShortcutNamesToggle.IsOn = _settings.HideShortcutNames;
            TopMostToggle.IsOn = _settings.AlwaysOnTopEnabled;
            OcrToggle.IsOn = _settings.OcrEnabled;
            RestartExplorerToggle.IsOn = _settings.RestartExplorerHotkeyEnabled;
            FileListToggle.IsOn = _settings.FileListExportEnabled;
            SelectComboByTag(FileListModeCombo, NormalizeFileListExportMode(_settings.FileListExportMode));
            FileListModeCombo.IsEnabled = _settings.FileListExportEnabled;
            SelectComboByTag(LanguageCombo, _settings.LanguageMode);
            SelectComboByTag(ThemeCombo, _settings.ThemeMode);
            LocalizationService.SetLanguageMode(_settings.LanguageMode);
            LocalizationService.Reload();
            ApplyTheme(_settings.ThemeMode);
            ApplyLocalization();

            try
            {
                if (_settings.ShortcutRenameMaps.Count > 0)
                {
                    _shortcutNameService.RestoreDesktopShortcutLabels(_settings.ShortcutRenameMaps);
                    if (!_settings.HideShortcutNames)
                    {
                        _settings.ShortcutRenameMaps.Clear();
                    }
                }

                if (_settings.HideShortcutNames)
                {
                    await ApplyShortcutLabelHiddenStateAsync(true);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"{Localize("labels.error")}: {ex.Message}", InfoBarSeverity.Warning);
            }

            EnableGlobalMiddleClickHook(_settings.MinimizeOnMiddleClickTitle);
        }
        finally
        {
            _isLoading = false;
        }

        if (!_hotkeyOcrRegistered)
        {
            SetStatus(Localize("ocr.hotkey.unavailable"), InfoBarSeverity.Warning);
        }
    }

    private void ConfigureWindow()
    {
        _appWindow = AppWindow.GetFromWindowId(Microsoft.UI.Win32Interop.GetWindowIdFromWindow(_hwnd));
        var appIconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico");
        if (File.Exists(appIconPath))
        {
            _appWindow.SetIcon(appIconPath);
        }

        _appWindow.Resize(new SizeInt32(1060, 760));
        _appWindow.Closing += AppWindow_Closing;
    }

    private void CardsGrid_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplyCardsLayout(e.NewSize.Width >= TwoColumnCardsMinWidth);
    }

    private void ApplyCardsLayout(bool twoColumns)
    {
        if (!twoColumns)
        {
            SetCardPosition(StartupCard, 2, 0, 2);
            SetCardPosition(MiddleClickCard, 3, 0, 2);
            SetCardPosition(IconsSaveCard, 4, 0, 2);
            SetCardPosition(IconsRestoreCard, 5, 0, 2);
            SetCardPosition(HideNamesCard, 6, 0, 2);
            SetCardPosition(HideArrowsCard, 7, 0, 2);
            SetCardPosition(HideDesktopArrowsCard, 8, 0, 2);
            SetCardPosition(TopMostCard, 9, 0, 2);
            SetCardPosition(OcrCard, 10, 0, 2);
            SetCardPosition(RestartExplorerCard, 11, 0, 2);
            SetCardPosition(FileListCard, 12, 0, 2);
            SetCardPosition(ThemeCard, 13, 0, 2);
            SetCardPosition(LanguageCard, 14, 0, 2);
            SetCardPosition(BackupCard, 15, 0, 2);
            SetCardPosition(RestoreCard, 16, 0, 2);
            return;
        }

        SetCardPosition(StartupCard, 2, 0, 1);
        SetCardPosition(MiddleClickCard, 2, 1, 1);
        SetCardPosition(IconsSaveCard, 3, 0, 1);
        SetCardPosition(IconsRestoreCard, 3, 1, 1);
        SetCardPosition(HideNamesCard, 4, 0, 1);
        SetCardPosition(HideArrowsCard, 4, 1, 1);
        SetCardPosition(HideDesktopArrowsCard, 5, 0, 1);
        SetCardPosition(TopMostCard, 5, 1, 1);
        SetCardPosition(OcrCard, 6, 0, 1);
        SetCardPosition(RestartExplorerCard, 6, 1, 1);
        SetCardPosition(FileListCard, 7, 0, 1);
        SetCardPosition(ThemeCard, 7, 1, 1);
        SetCardPosition(LanguageCard, 8, 0, 1);
        SetCardPosition(BackupCard, 8, 1, 1);
        SetCardPosition(RestoreCard, 9, 0, 2);
    }

    private static void SetCardPosition(FrameworkElement card, int row, int column, int columnSpan)
    {
        Grid.SetRow(card, row);
        Grid.SetColumn(card, column);
        Grid.SetColumnSpan(card, columnSpan);
    }

    private void AppWindow_Closing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_allowClose)
        {
            return;
        }

        args.Cancel = true;
        MinimizeToTray();
    }

    private void InitializeTrayIcon()
    {
        _trayIcon = new TrayIconHost(
            Localize("ui.app.title"),
            Localize("tray.open"),
            Localize("tray.exit"));
        _trayIcon.OpenRequested += (_, _) => _ = DispatcherQueue.TryEnqueue(ShowFromTray);
        _trayIcon.ExitRequested += (_, _) => _ = DispatcherQueue.TryEnqueue(ExitApplication);
    }

    private void InitializeWindowHook()
    {
        _wndProcDelegate = WndProc;
        var newProc = Marshal.GetFunctionPointerForDelegate(_wndProcDelegate);
        _oldWndProc = NativeMethods.SetWindowLongPtr(_hwnd, NativeMethods.GwlpWndProc, newProc);
    }

    private void RegisterHotkeys()
    {
        NativeMethods.RegisterHotKey(_hwnd, HotkeyTopMostId, NativeMethods.ModControl | NativeMethods.ModAlt, (uint)'T');
        _hotkeyOcrRegistered = NativeMethods.RegisterHotKey(_hwnd, HotkeyOcrId, NativeMethods.ModControl | NativeMethods.ModAlt, (uint)'O');
        NativeMethods.RegisterHotKey(_hwnd, HotkeyShowHideAppId, NativeMethods.ModControl | NativeMethods.ModAlt, (uint)'M');
        NativeMethods.RegisterHotKey(_hwnd, HotkeyRestartExplorerId, NativeMethods.ModControl | NativeMethods.ModAlt, (uint)'R');
        NativeMethods.RegisterHotKey(_hwnd, HotkeyFileListId, NativeMethods.ModControl | NativeMethods.ModAlt, (uint)'F');
    }

    private void UnregisterHotkeys()
    {
        NativeMethods.UnregisterHotKey(_hwnd, HotkeyTopMostId);
        NativeMethods.UnregisterHotKey(_hwnd, HotkeyOcrId);
        NativeMethods.UnregisterHotKey(_hwnd, HotkeyShowHideAppId);
        NativeMethods.UnregisterHotKey(_hwnd, HotkeyRestartExplorerId);
        NativeMethods.UnregisterHotKey(_hwnd, HotkeyFileListId);
    }

    private nint WndProc(nint hWnd, uint msg, nint wParam, nint lParam)
    {
        if (msg == NativeMethods.WmClose && !_allowClose)
        {
            MinimizeToTray();
            return nint.Zero;
        }

        if (msg == NativeMethods.WmSysCommand &&
            ((ulong)wParam.ToInt64() & 0xFFF0UL) == NativeMethods.ScClose &&
            !_allowClose)
        {
            MinimizeToTray();
            return nint.Zero;
        }

        if (msg == NativeMethods.WmNCMButtonDown && (int)wParam == NativeMethods.HtCaption && _settings.MinimizeOnMiddleClickTitle)
        {
            MinimizeToTray();
            return nint.Zero;
        }

        if (msg == NativeMethods.WmHotkey)
        {
            var id = wParam.ToInt32();
            _ = DispatcherQueue.TryEnqueue(async () =>
            {
                if (id == HotkeyTopMostId)
                {
                    if (_settings.AlwaysOnTopEnabled)
                    {
                        ToggleTopMostForWindow(NativeMethods.GetForegroundWindow());
                    }
                    return;
                }

                if (id == HotkeyOcrId)
                {
                    if (_settings.OcrEnabled)
                    {
                        await ExecuteOcrFromSelectionAsync();
                    }
                    return;
                }

                if (id == HotkeyShowHideAppId)
                {
                    ToggleMainWindowVisibility();
                    return;
                }

                if (id == HotkeyRestartExplorerId)
                {
                    if (_settings.RestartExplorerHotkeyEnabled)
                    {
                        RestartExplorerAction();
                    }
                    return;
                }

                if (id == HotkeyFileListId)
                {
                    if (_settings.FileListExportEnabled)
                    {
                        ExecuteExplorerFileListExport();
                    }
                }
            });
            return nint.Zero;
        }

        if (msg == NativeMethods.WmDestroy)
        {
            UnregisterHotkeys();
            DisableGlobalMiddleClickHook();
        }

        return NativeMethods.CallWindowProc(_oldWndProc, hWnd, msg, wParam, lParam);
    }

    private async void StartupToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        var targetState = StartupToggle.IsOn;
        var previousState = _settings.StartWithWindows;

        try
        {
            _settings.StartWithWindows = targetState;
            await _startupService.SetEnabledAsync(_settings.StartWithWindows);
            await _settingsService.SaveAsync(_settings);
            SetStatus($"{Localize("ui.startup.title")}: {(targetState ? Localize("ui.toggle.enabled") : Localize("ui.toggle.disabled"))}");
        }
        catch (Exception ex)
        {
            _settings.StartWithWindows = previousState;
            _isLoading = true;
            try
            {
                StartupToggle.IsOn = previousState;
            }
            finally
            {
                _isLoading = false;
            }

            SetStatus($"{Localize("ui.startup.title")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void MiddleClickToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        _settings.MinimizeOnMiddleClickTitle = MiddleClickToggle.IsOn;
        await _settingsService.SaveAsync(_settings);
        EnableGlobalMiddleClickHook(_settings.MinimizeOnMiddleClickTitle);
        SetStatus(_settings.MinimizeOnMiddleClickTitle ? Localize("middle.enabled") : Localize("middle.disabled"));
    }

    private async void HideShortcutNamesToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        try
        {
            if (HideShortcutNamesToggle.IsOn)
            {
                await ApplyShortcutLabelHiddenStateAsync(true);
                _settings.HideShortcutNames = true;
                SetStatus(Localize("labels.hidden"));
            }
            else
            {
                await ApplyShortcutLabelHiddenStateAsync(false);
                _settings.HideShortcutNames = false;
                SetStatus(Localize("labels.restored"));
            }

            await _settingsService.SaveAsync(_settings);
        }
        catch (Exception ex)
        {
            HideShortcutNamesToggle.IsOn = _settings.HideShortcutNames;
            SetStatus($"{Localize("labels.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void HideShortcutArrowsToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        var targetState = HideShortcutArrowsToggle.IsOn;

        try
        {
            _shellTweaksService.SetShortcutArrowsHidden(targetState);
            _settings.HideShortcutArrows = targetState;
            await _settingsService.SaveAsync(_settings);

            if (targetState && _shellTweaksService.NeedsLegacyMachineShortcutArrowCleanup())
            {
                var proceed = await ConfirmAsync(
                    Localize("legacyArrows.migration.title"),
                    Localize("legacyArrows.migration.internet"),
                    Localize("continue"));

                if (!proceed)
                {
                    SetStatus(Localize("legacyArrows.migration.pending"), InfoBarSeverity.Warning);
                    return;
                }

                var ok = _shellTweaksService.TryRestartAsAdminForLegacyShortcutArrowCleanupAndApply(targetState);
                if (!ok)
                {
                    SetStatus(Localize("legacyArrows.migration.cancel"), InfoBarSeverity.Warning);
                    return;
                }

                ExitApplication();
                return;
            }

            SetStatus(Localize("arrows.updated"), InfoBarSeverity.Success);

            var restartNow = await ConfirmAsync(
                Localize("explorer.title"),
                Localize("explorer.message"),
                Localize("explorer.restart"));

            if (restartNow)
            {
                _shellTweaksService.RestartExplorer();
                SetStatus(Localize("explorer.restarted"), InfoBarSeverity.Success);
            }
        }
        catch (Exception ex)
        {
            HideShortcutArrowsToggle.IsOn = _settings.HideShortcutArrows;
            SetStatus($"{Localize("arrows.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void HideDesktopShortcutArrowsToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        var targetState = HideDesktopShortcutArrowsToggle.IsOn;

        try
        {
            _shellTweaksService.SetDesktopShortcutArrowsHidden(targetState);
            _settings.HideDesktopShortcutArrows = targetState;
            await _settingsService.SaveAsync(_settings);

            if (!targetState && _shellTweaksService.NeedsLegacyMachineDesktopArrowCleanup())
            {
                var proceed = await ConfirmAsync(
                    Localize("legacyArrows.migration.title"),
                    Localize("legacyArrows.migration.desktop"),
                    Localize("continue"));

                if (!proceed)
                {
                    SetStatus(Localize("legacyArrows.migration.pending"), InfoBarSeverity.Warning);
                    return;
                }

                var ok = _shellTweaksService.TryRestartAsAdminForLegacyDesktopArrowCleanup();
                if (!ok)
                {
                    SetStatus(Localize("legacyArrows.migration.cancel"), InfoBarSeverity.Warning);
                    return;
                }

                ExitApplication();
                return;
            }

            if (targetState && _shellTweaksService.IsDesktopShortcutOverlayPotentiallyUnsupported())
            {
                SetStatus(Localize("desktopArrows.compat.warning"), InfoBarSeverity.Warning);
            }
            else
            {
                SetStatus(Localize("arrows.updated"), InfoBarSeverity.Success);
            }
            await ShowDesktopArrowsRestartDialogAndRestartAsync(targetState);
        }
        catch (Exception ex)
        {
            HideDesktopShortcutArrowsToggle.IsOn = _settings.HideDesktopShortcutArrows;
            SetStatus($"{Localize("arrows.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void OcrToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        _settings.OcrEnabled = OcrToggle.IsOn;
        await _settingsService.SaveAsync(_settings);
        SetStatus(_settings.OcrEnabled ? Localize("ocr.enabled") : Localize("ocr.disabled"));
    }

    private async void LanguageCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        var tag = GetSelectedTag(LanguageCombo);
        _settings.LanguageMode = tag;

        LocalizationService.SetLanguageMode(tag);
        LocalizationService.Reload();
        ApplyLocalization();
        await _settingsService.SaveAsync(_settings);
        SetStatus(Localize("lang.changed"));
    }

    private async void ThemeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        var tag = GetSelectedTag(ThemeCombo);
        _settings.ThemeMode = tag;
        ApplyTheme(tag);
        await _settingsService.SaveAsync(_settings);
    }

    private void ApplyTheme(string themeMode)
    {
        RootGrid.RequestedTheme = themeMode switch
        {
            "light" => ElementTheme.Light,
            "dark" => ElementTheme.Dark,
            _ => ElementTheme.Default
        };
    }

    private void RestartExplorerAction()
    {
        try
        {
            _shellTweaksService.RestartExplorerWithIconCacheRefresh();
            SetStatus(Localize("explorer.restarted"), InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("explorer.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void RestartExplorerToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        _settings.RestartExplorerHotkeyEnabled = RestartExplorerToggle.IsOn;
        await _settingsService.SaveAsync(_settings);
        SetStatus(_settings.RestartExplorerHotkeyEnabled ? Localize("restart.enabled") : Localize("restart.disabled"));
    }

    private async void FileListToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        _settings.FileListExportEnabled = FileListToggle.IsOn;
        FileListModeCombo.IsEnabled = _settings.FileListExportEnabled;
        await _settingsService.SaveAsync(_settings);
        SetStatus(_settings.FileListExportEnabled ? Localize("filelist.enabled") : Localize("filelist.disabled"));
    }

    private async void FileListModeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        _settings.FileListExportMode = NormalizeFileListExportMode(GetSelectedTag(FileListModeCombo));
        await _settingsService.SaveAsync(_settings);
    }

    private async void TopMostToggle_Toggled(object sender, RoutedEventArgs e)
    {
        if (_isLoading)
        {
            return;
        }

        _settings.AlwaysOnTopEnabled = TopMostToggle.IsOn;
        await _settingsService.SaveAsync(_settings);
        SetStatus(_settings.AlwaysOnTopEnabled ? Localize("top.feature.enabled") : Localize("top.feature.disabled"));
    }

    private async Task ExecuteOcrFromSelectionAsync()
    {
        try
        {
            SetStatus(Localize("ocr.select"));
            var text = await _ocrService.CaptureAreaAndExtractTextAsync();
            if (string.IsNullOrWhiteSpace(text))
            {
                SetStatus(Localize("ocr.empty"), InfoBarSeverity.Warning);
                return;
            }

            var package = new DataPackage();
            package.SetText(text);
            Clipboard.SetContent(package);

            SetStatus(Localize("ocr.done"), InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("ocr.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private void ExecuteExplorerFileListExport()
    {
        try
        {
            var mode = NormalizeFileListExportMode(_settings.FileListExportMode);
            var result = _explorerFileListExportService.ExportFromActiveExplorer(mode);

            if (string.Equals(result.Mode, ExplorerFileListExportService.ModeCsvFile, StringComparison.OrdinalIgnoreCase))
            {
                SetStatus($"{Localize("filelist.done.csv")} {result.CsvPath}", InfoBarSeverity.Success);
                return;
            }

            var package = new DataPackage();
            package.SetText(result.ClipboardText ?? string.Empty);
            Clipboard.SetContent(package);
            SetStatus(Localize("filelist.done.clipboard"), InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("filelist.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void BackupSettings_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            using var dialog = new WinForms.SaveFileDialog
            {
                Filter = ResolveJsonDialogFilter(),
                DefaultExt = "json",
                AddExtension = true,
                FileName = "easy-utilities-settings.json",
                OverwritePrompt = true,
                RestoreDirectory = true
            };

            if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
            {
                return;
            }

            // Persist latest in-memory state first, then export a signed backup envelope.
            await _settingsService.SaveAsync(_settings);

            var payload = new SettingsBackupPayload
            {
                BackupType = SettingsBackupType,
                Version = SettingsBackupVersion,
                CreatedAtUtc = DateTime.UtcNow,
                Settings = _settings
            };

            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            });

            await File.WriteAllTextAsync(dialog.FileName, json);
            SetStatus(Localize("backup.done"), InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("backup.errorPrefix")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void SaveDesktopIcons_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            using var dialog = new WinForms.SaveFileDialog
            {
                Filter = ResolveJsonDialogFilter(),
                DefaultExt = "json",
                AddExtension = true,
                FileName = "easy-utilities-desktop-icons.json",
                OverwritePrompt = true,
                RestoreDirectory = true
            };

            if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
            {
                return;
            }

            await _desktopIconLayoutService.SaveToFileAsync(dialog.FileName);
            SetStatus(Localize("icons.save.done"), InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("icons.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void RestoreDesktopIcons_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            using var dialog = new WinForms.OpenFileDialog
            {
                Filter = ResolveJsonDialogFilter(),
                DefaultExt = "json",
                CheckFileExists = true,
                RestoreDirectory = true
            };

            if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
            {
                return;
            }

            await _desktopIconLayoutService.RestoreFromFileAsync(dialog.FileName);
            SetStatus(Localize("icons.restore.done"), InfoBarSeverity.Success);

            var restartNow = await ConfirmAsync(
                Localize("explorer.title"),
                Localize("icons.restore.restart"),
                Localize("explorer.restart"));

            if (restartNow)
            {
                _shellTweaksService.RestartExplorer();
                SetStatus(Localize("explorer.restarted"), InfoBarSeverity.Success);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("icons.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async void RestoreSettings_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var applyWarnings = new List<string>();

            using var dialog = new WinForms.OpenFileDialog
            {
                Filter = ResolveJsonDialogFilter(),
                DefaultExt = "json",
                CheckFileExists = true,
                RestoreDirectory = true
            };

            if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
            {
                return;
            }

            var json = await File.ReadAllTextAsync(dialog.FileName);
            if (!TryParseSettingsBackup(json, out var restored))
            {
                SetStatus(Localize("restore.bad"), InfoBarSeverity.Warning);
                return;
            }

            _settings = restored;
            _settings.ShortcutRenameMaps ??= [];
            await _settingsService.SaveAsync(_settings);
            await _startupService.SetEnabledAsync(_settings.StartWithWindows);

            try
            {
                _shellTweaksService.SetShortcutArrowsHidden(_settings.HideShortcutArrows);
                if (_settings.HideShortcutArrows && _shellTweaksService.NeedsLegacyMachineShortcutArrowCleanup())
                {
                    applyWarnings.Add(Localize("legacyArrows.migration.pending"));
                }
            }
            catch (Exception ex)
            {
                applyWarnings.Add($"{Localize("arrows.error")}: {ex.Message}");
            }

            try
            {
                _shellTweaksService.SetDesktopShortcutArrowsHidden(_settings.HideDesktopShortcutArrows);
                if (!_settings.HideDesktopShortcutArrows && _shellTweaksService.NeedsLegacyMachineDesktopArrowCleanup())
                {
                    applyWarnings.Add(Localize("legacyArrows.migration.pending"));
                }
                else if (_settings.HideDesktopShortcutArrows &&
                         _shellTweaksService.IsDesktopShortcutOverlayPotentiallyUnsupported())
                {
                    applyWarnings.Add(Localize("desktopArrows.compat.warning"));
                }
            }
            catch (Exception ex)
            {
                applyWarnings.Add($"{Localize("arrows.error")}: {ex.Message}");
            }

            _isLoading = true;
            SelectComboByTag(LanguageCombo, _settings.LanguageMode);
            SelectComboByTag(ThemeCombo, _settings.ThemeMode);
            StartupToggle.IsOn = _settings.StartWithWindows;
            MiddleClickToggle.IsOn = _settings.MinimizeOnMiddleClickTitle;
            HideShortcutArrowsToggle.IsOn = _settings.HideShortcutArrows;
            HideDesktopShortcutArrowsToggle.IsOn = _settings.HideDesktopShortcutArrows;
            HideShortcutNamesToggle.IsOn = _settings.HideShortcutNames;
            TopMostToggle.IsOn = _settings.AlwaysOnTopEnabled;
            OcrToggle.IsOn = _settings.OcrEnabled;
            RestartExplorerToggle.IsOn = _settings.RestartExplorerHotkeyEnabled;
            FileListToggle.IsOn = _settings.FileListExportEnabled;
            SelectComboByTag(FileListModeCombo, NormalizeFileListExportMode(_settings.FileListExportMode));
            FileListModeCombo.IsEnabled = _settings.FileListExportEnabled;
            LocalizationService.SetLanguageMode(_settings.LanguageMode);
            LocalizationService.Reload();
            ApplyTheme(_settings.ThemeMode);
            ApplyLocalization();
            _isLoading = false;

            if (_settings.ShortcutRenameMaps.Count > 0)
            {
                _shortcutNameService.RestoreDesktopShortcutLabels(_settings.ShortcutRenameMaps);
                _settings.ShortcutRenameMaps.Clear();
            }

            await ApplyShortcutLabelHiddenStateAsync(_settings.HideShortcutNames);
            EnableGlobalMiddleClickHook(_settings.MinimizeOnMiddleClickTitle);
            if (applyWarnings.Count == 0)
            {
                SetStatus(Localize("restore.done"), InfoBarSeverity.Success);
            }
            else
            {
                SetStatus($"{Localize("restore.done")} {string.Join(" ", applyWarnings)}", InfoBarSeverity.Warning);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("restore.errorPrefix")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private void ToggleTopMostForWindow(nint target)
    {
        if (target == nint.Zero)
        {
            return;
        }

        var title = NativeMethods.GetWindowTitle(target);
        var exStyle = (nuint)NativeMethods.GetWindowLongPtr(target, NativeMethods.GwlExStyle);
        var isTopMost = (exStyle & NativeMethods.WsExTopmost) != 0;

        NativeMethods.SetWindowPos(
            target,
            new nint(isTopMost ? NativeMethods.HwndNotTopmost : NativeMethods.HwndTopmost),
            0,
            0,
            0,
            0,
            NativeMethods.SwpNomove | NativeMethods.SwpNosize | NativeMethods.SwpNoactivate);

        SetStatus($"{title}: {(isTopMost ? Localize("top.off") : Localize("top.on"))}", InfoBarSeverity.Success);
    }

    private void ToggleMainWindowVisibility()
    {
        if (_hwnd == nint.Zero)
        {
            return;
        }

        if (!NativeMethods.IsWindowVisible(_hwnd))
        {
            ShowFromTray();
            return;
        }

        MinimizeToTray();
    }

    private void ShowFromTray()
    {
        NativeMethods.ShowWindow(_hwnd, NativeMethods.SwShow);
        Activate();
    }

    private void ExitApplication()
    {
        _ = ExitApplicationAsync();
    }

    private async Task ExitApplicationAsync()
    {
        // Give ContentDialog close animation a moment to finish before exiting.
        await Task.Delay(100);

        _allowClose = true;

        try
        {
            await _settingsService.SaveAsync(_settings);
        }
        catch
        {
            // no-op
        }

        try
        {
            if (_appWindow is not null)
            {
                _appWindow.Closing -= AppWindow_Closing;
            }
        }
        catch
        {
            // no-op
        }

        try
        {
            DisableGlobalMiddleClickHook();
        }
        catch
        {
            // no-op
        }

        try
        {
            _trayIcon?.Dispose();
        }
        catch
        {
            // no-op
        }

        try
        {
            UnregisterHotkeys();
        }
        catch
        {
            // no-op
        }

        try
        {
            ProcessLifetimeService.CleanupOnExit();
        }
        catch
        {
            // no-op
        }

        try
        {
            Application.Current.Exit();
        }
        catch
        {
            Environment.Exit(0);
        }
    }

    private void EnableGlobalMiddleClickHook(bool enable)
    {
        if (!enable)
        {
            DisableGlobalMiddleClickHook();
            return;
        }

        if (_mouseHookHandle != nint.Zero)
        {
            return;
        }

        _mouseHookProc = GlobalMouseHookProc;
        _mouseHookHandle = NativeMethods.SetWindowsHookEx(NativeMethods.WhMouseLl, _mouseHookProc, nint.Zero, 0);
        if (_mouseHookHandle == nint.Zero)
        {
            SetStatus(Localize("middle.hook.failed"), InfoBarSeverity.Warning);
        }
    }

    private void DisableGlobalMiddleClickHook()
    {
        if (_mouseHookHandle == nint.Zero)
        {
            return;
        }

        _ = NativeMethods.UnhookWindowsHookEx(_mouseHookHandle);
        _mouseHookHandle = nint.Zero;
        _mouseHookProc = null;
    }

    private nint GlobalMouseHookProc(int nCode, nint wParam, nint lParam)
    {
        if (nCode < 0 || !_settings.MinimizeOnMiddleClickTitle || wParam.ToInt32() != NativeMethods.WmMButtonUp)
        {
            return NativeMethods.CallNextHookEx(_mouseHookHandle, nCode, wParam, lParam);
        }

        var hookData = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
        var hwnd = NativeMethods.WindowFromPoint(hookData.pt);
        var root = NativeMethods.GetAncestor(hwnd, NativeMethods.GaRoot);

        if (root == nint.Zero || !NativeMethods.IsWindowVisible(root))
        {
            return NativeMethods.CallNextHookEx(_mouseHookHandle, nCode, wParam, lParam);
        }

        if (!NativeMethods.GetWindowRect(root, out var rect))
        {
            return NativeMethods.CallNextHookEx(_mouseHookHandle, nCode, wParam, lParam);
        }

        var insideTopBand = hookData.pt.Y >= rect.Top && hookData.pt.Y <= rect.Top + 120 && hookData.pt.X >= rect.Left && hookData.pt.X <= rect.Right;
        if (insideTopBand)
        {
            _ = NativeMethods.PostMessage(root, NativeMethods.WmSysCommand, new nint(NativeMethods.ScMinimize), nint.Zero);
        }

        return NativeMethods.CallNextHookEx(_mouseHookHandle, nCode, wParam, lParam);
    }

    private async Task<bool> ConfirmAsync(string title, string message, string primaryButton)
    {
        var xamlRoot = await GetDialogXamlRootAsync();
        if (xamlRoot is null)
        {
            SetStatus(message, InfoBarSeverity.Warning);
            return false;
        }

        var dialog = new ContentDialog
        {
            Title = title,
            Content = message,
            PrimaryButtonText = primaryButton,
            CloseButtonText = Localize("cancel"),
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = xamlRoot
        };

        var result = await dialog.ShowAsync();
        return result == ContentDialogResult.Primary;
    }

    private async Task ShowInfoAsync(string title, string message)
    {
        var xamlRoot = await GetDialogXamlRootAsync();
        if (xamlRoot is null)
        {
            SetStatus(message, InfoBarSeverity.Warning);
            return;
        }

        var dialog = new ContentDialog
        {
            Title = title,
            Content = message,
            CloseButtonText = Localize("cancel"),
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = xamlRoot
        };

        _ = await dialog.ShowAsync();
    }

    private async Task ShowDesktopArrowsRestartDialogAndRestartAsync(bool hidden)
    {
        var message = hidden
            ? Localize("desktopArrows.enabled.restart")
            : Localize("desktopArrows.disabled.restart");

        var xamlRoot = await GetDialogXamlRootAsync();
        if (xamlRoot is null)
        {
            SetStatus(message, InfoBarSeverity.Warning);
            return;
        }

        var dialog = new ContentDialog
        {
            Title = Localize("ui.info.title"),
            Content = message,
            PrimaryButtonText = Localize("ok"),
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = xamlRoot
        };

        var result = await dialog.ShowAsync();
        if (result != ContentDialogResult.Primary)
        {
            return;
        }

        try
        {
            _shellTweaksService.RestartExplorer();
            SetStatus(Localize("explorer.restarted"), InfoBarSeverity.Success);
        }
        catch (Exception ex)
        {
            SetStatus($"{Localize("explorer.error")}: {ex.Message}", InfoBarSeverity.Error);
        }
    }

    private async Task<XamlRoot?> GetDialogXamlRootAsync()
    {
        for (var i = 0; i < 40; i++)
        {
            var xamlRoot = RootGrid.XamlRoot ?? (Content as FrameworkElement)?.XamlRoot;
            if (xamlRoot is not null)
            {
                return xamlRoot;
            }

            await Task.Delay(50);
        }

        return null;
    }

    private void SetStatus(string message, InfoBarSeverity severity = InfoBarSeverity.Informational)
    {
        _statusCts?.Cancel();
        _statusCts = new CancellationTokenSource();

        StatusBar.Severity = severity;
        StatusBar.Message = message;
        StatusBar.IsOpen = true;

        var token = _statusCts.Token;
        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(TimeSpan.FromSeconds(5), token);
                if (token.IsCancellationRequested)
                {
                    return;
                }

                _ = DispatcherQueue.TryEnqueue(() => StatusBar.IsOpen = false);
            }
            catch (TaskCanceledException)
            {
                // no-op
            }
        }, token);
    }

    private void ApplyLocalization()
    {
        SectionTitleText.Text = Localize("ui.section.system");
        AppTitleText.Text = Localize("ui.app.title");

        StartupTitleText.Text = Localize("ui.startup.title");
        StartupDescText.Text = Localize("ui.startup.desc");

        MiddleClickTitleText.Text = Localize("ui.middleClick.title");
        MiddleClickDescText.Text = Localize("ui.middleClick.desc");

        HideNamesTitleText.Text = Localize("ui.hideNames.title");
        HideNamesDescText.Text = Localize("ui.hideNames.desc");

        HideArrowsTitleText.Text = Localize("ui.hideArrows.title");
        HideArrowsDescText.Text = Localize("ui.hideArrows.desc");
        HideDesktopArrowsTitleText.Text = Localize("ui.hideDesktopArrows.title");
        HideDesktopArrowsDescText.Text = Localize("ui.hideDesktopArrows.desc");

        LanguageTitleText.Text = Localize("ui.language.title");
        LanguageDescText.Text = Localize("ui.language.desc");
        LanguageSystemItem.Content = Localize("ui.language.system");
        LanguageUkItem.Content = Localize("ui.language.uk");
        LanguageEnItem.Content = Localize("ui.language.en");

        ThemeTitleText.Text = Localize("ui.theme.title");
        ThemeDescText.Text = Localize("ui.theme.desc");
        ThemeSystemItem.Content = Localize("ui.theme.system");
        ThemeLightItem.Content = Localize("ui.theme.light");
        ThemeDarkItem.Content = Localize("ui.theme.dark");

        TopMostTitleText.Text = Localize("ui.topMost.title");
        TopMostDescText.Text = Localize("ui.topMost.desc");

        var toggleEnabled = Localize("ui.toggle.enabled");
        var toggleDisabled = Localize("ui.toggle.disabled");
        ApplyToggleLabels(StartupToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(MiddleClickToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(HideShortcutNamesToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(HideShortcutArrowsToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(HideDesktopShortcutArrowsToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(TopMostToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(OcrToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(RestartExplorerToggle, toggleEnabled, toggleDisabled);
        ApplyToggleLabels(FileListToggle, toggleEnabled, toggleDisabled);

        OcrTitleText.Text = Localize("ui.ocr.title");
        OcrDescText.Text = Localize("ui.ocr.desc");

        RestartExplorerTitleText.Text = Localize("ui.restartExplorer.title");
        RestartExplorerDescText.Text = Localize("ui.restartExplorer.desc");
        FileListTitleText.Text = Localize("ui.fileList.title");
        FileListDescText.Text = Localize("ui.fileList.desc");
        FileListModeClipboardItem.Content = Localize("ui.fileList.mode.clipboard");
        FileListModeCsvItem.Content = Localize("ui.fileList.mode.csv");

        BackupTitleText.Text = Localize("ui.backup.title");
        BackupDescText.Text = Localize("ui.backup.desc");
        BackupButton.Content = Localize("ui.backup.button");

        IconsSaveTitleText.Text = Localize("ui.iconsSave.title");
        IconsSaveDescText.Text = Localize("ui.iconsSave.desc");
        IconsSaveButton.Content = Localize("ui.iconsSave.button");

        IconsRestoreTitleText.Text = Localize("ui.iconsRestore.title");
        IconsRestoreDescText.Text = Localize("ui.iconsRestore.desc");
        IconsRestoreButton.Content = Localize("ui.iconsRestore.button");

        RestoreTitleText.Text = Localize("ui.restore.title");
        RestoreDescText.Text = Localize("ui.restore.desc");
        RestoreButton.Content = Localize("ui.restore.button");

        _trayIcon?.UpdateTexts(
            Localize("ui.app.title"),
            Localize("tray.open"),
            Localize("tray.exit"));
    }

    private string Localize(string key)
    {
        return LocalizationService.Get(key);
    }

    private static void SelectComboByTag(ComboBox comboBox, string tag)
    {
        foreach (var item in comboBox.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Tag?.ToString(), tag, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
                return;
            }
        }

        comboBox.SelectedIndex = 0;
    }

    private static string GetSelectedTag(ComboBox comboBox)
    {
        return (comboBox.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "system";
    }

    private static string NormalizeFileListExportMode(string? mode)
    {
        return string.Equals(mode, ExplorerFileListExportService.ModeCsvFile, StringComparison.OrdinalIgnoreCase)
            ? ExplorerFileListExportService.ModeCsvFile
            : ExplorerFileListExportService.ModeClipboard;
    }

    private string ResolveJsonDialogFilter()
    {
        const string fallback = "JSON files (*.json)|*.json|All files (*.*)|*.*";
        var localized = Localize("dialog.filter.json");

        if (string.IsNullOrWhiteSpace(localized))
        {
            return fallback;
        }

        var parts = localized.Split('|', StringSplitOptions.None);
        return parts.Length >= 2 && parts.Length % 2 == 0 ? localized : fallback;
    }

    private static void ApplyToggleLabels(ToggleSwitch toggle, string onContent, string offContent)
    {
        toggle.OnContent = onContent;
        toggle.OffContent = offContent;
    }

    private async Task ApplyShortcutLabelHiddenStateAsync(bool hide)
    {
        if (hide)
        {
            // Migrate from legacy rename mode if old mappings still exist.
            if (_settings.ShortcutRenameMaps.Count > 0)
            {
                _shortcutNameService.RestoreDesktopShortcutLabels(_settings.ShortcutRenameMaps);
                _settings.ShortcutRenameMaps.Clear();
            }

            _settings.DesktopLabelRestoreColor = _shortcutNameService.GetCurrentTextColor();
            _settings.DesktopLabelRestoreBkColor = _shortcutNameService.GetCurrentTextBackgroundColor();
            _settings.DesktopLabelRestoreShadowEnabled = _shortcutNameService.GetListViewShadowEnabled();

            _shortcutNameService.SetShortcutNamesHidden(
                hidden: true,
                restoreColor: _settings.DesktopLabelRestoreColor,
                restoreTextBkColor: _settings.DesktopLabelRestoreBkColor,
                restoreShadowEnabled: _settings.DesktopLabelRestoreShadowEnabled);
            return;
        }

        if (_settings.ShortcutRenameMaps.Count > 0)
        {
            _shortcutNameService.RestoreDesktopShortcutLabels(_settings.ShortcutRenameMaps);
            _settings.ShortcutRenameMaps.Clear();
        }

        _shortcutNameService.SetShortcutNamesHidden(
            hidden: false,
            restoreColor: _settings.DesktopLabelRestoreColor,
            restoreTextBkColor: _settings.DesktopLabelRestoreBkColor,
            restoreShadowEnabled: _settings.DesktopLabelRestoreShadowEnabled);
    }

    private static bool TryParseSettingsBackup(string json, out AppSettings settings)
    {
        settings = new AppSettings();

        var envelope = JsonSerializer.Deserialize<SettingsBackupPayload>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (envelope?.Settings is not null &&
            string.Equals(envelope.BackupType, SettingsBackupType, StringComparison.OrdinalIgnoreCase) &&
            envelope.Version == SettingsBackupVersion)
        {
            settings = envelope.Settings;
            settings.ShortcutRenameMaps ??= [];
            return true;
        }

        // Backward compatibility with legacy raw settings backup.
        try
        {
            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            var keyCount = 0;
            if (doc.RootElement.TryGetProperty("startWithWindows", out _)) keyCount++;
            if (doc.RootElement.TryGetProperty("minimizeOnMiddleClickTitle", out _)) keyCount++;
            if (doc.RootElement.TryGetProperty("languageMode", out _)) keyCount++;
            if (doc.RootElement.TryGetProperty("themeMode", out _)) keyCount++;
            if (doc.RootElement.TryGetProperty("alwaysOnTopEnabled", out _)) keyCount++;
            if (doc.RootElement.TryGetProperty("ocrEnabled", out _)) keyCount++;

            if (keyCount < 2)
            {
                return false;
            }

            var legacy = JsonSerializer.Deserialize<AppSettings>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            if (legacy is null)
            {
                return false;
            }

            legacy.ShortcutRenameMaps ??= [];
            settings = legacy;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private sealed class SettingsBackupPayload
    {
        public string BackupType { get; set; } = string.Empty;
        public int Version { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public AppSettings? Settings { get; set; }
    }

}
