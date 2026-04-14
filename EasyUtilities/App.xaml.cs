using System.Text;
using EasyUtilities.Services;
using Microsoft.Windows.AppLifecycle;
using Microsoft.UI.Xaml;

namespace EasyUtilities;

public partial class App : Application
{
    private MainWindow? _mainWindow;

    public static string[] Arguments { get; } = Environment.GetCommandLineArgs();

    public App()
    {
        UnhandledException += App_UnhandledException;
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
        AppDomain.CurrentDomain.ProcessExit += (_, _) => ProcessLifetimeService.CleanupOnExit();
        TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
    }

    protected override async void OnLaunched(Microsoft.UI.Xaml.LaunchActivatedEventArgs args)
    {
        _mainWindow = new MainWindow();
        _mainWindow.Activate();

        var hasShortcutAction = ArgumentParser.TryGetShortcutArrowAction(Arguments, out var hideArrows);
        var hasDesktopShortcutAction = ArgumentParser.TryGetDesktopShortcutArrowAction(Arguments, out var hideDesktopArrows);

        if (ArgumentParser.HasLegacyShortcutArrowCleanup(Arguments))
        {
            try
            {
                await _mainWindow.HandleStartupLegacyShortcutArrowCleanupAsync(showRestartPrompt: !hasShortcutAction);
            }
            catch (Exception ex)
            {
                LogFatal("Failed to cleanup legacy startup shortcut-arrow state.", ex);
            }
        }

        if (ArgumentParser.HasLegacyDesktopShortcutArrowCleanup(Arguments))
        {
            try
            {
                await _mainWindow.HandleStartupLegacyDesktopShortcutArrowCleanupAsync(showRestartPrompt: !hasDesktopShortcutAction);
            }
            catch (Exception ex)
            {
                LogFatal("Failed to cleanup legacy startup desktop shortcut-arrow state.", ex);
            }
        }

        if (hasShortcutAction)
        {
            try
            {
                await _mainWindow.HandleStartupShortcutArrowActionAsync(hideArrows);
            }
            catch (Exception ex)
            {
                LogFatal("Failed to apply startup shortcut-arrow action.", ex);
            }
        }

        if (hasDesktopShortcutAction)
        {
            try
            {
                await _mainWindow.HandleStartupDesktopShortcutArrowActionAsync(hideDesktopArrows);
            }
            catch (Exception ex)
            {
                LogFatal("Failed to apply startup desktop shortcut-arrow action.", ex);
            }
        }

        if (ArgumentParser.HasStartMinimized(Arguments) || IsStartupTaskActivation())
        {
            _mainWindow.MinimizeToTray();
        }
    }

    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        LogFatal("XAML unhandled exception", e.Exception);
    }

    private void CurrentDomain_UnhandledException(object? sender, System.UnhandledExceptionEventArgs e)
    {
        LogFatal("AppDomain unhandled exception", e.ExceptionObject as Exception);
    }

    private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        LogFatal("TaskScheduler unobserved exception", e.Exception);
        e.SetObserved();
    }

    private static void LogFatal(string title, Exception? exception)
    {
        try
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EasyUtilities");
            Directory.CreateDirectory(dir);
            var file = Path.Combine(dir, "fatal.log");

            var sb = new StringBuilder();
            sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}");
            sb.AppendLine(exception?.ToString() ?? "No managed exception object.");
            sb.AppendLine(new string('-', 70));

            File.AppendAllText(file, sb.ToString());
        }
        catch
        {
            // no-op
        }
    }

    private static bool IsStartupTaskActivation()
    {
        try
        {
            var activation = AppInstance.GetCurrent().GetActivatedEventArgs();
            return activation.Kind == ExtendedActivationKind.StartupTask;
        }
        catch
        {
            return false;
        }
    }
}
