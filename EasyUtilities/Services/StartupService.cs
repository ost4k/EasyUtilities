using Microsoft.Win32;
using Windows.ApplicationModel;

namespace EasyUtilities.Services;

public sealed class StartupService
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "EasyUtilities";
    private const string StartupTaskId = "EasyUtilitiesStartup";

    public async Task SetEnabledAsync(bool enabled)
    {
        if (HasPackageIdentity())
        {
            DeleteLegacyRunValue();
            await SetPackagedStartupEnabledAsync(enabled);
            return;
        }

        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true)
            ?? throw new InvalidOperationException("Could not open Run key.");

        if (enabled)
        {
            var exePath = Environment.ProcessPath ?? throw new InvalidOperationException("Could not detect executable path.");
            key.SetValue(ValueName, $"\"{exePath}\" --start-minimized");
            return;
        }

        key.DeleteValue(ValueName, throwOnMissingValue: false);
    }

    public async Task<bool> IsEnabledAsync()
    {
        if (HasPackageIdentity())
        {
            DeleteLegacyRunValue();
            var startupTask = await StartupTask.GetAsync(StartupTaskId);
            return startupTask.State is StartupTaskState.Enabled or StartupTaskState.EnabledByPolicy;
        }

        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        return key?.GetValue(ValueName) is string;
    }

    private static bool HasPackageIdentity()
    {
        try
        {
            _ = Package.Current;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static async Task SetPackagedStartupEnabledAsync(bool enabled)
    {
        var startupTask = await StartupTask.GetAsync(StartupTaskId);

        if (!enabled)
        {
            startupTask.Disable();
            return;
        }

        if (startupTask.State is StartupTaskState.Enabled or StartupTaskState.EnabledByPolicy)
        {
            return;
        }

        var state = await startupTask.RequestEnableAsync();
        if (state is not StartupTaskState.Enabled and not StartupTaskState.EnabledByPolicy)
        {
            throw new InvalidOperationException($"Could not enable startup task. Current state: {state}.");
        }
    }

    private static void DeleteLegacyRunValue()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: true);
        key?.DeleteValue(ValueName, throwOnMissingValue: false);
    }
}
