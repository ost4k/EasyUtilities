using Microsoft.Win32;

namespace EasyUtilities.Services;

public sealed class StartupService
{
    private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string ValueName = "EasyUtilities";

    public void SetEnabled(bool enabled)
    {
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

    public bool IsEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false);
        return key?.GetValue(ValueName) is string;
    }
}
