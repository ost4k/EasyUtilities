using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace EasyUtilities.Services;

public sealed class ShellTweaksService
{
    private const uint ShcneAssocChanged = 0x08000000;
    private const uint ShcnfIdList = 0x0000;
    private const string ShellIconsKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons";
    private const string OverlayValueName = "29";
    private const string ClassesRootPath = @"SOFTWARE\Classes";
    private const string InternetShortcutClass = "InternetShortcut";
    private const string FolderClass = "Folder";
    private const string DirectoryClass = "Directory";
    private const string DriveClass = "Drive";
    private const string LinkShortcutClass = "lnkfile";
    private const string PifShortcutClass = "piffile";
    private const string IsShortcutValueName = "IsShortcut";
    private const string DesktopShortcutOverlayIconFileName = "shortcut-overlay-empty.ico";

    public bool IsAdministrator()
    {
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    public void SetShortcutArrowsHidden(bool hide)
    {
        SetInternetShortcutIsShortcutInHive(RegistryHive.CurrentUser, hide);
        NotifyShellAssociationChanged();
    }

    public void SetDesktopShortcutArrowsHidden(bool hide)
    {
        // Safety first: ensure legacy/broken per-user overrides do not break shortcut launching.
        // Best effort only, must not block the main feature path.
        TryRepairUserShortcutClassAssociations();

        // Primary path: Shell Icons overlay override.
        if (hide)
        {
            var overlayIconPath = EnsureOverlayIconAvailable();
            var value = $"{overlayIconPath},0";
            SetOverlayValueInHive(RegistryHive.CurrentUser, value);
        }
        else
        {
            DeleteOverlayValueInHive(RegistryHive.CurrentUser);
        }

        NotifyShellAssociationChanged();
        RefreshIconCache();
    }

    public void RepairUserShortcutClassAssociations()
    {
        RepairShortcutClassAssociationOverrideInHive(RegistryHive.CurrentUser, LinkShortcutClass);
        RepairShortcutClassAssociationOverrideInHive(RegistryHive.CurrentUser, PifShortcutClass);

        var repairedDirectoryClass = EnsureCurrentUserClassDefaultIfEmpty(DirectoryClass, "File Folder");
        var repairedDriveClass = EnsureCurrentUserClassDefaultIfEmpty(DriveClass, "Drive");
        var folderLooksBroken = HasBrokenCurrentUserFolderOpenCommand();
        if (repairedDirectoryClass || repairedDriveClass || folderLooksBroken)
        {
            EnsureFolderOpenFallbackInCurrentUserOverride();
            RemoveEmptyCurrentUserShellKeyIfPresent(DirectoryClass);
            RemoveEmptyCurrentUserShellKeyIfPresent(DriveClass);
        }
    }

    public bool NeedsLegacyMachineShortcutArrowCleanup()
    {
        return HasInternetShortcutValueInHive(RegistryHive.LocalMachine);
    }

    public bool NeedsLegacyMachineDesktopArrowCleanup()
    {
        return HasOverlayValueInHive(RegistryHive.LocalMachine);
    }

    public void CleanupLegacyMachineShortcutArrowState()
    {
        SetInternetShortcutIsShortcutInHive(RegistryHive.LocalMachine, hide: true);
    }

    public void CleanupLegacyMachineDesktopArrowState()
    {
        DeleteOverlayValueInHive(RegistryHive.LocalMachine);
    }

    public bool TryRestartAsAdminForLegacyShortcutArrowCleanup()
    {
        return TryRestartAsAdmin("--cleanup-legacy-shortcut-arrows");
    }

    public bool TryRestartAsAdminForLegacyShortcutArrowCleanupAndApply(bool hide)
    {
        return TryRestartAsAdmin($"--cleanup-legacy-shortcut-arrows --set-shortcut-arrows={(hide ? "on" : "off")}");
    }

    public bool TryRestartAsAdminForLegacyDesktopArrowCleanup()
    {
        return TryRestartAsAdmin("--cleanup-legacy-desktop-shortcut-arrows");
    }

    public bool IsDesktopShortcutOverlayPotentiallyUnsupported()
    {
        try
        {
            var build = Environment.OSVersion.Version.Build;
            // Windows 11 24H2+ Insider/Canary builds may ignore Shell Icons\29.
            return build >= 26100;
        }
        catch
        {
            return false;
        }
    }

    public void RestartExplorer()
    {
        var taskKill = Process.Start(new ProcessStartInfo
        {
            FileName = "taskkill",
            Arguments = "/f /im explorer.exe",
            CreateNoWindow = true,
            UseShellExecute = false
        });
        taskKill?.WaitForExit(5000);

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            UseShellExecute = true
        });
    }

    public void RestartExplorerWithIconCacheRefresh()
    {
        var taskKill = Process.Start(new ProcessStartInfo
        {
            FileName = "taskkill",
            Arguments = "/f /im explorer.exe",
            CreateNoWindow = true,
            UseShellExecute = false
        });
        taskKill?.WaitForExit(5000);

        ClearIconCacheFiles();

        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            UseShellExecute = true
        });
    }

    private static bool TryRestartAsAdmin(string arguments)
    {
        var exePath = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(exePath))
        {
            return false;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                UseShellExecute = true,
                Verb = "runas",
                Arguments = arguments
            });
            ProcessLifetimeService.SuppressSiblingShutdownForCurrentExit();
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string EnsureOverlayIconAvailable()
    {
        var source = Path.Combine(AppContext.BaseDirectory, "Assets", DesktopShortcutOverlayIconFileName);
        if (!File.Exists(source))
        {
            throw new FileNotFoundException("Overlay icon was not found.", source);
        }

        var destinationDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "EasyUtilities");
        Directory.CreateDirectory(destinationDirectory);

        var destination = Path.Combine(destinationDirectory, DesktopShortcutOverlayIconFileName);
        File.Copy(source, destination, overwrite: true);
        return destination;
    }

    private static void SetOverlayValueInHive(RegistryHive hive, string value)
    {
        Exception? lastError = null;
        var wroteAny = false;
        foreach (var view in EnumerateViewsForHive(hive))
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(hive, view);
                using var shellIcons = baseKey.CreateSubKey(ShellIconsKeyPath, writable: true)
                    ?? throw new InvalidOperationException($"Cannot access {hive}\\{ShellIconsKeyPath} ({view}).");
                shellIcons.SetValue(OverlayValueName, value, RegistryValueKind.String);
                wroteAny = true;
            }
            catch (Exception ex)
            {
                lastError = ex;
            }
        }

        if (!wroteAny && lastError is not null)
        {
            throw lastError;
        }
    }

    private static void DeleteOverlayValueInHive(RegistryHive hive)
    {
        foreach (var view in EnumerateViewsForHive(hive))
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var shellIcons = baseKey.CreateSubKey(ShellIconsKeyPath, writable: true);
            shellIcons?.DeleteValue(OverlayValueName, throwOnMissingValue: false);
        }
    }

    private static bool HasOverlayValueInHive(RegistryHive hive)
    {
        foreach (var view in EnumerateViewsForHive(hive))
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var shellIcons = baseKey.OpenSubKey(ShellIconsKeyPath, writable: false);
            if (shellIcons?.GetValue(OverlayValueName) is not null)
            {
                return true;
            }
        }

        return false;
    }

    private static void SetInternetShortcutIsShortcutInHive(RegistryHive hive, bool hide)
    {
        using var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
        using var classes = baseKey.CreateSubKey(ClassesRootPath, writable: true)
            ?? throw new InvalidOperationException($"Cannot access {hive}\\{ClassesRootPath}.");
        using var key = classes.CreateSubKey(InternetShortcutClass, writable: true)
            ?? throw new InvalidOperationException($"Cannot access {hive}\\{ClassesRootPath}\\{InternetShortcutClass}.");

        if (hide)
        {
            key.DeleteValue(IsShortcutValueName, throwOnMissingValue: false);
            return;
        }

        key.SetValue(IsShortcutValueName, string.Empty, RegistryValueKind.String);
    }

    private static void RepairShortcutClassAssociationOverrideInHive(RegistryHive hive, string className)
    {
        using var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
        using var classes = baseKey.OpenSubKey(ClassesRootPath, writable: true);
        if (classes is null)
        {
            return;
        }

        using var key = classes.OpenSubKey(className, writable: true);
        if (key is null)
        {
            return;
        }

        key.SetValue(IsShortcutValueName, string.Empty, RegistryValueKind.String);
    }

    private void TryRepairUserShortcutClassAssociations()
    {
        try
        {
            RepairUserShortcutClassAssociations();
        }
        catch
        {
            // Keep non-critical repair from breaking toggle execution.
        }
    }

    private static bool EnsureCurrentUserClassDefaultIfEmpty(string className, string expectedDefault)
    {
        using var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
        using var classes = hkcu.OpenSubKey(ClassesRootPath, writable: true);
        if (classes is null)
        {
            return false;
        }

        using var key = classes.OpenSubKey(className, writable: true);
        if (key is null)
        {
            return false;
        }

        var current = key.GetValue(null)?.ToString();
        if (!string.IsNullOrWhiteSpace(current))
        {
            return false;
        }

        key.SetValue(null, expectedDefault, RegistryValueKind.String);
        return true;
    }

    private static bool HasBrokenCurrentUserFolderOpenCommand()
    {
        using var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
        using var classes = hkcu.OpenSubKey(ClassesRootPath, writable: false);
        if (classes is null)
        {
            return false;
        }

        using var folder = classes.OpenSubKey(FolderClass, writable: false);
        if (folder is null)
        {
            return false;
        }

        var folderDefault = folder.GetValue(null)?.ToString();
        if (string.IsNullOrWhiteSpace(folderDefault))
        {
            return true;
        }

        using var openCommand = folder.OpenSubKey(@"shell\open\command", writable: false);
        var command = openCommand?.GetValue(null)?.ToString();
        return string.IsNullOrWhiteSpace(command);
    }

    private static void EnsureFolderOpenFallbackInCurrentUserOverride()
    {
        using var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
        using var classes = hkcu.CreateSubKey(ClassesRootPath, writable: true)
            ?? throw new InvalidOperationException($"Cannot access {RegistryHive.CurrentUser}\\{ClassesRootPath}.");
        using var folder = classes.CreateSubKey(FolderClass, writable: true)
            ?? throw new InvalidOperationException($"Cannot access {RegistryHive.CurrentUser}\\{ClassesRootPath}\\{FolderClass}.");
        folder.SetValue(null, "Folder", RegistryValueKind.String);

        using var shell = folder.CreateSubKey("shell", writable: true)
            ?? throw new InvalidOperationException($"Cannot access {RegistryHive.CurrentUser}\\{ClassesRootPath}\\{FolderClass}\\shell.");
        shell.SetValue(null, "open", RegistryValueKind.String);

        using var command = shell.CreateSubKey(@"open\command", writable: true)
            ?? throw new InvalidOperationException($"Cannot access {RegistryHive.CurrentUser}\\{ClassesRootPath}\\{FolderClass}\\shell\\open\\command.");
        command.SetValue(null, @"%SystemRoot%\Explorer.exe ""%1""", RegistryValueKind.ExpandString);
    }

    private static void RemoveEmptyCurrentUserShellKeyIfPresent(string className)
    {
        using var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
        using var classes = hkcu.OpenSubKey(ClassesRootPath, writable: true);
        if (classes is null)
        {
            return;
        }

        using var classKey = classes.OpenSubKey(className, writable: true);
        if (classKey is null)
        {
            return;
        }

        using var shell = classKey.OpenSubKey("shell", writable: false);
        if (shell is null)
        {
            return;
        }

        if (shell.SubKeyCount > 0)
        {
            return;
        }

        var valueNames = shell.GetValueNames();
        if (valueNames.Length > 0)
        {
            return;
        }

        classKey.DeleteSubKey("shell", throwOnMissingSubKey: false);
    }

    private static bool HasInternetShortcutValueInHive(RegistryHive hive)
    {
        using var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
        using var classes = baseKey.OpenSubKey(ClassesRootPath, writable: false);
        using var key = classes?.OpenSubKey(InternetShortcutClass, writable: false);
        if (key is null)
        {
            return false;
        }

        return key.GetValue(IsShortcutValueName) is not null;
    }

    private static IEnumerable<RegistryView> EnumerateViewsForHive(RegistryHive hive)
    {
        // On 64-bit systems, a packaged process can run as x86 or x64.
        // Write/read both views so Explorer always sees Shell Icons override.
        if (Environment.Is64BitOperatingSystem &&
            (hive == RegistryHive.CurrentUser || hive == RegistryHive.LocalMachine))
        {
            yield return RegistryView.Registry64;
            yield return RegistryView.Registry32;
            yield break;
        }

        yield return RegistryView.Default;
    }

    private static void NotifyShellAssociationChanged()
    {
        SHChangeNotify(ShcneAssocChanged, ShcnfIdList, nint.Zero, nint.Zero);
    }

    private static void RefreshIconCache()
    {
        TryRunProcess("ie4uinit.exe", "-ClearIconCache");
        TryRunProcess("ie4uinit.exe", "-show");
    }

    private static void TryRunProcess(string fileName, string arguments)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                CreateNoWindow = true,
                UseShellExecute = false
            });
            process?.WaitForExit(5000);
        }
        catch
        {
            // Best-effort cache refresh only.
        }
    }

    private static void ClearIconCacheFiles()
    {
        try
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var rootIconCache = Path.Combine(localAppData, "IconCache.db");
            TryDeleteFile(rootIconCache);

            var explorerCacheDir = Path.Combine(localAppData, "Microsoft", "Windows", "Explorer");
            if (!Directory.Exists(explorerCacheDir))
            {
                return;
            }

            foreach (var pattern in new[] { "iconcache*.db", "thumbcache*.db" })
            {
                foreach (var file in Directory.EnumerateFiles(explorerCacheDir, pattern, SearchOption.TopDirectoryOnly))
                {
                    TryDeleteFile(file);
                }
            }
        }
        catch
        {
            // Best-effort cache cleanup only.
        }
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup only.
        }
    }

    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(uint wEventId, uint uFlags, nint dwItem1, nint dwItem2);
}
