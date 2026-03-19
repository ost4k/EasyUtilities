using System.Runtime.InteropServices;
using EasyUtilities.Models;
using Microsoft.Win32;

namespace EasyUtilities.Services;

public sealed class DesktopShortcutNameService
{
    private const int LvmFirst = 0x1000;
    private const int LvmGetTextColor = LvmFirst + 35;
    private const int LvmSetTextColor = LvmFirst + 36;
    private const int LvmGetBkColor = LvmFirst;
    private const int LvmGetTextBkColor = LvmFirst + 37;
    private const int LvmSetTextBkColor = LvmFirst + 38;
    private const int WmSetRedraw = 0x000B;
    private const int WmSettingChange = 0x001A;
    private const int ClrNone = unchecked((int)0xFFFFFFFF);
    private const int ProgmanSpawnWorker = 0x052C;
    private const int ColorDesktop = 1;
    private const char HiddenNameChar = '\u2063';
    private const uint HwndBroadcast = 0xFFFF;
    private const uint SmtoAbortIfHung = 0x0002;
    private const string ExplorerAdvancedKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";
    private const string ListviewShadowValue = "ListviewShadow";

    public int GetCurrentTextColor()
    {
        var listViews = FindDesktopListViews();
        if (listViews.Count == 0)
        {
            throw new InvalidOperationException("Не знайдено список іконок робочого столу.");
        }

        return unchecked((int)SendMessage(listViews[0], LvmGetTextColor, nint.Zero, nint.Zero).ToInt64());
    }

    public int GetCurrentTextBackgroundColor()
    {
        var listViews = FindDesktopListViews();
        if (listViews.Count == 0)
        {
            throw new InvalidOperationException("Не знайдено список іконок робочого столу.");
        }

        return unchecked((int)SendMessage(listViews[0], LvmGetTextBkColor, nint.Zero, nint.Zero).ToInt64());
    }

    public int GetListViewShadowEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(ExplorerAdvancedKeyPath, writable: false);
        var value = key?.GetValue(ListviewShadowValue);
        if (value is null)
        {
            return 1;
        }

        try
        {
            return Convert.ToInt32(value) != 0 ? 1 : 0;
        }
        catch
        {
            return 1;
        }
    }

    public List<ShortcutRenameMap> HideDesktopShortcutLabels(IReadOnlyCollection<ShortcutRenameMap>? existingMaps)
    {
        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (string.IsNullOrWhiteSpace(desktop) || !Directory.Exists(desktop))
        {
            throw new InvalidOperationException("Не знайдено робочий стіл користувача.");
        }

        var maps = existingMaps?.ToList() ?? [];
        var knownOriginals = new HashSet<string>(maps.Select(m => m.OriginalPath), StringComparer.OrdinalIgnoreCase);
        var knownHidden = new HashSet<string>(maps.Select(m => m.HiddenPath), StringComparer.OrdinalIgnoreCase);

        var shortcutFiles = Directory.GetFiles(desktop, "*.*", SearchOption.TopDirectoryOnly)
            .Where(path =>
            {
                var ext = Path.GetExtension(path);
                return ext.Equals(".lnk", StringComparison.OrdinalIgnoreCase) ||
                       ext.Equals(".url", StringComparison.OrdinalIgnoreCase) ||
                       ext.Equals(".pif", StringComparison.OrdinalIgnoreCase);
            })
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var index = 1;
        foreach (var shortcutPath in shortcutFiles)
        {
            if (knownOriginals.Contains(shortcutPath) || knownHidden.Contains(shortcutPath))
            {
                continue;
            }

            var extension = Path.GetExtension(shortcutPath);
            string targetPath;
            do
            {
                var hiddenFileName = $"{new string(HiddenNameChar, index)}{extension}";
                targetPath = Path.Combine(desktop, hiddenFileName);
                index++;
            }
            while (File.Exists(targetPath) || knownHidden.Contains(targetPath));

            File.Move(shortcutPath, targetPath);

            maps.Add(new ShortcutRenameMap
            {
                OriginalPath = shortcutPath,
                HiddenPath = targetPath
            });

            knownOriginals.Add(shortcutPath);
            knownHidden.Add(targetPath);
        }

        return maps;
    }

    public void SetShortcutNamesHidden(bool hidden, int restoreColor, int restoreTextBkColor, int restoreShadowEnabled)
    {
        var listViews = FindDesktopListViews();
        if (listViews.Count == 0)
        {
            throw new InvalidOperationException("Не знайдено список іконок робочого столу.");
        }

        if (hidden)
        {
            SetListViewShadowEnabled(false);
        }
        else if (restoreShadowEnabled >= 0)
        {
            SetListViewShadowEnabled(restoreShadowEnabled != 0);
        }

        foreach (var listView in listViews)
        {
            _ = SendMessage(listView, WmSetRedraw, nint.Zero, nint.Zero);

            if (hidden)
            {
                var bgColor = unchecked((int)SendMessage(listView, LvmGetBkColor, nint.Zero, nint.Zero).ToInt64());
                if (bgColor == ClrNone)
                {
                    bgColor = GetSysColor(ColorDesktop);
                }

                // Method 1: match text and text background to desktop background.
                _ = SendMessage(listView, LvmSetTextColor, nint.Zero, new nint(bgColor));
                _ = SendMessage(listView, LvmSetTextBkColor, nint.Zero, new nint(bgColor));
                // Method 2: fallback to transparent text background.
                _ = SendMessage(listView, LvmSetTextBkColor, nint.Zero, new nint(ClrNone));
            }
            else
            {
                var color = restoreColor == 0 ? ClrNone : restoreColor;
                var textBkColor = restoreTextBkColor == 0 ? ClrNone : restoreTextBkColor;
                _ = SendMessage(listView, LvmSetTextColor, nint.Zero, new nint(color));
                _ = SendMessage(listView, LvmSetTextBkColor, nint.Zero, new nint(textBkColor));
            }

            _ = SendMessage(listView, WmSetRedraw, new nint(1), nint.Zero);
            _ = InvalidateRect(listView, nint.Zero, true);
            _ = UpdateWindow(listView);
        }

        BroadcastSettingChange();
    }

    // Backward compatibility for older rename-based mode.
    public void RestoreDesktopShortcutLabels(IReadOnlyCollection<ShortcutRenameMap>? maps)
    {
        if (maps is null || maps.Count == 0)
        {
            return;
        }

        foreach (var map in maps)
        {
            if (string.IsNullOrWhiteSpace(map.HiddenPath) || string.IsNullOrWhiteSpace(map.OriginalPath))
            {
                continue;
            }

            if (!File.Exists(map.HiddenPath))
            {
                continue;
            }

            var restorePath = map.OriginalPath;
            if (File.Exists(restorePath))
            {
                var dir = Path.GetDirectoryName(restorePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                var baseName = Path.GetFileNameWithoutExtension(restorePath);
                var extension = Path.GetExtension(restorePath);
                var suffix = 1;

                do
                {
                    restorePath = Path.Combine(dir, $"{baseName} (restored {suffix}){extension}");
                    suffix++;
                }
                while (File.Exists(restorePath));
            }

            File.Move(map.HiddenPath, restorePath);
        }
    }

    private static List<nint> FindDesktopListViews()
    {
        var listViews = new List<nint>();

        var progman = FindWindow("Progman", null);
        if (progman != nint.Zero)
        {
            _ = SendMessageTimeout(progman, ProgmanSpawnWorker, nint.Zero, nint.Zero, 0, 1000, out _);
        }

        _ = EnumWindows((topHwnd, _) =>
        {
            var defView = FindWindowEx(topHwnd, nint.Zero, "SHELLDLL_DefView", null);
            if (defView == nint.Zero)
            {
                return true;
            }

            nint childAfter = nint.Zero;
            while (true)
            {
                var lv = FindWindowEx(defView, childAfter, "SysListView32", null);
                if (lv == nint.Zero)
                {
                    break;
                }

                if (!listViews.Contains(lv))
                {
                    listViews.Add(lv);
                }

                childAfter = lv;
            }

            return true;
        }, nint.Zero);

        if (listViews.Count == 0)
        {
            var shellView = FindWindow("SHELLDLL_DefView", null);
            if (shellView != nint.Zero)
            {
                var lv = FindWindowEx(shellView, nint.Zero, "SysListView32", null);
                if (lv != nint.Zero)
                {
                    listViews.Add(lv);
                }
            }
        }

        return listViews;
    }

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint FindWindow(string lpClassName, string? lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint FindWindowEx(nint parent, nint childAfter, string className, string? windowTitle);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint SendMessage(nint hWnd, int msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern nint SendMessageTimeout(nint hWnd, int msg, nint wParam, nint lParam, uint fuFlags, uint uTimeout, out nint lpdwResult);

    [DllImport("user32.dll")]
    private static extern bool InvalidateRect(nint hWnd, nint lpRect, bool bErase);

    [DllImport("user32.dll")]
    private static extern bool UpdateWindow(nint hWnd);

    [DllImport("user32.dll")]
    private static extern int GetSysColor(int nIndex);

    private static void SetListViewShadowEnabled(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(ExplorerAdvancedKeyPath, writable: true)
            ?? throw new InvalidOperationException("Cannot access Explorer advanced settings.");
        key.SetValue(ListviewShadowValue, enabled ? 1 : 0, RegistryValueKind.DWord);
    }

    private static void BroadcastSettingChange()
    {
        _ = SendMessageTimeout(
            new nint((int)HwndBroadcast),
            WmSettingChange,
            nint.Zero,
            nint.Zero,
            SmtoAbortIfHung,
            1000,
            out _);
    }
}
