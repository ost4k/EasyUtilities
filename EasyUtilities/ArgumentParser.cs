namespace EasyUtilities;

public static class ArgumentParser
{
    private const string LegacyShortcutArrowCleanupArg = "--cleanup-legacy-shortcut-arrows";
    private const string LegacyDesktopShortcutArrowCleanupArg = "--cleanup-legacy-desktop-shortcut-arrows";

    public static bool HasStartMinimized(string[] args) =>
        args.Any(a => string.Equals(a, "--start-minimized", StringComparison.OrdinalIgnoreCase));

    public static bool HasLegacyShortcutArrowCleanup(string[] args) =>
        args.Any(a => string.Equals(a, LegacyShortcutArrowCleanupArg, StringComparison.OrdinalIgnoreCase));

    public static bool HasLegacyDesktopShortcutArrowCleanup(string[] args) =>
        args.Any(a => string.Equals(a, LegacyDesktopShortcutArrowCleanupArg, StringComparison.OrdinalIgnoreCase));

    public static bool TryGetShortcutArrowAction(string[] args, out bool hideArrows)
    {
        hideArrows = false;
        var value = args.FirstOrDefault(a => a.StartsWith("--set-shortcut-arrows=", StringComparison.OrdinalIgnoreCase));
        if (value is null)
        {
            return false;
        }

        var split = value.Split('=', 2);
        if (split.Length != 2)
        {
            return false;
        }

        hideArrows = string.Equals(split[1], "on", StringComparison.OrdinalIgnoreCase);
        return true;
    }

    public static bool TryGetDesktopShortcutArrowAction(string[] args, out bool hideArrows)
    {
        hideArrows = false;
        var value = args.FirstOrDefault(a => a.StartsWith("--set-desktop-shortcut-arrows=", StringComparison.OrdinalIgnoreCase));
        if (value is null)
        {
            return false;
        }

        var split = value.Split('=', 2);
        if (split.Length != 2)
        {
            return false;
        }

        hideArrows = string.Equals(split[1], "on", StringComparison.OrdinalIgnoreCase);
        return true;
    }
}
