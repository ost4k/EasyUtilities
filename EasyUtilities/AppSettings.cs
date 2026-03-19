using EasyUtilities.Models;

namespace EasyUtilities;

public sealed class AppSettings
{
    public bool StartWithWindows { get; set; }
    public bool MinimizeOnMiddleClickTitle { get; set; }
    public bool HideShortcutArrows { get; set; }
    public bool HideDesktopShortcutArrows { get; set; }
    public bool HideShortcutNames { get; set; }
    public bool FileListExportEnabled { get; set; }
    public string FileListExportMode { get; set; } = "clipboard";
    public bool AlwaysOnTopEnabled { get; set; } = true;
    public bool OcrEnabled { get; set; } = true;
    public bool RestartExplorerHotkeyEnabled { get; set; } = true;
    public int DesktopLabelRestoreColor { get; set; }
    public int DesktopLabelRestoreBkColor { get; set; } = -1;
    public int DesktopLabelRestoreShadowEnabled { get; set; } = -1;
    public string LanguageMode { get; set; } = "system";
    public string ThemeMode { get; set; } = "system";
    public string HotkeyToggleTopMost { get; set; } = "Ctrl+Alt+T";
    public string HotkeyOcr { get; set; } = "Ctrl+Alt+O";
    public string HotkeyShowHideApp { get; set; } = "Ctrl+Alt+M";
    public List<ShortcutRenameMap> ShortcutRenameMaps { get; set; } = [];
}
