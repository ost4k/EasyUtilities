using System.Runtime.InteropServices;
using System.Text;

namespace EasyUtilities;

internal static class NativeMethods
{
    public const int GwlpWndProc = -4;
    public const int WmClose = 0x0010;
    public const int WmDestroy = 0x0002;
    public const int WmNCMButtonDown = 0x00A7;
    public const int WmNCRButtonUp = 0x00A5;
    public const int WmHotkey = 0x0312;
    public const int HtCaption = 2;
    public const uint SwHide = 0;
    public const uint SwShow = 5;
    public const uint SwMinimize = 6;
    public const uint ModAlt = 0x0001;
    public const uint ModControl = 0x0002;
    public const int HwndTopmost = -1;
    public const int HwndNotTopmost = -2;
    public const uint SwpNomove = 0x0002;
    public const uint SwpNosize = 0x0001;
    public const uint SwpNoactivate = 0x0010;
    public const int GwlExStyle = -20;
    public const uint WsExTopmost = 0x00000008;
    public const uint GaRoot = 2;
    public const int WhMouseLl = 14;
    public const int WmMButtonUp = 0x0208;
    public const uint WmSysCommand = 0x0112;
    public const uint ScMinimize = 0xF020;
    public const uint ScClose = 0xF060;
    public const uint MfBycommand = 0x00000000;
    public const uint MfSeparator = 0x00000800;
    public const uint MfChecked = 0x00000008;
    public const uint MfUnchecked = 0x00000000;
    public const uint TpmReturNcmd = 0x0100;
    public const uint TpmRightButton = 0x0002;
    public const uint MfString = 0x0000;
    public const byte VkLwin = 0x5B;
    public const byte VkShift = 0x10;
    public const byte VkS = 0x53;
    public const uint KeyeventfKeyup = 0x0002;
    public const uint MbIconInformation = 0x00000040;
    public const uint MbYesNo = 0x00000004;
    public const uint MbTopmost = 0x00040000;
    public const int IdYes = 6;

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    public static extern nint SetWindowLongPtr(nint hWnd, int nIndex, nint dwNewLong);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    public static extern nint GetWindowLongPtr(nint hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern nint CallWindowProc(nint lpPrevWndFunc, nint hWnd, uint msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(nint hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(nint hWnd, int id);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(nint hWnd, uint nCmdShow);

    [DllImport("user32.dll")]
    public static extern nint GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(nint hWnd);

    [DllImport("user32.dll")]
    public static extern nint WindowFromPoint(POINT point);

    [DllImport("user32.dll")]
    public static extern nint GetAncestor(nint hWnd, uint gaFlags);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(nint hWnd, out RECT lpRect);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(nint hWnd);

    [DllImport("user32.dll")]
    public static extern nint SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, nint hMod, uint dwThreadId);

    [DllImport("user32.dll")]
    public static extern bool UnhookWindowsHookEx(nint hhk);

    [DllImport("user32.dll")]
    public static extern nint CallNextHookEx(nint hhk, int nCode, nint wParam, nint lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    public static extern nint GetModuleHandle(string? lpModuleName);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(nint hWnd, uint msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    public static extern nint GetSystemMenu(nint hWnd, bool bRevert);

    [DllImport("user32.dll")]
    public static extern nint CreatePopupMenu();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool AppendMenu(nint hMenu, uint uFlags, nuint uIDNewItem, string lpNewItem);

    [DllImport("user32.dll")]
    public static extern int TrackPopupMenu(nint hMenu, uint uFlags, int x, int y, int nReserved, nint hWnd, nint prcRect);

    [DllImport("user32.dll")]
    public static extern bool DestroyMenu(nint hMenu);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool ModifyMenu(nint hMnu, uint uPosition, uint uFlags, nuint uIDNewItem, string lpNewItem);

    [DllImport("user32.dll")]
    public static extern uint CheckMenuItem(nint hmenu, uint uIDCheckItem, uint uCheck);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, nuint dwExtraInfo);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int MessageBox(nint hWnd, string text, string caption, uint type);

    public static void TriggerScreenSnipShortcut()
    {
        keybd_event(VkLwin, 0, 0, 0);
        keybd_event(VkShift, 0, 0, 0);
        keybd_event(VkS, 0, 0, 0);
        keybd_event(VkS, 0, KeyeventfKeyup, 0);
        keybd_event(VkShift, 0, KeyeventfKeyup, 0);
        keybd_event(VkLwin, 0, KeyeventfKeyup, 0);
    }

    public static string GetWindowTitle(nint hwnd)
    {
        var length = GetWindowTextLength(hwnd);
        if (length <= 0)
        {
            return "(без назви)";
        }

        var sb = new StringBuilder(length + 1);
        _ = GetWindowText(hwnd, sb, sb.Capacity);
        return string.IsNullOrWhiteSpace(sb.ToString()) ? "(без назви)" : sb.ToString();
    }

    public delegate nint LowLevelMouseProc(int nCode, nint wParam, nint lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public uint mouseData;
        public uint flags;
        public uint time;
        public nuint dwExtraInfo;
    }
}
