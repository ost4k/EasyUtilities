using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace EasyUtilities.TrayHost;

public sealed class TrayIconHost : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly ToolStripMenuItem _openMenuItem;
    private readonly ToolStripMenuItem _exitMenuItem;
    private readonly Icon? _customIcon;

    public event EventHandler? OpenRequested;
    public event EventHandler? ExitRequested;

    public TrayIconHost(string title, string openText, string exitText)
    {
        var menu = new ContextMenuStrip();
        _openMenuItem = new ToolStripMenuItem(openText, null, (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty));
        _exitMenuItem = new ToolStripMenuItem(exitText, null, (_, _) => ExitRequested?.Invoke(this, EventArgs.Empty));
        menu.Items.Add(_openMenuItem);
        menu.Items.Add(_exitMenuItem);

        // Use app icon from build output when available.
        var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico");
        if (File.Exists(iconPath))
        {
            _customIcon = new Icon(iconPath);
        }

        _notifyIcon = new NotifyIcon
        {
            Visible = true,
            Text = title,
            Icon = _customIcon ?? SystemIcons.Application,
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, _) => OpenRequested?.Invoke(this, EventArgs.Empty);
    }

    public void UpdateTexts(string title, string openText, string exitText)
    {
        _notifyIcon.Text = title;
        _openMenuItem.Text = openText;
        _exitMenuItem.Text = exitText;
    }

    public void Dispose()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _customIcon?.Dispose();
    }
}
