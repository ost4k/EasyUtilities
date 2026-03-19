using System.Runtime.InteropServices;

namespace EasyUtilities.Services;

public sealed class ExplorerFileListExportService
{
    public const string ModeClipboard = "clipboard";
    public const string ModeCsvFile = "csv";

    public FileListExportResult ExportFromActiveExplorer(string mode)
    {
        var folderPath = GetActiveExplorerFolderPath();
        if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
        {
            throw new InvalidOperationException("Active Explorer folder is unavailable.");
        }

        var items = CollectFileSystemItems(folderPath);

        if (string.Equals(mode, ModeCsvFile, StringComparison.OrdinalIgnoreCase))
        {
            var csvPath = BuildCsvFile(folderPath, items);
            return new FileListExportResult
            {
                FolderPath = folderPath,
                Mode = ModeCsvFile,
                CsvPath = csvPath
            };
        }

        var text = BuildClipboardText(items);
        return new FileListExportResult
        {
            FolderPath = folderPath,
            Mode = ModeClipboard,
            ClipboardText = text
        };
    }

    private static string GetActiveExplorerFolderPath()
    {
        var foreground = NativeMethods.GetForegroundWindow();
        if (foreground == nint.Zero)
        {
            throw new InvalidOperationException("No active window.");
        }

        var shellType = Type.GetTypeFromProgID("Shell.Application")
            ?? throw new InvalidOperationException("Shell.Application COM object is not available.");

        object? shellObject = null;
        object? windowsObject = null;

        try
        {
            shellObject = Activator.CreateInstance(shellType);
            if (shellObject is null)
            {
                throw new InvalidOperationException("Cannot create Shell.Application instance.");
            }

            dynamic shell = shellObject;
            windowsObject = shell.Windows();
            dynamic windows = windowsObject;

            var activeHwnd = foreground.ToInt64();
            var count = (int)windows.Count;
            for (var i = 0; i < count; i++)
            {
                object? windowObject = null;
                try
                {
                    windowObject = windows.Item(i);
                    if (windowObject is null)
                    {
                        continue;
                    }

                    dynamic window = windowObject;
                    var hwnd = Convert.ToInt64(window.HWND);
                    if (hwnd != activeHwnd)
                    {
                        continue;
                    }

                    object? documentObject = window.Document;
                    if (documentObject is null)
                    {
                        break;
                    }

                    dynamic document = documentObject;
                    object? folderObject = document.Folder;
                    if (folderObject is null)
                    {
                        break;
                    }

                    dynamic folder = folderObject;
                    object? selfObject = folder.Self;
                    if (selfObject is null)
                    {
                        break;
                    }

                    dynamic self = selfObject;
                    var path = self.Path as string;
                    if (string.IsNullOrWhiteSpace(path))
                    {
                        break;
                    }

                    return path;
                }
                catch
                {
                    // Try next explorer window.
                }
                finally
                {
                    if (windowObject is not null && Marshal.IsComObject(windowObject))
                    {
                        _ = Marshal.ReleaseComObject(windowObject);
                    }
                }
            }
        }
        finally
        {
            if (windowsObject is not null && Marshal.IsComObject(windowsObject))
            {
                _ = Marshal.ReleaseComObject(windowsObject);
            }

            if (shellObject is not null && Marshal.IsComObject(shellObject))
            {
                _ = Marshal.ReleaseComObject(shellObject);
            }
        }

        throw new InvalidOperationException("Active Explorer folder was not found.");
    }

    private static List<FileSystemExportItem> CollectFileSystemItems(string rootPath)
    {
        var items = new List<FileSystemExportItem>();
        var pendingFolders = new Stack<string>();
        pendingFolders.Push(rootPath);

        while (pendingFolders.Count > 0)
        {
            var currentFolder = pendingFolders.Pop();

            string[] childDirs;
            try
            {
                childDirs = Directory.GetDirectories(currentFolder, "*", SearchOption.TopDirectoryOnly)
                    .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            catch
            {
                childDirs = [];
            }

            for (var i = childDirs.Length - 1; i >= 0; i--)
            {
                pendingFolders.Push(childDirs[i]);
            }

            foreach (var dir in childDirs)
            {
                items.Add(CreateDirectoryItem(rootPath, dir));
            }

            string[] childFiles;
            try
            {
                childFiles = Directory.GetFiles(currentFolder, "*", SearchOption.TopDirectoryOnly)
                    .OrderBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            catch
            {
                childFiles = [];
            }

            foreach (var file in childFiles)
            {
                items.Add(CreateFileItem(rootPath, file));
            }
        }

        return items
            .OrderBy(i => i.RelativePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static FileSystemExportItem CreateDirectoryItem(string rootPath, string fullPath)
    {
        var relativePath = Path.GetRelativePath(rootPath, fullPath);
        DateTime? lastWriteUtc = null;

        try
        {
            lastWriteUtc = Directory.GetLastWriteTimeUtc(fullPath);
        }
        catch
        {
            // ignore metadata failure
        }

        return new FileSystemExportItem
        {
            Type = "DIR",
            Name = Path.GetFileName(fullPath),
            Extension = string.Empty,
            RelativePath = EnsureDirectorySuffix(relativePath),
            FullPath = fullPath,
            LastWriteTimeUtc = lastWriteUtc
        };
    }

    private static FileSystemExportItem CreateFileItem(string rootPath, string fullPath)
    {
        long? sizeBytes = null;
        DateTime? lastWriteUtc = null;

        try
        {
            var info = new FileInfo(fullPath);
            if (info.Exists)
            {
                sizeBytes = info.Length;
                lastWriteUtc = info.LastWriteTimeUtc;
            }
        }
        catch
        {
            // ignore metadata failure
        }

        return new FileSystemExportItem
        {
            Type = "FILE",
            Name = Path.GetFileName(fullPath),
            Extension = Path.GetExtension(fullPath),
            RelativePath = Path.GetRelativePath(rootPath, fullPath),
            FullPath = fullPath,
            SizeBytes = sizeBytes,
            LastWriteTimeUtc = lastWriteUtc
        };
    }

    private static string BuildClipboardText(IReadOnlyList<FileSystemExportItem> items)
    {
        if (items.Count == 0)
        {
            return string.Empty;
        }

        return string.Join(Environment.NewLine, items.Select(i => $"[{i.Type}] {i.RelativePath}"));
    }

    private static string BuildCsvFile(string folderPath, IReadOnlyList<FileSystemExportItem> items)
    {
        var fileName = $"easy-utilities-file-list-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
        var csvPath = Path.Combine(folderPath, fileName);

        using var writer = new StreamWriter(csvPath, false, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        writer.WriteLine("Type,RelativePath,Name,Extension,SizeBytes,LastWriteTimeUtc,FullPath");

        foreach (var item in items)
        {
            writer.WriteLine(string.Join(",",
                EscapeCsv(item.Type),
                EscapeCsv(item.RelativePath),
                EscapeCsv(item.Name),
                EscapeCsv(item.Extension),
                item.SizeBytes?.ToString() ?? string.Empty,
                EscapeCsv(item.LastWriteTimeUtc?.ToString("o") ?? string.Empty),
                EscapeCsv(item.FullPath)));
        }

        return csvPath;
    }

    private static string EnsureDirectorySuffix(string path)
    {
        return path.EndsWith(Path.DirectorySeparatorChar) || path.EndsWith(Path.AltDirectorySeparatorChar)
            ? path
            : path + Path.DirectorySeparatorChar;
    }

    private static string EscapeCsv(string value)
    {
        if (value.Contains('"'))
        {
            value = value.Replace("\"", "\"\"");
        }

        if (value.Contains(',') || value.Contains('"') || value.Contains('\r') || value.Contains('\n'))
        {
            return $"\"{value}\"";
        }

        return value;
    }
}

internal sealed class FileSystemExportItem
{
    public string Type { get; set; } = "FILE";
    public string RelativePath { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Extension { get; set; } = string.Empty;
    public long? SizeBytes { get; set; }
    public DateTime? LastWriteTimeUtc { get; set; }
    public string FullPath { get; set; } = string.Empty;
}

public sealed class FileListExportResult
{
    public string FolderPath { get; set; } = string.Empty;
    public string Mode { get; set; } = ExplorerFileListExportService.ModeClipboard;
    public string? ClipboardText { get; set; }
    public string? CsvPath { get; set; }
}
