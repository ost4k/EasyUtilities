using System.Text.Json;
using Microsoft.Win32;

namespace EasyUtilities.Services;

public sealed class DesktopIconLayoutService
{
    private const string BackupType = "easy-utilities.desktop-icon-layout";
    private const int BackupVersion = 1;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private static readonly string[] RootPaths =
    [
        @"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags",
        @"Software\Microsoft\Windows\Shell\Bags"
    ];

    public async Task SaveToFileAsync(string filePath)
    {
        var snapshot = CaptureSnapshot();
        await File.WriteAllTextAsync(filePath, JsonSerializer.Serialize(snapshot, JsonOptions));
    }

    public async Task RestoreFromFileAsync(string filePath)
    {
        var json = await File.ReadAllTextAsync(filePath);
        var snapshot = JsonSerializer.Deserialize<DesktopIconLayoutSnapshot>(json, JsonOptions);
        if (snapshot is null ||
            !string.Equals(snapshot.BackupType, BackupType, StringComparison.OrdinalIgnoreCase) ||
            snapshot.Version != BackupVersion ||
            snapshot.Keys.Count == 0)
        {
            throw new InvalidOperationException("Icon layout snapshot is empty or invalid.");
        }

        ApplySnapshot(snapshot);
    }

    private static DesktopIconLayoutSnapshot CaptureSnapshot()
    {
        var snapshot = new DesktopIconLayoutSnapshot
        {
            BackupType = BackupType,
            Version = BackupVersion,
            CreatedAtUtc = DateTime.UtcNow
        };

        using var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
        foreach (var rootPath in RootPaths)
        {
            using var root = hkcu.OpenSubKey(rootPath, writable: false);
            if (root is null)
            {
                continue;
            }

            CollectDesktopKeys(root, rootPath, snapshot.Keys);
        }

        return snapshot;
    }

    private static void CollectDesktopKeys(RegistryKey current, string currentPath, List<RegistryKeySnapshot> output)
    {
        if (current.Name.EndsWith(@"\Desktop", StringComparison.OrdinalIgnoreCase))
        {
            output.Add(CaptureKey(currentPath, current));
        }

        foreach (var subKeyName in current.GetSubKeyNames())
        {
            try
            {
                using var sub = current.OpenSubKey(subKeyName, writable: false);
                if (sub is null)
                {
                    continue;
                }

                CollectDesktopKeys(sub, $"{currentPath}\\{subKeyName}", output);
            }
            catch
            {
                // ignore inaccessible keys
            }
        }
    }

    private static RegistryKeySnapshot CaptureKey(string keyPath, RegistryKey key)
    {
        var snapshot = new RegistryKeySnapshot
        {
            KeyPath = keyPath
        };

        foreach (var valueName in key.GetValueNames())
        {
            try
            {
                var kind = key.GetValueKind(valueName);
                var value = key.GetValue(valueName);
                if (value is null)
                {
                    continue;
                }

                snapshot.Values.Add(RegistryValueSnapshot.From(valueName, kind, value));
            }
            catch
            {
                // skip problematic values
            }
        }

        return snapshot;
    }

    private static void ApplySnapshot(DesktopIconLayoutSnapshot snapshot)
    {
        using var hkcu = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);

        foreach (var keySnapshot in snapshot.Keys)
        {
            if (string.IsNullOrWhiteSpace(keySnapshot.KeyPath))
            {
                continue;
            }

            using var key = hkcu.CreateSubKey(keySnapshot.KeyPath, writable: true);
            if (key is null)
            {
                continue;
            }

            foreach (var existingName in key.GetValueNames())
            {
                key.DeleteValue(existingName, throwOnMissingValue: false);
            }

            foreach (var valueSnapshot in keySnapshot.Values)
            {
                valueSnapshot.Apply(key);
            }
        }
    }

    private sealed class DesktopIconLayoutSnapshot
    {
        public string BackupType { get; set; } = string.Empty;
        public int Version { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public List<RegistryKeySnapshot> Keys { get; set; } = [];
    }

    private sealed class RegistryKeySnapshot
    {
        public string KeyPath { get; set; } = string.Empty;
        public List<RegistryValueSnapshot> Values { get; set; } = [];
    }

    private sealed class RegistryValueSnapshot
    {
        public string Name { get; set; } = string.Empty;
        public RegistryValueKind Kind { get; set; }
        public string? StringValue { get; set; }
        public string[]? MultiStringValue { get; set; }
        public int? DwordValue { get; set; }
        public long? QwordValue { get; set; }
        public string? BinaryBase64 { get; set; }

        public static RegistryValueSnapshot From(string name, RegistryValueKind kind, object value)
        {
            var snapshot = new RegistryValueSnapshot
            {
                Name = name,
                Kind = kind
            };

            switch (kind)
            {
                case RegistryValueKind.String:
                case RegistryValueKind.ExpandString:
                    snapshot.StringValue = value.ToString() ?? string.Empty;
                    break;
                case RegistryValueKind.MultiString:
                    snapshot.MultiStringValue = value as string[] ?? [];
                    break;
                case RegistryValueKind.DWord:
                    snapshot.DwordValue = Convert.ToInt32(value);
                    break;
                case RegistryValueKind.QWord:
                    snapshot.QwordValue = Convert.ToInt64(value);
                    break;
                case RegistryValueKind.Binary:
                    snapshot.BinaryBase64 = Convert.ToBase64String((byte[])value);
                    break;
                default:
                    snapshot.StringValue = value.ToString() ?? string.Empty;
                    snapshot.Kind = RegistryValueKind.String;
                    break;
            }

            return snapshot;
        }

        public void Apply(RegistryKey key)
        {
            switch (Kind)
            {
                case RegistryValueKind.String:
                case RegistryValueKind.ExpandString:
                    key.SetValue(Name, StringValue ?? string.Empty, Kind);
                    break;
                case RegistryValueKind.MultiString:
                    key.SetValue(Name, MultiStringValue ?? [], RegistryValueKind.MultiString);
                    break;
                case RegistryValueKind.DWord:
                    key.SetValue(Name, DwordValue ?? 0, RegistryValueKind.DWord);
                    break;
                case RegistryValueKind.QWord:
                    key.SetValue(Name, QwordValue ?? 0L, RegistryValueKind.QWord);
                    break;
                case RegistryValueKind.Binary:
                    key.SetValue(Name, string.IsNullOrWhiteSpace(BinaryBase64) ? Array.Empty<byte>() : Convert.FromBase64String(BinaryBase64), RegistryValueKind.Binary);
                    break;
                default:
                    key.SetValue(Name, StringValue ?? string.Empty, RegistryValueKind.String);
                    break;
            }
        }
    }
}
