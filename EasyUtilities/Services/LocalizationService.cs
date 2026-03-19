using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Windows.ApplicationModel.Resources;

namespace EasyUtilities.Services;

public static class LocalizationService
{
    private const string LanguageSystem = "system";
    private const string LanguageUk = "uk-UA";
    private const string LanguageEn = "en-US";

    private static readonly object Sync = new();
    private static readonly Dictionary<string, string> En = new(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, string> Uk = new(StringComparer.OrdinalIgnoreCase);
    private static ResourceManager? _resourceManager;
    private static ResourceMap? _resourceMap;
    private static bool _resourceMapInitialized;

    private static string _languageMode = LanguageSystem;
    private static bool _loaded;

    public static void SetLanguageMode(string mode)
    {
        lock (Sync)
        {
            _languageMode = NormalizeLanguageMode(mode);
            try
            {
                EnsureLoaded();
            }
            catch
            {
                // Keep app running even if local resource files are malformed.
            }
        }
    }

    public static void Reload()
    {
        lock (Sync)
        {
            _loaded = false;
            En.Clear();
            Uk.Clear();
            try
            {
                EnsureLoaded();
            }
            catch
            {
                // Keep app running even if local resource files are malformed.
            }
        }
    }

    public static string Get(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return string.Empty;
        }

        lock (Sync)
        {
            try
            {
                EnsureLoaded();
            }
            catch
            {
                return key;
            }

            var active = ResolveActiveDictionary();
            if (active.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            if (En.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            var packagedValue = TryGetFromPackagedResources(key);
            if (!string.IsNullOrWhiteSpace(packagedValue) &&
                !string.Equals(packagedValue, key, StringComparison.OrdinalIgnoreCase))
            {
                return packagedValue;
            }
        }

        return key;
    }

    private static void EnsureLoaded()
    {
        EnsureResourceMap();

        if (_loaded)
        {
            return;
        }

        try
        {
            LoadLanguageFile(LanguageEn, En);
        }
        catch
        {
            En.Clear();
        }

        try
        {
            LoadLanguageFile(LanguageUk, Uk);
        }
        catch
        {
            Uk.Clear();
        }

        _loaded = true;
    }

    private static void EnsureResourceMap()
    {
        if (_resourceMapInitialized)
        {
            return;
        }

        _resourceMapInitialized = true;

        try
        {
            _resourceManager = new ResourceManager();
            _resourceMap = _resourceManager.MainResourceMap.TryGetSubtree("Resources") ?? _resourceManager.MainResourceMap;
        }
        catch
        {
            _resourceManager = null;
            _resourceMap = null;
        }
    }

    private static Dictionary<string, string> ResolveActiveDictionary()
    {
        if (string.Equals(_languageMode, LanguageSystem, StringComparison.OrdinalIgnoreCase))
        {
            var systemLanguage = CultureInfo.CurrentUICulture.Name;
            if (systemLanguage.StartsWith("uk", StringComparison.OrdinalIgnoreCase) && Uk.Count > 0)
            {
                return Uk;
            }

            // Default fallback for unsupported system locales is English.
            return En;
        }

        if (_languageMode.StartsWith("uk", StringComparison.OrdinalIgnoreCase) && Uk.Count > 0)
        {
            return Uk;
        }

        return En;
    }

    private static string NormalizeLanguageMode(string? mode)
    {
        if (string.IsNullOrWhiteSpace(mode))
        {
            return LanguageSystem;
        }

        if (mode.StartsWith("uk", StringComparison.OrdinalIgnoreCase))
        {
            return LanguageUk;
        }

        if (mode.StartsWith("en", StringComparison.OrdinalIgnoreCase))
        {
            return LanguageEn;
        }

        return LanguageSystem;
    }

    private static void LoadLanguageFile(string language, Dictionary<string, string> destination)
    {
        var fullPath = Path.Combine(AppContext.BaseDirectory, "Strings", language, "Resources.resw");
        if (!File.Exists(fullPath))
        {
            return;
        }

        var xml = LoadReswWithEncodingFallback(fullPath);
        var dataElements = xml.Root?.Elements("data") ?? [];
        foreach (var data in dataElements)
        {
            var name = data.Attribute("name")?.Value;
            var value = data.Element("value")?.Value;
            if (string.IsNullOrWhiteSpace(name) || value is null)
            {
                continue;
            }

            destination[name] = value;
        }
    }

    private static XDocument LoadReswWithEncodingFallback(string fullPath)
    {
        var bytes = File.ReadAllBytes(fullPath);

        // Preferred path: valid UTF-8.
        var utf8Text = SanitizeXmlText(Encoding.UTF8.GetString(bytes));
        if (!utf8Text.Contains('\uFFFD'))
        {
            return XDocument.Parse(utf8Text);
        }

        // Fallback for legacy/accidental ANSI(cp1251) encoded .resw files.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var cp1251 = SanitizeXmlText(Encoding.GetEncoding(1251).GetString(bytes));
        return XDocument.Parse(cp1251);
    }

    private static string SanitizeXmlText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        var sanitized = text;

        if (sanitized[0] == '\uFEFF')
        {
            sanitized = sanitized[1..];
        }

        // Handles cases where UTF-8 BOM bytes were decoded using a non-UTF encoding.
        if (sanitized.StartsWith("ï»¿", StringComparison.Ordinal) ||
            sanitized.StartsWith("п»ї", StringComparison.Ordinal))
        {
            sanitized = sanitized[3..];
        }

        return sanitized;
    }

    private static string? TryGetFromPackagedResources(string key)
    {
        if (_resourceManager is null || _resourceMap is null)
        {
            return null;
        }

        try
        {
            var context = _resourceManager.CreateResourceContext();
            if (!string.Equals(_languageMode, LanguageSystem, StringComparison.OrdinalIgnoreCase))
            {
                context.QualifierValues["language"] = _languageMode;
            }

            foreach (var candidateName in BuildPackagedCandidates(key))
            {
                var candidate = _resourceMap.TryGetValue(candidateName, context);
                var value = candidate?.ValueAsString;
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            return null;
        }
        catch
        {
            return null;
        }
    }

    private static IEnumerable<string> BuildPackagedCandidates(string key)
    {
        var slash = key.Replace('.', '/');

        return new[]
        {
            key,
            slash,
            $"Resources/{key}",
            $"Resources/{slash}"
        }.Distinct(StringComparer.OrdinalIgnoreCase);
    }
}
