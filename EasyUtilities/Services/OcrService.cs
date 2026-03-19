using System.Diagnostics;
using Windows.ApplicationModel.DataTransfer;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace EasyUtilities.Services;

public sealed class OcrService
{
    public async Task<string?> CaptureAreaAndExtractTextAsync()
    {
        Clipboard.Clear();
        TriggerScreenSnip();
        await Task.Delay(250);

        var deadline = DateTime.UtcNow.AddSeconds(20);
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(180);

            var data = Clipboard.GetContent();
            if (!data.Contains(StandardDataFormats.Bitmap))
            {
                continue;
            }

            var bitmapReference = await data.GetBitmapAsync();
            using var stream = await bitmapReference.OpenReadAsync();
            var decoder = await BitmapDecoder.CreateAsync(stream);
            var softwareBitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);

            var text = await RunOcrAsync(softwareBitmap);
            return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        }

        return null;
    }

    private static void TriggerScreenSnip()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = "ms-screenclip:",
                UseShellExecute = true
            });
            return;
        }
        catch
        {
            // fallback below
        }

        NativeMethods.TriggerScreenSnipShortcut();
    }

    private static async Task<string?> RunOcrAsync(SoftwareBitmap softwareBitmap)
    {
        var available = OcrEngine.AvailableRecognizerLanguages.ToList();
        var preferredLanguageBases = new[] { "uk", "ru", "en" };

        foreach (var languageBase in preferredLanguageBases)
        {
            var recognizerLanguage = ResolveRecognizerLanguage(available, languageBase);
            if (recognizerLanguage is null)
            {
                continue;
            }

            var engine = OcrEngine.TryCreateFromLanguage(recognizerLanguage);
            if (engine is null)
            {
                continue;
            }

            var result = await engine.RecognizeAsync(softwareBitmap);
            if (string.IsNullOrWhiteSpace(result.Text))
            {
                continue;
            }

            var normalized = result.Text.Trim();
            return normalized;
        }

        var fallback = OcrEngine.TryCreateFromUserProfileLanguages();
        if (fallback is null)
        {
            return null;
        }

        var fallbackResult = await fallback.RecognizeAsync(softwareBitmap);
        return string.IsNullOrWhiteSpace(fallbackResult.Text) ? null : fallbackResult.Text.Trim();
    }

    private static Language? ResolveRecognizerLanguage(IEnumerable<Language> available, string languageBase)
    {
        foreach (var language in available)
        {
            if (string.Equals(language.LanguageTag, languageBase, StringComparison.OrdinalIgnoreCase))
            {
                return language;
            }
        }

        foreach (var language in available)
        {
            if (language.LanguageTag.StartsWith($"{languageBase}-", StringComparison.OrdinalIgnoreCase))
            {
                return language;
            }
        }

        return null;
    }
}
