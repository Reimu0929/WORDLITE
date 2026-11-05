using System.Globalization;

namespace WordPadLite.WinUI.Services;

/// <summary>
/// Tiny persistent settings store for the unpackaged build. Values become a
/// simple "key=value" text file under %LOCALAPPDATA%\WordPadLite.
/// </summary>
internal static class AppSettings
{
    private static readonly Dictionary<string, string> Values = Load();

    private static string FolderPath => System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WordPadLite");

    private static string FilePath => System.IO.Path.Combine(FolderPath, "settings.ini");

    private static Dictionary<string, string> Load()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (!File.Exists(FilePath))
            {
                return result;
            }

            foreach (var line in File.ReadAllLines(FilePath))
            {
                var separator = line.IndexOf('=');
                if (separator <= 0)
                {
                    continue;
                }

                result[line[..separator].Trim()] = line[(separator + 1)..].Trim();
            }
        }
        catch
        {
            // A missing or unreadable settings file simply means "use defaults".
        }

        return result;
    }

    private static void Save()
    {
        try
        {
            Directory.CreateDirectory(FolderPath);
            File.WriteAllLines(FilePath, Values.Select(pair => $"{pair.Key}={pair.Value}"));
        }
        catch
        {
            // Persisting settings is optional.
        }
    }

    public static string Get(string key, string fallback) =>
        Values.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : fallback;

    public static bool GetBool(string key, bool fallback) =>
        bool.TryParse(Get(key, fallback.ToString()), out var value) ? value : fallback;

    public static double GetDouble(string key, double fallback) =>
        double.TryParse(Get(key, fallback.ToString(CultureInfo.InvariantCulture)), NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
            ? value
            : fallback;

    public static void Set(string key, string value)
    {
        Values[key] = value;
        Save();
    }

    public static void Set(string key, bool value) => Set(key, value ? "true" : "false");

    public static void Set(string key, double value) => Set(key, value.ToString(CultureInfo.InvariantCulture));
}
