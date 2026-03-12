using System.Text.Json;
using System.Text.Json.Serialization;

namespace AccessCli;

/// <summary>
/// GitHub Releases APIを使った自動更新チェック。
/// 起動時に非同期で最新バージョンを確認し、新しいものがあれば案内する。
/// </summary>
public static class Updater
{
    private const string CurrentVersion = "0.2.0";
    private const string ApiUrl = "https://api.github.com/repos/YOUR_ORG/access-cli/releases/latest";
    private static readonly TimeSpan CheckInterval = TimeSpan.FromHours(24);
    private static readonly string CacheFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "access-cli", "update-check.json");

    public static async Task CheckAsync()
    {
        try
        {
            // Rate limit: check at most once per day
            if (!ShouldCheck()) return;

            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "access-cli/" + CurrentVersion);
            // Use system proxy (configured via environment variable)
            client.Timeout = TimeSpan.FromSeconds(5);

            var json = await client.GetStringAsync(ApiUrl);
            var release = JsonSerializer.Deserialize<GhRelease>(json);
            if (release?.TagName is null) return;

            string latest = release.TagName.TrimStart('v');
            SaveCheckTime();

            if (IsNewer(latest, CurrentVersion))
            {
                Console.Error.WriteLine();
                Console.Error.WriteLine($"[UPDATE] 新しいバージョンが利用可能です: {latest} (現在: {CurrentVersion})");
                Console.Error.WriteLine($"  更新: dotnet tool update -g access-cli");
                Console.Error.WriteLine();
            }
        }
        catch
        {
            // 更新チェックの失敗は無視する
        }
    }

    private static bool ShouldCheck()
    {
        try
        {
            if (!File.Exists(CacheFile)) return true;
            var json = File.ReadAllText(CacheFile);
            var cache = JsonSerializer.Deserialize<CheckCache>(json);
            if (cache is null) return true;
            return DateTime.UtcNow - cache.LastCheck > CheckInterval;
        }
        catch { return true; }
    }

    private static void SaveCheckTime()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(CacheFile)!);
            var json = JsonSerializer.Serialize(new CheckCache(DateTime.UtcNow));
            File.WriteAllText(CacheFile, json);
        }
        catch { }
    }

    private static bool IsNewer(string latest, string current)
    {
        if (Version.TryParse(latest, out var lv) && Version.TryParse(current, out var cv))
            return lv > cv;
        return string.CompareOrdinal(latest, current) > 0;
    }

    private record GhRelease([property: JsonPropertyName("tag_name")] string? TagName);
    private record CheckCache(DateTime LastCheck);
}
