using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;

namespace OutlookPSTRecovery.Web;

internal static class Program
{
    private static readonly string SessionToken = Convert.ToHexString(RandomNumberGenerator.GetBytes(32));
    private static readonly object StateLock = new();
    private static readonly Dictionary<string, Candidate> Candidates = new(StringComparer.OrdinalIgnoreCase);
    private static string? SelectedDestination;
    private static readonly string AppData = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookPSTRecovery");
    private static readonly string HistoryPath = Path.Combine(AppData, "destination-history.json");

    [STAThread]
    private static async Task Main()
    {
        var builder = WebApplication.CreateBuilder();
        builder.Logging.ClearProviders();
        builder.WebHost.UseUrls("http://127.0.0.1:0");
        var app = builder.Build();

        app.Use(async (context, next) =>
        {
            context.Response.Headers.CacheControl = "no-store";
            if (context.Request.Path.StartsWithSegments("/api"))
            {
                var expectedOrigin = $"{context.Request.Scheme}://{context.Request.Host}";
                if (!string.Equals(context.Request.Headers.Origin.ToString(), expectedOrigin, StringComparison.OrdinalIgnoreCase))
                {
                    context.Response.StatusCode = StatusCodes.Status403Forbidden;
                    await context.Response.WriteAsJsonAsync(new { error = "許可されていないアクセス元です。" });
                    return;
                }
                if (!CryptographicOperations.FixedTimeEquals(
                    Encoding.UTF8.GetBytes(context.Request.Headers["X-Session-Token"].ToString()),
                    Encoding.UTF8.GetBytes(SessionToken)))
                {
                    context.Response.StatusCode = StatusCodes.Status403Forbidden;
                    await context.Response.WriteAsJsonAsync(new { error = "セッション認証に失敗しました。ページを再読み込みしてください。" });
                    return;
                }
            }
            await next();
        });

        app.MapGet("/", () =>
        {
            using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("OutlookPSTRecovery.Web.wwwroot.index.html");
            if (stream is null) return Results.Problem("画面ファイルを読み込めません。");
            using var reader = new StreamReader(stream);
            var html = reader.ReadToEnd().Replace("__SESSION_TOKEN__", SessionToken, StringComparison.Ordinal);
            return Results.Content(html, "text/html; charset=utf-8");
        });

        app.MapGet("/api/status", () => Results.Json(new { app = "Outlook PST Recovery", outlookRunning = IsOutlookRunning() }));

        app.MapPost("/api/scan", () =>
        {
            var found = ScanPstFiles();
            lock (StateLock)
            {
                Candidates.Clear();
                foreach (var item in found) Candidates[item.Id] = item;
            }
            return Results.Json(new
            {
                candidates = found.Select(x => new {
                    x.Id, x.Name, x.Path, x.Size, lastWrite = x.LastWrite.ToString("yyyy-MM-dd HH:mm:ss"),
                    x.InOneDrive, x.IsOffline, knownDestination = FindKnownDestination(x)
                }),
                count = found.Count,
                note = "検索はこのPCのアクセス可能なユーザーフォルダーを対象とします。すべてのドライブやクラウド上だけのファイルを網羅するものではありません。"
            });
        });

        app.MapPost("/api/select-destination", async () =>
        {
            string? selected = await RunFolderPickerAsync();
            if (!string.IsNullOrWhiteSpace(selected) && Directory.Exists(selected))
            {
                lock (StateLock) SelectedDestination = selected;
            }
            return Results.Json(new { path = selected });
        });

        app.MapPost("/api/open-folder", async (OpenFolderRequest request) =>
        {
            string? folder = null;
            lock (StateLock)
            {
                if (request.Id is not null && Candidates.TryGetValue(request.Id, out var candidate))
                    folder = Path.GetDirectoryName(candidate.Path);
            }
            if (request.Kind == "destination")
            {
                lock (StateLock) folder = SelectedDestination;
                if (folder is null) return Results.BadRequest(new { error = "先に保存先フォルダーを選択してください。" });
            }
            if (folder is null || !Directory.Exists(folder))
                return Results.BadRequest(new { error = "フォルダーが見つかりません。" });
            OpenExplorer(folder);
            await Task.CompletedTask;
            return Results.Json(new { ok = true, path = folder });
        });

        app.MapPost("/api/remember-destination", (RememberRequest request) =>
        {
            Candidate? candidate;
            lock (StateLock) Candidates.TryGetValue(request.Id, out candidate);
            if (candidate is null) return Results.BadRequest(new { error = "対象ファイルを再検索してください。" });
            var destination = request.Destination;
            if (string.IsNullOrWhiteSpace(destination) || !Directory.Exists(destination))
                return Results.BadRequest(new { error = "存在する保存先フォルダーを指定してください。" });
            try
            {
                Directory.CreateDirectory(AppData);
                var history = LoadHistory();
                history.RemoveAll(h => string.Equals(h.Name, candidate.Name, StringComparison.OrdinalIgnoreCase));
                history.Add(new DestinationHistory(candidate.Name, candidate.Size, candidate.LastWriteUtc.Ticks, destination, DateTime.UtcNow));
                File.WriteAllText(HistoryPath, JsonSerializer.Serialize(history, JsonOptions), Encoding.UTF8);
                lock (StateLock) SelectedDestination = destination;
                return Results.Json(new { ok = true, message = "保存先の履歴を記録しました。" });
            }
            catch (Exception ex) { return Results.Problem("保存先履歴を記録できません: " + ex.Message); }
        });

        app.MapPost("/api/backup", async (BackupRequest request) =>
        {
            Candidate? candidate;
            lock (StateLock) Candidates.TryGetValue(request.Id, out candidate);
            if (candidate is null) return Results.BadRequest(new { error = "対象ファイルを再検索してください。" });
            if (IsOutlookRunning()) return Results.BadRequest(new { error = "Outlookを終了してからバックアップしてください。" });
            try
            {
                var source = new FileInfo(candidate.Path);
                if (!source.Exists) return Results.BadRequest(new { error = "元ファイルが見つかりません。再検索してください。" });
                if ((source.Attributes & FileAttributes.Offline) != 0)
                    return Results.BadRequest(new { error = "このファイルはオンラインのみの可能性があります。エクスプローラーで「このデバイス上で常に保持する」を選び、ダウンロード完了後に再検索してください。" });
                var backupDir = Path.Combine(AppData, "Backups", DateTime.Now.ToString("yyyyMMdd-HHmmss"));
                Directory.CreateDirectory(backupDir);
                var destination = Path.Combine(backupDir, candidate.Name);
                if (File.Exists(destination)) destination = Path.Combine(backupDir, Guid.NewGuid().ToString("N") + "_" + candidate.Name);
                await using (var input = new FileStream(candidate.Path, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, true))
                await using (var output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1024 * 1024, true))
                    await input.CopyToAsync(output);
                var sourceHash = await HashFileAsync(candidate.Path);
                var backupHash = await HashFileAsync(destination);
                if (!CryptographicOperations.FixedTimeEquals(Convert.FromHexString(sourceHash), Convert.FromHexString(backupHash)))
                {
                    File.Delete(destination);
                    return Results.Problem("バックアップ検証に失敗したため、作成したコピーを削除しました。");
                }
                return Results.Json(new { ok = true, path = destination, sha256 = sourceHash, message = "バックアップとSHA-256検証が完了しました。元のPSTは変更していません。" });
            }
            catch (Exception ex) { return Results.Problem("バックアップに失敗しました: " + ex.Message); }
        });

        await app.StartAsync();
        var server = app.Services.GetRequiredService<IServer>();
        var addresses = server.Features.Get<IServerAddressesFeature>()?.Addresses;
        var address = addresses?.FirstOrDefault(a => a.StartsWith("http://127.0.0.1:", StringComparison.Ordinal));
        if (address is null) throw new InvalidOperationException("ローカルWebサーバーを開始できませんでした。");
        try { Process.Start(new ProcessStartInfo(address) { UseShellExecute = true }); }
        catch { MessageBox.Show($"ブラウザーを起動できませんでした。次のURLを開いてください。\n{address}", "Outlook PST Recovery"); }
        await app.WaitForShutdownAsync();
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private static List<Candidate> ScanPstFiles()
    {
        var results = new Dictionary<string, Candidate>(StringComparer.OrdinalIgnoreCase);
        foreach (var root in GetSearchRoots().Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (!Directory.Exists(root)) continue;
            var stack = new Stack<string>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                var dir = stack.Pop();
                try
                {
                    var di = new DirectoryInfo(dir);
                    if ((di.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                    if (!string.Equals(dir, root, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(di.Name, "AppData", StringComparison.OrdinalIgnoreCase)) continue;
                    foreach (var file in di.EnumerateFiles("*.pst", SearchOption.TopDirectoryOnly))
                    {
                        try
                        {
                            var id = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(file.FullName.ToUpperInvariant())));
                            var full = file.FullName;
                            results[full] = new Candidate(id, file.Name, full, file.Length, file.LastWriteTime,
                                file.LastWriteTimeUtc, IsOneDrivePath(full), (file.Attributes & FileAttributes.Offline) != 0);
                        }
                        catch { }
                    }
                    foreach (var sub in di.EnumerateDirectories())
                    {
                        if ((sub.Attributes & FileAttributes.ReparsePoint) != 0) continue;
                        if (sub.Name.Equals("AppData", StringComparison.OrdinalIgnoreCase)) continue;
                        stack.Push(sub.FullName);
                    }
                }
                catch { }
            }
        }
        return results.Values.OrderByDescending(x => x.LastWrite).ToList();
    }

    private static IEnumerable<string> GetSearchRoots()
    {
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        yield return profile;
        yield return docs;
        yield return Path.Combine(local, "Microsoft", "Outlook");
        var oneDriveEnv = Environment.GetEnvironmentVariable("OneDrive");
        if (!string.IsNullOrWhiteSpace(oneDriveEnv)) yield return oneDriveEnv;
        foreach (var key in new[] { "OneDriveCommercial", "OneDriveConsumer" })
        {
            var value = Environment.GetEnvironmentVariable(key);
            if (!string.IsNullOrWhiteSpace(value)) yield return value;
        }
    }

    private static bool IsOneDrivePath(string path) =>
        path.Contains("\\OneDrive", StringComparison.OrdinalIgnoreCase) ||
        path.Contains("/OneDrive", StringComparison.OrdinalIgnoreCase);

    private static bool IsOutlookRunning() => Process.GetProcessesByName("OUTLOOK").Length > 0;

    private static string? FindKnownDestination(Candidate c)
    {
        var entry = LoadHistory().LastOrDefault(h => h.Name.Equals(c.Name, StringComparison.OrdinalIgnoreCase)
            && h.Size == c.Size && h.LastWriteUtcTicks == c.LastWriteUtc.Ticks && Directory.Exists(h.Destination));
        return entry?.Destination;
    }

    private static List<DestinationHistory> LoadHistory()
    {
        try { return File.Exists(HistoryPath) ? JsonSerializer.Deserialize<List<DestinationHistory>>(File.ReadAllText(HistoryPath)) ?? new() : new(); }
        catch { return new(); }
    }

    private static async Task<string?> RunFolderPickerAsync()
    {
        var completion = new TaskCompletionSource<string?>();
        var thread = new Thread(() =>
        {
            try
            {
                using var dialog = new FolderBrowserDialog { Description = "PSTの保存先フォルダーを選択してください", UseDescriptionForTitle = true, ShowNewFolderButton = true };
                completion.SetResult(dialog.ShowDialog() == DialogResult.OK ? dialog.SelectedPath : null);
            }
            catch (Exception ex) { completion.SetException(ex); }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        return await completion.Task;
    }

    private static void OpenExplorer(string path)
    {
        var start = new ProcessStartInfo("explorer.exe") { UseShellExecute = true };
        start.ArgumentList.Add(path);
        Process.Start(start);
    }

    private static async Task<string> HashFileAsync(string path)
    {
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, true);
        var hash = await SHA256.HashDataAsync(stream);
        return Convert.ToHexString(hash);
    }

    private sealed record Candidate(string Id, string Name, string Path, long Size, DateTime LastWrite, DateTime LastWriteUtc, bool InOneDrive, bool IsOffline);
    private sealed record DestinationHistory(string Name, long Size, long LastWriteUtcTicks, string Destination, DateTime RememberedUtc);
    private sealed record OpenFolderRequest(string Kind, string? Id);
    private sealed record RememberRequest(string Id, string Destination);
    private sealed record BackupRequest(string Id);
}
