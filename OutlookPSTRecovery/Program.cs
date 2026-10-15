using Microsoft.Win32;
using System.Diagnostics;
using System.Security.Cryptography;

namespace OutlookPSTRecovery;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed record PstCandidate(
    string Path,
    long Size,
    DateTime LastWrite,
    bool InOneDrive,
    bool IsOffline);

internal sealed class MainForm : Form
{
    private readonly Button _analyze = new() { Text = "1. 解析", Width = 120, Height = 36 };
    private readonly Button _repair = new() { Text = "2. 解析結果から修復", Width = 190, Height = 36, Enabled = false };
    private readonly ListView _files = new() { View = View.Details, FullRowSelect = true, GridLines = true, MultiSelect = false };
    private readonly TextBox _details = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill };
    private readonly Label _status = new() { AutoSize = true, Text = "最初に「1. 解析」を実行してください。" };
    private readonly List<PstCandidate> _candidates = new();

    public MainForm()
    {
        Text = "Outlook PST 一括復旧ツール";
        Width = 1120;
        Height = 690;
        MinimumSize = new Size(900, 560);
        StartPosition = FormStartPosition.CenterScreen;

        var intro = new Label
        {
            AutoSize = false,
            Height = 58,
            Dock = DockStyle.Top,
            Padding = new Padding(10),
            Text = "Windows 11 / クラシック版 Outlook / POP 向け\r\n解析は読み取り専用です。修復はPSTをローカルへコピーして検証し、Outlookのメール設定画面を開きます。元ファイルは削除・上書きしません。"
        };

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 48, Padding = new Padding(8, 4, 8, 4) };
        buttons.Controls.Add(_analyze);
        buttons.Controls.Add(_repair);
        buttons.Controls.Add(_status);

        _files.Columns.Add("状態", 110);
        _files.Columns.Add("OneDrive", 80);
        _files.Columns.Add("サイズ", 100);
        _files.Columns.Add("更新日時", 145);
        _files.Columns.Add("PSTファイルの場所", 630);
        _files.Dock = DockStyle.Fill;

        var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 360 };
        split.Panel1.Controls.Add(_files);
        split.Panel2.Controls.Add(_details);

        Controls.Add(split);
        Controls.Add(buttons);
        Controls.Add(intro);

        _analyze.Click += (_, _) => Analyze();
        _repair.Click += (_, _) => RepairSelected();
        _files.SelectedIndexChanged += (_, _) => UpdateSelectionDetails();
    }

    private void Analyze()
    {
        _files.Items.Clear();
        _details.Clear();
        _candidates.Clear();
        _repair.Enabled = false;
        _status.Text = "解析中…";
        UseWaitCursor = true;

        try
        {
            var roots = GetSearchRoots();
            var files = new Dictionary<string, FileInfo>(StringComparer.OrdinalIgnoreCase);
            var messages = new List<string>
            {
                $"PC: {Environment.MachineName}",
                $"ユーザー: {Environment.UserDomainName}\\{Environment.UserName}",
                $"ドキュメント: {GetDocumentsPath()}",
                $"Outlook起動中: {(IsOutlookRunning() ? "はい（修復には終了が必要）" : "いいえ")}",
                "",
                "検索対象:"
            };
            messages.AddRange(roots.Select(p => $"  {(Directory.Exists(p) ? "検出" : "なし")}  {p}"));

            foreach (var root in roots.Where(Directory.Exists))
            {
                try
                {
                    foreach (var file in Directory.EnumerateFiles(root, "*.pst", SearchOption.AllDirectories))
                    {
                        try { files[file] = new FileInfo(file); }
                        catch { /* inaccessible entry */ }
                    }
                }
                catch (Exception ex)
                {
                    messages.Add($"検索警告: {root} — {ex.Message}");
                }
            }

            var oneDriveRoots = GetOneDriveRoots();
            foreach (var file in files.Values.OrderBy(f => f.FullName, StringComparer.OrdinalIgnoreCase))
            {
                var inOneDrive = oneDriveRoots.Any(root =>
                    file.FullName.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar,
                        StringComparison.OrdinalIgnoreCase));
                var offline = (file.Attributes & FileAttributes.Offline) != 0;
                var candidate = new PstCandidate(file.FullName, file.Length, file.LastWriteTime, inOneDrive, offline);
                _candidates.Add(candidate);

                var state = offline ? "要確認: オンラインのみ" : (inOneDrive ? "OneDrive内" : "検出");
                var item = new ListViewItem(state);
                item.SubItems.Add(inOneDrive ? "はい" : "いいえ");
                item.SubItems.Add(FormatBytes(file.Length));
                item.SubItems.Add(file.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss"));
                item.SubItems.Add(file.FullName);
                item.Tag = candidate;
                _files.Items.Add(item);
            }

            var profileNames = GetOutlookProfileNames();
            messages.Add("");
            messages.Add($"検出PST数: {_candidates.Count}");
            messages.Add("Outlookプロファイル名（設定値は変更していません）:");
            messages.AddRange(profileNames.Count > 0 ? profileNames.Select(p => $"  {p}") : new[] { "  検出できませんでした" });
            messages.Add("");
            messages.Add("判定:");
            if (_candidates.Count == 0)
            {
                messages.Add("PSTが標準の検索場所で見つかりません。OneDriveのWeb上・別フォルダー・別ユーザーの保存先を確認してください。");
            }
            else if (_candidates.Any(c => c.InOneDrive))
            {
                messages.Add("OneDrive配下にPSTがあります。同期中のままOutlookで使用しないでください。");
            }
            else
            {
                messages.Add("検索範囲内にPSTが見つかりました。サイズ・更新日時・保存場所を確認し、正しいファイルを選択してください。");
            }
            messages.Add("");
            messages.Add("注意: このツールはOutlookのプロファイルを直接書き換えません。修復後にPOPアカウントの配信先と既定データファイルの確認が必要です。");

            _details.Text = string.Join(Environment.NewLine, messages);
            _repair.Enabled = _candidates.Count > 0;
            _status.Text = $"解析完了：PST {_candidates.Count} 件";
            if (_files.Items.Count > 0) _files.Items[0].Selected = true;
        }
        catch (Exception ex)
        {
            _status.Text = "解析に失敗しました。";
            MessageBox.Show(this, ex.ToString(), "解析エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private void UpdateSelectionDetails()
    {
        if (_files.SelectedItems.Count == 0) return;
        if (_files.SelectedItems[0].Tag is not PstCandidate c) return;

        _details.AppendText(Environment.NewLine + Environment.NewLine +
            "選択中のPST:" + Environment.NewLine +
            $"場所: {c.Path}" + Environment.NewLine +
            $"サイズ: {FormatBytes(c.Size)}" + Environment.NewLine +
            $"更新日時: {c.LastWrite:yyyy/MM/dd HH:mm:ss}" + Environment.NewLine +
            $"OneDrive配下: {(c.InOneDrive ? "はい" : "いいえ")}" + Environment.NewLine +
            $"オンラインのみの可能性: {(c.IsOffline ? "あり（修復前にローカルにダウンロードしてください）" : "属性上は検出されず")}");
    }

    private void RepairSelected()
    {
        if (_files.SelectedItems.Count == 0 || _files.SelectedItems[0].Tag is not PstCandidate candidate)
        {
            MessageBox.Show(this, "先に解析結果からPSTを1件選択してください。");
            return;
        }

        if (IsOutlookRunning())
        {
            MessageBox.Show(this, "Outlookが起動しています。Outlookを通常終了し、タスクマネージャーでOUTLOOK.EXEが終了したことを確認してから再実行してください。元ファイル保護のため、今回は修復を中止しました。",
                "Outlook起動中", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (candidate.IsOffline)
        {
            MessageBox.Show(this, "このファイルはオンラインのみのプレースホルダーの可能性があります。エクスプローラーで右クリックし「このデバイス上で常に保持する」を選び、ダウンロード完了後に解析し直してください。",
                "オンラインファイル", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var answer = MessageBox.Show(this,
            "選択したPSTをローカルの復旧フォルダーへコピーし、SHA-256で照合します。元ファイルは削除・変更しません。\r\n\r\n" +
            candidate.Path + "\r\n\r\n続行しますか？",
            "修復の確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (answer != DialogResult.Yes) return;

        UseWaitCursor = true;
        _repair.Enabled = false;
        try
        {
            var targetDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookPSTRecovery", "Restored");
            Directory.CreateDirectory(targetDir);

            var baseName = Path.GetFileNameWithoutExtension(candidate.Path);
            var extension = Path.GetExtension(candidate.Path);
            var target = Path.Combine(targetDir, Path.GetFileName(candidate.Path));
            if (File.Exists(target))
                target = Path.Combine(targetDir, $"{baseName}-{DateTime.Now:yyyyMMdd-HHmmss}{extension}");
            var counter = 2;
            while (File.Exists(target))
            {
                target = Path.Combine(targetDir, $"{baseName}-{DateTime.Now:yyyyMMdd-HHmmss}-{counter++}{extension}");
            }

            var temp = target + ".copying";
            if (File.Exists(temp))
                throw new IOException($"一時ファイルが既に存在します。手動で確認してください: {temp}");

            File.Copy(candidate.Path, temp, false);
            string sourceHash;
            string copiedHash;
            try
            {
                using var source = File.Open(candidate.Path, FileMode.Open, FileAccess.Read, FileShare.Read);
                sourceHash = Convert.ToHexString(SHA256.HashData(source));
                using var copy = File.Open(temp, FileMode.Open, FileAccess.Read, FileShare.Read);
                copiedHash = Convert.ToHexString(SHA256.HashData(copy));
            }
            catch
            {
                // Keep any temporary file for forensic review rather than deleting user data automatically.
                throw;
            }

            if (!string.Equals(sourceHash, copiedHash, StringComparison.Ordinal))
                throw new IOException("SHA-256が一致しません。一時ファイルを残しました。元ファイルは変更していません。");

            File.Move(temp, target, false);
            var logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookPSTRecovery", "Logs");
            Directory.CreateDirectory(logDir);
            var log = Path.Combine(logDir, $"Repair-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
            File.WriteAllText(log,
                $"日時: {DateTime.Now:yyyy/MM/dd HH:mm:ss}\r\n元ファイル: {candidate.Path}\r\n復旧コピー: {target}\r\nSHA256: {sourceHash}\r\n元ファイル削除: いいえ\r\nOutlookプロファイル変更: いいえ\r\n",
                System.Text.Encoding.UTF8);

            var next = MessageBox.Show(this,
                $"復旧コピーを作成し、SHA-256の一致を確認しました。\r\n\r\n保存先:\r\n{target}\r\n\r\n" +
                "次にOutlookのメール設定を開きます。［データファイル］タブで既存PSTの参照先を確認してください。POPアカウントの配信先を変更する前に、正しいPSTであることを確認してください。\r\n\r\n設定画面を開きますか？",
                "コピー検証成功", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (next == DialogResult.Yes)
            {
                try { Process.Start(new ProcessStartInfo("control.exe", "mlcfg32.cpl") { UseShellExecute = true }); }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"メール設定画面を自動で開けませんでした。コントロールパネルの「Mail (Microsoft Outlook)」を開いてください。\r\n{ex.Message}");
                }
            }

            _status.Text = "復旧コピー作成・検証済み（Outlook設定は未変更）";
            MessageBox.Show(this, $"完了しました。\r\n\r\n復旧コピー: {target}\r\nログ: {log}\r\n\r\n元のPSTは保持されています。Outlookでメールの送受信・過去メールを確認するまで元ファイルを削除しないでください。",
                "修復処理完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _status.Text = "修復に失敗しました。";
            MessageBox.Show(this, ex.ToString(), "修復エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
            _repair.Enabled = _candidates.Count > 0;
        }
    }

    private static bool IsOutlookRunning() => Process.GetProcessesByName("OUTLOOK").Length > 0;

    private static string FormatBytes(long bytes)
    {
        string[] units = ["B", "KB", "MB", "GB", "TB"];
        double size = bytes;
        var i = 0;
        while (size >= 1024 && i < units.Length - 1) { size /= 1024; i++; }
        return $"{size:0.##} {units[i]}";
    }

    private static string GetDocumentsPath()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders");
            var value = key?.GetValue("Personal") as string;
            if (!string.IsNullOrWhiteSpace(value))
            {
                value = Environment.ExpandEnvironmentVariables(value);
                if (Directory.Exists(value)) return value;
            }
        }
        catch { }
        return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    }

    private static List<string> GetOneDriveRoots()
    {
        var roots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in new[] { "OneDrive", "OneDriveCommercial", "OneDriveConsumer" })
        {
            var value = Environment.GetEnvironmentVariable(name);
            if (!string.IsNullOrWhiteSpace(value) && Directory.Exists(value)) roots.Add(Path.GetFullPath(value));
        }

        try
        {
            using var accounts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts");
            if (accounts != null)
            {
                foreach (var subName in accounts.GetSubKeyNames())
                {
                    using var sub = accounts.OpenSubKey(subName);
                    var value = sub?.GetValue("UserFolder") as string;
                    if (!string.IsNullOrWhiteSpace(value) && Directory.Exists(value)) roots.Add(Path.GetFullPath(value));
                }
            }
        }
        catch { }
        return roots.ToList();
    }

    private static List<string> GetSearchRoots()
    {
        var roots = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var docs = GetDocumentsPath();
        roots.Add(Path.Combine(docs, "Outlook Files"));
        roots.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Documents", "Outlook Files"));
        roots.Add(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook"));
        foreach (var od in GetOneDriveRoots())
        {
            roots.Add(Path.Combine(od, "Documents", "Outlook Files"));
            roots.Add(Path.Combine(od, "Outlook Files"));
        }
        return roots.ToList();
    }

    private static List<string> GetOutlookProfileNames()
    {
        var result = new List<string>();
        try
        {
            using var profiles = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles");
            if (profiles != null) result.AddRange(profiles.GetSubKeyNames());
        }
        catch { }
        return result;
    }
}
