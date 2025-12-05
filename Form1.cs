using Microsoft.Win32;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EXIF_Remover
{
    public partial class Form1 : Form
    {
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunValueName = "EXIFRemover";
        private bool _trayBalloonShown;
        private readonly string[] _startupFiles;
        private readonly bool _invokedFromShell;
        private bool _startupMinimized;
        private string buildType = ""; // Debug|Release
        private string versionThis = "";
        private string _rootBasePath;

        // Session-scoped undo manager
        private readonly UndoSession _undo;

        // Track clickable restore tokens in the log
        private readonly List<RestoreToken> _restoreTokens = new List<RestoreToken>();
        private bool _logMouseHandlersAttached;

        // EXIF tag for DateTimeOriginal
        private const int ExifTagDateTimeOriginal = 0x9003;

        private const string ContextMenuName = "Remove EXIF data";
        private static readonly string[] ContextExtensions = { ".jpg", ".jpeg", ".png" };

        private const int WM_COPYDATA = 0x004A;

        private readonly List<string> _pendingFiles = new List<string>();
        private readonly Timer _batchTimer;
        private const int BatchDelayMs = 150;
        private bool _batchProcessing;

        [StructLayout(LayoutKind.Sequential)]
        private struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;
            public IntPtr lpData;
        }

        public Form1() : this(Array.Empty<string>()) { }

        public Form1(string[] files)
        {
            _startupFiles = files ?? Array.Empty<string>();
            _invokedFromShell = _startupFiles.Length > 0;

            // Clean the entire global temp root BEFORE any session backup is created
            CleanGlobalTempRoot();

            // IMPORTANT: instantiate the session-scoped undo manager
            _undo = new UndoSession();

            InitializeComponent();

            this.Text = "EXIF Remover";

            // Ensure log mouse handlers are attached once
            AttachLogClickHandlers();

            linkLabel1.LinkClicked += (s, e) => {
                try { Process.Start("https://github.com/HovKlan-DH/EXIF-Remover"); }
                catch (Exception ex) { LogErrorEarly("OpenLink", ex); }
            };

            _startupMinimized = Environment.GetCommandLineArgs()
                .Skip(1)
                .Any(a => string.Equals(a, "--minimized", StringComparison.OrdinalIgnoreCase));

            try
            {
                using (var rk = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false))
                {
                    var val = rk?.GetValue(RunValueName) as string;
                    chkStartup.Checked = !string.IsNullOrWhiteSpace(val);
                }
            }
            catch { }

            _batchTimer = new Timer { Interval = BatchDelayMs };
            _batchTimer.Tick += (s, e) => FlushPendingBatch();

            try
            {
                RegisterContextMenus();
                ValidateContextMenus();
            }
            catch (Exception ex)
            {
                LogErrorEarly("ContextMenuReg", ex);
            }

            this.FormClosed += (s, e) => { try { UnregisterContextMenus(); } catch { } };

            this.Load += (s, e) =>
            {
                if (_startupMinimized)
                {
                    WindowState = FormWindowState.Minimized;
                    BeginInvoke(new Action(() => MinimizeToTray()));
                }

                var launchFiles = _startupFiles;
                if (launchFiles.Length == 0)
                {
                    var envArgs = Environment.GetCommandLineArgs();
                    var extra = envArgs != null && envArgs.Length > 1
                        ? envArgs
                            .Skip(1)
                            .Where(a => !string.Equals(a, "--minimized", StringComparison.OrdinalIgnoreCase))
                            // Accept directories too
                            .Where(a => File.Exists(a) || Directory.Exists(a))
                            .ToArray()
                        : Array.Empty<string>();
                    launchFiles = extra;
                }

                if (launchFiles.Length > 0)
                    QueueFiles(launchFiles);
            };

            this.FormClosed += (s, e) =>
            {
                try
                {
                    UnregisterContextMenus();
                }
                catch { }
                try
                {
                    _undo.Cleanup();
                }
                catch { }
            };

            GetAssemblyVersion();
            label1.Text = $"Version {versionThis}";
            label1.Location = new System.Drawing.Point(
                this.ClientSize.Width - label1.Width - 10,
                label1.Location.Y);
        }

        private static void CleanGlobalTempRoot()
        {
            try
            {
                var root = Path.Combine(Path.GetTempPath(), "EXIFRemover");
                if (Directory.Exists(root))
                {
                    // Attempt to remove everything under EXIFRemover (including SessionBackups)
                    Directory.Delete(root, recursive: true);
                }
            }
            catch
            {
                // Swallow exceptions: if some files are locked, we just proceed.
                // Optionally could retry on exit, but user requested only at launch.
            }
        }

        private void GetAssemblyVersion()
        {
            try
            {
                Assembly assemblyInfo = Assembly.GetExecutingAssembly();
                string assemblyVersion = FileVersionInfo.GetVersionInfo(assemblyInfo.Location).FileVersion;
                string year = assemblyVersion.Substring(0, 4);
                string month = assemblyVersion.Substring(5, 2);
                string day = assemblyVersion.Substring(8, 2);
                string rev = assemblyVersion.Substring(11);
                switch (month)
                {
                    case "01": month = "January"; break;
                    case "02": month = "February"; break;
                    case "03": month = "March"; break;
                    case "04": month = "April"; break;
                    case "05": month = "May"; break;
                    case "06": month = "June"; break;
                    case "07": month = "July"; break;
                    case "08": month = "August"; break;
                    case "09": month = "September"; break;
                    case "10": month = "October"; break;
                    case "11": month = "November"; break;
                    case "12": month = "December"; break;
                    default: month = "Unknown"; break;
                }
                day = day.TrimStart(new Char[] { '0' });
                day = day.TrimEnd(new Char[] { '.' });
                string date = year + "-" + month + "-" + day;

                rev = "(rev. " + rev + ")";
                rev = buildType == "Debug" ? rev : "";
                string buildTypeTmp = buildType == "Debug" ? "# DEVELOPMENT " : "";

                versionThis = (date + " " + buildTypeTmp + rev).Trim();
            }
            catch { }
        }

        private void MinimizeToTray()
        {
            try
            {
                if (notifyIcon == null) return;
                notifyIcon.Visible = true;
                ShowInTaskbar = false;
                Hide();

                if (!_startupMinimized && !_trayBalloonShown)
                {
                    _trayBalloonShown = true;
                    try
                    {
                        notifyIcon.BalloonTipText = "Minimized to tray. Double-click to restore.";
                        notifyIcon.ShowBalloonTip(2000);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
                MinimizeToTray();
        }

        private void NotifyIcon_DoubleClick(object sender, EventArgs e)
        {
            RestoreFromTray();
        }

        private void RestoreFromTray()
        {
            try
            {
                if (notifyIcon != null) notifyIcon.Visible = false;
                ShowInTaskbar = true;
                Show();
                WindowState = FormWindowState.Normal;
                try { Activate(); } catch { }
            }
            catch { }
        }

        private void ChkStartup_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                using (var rk = Registry.CurrentUser.CreateSubKey(RunKeyPath))
                {
                    if (rk == null) return;
                    if (chkStartup.Checked)
                    {
                        var exe = Application.ExecutablePath;
                        rk.SetValue(RunValueName, $"\"{exe}\" --minimized", RegistryValueKind.String);
                    }
                    else
                    {
                        rk.DeleteValue(RunValueName, throwOnMissingValue: false);
                    }
                }
            }
            catch (Exception ex)
            {
                LogErrorEarly("StartupReg", ex);
                try
                {
                    using (var rk = Registry.CurrentUser.OpenSubKey(RunKeyPath, writable: false))
                    {
                        var val = rk?.GetValue(RunValueName) as string;
                        chkStartup.Checked = !string.IsNullOrWhiteSpace(val);
                    }
                }
                catch { }
            }
        }

        private void TrayShowMenuItem_Click(object sender, EventArgs e)
        {
            RestoreFromTray();
        }

        private void TrayExitMenuItem_Click(object sender, EventArgs e)
        {
            try { notifyIcon.Visible = false; } catch { }
            Close();
        }

        // NEW: Expands directories recursively into supported file paths.
        private IEnumerable<string> ExpandPaths(IEnumerable<string> inputs)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in inputs ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(item)) continue;

                if (File.Exists(item))
                {
                    if (seen.Add(item))
                        yield return item;
                    continue;
                }

                if (!Directory.Exists(item)) continue;

                var stack = new Stack<string>();
                stack.Push(item);

                while (stack.Count > 0)
                {
                    var dir = stack.Pop();

                    IEnumerable<string> files;
                    try
                    {
                        files = Directory.EnumerateFiles(dir, "*.*", SearchOption.TopDirectoryOnly);
                    }
                    catch (Exception ex)
                    {
                        LogErrorEarly("DirEnumFiles", ex);
                        continue;
                    }

                    foreach (var f in files)
                    {
                        if (!seen.Contains(f) && IsSupported(f))
                        {
                            seen.Add(f);
                            yield return f;
                        }
                    }

                    IEnumerable<string> subDirs;
                    try
                    {
                        subDirs = Directory.EnumerateDirectories(dir);
                    }
                    catch (Exception ex)
                    {
                        LogErrorEarly("DirEnumDirs", ex);
                        continue;
                    }

                    foreach (var sd in subDirs)
                    {
                        try
                        {
                            var attr = File.GetAttributes(sd);
                            if ((attr & FileAttributes.System) != 0) continue;
                        }
                        catch { }
                        stack.Push(sd);
                    }
                }
            }
        }

        // Attach mouse handlers to enable clicking “[Restore original]”
        private void AttachLogClickHandlers()
        {
            try
            {
                if (_logMouseHandlersAttached || richTextBox1 == null) return;
                richTextBox1.MouseMove += TxtLog_MouseMove;
                richTextBox1.MouseClick += TxtLog_MouseClick;
                richTextBox1.MouseLeave += RichTextBox1_MouseLeave; // reset on leave
                _logMouseHandlersAttached = true;
            }
            catch (Exception ex)
            {
                LogErrorEarly("AttachLogClick", ex);
            }
        }

        private void TxtLog_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (richTextBox1 == null) return;

                // Compute whether we are over a restore token
                int idx = richTextBox1.GetCharIndexFromPosition(e.Location);
                bool overToken = _restoreTokens.Any(t => idx >= t.StartIndex && idx < t.EndIndex);

                // Only update cursor when state changes to avoid flicker
                var desiredCursor = overToken ? Cursors.Hand : Cursors.IBeam;
                if (!ReferenceEquals(richTextBox1.Cursor, desiredCursor))
                {
                    richTextBox1.Cursor = desiredCursor;
                }
            }
            catch { }
        }

        private void TxtLog_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button != MouseButtons.Left || richTextBox1 == null) return;
                int idx = richTextBox1.GetCharIndexFromPosition(e.Location);
                var token = _restoreTokens.FirstOrDefault(t => idx >= t.StartIndex && idx < t.EndIndex);
                if (token != null && !string.IsNullOrWhiteSpace(token.OriginalPath))
                {
                    RestoreFile(token.OriginalPath);
                }
            }
            catch (Exception ex)
            {
                LogErrorEarly("LogClickRestore", ex);
            }
        }

        private void RichTextBox1_MouseLeave(object sender, EventArgs e)
        {
            try
            {
                if (richTextBox1 == null) return;
                // When leaving the log, let other UI controls manage their own cursors.
                // Reset to default Arrow so moving over empty areas won’t “toggle”.
                if (!ReferenceEquals(richTextBox1.Cursor, Cursors.Arrow))
                {
                    richTextBox1.Cursor = Cursors.Arrow;
                }
            }
            catch { }
        }

        private static DateTime? TryGetExifDateTimeOriginal(string filePath)
        {
            try
            {
                using (var img = Image.FromFile(filePath))
                {
                    var prop = img.PropertyItems?.FirstOrDefault(p => p.Id == ExifTagDateTimeOriginal);
                    if (prop == null || prop.Value == null || prop.Value.Length == 0)
                        return null;

                    // EXIF DateTimeOriginal is ASCII "YYYY:MM:DD HH:MM:SS" possibly with trailing null
                    var s = System.Text.Encoding.ASCII.GetString(prop.Value).Trim('\0', ' ', '\r', '\n');
                    if (string.IsNullOrWhiteSpace(s))
                        return null;

                    // Normalize to "yyyy-MM-dd HH:mm:ss" for parsing
                    // EXIF uses colons in the date portion: "YYYY:MM:DD HH:MM:SS"
                    // We'll parse it directly.
                    DateTime dt;
                    if (DateTime.TryParseExact(
                            s,
                            "yyyy:MM:dd HH:mm:ss",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None,
                            out dt))
                    {
                        return dt; // EXIF is local time (no TZ); use as local LastWriteTime
                    }

                    // Fallback attempts (rare variations)
                    if (DateTime.TryParse(s, out dt))
                        return dt;

                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static void SafeSetLastWrite(string path, DateTime dtLocal)
        {
            try
            {
                // Set local time; Windows will translate internally
                File.SetLastWriteTime(path, dtLocal);

                // Also set UTC explicitly to avoid any viewer discrepancies
                var dtUtc = DateTime.SpecifyKind(dtLocal, DateTimeKind.Local).ToUniversalTime();
                File.SetLastWriteTimeUtc(path, dtUtc);
            }
            catch { }
        }



        private bool HasBackup(string originalPath)
        {
            try
            {
                var field = typeof(UndoSession).GetField("_map", BindingFlags.NonPublic | BindingFlags.Instance);
                var dict = field?.GetValue(_undo) as System.Collections.IDictionary;
                if (dict == null) return false;
                return dict.Contains(originalPath);
            }
            catch { return false; }
        }

        // Call this before any destructive write to register the original file's backup.
        private void EnsureSessionBackup(string originalPath)
        {
            try
            {
                if (_undo == null) return; // defensive guard
                _undo.RegisterBackup(originalPath);
            }
            catch (Exception ex)
            {
                LogErrorEarly("UndoBackup", ex);
            }
        }

        // Optional: expose restore helpers (can be wired to UI buttons/menus)
        private void RestoreFile(string path)
        {
            try
            {
                if (_undo.Restore(path))
                    LogLine($"[Restored original] [{GetRelativePath(_rootBasePath, path)}]", bracketTimestamp: false);
                else
                    LogLine($"[Restore not available] [{GetRelativePath(_rootBasePath, path)}]", bracketTimestamp: false);
            }
            catch (Exception ex)
            {
                LogErrorEarly("UndoRestoreFile", ex);
            }
        }

        private string ResolveFromUndoMapBySuffix(string display)
        {
            try
            {
                // Access the undo map via reflection without adding public API
                var backupsField = typeof(UndoSession).GetField("_map", BindingFlags.NonPublic | BindingFlags.Instance);
                var dict = backupsField?.GetValue(_undo) as System.Collections.IDictionary;
                if (dict == null) return null;

                foreach (System.Collections.DictionaryEntry de in dict)
                {
                    var original = de.Key as string;
                    if (string.IsNullOrWhiteSpace(original)) continue;

                    if (string.Equals(Path.GetFileName(original), display, StringComparison.OrdinalIgnoreCase))
                        return original;

                    if (original.EndsWith(display, StringComparison.OrdinalIgnoreCase))
                        return original;
                }
            }
            catch { }
            return null;
        }

        private static string NormalizeDir(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            path = Path.GetFullPath(path.Trim());
            if (!path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                path += Path.DirectorySeparatorChar;
            return path;
        }

        private string ComputeCommonBase(IEnumerable<string> fullPaths)
        {
            var dirs = fullPaths
                .Where(File.Exists)
                .Select(p => Path.GetDirectoryName(p))
                .Where(d => !string.IsNullOrWhiteSpace(d))
                .Select(d => NormalizeDir(d))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (dirs.Count == 0) return null;
            if (dirs.Count == 1) return dirs[0];

            var first = dirs[0];
            var segments = first.Split(Path.DirectorySeparatorChar)
                .Where(s => s.Length > 0)
                .ToList();

            int commonCount = segments.Count;
            for (int i = 1; i < dirs.Count && commonCount > 0; i++)
            {
                var otherSegs = dirs[i].Split(Path.DirectorySeparatorChar).Where(s => s.Length > 0).ToList();
                int j = 0;
                while (j < commonCount && j < otherSegs.Count &&
                       string.Equals(segments[j], otherSegs[j], StringComparison.OrdinalIgnoreCase))
                {
                    j++;
                }
                commonCount = j;
            }

            if (commonCount == 0) return null;

            var common = string.Join(Path.DirectorySeparatorChar.ToString(), segments.Take(commonCount));
            common = NormalizeDir(common);
            return common;
        }

        private string GetRelativePath(string basePath, string fullPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(basePath)) return Path.GetFileName(fullPath);
                basePath = NormalizeDir(basePath);
                var fp = Path.GetFullPath(fullPath);
                if (fp.StartsWith(basePath, StringComparison.OrdinalIgnoreCase))
                {
                    var rel = fp.Substring(basePath.Length);
                    return rel.Length == 0 ? Path.GetFileName(fullPath) : rel;
                }
                return Path.GetFileName(fullPath);
            }
            catch
            {
                return Path.GetFileName(fullPath);
            }
        }

        private void QueueFiles(IEnumerable<string> files)
        {
            var expanded = ExpandPaths(files)
                .Select(f => f.Trim())
                .Where(f => f.Length > 0)
                .ToArray();

            if (expanded.Length == 0) return;

            // Do NOT set _rootBasePath here; it's computed per batch in ProcessBatchAsync.
            lock (_pendingFiles)
            {
                foreach (var f in expanded)
                {
                    if (!_pendingFiles.Contains(f, StringComparer.OrdinalIgnoreCase))
                        _pendingFiles.Add(f);
                }
                _batchTimer.Stop();
                _batchTimer.Start();
            }
        }

        private void FlushPendingBatch()
        {
            List<string> batch;
            lock (_pendingFiles)
            {
                if (_pendingFiles.Count == 0) return;
                batch = _pendingFiles.ToList();
                _pendingFiles.Clear();
                _batchTimer.Stop();
            }

            if (_batchProcessing) return;
            _ = ProcessBatchAsync(batch);
        }

        private async Task ProcessBatchAsync(List<string> files)
        {
            if (files == null || files.Count == 0) return;

            // Recompute base path for this batch so relative paths include subfolders correctly.
            // Use the files list (expanded paths) to derive a common base.
            _rootBasePath = ComputeCommonBase(files);

            _batchProcessing = true;
            try
            {
                var targets = files
                    .Where(File.Exists)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();

                if (targets.Length == 0)
                {
                    LogLine("No valid files.", bracketTimestamp: false);
                    return;
                }

                string plural = targets.Length == 1 ? "file" : "files";
                LogLine($"Processing {targets.Length} {plural}:", bracketTimestamp: false);

                int removedCount = 0, skippedCount = 0, errorCount = 0;

                await Task.Run(() =>
                {
                    foreach (var path in targets)
                    {
                        var displayName = GetRelativePath(_rootBasePath, path);

                        try
                        {
                            var originalSize = SafeFileLength(path);
                            bool isSupported = IsSupported(path);
                            bool changed = false;

                            if (isSupported)
                            {
                                if (IsJpeg(path))
                                    changed = RemoveExifFromJpeg(path);
                                else if (IsPng(path))
                                    changed = RemoveExifFromPng(path);
                            }

                            var newSize = SafeFileLength(path);

                            if (!isSupported)
                            {
                                LogLine($"[Unsupported file] [{displayName}]", bracketTimestamp: false);
                                skippedCount++;
                            }
                            else if (changed)
                            {
//                                LogLine($"[Removed EXIF] [Org={FormatKb(originalSize)}] [New={FormatKb(newSize)}] [{displayName}]", bracketTimestamp: false);
                                LogLine($"[Removed EXIF] [{displayName}]", bracketTimestamp: false);
                                removedCount++;
                            }
                            else
                            {
                                LogLine($"[No EXIF present] [{displayName}]", bracketTimestamp: false);
                                skippedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogLine($"[Error] [{ex.GetType().Name}: {ex.Message}] [{displayName}]", bracketTimestamp: false);
                            errorCount++;
                        }
                    }
                });

                if (_invokedFromShell)
                {
                    try
                    {
                        MessageBox.Show(
                            this,
                            $"Removed: {removedCount}\r\nNo EXIF / Unsupported: {skippedCount}\r\nErrors: {errorCount}",
                            "EXIF Remover",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch { }
                    BeginInvoke(new Action(Close));
                }
            }
            finally
            {
                _batchProcessing = false;
            }
        }

        private static void RegisterContextMenus()
        {
            var exe = Application.ExecutablePath;

            // Image file verbs (SystemFileAssociations, progId, legacy jpegfile/pngfile) => use "%1"
            foreach (var ext in ContextExtensions)
            {
                WriteVerbWithArgs($@"Software\Classes\SystemFileAssociations\{ext}\shell\{ContextMenuName}", exe, "%1");

                try
                {
                    using (var extKey = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{ext}", writable: false))
                    {
                        var progId = extKey?.GetValue(null) as string;
                        if (!string.IsNullOrWhiteSpace(progId))
                        {
                            WriteVerbWithArgs($@"Software\Classes\{progId}\shell\{ContextMenuName}", exe, "%1");
                        }
                    }
                }
                catch { }

                if (string.Equals(ext, ".jpg", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".jpeg", StringComparison.OrdinalIgnoreCase))
                {
                    WriteVerbWithArgs(@"Software\Classes\jpegfile\shell\Remove EXIF data", exe, "%1");
                }
                else if (string.Equals(ext, ".png", StringComparison.OrdinalIgnoreCase))
                {
                    WriteVerbWithArgs(@"Software\Classes\pngfile\shell\Remove EXIF data", exe, "%1");
                }
            }

            // Folder item verb (multi-select folders) => use "%1"
            WriteVerbWithArgs(@"Software\Classes\Directory\shell\Remove EXIF data", exe, "%1");

            // Folder background verb (right-click empty area) => use "%V" (the viewed folder)
            WriteVerbWithArgs(@"Software\Classes\Directory\Background\shell\Remove EXIF data", exe, "%V");
        }

        private static void WriteVerbWithArgs(string basePath, string exe, string argumentToken)
        {
            using (var shellKey = Registry.CurrentUser.CreateSubKey(basePath))
            {
                if (shellKey == null) return;
                shellKey.SetValue(null, ContextMenuName);
                shellKey.SetValue("Icon", exe);
                using (var cmdKey = shellKey.CreateSubKey("command"))
                {
                    if (cmdKey == null) return;
                    // Proper token per verb type
                    cmdKey.SetValue(null, $"\"{exe}\" \"{argumentToken}\"");
                    cmdKey.Flush();
                }
                shellKey.Flush();
            }
        }

        private static void ValidateContextMenus()
        {
            var exe = Application.ExecutablePath;

            // Ensure image file verbs use "%1"
            foreach (var ext in ContextExtensions)
            {
                var paths = new[]
                {
            $@"Software\Classes\SystemFileAssociations\{ext}\shell\{ContextMenuName}\command",
        };

                foreach (var cmdKeyPath in paths)
                {
                    using (var cmdKey = Registry.CurrentUser.OpenSubKey(cmdKeyPath, writable: true))
                    {
                        var val = cmdKey?.GetValue(null) as string;
                        if (cmdKey == null || string.IsNullOrWhiteSpace(val))
                        {
                            using (var fix = Registry.CurrentUser.CreateSubKey(cmdKeyPath))
                            {
                                fix?.SetValue(null, $"\"{exe}\" \"%1\"");
                                fix?.Flush();
                            }
                        }
                    }
                }
            }

            // jpegfile/pngfile legacy keys (optional harden)
            using (var fixJ = Registry.CurrentUser.CreateSubKey(@"Software\Classes\jpegfile\shell\Remove EXIF data\command"))
            { fixJ?.SetValue(null, $"\"{exe}\" \"%1\""); fixJ?.Flush(); }
            using (var fixP = Registry.CurrentUser.CreateSubKey(@"Software\Classes\pngfile\shell\Remove EXIF data\command"))
            { fixP?.SetValue(null, $"\"{exe}\" \"%1\""); fixP?.Flush(); }

            // Folder item verb => "%1"
            using (var fixDir = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Directory\shell\Remove EXIF data\command"))
            { fixDir?.SetValue(null, $"\"{exe}\" \"%1\""); fixDir?.Flush(); }

            // Folder background verb => "%V"
            using (var fixBg = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Directory\Background\shell\Remove EXIF data\command"))
            { fixBg?.SetValue(null, $"\"{exe}\" \"%V\""); fixBg?.Flush(); }
        }

        private static void UnregisterContextMenus()
        {
            // Remove image verbs
            foreach (var ext in ContextExtensions)
            {
                try
                {
                    Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\SystemFileAssociations\{ext}\shell\{ContextMenuName}", false);
                }
                catch { }
            }

            // NEW: Remove folder verbs
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\Directory\shell\Remove EXIF data", false); } catch { }
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\Directory\Background\shell\Remove EXIF data", false); } catch { }
        }

        private void BtnSelectFiles_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog
            {
                Title = "Select image files",
                Filter = "Images|*.jpg;*.jpeg;*.png",
                Multiselect = true
            })
            {
                if (ofd.ShowDialog(this) == DialogResult.OK)
                    QueueFiles(ofd.FileNames);
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths != null && paths.Length > 0)
                QueueFiles(paths);
        }

        private static bool IsSupported(string path)
        {
            var ext = Path.GetExtension(path);
            return IsJpegExt(ext) || IsPngExt(ext);
        }       

        private static bool IsJpeg(string path) => IsJpegExt(Path.GetExtension(path));
        private static bool IsPng(string path) => IsPngExt(Path.GetExtension(path));

        private static bool IsJpegExt(string ext) =>
            ext != null && (ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) ||
                            ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase));

        private static bool IsPngExt(string ext) =>
            ext != null && ext.Equals(".png", StringComparison.OrdinalIgnoreCase);

        private static long SafeFileLength(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                return fi.Exists ? fi.Length : 0L;
            }
            catch
            {
                return 0L;
            }
        }

        private static bool StartsWithAscii(byte[] data, string ascii)
        {
            if (data == null) return false;
            int n = ascii.Length;
            if (data.Length < n) return false;
            for (int i = 0; i < n; i++)
                if (data[i] != (byte)ascii[i]) return false;
            return true;
        }

        private static bool RemoveExifFromJpeg(string filePath)
        {
            // Capture EXIF DateTimeOriginal before we modify the file
            var exifDate = TryGetExifDateTimeOriginal(filePath);

            var tmpPath = filePath + ".exif_stripped_tmp";
            using (var src = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var br = new BinaryReader(src))
            {
                if (br.ReadByte() != 0xFF || br.ReadByte() != 0xD8)
                    return false;

                bool removed = false;
                using (var dst = new FileStream(tmpPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var bw = new BinaryWriter(dst))
                {
                    bw.Write((byte)0xFF);
                    bw.Write((byte)0xD8);

                    while (true)
                    {
                        int markerVal = ReadNextMarker(br);
                        if (markerVal == -1) break;

                        byte marker = (byte)markerVal;

                        if (marker == 0xD9)
                        {
                            bw.Write((byte)0xFF);
                            bw.Write(marker);
                            break;
                        }

                        if (marker == 0xDA)
                        {
                            ushort sosLen = ReadBEUInt16(br);
                            if (sosLen < 2) throw new InvalidDataException("Invalid SOS length");
                            var sosData = br.ReadBytes(sosLen - 2);
                            if (sosData.Length != sosLen - 2) throw new EndOfStreamException();

                            bw.Write((byte)0xFF);
                            bw.Write(marker);
                            WriteBEUInt16(bw, sosLen);
                            bw.Write(sosData);

                            CopyRemainder(br, bw);
                            break;
                        }

                        if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7))
                        {
                            bw.Write((byte)0xFF);
                            bw.Write(marker);
                            continue;
                        }

                        ushort len = ReadBEUInt16(br);
                        if (len < 2) throw new InvalidDataException("Invalid segment length");
                        int dataLen = len - 2;
                        var data = br.ReadBytes(dataLen);
                        if (data.Length != dataLen) throw new EndOfStreamException();

                        bool skip = false;
                        if (marker == 0xE1)
                        {
                            if (dataLen >= 6 &&
                                data[0] == (byte)'E' && data[1] == (byte)'x' && data[2] == (byte)'i' &&
                                data[3] == (byte)'f' && data[4] == 0 && data[5] == 0)
                            {
                                // EXIF APP1 -> remove
                                skip = true;
                            }
                            else if (StartsWithAscii(data, "http://ns.adobe.com/xap/1.0/") ||
                                     StartsWithAscii(data, "http://ns.adobe.com/xmp/extension/"))
                            {
                                // XMP -> remove
                                skip = true;
                            }
                        }

                        if (!skip)
                        {
                            bw.Write((byte)0xFF);
                            bw.Write(marker);
                            WriteBEUInt16(bw, len);
                            bw.Write(data);
                        }
                        else
                        {
                            removed = true;
                        }
                    }
                }

                if (!removed)
                {
                    TryDelete(tmpPath);
                    return false;
                }
            }

            ReplaceFile(tmpPath, filePath);

            // Visual match: set file's LastWriteTime to EXIF DateTimeOriginal if available
            if (exifDate.HasValue)
                SafeSetLastWrite(filePath, exifDate.Value);

            return true;
        }

        private static bool RemoveExifFromPng(string filePath)
        {
            // Attempt to capture EXIF DateTimeOriginal before modify (if present via PropertyItem; often not for PNG)
            var exifDate = TryGetExifDateTimeOriginal(filePath);

            var tmpPath = filePath + ".exif_stripped_tmp";
            using (var src = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var br = new BinaryReader(src))
            {
                var sig = br.ReadBytes(8);
                var pngSig = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
                if (sig.Length != 8 || !ByteSeqEqual(sig, pngSig)) return false;

                bool removed = false;
                using (var dst = new FileStream(tmpPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var bw = new BinaryWriter(dst))
                {
                    bw.Write(sig);

                    while (true)
                    {
                        var lenBytes = br.ReadBytes(4);
                        if (lenBytes.Length == 0) break;
                        if (lenBytes.Length < 4) throw new EndOfStreamException();

                        uint len = ToUInt32BE(lenBytes, 0);
                        var typeBytes = br.ReadBytes(4);
                        if (typeBytes.Length < 4) throw new EndOfStreamException();

                        if (len > int.MaxValue) throw new InvalidDataException("Chunk too large");
                        var data = br.ReadBytes((int)len);
                        if (data.Length != (int)len) throw new EndOfStreamException();

                        var crcBytes = br.ReadBytes(4);
                        if (crcBytes.Length < 4) throw new EndOfStreamException();

                        bool isEXIF = typeBytes[0] == (byte)'e' && typeBytes[1] == (byte)'X' &&
                                      typeBytes[2] == (byte)'I' && typeBytes[3] == (byte)'f';
                        bool isIEND = typeBytes[0] == (byte)'I' && typeBytes[1] == (byte)'E' &&
                                      typeBytes[2] == (byte)'N' && typeBytes[3] == (byte)'D';

                        if (!isEXIF)
                        {
                            bw.Write(lenBytes);
                            bw.Write(typeBytes);
                            bw.Write(data);
                            bw.Write(crcBytes);
                        }
                        else
                        {
                            removed = true;
                        }

                        if (isIEND) break;
                    }
                }

                if (!removed)
                {
                    TryDelete(tmpPath);
                    return false;
                }
            }

            ReplaceFile(tmpPath, filePath);

            // Visual match for PNG if we managed to read EXIF date through PropertyItems (rare)
            if (exifDate.HasValue)
                SafeSetLastWrite(filePath, exifDate.Value);

            return true;
        }

        private static int ReadNextMarker(BinaryReader br)
        {
            int b;
            do
            {
                b = br.BaseStream.ReadByte();
                if (b == -1) return -1;
            } while (b != 0xFF);

            int marker;
            do
            {
                marker = br.BaseStream.ReadByte();
                if (marker == -1) return -1;
            } while (marker == 0xFF);

            return marker;
        }

        private static ushort ReadBEUInt16(BinaryReader br)
        {
            int hi = br.ReadByte();
            int lo = br.ReadByte();
            if (hi < 0 || lo < 0) throw new EndOfStreamException();
            return (ushort)((hi << 8) | lo);
        }

        private static void WriteBEUInt16(BinaryWriter bw, ushort value)
        {
            bw.Write((byte)(value >> 8));
            bw.Write((byte)(value & 0xFF));
        }

        private static void CopyRemainder(BinaryReader br, BinaryWriter bw)
        {
            var buffer = new byte[81920];
            int n;
            while ((n = br.Read(buffer, 0, buffer.Length)) > 0)
                bw.Write(buffer, 0, n);
        }

        private static bool ByteSeqEqual(byte[] a, byte[] b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null || a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            return true;
        }

        private static uint ToUInt32BE(byte[] buffer, int offset)
        {
            return (uint)((buffer[offset] << 24) | (buffer[offset + 1] << 16) | (buffer[offset + 2] << 8) | buffer[offset + 3]);
        }

        private static void ReplaceFile(string source, string destination)
        {
            // Before overwriting, register a session backup of the destination (if it exists).
            // NOTE: This requires access to the Form1 instance; we’ll guard via Application.OpenForms.
            try
            {
                var form = Application.OpenForms.OfType<Form1>().FirstOrDefault();
                if (form != null && File.Exists(destination))
                    form.EnsureSessionBackup(destination);
            }
            catch { }

            DateTime? creationUtc = null;
            DateTime? lastWriteUtc = null;
            DateTime? lastAccessUtc = null;

            try
            {
                if (File.Exists(destination))
                {
                    creationUtc = File.GetCreationTimeUtc(destination);
                    lastWriteUtc = File.GetLastWriteTimeUtc(destination);
                    lastAccessUtc = File.GetLastAccessTimeUtc(destination);

                    var attrs = File.GetAttributes(destination);
                    if ((attrs & FileAttributes.ReadOnly) != 0)
                        File.SetAttributes(destination, attrs & ~FileAttributes.ReadOnly);
                }
            }
            catch { }

            File.Copy(source, destination, true);

            try
            {
                if (creationUtc.HasValue) File.SetCreationTimeUtc(destination, creationUtc.Value);
                if (lastWriteUtc.HasValue) File.SetLastWriteTimeUtc(destination, lastWriteUtc.Value);
                if (lastAccessUtc.HasValue) File.SetLastAccessTimeUtc(destination, lastAccessUtc.Value);
            }
            catch { }

            TryDelete(source);
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        private void LogLine(string text, bool bracketTimestamp)
        {
            try
            {
                var ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var line = bracketTimestamp ? $"[{ts}] {text}" : $"{ts} {text}";
                AppendToTextLog(line + Environment.NewLine);
            }
            catch { }
        }

        private void LogErrorEarly(string tag, Exception ex)
        {
            LogLine($"[Error] [{tag}] [{ex.GetType().Name}: {ex.Message}]", bracketTimestamp: false);
        }

        // Override AppendToTextLog to append clickable “[Restore original]” when a backup exists
        private void AppendToTextLog(string content)
        {
            if (richTextBox1 == null) return;

            if (richTextBox1.InvokeRequired)
            {
                richTextBox1.BeginInvoke(new Action<string>(AppendToTextLog), content);
                return;
            }

            string line = content;
            if (line.EndsWith(Environment.NewLine))
                line = line.Substring(0, line.Length - Environment.NewLine.Length);

            // Try to parse: "<timestamp> [<tag>] [<display>]" capturing prefix, tag, display
            // Prefix includes timestamp and any text before the first bracket
            var m = Regex.Match(line, @"^(?<prefix>.*?)(?:\s*)\[(?<tag>Removed EXIF|No EXIF present|Unsupported file|Error)\]\s*\[(?<display>[^\[\]]+)\]\s*$");
            bool hasDisplay = m.Success;
            string prefix = hasDisplay ? m.Groups["prefix"].Value : null;
            string tag = hasDisplay ? m.Groups["tag"].Value : null;
            string display = hasDisplay ? m.Groups["display"].Value?.Trim() : null;

            bool isRemovedExif = hasDisplay && string.Equals(tag, "Removed EXIF", StringComparison.OrdinalIgnoreCase);

            // Styles
            var baseFont = richTextBox1.Font ?? SystemFonts.DefaultFont;
            var normalFont = baseFont;
            var boldFont = new Font(baseFont, baseFont.Style | FontStyle.Bold);
            var normalColor = richTextBox1.ForeColor;

            // Link style
            Font underlineFont;
            try
            {
                underlineFont = new Font(baseFont, baseFont.Style | FontStyle.Underline);
            }
            catch
            {
                underlineFont = baseFont;
            }
            var linkColor = Color.Blue;
            string tokenText = " [Restore original]";

            // Resolve candidate original path for restore token (only for Removed EXIF)
            string candidate = null;
            if (isRemovedExif && !string.IsNullOrWhiteSpace(display))
            {
                try
                {
                    if (Path.IsPathRooted(display))
                        candidate = display;
                    else if (!string.IsNullOrWhiteSpace(_rootBasePath))
                        candidate = Path.Combine(_rootBasePath, display);
                }
                catch { candidate = null; }

                if (string.IsNullOrWhiteSpace(candidate) || !File.Exists(candidate))
                {
                    candidate = ResolveFromUndoMapBySuffix(display);
                }
            }

            // Begin styled append
            int baseLength = richTextBox1.TextLength;
            richTextBox1.SelectionStart = baseLength;
            richTextBox1.SelectionLength = 0;

            // 1) Append prefix and tag normally
            if (hasDisplay)
            {
                // Prefix (e.g., "2025-12-05 14:00:14")
                richTextBox1.SelectionColor = normalColor;
                richTextBox1.SelectionFont = normalFont;
                richTextBox1.AppendText(prefix?.TrimEnd() ?? string.Empty);

                // Space and [tag]
                if (!string.IsNullOrEmpty(prefix)) richTextBox1.AppendText(" ");
                richTextBox1.AppendText("[");
                richTextBox1.AppendText(tag);
                richTextBox1.AppendText("] ");
                richTextBox1.AppendText("[");
            }
            else
            {
                // Fallback: not in expected format, append entire content plainly
                richTextBox1.SelectionColor = normalColor;
                richTextBox1.SelectionFont = normalFont;
                richTextBox1.AppendText(content);
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.ScrollToCaret();
                return;
            }

            // 2) Append display with folder normal and filename bold
            // display may be like "Folder\Sub\IMG_1234.JPEG" or "IMG_1234.JPEG"
            string fileName = display != null ? Path.GetFileName(display) : string.Empty;
            string folderPart = string.Empty;
            if (!string.IsNullOrEmpty(display) && !string.IsNullOrEmpty(fileName) && display.Length > fileName.Length)
                folderPart = display.Substring(0, display.Length - fileName.Length);

            // Folder part normal (if any)
            richTextBox1.SelectionColor = normalColor;
            richTextBox1.SelectionFont = normalFont;
            if (!string.IsNullOrEmpty(folderPart))
                richTextBox1.AppendText(folderPart);

            // Filename bold
            if (!string.IsNullOrEmpty(fileName))
            {
                richTextBox1.SelectionColor = normalColor;
                richTextBox1.SelectionFont = boldFont;
                richTextBox1.AppendText(fileName);
            }

            // Close display bracket
            richTextBox1.SelectionColor = normalColor;
            richTextBox1.SelectionFont = normalFont;
            richTextBox1.AppendText("]");

            // 3) Optional restore token for Removed EXIF if backup exists
            bool appendedToken = false;
            if (isRemovedExif && !string.IsNullOrWhiteSpace(candidate) && HasBackup(candidate))
            {
                int tokenStart = richTextBox1.TextLength + 1; // space before token
                richTextBox1.AppendText(" ");
                richTextBox1.SelectionStart = richTextBox1.TextLength;
                richTextBox1.SelectionColor = linkColor;
                richTextBox1.SelectionFont = underlineFont;
                richTextBox1.AppendText(tokenText.TrimStart());
                int tokenEnd = richTextBox1.TextLength;

                _restoreTokens.Add(new RestoreToken
                {
                    StartIndex = tokenStart,
                    EndIndex = tokenEnd,
                    OriginalPath = candidate
                });

                appendedToken = true;
            }

            // Newline and scroll
            richTextBox1.SelectionColor = normalColor;
            richTextBox1.SelectionFont = normalFont;
            richTextBox1.AppendText(Environment.NewLine);
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.ScrollToCaret();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_COPYDATA)
            {
                try
                {
                    var cds = (COPYDATASTRUCT)Marshal.PtrToStructure(m.LParam, typeof(COPYDATASTRUCT));
                    if (cds.cbData > 0 && cds.lpData != IntPtr.Zero)
                    {
                        var chars = cds.cbData / 2;
                        var text = Marshal.PtrToStringUni(cds.lpData, Math.Max(chars, 0)) ?? string.Empty;
                        text = text.TrimEnd('\0');

                        var files = text
                            .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => s.Trim('\r'))
                            // Accept both files and directories received from the forwarder
                            .Where(p => File.Exists(p) || Directory.Exists(p))
                            .ToArray();

                        if (files.Length > 0)
                            QueueFiles(files);
                    }
                }
                catch (Exception ex)
                {
                    LogErrorEarly("WM_COPYDATA", ex);
                }

                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }

        // Helper class to track clickable restore tokens
        private sealed class RestoreToken
        {
            public int StartIndex { get; set; }
            public int EndIndex { get; set; }
            public string OriginalPath { get; set; }
        }
    }


    // Session-only Undo manager
    internal sealed class UndoSession
    {
        private readonly string _sessionRoot;
        private readonly ConcurrentDictionary<string, string> _map; // original -> backup

        

        public UndoSession()
        {
            _map = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var baseTemp = Path.Combine(Path.GetTempPath(), "EXIFRemover", "SessionBackups");
            var stamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss_fff");
            _sessionRoot = Path.Combine(baseTemp, stamp);
            try
            {
                Directory.CreateDirectory(_sessionRoot);
            }
            catch { }
        }

        public void RegisterBackup(string originalPath)
        {
            if (string.IsNullOrWhiteSpace(originalPath) || !File.Exists(originalPath))
                return;

            // Avoid duplicate work: if already backed up this exact file version (based on last write time and size),
            // you can skip. For simplicity, re-copy if not present.
            if (_map.ContainsKey(originalPath))
                return;

            var rel = MakeSafeRelative(originalPath);
            var backupPath = Path.Combine(_sessionRoot, rel);

            var backupDir = Path.GetDirectoryName(backupPath);
            if (!string.IsNullOrWhiteSpace(backupDir))
            {
                try { Directory.CreateDirectory(backupDir); } catch { }
            }

            // Copy original to backup location
            try
            {
                File.Copy(originalPath, backupPath, overwrite: true);
                _map[originalPath] = backupPath;
            }
            catch
            {
                // If copy fails, do not track
            }
        }

        public bool Restore(string originalPath)
        {
            if (string.IsNullOrWhiteSpace(originalPath))
                return false;

            string backupPath;
            if (!_map.TryGetValue(originalPath, out backupPath))
                return false;

            if (!File.Exists(backupPath))
                return false;

            try
            {
                // Preserve current file timestamps if needed; simply overwrite with backup
                var attrs = File.GetAttributes(originalPath);
                if ((attrs & FileAttributes.ReadOnly) != 0)
                    File.SetAttributes(originalPath, attrs & ~FileAttributes.ReadOnly);

                File.Copy(backupPath, originalPath, overwrite: true);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public int RestoreFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return 0;

            var normFolder = NormalizeDir(folderPath);
            int count = 0;

            foreach (var kvp in _map.ToArray())
            {
                var original = kvp.Key;
                if (original.StartsWith(normFolder, StringComparison.OrdinalIgnoreCase))
                {
                    if (Restore(original))
                        count++;
                }
            }

            return count;
        }

        public void Cleanup()
        {
            try
            {
                if (Directory.Exists(_sessionRoot))
                    Directory.Delete(_sessionRoot, recursive: true);
            }
            catch { }
        }

        private static string MakeSafeRelative(string fullPath)
        {
            // Create a path under the session root mirroring the drive and directory.
            // E.g. "C:\Images\A\B.jpg" -> "C\Images\A\B.jpg"
            try
            {
                var fp = Path.GetFullPath(fullPath);
                var root = Path.GetPathRoot(fp) ?? string.Empty;
                var drivePart = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                                    .Replace(":", string.Empty);
                var withoutRoot = fp.Substring(root.Length);
                var combined = Path.Combine(drivePart, withoutRoot);
                return combined;
            }
            catch
            {
                // Fallback: use file name only
                return Path.GetFileName(fullPath);
            }
        }

        private static string NormalizeDir(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            path = Path.GetFullPath(path.Trim());
            if (!path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
                path += Path.DirectorySeparatorChar;
            return path;
        }
    }

}