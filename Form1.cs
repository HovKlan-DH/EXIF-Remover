using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

            InitializeComponent();
            this.Text = "EXIF Remover";

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
                            .Where(File.Exists)
                            .ToArray()
                        : Array.Empty<string>();
                    launchFiles = extra;
                }

                if (launchFiles.Length > 0)
                    QueueFiles(launchFiles);
            };

            GetAssemblyVersion();
            label1.Text = $"Version {versionThis}";
            label1.Location = new System.Drawing.Point(
                this.ClientSize.Width - label1.Width - 10,
                label1.Location.Y);
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

            // Establish a common base path once (first time files are queued)
            if (_rootBasePath == null)
                _rootBasePath = ComputeCommonBase(expanded);

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
                                LogLine($"[Unsupported] [{displayName}]", bracketTimestamp: false);
                                skippedCount++;
                            }
                            else if (changed)
                            {
                                LogLine($"[Removed EXIF] [Org={FormatKb(originalSize)}] [New={FormatKb(newSize)}] [{displayName}]", bracketTimestamp: false);
                                removedCount++;
                            }
                            else
                            {
                                LogLine($"[No EXIF] [{displayName}]", bracketTimestamp: false);
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
            foreach (var ext in ContextExtensions)
            {
                WriteVerb($@"Software\Classes\SystemFileAssociations\{ext}\shell\{ContextMenuName}", exe);

                try
                {
                    using (var extKey = Registry.CurrentUser.OpenSubKey($@"Software\Classes\{ext}", writable: false))
                    {
                        var progId = extKey?.GetValue(null) as string;
                        if (!string.IsNullOrWhiteSpace(progId))
                        {
                            WriteVerb($@"Software\Classes\{progId}\shell\{ContextMenuName}", exe);
                        }
                    }
                }
                catch { }

                if (string.Equals(ext, ".jpg", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".jpeg", StringComparison.OrdinalIgnoreCase))
                {
                    WriteVerb(@"Software\Classes\jpegfile\shell\Remove EXIF data", exe);
                }
                else if (string.Equals(ext, ".png", StringComparison.OrdinalIgnoreCase))
                {
                    WriteVerb(@"Software\Classes\pngfile\shell\Remove EXIF data", exe);
                }
            }
        }

        private static void WriteVerb(string basePath, string exe)
        {
            using (var shellKey = Registry.CurrentUser.CreateSubKey(basePath))
            {
                if (shellKey == null) return;
                shellKey.SetValue(null, ContextMenuName);
                shellKey.SetValue("Icon", exe);
                using (var cmdKey = shellKey.CreateSubKey("command"))
                {
                    if (cmdKey == null) return;
                    cmdKey.SetValue(null, $"\"{exe}\" \"%1\"");
                    cmdKey.Flush();
                }
                shellKey.Flush();
            }
        }

        private static void ValidateContextMenus()
        {
            foreach (var ext in ContextExtensions)
            {
                var cmdKeyPath = $@"Software\Classes\SystemFileAssociations\{ext}\shell\{ContextMenuName}\command";
                using (var cmdKey = Registry.CurrentUser.OpenSubKey(cmdKeyPath, writable: true))
                {
                    var val = cmdKey?.GetValue(null) as string;
                    if (cmdKey == null || string.IsNullOrWhiteSpace(val))
                    {
                        using (var fix = Registry.CurrentUser.CreateSubKey(cmdKeyPath))
                        {
                            fix?.SetValue(null, $"\"{Application.ExecutablePath}\" \"%1\"");
                            fix?.Flush();
                        }
                    }
                }
            }
        }

        private static void UnregisterContextMenus()
        {
            foreach (var ext in ContextExtensions)
            {
                try
                {
                    Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\SystemFileAssociations\{ext}\shell\{ContextMenuName}", false);
                }
                catch { }
            }
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

        private static string FormatKb(long bytes)
        {
            if (bytes <= 0) return "0KB";
            long kb = bytes / 1024;
            if (kb == 0) kb = 1;
            return kb + "KB";
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
                                skip = true;
                            }
                            else if (StartsWithAscii(data, "http://ns.adobe.com/xap/1.0/") ||
                                     StartsWithAscii(data, "http://ns.adobe.com/xmp/extension/"))
                            {
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
            return true;
        }

        private static bool RemoveExifFromPng(string filePath)
        {
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

        private void AppendToTextLog(string content)
        {
            if (txtLog == null) return;

            if (txtLog.InvokeRequired)
            {
                txtLog.BeginInvoke(new Action<string>(AppendToTextLog), content);
                return;
            }

            txtLog.AppendText(content);
            txtLog.SelectionStart = txtLog.TextLength;
            txtLog.ScrollToCaret();
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
                            .Where(File.Exists)
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
    }
}