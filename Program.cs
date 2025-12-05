using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace EXIF_Remover
{
    internal static class Program
    {
        private const string MutexName = "EXIF_Remover_SingleInstance_v1";
        private const int WM_COPYDATA = 0x004A;
        private const string MainWindowTitle = "EXIF Remover"; // must match Form1.Text

        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var argv = (args?.Length > 0 ? args : Environment.GetCommandLineArgs()?.Skip(1)?.ToArray())
                       ?? Array.Empty<string>();
            // Accept both files and directories from shell/context menu
            var startupFiles = argv.Where(p => File.Exists(p) || Directory.Exists(p)).ToArray();

            using (var mutex = new Mutex(true, MutexName, out bool isOwner))
            {
                if (isOwner)
                {
                    Application.Run(new Form1(startupFiles));
                }
                else
                {
                    if (startupFiles.Length > 0)
                    {
                        var payload = string.Join("\n", startupFiles);
                        var hwnd = FindPrimaryWindow();
                        if (hwnd != IntPtr.Zero)
                        {
                            SendString(hwnd, payload);
                        }
                    }
                    // exit immediately
                }
            }
        }

        private static IntPtr FindPrimaryWindow()
        {
            var hwnd = Native.FindWindow(null, MainWindowTitle);
            if (hwnd != IntPtr.Zero) return hwnd;

            try
            {
                var exeName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);
                var procs = Process.GetProcessesByName(exeName);
                foreach (var p in procs)
                {
                    if (p.Id == Process.GetCurrentProcess().Id) continue;
                    if (p.MainWindowHandle != IntPtr.Zero) return p.MainWindowHandle;
                }
            }
            catch
            {
            }

            return IntPtr.Zero;
        }

        private static void SendString(IntPtr hwnd, string text)
        {
            var managed = (text ?? string.Empty) + "\0";
            var hMem = Marshal.StringToHGlobalUni(managed);
            try
            {
                var cds = new Native.COPYDATASTRUCT
                {
                    dwData = IntPtr.Zero,
                    cbData = managed.Length * 2,
                    lpData = hMem
                };
                Native.SendMessage(hwnd, WM_COPYDATA, IntPtr.Zero, ref cds);
            }
            finally
            {
                if (hMem != IntPtr.Zero) Marshal.FreeHGlobal(hMem);
            }
        }

        /*
        private static void TryAppendLog(string message)
        {
            try
            {
                var root = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "EXIFRemover");
                var path = Path.Combine(root, "EXIF_Remover.forwarder.log");
                Directory.CreateDirectory(root);
                File.AppendAllText(path, $"[{DateTime.Now:O}] {message}{Environment.NewLine}");
            }
            catch
            {
            }
        }
        */

        private static class Native
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct COPYDATASTRUCT
            {
                public IntPtr dwData;
                public int cbData;
                public IntPtr lpData;
            }

            [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = false)]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = false)]
            public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, ref COPYDATASTRUCT lParam);
        }
    }
}