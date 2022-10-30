using Microsoft.Win32.SafeHandles;
using System;
using System.Diagnostics;
using System.Linq;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Clippy
{
    internal static class Program
    {
        private static readonly string MainMutexName = "ClippyApplicationMutex";
        private static Mutex MyMutex;
        private static int parentPID;
        //private static ResourceManager CurrentLangResourceManager;

        [STAThread]
        private static void Main(string[] args)
        {
            bool elevatedMode = args.Contains("/elevated");
            if (elevatedMode)
            {
                parentPID = Int32.Parse(args[1]);
                string elevatedMutexName = "ClippyElevatedMutex" + parentPID;
                MyMutex = new Mutex(true, elevatedMutexName);
                if (!MyMutex.WaitOne(0, false))
                    return;
                Process parentProcess;
                try
                {
                    parentProcess = Process.GetProcessById(parentPID);
                }
                catch
                {
                    return;
                }
                EventWaitHandle pasteEventWaiter = Paster.GetPasteEventWaiter(parentPID);
                EventWaitHandle sendCharsFastEventWaiter = Paster.GetSendCharsEventWaiter(parentPID, false);
                EventWaitHandle sendCharsSlowEventWaiter = Paster.GetSendCharsEventWaiter(parentPID, true); // not implemented

                // https://social.msdn.microsoft.com/Forums/vstudio/en-US/ade69140-490f-4070-b8b5-08ac80d1c557/how-to-waitany-on-a-process-a-manualresetevent?forum=netfxbcl
                ManualResetEvent processExitWaiter = new ManualResetEvent(true);
                processExitWaiter.SafeWaitHandle = new SafeWaitHandle(parentProcess.Handle, false);

                WaitHandle[] waitHandles = new WaitHandle[4] { pasteEventWaiter, processExitWaiter, sendCharsFastEventWaiter, sendCharsSlowEventWaiter };
                while (true)
                {
                    try
                    {
                        Process.GetProcessById(parentPID);
                    }
                    catch
                    {
                        break;
                    }
                    int eventIndex = WaitHandle.WaitAny(waitHandles, 60000); // sometimes waitAny begins consume 100% cpu core. Set timeout is workaround try and it seems working well.
                    if (eventIndex != WaitHandle.WaitTimeout)
                    {
                        if (waitHandles[eventIndex] == pasteEventWaiter)
                        {
                            Paster.SendPaste();
                            pasteEventWaiter.Reset();
                        }
                        else if (waitHandles[eventIndex] == sendCharsFastEventWaiter)
                        {
                            Paster.SendChars(null, false);
                            sendCharsFastEventWaiter.Reset();
                        }
                        //else if (waitHandles[eventIndex] == sendCharsSlowEventWaiter)
                        //{
                        //    Paster.SendChars(null, true);
                        //    sendCharsSlowEventWaiter.Reset();
                        //}
                        else if (waitHandles[eventIndex] == processExitWaiter)
                        {
                            break;
                        }
                    }
                    else
                    {
                        break; // sometimes waitAny begins consume 100% cpu core. This is workaround try.
                    }
                }
            }
            else
            {
                if (IsSingleInstance(MainMutexName))
                {
                    //string UserSettingsPath = Application.StartupPath + "UserSettings";
                    //bool PortableMode = true;
                    //string portableFlagFileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\PortableMode.txt";
                    //PortableMode = args.Contains("/p") || args.Contains("/portable") || File.Exists(portableFlagFileName);
                    // http://stackoverflow.com/questions/1382617/how-to-make-designer-generated-net-application-settings-portable/2579399#2579399
                    //UserSettingsPath = MakePortable(Properties.Settings.Default, PortableMode);

                    Application.EnableVisualStyles();
                    Application.SetHighDpiMode(HighDpiMode.SystemAware);
                    Application.SetCompatibleTextRenderingDefault(false);
                    Main Main = new Main(args.Contains("/m"));
                    Application.AddMessageFilter(new TestMessageFilter());
                    Application.Run(Main);
                    //Properties.Settings.Default.Save();
                }
            }
        }

        //private static string MakePortable(ApplicationSettingsBase settings, bool PortableMode = true)
        //{
        //    var portableSettingsProvider = new PortableSettingsProvider(settings.GetType().Name + ".settings", PortableMode);
        //    settings.Providers.Add(portableSettingsProvider);
        //    foreach (SettingsProperty prop in settings.Properties)
        //        prop.Provider = portableSettingsProvider;
        //    settings.Reload();
        //    return portableSettingsProvider.GetAppSettingsPath();
        //}

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private static bool IsSingleInstance(string MutexName)
        {
            MyMutex = new Mutex(true, MutexName);
            if (MyMutex.WaitOne(0, false))
                return true;
            Process[] allProcess = Process.GetProcesses();
            Process currentProcess = Process.GetCurrentProcess();
            foreach (Process process in allProcess)
            {
                if (currentProcess.ProcessName == process.ProcessName && currentProcess.Id != process.Id)
                {
                    IntPtr hWnd = GetProcessMainWindow(process.Id);
                    if (hWnd != IntPtr.Zero)
                    {
                        //string dummy;
                        //CurrentLangResourceManager = Clippy.Main.getResourceManager(out dummy);
                        ShowWindow(hWnd, 1);
                        SetForegroundWindow(hWnd);
                        //MessageBox.Show(CurrentLangResourceManager.GetString("ActivatedExistedProcess"), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information,
                        //    MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
            }
            return false;
        }

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr parentWindow, IntPtr previousChildWindow, string windowClass, string windowTitle);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowThreadProcessId(IntPtr window, out int process);

        [DllImport("user32.dll")]
        private static extern IntPtr GetProp(IntPtr hWnd, string lpString);

        private static IntPtr GetProcessMainWindow(int process)
        {
            IntPtr pLast = IntPtr.Zero;
            do
            {
                pLast = FindWindowEx(IntPtr.Zero, pLast, null, null);
                int iProcess_;
                GetWindowThreadProcessId(pLast, out iProcess_);
                if (iProcess_ == process)
                {
                    IntPtr propVal = GetProp(pLast, Clippy.Main.IsMainPropName);
                    if (propVal.ToInt32() == 1)
                        break;
                }
            } while (pLast != IntPtr.Zero);
            return pLast;
        }

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern long GetClassName(IntPtr hwnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        // To support exit by taskkill signal http://qaru.site/questions/1648501/windows-application-handle-taskkill
        //[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        private class TestMessageFilter : IMessageFilter
        {
            public bool PreFilterMessage(ref Message m)
            {
                if (m.Msg == /*WM_CLOSE*/ 0x10)
                {
                    int nChars = 100;
                    StringBuilder className = new StringBuilder(nChars);
                    GetClassName(m.HWnd, className, nChars);
                    if (className.ToString() != "URL Moniker Notification Window") // on some HTML clips with ShowNativeFormatting=True somehow this windows sends this message
                    {
                        Application.Exit();
                        return true;
                    }
                }
                return false;
            }
        }
    }
}