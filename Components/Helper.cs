using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Drawing;

namespace Clippy
{
    // http://stackoverflow.com/questions/1220213/detect-if-running-as-administrator-with-or-without-elevated-privileges
    public static class UacHelper
    {
        private const string uacRegistryKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System";
        private const string uacRegistryValue = "EnableLUA";

        private static uint STANDARD_RIGHTS_READ = 0x00020000;
        private static uint TOKEN_QUERY = 0x0008;
        private static uint TOKEN_READ = (STANDARD_RIGHTS_READ | TOKEN_QUERY);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool OpenProcessToken(IntPtr ProcessHandle, UInt32 DesiredAccess, out IntPtr TokenHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool GetTokenInformation(IntPtr TokenHandle, TOKEN_INFORMATION_CLASS TokenInformationClass, IntPtr TokenInformation, uint TokenInformationLength, out uint ReturnLength);

        public enum TOKEN_INFORMATION_CLASS
        {
            TokenUser = 1,
            TokenGroups,
            TokenPrivileges,
            TokenOwner,
            TokenPrimaryGroup,
            TokenDefaultDacl,
            TokenSource,
            TokenType,
            TokenImpersonationLevel,
            TokenStatistics,
            TokenRestrictedSids,
            TokenSessionId,
            TokenGroupsAndPrivileges,
            TokenSessionReference,
            TokenSandBoxInert,
            TokenAuditPolicy,
            TokenOrigin,
            TokenElevationType,
            TokenLinkedToken,
            TokenElevation,
            TokenHasRestrictions,
            TokenAccessInformation,
            TokenVirtualizationAllowed,
            TokenVirtualizationEnabled,
            TokenIntegrityLevel,
            TokenUIAccess,
            TokenMandatoryPolicy,
            TokenLogonSid,
            MaxTokenInfoClass
        }

        public enum TOKEN_ELEVATION_TYPE
        {
            TokenElevationTypeDefault = 1,
            TokenElevationTypeFull,
            TokenElevationTypeLimited
        }

        public static bool IsUacEnabled
        {
            get
            {
                using (RegistryKey uacKey = Registry.LocalMachine.OpenSubKey(uacRegistryKey, false))
                {
                    bool result = uacKey.GetValue(uacRegistryValue).Equals(1);
                    return result;
                }
            }
        }

        public static bool IsProcessAccessible(int PID = 0)
        {
            IntPtr tokenHandle = IntPtr.Zero;
            Process TargetProcess = Process.GetProcessById(PID);
            try
            {
                OpenProcessToken(TargetProcess.Handle, TOKEN_READ, out tokenHandle);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool IsProcessElevated()
        {
            if (IsUacEnabled)
            {
                IntPtr tokenHandle = IntPtr.Zero;
                if (!OpenProcessToken(Process.GetCurrentProcess().Handle, TOKEN_READ, out tokenHandle))
                {
                    throw new ApplicationException("Could not get process token.  Win32 Error Code: " +
                                                    Marshal.GetLastWin32Error());
                }

                try
                {
                    TOKEN_ELEVATION_TYPE elevationResult = TOKEN_ELEVATION_TYPE.TokenElevationTypeDefault;

                    //int elevationResultSize = Marshal.SizeOf(typeof(TOKEN_ELEVATION_TYPE));
                    int elevationResultSize = Marshal.SizeOf((int)elevationResult);
                    uint returnedSize = 0;

                    IntPtr elevationTypePtr = Marshal.AllocHGlobal(elevationResultSize);
                    try
                    {
                        bool success = GetTokenInformation(tokenHandle, TOKEN_INFORMATION_CLASS.TokenElevationType,
                                                            elevationTypePtr, (uint)elevationResultSize,
                                                            out returnedSize);
                        if (success)
                        {
                            elevationResult = (TOKEN_ELEVATION_TYPE)Marshal.ReadInt32(elevationTypePtr);
                            bool isProcessAdmin = elevationResult == TOKEN_ELEVATION_TYPE.TokenElevationTypeFull;
                            return isProcessAdmin;
                        }
                        else
                        {
                            throw new ApplicationException("Unable to determine the current elevation.");
                        }
                    }
                    finally
                    {
                        if (elevationTypePtr != IntPtr.Zero)
                            Marshal.FreeHGlobal(elevationTypePtr);
                    }
                }
                finally
                {
                    if (tokenHandle != IntPtr.Zero)
                        CloseHandle(tokenHandle);
                }
            }
            else
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                bool result = principal.IsInRole(WindowsBuiltInRole.Administrator)
                            || principal.IsInRole(0x200); //Domain Administrator
                return result;
            }
        }
    }

    public static class DataTableEnumerate
    {
        public static void Fill<T>(this IEnumerable<T> Ts, ref DataTable dt) where T : class
        {
            //Get Enumerable Type
            Type tT = typeof(T);

            //Get Collection of NoVirtual properties
            var T_props = tT.GetProperties().Where(p => !p.GetGetMethod().IsVirtual).ToArray();

            //Fill Schema
            foreach (PropertyInfo p in T_props)
                dt.Columns.Add(p.Name, p.GetMethod.ReturnParameter.ParameterType.BaseType);

            //Fill Data
            foreach (T t in Ts)
            {
                DataRow row = dt.NewRow();

                foreach (PropertyInfo p in T_props)
                    row[p.Name] = p.GetValue(t);

                dt.Rows.Add(row);
            }

        }
    }
    public enum PasteMethod
    {
        Standard,
        Text,
        Line,
        SendCharsFast,
        SendCharsSlow,
        File,
        Null
    };

    public delegate int WindowProcDelegate(IntPtr hw, IntPtr uMsg, IntPtr wParam, IntPtr lParam);

    /// <summary>
    /// Windows Event Messages sent to the WindowProc
    /// </summary>
    public enum Msgs
    {
        WM_NULL = 0x0000,
        WM_CREATE = 0x0001,
        WM_DESTROY = 0x0002,
        WM_MOVE = 0x0003,
        WM_SIZE = 0x0005,
        WM_ACTIVATE = 0x0006,
        WM_SETFOCUS = 0x0007,
        WM_KILLFOCUS = 0x0008,
        WM_ENABLE = 0x000A,
        WM_SETREDRAW = 0x000B,
        WM_SETTEXT = 0x000C,
        WM_GETTEXT = 0x000D,
        WM_GETTEXTLENGTH = 0x000E,
        WM_PAINT = 0x000F,
        WM_CLOSE = 0x0010,
        WM_QUERYENDSESSION = 0x0011,
        WM_QUIT = 0x0012,
        WM_QUERYOPEN = 0x0013,
        WM_ERASEBKGND = 0x0014,
        WM_SYSCOLORCHANGE = 0x0015,
        WM_ENDSESSION = 0x0016,
        WM_SHOWWINDOW = 0x0018,
        WM_WININICHANGE = 0x001A,
        WM_SETTINGCHANGE = 0x001A,
        WM_DEVMODECHANGE = 0x001B,
        WM_ACTIVATEAPP = 0x001C,
        WM_FONTCHANGE = 0x001D,
        WM_TIMECHANGE = 0x001E,
        WM_CANCELMODE = 0x001F,
        WM_SETCURSOR = 0x0020,
        WM_MOUSEACTIVATE = 0x0021,
        WM_CHILDACTIVATE = 0x0022,
        WM_QUEUESYNC = 0x0023,
        WM_GETMINMAXINFO = 0x0024,
        WM_PAINTICON = 0x0026,
        WM_ICONERASEBKGND = 0x0027,
        WM_NEXTDLGCTL = 0x0028,
        WM_SPOOLERSTATUS = 0x002A,
        WM_DRAWITEM = 0x002B,
        WM_MEASUREITEM = 0x002C,
        WM_DELETEITEM = 0x002D,
        WM_VKEYTOITEM = 0x002E,
        WM_CHARTOITEM = 0x002F,
        WM_SETFONT = 0x0030,
        WM_GETFONT = 0x0031,
        WM_SETHOTKEY = 0x0032,
        WM_GETHOTKEY = 0x0033,
        WM_QUERYDRAGICON = 0x0037,
        WM_COMPAREITEM = 0x0039,
        WM_GETOBJECT = 0x003D,
        WM_COMPACTING = 0x0041,
        WM_COMMNOTIFY = 0x0044,
        WM_WINDOWPOSCHANGING = 0x0046,
        WM_WINDOWPOSCHANGED = 0x0047,
        WM_POWER = 0x0048,
        WM_COPYDATA = 0x004A,
        WM_CANCELJOURNAL = 0x004B,
        WM_NOTIFY = 0x004E,
        WM_INPUTLANGCHANGEREQUEST = 0x0050,
        WM_INPUTLANGCHANGE = 0x0051,
        WM_TCARD = 0x0052,
        WM_HELP = 0x0053,
        WM_USERCHANGED = 0x0054,
        WM_NOTIFYFORMAT = 0x0055,
        WM_CONTEXTMENU = 0x007B,
        WM_STYLECHANGING = 0x007C,
        WM_STYLECHANGED = 0x007D,
        WM_DISPLAYCHANGE = 0x007E,
        WM_GETICON = 0x007F,
        WM_SETICON = 0x0080,
        WM_NCCREATE = 0x0081,
        WM_NCDESTROY = 0x0082,
        WM_NCCALCSIZE = 0x0083,
        WM_NCHITTEST = 0x0084,
        WM_NCPAINT = 0x0085,
        WM_NCACTIVATE = 0x0086,
        WM_GETDLGCODE = 0x0087,
        WM_SYNCPAINT = 0x0088,
        WM_NCMOUSEMOVE = 0x00A0,
        WM_NCLBUTTONDOWN = 0x00A1,
        WM_NCLBUTTONUP = 0x00A2,
        WM_NCLBUTTONDBLCLK = 0x00A3,
        WM_NCRBUTTONDOWN = 0x00A4,
        WM_NCRBUTTONUP = 0x00A5,
        WM_NCRBUTTONDBLCLK = 0x00A6,
        WM_NCMBUTTONDOWN = 0x00A7,
        WM_NCMBUTTONUP = 0x00A8,
        WM_NCMBUTTONDBLCLK = 0x00A9,
        WM_NCXBUTTONDOWN = 0x00AB,
        WM_NCXBUTTONUP = 0x00AC,
        WM_KEYDOWN = 0x0100,
        WM_KEYUP = 0x0101,
        WM_CHAR = 0x0102,
        WM_DEADCHAR = 0x0103,
        WM_SYSKEYDOWN = 0x0104,
        WM_SYSKEYUP = 0x0105,
        WM_SYSCHAR = 0x0106,
        WM_SYSDEADCHAR = 0x0107,
        WM_KEYLAST = 0x0108,
        WM_IME_STARTCOMPOSITION = 0x010D,
        WM_IME_ENDCOMPOSITION = 0x010E,
        WM_IME_COMPOSITION = 0x010F,
        WM_IME_KEYLAST = 0x010F,
        WM_INITDIALOG = 0x0110,
        WM_COMMAND = 0x0111,
        WM_SYSCOMMAND = 0x0112,
        WM_TIMER = 0x0113,
        WM_HSCROLL = 0x0114,
        WM_VSCROLL = 0x0115,
        WM_INITMENU = 0x0116,
        WM_INITMENUPOPUP = 0x0117,
        WM_MENUSELECT = 0x011F,
        WM_MENUCHAR = 0x0120,
        WM_ENTERIDLE = 0x0121,
        WM_MENURBUTTONUP = 0x0122,
        WM_MENUDRAG = 0x0123,
        WM_MENUGETOBJECT = 0x0124,
        WM_UNINITMENUPOPUP = 0x0125,
        WM_MENUCOMMAND = 0x0126,
        WM_CTLCOLORMSGBOX = 0x0132,
        WM_CTLCOLOREDIT = 0x0133,
        WM_CTLCOLORLISTBOX = 0x0134,
        WM_CTLCOLORBTN = 0x0135,
        WM_CTLCOLORDLG = 0x0136,
        WM_CTLCOLORSCROLLBAR = 0x0137,
        WM_CTLCOLORSTATIC = 0x0138,
        WM_MOUSEMOVE = 0x0200,
        WM_LBUTTONDOWN = 0x0201,
        WM_LBUTTONUP = 0x0202,
        WM_LBUTTONDBLCLK = 0x0203,
        WM_RBUTTONDOWN = 0x0204,
        WM_RBUTTONUP = 0x0205,
        WM_RBUTTONDBLCLK = 0x0206,
        WM_MBUTTONDOWN = 0x0207,
        WM_MBUTTONUP = 0x0208,
        WM_MBUTTONDBLCLK = 0x0209,
        WM_MOUSEWHEEL = 0x020A,
        WM_XBUTTONDOWN = 0x020B,
        WM_XBUTTONUP = 0x020C,
        WM_XBUTTONDBLCLK = 0x020D,
        WM_PARENTNOTIFY = 0x0210,
        WM_ENTERMENULOOP = 0x0211,
        WM_EXITMENULOOP = 0x0212,
        WM_NEXTMENU = 0x0213,
        WM_SIZING = 0x0214,
        WM_CAPTURECHANGED = 0x0215,
        WM_MOVING = 0x0216,
        WM_DEVICECHANGE = 0x0219,
        WM_MDICREATE = 0x0220,
        WM_MDIDESTROY = 0x0221,
        WM_MDIACTIVATE = 0x0222,
        WM_MDIRESTORE = 0x0223,
        WM_MDINEXT = 0x0224,
        WM_MDIMAXIMIZE = 0x0225,
        WM_MDITILE = 0x0226,
        WM_MDICASCADE = 0x0227,
        WM_MDIICONARRANGE = 0x0228,
        WM_MDIGETACTIVE = 0x0229,
        WM_MDISETMENU = 0x0230,
        WM_ENTERSIZEMOVE = 0x0231,
        WM_EXITSIZEMOVE = 0x0232,
        WM_DROPFILES = 0x0233,
        WM_MDIREFRESHMENU = 0x0234,
        WM_IME_SETCONTEXT = 0x0281,
        WM_IME_NOTIFY = 0x0282,
        WM_IME_CONTROL = 0x0283,
        WM_IME_COMPOSITIONFULL = 0x0284,
        WM_IME_SELECT = 0x0285,
        WM_IME_CHAR = 0x0286,
        WM_IME_REQUEST = 0x0288,
        WM_IME_KEYDOWN = 0x0290,
        WM_IME_KEYUP = 0x0291,
        WM_MOUSEHOVER = 0x02A1,
        WM_MOUSELEAVE = 0x02A3,
        WM_CUT = 0x0300,
        WM_COPY = 0x0301,
        WM_PASTE = 0x0302,
        WM_CLEAR = 0x0303,
        WM_UNDO = 0x0304,
        WM_RENDERFORMAT = 0x0305,
        WM_RENDERALLFORMATS = 0x0306,
        WM_DESTROYCLIPBOARD = 0x0307,
        WM_DRAWCLIPBOARD = 0x0308,
        WM_PAINTCLIPBOARD = 0x0309,
        WM_VSCROLLCLIPBOARD = 0x030A,
        WM_SIZECLIPBOARD = 0x030B,
        WM_ASKCBFORMATNAME = 0x030C,
        WM_CHANGECBCHAIN = 0x030D,
        WM_HSCROLLCLIPBOARD = 0x030E,
        WM_QUERYNEWPALETTE = 0x030F,
        WM_PALETTEISCHANGING = 0x0310,
        WM_PALETTECHANGED = 0x0311,
        WM_HOTKEY = 0x0312,
        WM_PRINT = 0x0317,
        WM_PRINTCLIENT = 0x0318,
        WM_CLIPBOARDUPDATE = 0x031D,
        WM_HANDHELDFIRST = 0x0358,
        WM_HANDHELDLAST = 0x035F,
        WM_AFXFIRST = 0x0360,
        WM_AFXLAST = 0x037F,
        WM_PENWINFIRST = 0x0380,
        WM_PENWINLAST = 0x038F,
        WM_APP = 0x8000,
        WM_USER = 0x0400
    }


    /// <summary>
    /// Windows User32 DLL declarations
    /// </summary>
    public class User32
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SetClipboardViewer(IntPtr hWnd);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern bool ChangeClipboardChain(
            IntPtr hWndRemove,  // handle to window to remove
            IntPtr hWndNewNext  // handle to next window
            );

        [DllImport("user32.dll")]
        public static extern IntPtr GetClipboardViewer();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, IntPtr wParam, IntPtr lParam);

    }

    public partial class Main : Form
    {

        //[DllImport("user32.dll")]
        //public static extern IntPtr GetParent(IntPtr hWnd);

        //[DllImport("gdi32.dll")]
        //private static extern IntPtr SetMetaFileBitsEx(uint _bufferSize, byte[] _buffer);

        //[DllImport("user32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        //[DllImport("gdi32.dll")]
        //private static extern IntPtr CopyMetaFile(IntPtr hWmf, string filename);

        //[DllImport("gdi32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool DeleteEnhMetaFile(IntPtr hEmf);

        //[DllImport("gdi32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool DeleteMetaFile(IntPtr hWmf);

        //[DllImport("gdiplus.dll")]
        //private static extern uint GdipEmfToWmfBits(IntPtr _hEmf, uint _bufferSize, byte[] _buffer, int _mappingMode, EmfToWmfBitsFlags _flags);

        //[DllImport("user32.dll")]
        //private static extern IntPtr GetFocus();

        //[DllImport("user32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool EnableWindow(IntPtr hwnd, bool bEnable);

        //[DllImport("user32.dll")]
        //private static extern IntPtr GetActiveWindow();

        //[DllImport("User32.dll")]
        //private static extern short GetAsyncKeyState(Keys vKey);

        //[DllImport("user32", SetLastError = true)]
        //private static extern int GetCaretPos(out System.Drawing.Point p);

        //[DllImport("kernel32.dll")]
        //private static extern uint GetCurrentThreadId();

        //[DllImport("user32.dll")]
        //private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        //[DllImport("user32.dll", SetLastError = true)]
        //private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        //// Code For OpenWithDialog Box
        //[DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool ShellExecuteEx(ref ShellExecuteInfo lpExecInfo);

        //[DllImport("user32", SetLastError = true)]
        //private static extern int SetCaretPos(int x, int y);

        //[DllImport("user32.dll")]
        //private static extern IntPtr SetFocus(IntPtr hWnd);

        //[DllImport("User32.dll")]
        //private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);


        [DllImport("user32.dll", EntryPoint = "GetGUIThreadInfo")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetGUIThreadInfo(uint tId, out GUITHREADINFO threadInfo);
        [DllImport("kernel32.dll")]
        public static extern IntPtr OpenProcess(uint processAccess, bool bInheritHandle, int processId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr handle, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool MoveWindow(IntPtr hwnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr hObject);

        [DllImport("user32.dll")]
        private static extern IntPtr GetClipboardOwner();

        [DllImport("User32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ClientToScreen(IntPtr hWnd, out System.Drawing.Point position);

        [DllImport("psapi.dll", CharSet = CharSet.Unicode)]
        //private static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, [Out] StringBuilder lpBaseName, [In][MarshalAs(UnmanagedType.U4)] int nSize);
        private static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder lpBaseName, int nSize);

        [DllImport("user32.dll")]
        private static extern IntPtr GetOpenClipboardWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern int GetWindowText(IntPtr hWnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("User32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProp(IntPtr hWnd, string lpString, IntPtr hData);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
           WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll")]
        private static extern uint ActivateKeyboardLayout(uint hkl, uint flags);

        [DllImport("user32.dll")]
        private static extern uint GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint LoadKeyboardLayout(StringBuilder pwszKLID, uint flags);

        // http://stackoverflow.com/questions/9501771/how-to-avoid-a-win32-exception-when-accessing-process-mainmodule-filename-in-c
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public static string GetProcessMainModuleFullName(int pid)
        {
            string result = null;
            var processHandle = OpenProcess(0x0400 | 0x0010, false, pid);
            if (processHandle == IntPtr.Zero)
            {
                // Not enough priviledges. Need to call it elevated
                return null;
            }
            const int lengthSb = 1000; // Dirty
            var sb = new StringBuilder(lengthSb);
            
            // Possibly there is no such fuction in Windows 7 https://stackoverflow.com/a/321343/4085971
            if (GetModuleFileNameEx(processHandle, IntPtr.Zero, sb, sb.Capacity) > 0)
            {
                result = sb.ToString();
            }
            CloseHandle(processHandle);
            return result;
        }

        private static GUITHREADINFO GetGuiInfo(IntPtr hWindow, out Point point)
        {
            int pid;
            uint remoteThreadId = GetWindowThreadProcessId(hWindow, out pid);
            var guiInfo = new GUITHREADINFO();
            guiInfo.cbSize = (uint)Marshal.SizeOf(guiInfo);
            GetGUIThreadInfo(remoteThreadId, out guiInfo);
            point = new Point(0, 0);
            ClientToScreen(guiInfo.hwndCaret, out point);

            return guiInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public uint cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        };

        public struct LastClip
        {
            public DateTime Created;
            public int ID;
            public int ProcessID;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        public void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            // http://stackoverflow.com/a/10280800/4085971
            UpdateLastActiveParentWindow(hwnd);
        }

        //private IntPtr GetFocusWindow(int maxWait = 100)
        //{
        //    IntPtr hwnd = GetForegroundWindow();
        //    int pid;
        //    uint remoteThreadId = GetWindowThreadProcessId(hwnd, out pid);
        //    uint currentThreadId = GetCurrentThreadId();
        //    //AttachTrheadInput is needed so we can get the handle of a focused window in another app
        //    AttachThreadInput(remoteThreadId, currentThreadId, true);
        //    Stopwatch stopWatch = new Stopwatch();
        //    stopWatch.Start();
        //    while (stopWatch.ElapsedMilliseconds < maxWait)
        //    {
        //        hwnd = GetFocus();
        //        if (hwnd != IntPtr.Zero)
        //            break;
        //        Thread.Sleep(5);
        //    }
        //    AttachThreadInput(remoteThreadId, currentThreadId, false);
        //    return hwnd;
        //}

        //internal static class NativeMethods
        //{
        //    [DllImport("user32.dll", SetLastError = true)]
        //    internal static extern uint SendInput(uint nInputs, ref NativeStructs.Input pInputs, int cbSize);
        //}

        //private enum GetWindowCmd : uint
        //{
        //    GW_HWNDFIRST = 0,
        //    GW_HWNDLAST = 1,
        //    GW_HWNDNEXT = 2,
        //    GW_HWNDPREV = 3,
        //    GW_OWNER = 4,
        //    GW_CHILD = 5,
        //    GW_ENABLEDPOPUP = 6
        //}

        protected override void WndProc(ref Message m)
        {
            if (settings.FastWindowOpen)
            {
                if (m.Msg == WM_SYSCOMMAND)
                {
                    if (m.WParam.ToInt32() == SC_MINIMIZE) // TODO catch MinimizeAll command if it is possible
                    {
                        m.Result = IntPtr.Zero;
                        Close();
                        return;
                    }
                }
            }
            switch ((Msgs)m.Msg)
            {
                case Msgs.WM_CLIPBOARDUPDATE:
                    Debug.WriteLine("WM_CLIPBOARDUPDATE", "WndProc");
                    if (UsualClipboardMode)
                        UsualClipboardMode = false;
                    else if (!captureTimer.Enabled)
                    {
                        captureTimer.Interval = 100;
                        captureTimer.Start();
                    }
                    break;

                case Msgs.WM_NCMOUSEMOVE:
                    if (PreviousCursor != Cursor.Position)
                    {
                        titleToolTipBeforeTimer.Stop();
                        titleToolTip.Hide(this);
                    }
                    if (m.WParam == new IntPtr(0x0002)) // HT_CAPTION
                    {
                        if (!titleToolTipShown)
                            titleToolTipBeforeTimer.Start();
                    }
                    PreviousCursor = Cursor.Position;
                    return;

                case Msgs.WM_MOVE:
                    titleToolTipShown = true;
                    titleToolTipBeforeTimer.Stop();
                    titleToolTip.Hide(this);
                    return;

                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        // RTF Image Format
        // {\pict\wmetafile8\picw[A]\pich[B]\picwgoal[C]\pichgoal[D]
        //
        // A    = (Image Width in Pixels / Graphics.DpiX) * 2540
        //
        // B    = (Image Height in Pixels / Graphics.DpiX) * 2540
        //
        // C    = (Image Width in Pixels / Graphics.DpiX) * 1440
        //
        // D    = (Image Height in Pixels / Graphics.DpiX) * 1440

        //public void RestoreClipboard()
        //{
        //    DataObject o = new DataObject();
        //    foreach (string format in clipboardContents.Keys)
        //    {
        //        o.SetData(format, clipboardContents[format]);
        //    }
        //    SetClipboardDataObject(o, false);
        //}

        //public string GetImageForRTF(Image img, int width = 0, int height = 0)
        //{
        //    //string newPath = Path.Combine(Environment.CurrentDirectory, path);
        //    //Image img = Image.FromFile(newPath);
        //    if (width == 0)
        //        width = img.Width;
        //    if (height == 0)
        //        height = img.Width;
        //    MemoryStream stream = new MemoryStream();
        //    img.Save(stream, ImageFormat.Bmp);
        //    byte[] bytes = stream.ToArray();
        //    string str = BitConverter.ToString(bytes, 0).Replace("-", string.Empty);
        //    //string str = System.Text.Encoding.UTF8.GetString(bytes);
        //    string mpic = @"{\pict\wbitmapN\picw" + img.Width + @"\pich" + img.Height + @"\picwgoal" + width +
        //                  @"\pichgoal" + height + @"\bin " + str + "}";
        //    return mpic;
        //}

        //private static IntPtr GetTopParentWindow(IntPtr hForegroundWindow)
        //{
        //    while (true)
        //    {
        //        IntPtr temp = GetParent(hForegroundWindow);
        //        if (temp.Equals(IntPtr.Zero)) break;
        //        hForegroundWindow = temp;
        //    }

        //    return hForegroundWindow;
        //}

        //public static T ParseEnum<T>(string value)
        //{
        //    return (T)Enum.Parse(typeof(T), value, true);
        //}

        //public static ImageFormat GetImageFormat(Image img)
        //{
        //    if (img.RawFormat.Equals(ImageFormat.Jpeg))
        //        return ImageFormat.Jpeg;
        //    if (img.RawFormat.Equals(ImageFormat.Bmp))
        //        return ImageFormat.Bmp;
        //    if (img.RawFormat.Equals(ImageFormat.Png))
        //        return ImageFormat.Png;
        //    if (img.RawFormat.Equals(ImageFormat.Emf))
        //        return ImageFormat.Emf;
        //    if (img.RawFormat.Equals(ImageFormat.Exif))
        //        return ImageFormat.Exif;
        //    if (img.RawFormat.Equals(ImageFormat.Gif))
        //        return ImageFormat.Gif;
        //    if (img.RawFormat.Equals(ImageFormat.Icon))
        //        return ImageFormat.Icon;
        //    if (img.RawFormat.Equals(ImageFormat.MemoryBmp))
        //        return ImageFormat.MemoryBmp;
        //    if (img.RawFormat.Equals(ImageFormat.Tiff))
        //        return ImageFormat.Tiff;
        //    else
        //        return ImageFormat.Wmf;
        //}

        //public void BackupClipboard()
        //{
        //    clipboardContents.Clear();
        //    IDataObject o = Clipboard.GetDataObject();
        //    foreach (string format in o.GetFormats())
        //    {
        //        object data;
        //        try
        //        {
        //            data = o.GetData(format);
        //        }
        //        catch
        //        {
        //            Debug.WriteLine(String.Format(Properties.Resources.FailedToReadFormatFromClipboard, format));
        //            continue;
        //        }
        //        clipboardContents.Add(format, data);
        //    }
        //}

        //private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    ClearFilter_Click();
        //}

        //private int CountLines(string str, int position = 0)
        //{
        //    if (str == null)
        //        throw new ArgumentNullException("str");
        //    if (str == string.Empty)
        //        return 0;
        //    int index = -1;
        //    int count = 0;
        //    if (position == 0)
        //        position = str.Length;

        //    while (-1 != (index = str.IndexOf("\n", index + 1)) && position > index)
        //    {
        //        count++;
        //    }
        //    return count + 1;
        //}

        //private string FormatByteSize(int byteSize)
        //{
        //    string[] sizes = { MultiLangByteUnit(), MultiLangKiloByteUnit(), MultiLangMegaByteUnit() };
        //    double len = byteSize;
        //    int order = 0;
        //    while (len >= 1024 && order < sizes.Length - 1)
        //    {
        //        order++;
        //        len = len / 1024;
        //    }

        //    // Adjust the format string to your preferences. For example "{0:0.#}{1}" would
        //    // show a single decimal place, and no space.
        //    string result = String.Format("{0:0.#} {1}", len, sizes[order]);
        //    return result;
        //}

        //// To hide on start
        //protected override void SetVisibleCore(bool value)
        //{
        //    if (!allowVisible)
        //    {
        //        value = false;
        //        if (!this.IsHandleCreated) CreateHandle();
        //    }
        //    base.SetVisibleCore(value);
        //}

        //private void gotoLastToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    GotoLastRow();
        //}

        //private void PrepareTableGrid()
        //{
        //    //ReadFilterText();
        //    foreach (DataGridViewRow row in dataGridView.Rows)
        //    {
        //        PrepareRow(row);
        //    }
        //    dataGridView.Update();
        //}

        //private void RestoreTextSelection(int NewSelectionStart = -1, int NewSelectionLength = -1)
        //{
        //}

        //private void richTextBox_KeyDown(object sender, KeyEventArgs e)
        //{
        //}

        //private void SaveClipUrl(string Url)
        //{
        //    string sql = "Update Clips set Url = @Url where Id = @Id";
        //    SQLiteCommand command = new SQLiteCommand(sql, m_dbConnection);
        //    command.Parameters.AddWithValue("@Id", LoadedClipRowReader["Id"]);
        //    command.Parameters.AddWithValue("@Url", Url);
        //    command.ExecuteNonQuery();
        //}

        //private void Text_CursorChanged(object sender, EventArgs e)
        //{
        //    // This event not working. Why? Decided to use Click instead.
        //}


        //public string AssemblyTitle
        //{
        //    get
        //    {
        //        object[] attributes = Assembly.GetExecutingAssembly()
        //            .GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
        //        if (attributes.Length > 0)
        //        {
        //            AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
        //            if (!String.IsNullOrEmpty(titleAttribute.Title))
        //            {
        //                return titleAttribute.Title;
        //            }
        //        }
        //        return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().Location);
        //    }
        //}

        //public string AssemblyVersion
        //{
        //    get { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); }
        //}

        //public string AssemblyDescription
        //{
        //    get
        //    {
        //        object[] attributes =
        //            Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
        //        if (attributes.Length == 0)
        //        {
        //            return "";
        //        }
        //        return ((AssemblyDescriptionAttribute)attributes[0]).Description;
        //    }
        //}

        //public string AssemblyCopyright
        //{
        //    get
        //    {
        //        object[] attributes =
        //            Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
        //        if (attributes.Length == 0)
        //        {
        //            return "";
        //        }
        //        return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
        //    }
        //}

        //public string AssemblyCompany
        //{
        //    get
        //    {
        //        object[] attributes =
        //            Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
        //        if (attributes.Length == 0)
        //        {
        //            return "";
        //        }
        //        return ((AssemblyCompanyAttribute)attributes[0]).Company;
        //    }
        //}
        //public static string GetEmbedImageString(Bitmap image, int width = 0, int height = 0)
        //{
        //    Metafile metafile = null;
        //    float dpiX;
        //    float dpiY;
        //    if (height == 0 || height > image.Height)
        //        height = image.Height;
        //    if (width == 0 || width > image.Width)
        //        width = image.Width;
        //    using (Graphics g = Graphics.FromImage(image))
        //    {
        //        IntPtr hDC = g.GetHdc();
        //        metafile = new Metafile(hDC, EmfType.EmfOnly);
        //        g.ReleaseHdc(hDC);
        //    }

        //    using (Graphics g = Graphics.FromImage(metafile))
        //    {
        //        g.DrawImage(image, 0, 0);
        //        dpiX = g.DpiX;
        //        dpiY = g.DpiY;
        //    }

        //    IntPtr _hEmf = metafile.GetHenhmetafile();
        //    uint _bufferSize = GdipEmfToWmfBits(_hEmf, 0, null, MM_ANISOTROPIC,
        //        EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);
        //    byte[] _buffer = new byte[_bufferSize];
        //    GdipEmfToWmfBits(_hEmf, _bufferSize, _buffer, MM_ANISOTROPIC,
        //        EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);
        //    IntPtr hmf = SetMetaFileBitsEx(_bufferSize, _buffer);
        //    string tempfile = Path.GetTempFileName();
        //    CopyMetaFile(hmf, tempfile);
        //    DeleteMetaFile(hmf);
        //    DeleteEnhMetaFile(_hEmf);

        //    var stream = new MemoryStream();
        //    byte[] data = File.ReadAllBytes(tempfile);
        //    //File.Delete (tempfile);
        //    int count = data.Length;
        //    stream.Write(data, 0, count);

        //    string proto = @"{\rtf1{\pict\wmetafile8\picw" + (int)(((float)width / dpiX) * 2540)
        //                   + @"\pich" + (int)(((float)height / dpiY) * 2540)
        //                   + @"\picwgoal" + (int)(((float)width / dpiX) * 1440)
        //                   + @"\pichgoal" + (int)(((float)height / dpiY) * 1440)
        //                   + " "
        //                   + BitConverter.ToString(stream.ToArray()).Replace("-", "")
        //                   + "}}";
        //    return proto;
        //}
        //[Flags]
        //private enum EmfToWmfBitsFlags
        //{
        //    EmfToWmfBitsFlagsDefault = 0x00000000,
        //    EmfToWmfBitsFlagsEmbedEmf = 0x00000001,
        //    EmfToWmfBitsFlagsIncludePlaceable = 0x00000002,
        //    EmfToWmfBitsFlagsNoXORClip = 0x00000004
        //}
        //private string MultiLangKiloByteUnit()
        //{
        //    return Properties.Resources.KiloByteUnit;
        //}

        //private string MultiLangMegaByteUnit()
        //{
        //    return Properties.Resources.MegaByteUnit;
        //}
        //[Serializable]
        //public struct ShellExecuteInfo
        //{
        //    public string Class;
        //    public string Directory;
        //    public string File;
        //    public IntPtr hkeyClass;
        //    public uint HotKey;
        //    public IntPtr hwnd;
        //    public IntPtr Icon;
        //    public IntPtr IDList;
        //    public IntPtr InstApp;
        //    public uint Mask;
        //    public IntPtr Monitor;
        //    public string Parameters;
        //    public uint Show;
        //    public int Size;
        //    public string Verb;
        //}
        //internal static class NativeStructs
        //{
        //    [StructLayout(LayoutKind.Sequential)]
        //    internal struct Input
        //    {
        //        public NativeEnums.SendInputEventType type;
        //        public MouseInput mouseInput;
        //    }

        //    [StructLayout(LayoutKind.Sequential)]
        //    internal struct MouseInput
        //    {
        //        public int dx;
        //        public int dy;
        //        public uint mouseData;
        //        public NativeEnums.MouseEventFlags dwFlags;
        //        public uint time;
        //        public IntPtr dwExtraInfo;
        //    }
        //}
        //internal static class NativeEnums
        //{
        //    [Flags]
        //    internal enum MouseEventFlags : uint
        //    {
        //        Move = 0x0001,
        //        LeftDown = 0x0002,
        //        LeftUp = 0x0004,
        //        RightDown = 0x0008,
        //        RightUp = 0x0010,
        //        MiddleDown = 0x0020,
        //        MiddleUp = 0x0040,
        //        XDown = 0x0080,
        //        XUp = 0x0100,
        //        Wheel = 0x0800,
        //        Absolute = 0x8000,
        //    }

        //    internal enum SendInputEventType : int
        //    {
        //        Mouse = 0,
        //        Keyboard = 1,
        //        Hardware = 2,
        //    }
        //}
        //private static bool GetNullableBoolFromSqlReader(SQLiteDataReader reader, string columnName)
        //{
        //    int columnIndex = reader.GetOrdinal(columnName);
        //    bool result;
        //    bool DefaultValue = false;
        //    if (!reader.IsDBNull(columnIndex))
        //        result = reader.GetBoolean(columnIndex);
        //    else
        //        result = DefaultValue;
        //    return result;
        //}
        //private void UpdateNewDBFieldsBackground(SQLiteCommand commandUpdate, string fieldsNeedUpdateText, string fieldsNeedSelectText, StringCollection patternNamesNeedUpdate)
        //{
        //    commandUpdate.CommandText = "UPDATE Clips SET " + fieldsNeedUpdateText + " WHERE Id = @Id";
        //    commandUpdate.Parameters.AddWithValue("Id", 0);
        //    SQLiteCommand commandSelect = new SQLiteCommand("", m_dbConnection);
        //    commandSelect.CommandText = "SELECT Id, Text " + fieldsNeedSelectText + " FROM Clips";
        //    using (SQLiteDataReader reader = commandSelect.ExecuteReader())
        //    {
        //        while (reader.Read())
        //        {
        //            if (stopUpdateDBThread)
        //                return;
        //            int ClipId = reader.GetInt32(reader.GetOrdinal("Id"));
        //            string plainText = reader.GetString(reader.GetOrdinal("Text"));
        //            commandUpdate.Parameters["id"].Value = ClipId;
        //            bool needWrite = false;
        //            foreach (string patternName in patternNamesNeedUpdate)
        //            {
        //                needWrite = Regex.IsMatch(plainText, TextPatterns[patternName], RegexOptions.IgnoreCase);
        //                commandUpdate.Parameters[patternName].Value = needWrite;
        //            }
        //            if (needWrite)
        //                commandUpdate.ExecuteNonQuery();
        //        }
        //    }
        //}

    }
}

/// <summary>
/// The enumeration of possible modifiers.
/// </summary>
[Flags]
public enum EnumModifierKeys : uint
{
    Alt = 1,
    Control = 2,
    Shift = 4,
    Win = 8
}

/// <summary>
/// Event Args for the event that is fired after the hot key has been pressed.
/// </summary>
public class KeyPressedEventArgs : EventArgs
{
    private Keys _key;
    private EnumModifierKeys _modifier;

    internal KeyPressedEventArgs(EnumModifierKeys modifier, Keys key)
    {
        _modifier = modifier;
        _key = key;
    }

    public Keys Key
    {
        get { return _key; }
    }

    public EnumModifierKeys Modifier
    {
        get { return _modifier; }
    }
}

public sealed class KeyboardHook : IDisposable
{
    // http://stackoverflow.com/questions/2450373/set-global-hotkeys-using-c-sharp

    private int _currentId;

    private Window _window = new Window();

    private ResourceManager resourceManager;

    public KeyboardHook(ResourceManager resourceManager)
    {
        this.resourceManager = resourceManager;
        // register the event of the inner native window.
        _window.KeyPressed += delegate (object sender, KeyPressedEventArgs args)
        {
            if (KeyPressed != null)
                KeyPressed(this, args);
        };
    }

    /// <summary>
    /// A hot key has been pressed.
    /// </summary>
    public event EventHandler<KeyPressedEventArgs> KeyPressed;

    public static string HotkeyTitle(Keys key, EnumModifierKeys modifier)
    {
        string hotkeyTitle = "";
        if ((modifier & EnumModifierKeys.Win) != 0)
            hotkeyTitle += Keys.Control.ToString() + " + ";
        if ((modifier & EnumModifierKeys.Control) != 0)
            hotkeyTitle += Keys.Control.ToString() + " + ";
        if ((modifier & EnumModifierKeys.Alt) != 0)
            hotkeyTitle += Keys.Alt.ToString() + " + ";
        if ((modifier & EnumModifierKeys.Shift) != 0)
            hotkeyTitle += Keys.Shift.ToString() + " + ";
        hotkeyTitle += key.ToString();
        return hotkeyTitle;
    }

    /// <summary>
    /// Registers a hot key in the system.
    /// </summary>
    /// <param name="modifier">The modifiers that are associated with the hot key.</param>
    /// <param name="key">The key itself that is associated with the hot key.</param>
    public void RegisterHotKey(EnumModifierKeys modifier, Keys key)
    {
        // increment the counter.
        _currentId = _currentId + 1;

        // register the hot key.
        if (!RegisterHotKey(_window.Handle, _currentId, (uint)modifier, (uint)key))
        {
            string hotkeyTitle = HotkeyTitle(key, modifier);
            //int ErrorCode = Marshal.GetLastWin32Error(); // 0 always
            //string errorText = resourceManager.GetString("CouldNotRegisterHotkey") + " \"" + hotkeyTitle + "\": " + ErrorCode;
            //string errorText = resourceManager.GetString("CouldNotRegisterHotkey") + " " + hotkeyTitle;
            
            MessageBox.Show("Cannot register key: "+hotkeyTitle, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            //throw new InvalidOperationException(hotkeyTitle);
            throw new Exception(hotkeyTitle);
        }
    }

    public void UnregisterHotKeys()
    {
        // unregister all the registered hot keys.
        for (int i = _currentId; i > 0; i--)
        {
            UnregisterHotKey(_window.Handle, i);
        }
    }

    // Registers a hot key with Windows.
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    // Unregisters the hot key with Windows.
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    /// <summary>
    /// Represents the window that is used internally to get the messages.
    /// </summary>
    private sealed class Window : NativeWindow, IDisposable
    {
        private const int WM_HOTKEY = 0x0312;

        public Window()
        {
            // create the handle for the window.
            this.CreateHandle(new CreateParams());
        }

        public event EventHandler<KeyPressedEventArgs> KeyPressed;

        /// <summary>
        /// Overridden to get the notifications.
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            // check if we got a hot key pressed.
            if (m.Msg == WM_HOTKEY)
            {
                // get the keys.
                Keys key = (Keys)(((int)m.LParam >> 16) & 0xFFFF);
                EnumModifierKeys modifier = (EnumModifierKeys)((int)m.LParam & 0xFFFF);

                // invoke the event to notify the parent.
                if (KeyPressed != null)
                    KeyPressed(this, new KeyPressedEventArgs(modifier, key));
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            this.DestroyHandle();
        }

        #endregion IDisposable Members
    }

    #region IDisposable Members

    public void Dispose()
    {
        UnregisterHotKeys();
        // dispose the inner native window.
        _window.Dispose();
    }

    #endregion IDisposable Members
}