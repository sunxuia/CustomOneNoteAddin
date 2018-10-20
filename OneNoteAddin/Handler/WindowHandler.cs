using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OneNoteAddin.Handler
{
    public class WindowHandler
    {
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("USER32", SetLastError = true)]
        static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        private delegate int HookProc(int nCode, int wParam, IntPtr lParam);

        public IntPtr Handle { get; set; }

        public IntPtr ThisHandle { get; private set; }

        // 遍历窗口句柄
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(WNDENUMPROC lpEnumFunc, int lParam);

        // 委托, 第一个参数是窗口句柄, 第二个是EnumWindows 中传入的lParam
        private delegate bool WNDENUMPROC(IntPtr hWnd, int lParam);

        // 获得窗口句柄的类名
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        // 获得窗口句柄的标题
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern int GetWindowTextA(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        // 获得窗口句柄的进程号
        [DllImport("user32")]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int pid);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);


        public bool SetWindow(string className, string title)
        {
            // 遍历所有窗口句柄
            EnumWindows((hWnd, lParam) =>
            {
                StringBuilder sb = new StringBuilder();
                GetClassName(hWnd, sb, className.Length + 1); // 获得窗口句柄的列名
                if (sb.ToString() == className)
                {
                    sb.Clear();
                    // 获得窗口标题
                    GetWindowTextA(hWnd, sb, 1024);
                    // 转换标题的编码
                    string windowTitle = Encoding.UTF8.GetString(Encoding.Unicode.GetBytes(sb.ToString()));
                    if (windowTitle.Contains(title))
                    {
                        Handle = hWnd;
                        return false;
                    }
                }
                return true;
            }, 0);
            return Handle != IntPtr.Zero;
        }

        public static bool IsCapsLockPressed()
        {
            return GetKeyState(20) == 1;
        }

        public static void SendCapsLock()
        {
            bool caps = IsCapsLockPressed();
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            const int KEYEVENTF_KEYUP = 0x2;
            keybd_event(0x14, 0x45, KEYEVENTF_EXTENDEDKEY, (UIntPtr)0);
            keybd_event(0x14, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, (UIntPtr)0);
            while (caps == IsCapsLockPressed())
            {
                Thread.Sleep(100);
            }
        }

        public WindowHandler()
        {
            ThisHandle = GetForegroundWindow();
        }

        public void ShowWindow(bool show)
        {
            ShowWindow(Handle, show ? 1u : 0);
        }

        public void ActiveWindow()
        {
            SetForegroundWindow(Handle);
            while (GetForegroundWindow() != Handle)
            {
                Thread.Sleep(100);
            }
        }

        public void DeactiveWindow()
        {
            SetForegroundWindow(ThisHandle);
            while (GetForegroundWindow() != ThisHandle)
            {
                Thread.Sleep(100);
            }
        }

        public void CloseWindow()
        {
            SendMessage(Handle, 0x10, 0, 0);
        }

        public bool IsWindowClosed()
        {
            try
            {
                return !IsWindow(Handle);
            }
            catch (Exception)
            {
                return true;
            }
        }

        public bool IsWindowVisible()
        {
            return IsWindowVisible(Handle);
        }
    }
}
