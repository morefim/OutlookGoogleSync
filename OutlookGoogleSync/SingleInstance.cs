using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;

namespace OutlookGoogleSync
{
    static public class SingleInstance
    {
        static public class WinApi
        {
            [DllImport("user32")]
            public static extern int RegisterWindowMessage(string message);

            public static int RegisterWindowMessage(string format, params object[] args)
            {
                var message = string.Format(format, args);
                return RegisterWindowMessage(message);
            }

            public const int HwndBroadcast = 0xffff;
            public const int SwShownormal = 1;

            [DllImport("user32")]
            public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            public static void ShowToFront(IntPtr window)
            {
                ShowWindow(window, SwShownormal);
                SetForegroundWindow(window);
            }
        }

        static public class ProgramInfo
        {
            static public string AssemblyGuid
            {
                get
                {
                    var attributes = Assembly.GetEntryAssembly().GetCustomAttributes(typeof(GuidAttribute), false);
                    if (attributes.Length == 0)
                    {
                        return string.Empty;
                    }
                    return ((GuidAttribute)attributes[0]).Value;

                    /*//Assembly asm = Assembly.GetExecutingAssembly();
                    //return asm.GetType().GUID.ToString();
                    return "{3CFC1234-40E7-4988-AC4A-86BF004706AC}";*/
                }
            }
        } 

        public static readonly int WmShowfirstinstance = WinApi.RegisterWindowMessage("WM_SHOWFIRSTINSTANCE|{0}", ProgramInfo.AssemblyGuid);
        static Mutex _mutex;

        static public bool Start()
        {
            var onlyInstance = false;
            var mutexName = $"Local\\{ProgramInfo.AssemblyGuid}";

            // if you want your app to be limited to a single instance
            // across ALL SESSIONS (multiple users & terminal services), then use the following line instead:
            // string mutexName = String.Format("Global\\{0}", ProgramInfo.AssemblyGuid);

            _mutex = new Mutex(true, mutexName, out onlyInstance);
            return onlyInstance;
        }

        static public void ShowFirstInstance()
        {
            WinApi.PostMessage((IntPtr)WinApi.HwndBroadcast, WmShowfirstinstance, IntPtr.Zero, IntPtr.Zero);
        }

        static public void Stop()
        {
            _mutex.ReleaseMutex();
        }
    }
}
