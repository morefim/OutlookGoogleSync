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
                var message = String.Format(format, args);
                return RegisterWindowMessage(message);
            }

            public const int HWND_BROADCAST = 0xffff;
            public const int SW_SHOWNORMAL = 1;

            [DllImport("user32")]
            public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            public static void ShowToFront(IntPtr window)
            {
                ShowWindow(window, SW_SHOWNORMAL);
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
                        return String.Empty;
                    }
                    return ((GuidAttribute)attributes[0]).Value;

                    /*//Assembly asm = Assembly.GetExecutingAssembly();
                    //return asm.GetType().GUID.ToString();
                    return "{3CFC1234-40E7-4988-AC4A-86BF004706AC}";*/
                }
            }
        } 

        public static readonly int WM_SHOWFIRSTINSTANCE = WinApi.RegisterWindowMessage("WM_SHOWFIRSTINSTANCE|{0}", ProgramInfo.AssemblyGuid);
        static Mutex mutex;

        static public bool Start()
        {
            var onlyInstance = false;
            var mutexName = String.Format("Local\\{0}", ProgramInfo.AssemblyGuid);

            // if you want your app to be limited to a single instance
            // across ALL SESSIONS (multiple users & terminal services), then use the following line instead:
            // string mutexName = String.Format("Global\\{0}", ProgramInfo.AssemblyGuid);

            mutex = new Mutex(true, mutexName, out onlyInstance);
            return onlyInstance;
        }

        static public void ShowFirstInstance()
        {
            WinApi.PostMessage((IntPtr)WinApi.HWND_BROADCAST, WM_SHOWFIRSTINSTANCE, IntPtr.Zero, IntPtr.Zero);
        }

        static public void Stop()
        {
            mutex.ReleaseMutex();
        }
    }
}
