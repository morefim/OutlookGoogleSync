using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OutlookGoogleSync
{
    public static class ByPassKlaLogoutPolicy
    {
        [Flags]
        public enum ExecutionState : uint
        {
            EsAwayModeRequired = 0x00000040,   
            EsContinuous       = 0x80000000, 
            EsDisplayRequired  = 0x00000010,
            EsSystemRequired   = 0x00000001,
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);

        private static ExecutionState SetStartup() 
        {
            var executionState = SetThreadExecutionState(ExecutionState.EsContinuous | ExecutionState.EsDisplayRequired);
            Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true)
                ?.SetValue("Bypass Logout Policy", Application.ExecutablePath);
            return executionState;
        }

    }
}
