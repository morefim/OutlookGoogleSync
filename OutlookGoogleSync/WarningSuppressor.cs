﻿using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace OutlookGoogleSync
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class WarningSuppressor : IDisposable
    {
        private bool _dispose;
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        /// <summary>
        /// Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.
        /// </summary>
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption(IntPtr zeroOnly, string lpWindowName);

        //[DllImport("user32.dll", CharSet = CharSet.Auto)]
        //private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        //[DllImport("user32.dll", CharSet = CharSet.Auto)]
        //public static extern int SendMessage(IntPtr hWnd, uint msg, int wParam, IntPtr lParam);

        //[DllImport("user32.dll", SetLastError = true)]
        //public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        private const uint WM_CLOSE = 0x0010;
        private const int WM_LBUTTONDOWN = 0x201;
        private const int WM_LBUTTONUP = 0x202;
        private const int MK_LBUTTON = 1;

        private const int AllowButtonID = 0x12A6;
        private const int AllowCheckboxID = 0x12A3;
        private const int AllowComboboxID = 0x12A5;

        private const string WindowClassName = "#32770";

        const int WM_COMMAND = 0x0111;
        const int BN_CLICKED = 0x0000;
        private const int BM_CLICK = 0x00F5;
        const string fn = @"C:\Windows\system32\calc.exe";

        [DllImport("user32.dll")]
        static extern IntPtr GetDlgItem(IntPtr hWnd, int nIDDlgItem);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SetActiveWindow(IntPtr hWnd);

        public WarningSuppressor()
        {
            _dispose = false;
            Task.Run(WaitForMessageBoxAndSuppress);
        }

        public void Dispose()
        {
            _dispose = true;
        }

        private void WaitForMessageBoxAndSuppress()
        {
            while (!_dispose)
            {
                IntPtr windowPtr = FindWindowByCaption(IntPtr.Zero, "Microsoft Outlook");
                if (windowPtr == IntPtr.Zero)
                {
                    Debug.WriteLine("Window not found");
                    Task.Delay(50);
                    continue;
                }

                Debug.WriteLine("Window found, handle ID: {0:X}", windowPtr);
                IntPtr hWndButton = GetDlgItem(windowPtr, AllowButtonID);
                if (hWndButton == IntPtr.Zero)
                {
                    Debug.WriteLine("Button not found");
                    continue;
                }
                Debug.WriteLine("Button found, handle ID: {0:X}", hWndButton);

                //int wParam = (BN_CLICKED << 16) | (AllowButtonID & 0xffff);
                //IntPtr sentStatus = SendMessage(windowPtr, BN_CLICKED, wParam, hWndButton);
                //Debug.WriteLine("Message sent, status {0:X}", sentStatus);

                IntPtr activeWindow = SetActiveWindow(windowPtr);
                Debug.WriteLine("Active Window set, handle ID: {0:X}", hWndButton);

                //IntPtr sentStatus = SendMessage(hWndButton, BM_CLICK, 0, IntPtr.Zero);
                //IntPtr sentStatus = SendMessage(windowPtr, WM_COMMAND, AllowButtonID, hWndButton);

                //SendMessage(hWndButton, WM_LBUTTONDOWN, MK_LBUTTON, IntPtr.Zero);
                //IntPtr sentStatus = SendMessage(hWndButton, WM_LBUTTONUP, MK_LBUTTON, IntPtr.Zero);
                
                //Debug.WriteLine("Message sent, status {0:X}", sentStatus);
            }
        }
    }

} 