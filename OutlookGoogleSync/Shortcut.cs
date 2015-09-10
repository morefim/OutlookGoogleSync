using System;
using IWshRuntimeLibrary;
using Shell32;
using System.Windows.Forms;
using System.IO;
using File = System.IO.File;

namespace OutlookGoogleSync
{
    class Shortcut
    {
        // Make sure you use try/catch block because your App may has no permissions on the target path!
        public void Test()
        {
            try
            {
                CreateShortcut(@"C:\temp", @"C:\MyShortcutFile.lnk", "Custom Shortcut", "/param", "Ctrl+F", @"c:\");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Create Windows Shorcut
        /// </summary>
        /// <param name="SourceFile">A file you want to make shortcut to</param>
        /// <param name="ShortcutFile">Path and shorcut file name including file extension (.lnk)</param>
        public static void CreateShortcut(string SourceFile, string ShortcutFile)
        {
            CreateShortcut(SourceFile, ShortcutFile, null, null, null, null);
        }

        /// <summary>
        /// Create Windows Shorcut
        /// </summary>
        /// <param name="SourceFile">A file you want to make shortcut to</param>
        /// <param name="ShortcutFile">Path and shorcut file name including file extension (.lnk)</param>
        /// <param name="Description">Shortcut description</param>
        /// <param name="Arguments">Command line arguments</param>
        /// <param name="HotKey">Shortcut hot key as a string, for example "Ctrl+F"</param>
        /// <param name="WorkingDirectory">"Start in" shorcut parameter</param>
        public static void CreateShortcut(string SourceFile, string ShortcutFile, string Description, string Arguments, string HotKey, string WorkingDirectory)
        {
            // Check necessary parameters first:
            if (String.IsNullOrEmpty(SourceFile))
                throw new ArgumentNullException("SourceFile");
            if (String.IsNullOrEmpty(ShortcutFile))
                throw new ArgumentNullException("ShortcutFile");

            // Create WshShellClass instance:
            var wshShell = new WshShellClass();

            // Create shortcut object:
            var shorcut = (IWshShortcut)wshShell.CreateShortcut(ShortcutFile);

            // Assign shortcut properties:
            shorcut.TargetPath = SourceFile;
            shorcut.Description = Description;
            if (!String.IsNullOrEmpty(Arguments))
                shorcut.Arguments = Arguments;
            if (!String.IsNullOrEmpty(HotKey))
                shorcut.Hotkey = HotKey;
            if (!String.IsNullOrEmpty(WorkingDirectory))
                shorcut.WorkingDirectory = WorkingDirectory;

            // Save the shortcut:
            shorcut.Save();
        }

        public static void CreateStartupFolderShortcut()
        {
            if (IsStartupFolderShortcutExists())
                return;

            var wshShell = new WshShellClass();
            IWshShortcut shortcut;
            var startUpFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);

            // Create the shortcut
            shortcut = (IWshShortcut)wshShell.CreateShortcut(startUpFolderPath + "\\" + Application.ProductName + ".lnk");

            shortcut.TargetPath = Application.ExecutablePath;
            shortcut.WorkingDirectory = Application.StartupPath;
            shortcut.Description = "Launch " + Application.ProductName;
            // shortcut.IconLocation = Application.StartupPath + @"\App.ico";
            shortcut.Save();
        }

        public static string GetShortcutTargetFile(string shortcutFilename)
        {
            if (shortcutFilename.Length == 0)
                shortcutFilename = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + Application.ProductName + ".lnk";

            var pathOnly = Path.GetDirectoryName(shortcutFilename);
            var filenameOnly = Path.GetFileName(shortcutFilename);

            Shell shell = new ShellClass();
            var folder = shell.NameSpace(pathOnly);
            var folderItem = folder.ParseName(filenameOnly);
            if (folderItem != null)
            {
                var link = (ShellLinkObject)folderItem.GetLink;
                return link.Path;
            }

            return String.Empty; // Not found
        }

        public static bool IsStartupFolderShortcutExists()
        {
            return GetShortcutTargetFile(string.Empty) == Application.ExecutablePath;
        }

        public static void DeleteStartupFolderShortcuts(string targetExeName)
        {
            var startUpFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);

            var di = new DirectoryInfo(startUpFolderPath);
            var files = di.GetFiles("*.lnk");

            foreach (var fi in files)
            {
                var shortcutTargetFile = GetShortcutTargetFile(fi.FullName);

                if (shortcutTargetFile.EndsWith(targetExeName, StringComparison.InvariantCultureIgnoreCase))
                {
                    File.Delete(fi.FullName);
                }
            }
        }
    }
}
