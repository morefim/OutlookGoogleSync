using System;
using IWshRuntimeLibrary;
using System.Windows.Forms;
using System.IO;
using File = System.IO.File;

namespace OutlookGoogleSync
{
    internal class Shortcut
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
        /// Create Windows Shortcut
        /// </summary>
        /// <param name="sourceFile">A file you want to make shortcut to</param>
        /// <param name="shortcutFile">Path and shortcut file name including file extension (.lnk)</param>
        public static void CreateShortcut(string sourceFile, string shortcutFile)
        {
            CreateShortcut(sourceFile, shortcutFile, null, null, null, null);
        }

        /// <summary>
        /// Create Windows Shortcut
        /// </summary>
        /// <param name="sourceFile">A file you want to make shortcut to</param>
        /// <param name="shortcutFile">Path and shortcut file name including file extension (.lnk)</param>
        /// <param name="description">Shortcut description</param>
        /// <param name="arguments">Command line arguments</param>
        /// <param name="hotKey">Shortcut hot key as a string, for example "Ctrl+F"</param>
        /// <param name="workingDirectory">"Start in" shortcut parameter</param>
        public static void CreateShortcut(string sourceFile, string shortcutFile, string description, string arguments, string hotKey, string workingDirectory)
        {
            // Check necessary parameters first:
            if (string.IsNullOrEmpty(sourceFile))
                throw new ArgumentNullException(nameof(sourceFile));
            if (string.IsNullOrEmpty(shortcutFile))
                throw new ArgumentNullException(nameof(shortcutFile));

            // Create WshShellClass instance:
            var wshShell = new WshShellClass();

            // Create shortcut object:
            var shortcut = (IWshShortcut)wshShell.CreateShortcut(shortcutFile);

            // Assign shortcut properties:
            shortcut.TargetPath = sourceFile;
            shortcut.Description = description;
            if (!string.IsNullOrEmpty(arguments))
                shortcut.Arguments = arguments;
            if (!string.IsNullOrEmpty(hotKey))
                shortcut.Hotkey = hotKey;
            if (!string.IsNullOrEmpty(workingDirectory))
                shortcut.WorkingDirectory = workingDirectory;

            // Save the shortcut:
            shortcut.Save();
        }

        public static void CreateStartupFolderShortcut()
        {
            if (IsStartupFolderShortcutExists())
                return;

            var wshShell = new WshShellClass();
            var startUpFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);

            // Create the shortcut
            var shortcut = (IWshShortcut)wshShell.CreateShortcut(startUpFolderPath + "\\" + Application.ProductName + ".lnk");

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

            if (!File.Exists(shortcutFilename)) return string.Empty;
            // WshShellClass shell = new WshShellClass();
            var shell = new WshShell(); //Create a new WshShell Interface
            var link = (IWshShortcut)shell.CreateShortcut(shortcutFilename); //Link the interface to our shortcut

            return link.TargetPath; //Show the target in a MessageBox using IWshShortcut
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
