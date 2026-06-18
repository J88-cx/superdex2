using GestureSign.Common.Configuration;
using GestureSign.Common.Localization;
using IWshRuntimeLibrary;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using File = System.IO.File;

namespace GestureSign.ControlPanel.Common
{
    static class StartupHelper
    {
        private static string DaemonPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, GestureSign.Common.Constants.DaemonFileName);

        private static string StartupLnkPath => Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + GestureSign.Common.Constants.ProductName + ".lnk";

        public static bool IsRunAsAdmin => AppConfig.RunAsAdmin;

        private static string GetLnkTargetPath(string filepath)
        {
            using (var br = new BinaryReader(File.OpenRead(filepath)))
            {
                br.ReadBytes(0x14);
                uint lflags = br.ReadUInt32();
                if ((lflags & 0x01) == 1)
                {
                    br.ReadBytes(0x34);
                    var skip = br.ReadUInt16();
                    br.ReadBytes(skip);
                }
                var length = br.ReadUInt32();
                br.ReadBytes(0x0C);
                var lbpos = br.ReadUInt32();
                br.ReadBytes((int)lbpos - 0x14);
                var size = length - lbpos - 0x02;
                var bytePath = br.ReadBytes((int)size);
                int index = Array.IndexOf(bytePath, (byte)0x00);
                var path = index < 0 ? System.Text.Encoding.Default.GetString(bytePath, 0, bytePath.Length) :
                    System.Text.Encoding.Unicode.GetString(bytePath, index + 1, bytePath.Length - index - 1);
                return path;
            }
        }

        private static void CreateLnk(string lnkPath, string targetPath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortCut = (IWshShortcut)shell.CreateShortcut(lnkPath);
            shortCut.TargetPath = targetPath;
            shortCut.WindowStyle = 7;
            shortCut.Arguments = "";
            shortCut.Description = Application.ResourceAssembly.GetName().Version.ToString();
            shortCut.Save();
        }

        private static bool AddStartupTask(string filePath)
        {
            try
            {
                string taskXml = Properties.Resources.StartGestureSignTask.Replace("GestureSignFilePath", filePath);
                string xmlFilePath = Path.Combine(AppConfig.LocalApplicationDataPath, "StartGestureSignTask.xml");
                File.WriteAllText(xmlFilePath, taskXml, System.Text.Encoding.Unicode);

                using (Process schtasks = new Process())
                {
                    string arguments = string.Format(" /create /tn StartGestureSign /f /xml \"{0}\"", xmlFilePath);
                    schtasks.StartInfo = new ProcessStartInfo("schtasks.exe", arguments)
                    {
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        UseShellExecute = true,
                        Verb = "runas",
                    };
                    schtasks.Start();
                    schtasks.WaitForExit();
                }
                if (File.Exists(xmlFilePath))
                    File.Delete(xmlFilePath);
            }
            catch (Exception exception)
            {
                GestureSign.Common.Log.Logging.LogAndNotice(exception);
                return false;
            }

            return true;
        }
        private static bool DelStartupTask()
        {
            try
            {
                using (Process schtasks = new Process())
                {
                    schtasks.StartInfo = new ProcessStartInfo("schtasks.exe", " /delete /tn StartGestureSign /f")
                    {
                        CreateNoWindow = true,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        UseShellExecute = true,
                        Verb = "runas",
                    };
                    schtasks.Start();
                    schtasks.WaitForExit();
                }
            }
            catch (Exception exception)
            {
                GestureSign.Common.Log.Logging.LogAndNotice(exception);
                return false;
            }

            return true;
        }

        // 以下方法原本使用 Windows.ApplicationModel.StartupTask，但 WinRT API 在 WPF 中不可用，已注释处理。
        public static Task<bool> CheckStoreAppStartupStatus()
        {
            // return await StartupTask.GetAsync("GestureSignTask")...
            return Task.FromResult(false);
        }

        public static Task<bool> EnableStoreAppStartup()
        {
            // return await StartupTask.RequestEnableAsync()...
            return Task.FromResult(false);
        }

        public static Task<bool> DisableStoreAppStartup()
        {
            // startupTask.Disable();
            return Task.FromResult(true);
        }
        public static bool GetStartupStatus()
        {
            try
            {
                string startupLnkPath = StartupLnkPath;
                if (File.Exists(startupLnkPath))
                {
                    var targetPath = GetLnkTargetPath(startupLnkPath);
                    var daemonPath = DaemonPath;
                    if (daemonPath != targetPath)
                    {
                        CreateLnk(startupLnkPath, daemonPath);
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, LocalizationProvider.Instance.GetTextValue("Messages.Error"), MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        public static bool EnableNormalStartup()
        {
            CreateLnk(StartupLnkPath, DaemonPath);
            return true;
        }

        public static bool DisableNormalStartup()
        {
            if (File.Exists(StartupLnkPath))
            {
                try
                {
                    File.Delete(StartupLnkPath);
                }
                catch (Exception exception)
                {
                    GestureSign.Common.Log.Logging.LogAndNotice(exception);
                    return false;
                }
            }
            return true;
        }

        public static bool EnableHighPrivilegeStartup()
        {
            return AddStartupTask(DaemonPath);
        }

        public static bool DisableHighPrivilegeStartup()
        {
            return DelStartupTask();
        }
    }
}
