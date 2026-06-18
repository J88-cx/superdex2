using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml;
using GestureSign.Common.Localization;
using GestureSign.Common.Log;
using Microsoft.Win32;

namespace GestureSign.CorePlugins.LaunchApp
{
    public partial class LaunchAppView : UserControl
    {
        public LaunchAppView()
        {
            InitializeComponent();
        }

        [DllImport("shlwapi.dll", BestFitMapping = false, CharSet = CharSet.Unicode, ExactSpelling = true,
            SetLastError = false, ThrowOnUnmappableChar = true)]
        private static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf,
            IntPtr ppvReserved);

        public KeyValuePair<string, string> SelectedAppInfo
        {
            get
            {
                Model model = comboBox.SelectedItem as Model;
                return model?.AppInfo ?? new KeyValuePair<string, string>();
            }
        }

        private void comboBox_Loaded(object sender, RoutedEventArgs e)
        {
            var apps = new List<Model>(10);

            var getAppsTask = Task.Run(() =>
            {
                if (comboBox.ItemsSource != null) return;

                var appXInfos = GetAppXInfosFromReg();
                if (appXInfos == null || appXInfos.Count == 0) return;

                foreach (var kvp in appXInfos)
                {
                    try
                    {
                        string appUserModelId = kvp.Key;
                        var appXInfo = kvp.Value;

                        string displayName = ExtractDisplayNameFromResource(appXInfo.ApplicationName);
                        if (string.IsNullOrEmpty(displayName)) continue;

                        string logoPath = ExtractDisplayIconFromResource(appXInfo.ApplicationIcon);
                        if (string.IsNullOrEmpty(logoPath)) continue;

                        var model = new Model
                        {
                            AppInfo = new KeyValuePair<string, string>(appUserModelId, displayName),
                            Logo = logoPath,
                            BackgroundColor = SystemParameters.WindowGlassBrush
                        };
                        model.BackgroundColor.Freeze();
                        apps.Add(model);
                    }
                    catch (Exception ex)
                    {
                        Logging.LogException(ex);
                        continue;
                    }
                }
            });
            getAppsTask.ContinueWith(task =>
            {
                string message = null;
                if (task.Exception != null)
                {
                    foreach (var item in task.Exception.InnerExceptions)
                    {
                        Logging.LogException(item);
                        message = item.Message;
                    }
                }

                comboBox.Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (message != null)
                    {
                        TipTextBlock.Text = message;
                        return;
                    }
                    if (apps.Count != 0)
                        comboBox.ItemsSource = apps.OrderBy(app => app.AppInfo.Value).ToList();

                    var launchApp = DataContext as LaunchApp;
                    if (launchApp != null)
                    {
                        comboBox.SelectedItem = (comboBox.ItemsSource as List<Model>)?.Find(m => m.AppInfo.Equals(launchApp.AppInfo));
                    }
                    TipTextBlock.Text = LocalizationProvider.Instance.GetTextValue("CorePlugins.LaunchApp.Tip");
                    comboBox.Visibility = Visibility.Visible;
                }));
            });
        }

        private string ExtractDisplayNameFromResource(string resourceKey)
        {
            if (string.IsNullOrEmpty(resourceKey)) return string.Empty;
            var outBuff = new StringBuilder(128);
            SHLoadIndirectString(resourceKey, outBuff, outBuff.Capacity, IntPtr.Zero);
            return outBuff.ToString();
        }

        private string ExtractDisplayIconFromResource(string logoPath)
        {
            if (string.IsNullOrEmpty(logoPath)) return null;
            if (File.Exists(logoPath)) return logoPath;

            var scale100LogoPath = Path.ChangeExtension(logoPath, "scale-100.png");
            if (File.Exists(scale100LogoPath)) return scale100LogoPath;

            var directory = Path.GetDirectoryName(logoPath);
            var pattern = Path.GetFileNameWithoutExtension(logoPath) + ".*" + Path.GetExtension(logoPath);
            if (Directory.Exists(directory))
            {
                try
                {
                    var logoPaths = Directory.GetFiles(directory, pattern);
                    if (logoPaths.Length != 0)
                        return logoPaths[0];
                }
                catch { }
            }

            var localized = Path.Combine(Path.GetDirectoryName(logoPath), "en-us", Path.GetFileName(logoPath));
            if (localized == null) return null;
            localized = Path.ChangeExtension(localized, "scale-100.png");
            if (File.Exists(localized)) return localized;

            return null;
        }
        private Dictionary<string, AppXInfo> GetAppXInfosFromReg()
        {
            var dic = new Dictionary<string, AppXInfo>(10);

            using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Classes\"))
            {
                if (key == null) return null;
                var appKeys = from k in key.GetSubKeyNames()
                              where k.StartsWith("AppX")
                              select k;
                foreach (var appKey in appKeys)
                {
                    using (var appRegKey = key.OpenSubKey(appKey))
                    {
                        if (appRegKey == null) continue;
                        using (var applicationKey = appRegKey.OpenSubKey("Application"))
                        {
                            if (applicationKey == null) continue;
                            var appUserModelId = applicationKey.GetValue("AppUserModelId")?.ToString();
                            var applicationIcon = applicationKey.GetValue("ApplicationIcon")?.ToString();
                            var applicationName = applicationKey.GetValue("ApplicationName")?.ToString();

                            if (!string.IsNullOrEmpty(appUserModelId) && !dic.ContainsKey(appUserModelId))
                                dic.Add(appUserModelId, new AppXInfo
                                {
                                    ApplicationIcon = applicationIcon,
                                    ApplicationName = applicationName
                                });
                        }
                    }
                }
            }
            return dic;
        }

        private string GetAppUserModelId(string packageFullName)
        {
            var str = string.Empty;
            using (var key = Registry.CurrentUser.OpenSubKey(
                $@"SOFTWARE\Classes\ActivatableClasses\Package\{packageFullName}\Server\"))
            {
                if (key == null) return str;

                var appKeys = from k in key.GetSubKeyNames()
                              where k.StartsWith("Appex")
                              select k;

                foreach (var appKey in appKeys)
                {
                    using (var serverKey = key.OpenSubKey(appKey))
                    {
                        if (serverKey?.GetValue("AppUserModelId") == null) continue;
                        str = serverKey.GetValue("AppUserModelId").ToString();
                        break;
                    }
                }
            }
            return str;
        }

        private class AppXInfo
        {
            public string ApplicationIcon { get; set; }
            public string ApplicationName { get; set; }
        }

        private class Model
        {
            public string Logo { get; set; }
            public KeyValuePair<string, string> AppInfo { get; set; }
            public Brush BackgroundColor { get; set; }
        }

        private struct AppInfo
        {
            public string ID { get; set; }
            public string BackgroundColor { get; set; }
            public string DisplayName { get; set; }
            public string Logo { get; set; }
        }
    }
}
