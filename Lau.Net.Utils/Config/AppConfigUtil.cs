using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Config
{
    public static class AppConfigUtil
    {
        /// <summary>
        /// 获取当前应用程序配置文件
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static string GetValue(string key)
        {
            var value = ConfigurationManager.AppSettings[key];
            if (value == null)
            {
                return "";
            }
            return value;
        }

        public static void SetValue(string key, string value)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            if (config.AppSettings.Settings[key] != null)
            {
                config.AppSettings.Settings[key].Value = value;
            }
            else
            {
                config.AppSettings.Settings.Add(key, value);
            }

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static void SetValue(Dictionary<string, string> keyValues)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            foreach (var kvp in keyValues)
            {
                if (config.AppSettings.Settings[kvp.Key] != null)
                {
                    config.AppSettings.Settings[kvp.Key].Value = kvp.Value;
                }
                else
                {
                    config.AppSettings.Settings.Add(kvp.Key, kvp.Value);
                }
            }

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
    }
}
