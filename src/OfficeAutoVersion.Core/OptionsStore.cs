using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 配置的持久化：%APPDATA%\OfficeAutoVersion\settings.json，
    /// 形如 { "Word": {...}, "Excel": {...}, "PowerPoint": {...} }。
    /// 使用内置的轻量 JSON 读写，无需第三方依赖。
    /// </summary>
    public static class OptionsStore
    {
        /// <summary>插件数据目录（%APPDATA%\OfficeAutoVersion）。</summary>
        public static string AppDataFolder
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficeAutoVersion");
            }
        }

        /// <summary>默认设置文件路径。</summary>
        public static string SettingsFilePath
        {
            get { return Path.Combine(AppDataFolder, "settings.json"); }
        }

        /// <summary>测试/高级场景用：覆盖设置文件路径（为空表示用默认路径）。</summary>
        public static string OverrideSettingsPath { get; set; }

        private static string FilePath
        {
            get
            {
                return string.IsNullOrEmpty(OverrideSettingsPath) ? SettingsFilePath : OverrideSettingsPath;
            }
        }

        /// <summary>读取某个应用的配置；不存在或解析失败时返回默认配置。</summary>
        public static VersionOptions Load(string appKey)
        {
            try
            {
                var all = LoadAll();
                VersionOptions options;
                if (all.TryGetValue(appKey ?? string.Empty, out options) && options != null)
                {
                    return options;
                }
            }
            catch
            {
                // 忽略，返回默认值
            }
            return new VersionOptions();
        }

        /// <summary>保存某个应用的配置（不影响其它应用）。</summary>
        public static void Save(string appKey, VersionOptions options)
        {
            var all = LoadAll();
            all[appKey ?? string.Empty] = options ?? new VersionOptions();
            var directory = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }
            File.WriteAllText(FilePath, MiniJson.Serialize(all), Encoding.UTF8);
        }

        /// <summary>读取全部应用的配置。</summary>
        public static Dictionary<string, VersionOptions> LoadAll()
        {
            var result = new Dictionary<string, VersionOptions>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (File.Exists(FilePath))
                {
                    var root = MiniJson.Parse(File.ReadAllText(FilePath)) as Dictionary<string, object>;
                    if (root != null)
                    {
                        foreach (var pair in root)
                        {
                            var options = MiniJson.ToVersionOptions(pair.Value as Dictionary<string, object>);
                            if (options != null && !string.IsNullOrEmpty(pair.Key))
                            {
                                result[pair.Key] = options;
                            }
                        }
                    }
                }
            }
            catch
            {
                // 配置损坏时按空处理，避免插件无法启动
            }
            return result;
        }
    }
}
