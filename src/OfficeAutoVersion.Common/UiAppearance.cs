using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeAutoVersion.Core;

namespace OfficeAutoVersion.Common
{
    /// <summary>界面配色偏好：跟随 Office（默认）/ 强制浅色 / 强制深色。</summary>
    internal enum ThemePreference
    {
        Automatic = 0,
        Light = 1,
        Dark = 2,
    }

    /// <summary>
    /// 配色偏好的持久化：%APPDATA%\OfficeAutoVersion\ui.json。
    /// 单独一个文件而不是塞进 settings.json，因为它跟“哪个应用”无关，是一份全局偏好；
    /// 也避免和按应用分组的版本配置互相干扰。
    /// </summary>
    internal static class UiAppearance
    {
        private const string KeyAppearance = "appearance";

        private static bool _applied;
        private static ThemePreference _appliedPreference = ThemePreference.Automatic;

        /// <summary>测试用：覆盖配置文件路径。</summary>
        public static string OverrideFilePath { get; set; }

        /// <summary>
        /// 测试/截图用：为 true 时 <see cref="EnsureApplied"/> 不再按偏好覆盖调色板，
        /// 让调用方自己 <see cref="UiTheme.Apply"/>（渲染浅色/深色两套截图时用）。
        /// </summary>
        public static bool IgnorePreference { get; set; }

        public static string FilePath
        {
            get
            {
                return string.IsNullOrEmpty(OverrideFilePath)
                    ? Path.Combine(OptionsStore.AppDataFolder, "ui.json")
                    : OverrideFilePath;
            }
        }

        /// <summary>读取配色偏好；文件不存在或损坏时按“跟随 Office”处理。</summary>
        public static ThemePreference Load()
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    return ThemePreference.Automatic;
                }
                var root = MiniJson.Parse(File.ReadAllText(FilePath)) as Dictionary<string, object>;
                if (root == null)
                {
                    return ThemePreference.Automatic;
                }
                object raw;
                if (!root.TryGetValue(KeyAppearance, out raw) || raw == null)
                {
                    return ThemePreference.Automatic;
                }
                switch (Convert.ToString(raw).Trim().ToLowerInvariant())
                {
                    case "light":
                        return ThemePreference.Light;
                    case "dark":
                        return ThemePreference.Dark;
                    default:
                        return ThemePreference.Automatic;
                }
            }
            catch
            {
                return ThemePreference.Automatic;
            }
        }

        /// <summary>保存配色偏好。下次打开窗口时生效（当前已创建的控件不会重绘）。</summary>
        public static void Save(ThemePreference preference)
        {
            var map = new Dictionary<string, object> { { KeyAppearance, ToText(preference) } };
            var directory = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }
            File.WriteAllText(FilePath, MiniJson.Serialize(map), Encoding.UTF8);
            _applied = false; // 让下一个窗口重新取一次配色
        }

        /// <summary>
        /// 窗体创建前调用：把偏好落到 <see cref="UiTheme"/> 上。
        /// 每次打开窗口都会走一遍（读一个小文件），因此改完设置下次打开就生效。
        /// </summary>
        public static void EnsureApplied()
        {
            if (IgnorePreference)
            {
                return;
            }
            var preference = Load();
            if (_applied && _appliedPreference == preference)
            {
                return;
            }
            Apply(preference);
        }

        /// <summary>按偏好切换调色板：跟随 Office 时取注册表里的主题，强制时忽略注册表。</summary>
        public static void Apply(ThemePreference preference)
        {
            OfficeThemeKind kind;
            switch (preference)
            {
                case ThemePreference.Light:
                    kind = OfficeThemeKind.Colorful;
                    break;
                case ThemePreference.Dark:
                    kind = OfficeThemeKind.Black;
                    break;
                default:
                    kind = OfficeTheme.Kind;
                    break;
            }
            UiTheme.Apply(kind, OfficeTheme.HostAccent);
            _applied = true;
            _appliedPreference = preference;
        }

        /// <summary>界面上的文字说明（同时说明当前实际取到的是哪种主题）。</summary>
        public static string Describe(ThemePreference preference)
        {
            switch (preference)
            {
                case ThemePreference.Light:
                    return "始终使用浅色界面";
                case ThemePreference.Dark:
                    return "始终使用深色界面";
                default:
                    return "跟随 Office 主题（当前：" + UiTheme.ThemeName + "）";
            }
        }

        public static string ToText(ThemePreference preference)
        {
            switch (preference)
            {
                case ThemePreference.Light:
                    return "light";
                case ThemePreference.Dark:
                    return "dark";
                default:
                    return "auto";
            }
        }
    }
}
