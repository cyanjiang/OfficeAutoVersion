using System;
using System.Diagnostics;
using System.Drawing;
using Microsoft.Win32;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// Office 的界面主题（“文件 - 选项 - 常规 - Office 主题”）。括号里的数值是注册表里的取值。
    /// </summary>
    internal enum OfficeThemeKind
    {
        /// <summary>彩色（默认，浅色界面 + 各应用品牌色）。</summary>
        Colorful = 0,

        /// <summary>白色（最亮的浅色界面）。</summary>
        White = 3,

        /// <summary>深灰色（深色界面）。</summary>
        DarkGray = 4,

        /// <summary>黑色（最深，纯黑画布）。</summary>
        Black = 5,
    }

    /// <summary>
    /// 读取宿主 Office 的界面主题与品牌色，让插件对话框跟 Office 自身配色一致：
    ///   • 浅色 / 深色：HKCU\Software\Microsoft\Office\16.0\Common 的 "UI Theme"
    ///   • 主色：Word 蓝 #185ABD、Excel 绿 #107C41、PowerPoint 橙红 #C43E1C
    /// 任何一步失败都退回“Office 默认浅色 + Office 蓝”，绝不影响功能。
    /// </summary>
    internal static class OfficeTheme
    {
        private const string CommonKeyPath = @"Software\Microsoft\Office\16.0\Common";
        private const string ThemeValueName = "UI Theme";

        /// <summary>Office 通用蓝（读不到宿主、或非三大应用时使用）。</summary>
        public static readonly Color DefaultAccent = Color.FromArgb(0x0F, 0x6C, 0xBD);

        private static bool _detected;
        private static OfficeThemeKind _kind = OfficeThemeKind.Colorful;
        private static string _source = "Office 默认（浅色）";

        /// <summary>当前 Office 主题。</summary>
        public static OfficeThemeKind Kind
        {
            get
            {
                DetectOnce();
                return _kind;
            }
        }

        /// <summary>当前是否为深色主题（深灰 / 黑色）。</summary>
        public static bool IsDark
        {
            get { return IsDarkKind(Kind); }
        }

        /// <summary>主题来源说明，用于设置界面与日志（便于排查“为什么是深色”）。</summary>
        public static string Source
        {
            get
            {
                DetectOnce();
                return _source;
            }
        }

        public static bool IsDarkKind(OfficeThemeKind kind)
        {
            return kind == OfficeThemeKind.DarkGray || kind == OfficeThemeKind.Black;
        }

        /// <summary>宿主应用名（Word / Excel / PowerPoint）。取不到返回 null。</summary>
        public static string HostName
        {
            get
            {
                try
                {
                    switch (HostProcessName())
                    {
                        case "WINWORD": return "Word";
                        case "EXCEL": return "Excel";
                        case "POWERPNT": return "PowerPoint";
                        default: return null;
                    }
                }
                catch
                {
                    return null;
                }
            }
        }

        /// <summary>宿主应用的品牌色（Office 2021/2024 的图标主色）。</summary>
        public static Color HostAccent
        {
            get
            {
                try
                {
                    return AccentFor(HostProcessName());
                }
                catch
                {
                    return DefaultAccent;
                }
            }
        }

        /// <summary>按进程名取品牌色。Word/Excel/PowerPoint 各用自家的主色，其它一律 Office 蓝。</summary>
        internal static Color AccentFor(string processName)
        {
            switch ((processName ?? string.Empty).ToUpperInvariant())
            {
                case "WINWORD":
                    return Color.FromArgb(0x18, 0x5A, 0xBD); // Word 蓝
                case "EXCEL":
                    return Color.FromArgb(0x10, 0x7C, 0x41); // Excel 绿
                case "POWERPNT":
                    return Color.FromArgb(0xC4, 0x3E, 0x1C); // PowerPoint 橙红
                default:
                    return DefaultAccent;
            }
        }

        /// <summary>
        /// Office 注册表取值映射：0=彩色（默认）3=白色 4=深灰色 5=黑色。
        /// 未知取值按浅色处理——宁可界面偏亮，也不要把浅色用户误判成深色。
        /// </summary>
        internal static OfficeThemeKind FromRegistryValue(int value)
        {
            switch (value)
            {
                case 3:
                    return OfficeThemeKind.White;
                case 4:
                    return OfficeThemeKind.DarkGray;
                default:
                    return value >= 5 ? OfficeThemeKind.Black : OfficeThemeKind.Colorful;
            }
        }

        /// <summary>测试/预览用：直接指定主题，跳过注册表。</summary>
        internal static void Override(OfficeThemeKind kind, string source)
        {
            _detected = true;
            _kind = kind;
            _source = source ?? "手动指定";
        }

        /// <summary>恢复为“按注册表自动检测”。</summary>
        internal static void ResetToDetected()
        {
            _detected = false;
            _kind = OfficeThemeKind.Colorful;
            _source = "Office 默认（浅色）";
            DetectOnce();
        }

        private static string HostProcessName()
        {
            return Process.GetCurrentProcess().ProcessName;
        }

        private static void DetectOnce()
        {
            if (_detected)
            {
                return;
            }
            _detected = true;
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(CommonKeyPath))
                {
                    if (key == null)
                    {
                        _source = "未找到 Office 注册表项，按浅色处理";
                        return;
                    }
                    var raw = key.GetValue(ThemeValueName);
                    if (raw == null)
                    {
                        _source = "Office 未显式设置主题，按默认浅色处理";
                        return;
                    }
                    var value = Convert.ToInt32(raw);
                    _kind = FromRegistryValue(value);
                    _source = "Office 主题（注册表 UI Theme=" + value + "）";
                }
            }
            catch (Exception ex)
            {
                _source = "读取 Office 主题失败，按浅色处理：" + ex.Message;
            }
        }
    }
}
