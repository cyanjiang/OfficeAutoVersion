using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// 深色主题下，让 Windows 自己绘制的那部分也跟着变深：
    ///   • 标题栏：DWM 的“沉浸式深色”（Win10 1809+）
    ///   • 滚动条 / 列表头：DarkMode_Explorer 视觉样式
    /// 全部是尽力而为的调用——系统不支持或返回失败都保持原样，绝不影响窗口本身。
    /// </summary>
    internal static class NativeTheme
    {
        private const int DwmUseImmersiveDarkMode = 20;
        private const int DwmUseImmersiveDarkModeLegacy = 19;

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        private static extern int SetWindowTheme(IntPtr hwnd, string subAppName, string subIdList);

        /// <summary>给窗体套深色标题栏（仅深色主题；窗体句柄还没建好时等句柄创建后再套）。</summary>
        public static void Apply(Form form)
        {
            if (form == null || !UiTheme.IsDark)
            {
                return;
            }
            try
            {
                if (form.IsHandleCreated)
                {
                    ApplyHandle(form.Handle);
                }
                else
                {
                    form.HandleCreated += OnFormHandleCreated;
                }
            }
            catch
            {
                // 外观优化，失败忽略
            }
        }

        /// <summary>给列表、文本框等控件套深色滚动条。</summary>
        public static void ApplyToControl(Control control)
        {
            if (control == null || !UiTheme.IsDark)
            {
                return;
            }
            try
            {
                if (control.IsHandleCreated)
                {
                    ApplyHandle(control.Handle);
                }
                else
                {
                    control.HandleCreated += OnControlHandleCreated;
                }
            }
            catch
            {
                // 外观优化，失败忽略
            }
        }

        private static void OnFormHandleCreated(object sender, EventArgs e)
        {
            var form = sender as Form;
            if (form == null)
            {
                return;
            }
            try
            {
                form.HandleCreated -= OnFormHandleCreated;
                ApplyHandle(form.Handle);
            }
            catch
            {
                // 忽略
            }
        }

        private static void OnControlHandleCreated(object sender, EventArgs e)
        {
            var control = sender as Control;
            if (control == null)
            {
                return;
            }
            try
            {
                control.HandleCreated -= OnControlHandleCreated;
                ApplyHandle(control.Handle);
            }
            catch
            {
                // 忽略
            }
        }

        private static void ApplyHandle(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
            {
                return;
            }
            try
            {
                var enabled = 1;
                if (DwmSetWindowAttribute(hwnd, DwmUseImmersiveDarkMode, ref enabled, sizeof(int)) != 0)
                {
                    // 老版本 Windows 用的是属性 19
                    DwmSetWindowAttribute(hwnd, DwmUseImmersiveDarkModeLegacy, ref enabled, sizeof(int));
                }
            }
            catch
            {
                // 忽略
            }
            try
            {
                SetWindowTheme(hwnd, "DarkMode_Explorer", null);
            }
            catch
            {
                // 忽略
            }
        }
    }
}
