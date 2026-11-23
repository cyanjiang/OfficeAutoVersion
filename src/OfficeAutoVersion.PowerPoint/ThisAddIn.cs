using System;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;
// 注意：别名不能叫 Word/Excel/PowerPoint——会被项目自身的命名空间
// OfficeAutoVersion.PowerPoint 等遮蔽，导致 PowerPoint.Presentation 解析失败。
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace OfficeAutoVersion.PowerPoint
{
    /// <summary>
    /// PowerPoint 自动版本加载项：在 PresentationBeforeSave 事件中把演示文稿的当前磁盘状态
    /// 复制为一份带时间戳的版本快照（保存的是“本次保存之前”的内容）。
    /// </summary>
    public partial class ThisAddIn
    {
        private const string AppKey = "PowerPoint";

        private VersionManager _versionManager;
        private VersioningRibbon _ribbon;
        private bool _suppressAutoVersion;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _versionManager = new VersionManager(OptionsStore.Load(AppKey));
            Application.PresentationBeforeSave += Application_PresentationBeforeSave;
            TryLog("PowerPoint 加载项已启动，已挂接 PresentationBeforeSave");
        }

        /// <summary>启动自检日志：加载项没生效时，能一眼看出是“没启动”还是“事件没触发”。</summary>
        private static void TryLog(string message)
        {
            try
            {
                Logger.Default.Debug(message);
            }
            catch
            {
                // 日志失败不影响加载项
            }
        }

        // 注意：EApplication_PresentationBeforeSaveEventHandler 的签名为
        // (Presentation Pres, ref bool Cancel)。
        private void Application_PresentationBeforeSave(PowerPointInterop.Presentation presentation, ref bool cancel)
        {
            cancel = false; // 本插件绝不拦截用户的保存操作
            if (_suppressAutoVersion)
            {
                return;
            }
            TryCreateVersion(presentation);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                Application.PresentationBeforeSave -= Application_PresentationBeforeSave;
            }
            catch
            {
                // 卸载时忽略
            }
        }

        private void TryCreateVersion(PowerPointInterop.Presentation presentation)
        {
            try
            {
                if (_versionManager != null)
                {
                    _versionManager.CreateVersion(GetPresentationPath(presentation));
                }
            }
            catch
            {
                // 自动版本失败绝不能影响用户的保存操作
            }
        }

        private static string GetPresentationPath(PowerPointInterop.Presentation presentation)
        {
            try
            {
                return presentation == null ? null : presentation.FullName;
            }
            catch
            {
                // 从未保存过的演示文稿取路径会抛异常
                return null;
            }
        }

        private string GetActivePresentationPath()
        {
            try
            {
                return GetPresentationPath(Application.ActivePresentation);
            }
            catch
            {
                return null;
            }
        }

        private void ReloadOptions()
        {
            _versionManager = new VersionManager(OptionsStore.Load(AppKey));
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            // 注意：这里传的是“取 Application 的委托”，不是 Application 本身。
            // VSTO 可能在加载项初始化完成之前就调用本方法，此刻 Application 仍为 null；
            // 若把 null 存进动作类，之后每次比较/恢复都会抛 NullReferenceException。
            TryLog("PowerPoint 功能区已创建，Application=" + (Application == null ? "null" : "ok"));
            var actions = new PowerPointVersionActions(delegate { return Application; }, delegate { return _versionManager; },
                SetSuppressAutoVersion);
            _ribbon = new VersioningRibbon(AppKey, GetActivePresentationPath, ReloadOptions, actions);
            return _ribbon;
        }

        /// <summary>恢复版本时临时关闭自动快照，避免一次恢复产生两份重复版本。</summary>
        internal void SetSuppressAutoVersion(bool value)
        {
            _suppressAutoVersion = value;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
