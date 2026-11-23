using System;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;
// 注意：别名不能叫 Word/Excel/PowerPoint——会被项目自身的命名空间
// OfficeAutoVersion.Excel 等遮蔽，导致 Excel.Workbook 解析失败。
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAutoVersion.Excel
{
    /// <summary>
    /// Excel 自动版本加载项：在 WorkbookBeforeSave 事件中把工作簿的当前磁盘状态
    /// 复制为一份带时间戳的版本快照（保存的是“本次保存之前”的内容）。
    /// </summary>
    public partial class ThisAddIn
    {
        private const string AppKey = "Excel";

        private VersionManager _versionManager;
        private VersioningRibbon _ribbon;
        private bool _suppressAutoVersion;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _versionManager = new VersionManager(OptionsStore.Load(AppKey));
            Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
            TryLog("Excel 加载项已启动，已挂接 WorkbookBeforeSave");
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

        // 注意：AppEvents_WorkbookBeforeSaveEventHandler 的签名为
        // (Workbook Wb, bool SaveAsUI, ref bool Cancel)。
        private void Application_WorkbookBeforeSave(ExcelInterop.Workbook workbook, bool saveAsUi, ref bool cancel)
        {
            cancel = false; // 本插件绝不拦截用户的保存操作
            if (_suppressAutoVersion)
            {
                return;
            }
            TryCreateVersion(workbook);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                Application.WorkbookBeforeSave -= Application_WorkbookBeforeSave;
            }
            catch
            {
                // 卸载时忽略
            }
        }

        private void TryCreateVersion(ExcelInterop.Workbook workbook)
        {
            try
            {
                if (_versionManager != null)
                {
                    _versionManager.CreateVersion(GetWorkbookPath(workbook));
                }
            }
            catch
            {
                // 自动版本失败绝不能影响用户的保存操作
            }
        }

        private static string GetWorkbookPath(ExcelInterop.Workbook workbook)
        {
            try
            {
                return workbook == null ? null : workbook.FullName;
            }
            catch
            {
                return null;
            }
        }

        private string GetActiveWorkbookPath()
        {
            try
            {
                return GetWorkbookPath(Application.ActiveWorkbook);
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
            TryLog("Excel 功能区已创建，Application=" + (Application == null ? "null" : "ok"));
            var actions = new ExcelVersionActions(delegate { return Application; }, delegate { return _versionManager; },
                SetSuppressAutoVersion);
            _ribbon = new VersioningRibbon(AppKey, GetActiveWorkbookPath, ReloadOptions, actions);
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
