using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using OfficeAutoVersion.Core;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// 三个 Office 加载项（Word/Excel/PowerPoint）共享的功能区实现（Ribbon XML）。
    /// 由各加载项的 ThisAddIn.CreateRibbonExtensibilityObject() 返回本对象。
    /// </summary>
    [ComVisible(true)]
    public sealed class VersioningRibbon : Office.IRibbonExtensibility
    {
        private readonly string _appKey;              // "Word" / "Excel" / "PowerPoint"
        private readonly Func<string> _getActivePath; // 获取当前活动文档的完整路径
        private readonly Action _optionsChanged;      // 通知加载项重新加载配置
        private readonly IVersionHostActions _hostActions; // 比较/恢复（各应用自带能力不同）
        private Office.IRibbonUI _ribbon;

        public VersioningRibbon(string appKey, Func<string> getActivePath, Action optionsChanged = null,
            IVersionHostActions hostActions = null)
        {
            _appKey = appKey;
            _getActivePath = getActivePath;
            _optionsChanged = optionsChanged;
            _hostActions = hostActions;
        }

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("OfficeAutoVersion.Common.Ribbon.xml");
        }

        #endregion

        /// <summary>功能区加载完成，保存 IRibbonUI 以便后续刷新控件状态。</summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        /// <summary>复选框“自动建立版本”的当前勾选状态。</summary>
        public bool GetAutoVersionChecked(Office.IRibbonControl control)
        {
            return OptionsStore.Load(_appKey).Enabled;
        }

        /// <summary>用户勾选 / 取消“自动建立版本”复选框。</summary>
        public void OnAutoVersionChecked(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                var options = OptionsStore.Load(_appKey);
                options.Enabled = pressed;
                OptionsStore.Save(_appKey, options);
                if (_optionsChanged != null)
                {
                    _optionsChanged();
                }
                if (_ribbon != null)
                {
                    _ribbon.Invalidate();
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        /// <summary>打开“版本历史”窗口：列出所有版本，可比较（Office 自带功能）与恢复。</summary>
        public void OnVersionHistory(Office.IRibbonControl control)
        {
            try
            {
                var path = _getActivePath != null ? _getActivePath() : null;
                if (string.IsNullOrEmpty(path))
                {
                    MessageBox.Show("当前没有已保存到磁盘的文档。\n请先保存文档，再查看版本历史。", "自动版本",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var manager = new VersionManager(OptionsStore.Load(_appKey));
                VersionHistoryAction pendingAction;
                string pendingVersion;
                using (var form = new VersionHistoryForm(_appKey, path, manager, _hostActions))
                {
                    // 窗口只负责“选择”，不在模态期间操作 Office
                    form.IsModal = true;
                    form.ShowDialog();
                    pendingAction = form.PendingAction;
                    pendingVersion = form.PendingVersionPath;
                }

                if (pendingAction == VersionHistoryAction.None || string.IsNullOrEmpty(pendingVersion))
                {
                    return;
                }
                if (_hostActions == null)
                {
                    MessageBox.Show("当前加载项没有提供比较 / 恢复能力。", "自动版本",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 窗口已关闭：此时调用 Office 对象模型才可靠
                // （模态窗口显示期间 Excel 的 ActiveWorkbook 会变成 null）
                if (pendingAction == VersionHistoryAction.Compare)
                {
                    try
                    {
                        _hostActions.CompareWithCurrent(path, pendingVersion);
                    }
                    catch (Exception ex)
                    {
                        ShowError("比较失败", ex);
                    }
                }
                else if (pendingAction == VersionHistoryAction.Restore)
                {
                    try
                    {
                        var snapshot = _hostActions.RestoreVersion(path, pendingVersion);
                        VersionHistoryForm.ShowRestoreResult(null, pendingVersion, snapshot);
                    }
                    catch (Exception ex)
                    {
                        ShowError("恢复失败", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        /// <summary>打开当前文档的版本文件夹。</summary>
        public void OnOpenVersionFolder(Office.IRibbonControl control)
        {
            try
            {
                var path = _getActivePath != null ? _getActivePath() : null;
                if (string.IsNullOrEmpty(path))
                {
                    MessageBox.Show("当前没有已保存到磁盘的文档。", "自动版本",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                var options = OptionsStore.Load(_appKey);
                var directory = VersionPaths.GetVersionsDirectory(path, options);
                if (Directory.Exists(directory))
                {
                    Process.Start("explorer.exe", "\"" + directory + "\"");
                }
                else
                {
                    MessageBox.Show("该文档还没有生成任何版本。\n保存文档后，版本将出现在：\n" + directory,
                        "自动版本", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        /// <summary>打开设置对话框。</summary>
        public void OnSettings(Office.IRibbonControl control)
        {
            try
            {
                var currentPath = _getActivePath != null ? _getActivePath() : null;
                using (var form = new SettingsForm(_appKey, currentPath))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        if (_optionsChanged != null)
                        {
                            _optionsChanged();
                        }
                        if (_ribbon != null)
                        {
                            _ribbon.InvalidateControl("chkAutoVersionEnabled");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowError(ex);
            }
        }

        private static void ShowError(Exception ex)
        {
            ShowError("操作失败", ex);
        }

        private static void ShowError(string title, Exception ex)
        {
            try
            {
                Logger.Default.Error(title, ex);
            }
            catch
            {
                // 忽略
            }
            MessageBox.Show(title + "：" + ex.Message, "自动版本",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    return null;
                }
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
