using System;
using System.IO;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;

namespace OfficeAutoVersion.PowerPoint
{
    /// <summary>
    /// PowerPoint 侧的“比较 / 恢复”实现。
    /// 比较调用 PowerPoint 自带的“审阅 → 比较”（Presentation.Merge），
    /// 差异会以修订形式出现在“审阅”窗格里，可逐条接受/拒绝；
    /// 恢复用文件级替换（先保存并快照当前状态，再关闭演示文稿、覆盖文件、重新打开）。
    /// </summary>
    internal sealed class PowerPointVersionActions : IVersionHostActions
    {
        // 注意：保存“取 Application 的委托”，不要保存 Application 本身。
        // VSTO 可能在加载项完成初始化之前就调用 CreateRibbonExtensibilityObject()，
        // 那一刻 Application 还是 null，存下来会让比较/恢复全部以 NullReferenceException 失败。
        private readonly Func<PowerPointInterop.Application> _getApp;
        private readonly Func<VersionManager> _getManager;
        private readonly Action<bool> _setSuppressAutoVersion;

        public PowerPointVersionActions(Func<PowerPointInterop.Application> getApp, Func<VersionManager> getManager,
            Action<bool> setSuppressAutoVersion)
        {
            _getApp = getApp;
            _getManager = getManager;
            _setSuppressAutoVersion = setSuppressAutoVersion;
        }

        /// <summary>调用时刻的 PowerPoint Application，永远取最新的宿主对象。</summary>
        private PowerPointInterop.Application GetApp()
        {
            var app = _getApp != null ? _getApp() : null;
            if (app == null)
            {
                throw new InvalidOperationException(
                    "PowerPoint 还没准备好（加载项尚未完成初始化）。请稍等一秒后重试；若持续出现，请重启 PowerPoint。");
            }
            return app;
        }

        public string CompareButtonText
        {
            get { return "比较差异（PPT 比较）"; }
        }

        public string CompareHint
        {
            get
            {
                return "调用 PowerPoint 自带的“比较”（审阅 → 比较），差异会出现在“审阅”窗格，可逐条接受或拒绝。";
            }
        }

        public void CompareWithCurrent(string currentFilePath, string versionFilePath)
        {
            GetApp(); // 先确认宿主可用（未初始化时给出明确提示，而不是空引用）
            var presentation = ResolvePresentation(currentFilePath);
            if (presentation == null)
            {
                throw new InvalidOperationException("找不到要比较的演示文稿，可能已被移动或删除：" +
                                                    (currentFilePath ?? "(未知路径)"));
            }

            ActivateQuietly(presentation);
            // Presentation.Merge = PowerPoint 自带“审阅 → 比较”
            presentation.Merge(versionFilePath);
        }

        public string RestoreVersion(string currentFilePath, string versionFilePath)
        {
            var path = currentFilePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                throw new InvalidOperationException("请先把演示文稿保存到磁盘，再恢复版本。");
            }
            if (!File.Exists(versionFilePath))
            {
                throw new InvalidOperationException("该版本文件已被移动或删除：" + versionFilePath);
            }

            GetApp(); // 先确认宿主可用
            var presentation = FindOpenPresentation(path);

            // 1) 保存当前内容（含未保存修改），期间临时关掉自动快照，只留一份安全副本
            if (presentation != null &&
                presentation.Saved == Microsoft.Office.Core.MsoTriState.msoFalse)
            {
                if (_setSuppressAutoVersion != null)
                {
                    _setSuppressAutoVersion(true);
                }
                try
                {
                    presentation.Save();
                }
                finally
                {
                    if (_setSuppressAutoVersion != null)
                    {
                        _setSuppressAutoVersion(false);
                    }
                }
            }

            var manager = _getManager != null ? _getManager() : null;
            var snapshot = manager != null ? manager.CreateVersion(path) : null;
            if (snapshot == null && manager != null)
            {
                var versions = manager.ListVersions(path);
                if (versions.Count > 0)
                {
                    snapshot = versions[0];
                }
            }

            // 2) 关闭演示文稿 → 覆盖文件 → 重新打开
            if (presentation != null)
            {
                presentation.Close();
            }
            // PowerPoint 关闭后可能还会短暂占用文件，先等句柄释放再覆盖
            VersionManager.WaitForFileUnlocked(path, 3000);
            try
            {
                VersionManager.ReplaceFileRobust(versionFilePath, path);
            }
            catch
            {
                // 覆盖失败时把原文件重新打开，避免用户“文稿被关掉却没恢复成”
                try
                {
                    if (File.Exists(path))
                    {
                        ActivateQuietly(OpenPresentation(path));
                    }
                }
                catch
                {
                    // 重新打开失败就保持原样，异常照旧抛给界面
                }
                throw;
            }
            ActivateQuietly(OpenPresentation(path));
            return snapshot;
        }

        /// <summary>
        /// PowerPoint 用“幻灯片N · 形状名”定位，幻灯片序号已含在位置里，因此不再提供页码。
        /// </summary>
        public System.Collections.Generic.Dictionary<string, int> GetTextPageMap(string documentFilePath)
        {
            return null;
        }

        /// <summary>
        /// 按路径定位已打开的演示文稿；没打开则自动打开（文件存在时）。
        /// 不能依赖 Application.ActivePresentation：版本历史窗口是模态窗口，
        /// 打开期间 ActivePresentation 可能返回 null。
        /// </summary>
        private PowerPointInterop.Presentation ResolvePresentation(string path)
        {
            var opened = FindOpenPresentation(path);
            if (opened != null)
            {
                return opened;
            }
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                return OpenPresentation(path);
            }
            return null;
        }

        private PowerPointInterop.Presentation FindOpenPresentation(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }

            var app = GetApp();
            var count = 0;
            try
            {
                count = app.Presentations.Count;
            }
            catch
            {
                return null;
            }

            for (var i = 1; i <= count; i++)
            {
                PowerPointInterop.Presentation presentation = null;
                string fullName = null;
                try
                {
                    presentation = app.Presentations[i];
                    fullName = presentation.FullName;
                }
                catch
                {
                    continue;
                }
                if (string.Equals(fullName, path, StringComparison.OrdinalIgnoreCase))
                {
                    return presentation;
                }
            }
            return null;
        }

        private PowerPointInterop.Presentation OpenPresentation(string path)
        {
            return GetApp().Presentations.Open(path,
                Microsoft.Office.Core.MsoTriState.msoFalse,   // ReadOnly
                Microsoft.Office.Core.MsoTriState.msoFalse,   // Untitled
                Microsoft.Office.Core.MsoTriState.msoTrue);   // WithWindow
        }

        private static void ActivateQuietly(PowerPointInterop.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }
            try
            {
                presentation.Windows[1].Activate();
            }
            catch
            {
                // 模态窗口 / 无窗口打开时可能失败，忽略
            }
        }
    }
}
