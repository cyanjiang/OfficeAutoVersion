using System;
using System.IO;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAutoVersion.Excel
{
    /// <summary>
    /// Excel 侧的“比较 / 恢复”实现。
    /// Excel 没有内置的“比较工作簿”命令（旧版的 Spreadsheet Compare 独立工具在 Office 2024 已移除），
    /// 因此比较使用 Excel 自带的“并排查看”（视图 → 并排查看），把版本文件与当前工作簿同步滚动对比；
    /// 恢复用文件级替换（先保存并快照当前状态，再关闭工作簿、覆盖文件、重新打开）。
    /// </summary>
    internal sealed class ExcelVersionActions : IVersionHostActions
    {
        // 注意：这里必须保存“取 Application 的委托”，不能保存 Application 本身。
        // VSTO 可能在加载项完成初始化之前就调用 CreateRibbonExtensibilityObject()，
        // 那一刻 ThisAddIn.Application 还是 null；把 null 存下来，之后每次比较/恢复
        // 都会以 NullReferenceException 失败（用户看到的就是“比较失败，未将对象引用设置到对象的实例”）。
        private readonly Func<ExcelInterop.Application> _getApp;
        private readonly Func<VersionManager> _getManager;
        private readonly Action<bool> _setSuppressAutoVersion;

        public ExcelVersionActions(Func<ExcelInterop.Application> getApp, Func<VersionManager> getManager,
            Action<bool> setSuppressAutoVersion)
        {
            _getApp = getApp;
            _getManager = getManager;
            _setSuppressAutoVersion = setSuppressAutoVersion;
        }

        /// <summary>调用时刻的 Excel Application，永远取最新的宿主对象。</summary>
        private ExcelInterop.Application GetApp()
        {
            var app = _getApp != null ? _getApp() : null;
            if (app == null)
            {
                throw new InvalidOperationException(
                    "Excel 还没准备好（加载项尚未完成初始化）。请稍等一秒后重试；若持续出现，请重启 Excel。");
            }
            return app;
        }

        public string CompareButtonText
        {
            get { return "并排查看（Excel 自带）"; }
        }

        public string CompareHint
        {
            get
            {
                return "用 Excel 自带的“并排查看”把该版本以只读方式打开，与当前工作簿并排同步滚动对比。";
            }
        }

        public void CompareWithCurrent(string currentFilePath, string versionFilePath)
        {
            var app = GetApp(); // 先确认宿主可用，再动手打开任何文件
            var current = ResolveWorkbook(currentFilePath);
            if (current == null)
            {
                throw new InvalidOperationException("找不到要比较的工作簿，可能已被移动或删除：" +
                                                    (currentFilePath ?? "(未知路径)"));
            }

            // 以只读方式打开版本文件（不加入最近使用列表）
            var version = OpenReadOnly(versionFilePath, false);

            try
            {
                // 激活要在“两个窗口都存在”之后再试；模态窗口期间激活可能被系统拒绝，失败不影响比较
                ActivateQuietly(current);
                // Excel 自带“视图 → 并排查看”（CompareSideBySideWith 在 Windows 集合上）
                app.Windows.CompareSideBySideWith(version.Name);
                app.Windows.Arrange(ExcelInterop.XlArrangeStyle.xlArrangeStyleVertical,
                    Type.Missing, Type.Missing, Type.Missing);
                try
                {
                    // “同步滚动”是用户的持久视图设置，可能被关掉过；比较时显式打开，两表才会一起滚
                    app.Windows.SyncScrollingSideBySide = true;
                }
                catch
                {
                    // 个别版本不支持该设置，忽略
                }
                ActivateQuietly(version);
            }
            catch
            {
                // 比较失败时版本文件至少已经打开，用户可以手工并排；把异常继续抛给界面提示
                throw;
            }
        }

        public string RestoreVersion(string currentFilePath, string versionFilePath)
        {
            var path = currentFilePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                throw new InvalidOperationException("请先把工作簿保存到磁盘，再恢复版本。");
            }
            if (!File.Exists(versionFilePath))
            {
                throw new InvalidOperationException("该版本文件已被移动或删除：" + versionFilePath);
            }

            var app = GetApp(); // 先确认宿主可用
            var workbook = FindOpenWorkbook(path);

            // 1) 保存当前内容（含未保存修改），期间临时关掉自动快照，只留一份安全副本
            if (workbook != null && !workbook.Saved)
            {
                if (_setSuppressAutoVersion != null)
                {
                    _setSuppressAutoVersion(true);
                }
                try
                {
                    workbook.Save();
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

            // 2) 关闭工作簿（不保存）→ 覆盖文件 → 重新打开
            if (workbook != null)
            {
                workbook.Close(false);
            }
            // Excel 关闭后可能还会短暂占用文件，先等句柄释放再覆盖
            VersionManager.WaitForFileUnlocked(path, 3000);
            try
            {
                VersionManager.ReplaceFileRobust(versionFilePath, path);
            }
            catch
            {
                // 覆盖失败时把原文件重新打开，避免用户“工作簿被关掉却没恢复成”
                try
                {
                    if (File.Exists(path))
                    {
                        ActivateQuietly(app.Workbooks.Open(path));
                    }
                }
                catch
                {
                    // 重新打开失败就保持原样，异常照旧抛给界面
                }
                throw;
            }

            var reopened = app.Workbooks.Open(path);
            ActivateQuietly(reopened);
            return snapshot;
        }

        /// <summary>
        /// Excel 没有“页”的概念（差异摘要用“工作表!单元格地址”定位），因此不提供页码。
        /// </summary>
        public System.Collections.Generic.Dictionary<string, int> GetTextPageMap(string documentFilePath)
        {
            return null;
        }

        /// <summary>
        /// 按路径定位已打开的工作簿；没打开则自动打开（文件存在时）。
        /// 不能依赖 Application.ActiveWorkbook：版本历史窗口是模态窗口，
        /// 打开期间 ActiveWorkbook 会返回 null，正是“当前没有打开的工作簿”提示的来源。
        /// </summary>
        private ExcelInterop.Workbook ResolveWorkbook(string path)
        {
            var opened = FindOpenWorkbook(path);
            if (opened != null)
            {
                return opened;
            }
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                return GetApp().Workbooks.Open(path);
            }
            return null;
        }

        private ExcelInterop.Workbook FindOpenWorkbook(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }
            var app = GetApp();
            var count = 0;
            try
            {
                count = app.Workbooks.Count;
            }
            catch
            {
                return null;
            }

            for (var i = 1; i <= count; i++)
            {
                ExcelInterop.Workbook workbook = null;
                string fullName = null;
                try
                {
                    workbook = app.Workbooks[i];
                    fullName = workbook.FullName;
                }
                catch
                {
                    continue;
                }
                if (string.Equals(fullName, path, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }
            return null;
        }

        private ExcelInterop.Workbook OpenReadOnly(string path, bool addToRecent)
        {
            // Open(Filename, UpdateLinks, ReadOnly, Format, Password, ...)
            return GetApp().Workbooks.Open(path, Type.Missing, true, Type.Missing,
                Type.Missing, Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, addToRecent);
        }

        private static void ActivateQuietly(ExcelInterop.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }
            try
            {
                workbook.Activate();
            }
            catch
            {
                // 模态窗口 / 工作簿被隐藏时可能无法激活，不影响比较结果
            }
        }
    }
}
