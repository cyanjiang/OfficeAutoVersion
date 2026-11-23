using System;
using System.IO;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;
using WordInterop = Microsoft.Office.Interop.Word;

namespace OfficeAutoVersion.Word
{
    /// <summary>
    /// Word 侧的“比较 / 恢复”实现：
    /// 比较调用 Word 自带的“审阅 → 比较”（Application.CompareDocuments，输出带修订标记的新文档）；
    /// 恢复用文件级替换（先保存并快照当前状态，再关闭文档、覆盖文件、重新打开）。
    /// </summary>
    internal sealed class WordVersionActions : IVersionHostActions
    {
        // 注意：保存“取 Application 的委托”，不要保存 Application 本身。
        // VSTO 可能在加载项完成初始化之前就调用 CreateRibbonExtensibilityObject()，
        // 那一刻 ThisAddIn.Application 还是 null，存下来就会让每次比较/恢复都以
        // NullReferenceException 失败。
        private readonly Func<WordInterop.Application> _getApp;
        private readonly Func<VersionManager> _getManager;
        private readonly Action<bool> _setSuppressAutoVersion;

        public WordVersionActions(Func<WordInterop.Application> getApp, Func<VersionManager> getManager,
            Action<bool> setSuppressAutoVersion)
        {
            _getApp = getApp;
            _getManager = getManager;
            _setSuppressAutoVersion = setSuppressAutoVersion;
        }

        /// <summary>调用时刻的 Word Application，永远取最新的宿主对象。</summary>
        private WordInterop.Application GetApp()
        {
            var app = _getApp != null ? _getApp() : null;
            if (app == null)
            {
                throw new InvalidOperationException(
                    "Word 还没准备好（加载项尚未完成初始化）。请稍等一秒后重试；若持续出现，请重启 Word。");
            }
            return app;
        }

        public string CompareButtonText
        {
            get { return "比较差异（Word 比较）"; }
        }

        public string CompareHint
        {
            get { return "调用 Word 自带的“比较文档”，生成一份带修订标记的对比文档；原文档不会改动。"; }
        }

        public void CompareWithCurrent(string currentFilePath, string versionFilePath)
        {
            // 注意：不能再用 Documents.Open 打开“当前文档”——Word 对已打开的文档会返回同一个
            // Document 对象，之后关闭临时文档就会把用户正在编辑的文档一起关掉。
            // 因此直接把用户已打开的文档当作“修订版”，只临时打开版本文件。
            var app = GetApp(); // 先确认宿主可用，再动手打开任何文件
            var current = ResolveDocument(currentFilePath);
            if (current == null)
            {
                throw new InvalidOperationException("找不到要比较的文档，可能已被移动或删除：" +
                                                    (currentFilePath ?? "(未知路径)"));
            }

            WordInterop.Document original = null;
            try
            {
                // Original = 旧版本，Revised = 用户当前的文档
                original = OpenDocument(versionFilePath, true);

                var result = app.CompareDocuments(
                    original,
                    current,
                    WordInterop.WdCompareDestination.wdCompareDestinationNew,
                    WordInterop.WdGranularity.wdGranularityWordLevel,
                    true,   // 比较格式
                    true,   // 大小写变化
                    true,   // 空白字符
                    true,   // 表格
                    true,   // 页眉页脚
                    true,   // 脚注尾注
                    true,   // 文本框
                    true,   // 域
                    true,   // 批注
                    true,   // 移动
                    "自动版本",
                    false); // 忽略比较警告

                if (result != null)
                {
                    ActivateQuietly(result);
                }
            }
            finally
            {
                CloseQuietly(original);
            }
        }

        public string RestoreVersion(string currentFilePath, string versionFilePath)
        {
            var path = currentFilePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                throw new InvalidOperationException("请先把文档保存到磁盘，再恢复版本。");
            }
            if (!File.Exists(versionFilePath))
            {
                throw new InvalidOperationException("该版本文件已被移动或删除：" + versionFilePath);
            }

            GetApp(); // 先确认宿主可用（未就绪时给出明确提示，而不是空引用）
            var document = FindOpenDocument(path);

            // 1) 先保存（把用户看得见的修改落到磁盘），保存期间临时关掉自动快照，
            //    这样下面只产生“一份”安全副本——正好是当前看到的内容。
            if (document != null && !document.Saved)
            {
                if (_setSuppressAutoVersion != null)
                {
                    _setSuppressAutoVersion(true);
                }
                try
                {
                    document.Save();
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
                // 内容与最新版本一致（去重跳过）时，那份最新版本就是安全副本
                var versions = manager.ListVersions(path);
                if (versions.Count > 0)
                {
                    snapshot = versions[0];
                }
            }

            // 2) 关闭文档（不保存）→ 用版本文件覆盖 → 重新打开
            if (document != null)
            {
                var missing = Type.Missing;
                object doNotSave = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                document.Close(ref doNotSave, ref missing, ref missing);
            }

            // Word 关闭文档后还会短暂占用文件（含 ~$ 所有者文件），先等句柄释放再覆盖
            VersionManager.WaitForFileUnlocked(path, 3000);
            try
            {
                VersionManager.ReplaceFileRobust(versionFilePath, path);
            }
            catch
            {
                // 覆盖失败（文件仍被占用）：把原文档重新打开，避免用户“文档被关掉却没恢复成”的双重损失
                try
                {
                    if (File.Exists(path))
                    {
                        ActivateQuietly(OpenDocument(path, false));
                    }
                }
                catch
                {
                    // 重新打开也失败就保持原样，异常照旧抛给界面
                }
                throw;
            }

            ActivateQuietly(OpenDocument(path, false));
            return snapshot;
        }

        /// <summary>
        /// 段落文本 → 页码。页码只能由排版结果得出（docx 里不存页码），所以必须问 Word 本身。
        /// 只读遍历，不改动文档；任何异常都退化为“没有页码”，不影响差异摘要本身。
        /// </summary>
        public System.Collections.Generic.Dictionary<string, int> GetTextPageMap(string documentFilePath)
        {
            try
            {
                GetApp(); // 宿主可用性检查（未就绪时下面会整体退化为无页码）
                var document = FindOpenDocument(documentFilePath);
                if (document == null)
                {
                    return null;
                }

                var map = new System.Collections.Generic.Dictionary<string, int>(StringComparer.Ordinal);
                var paragraphs = document.Paragraphs;
                var count = paragraphs.Count;
                // 超大文档（几千段）逐段取页码较慢，限制一个上限，超出部分不标页码
                const int maxParagraphs = 6000;
                if (count > maxParagraphs)
                {
                    count = maxParagraphs;
                }

                for (var i = 1; i <= count; i++)
                {
                    try
                    {
                        var range = paragraphs[i].Range;
                        if (range == null)
                        {
                            continue;
                        }
                        var text = range.Text;
                        if (string.IsNullOrEmpty(text))
                        {
                            continue;
                        }
                        var key = ContentDiff.NormalizeKey(text);
                        if (key.Length == 0 || map.ContainsKey(key))
                        {
                            continue; // 重复段落取首次出现的页码
                        }
                        var pageValue = range.Information[WordInterop.WdInformation.wdActiveEndPageNumber];
                        if (pageValue == null)
                        {
                            continue;
                        }
                        var pageText = pageValue.ToString();
                        int page = 0;
                        if (!int.TryParse(pageText, out page) || page <= 0)
                        {
                            continue;
                        }
                        map[key] = page;
                    }
                    catch
                    {
                        // 个别段落（含控件、域等）可能访问失败，跳过即可
                    }
                }
                return map;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 按路径定位已打开的文档；没打开则自动打开（文件存在时）。
        /// 不能依赖 Application.ActiveDocument：版本历史窗口是模态窗口，
        /// 打开期间 ActiveDocument 可能返回 null。
        /// </summary>
        private WordInterop.Document ResolveDocument(string path)
        {
            var opened = FindOpenDocument(path);
            if (opened != null)
            {
                return opened;
            }
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                return OpenDocument(path, false);
            }
            return null;
        }

        private WordInterop.Document FindOpenDocument(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }
            var app = GetApp();
            var count = 0;
            try
            {
                count = app.Documents.Count;
            }
            catch
            {
                return null;
            }

            // 注意：Word 的 Documents 集合里可能残留无法访问的“幽灵”项，逐个 try/catch 跳过
            for (var i = 1; i <= count; i++)
            {
                WordInterop.Document document = null;
                string fullName = null;
                try
                {
                    document = app.Documents[i];
                    fullName = document.FullName;
                }
                catch
                {
                    continue;
                }
                if (string.Equals(fullName, path, StringComparison.OrdinalIgnoreCase))
                {
                    return document;
                }
            }
            return null;
        }

        private WordInterop.Document OpenDocument(string path, bool readOnly)
        {
            object fileName = path;
            object missing = Type.Missing;
            object readOnlyFlag = readOnly;
            object addToRecent = false;
            return GetApp().Documents.Open(ref fileName, ref missing, ref readOnlyFlag, ref addToRecent,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);
        }

        private static void ActivateQuietly(WordInterop.Document document)
        {
            if (document == null)
            {
                return;
            }
            try
            {
                document.Activate();
            }
            catch
            {
                // 模态窗口期间激活可能失败，不影响操作本身
            }
        }

        private static void CloseQuietly(WordInterop.Document document)
        {
            if (document == null)
            {
                return;
            }
            try
            {
                object doNotSave = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                object missing = Type.Missing;
                document.Close(ref doNotSave, ref missing, ref missing);
            }
            catch
            {
                // 关闭临时文档失败可忽略
            }
        }
    }
}
