using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeAutoVersion.Core
{
    /// <summary>差异摘要的统一入口：读两个文件 → 提取内容 → 选择合适的比较方式。</summary>
    public static class VersionDiffer
    {
        /// <summary>
        /// 比较“当前文档”与“版本文件”，方向为：当前文档是新状态、版本是旧状态，
        /// 摘要回答“相比这个版本，现在多了/少了/改了什么”。
        /// 失败时抛 InvalidOperationException，消息可直接展示给用户。
        /// </summary>
        public static DiffSummary Diff(string currentPath, string versionPath)
        {
            List<ContentItem> current;
            string currentError;
            if (!OpcContent.TryExtract(currentPath, out current, out currentError))
            {
                throw new InvalidOperationException("无法读取当前文件内容。" + currentError);
            }
            List<ContentItem> version;
            string versionError;
            if (!OpcContent.TryExtract(versionPath, out version, out versionError))
            {
                throw new InvalidOperationException("无法读取版本文件内容。" + versionError);
            }

            var currentLabel = "当前文档（" + Path.GetFileName(currentPath) + "）";
            var versionLabel = "版本（" + Path.GetFileName(versionPath) + "）";
            var ext = Path.GetExtension(currentPath ?? "").ToLowerInvariant();

            if (ext == ".xlsx" || ext == ".pptx")
            {
                // Excel 单元格有稳定地址、PowerPoint 形状有稳定名称，按键值对比，与顺序无关；
                // 语义：oldItems=版本（旧状态），newItems=当前（新状态）
                return ContentDiff.CompareKeyed(Items(version), Items(current), currentLabel, versionLabel);
            }
            // Word 段落按文档顺序对齐。比较只看段落内容，不看段落编号：
            // 编号会随前面段落的增删而平移，把它算进比较键会把没改的部分大面积误报成差异。
            return ContentDiff.CompareLines(version, current, currentLabel, versionLabel);
        }

        private static IEnumerable<KeyValuePair<string, string>> Items(List<ContentItem> items)
        {
            foreach (var item in items)
            {
                yield return new KeyValuePair<string, string>(item.Location, item.Text);
            }
        }
    }
}
