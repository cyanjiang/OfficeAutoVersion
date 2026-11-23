using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAutoVersion.Core
{
    /// <summary>单个差异条目的类型。</summary>
    public enum DiffKind
    {
        Added = 0,   // 新增（版本里有、当前没有 → 相对“恢复方向”按 当前=新、版本=旧 记述）
        Removed = 1, // 删除
        Changed = 2, // 修改
    }

    /// <summary>一个差异条目：位置 + 旧值 + 新值。</summary>
    public sealed class DiffEntry
    {
        public DiffEntry(DiffKind kind, string location, string oldValue, string newValue)
        {
            Kind = kind;
            Location = location ?? "";
            OldValue = oldValue ?? "";
            NewValue = newValue ?? "";
        }

        public DiffKind Kind { get; private set; }
        /// <summary>差异发生的位置，例如 "Sheet1!B3"、"第12段"、"幻灯片2 · 标题"。</summary>
        public string Location { get; private set; }
        /// <summary>旧值（版本文件里的内容）。</summary>
        public string OldValue { get; private set; }
        public string NewValue { get; private set; }

        /// <summary>改动位置的前一段内容（当前文档里的，用于定位），没有则为空。</summary>
        public string ContextBefore { get; set; }
        /// <summary>改动位置的前一段的位置标识，如“第51段”。</summary>
        public string ContextBeforeLocation { get; set; }
        /// <summary>改动位置的后一段内容。</summary>
        public string ContextAfter { get; set; }
        /// <summary>改动位置的后一段的位置标识。</summary>
        public string ContextAfterLocation { get; set; }
    }

    /// <summary>
    /// 两个文件的内容差异摘要：统计 + 明细（条数有上限，超出部分只计数）。
    /// 方向约定：current（当前文档）是新状态，version（所选版本）是旧状态，
    /// 即“当前比这个版本多了什么 / 少了什么 / 改了什么”。
    /// </summary>
    public sealed class DiffSummary
    {
        public string LeftLabel;   // 当前文档的说明（通常是“当前”）
        public string RightLabel;  // 版本的说明（通常是版本文件名）

        public int AddedCount;     // 当前比版本多出来的
        public int RemovedCount;   // 当前比版本少掉的
        public int ChangedCount;   // 两边都有但内容不同的

        public List<DiffEntry> Entries = new List<DiffEntry>();
        /// <summary>未展示的条目数（Entries 达到上限时）。</summary>
        public int HiddenCount;

        public int TotalCount
        {
            get { return AddedCount + RemovedCount + ChangedCount; }
        }

        /// <summary>生成可直接展示 / 保存的纯文本摘要。</summary>
        public string ToText()
        {
            return ToText(null);
        }

        public string ToText(Func<string, int> pageOf)
        {
            return ToText(pageOf, true);
        }

        /// <summary>
        /// 生成纯文本摘要。pageOf 非空时，为每条差异标注它在当前文档里的页码
        /// （页码只能由排版结果得出，需宿主应用配合；取不到页码的条目不标注）。
        /// includeHeader=false 时省略标题和统计行（窗口已用图形化头部展示过），只留说明和明细。
        /// </summary>
        public string ToText(Func<string, int> pageOf, bool includeHeader)
        {
            var sb = new StringBuilder();
            if (includeHeader)
            {
                sb.AppendLine("差异摘要：当前文档  vs  " + (RightLabel ?? "所选版本"));
                sb.AppendLine(new string('=', 50));
                sb.AppendLine(string.Format("新增 {0} 处    删除 {1} 处    修改 {2} 处    共 {3} 处差异",
                    AddedCount, RemovedCount, ChangedCount, TotalCount));
                if (HiddenCount > 0)
                {
                    sb.AppendLine(string.Format("（明细最多列出 {0} 条，另有 {1} 条未列出）", Entries.Count, HiddenCount));
                }
            }
            else if (HiddenCount > 0)
            {
                sb.AppendLine(string.Format("（明细最多列出 {0} 条，另有 {1} 条未列出）", Entries.Count, HiddenCount));
            }

            var views = BuildViews(pageOf);
            AppendPageSummary(sb, views);

            if (pageOf != null)
            {
                sb.AppendLine("说明：段号是当前文档里的段落顺序号（空段落也计数）；页码按当前文档排版计算，仅作定位参考。");
            }
            if (TotalCount == 0)
            {
                sb.AppendLine();
                sb.AppendLine("两个文件的内容完全一致。");
                return sb.ToString();
            }

            if (includeHeader)
            {
                sb.AppendLine(new string('-', 50));
            }
            for (var i = 0; i < views.Count; i++)
            {
                var view = views[i];
                var entry = view.Entry;
                sb.AppendLine();
                sb.AppendLine(string.Format("{0}. {1}{2}", i + 1, TypeLabel(entry.Kind), FormatLocation(view)));

                if (entry.Kind == DiffKind.Added)
                {
                    sb.AppendLine("   内容：" + OneLine(entry.NewValue));
                    AppendPosition(sb, entry);
                }
                else if (entry.Kind == DiffKind.Removed)
                {
                    sb.AppendLine("   原文：" + OneLine(entry.OldValue));
                    AppendPosition(sb, entry);
                }
                else
                {
                    sb.AppendLine("   原文：" + OneLine(entry.OldValue));
                    sb.AppendLine("   现在：" + OneLine(entry.NewValue));
                    sb.AppendLine("   变化：" + TextDiff.Describe(entry.OldValue, entry.NewValue));
                    AppendPosition(sb, entry);
                }
            }
            return sb.ToString();
        }

        /// <summary>一条明细在展示时用到的派生信息（页码等）。</summary>
        private sealed class EntryView
        {
            public DiffEntry Entry;
            public int Page;
            /// <summary>页码是按邻近段落推算的（删除项），标注时要写“约”。</summary>
            public bool Approximate;
        }

        private List<EntryView> BuildViews(Func<string, int> pageOf)
        {
            var views = new List<EntryView>();
            foreach (var entry in Entries)
            {
                var view = new EntryView { Entry = entry };
                if (pageOf != null)
                {
                    // 新增/修改：文本在当前文档里，直接查；
                    // 删除：原文已经不在当前文档里，用它插回原位的前后段落推算大致页码。
                    var probe = entry.Kind == DiffKind.Removed ? null : entry.NewValue;
                    var page = ProbePage(pageOf, probe);
                    if (page <= 0 && entry.Kind == DiffKind.Removed)
                    {
                        page = ProbePage(pageOf, entry.ContextAfter);
                        if (page <= 0)
                        {
                            page = ProbePage(pageOf, entry.ContextBefore);
                        }
                        view.Approximate = page > 0;
                    }
                    view.Page = page;
                }
                views.Add(view);
            }
            return views;
        }

        private static int ProbePage(Func<string, int> pageOf, string text)
        {
            if (pageOf == null || string.IsNullOrEmpty(text))
            {
                return 0;
            }
            try
            {
                return pageOf(text);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>“涉及页：第 6、12 页（第 6 页 1 处；第 12 页 1 处）”</summary>
        private static void AppendPageSummary(StringBuilder sb, List<EntryView> views)
        {
            var counts = new SortedDictionary<int, int>();
            foreach (var view in views)
            {
                if (view.Page <= 0)
                {
                    continue;
                }
                int old = 0;
                counts.TryGetValue(view.Page, out old);
                counts[view.Page] = old + 1;
            }
            if (counts.Count == 0)
            {
                return;
            }
            var pages = new StringBuilder();
            var detail = new StringBuilder();
            foreach (var kv in counts)
            {
                if (pages.Length > 0)
                {
                    pages.Append('、');
                    detail.Append("；");
                }
                pages.Append(kv.Key);
                detail.Append("第 ").Append(kv.Key).Append(" 页 ").Append(kv.Value).Append(" 处");
            }
            sb.AppendLine("涉及页：第 " + pages + " 页（" + detail + "）");
        }

        private static string TypeLabel(DiffKind kind)
        {
            switch (kind)
            {
                case DiffKind.Added: return "【新增】";
                case DiffKind.Removed: return "【删除】";
                default: return "【修改】";
            }
        }

        /// <summary>位置展示：Word 段落写成“第 6 页 · 第 52 段”，其他（单元格/形状）保留原标识。</summary>
        private static string FormatLocation(EntryView view)
        {
            var location = view.Entry.Location ?? "";
            var isParagraph = location.StartsWith("第", StringComparison.Ordinal) ||
                              location.StartsWith("版本第", StringComparison.Ordinal);
            var pagePart = view.Page > 0
                ? (view.Approximate ? "约第 " + view.Page + " 页" : "第 " + view.Page + " 页")
                : "";
            if (!isParagraph)
            {
                return pagePart.Length > 0 ? location + " · " + pagePart : location;
            }
            var segment = FormatSegment(location);
            return pagePart.Length > 0 ? pagePart + " · " + segment : segment;
        }

        /// <summary>“第98段” → “第 98 段”；“版本第98段” → “版本第 98 段”。</summary>
        private static string FormatSegment(string location)
        {
            var text = location ?? "";
            var prefixLength = text.StartsWith("版本", StringComparison.Ordinal) ? 2 : 0;
            if (!text.EndsWith("段", StringComparison.Ordinal))
            {
                return text;
            }
            var digitStart = prefixLength;
            if (digitStart < text.Length && text[digitStart] == '第')
            {
                digitStart++;
            }
            if (digitStart >= text.Length - 1)
            {
                return text;
            }
            var digits = text.Substring(digitStart, text.Length - digitStart - 1);
            int number;
            if (!int.TryParse(digits, out number))
            {
                return text;
            }
            return text.Substring(0, prefixLength) + "第 " + number + " 段";
        }

        /// <summary>“位置：在第 51 段「…」之后、第 53 段「…」之前”</summary>
        private static void AppendPosition(StringBuilder sb, DiffEntry entry)
        {
            var hasBefore = !string.IsNullOrEmpty(entry.ContextBefore);
            var hasAfter = !string.IsNullOrEmpty(entry.ContextAfter);
            if (!hasBefore && !hasAfter)
            {
                return;
            }
            var label = entry.Kind == DiffKind.Removed ? "   大致位置：" : "   位置：";
            var text = new StringBuilder();
            if (hasBefore && hasAfter)
            {
                text.Append("在第 ").Append(SegmentNumber(entry.ContextBeforeLocation))
                    .Append(" 段「").Append(Snippet(entry.ContextBefore)).Append("」之后、第 ")
                    .Append(SegmentNumber(entry.ContextAfterLocation))
                    .Append(" 段「").Append(Snippet(entry.ContextAfter)).Append("」之前");
            }
            else if (hasBefore)
            {
                text.Append("紧接第 ").Append(SegmentNumber(entry.ContextBeforeLocation))
                    .Append(" 段「").Append(Snippet(entry.ContextBefore)).Append("」之后");
            }
            else
            {
                text.Append("紧接第 ").Append(SegmentNumber(entry.ContextAfterLocation))
                    .Append(" 段「").Append(Snippet(entry.ContextAfter)).Append("」之前");
            }
            sb.AppendLine(label + text);
        }

        /// <summary>从“第51段”这样的位置标识里取出数字；取不到时退回原文字。</summary>
        private static string SegmentNumber(string location)
        {
            var text = location ?? "";
            var digits = new StringBuilder();
            foreach (var c in text)
            {
                if (c >= '0' && c <= '9')
                {
                    digits.Append(c);
                }
            }
            return digits.Length > 0 ? digits.ToString() : text;
        }

        /// <summary>上下文片段：一行、截断，用作定位锚点。</summary>
        private static string Snippet(string text)
        {
            var line = OneLine(text);
            return line.Length > 30 ? line.Substring(0, 30) + "…" : line;
        }

        /// <summary>把多行值压成一行并截断，避免摘要被长段落撑爆。</summary>
        private static string OneLine(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "(空)";
            }
            var line = text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (line.Length == 0)
            {
                return "(空)";
            }
            if (line.Length > 160)
            {
                line = line.Substring(0, 160) + "…";
            }
            return line;
        }
    }

    /// <summary>
    /// 两种内容比较方式：
    /// 1) 键值比较——按“位置 → 内容”的字典对比（Excel 单元格、PowerPoint 形状）；
    /// 2) 行比较——按顺序对段落做 LCS 对齐（Word 段落）。
    /// </summary>
    public static class ContentDiff
    {
        /// <summary>摘要中最多列出的明细条数；超出只计数。</summary>
        public const int MaxEntries = 400;

        /// <summary>LCS 的规模上限：任一边超过此长度时退化为裁剪 + 顺序对齐（见 CompareLines）。</summary>
        private const int LcsLimit = 3000;

        /// <summary>
        /// 键值比较。oldItems/newItems 的 Key 是稳定的位置标识（如 Sheet1!B3），
        /// Value 是该位置的内容。Key 相同但 Value 不同记为修改；只在一边存在记为新增/删除。
        /// </summary>
        public static DiffSummary CompareKeyed(IEnumerable<KeyValuePair<string, string>> oldItems,
            IEnumerable<KeyValuePair<string, string>> newItems, string leftLabel, string rightLabel)
        {
            var summary = new DiffSummary { LeftLabel = leftLabel, RightLabel = rightLabel };
            var oldMap = new Dictionary<string, string>();
            foreach (var kv in oldItems)
            {
                oldMap[kv.Key] = kv.Value ?? "";
            }
            var seen = new HashSet<string>();

            // 先按新文件的顺序输出“新增/修改”，保持自然阅读顺序
            foreach (var kv in newItems)
            {
                string oldValue;
                var exists = oldMap.TryGetValue(kv.Key, out oldValue);
                var newValue = kv.Value ?? "";
                if (!exists)
                {
                    summary.AddedCount++;
                    Add(summary, new DiffEntry(DiffKind.Added, kv.Key, null, newValue));
                }
                else
                {
                    seen.Add(kv.Key);
                    if (!string.Equals(oldValue, newValue, StringComparison.Ordinal))
                    {
                        summary.ChangedCount++;
                        Add(summary, new DiffEntry(DiffKind.Changed, kv.Key, oldValue, newValue));
                    }
                }
            }

            // 版本里剩下的就是被删除的
            foreach (var kv in oldItems)
            {
                if (!seen.Contains(kv.Key))
                {
                    summary.RemovedCount++;
                    Add(summary, new DiffEntry(DiffKind.Removed, kv.Key, kv.Value ?? "", null));
                }
            }
            return summary;
        }

        /// <summary>
        /// 行（段落）比较。只比内容本身，不比位置编号——否则在某段之前插入/删除一段，
        /// 后面所有段落的编号都会平移，会把没改的段落全部误报成差异。
        /// 先裁掉公共前缀/后缀（版本间差异通常集中在中间），
        /// 剩余部分不大时用 LCS 对齐，太大时退化为顺序配对并如实标注。
        /// 相邻的一处删除 + 一处新增会配对成“修改”，更符合阅读直觉。
        /// </summary>
        public static DiffSummary CompareLines(IList<ContentItem> oldItems, IList<ContentItem> newItems,
            string leftLabel, string rightLabel)
        {
            var summary = new DiffSummary { LeftLabel = leftLabel, RightLabel = rightLabel };
            var n = oldItems.Count;
            var m = newItems.Count;

            // 去掉两端完全相同的部分
            int start = 0;
            while (start < n && start < m && Same(oldItems[start], newItems[start]))
            {
                start++;
            }
            int endOld = n, endNew = m;
            while (endOld > start && endNew > start && Same(oldItems[endOld - 1], newItems[endNew - 1]))
            {
                endOld--;
                endNew--;
            }

            var ops = new List<Op>();
            if (endOld - start <= LcsLimit && endNew - start <= LcsLimit)
            {
                LcsOps(ops, oldItems, newItems, start, endOld, endNew);
            }
            else
            {
                // 大文档兜底：顺序配对，多出来的按插入/删除计
                var count = Math.Min(endOld, endNew) - start;
                for (var i = 0; i < count; i++)
                {
                    if (!Same(oldItems[start + i], newItems[start + i]))
                    {
                        ops.Add(new Op(DiffKind.Changed, start + i, start + i));
                    }
                }
                for (var i = start + count; i < endOld; i++)
                {
                    ops.Add(new Op(DiffKind.Removed, i, start + count));
                }
                for (var j = start + count; j < endNew; j++)
                {
                    ops.Add(new Op(DiffKind.Added, -1, j));
                }
            }

            // 相邻的“删除+新增”配对成“修改”（1:1 贪心）
            for (var k = 0; k < ops.Count - 1; k++)
            {
                if (ops[k].Kind == DiffKind.Removed && ops[k + 1].Kind == DiffKind.Added)
                {
                    ops[k] = new Op(DiffKind.Changed, ops[k].OldIndex, ops[k + 1].NewIndex);
                    ops.RemoveAt(k + 1);
                }
            }

            foreach (var op in ops)
            {
                switch (op.Kind)
                {
                    case DiffKind.Changed:
                    {
                        summary.ChangedCount++;
                        var entry = new DiffEntry(DiffKind.Changed, newItems[op.NewIndex].Location,
                            oldItems[op.OldIndex].Text, newItems[op.NewIndex].Text);
                        // 改动处的上下文取当前文档的相邻段落，方便按“前后是什么”翻到原处
                        FillContext(entry, newItems, op.NewIndex - 1, op.NewIndex + 1);
                        Add(summary, entry);
                        break;
                    }
                    case DiffKind.Added:
                    {
                        summary.AddedCount++;
                        var entry = new DiffEntry(DiffKind.Added, newItems[op.NewIndex].Location,
                            null, newItems[op.NewIndex].Text);
                        FillContext(entry, newItems, op.NewIndex - 1, op.NewIndex + 1);
                        Add(summary, entry);
                        break;
                    }
                    default:
                    {
                        summary.RemovedCount++;
                        // 删除条目只存在于版本文件里，编号是版本文件的编号，明确标注；
                        // 当前文档里没有它，用“插回原位”的前后段落给它定位。
                        var entry = new DiffEntry(DiffKind.Removed,
                            "版本" + oldItems[op.OldIndex].Location, oldItems[op.OldIndex].Text, null);
                        FillContext(entry, newItems, op.NewIndex - 1, op.NewIndex);
                        Add(summary, entry);
                        break;
                    }
                }
            }
            return summary;
        }

        /// <summary>
        /// 记录改动处前后各一段的内容与位置。下标越界（文件开头/结尾）时留空。
        /// </summary>
        private static void FillContext(DiffEntry entry, IList<ContentItem> items,
            int beforeIndex, int afterIndex)
        {
            if (entry == null || items == null)
            {
                return;
            }
            if (beforeIndex >= 0 && beforeIndex < items.Count)
            {
                entry.ContextBefore = items[beforeIndex].Text;
                entry.ContextBeforeLocation = items[beforeIndex].Location;
            }
            if (afterIndex >= 0 && afterIndex < items.Count)
            {
                entry.ContextAfter = items[afterIndex].Text;
                entry.ContextAfterLocation = items[afterIndex].Location;
            }
        }

        private static bool Same(ContentItem a, ContentItem b)
        {
            return string.Equals(a.Text, b.Text, StringComparison.Ordinal);
        }

        /// <summary>
        /// 比较过程中的一个操作：修改/新增/删除，带两边的原始下标。
        /// 删除项没有 OldIndex 以外的对应内容，NewIndex 记的是它在新文档里的插回位置。
        /// </summary>
        private sealed class Op
        {
            public Op(DiffKind kind, int oldIndex, int newIndex)
            {
                Kind = kind;
                OldIndex = oldIndex;
                NewIndex = newIndex;
            }

            public DiffKind Kind { get; private set; }
            public int OldIndex { get; private set; }
            public int NewIndex { get; private set; }
        }

        /// <summary>LCS 动态规划：把公共子序列对齐，其余部分按编辑脚本顺序记为删除/插入。</summary>
        private static void LcsOps(List<Op> ops, IList<ContentItem> oldItems, IList<ContentItem> newItems,
            int start, int endOld, int endNew)
        {
            var n = endOld - start;
            var m = endNew - start;
            // lcs[i,j] = 旧文 i.. 与新文 j.. 的最长公共子序列长度
            var lcs = new int[n + 1, m + 1];
            for (var i = n - 1; i >= 0; i--)
            {
                for (var j = m - 1; j >= 0; j--)
                {
                    lcs[i, j] = Same(oldItems[start + i], newItems[start + j])
                        ? lcs[i + 1, j + 1] + 1
                        : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
                }
            }

            var i2 = 0;
            var j2 = 0;
            while (i2 < n && j2 < m)
            {
                if (Same(oldItems[start + i2], newItems[start + j2]))
                {
                    i2++;
                    j2++;
                }
                else if (lcs[i2 + 1, j2] >= lcs[i2, j2 + 1])
                {
                    // 删除：记录它在新文档里的“插回位置”，用来取前后段落做定位
                    ops.Add(new Op(DiffKind.Removed, start + i2, start + j2));
                    i2++;
                }
                else
                {
                    ops.Add(new Op(DiffKind.Added, -1, start + j2));
                    j2++;
                }
            }
            for (; i2 < n; i2++)
            {
                ops.Add(new Op(DiffKind.Removed, start + i2, start + j2));
            }
            for (; j2 < m; j2++)
            {
                ops.Add(new Op(DiffKind.Added, -1, start + j2));
            }
        }

        /// <summary>
        /// 规范化文本用作查找键：把各类空白/表格单元结束符等折成单个空格并去首尾。
        /// 用于把 XML 里提取的段落文本与 Word 对象模型里的段落文本对上号。
        /// </summary>
        public static string NormalizeKey(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }
            var sb = new StringBuilder(text.Length);
            var lastWasSpace = false;
            foreach (var c in text)
            {
                if (c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\f' || c == '\v' ||
                    c == '\u0007' /* 表格单元格结束符 */ || c == '\u00A0' || c == '\u3000' || c == '\u200B')
                {
                    if (!lastWasSpace)
                    {
                        sb.Append(' ');
                        lastWasSpace = true;
                    }
                }
                else
                {
                    sb.Append(c);
                    lastWasSpace = false;
                }
            }
            return sb.ToString().Trim();
        }

        private static void Add(DiffSummary summary, DiffEntry entry)
        {
            if (summary.Entries.Count < MaxEntries)
            {
                summary.Entries.Add(entry);
            }
            else
            {
                summary.HiddenCount++;
            }
        }
    }
}
