using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 字符级差异描述：把“原文 → 现在”的变化说成人话，例如
    /// “在「证明」前插入「信用」”“「30 天」改为「60 天」”。
    /// 只用于展示，不参与统计口径。
    /// </summary>
    public static class TextDiff
    {
        /// <summary>每处引用最多显示多少字。</summary>
        private const int QuoteLimit = 24;
        /// <summary>引用前后文最多多少字。</summary>
        private const int ContextLimit = 8;
        /// <summary>最多列出几处字符改动。</summary>
        private const int MaxHunks = 3;
        /// <summary>参与 LCS 的最大字数（两边都要满足），超出退化为“整体改写”。</summary>
        private const int LcsCharLimit = 1200;

        /// <summary>
        /// 生成一句可直接阅读的变化说明。oldText 为原文（版本），newText 为现在。
        /// </summary>
        public static string Describe(string oldText, string newText)
        {
            var a = oldText ?? "";
            var b = newText ?? "";
            if (string.Equals(a, b, StringComparison.Ordinal))
            {
                return "内容相同";
            }
            if (a.Length == 0)
            {
                return "原为空段落，现在填入「" + Clip(b) + "」";
            }
            if (b.Length == 0)
            {
                return "内容被清空（原「" + Clip(a) + "」）";
            }

            List<Hunk> hunks;
            if (!TryDiff(a, b, out hunks))
            {
                return "整段改写（" + a.Length + " 字 → " + b.Length + " 字）";
            }
            if (hunks.Count == 0)
            {
                return "内容相同";
            }

            var sb = new StringBuilder();
            var shown = Math.Min(hunks.Count, MaxHunks);
            for (var i = 0; i < shown; i++)
            {
                if (i > 0)
                {
                    sb.Append('；');
                }
                if (hunks.Count > 1)
                {
                    sb.Append(Mark(i)).Append(' ');
                }
                sb.Append(Render(hunks[i]));
            }
            if (hunks.Count > shown)
            {
                sb.Append("（共 ").Append(hunks.Count).Append(" 处字符改动，仅列前 ")
                    .Append(shown).Append(" 处）");
            }
            sb.Append("（").Append(a.Length).Append(" 字 → ").Append(b.Length).Append(" 字）");
            return sb.ToString();
        }

        /// <summary>一处字符级改动：删除什么、插入什么、前后各是什么内容。</summary>
        private sealed class Hunk
        {
            public Hunk(string deleted, string inserted, string before, string after)
            {
                Deleted = deleted ?? "";
                Inserted = inserted ?? "";
                Before = before ?? "";
                After = after ?? "";
            }

            public string Deleted { get; private set; }
            public string Inserted { get; private set; }
            public string Before { get; private set; }
            public string After { get; private set; }
        }

        /// <summary>
        /// 按字符做 LCS 对齐，把差异切成若干处改动。返回 false 表示规模太大、不做细粒度比较。
        /// </summary>
        private static bool TryDiff(string a, string b, out List<Hunk> hunks)
        {
            hunks = new List<Hunk>();

            // 先剥掉公共前后缀：文档里通常只有个别字不同，剥完规模就很小了
            var start = 0;
            var maxStart = Math.Min(a.Length, b.Length);
            while (start < maxStart && a[start] == b[start])
            {
                start++;
            }
            var endA = a.Length;
            var endB = b.Length;
            while (endA > start && endB > start && a[endA - 1] == b[endB - 1])
            {
                endA--;
                endB--;
            }
            if (endA == start && endB == start)
            {
                return true; // 内容相同
            }

            var midA = a.Substring(start, endA - start);
            var midB = b.Substring(start, endB - start);
            if (midA.Length > LcsCharLimit || midB.Length > LcsCharLimit)
            {
                // 中间部分太大：只报“这一块被改写”，避免几十万次比较卡住界面
                hunks.Add(new Hunk(midA, midB, Tail(a, start), Head(a, endA)));
                return true;
            }

            var n = midA.Length;
            var m = midB.Length;
            var lcs = new int[n + 1, m + 1];
            for (var i = n - 1; i >= 0; i--)
            {
                for (var j = m - 1; j >= 0; j--)
                {
                    lcs[i, j] = midA[i] == midB[j]
                        ? lcs[i + 1, j + 1] + 1
                        : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
                }
            }

            var x = 0;
            var y = 0;
            while (x < n || y < m)
            {
                if (x < n && y < m && midA[x] == midB[y])
                {
                    x++;
                    y++;
                    continue;
                }
                // 收集连续的一段改动
                var delStart = x;
                var del = new StringBuilder();
                var ins = new StringBuilder();
                while (x < n || y < m)
                {
                    if (x < n && y < m && midA[x] == midB[y])
                    {
                        break;
                    }
                    if (y >= m)
                    {
                        del.Append(midA[x]);
                        x++;
                    }
                    else if (x >= n)
                    {
                        ins.Append(midB[y]);
                        y++;
                    }
                    else if (lcs[x + 1, y] >= lcs[x, y + 1])
                    {
                        del.Append(midA[x]);
                        x++;
                    }
                    else
                    {
                        ins.Append(midB[y]);
                        y++;
                    }
                }
                hunks.Add(new Hunk(del.ToString(), ins.ToString(),
                    Tail(a, start + delStart), Head(a, start + x)));
            }
            return true;
        }

        /// <summary>渲染一处改动，尽量说明“改在什么地方”。</summary>
        private static string Render(Hunk h)
        {
            var before = h.Before;
            var after = h.After;

            if (h.Deleted.Length == 0)
            {
                if (before.Length > 0)
                {
                    return "在「" + Clip(before) + "」后插入「" + Clip(h.Inserted) + "」";
                }
                if (after.Length > 0)
                {
                    return "在「" + Clip(after) + "」前插入「" + Clip(h.Inserted) + "」";
                }
                return "插入「" + Clip(h.Inserted) + "」";
            }

            if (h.Inserted.Length == 0)
            {
                return "删除「" + Clip(h.Deleted) + "」" +
                       (after.Length > 0 ? "（其后为「" + Clip(after) + "」）" : "");
            }

            // 替换：把改动处连同紧邻的上下文一起对比，读者能看出改在句子的哪个部位
            string oldWindow;
            string newWindow;
            if (before.Length > 0)
            {
                oldWindow = before + h.Deleted;
                newWindow = before + h.Inserted;
            }
            else if (after.Length > 0)
            {
                oldWindow = h.Deleted + after;
                newWindow = h.Inserted + after;
            }
            else
            {
                oldWindow = h.Deleted;
                newWindow = h.Inserted;
            }
            return "「" + Clip(oldWindow) + "」→「" + Clip(newWindow) + "」";
        }

        /// <summary>位置序号：①②③…</summary>
        private static string Mark(int index)
        {
            const string marks = "①②③④⑤⑥⑦⑧⑨";
            return index >= 0 && index < marks.Length ? marks[index].ToString() : (index + 1) + ".";
        }

        /// <summary>取 text 中 end 之前的最后若干个字（作为前文）。</summary>
        private static string Tail(string text, int end)
        {
            if (string.IsNullOrEmpty(text) || end <= 0)
            {
                return "";
            }
            if (end > text.Length)
            {
                end = text.Length;
            }
            var start = Math.Max(0, end - ContextLimit);
            return text.Substring(start, end - start);
        }

        /// <summary>取 text 中 start 之后的若干个字（作为后文）。</summary>
        private static string Head(string text, int start)
        {
            if (string.IsNullOrEmpty(text) || start >= text.Length)
            {
                return "";
            }
            if (start < 0)
            {
                start = 0;
            }
            var length = Math.Min(ContextLimit, text.Length - start);
            return text.Substring(start, length);
        }

        /// <summary>引用文本：压成一行、去首尾空白，过长截断。</summary>
        private static string Clip(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "空";
            }
            var line = text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (line.Length == 0)
            {
                return "空";
            }
            return line.Length > QuoteLimit ? line.Substring(0, QuoteLimit) + "…" : line;
        }
    }
}
