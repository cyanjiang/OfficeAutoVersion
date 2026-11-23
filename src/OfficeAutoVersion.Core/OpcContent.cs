using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace OfficeAutoVersion.Core
{
    /// <summary>从 Office 开放格式（docx/xlsx/pptx，即 OPC zip 包）里提取“带位置的内容行”。</summary>
    public sealed class ContentItem
    {
        public ContentItem(string location, string text)
        {
            Location = location ?? "";
            Text = text ?? "";
        }

        /// <summary>便于人阅读的位置标识，如 "Sheet1!B3"、"第12段"、"幻灯片2 · 标题"。</summary>
        public string Location { get; private set; }
        /// <summary>该位置的内容（用于比较）。</summary>
        public string Text { get; private set; }
    }

    /// <summary>
    /// 内容提取器：纯文件解析（zip + XML），不打开 Office、不碰对象模型，
    /// 因此在模态窗口期间也可以安全调用。
    /// </summary>
    public static class OpcContent
    {
        /// <summary>
        /// 提取文件内容。返回 false 时 error 给出用户可读的原因
        /// （旧版二进制格式不支持、文件损坏等）。
        /// </summary>
        public static bool TryExtract(string path, out List<ContentItem> items, out string error)
        {
            items = null;
            error = null;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                error = "文件不存在：" + path;
                return false;
            }

            try
            {
                var ext = Path.GetExtension(path).ToLowerInvariant();
                switch (ext)
                {
                    case ".docx":
                        items = ExtractDocx(path);
                        return true;
                    case ".xlsx":
                        items = ExtractXlsx(path);
                        return true;
                    case ".pptx":
                        items = ExtractPptx(path);
                        return true;
                    case ".doc":
                    case ".xls":
                    case ".ppt":
                        error = "旧版二进制格式（" + ext + "）不支持差异摘要，请改用“比较”功能查看。";
                        return false;
                    default:
                        error = "不支持的文件类型：" + ext;
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "读取文件内容失败：" + ex.Message;
                return false;
            }
        }

        // ------------------------------------------------------------ Word（docx）

        /// <summary>
        /// 打开 OPC 包。注意：当前文档必然正被 Office 打开并持有写句柄，
        /// ZipFile.OpenRead 内部只申请 FileShare.Read，会与 Office 的写入共享冲突，
        /// 因此用最宽松的共享模式手工打开；保存刚结束的瞬间仍可能撞锁，带几次重试。
        /// </summary>
        private static ZipArchive OpenZip(string path)
        {
            const int maxAttempts = 4;
            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                        FileShare.ReadWrite | FileShare.Delete);
                    return new ZipArchive(stream, ZipArchiveMode.Read);
                }
                catch (IOException)
                {
                    if (attempt >= maxAttempts)
                    {
                        throw;
                    }
                    System.Threading.Thread.Sleep(120);
                }
            }
        }

        /// <summary>Word：正文里每个段落一行（含表格中的段落），位置用“第N段”。</summary>
        private static List<ContentItem> ExtractDocx(string path)
        {
            var items = new List<ContentItem>();
            using (var zip = OpenZip(path))
            {
                var entry = zip.GetEntry("word/document.xml");
                if (entry == null)
                {
                    throw new InvalidDataException("不是有效的 Word 文档（缺少 word/document.xml）");
                }
                var doc = LoadXml(entry);
                // w 命名空间
                var ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                var paragraphs = doc.SelectNodes("//w:p", ns);
                if (paragraphs == null)
                {
                    return items;
                }
                var index = 0;
                foreach (XmlNode p in paragraphs)
                {
                    index++;
                    var text = ConcatParagraphText(p);
                    if (text.Length == 0)
                    {
                        continue; // 空段落没有比较价值，跳过（段落编号仍按文档顺序递增）
                    }
                    items.Add(new ContentItem("第" + index + "段", text));
                }
            }
            return items;
        }

        // ------------------------------------------------------------ Excel（xlsx）

        /// <summary>Excel：每个有内容的单元格一行，位置用 "工作表!单元格地址"。</summary>
        private static List<ContentItem> ExtractXlsx(string path)
        {
            var items = new List<ContentItem>();
            using (var zip = OpenZip(path))
            {
                // 1) 共享字符串表
                var shared = new List<string>();
                var sst = zip.GetEntry("xl/sharedStrings.xml");
                if (sst != null)
                {
                    var doc = LoadXml(sst);
                    var ns = new XmlNamespaceManager(doc.NameTable);
                    ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    foreach (XmlNode si in doc.SelectNodes("//s:si", ns))
                    {
                        shared.Add(ConcatNodeText(si,
                            "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "t"));
                    }
                }

                // 2) 工作表清单：sheet 名 → r:id
                var wb = zip.GetEntry("xl/workbook.xml");
                if (wb == null)
                {
                    throw new InvalidDataException("不是有效的 Excel 文档（缺少 xl/workbook.xml）");
                }
                var wbDoc = LoadXml(wb);
                var wbNs = new XmlNamespaceManager(wbDoc.NameTable);
                wbNs.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                // 3) r:id → 目标文件
                var rels = new Dictionary<string, string>();
                var relEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relEntry != null)
                {
                    var relDoc = LoadXml(relEntry);
                    foreach (XmlNode rel in relDoc.DocumentElement.ChildNodes)
                    {
                        var id = rel.Attributes != null ? rel.Attributes["Id"] : null;
                        var target = rel.Attributes != null ? rel.Attributes["Target"] : null;
                        if (id != null && target != null)
                        {
                            rels[id.Value] = target.Value.Replace("/xl/", "").TrimStart('/');
                        }
                    }
                }

                // 4) 逐个工作表读单元格
                foreach (XmlNode sheet in wbDoc.SelectNodes("//s:sheets/s:sheet", wbNs))
                {
                    var name = sheet.Attributes != null && sheet.Attributes["name"] != null
                        ? sheet.Attributes["name"].Value
                        : "Sheet";
                    var rid = sheet.Attributes != null && sheet.Attributes["r:id"] != null
                        ? sheet.Attributes["r:id"].Value
                        : null;
                    if (rid == null || !rels.ContainsKey(rid))
                    {
                        continue;
                    }
                    var target = rels[rid];
                    if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                    {
                        target = "xl/" + target;
                    }
                    var sheetEntry = zip.GetEntry(target);
                    if (sheetEntry == null)
                    {
                        continue;
                    }
                    ExtractSheet(sheetEntry, name, shared, items);
                }
            }
            return items;
        }

        private static void ExtractSheet(ZipArchiveEntry entry, string sheetName, List<string> shared,
            List<ContentItem> items)
        {
            var doc = LoadXml(entry);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            foreach (XmlNode cell in doc.SelectNodes("//s:c", ns))
            {
                var address = cell.Attributes != null && cell.Attributes["r"] != null
                    ? cell.Attributes["r"].Value
                    : "?";
                var type = cell.Attributes != null && cell.Attributes["t"] != null
                    ? cell.Attributes["t"].Value
                    : "n";

                string value = null;
                if (type == "s")
                {
                    // 共享字符串：v 的文本是索引
                    var vNode = cell.SelectSingleNode("s:v", ns);
                    int index;
                    if (vNode != null && int.TryParse(vNode.InnerText, NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out index) &&
                        index >= 0 && index < shared.Count)
                    {
                        value = shared[index];
                    }
                }
                else if (type == "inlineStr")
                {
                    value = ConcatNodeText(cell,
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "t");
                }
                else
                {
                    var vNode = cell.SelectSingleNode("s:v", ns);
                    if (vNode != null)
                    {
                        value = vNode.InnerText;
                    }
                }

                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }
                items.Add(new ContentItem(sheetName + "!" + address, value));
            }
        }

        // ------------------------------------------------------------ PowerPoint（pptx）

        /// <summary>PowerPoint：每张幻灯片的每个有文字的形状一行，位置用形状名。</summary>
        private static List<ContentItem> ExtractPptx(string path)
        {
            var items = new List<ContentItem>();
            using (var zip = OpenZip(path))
            {
                // 幻灯片文件名形如 ppt/slides/slide12.xml，按数字排序保证顺序稳定
                var slideNumbers = new List<int>();
                foreach (var e in zip.Entries)
                {
                    var m = System.Text.RegularExpressions.Regex.Match(e.FullName,
                        @"^ppt/slides/slide(\d+)\.xml$");
                    if (m.Success)
                    {
                        slideNumbers.Add(int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture));
                    }
                }
                slideNumbers.Sort();

                foreach (var number in slideNumbers)
                {
                    var entry = zip.GetEntry("ppt/slides/slide" + number + ".xml");
                    if (entry == null)
                    {
                        continue;
                    }
                    var doc = LoadXml(entry);
                    var ns = new XmlNamespaceManager(doc.NameTable);
                    ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                    ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    var shapes = doc.SelectNodes("//p:sp", ns);
                    if (shapes == null)
                    {
                        continue;
                    }
                    var shapeIndex = 0;
                    foreach (XmlNode sp in shapes)
                    {
                        shapeIndex++;
                        var text = ConcatNodeText(sp,
                            "http://schemas.openxmlformats.org/drawingml/2006/main", "t");
                        if (text.Length == 0)
                        {
                            continue;
                        }
                        // 优先用形状名（“标题 1”“内容占位符 2”），退化用序号
                        var nameNode = sp.SelectSingleNode("p:nvSpPr/p:cNvPr/@name", ns);
                        var shapeName = nameNode != null && nameNode.Value.Length > 0
                            ? nameNode.Value
                            : "形状" + shapeIndex;
                        items.Add(new ContentItem("幻灯片" + number + " · " + shapeName, text));
                    }
                }
            }
            return items;
        }

        // ------------------------------------------------------------ 工具

        /// <summary>把 zip 条目读成 XmlDocument。</summary>
        private static XmlDocument LoadXml(ZipArchiveEntry entry)
        {
            using (var stream = entry.Open())
            {
                var doc = new XmlDocument();
                doc.Load(stream);
                return doc;
            }
        }

        /// <summary>拼接节点下指定本地名的所有子孙节点的文本（按文档顺序）。</summary>
        private static string ConcatNodeText(XmlNode root, string nsUri, string localName)
        {
            var sb = new StringBuilder();
            AppendNodeText(root, nsUri, localName, sb);
            return sb.ToString();
        }

        private const string WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>
        /// 段落文本：除了 w:t，还要把制表符（w:tab）和换行（w:br/w:cr）算进去，
        /// 否则同一段落里“有制表符/手动换行”的内容会被压成连续字符串，与 Word 里的实际文本对不上。
        /// </summary>
        private static string ConcatParagraphText(XmlNode paragraph)
        {
            var sb = new StringBuilder();
            AppendParagraphText(paragraph, sb);
            return sb.ToString();
        }

        private static void AppendParagraphText(XmlNode node, StringBuilder sb)
        {
            if (node.NodeType == XmlNodeType.Element && node.NamespaceURI == WordNs)
            {
                switch (node.LocalName)
                {
                    case "t":
                        sb.Append(node.InnerText);
                        return;
                    case "tab":
                        sb.Append('\t');
                        return;
                    case "br":
                    case "cr":
                        sb.Append('\n');
                        return;
                }
            }
            foreach (XmlNode child in node.ChildNodes)
            {
                AppendParagraphText(child, sb);
            }
        }

        private static void AppendNodeText(XmlNode node, string nsUri, string localName, StringBuilder sb)
        {
            if (node.NodeType == XmlNodeType.Element && node.LocalName == localName && node.NamespaceURI == nsUri)
            {
                sb.Append(node.InnerText);
                return;
            }
            foreach (XmlNode child in node.ChildNodes)
            {
                AppendNodeText(child, nsUri, localName, sb);
            }
        }
    }
}
