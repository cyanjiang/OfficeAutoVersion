using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Core;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;

namespace OfficeAutoVersion.Selftest
{
    /// <summary>
    /// 核心版本引擎 + Common 层（功能区/设置窗体）的自测程序。
    /// 直接链接交付源码编译，测的就是线上跑的代码。
    /// 运行方式：dotnet run --project tests/OfficeAutoVersion.Selftest
    /// </summary>
    public static class Program
    {
        private static int _passed;
        private static int _failed;
        private static readonly List<string> Failures = new List<string>();

        [STAThread]
        public static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            // 诊断模式：直接对两个真实文件生成差异摘要（不跑自测），
            // 用于在用户机器上核对摘要是否准确。
            //   OfficeAutoVersion.Selftest.exe diff <当前文件> <版本文件>
            if (args != null && args.Length >= 3 &&
                string.Equals(args[0], "diff", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var diagnostic = VersionDiffer.Diff(args[1], args[2]);
                    Console.WriteLine(diagnostic.ToText());
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("失败：" + ex.Message);
                    return 1;
                }
            }

            // UI 快照模式（开发期用）：把三个窗口各画一张 PNG，方便检查排版、裁字、尺寸。
            //   OfficeAutoVersion.Selftest.exe ui <输出目录>
            if (args != null && args.Length >= 2 &&
                string.Equals(args[0], "ui", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    RenderUi(args[1]);
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("失败：" + ex);
                    return 1;
                }
            }

            var root = Path.Combine(Path.GetTempPath(), "OfficeAutoVersionSelftest_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(root);
            var settingsPath = Path.Combine(root, "settings.json");
            OptionsStore.OverrideSettingsPath = settingsPath;

            try
            {
                RunAll(root);
            }
            catch (Exception ex)
            {
                Fail("自测程序异常", ex.ToString());
            }
            finally
            {
                try
                {
                    Directory.Delete(root, true);
                }
                catch
                {
                    // 忽略清理失败
                }
            }

            Console.WriteLine();
            Console.WriteLine("=====================================");
            Console.WriteLine("通过 " + _passed + " 项，失败 " + _failed + " 项。");
            if (_failed > 0)
            {
                foreach (var failure in Failures)
                {
                    Console.WriteLine("  [FAIL] " + failure);
                }
            }
            Console.WriteLine("=====================================");
            return _failed == 0 ? 0 : 1;
        }

        private static void RunAll(string root)
        {
            Section("路径与快照");
            TestSnapshotBasics(root);

            Section("去重与保留数量");
            TestDedupeAndPrune(root);

            Section("文件被占用时的去重");
            TestDedupeWhileLocked(root);

            Section("集中保存目录");
            TestCentralStore(root);

            Section("设置持久化");
            TestOptionsStore(root);

            Section("文件名格式");
            TestFileNamePattern(root);

            Section("边界情况");
            TestEdgeCases(root);

            Section("功能区");
            TestRibbon(root);

            Section("设置窗体");
            TestSettingsForm(root);

            Section("界面配色（跟随 Office 主题）");
            TestTheme(root);

            Section("版本历史窗体（比较/恢复）");
            TestVersionHistoryForm(root);

            Section("差异摘要（内容提取与比较）");
            TestDiffEngine(root);

            Section("日志");
            TestLogger(root);

            Section("加载项类（Word/Excel/PowerPoint）");
            TestAddInClasses();
        }

        // ------------------------------------------------------------ 路径与快照

        private static void TestSnapshotBasics(string root)
        {
            var dir = Path.Combine(root, "snapshot");
            Directory.CreateDirectory(dir);
            var doc = Path.Combine(dir, "报告 v1.docx");
            File.WriteAllText(doc, "第一版内容", System.Text.Encoding.UTF8);

            var logger = new Logger(Path.Combine(root, "snapshot.log"));
            var options = new VersionOptions { MaxVersionsToKeep = 5, SkipIfUnchanged = true, LogEnabled = true };
            var manager = new VersionManager(options, logger);

            Check(manager.GetVersionsDirectory(doc).Equals(Path.Combine(dir, "报告 v1_Versions"), StringComparison.OrdinalIgnoreCase),
                "默认模式：版本文件夹 = 原文件同目录下的「报告 v1_Versions」");

            var first = manager.CreateVersion(doc);
            Check(first != null && File.Exists(first), "保存前调用 → 生成版本文件");
            Check(first != null && Regex.IsMatch(Path.GetFileName(first), @"^报告 v1_\d{8}-\d{6}\.docx$"),
                "版本文件名符合「名称_yyyyMMdd-HHmmss.docx」: " + (first == null ? "(null)" : Path.GetFileName(first)));
            Check(first != null && File.ReadAllText(first, System.Text.Encoding.UTF8) == "第一版内容", "版本内容 = 保存前的磁盘内容");

            var second = manager.CreateVersion(doc);
            Check(second == null, "内容未变化 → 跳过，不产生重复版本");

            File.WriteAllText(doc, "第二版内容", System.Text.Encoding.UTF8);
            var third = manager.CreateVersion(doc);
            Check(third != null && File.Exists(third), "内容变化后再次保存 → 生成新版本");
            Check(third != first, "新旧版本文件名不同");
            Check(manager.ListVersions(doc).Count == 2, "累计 2 个版本");
        }

        // ------------------------------------------------------------ 去重与保留数量

        private static void TestDedupeAndPrune(string root)
        {
            var dir = Path.Combine(root, "prune");
            Directory.CreateDirectory(dir);
            var doc = Path.Combine(dir, "数据.xlsx");
            var logger = new Logger(Path.Combine(root, "prune.log"));
            var options = new VersionOptions { MaxVersionsToKeep = 3, SkipIfUnchanged = false };
            var manager = new VersionManager(options, logger);

            for (var i = 1; i <= 6; i++)
            {
                File.WriteAllText(doc, "第 " + i + " 次保存", System.Text.Encoding.UTF8);
                manager.CreateVersion(doc);
                var expected = Math.Min(i, 3);
                var actual = manager.ListVersions(doc).Count;
                Check(actual == expected, "第 " + i + " 次保存后版本数 = " + expected + "（实际 " + actual + "）");
                System.Threading.Thread.Sleep(1100); // 保证时间戳可区分，验证“保留最新”的排序
            }

            var versions = manager.ListVersions(doc);
            var contents = versions.Select(v => File.ReadAllText(v, System.Text.Encoding.UTF8)).ToList();
            Check(contents.Contains("第 6 次保存") && contents.Contains("第 5 次保存") && contents.Contains("第 4 次保存"),
                "超出上限后保留的是最新的 3 个版本");
            Check(!contents.Contains("第 3 次保存") && !contents.Contains("第 1 次保存"), "更早的版本已被清理");
        }

        // ------------------------------------------------------------ 集中保存目录

        private static void TestCentralStore(string root)
        {
            var docDir = Path.Combine(root, "central", "docs", "sub");
            Directory.CreateDirectory(docDir);
            var doc = Path.Combine(docDir, "月报.docx");
            File.WriteAllText(doc, "月报内容", System.Text.Encoding.UTF8);

            var options = new VersionOptions
            {
                StorageMode = VersionStorageMode.CentralStore,
                CentralRoot = Path.Combine(root, "central", "store"),
                MaxVersionsToKeep = 10,
            };
            var manager = new VersionManager(options, new Logger(Path.Combine(root, "central.log")));

            var expected = manager.GetVersionsDirectory(doc);
            var centralRoot = Path.Combine(root, "central", "store");
            Check(expected.StartsWith(centralRoot, StringComparison.OrdinalIgnoreCase),
                "集中模式：版本目录位于统一根目录之下");
            Check(expected.EndsWith(Path.Combine("docs", "sub", "月报_Versions"), StringComparison.OrdinalIgnoreCase),
                "集中模式：镜像原文件目录结构 -> " + expected);
            Check(!expected.Substring(centralRoot.Length).Contains(":"), "镜像出去的路径中不含盘符冒号");

            var version = manager.CreateVersion(doc);
            Check(version != null && version.StartsWith(expected, StringComparison.OrdinalIgnoreCase), "集中模式下快照落盘位置正确");
        }

        // ------------------------------------------------------------ 设置持久化

        private static void TestOptionsStore(string root)
        {
            var word = new VersionOptions
            {
                Enabled = true,
                StorageMode = VersionStorageMode.CentralStore,
                VersionFolderName = "{name}_备份",
                CentralRoot = @"D:\Versions",
                MaxVersionsToKeep = 7,
                SkipIfUnchanged = false,
                FileNamePattern = "{name}_{yyyyMMdd}",
                LogEnabled = false,
            };
            OptionsStore.Save("Word", word);

            var loaded = OptionsStore.Load("Word");
            Check(loaded.Enabled == word.Enabled, "设置往返：Enabled");
            Check(loaded.StorageMode == VersionStorageMode.CentralStore, "设置往返：StorageMode");
            Check(loaded.VersionFolderName == "{name}_备份", "设置往返：VersionFolderName");
            Check(loaded.CentralRoot == @"D:\Versions", "设置往返：CentralRoot");
            Check(loaded.MaxVersionsToKeep == 7, "设置往返：MaxVersionsToKeep");
            Check(loaded.SkipIfUnchanged == false, "设置往返：SkipIfUnchanged");
            Check(loaded.FileNamePattern == "{name}_{yyyyMMdd}", "设置往返：FileNamePattern");
            Check(loaded.LogEnabled == false, "设置往返：LogEnabled");

            var excel = OptionsStore.Load("Excel");
            Check(excel.Enabled && excel.MaxVersionsToKeep == 20 && excel.VersionFolderName == VersionPaths.DefaultVersionFolderName,
                "未配置的应用返回默认值（Excel）");

            var json = File.ReadAllText(OptionsStore.OverrideSettingsPath);
            Check(json.Contains("\"Word\"") && !json.Contains("Excel"), "各应用配置互相隔离");

            var serialized = MiniJson.Serialize(new Dictionary<string, VersionOptions> { { "Word", word } });
            Check(serialized.Contains("\"Word\"") && serialized.Contains("\"Enabled\":true"), "MiniJson 序列化结果正确");
            var parsed = MiniJson.Parse(serialized) as Dictionary<string, object>;
            Check(parsed != null && parsed.Count == 1, "MiniJson 解析回字典");
            var back = MiniJson.ToVersionOptions(parsed == null ? null : parsed["Word"] as Dictionary<string, object>);
            Check(back != null && back.MaxVersionsToKeep == 7 && back.StorageMode == VersionStorageMode.CentralStore, "MiniJson 解析回配置对象");

            File.WriteAllText(OptionsStore.OverrideSettingsPath, "{ 这不是合法 JSON ");
            var fallback = OptionsStore.Load("Word");
            Check(fallback != null && fallback.Enabled && fallback.MaxVersionsToKeep == 20, "配置文件损坏时安全回退到默认值");
        }

        // ------------------------------------------------------------ 文件名格式

        private static void TestFileNamePattern(string root)
        {
            var dir = Path.Combine(root, "pattern");
            Directory.CreateDirectory(dir);

            var doc = Path.Combine(dir, "报告 v1.docx");
            File.WriteAllText(doc, "内容", System.Text.Encoding.UTF8);
            var manager = new VersionManager(
                new VersionOptions { FileNamePattern = "{name}-{yyyyMMdd}", SkipIfUnchanged = false, MaxVersionsToKeep = 0 },
                new Logger(Path.Combine(root, "pattern.log")));
            var version = manager.CreateVersion(doc);
            Check(version != null && Regex.IsMatch(Path.GetFileName(version), @"^报告 v1-\d{8}\.docx$"),
                "自定义格式 {name}-{yyyyMMdd}: " + (version == null ? "(null)" : Path.GetFileName(version)));

            var doc2 = Path.Combine(dir, "总结.docx");
            File.WriteAllText(doc2, "内容", System.Text.Encoding.UTF8);
            var manager2 = new VersionManager(
                new VersionOptions { FileNamePattern = "{name}_{yyyy年MM月dd日}", SkipIfUnchanged = false, MaxVersionsToKeep = 0 },
                new Logger(Path.Combine(root, "pattern2.log")));
            var version2 = manager2.CreateVersion(doc2);
            Check(version2 != null && Regex.IsMatch(Path.GetFileName(version2), @"^总结_\d{4}年\d{2}月\d{2}日\.docx$"),
                "中文日期格式: " + (version2 == null ? "(null)" : Path.GetFileName(version2)));

            var doc3 = Path.Combine(dir, "无扩展名模式.docx");
            File.WriteAllText(doc3, "内容", System.Text.Encoding.UTF8);
            var manager3 = new VersionManager(
                new VersionOptions { FileNamePattern = "{name}", SkipIfUnchanged = false, MaxVersionsToKeep = 0 },
                new Logger(Path.Combine(root, "pattern3.log")));
            var version3 = manager3.CreateVersion(doc3);
            Check(version3 != null && Path.GetExtension(version3) == ".docx", "格式不含 {ext} 时自动补齐扩展名");

            var folderName = VersionPaths.GetVersionsDirectory(doc, new VersionOptions { VersionFolderName = "{name}_Versions" });
            Check(Path.GetFileName(folderName) == "报告 v1_Versions", "版本文件夹名支持 {name} 占位符");
        }

        // ------------------------------------------------------------ 边界情况

        private static void TestEdgeCases(string root)
        {
            var dir = Path.Combine(root, "edge");
            Directory.CreateDirectory(dir);
            var logger = new Logger(Path.Combine(root, "edge.log"));
            var manager = new VersionManager(new VersionOptions(), logger);

            Check(manager.CreateVersion(null) == null, "文件路径为 null → 安全跳过");
            Check(manager.CreateVersion("") == null, "文件路径为空 → 安全跳过");
            Check(manager.CreateVersion(Path.Combine(dir, "不存在的文件.docx")) == null, "文件不存在（从未保存过）→ 安全跳过");

            var disabled = new VersionManager(new VersionOptions { Enabled = false }, logger);
            var doc = Path.Combine(dir, "开关.docx");
            File.WriteAllText(doc, "内容", System.Text.Encoding.UTF8);
            Check(disabled.CreateVersion(doc) == null, "总开关关闭 → 不生成版本");

            var rapid = new VersionManager(
                new VersionOptions { SkipIfUnchanged = false, MaxVersionsToKeep = 0 }, logger);
            var rapidDoc = Path.Combine(dir, "连续保存.docx");
            var created = new List<string>();
            for (var i = 0; i < 4; i++)
            {
                File.WriteAllText(rapidDoc, "第" + i + "版", System.Text.Encoding.UTF8);
                var version = rapid.CreateVersion(rapidDoc);
                if (version != null)
                {
                    created.Add(version);
                }
            }
            Check(created.Count == 4 && created.Distinct().Count() == 4 && created.All(File.Exists),
                "同一秒内 4 次连续保存 → 4 个互不冲突的版本文件");

            var weird = Path.Combine(dir, "项目 A&B（终版）v2.3.docx");
            File.WriteAllText(weird, "内容", System.Text.Encoding.UTF8);
            var weirdVersion = rapid.CreateVersion(weird);
            Check(weirdVersion != null && File.Exists(weirdVersion), "文件名含空格/括号/& → 正常生成版本");

            var generated = VersionPaths.BuildVersionFileName(weird, new VersionOptions { FileNamePattern = "{name}_{ts}" }, new DateTime(2026, 10, 25, 23, 55, 1));
            Check(generated.IndexOfAny(Path.GetInvalidFileNameChars()) < 0, "生成的文件名不含非法字符");
        }

        // ------------------------------------------------------------ 功能区

        private static void TestRibbon(string root)
        {
            var doc = Path.Combine(root, "ribbon", "演示.docx");
            var ribbon = new VersioningRibbon("Word", delegate () { return doc; }, null);
            var ui = new FakeRibbonUi();
            ribbon.Ribbon_Load(ui);

            var xml = ribbon.GetCustomUI(null);
            Check(!string.IsNullOrEmpty(xml) && xml.Contains("chkAutoVersionEnabled") && xml.Contains("自动版本"),
                "GetCustomUI 返回包含“自动建立版本”复选框的功能区 XML");
            Check(xml.Contains("getPressed=\"GetAutoVersionChecked\"") && xml.Contains("onAction=\"OnAutoVersionChecked\""),
                "复选框已绑定到勾选状态与切换回调");

            Check(ribbon.GetAutoVersionChecked(new FakeRibbonControl("chkAutoVersionEnabled")) == true, "默认状态下复选框为“勾选”");

            ribbon.OnAutoVersionChecked(new FakeRibbonControl("chkAutoVersionEnabled"), false);
            Check(OptionsStore.Load("Word").Enabled == false, "取消勾选 → 配置已持久化");
            Check(ui.InvalidateAllCount == 1, "切换复选框后刷新功能区");

            ribbon.OnAutoVersionChecked(new FakeRibbonControl("chkAutoVersionEnabled"), true);
            Check(OptionsStore.Load("Word").Enabled == true, "重新勾选 → 配置已持久化");

            var xmlDocument = XDocument.Parse(xml);
            var ids = xmlDocument.Descendants().Select(e => (string)e.Attribute("id")).Where(id => id != null).ToList();
            Check(ids.Contains("chkAutoVersionEnabled") && ids.Contains("btnOpenVersionFolder") && ids.Contains("btnSettings"),
                "功能区包含：复选框 / 打开版本文件夹 / 设置");
            Check(ids.Contains("btnVersionHistory"), "功能区包含“版本历史”按钮（比较 / 恢复入口）");
            Check(ids.Contains("OnVersionHistory") || xml.Contains("OnVersionHistory"),
                "“版本历史”按钮已绑定到 OnVersionHistory 回调");
        }

        // ------------------------------------------------------------ 设置窗体

        /// <summary>版本历史窗体：列表内容、按钮文案、以及“比较”是否转交给宿主（Office 自带能力）。</summary>
        private static void TestVersionHistoryForm(string root)
        {
            var dir = Path.Combine(root, "history");
            Directory.CreateDirectory(dir);
            var doc = Path.Combine(dir, "报告.docx");
            var manager = new VersionManager(new VersionOptions(), new Logger(Path.Combine(dir, "history.log")));

            File.WriteAllText(doc, "第一版");
            manager.CreateVersion(doc);
            File.WriteAllText(doc, "第二版");
            manager.CreateVersion(doc);

            var actions = new FakeHostActions();
            using (var form = new OfficeAutoVersion.Common.VersionHistoryForm("Word", doc, manager, actions))
            {
                Check(form.VersionCount == 2, "版本历史列出 2 个版本 -> 实际 " + form.VersionCount);
                Check(!string.IsNullOrEmpty(form.SelectedVersionPath), "默认选中最新版本");

                var selected = form.SelectedVersionPath;
                form.CompareSelected();
                Check(actions.ComparedPath == selected, "点“比较” -> 把选中的版本转交给宿主（Office 自带比较）");
                // 回归保护：比较必须显式带上“当前文档路径”，不能靠应用自己的 ActiveDocument
                // （版本历史是模态窗口，打开期间 ActiveDocument/ActiveWorkbook 可能为 null）
                Check(actions.ComparedCurrentPath == doc,
                    "“比较”显式传入当前文档路径 -> " + (actions.ComparedCurrentPath ?? "<null>"));
            }

            // 模态场景（真机就是这种）：窗口只记录意图，绝不在模态期间调用 Office
            var modalActions = new FakeHostActions();
            using (var form = new OfficeAutoVersion.Common.VersionHistoryForm("Word", doc, manager, modalActions))
            {
                form.IsModal = true;
                var selected = form.SelectedVersionPath;
                form.CompareSelected();
                Check(form.PendingAction == OfficeAutoVersion.Common.VersionHistoryAction.Compare,
                    "模态窗口：比较只记录动作，窗口关闭后再执行");
                Check(form.PendingVersionPath == selected, "模态窗口：记录了要比较的那个版本");
                Check(modalActions.ComparedPath == null && modalActions.ComparedCurrentPath == null,
                    "模态窗口：没有在显示期间调用 Office（避免 ActiveWorkbook 为 null）");
            }

            // 宿主动作为空时不应崩溃，也不应误触发
            using (var form = new OfficeAutoVersion.Common.VersionHistoryForm("Word", doc, manager, null))
            {
                Check(form.VersionCount == 2, "无宿主动作时仍能列出历史");
                form.CompareSelected();
                Check(true, "无宿主动作时“比较”安全退出");
            }

            // 没有版本的文档
            var emptyDoc = Path.Combine(dir, "空文档.docx");
            File.WriteAllText(emptyDoc, "x");
            using (var form = new OfficeAutoVersion.Common.VersionHistoryForm("Word", emptyDoc, manager, actions))
            {
                Check(form.VersionCount == 0, "没有版本时列表为空");
                Check(form.SelectedVersionPath == null, "没有版本时没有选中项");
            }
        }

        // ------------------------------------------------------------ 差异摘要

        /// <summary>
        /// 差异摘要：手工构造最小 xlsx/docx/pptx 包，验证内容提取与比较。
        /// 纯文件解析，不需要 Office 在场。
        /// </summary>
        private static void TestDiffEngine(string root)
        {
            var dir = Path.Combine(root, "diff");
            Directory.CreateDirectory(dir);

            // ---------- Excel：按键值（单元格地址）比较 ----------
            var xlsxOld = Path.Combine(dir, "old.xlsx");
            var xlsxNew = Path.Combine(dir, "new.xlsx");
            WriteXlsx(xlsxOld, new Dictionary<string, string>
            {
                { "Sheet1!A1", "甲" }, { "Sheet1!B2", "100" }, { "Sheet2!A1", "旧值" },
            });
            WriteXlsx(xlsxNew, new Dictionary<string, string>
            {
                { "Sheet1!A1", "乙" }, { "Sheet1!B2", "100" }, { "Sheet1!C3", "新格子" },
            });
            var xlsx = VersionDiffer.Diff(xlsxNew, xlsxOld);
            Check(xlsx.AddedCount == 1 && xlsx.RemovedCount == 1 && xlsx.ChangedCount == 1 && xlsx.TotalCount == 3,
                "xlsx：新增 1 / 删除 1 / 修改 1 -> 实际 " + xlsx.AddedCount + "/" + xlsx.RemovedCount + "/" + xlsx.ChangedCount);
            var changed = xlsx.Entries.FirstOrDefault(e => e.Kind == DiffKind.Changed);
            Check(changed != null && changed.Location == "Sheet1!A1" && changed.OldValue == "甲" && changed.NewValue == "乙",
                "xlsx：修改条目带地址与新旧值（" + (changed != null ? changed.Location : "<无>") + "）");
            Check(xlsx.Entries.Any(e => e.Kind == DiffKind.Added && e.Location == "Sheet1!C3"),
                "xlsx：新增条目定位到 Sheet1!C3");
            Check(xlsx.Entries.Any(e => e.Kind == DiffKind.Removed && e.Location == "Sheet2!A1"),
                "xlsx：删除条目定位到 Sheet2!A1");
            Check(xlsx.ToText().Contains("共 3 处差异"), "摘要文本包含统计行");
            Check(VersionDiffer.Diff(xlsxOld, xlsxOld).TotalCount == 0, "内容完全一致 → 0 处差异");

            // ---------- Word：按段落顺序（LCS）比较 ----------
            var docxOld = Path.Combine(dir, "old.docx");
            var docxNew = Path.Combine(dir, "new.docx");
            WriteDocx(docxOld, new[] { "第一段", "第二段", "第三段" });
            WriteDocx(docxNew, new[] { "第一段", "第二段改", "第三段", "新增的第四段" });
            var docx = VersionDiffer.Diff(docxNew, docxOld);
            Check(docx.ChangedCount == 1 && docx.AddedCount == 1 && docx.RemovedCount == 0 && docx.TotalCount == 2,
                "docx：改一段 + 加一段 -> 修改 1 / 新增 1，实际 "
                + docx.ChangedCount + "/" + docx.AddedCount + "/" + docx.RemovedCount);
            var docxChanged = docx.Entries.FirstOrDefault(e => e.Kind == DiffKind.Changed);
            Check(docxChanged != null && docxChanged.Location == "第2段" &&
                  docxChanged.OldValue == "第二段" && docxChanged.NewValue == "第二段改",
                "docx：修改条目带段落号与新旧内容（" + (docxChanged != null ? docxChanged.Location : "<无>") + "）");
            Check(docx.Entries.Any(e => e.Kind == DiffKind.Added && e.Location == "第4段" && e.NewValue == "新增的第四段"),
                "docx：新增条目定位到第4段");

            // 关键回归：在开头插入一段，后面所有段落编号都平移，但内容没变 → 不能误报
            var shiftOld = Path.Combine(dir, "shift-old.docx");
            var shiftNew = Path.Combine(dir, "shift-new.docx");
            WriteDocx(shiftOld, new[] { "甲", "乙", "丙", "丁" });
            WriteDocx(shiftNew, new[] { "开头新加的一段", "甲", "乙", "丙", "丁" });
            var shift = VersionDiffer.Diff(shiftNew, shiftOld);
            Check(shift.AddedCount == 1 && shift.RemovedCount == 0 && shift.ChangedCount == 0,
                "docx：开头插入一段只报 1 处新增，后面的段落不因编号平移被误报 -> 实际 "
                + shift.AddedCount + "/" + shift.RemovedCount + "/" + shift.ChangedCount);

            // 删除段落：位置标注为“版本第N段”（编号属于版本文件，不是当前文档）
            var delOld = Path.Combine(dir, "del-old.docx");
            var delNew = Path.Combine(dir, "del-new.docx");
            WriteDocx(delOld, new[] { "甲", "乙", "丙" });
            WriteDocx(delNew, new[] { "甲", "丙" });
            var del = VersionDiffer.Diff(delNew, delOld);
            Check(del.RemovedCount == 1 && del.AddedCount == 0 && del.ChangedCount == 0 &&
                  del.Entries.Any(e => e.Kind == DiffKind.Removed && e.Location == "版本第2段" && e.OldValue == "乙"),
                "docx：删除条目标注为“版本第N段”，计数不虚增 -> "
                + del.RemovedCount + "/" + del.AddedCount + "/" + del.ChangedCount);

            // 页码标注：页码来自宿主排版，传入查找器后应出现在条目行上
            var withPage = docx.ToText(delegate(string text) { return text == "第二段改" ? 5 : 0; });
            Check(withPage.Contains("第 5 页 · 第 2 段"),
                "docx：摘要按宿主提供的页码定位修改处 -> " + LineOf(withPage, "【修改】"));
            Check(withPage.Contains("涉及页：第 5 页"),
                "docx：摘要头部汇总涉及的页码 -> " + LineOf(withPage, "涉及页"));
            Check(withPage.Contains("变化：在「第二段」后插入「改」"),
                "docx：修改条目给出逐字变化说明 -> " + LineOf(withPage, "变化"));
            Check(withPage.Contains("位置：在第 1 段「第一段」之后、第 3 段「第三段」之前"),
                "docx：条目给出前后段落作为定位锚点 -> " + LineOf(withPage, "位置"));
            Check(!docx.ToText().Contains("页"), "没有页码来源时摘要不出现页码字样");

            // 逐字差异描述：插入、替换、删除三种说法都要能读懂
            Check(TextDiff.Describe("证明", "信用证明").StartsWith("在「证明」前插入「信用」"),
                "逐字差异：前缀插入 -> " + TextDiff.Describe("证明", "信用证明"));
            Check(TextDiff.Describe("30天", "60天").StartsWith("「30天」→「60天」"),
                "逐字差异：局部替换 -> " + TextDiff.Describe("30天", "60天"));
            Check(TextDiff.Describe("甲、乙", "甲").StartsWith("删除「、乙」"),
                "逐字差异：删除片段 -> " + TextDiff.Describe("甲、乙", "甲"));
            Check(TextDiff.Describe("同一段", "同一段") == "内容相同", "逐字差异：内容相同");

            // 规范化：制表符、换行、表格单元结束符都折成单个空格，用于和 Word 对象模型的文本对齐
            Check(ContentDiff.NormalizeKey("  a\tb\r\nc\u0007 ") == "a b c",
                "文本规范化：各类空白折成单空格 -> " + ContentDiff.NormalizeKey("  a\tb\r\nc\u0007 "));

            // ---------- PowerPoint：按形状（键值）比较 ----------
            var pptxOld = Path.Combine(dir, "old.pptx");
            var pptxNew = Path.Combine(dir, "new.pptx");
            WritePptx(pptxOld, new[]
            {
                new[] { new[] { "标题 1", "旧标题" }, new[] { "内容 2", "要点A" } },
                new[] { new[] { "标题 1", "第二页" } },
            });
            WritePptx(pptxNew, new[]
            {
                new[] { new[] { "标题 1", "新标题" }, new[] { "内容 2", "要点A" } },
                new[] { new[] { "标题 1", "第二页" } },
            });
            var pptx = VersionDiffer.Diff(pptxNew, pptxOld);
            Check(pptx.ChangedCount == 1 && pptx.TotalCount == 1,
                "pptx：标题形状修改 1 处 -> 实际 " + pptx.ChangedCount);
            var pptEntry = pptx.Entries.FirstOrDefault();
            Check(pptEntry != null && pptEntry.Location == "幻灯片1 · 标题 1" && pptEntry.OldValue == "旧标题" && pptEntry.NewValue == "新标题",
                "pptx：修改条目定位到形状（" + (pptEntry != null ? pptEntry.Location : "<无>") + "）");

            // ---------- 不支持 / 损坏的文件给出可读提示 ----------
            var legacy = Path.Combine(dir, "legacy.doc");
            File.WriteAllText(legacy, "not a zip");
            List<ContentItem> unused;
            string error;
            Check(!OpcContent.TryExtract(legacy, out unused, out error) && error.Contains("旧版"),
                "旧版二进制格式给出可读提示 -> " + (error ?? "<null>"));
            var broken = Path.Combine(dir, "broken.xlsx");
            File.WriteAllText(broken, "junk");
            Check(!OpcContent.TryExtract(broken, out unused, out error),
                "损坏的文件返回失败而不是抛异常");

            // ---------- 被占用的当前文件（模拟 Office 打开时持有写句柄）----------
            var locked = Path.Combine(dir, "locked.docx");
            WriteDocx(locked, new[] { "占用中也能读" });
            using (var hold = new FileStream(locked, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
            {
                // Word 打开文档时正是这种持有方式：自己可读写、只允许别人读。
                // 若按 ZipFile.OpenRead 的 FileShare.Read 打开会撞共享冲突，必须用最宽松共享模式。
                List<ContentItem> lockedItems;
                string lockedError;
                Check(OpcContent.TryExtract(locked, out lockedItems, out lockedError)
                        && lockedItems != null && lockedItems.Count == 1,
                    "文件被 Office 式占用时仍能提取内容 -> "
                    + (lockedItems != null ? lockedItems.Count + " 条" : (lockedError ?? "<null>")));
            }

            // ---------- 界面入口 ----------
            var doc = Path.Combine(dir, "汇总.docx");
            WriteDocx(doc, new[] { "正文" });
            var manager = new VersionManager(new VersionOptions(), new Logger(Path.Combine(dir, "diff.log")));
            manager.CreateVersion(doc);
            WriteDocx(doc, new[] { "正文", "多出来的一段" });
            manager.CreateVersion(doc);
            using (var form = new OfficeAutoVersion.Common.VersionHistoryForm("Word", doc, manager, null))
            {
                Check(FindButtonText(form, "差异摘要"), "版本历史窗口包含“差异摘要”按钮");
            }
            using (var summaryForm = new OfficeAutoVersion.Common.DiffSummaryForm(xlsx, xlsxNew, xlsxOld))
            {
                var text = FindTextBoxText(summaryForm);
                Check(text != null && text.Contains("【修改】") && text.Contains("Sheet1!A1") && text.Contains("原文："),
                    "差异摘要窗体展示明细（逐条含原文）-> " +
                    (text == null ? "<null>" : text.Length + "字, 首行=" + LineOf(text, "【")));
            }
            using (var summaryForm = new OfficeAutoVersion.Common.DiffSummaryForm(docx, docxNew, docxOld,
                delegate(string text) { return text == "第二段改" ? 7 : 0; }))
            {
                var text = FindTextBoxText(summaryForm);
                Check(text != null && text.Contains("第 7 页"), "差异摘要窗体把宿主给的页码标出来");
            }
        }

        /// <summary>从摘要文本里取出含指定关键字的第一个非空行，用于失败时看清实际输出。</summary>
        private static string LineOf(string text, string keyword)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "<空>";
            }
            foreach (var line in text.Split('\n'))
            {
                var trimmed = line.Trim();
                if (trimmed.Length > 0 && trimmed.Contains(keyword))
                {
                    return trimmed;
                }
            }
            return "<未找到“" + keyword + "”>";
        }

        // ------------------------------------------------------------ UI 快照

        /// <summary>把三个窗口按浅色/深色各渲染一套 PNG，开发期用肉眼核对排版（不参与自测计分）。</summary>
        private static void RenderUi(string outDir)
        {
            Directory.CreateDirectory(outDir);
            // 截图不读本机偏好，固定按下面两套调色板渲染，保证可复现
            UiAppearance.IgnorePreference = true;

            var longPath = @"D:\项目资料\2026年度\第三季度\市场部\华东大区\客户回访记录\季度经营分析报告（最终版）.docx";

            var docDir = Path.Combine(outDir, "docs");
            Directory.CreateDirectory(docDir);
            var doc = Path.Combine(docDir, "投标文件.docx");
            var manager = new VersionManager(new VersionOptions(), new Logger(Path.Combine(docDir, "ui.log")));
            WriteDocx(doc, new[] { "第一章 投标函", "致：某某单位", "我方愿意参加本次投标。" });
            manager.CreateVersion(doc);
            WriteDocx(doc, new[] { "第一章 投标函", "致：某某单位", "我方愿意参加本次投标，报价见附件。" });
            manager.CreateVersion(doc);
            WriteDocx(doc, new[] { "第一章 投标函", "致：某某单位", "我方愿意参加本次投标，报价见附件。", "法人代表授权书" });
            manager.CreateVersion(doc);

            var versions = manager.ListVersions(doc);
            RenderSet(outDir, string.Empty, OfficeThemeKind.Colorful, longPath, doc, manager, versions);
            RenderSet(outDir, "-dark", OfficeThemeKind.Black, longPath, doc, manager, versions);

            UiTheme.ApplyAutomatic();
            UiAppearance.IgnorePreference = false;
        }

        private static void RenderSet(string outDir, string suffix, OfficeThemeKind kind, string longPath,
            string doc, VersionManager manager, List<string> versions)
        {
            UiTheme.Apply(kind, OfficeTheme.AccentFor("WINWORD"));
            WriteFormShot(new SettingsForm("Word", longPath), Path.Combine(outDir, "settings" + suffix + ".png"));
            WriteFormShot(new VersionHistoryForm("Word", doc, manager, new FakeHostActions()),
                Path.Combine(outDir, "history" + suffix + ".png"));
            if (versions.Count > 0)
            {
                var summary = VersionDiffer.Diff(doc, versions[versions.Count - 1]);
                WriteFormShot(new DiffSummaryForm(summary, doc, versions[versions.Count - 1]),
                    Path.Combine(outDir, "diff" + suffix + ".png"));
            }
        }

        private static void WriteFormShot(Form form, string path)
        {
            using (form)
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(30, 30);
                form.Show();
                Application.DoEvents();
                System.Threading.Thread.Sleep(400);
                Application.DoEvents();
                using (var bitmap = new Bitmap(form.Width, form.Height))
                {
                    form.DrawToBitmap(bitmap, new Rectangle(0, 0, form.Width, form.Height));
                    bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                }
                Console.WriteLine("shot=" + path + " " + form.Width + "x" + form.Height);
                form.Close();
            }
        }

        /// <summary>粗略亮度（0-255），用来判断底色是深还是浅。</summary>
        private static double Brightness(Color color)
        {
            return 0.299 * color.R + 0.587 * color.G + 0.114 * color.B;
        }

        /// <summary>WCAG 对比度（1:1 - 21:1）。</summary>
        private static double Contrast(Color a, Color b)
        {
            var la = RelativeLuminance(a);
            var lb = RelativeLuminance(b);
            var high = Math.Max(la, lb);
            var low = Math.Min(la, lb);
            return (high + 0.05) / (low + 0.05);
        }

        private static double RelativeLuminance(Color color)
        {
            return 0.2126 * Channel(color.R) + 0.7152 * Channel(color.G) + 0.0722 * Channel(color.B);
        }

        private static double Channel(int value)
        {
            var s = value / 255.0;
            return s <= 0.03928 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
        }

        private static bool FindComboBoxItem(Control root, string text)
        {
            foreach (Control control in root.Controls)
            {
                var combo = control as ComboBox;
                if (combo != null)
                {
                    foreach (var item in combo.Items)
                    {
                        if (string.Equals(Convert.ToString(item), text, StringComparison.Ordinal))
                        {
                            return true;
                        }
                    }
                }
                if (FindComboBoxItem(control, text))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool FindButtonText(Control root, string text)
        {
            foreach (Control control in root.Controls)
            {
                var button = control as Button;
                if (button != null && button.Text == text)
                {
                    return true;
                }
                if (FindButtonText(control, text))
                {
                    return true;
                }
            }
            return false;
        }

        private static string FindTextBoxText(Control root)
        {
            foreach (Control control in root.Controls)
            {
                var box = control as TextBox;
                if (box != null)
                {
                    return box.Text;
                }
                var nested = FindTextBoxText(control);
                if (nested != null)
                {
                    return nested;
                }
            }
            return null;
        }

        // ---------- 最小 Office 文件构造（仅供自测） ----------

        /// <summary>把若干文本条目写成一个 zip 包。</summary>
        private static void WriteZip(string path, Dictionary<string, string> entries)
        {
            using (var stream = new FileStream(path, FileMode.Create))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                foreach (var kv in entries)
                {
                    var entry = zip.CreateEntry(kv.Key);
                    using (var writer = new StreamWriter(entry.Open(), new System.Text.UTF8Encoding(false)))
                    {
                        writer.Write(kv.Value);
                    }
                }
            }
        }

        /// <summary>写一个最小的 xlsx：单元格键为 “工作表名!地址”（如 Sheet1!A1），值一律存为共享字符串。</summary>
        private static void WriteXlsx(string path, Dictionary<string, string> cells)
        {
            var shared = cells.Values.Distinct().ToList();
            var sharedIndex = shared.Select((value, index) => new { value, index })
                .ToDictionary(x => x.value, x => x.index);

            var sheets = new List<string>();
            foreach (var key in cells.Keys)
            {
                var name = key.Split('!')[0];
                if (!sheets.Contains(name))
                {
                    sheets.Add(name);
                }
            }

            var workbook = new StringBuilder();
            workbook.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            workbook.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            workbook.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>");
            for (var i = 0; i < sheets.Count; i++)
            {
                workbook.Append("<sheet name=\"" + sheets[i] + "\" sheetId=\"" + (i + 1) + "\" r:id=\"rId" + (i + 1) + "\"/>");
            }
            workbook.Append("</sheets></workbook>");

            var rels = new StringBuilder();
            rels.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            rels.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (var i = 0; i < sheets.Count; i++)
            {
                rels.Append("<Relationship Id=\"rId" + (i + 1) + "\" ");
                rels.Append("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" ");
                rels.Append("Target=\"worksheets/sheet" + (i + 1) + ".xml\"/>");
            }
            rels.Append("</Relationships>");

            var sst = new StringBuilder();
            sst.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sst.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + shared.Count + "\" uniqueCount=\"" + shared.Count + "\">");
            foreach (var value in shared)
            {
                sst.Append("<si><t>" + value + "</t></si>");
            }
            sst.Append("</sst>");

            var entries = new Dictionary<string, string>
            {
                { "xl/workbook.xml", workbook.ToString() },
                { "xl/_rels/workbook.xml.rels", rels.ToString() },
                { "xl/sharedStrings.xml", sst.ToString() },
            };
            for (var i = 0; i < sheets.Count; i++)
            {
                var sheetName = sheets[i];
                var xml = new StringBuilder();
                xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                foreach (var key in cells.Keys)
                {
                    if (key.Split('!')[0] != sheetName)
                    {
                        continue;
                    }
                    var address = key.Split('!')[1];
                    var value = cells[key];
                    xml.Append("<row r=\"1\"><c r=\"" + address + "\" t=\"s\"><v>" + sharedIndex[value] + "</v></c></row>");
                }
                xml.Append("</sheetData></worksheet>");
                entries["xl/worksheets/sheet" + (i + 1) + ".xml"] = xml.ToString();
            }
            WriteZip(path, entries);
        }

        /// <summary>写一个最小的 docx：每段一个 w:p/w:r/w:t。</summary>
        private static void WriteDocx(string path, string[] paragraphs)
        {
            var xml = new StringBuilder();
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            xml.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");
            foreach (var paragraph in paragraphs)
            {
                xml.Append("<w:p><w:r><w:t>" + paragraph + "</w:t></w:r></w:p>");
            }
            xml.Append("</w:body></w:document>");
            WriteZip(path, new Dictionary<string, string> { { "word/document.xml", xml.ToString() } });
        }

        /// <summary>写一个最小的 pptx：slides[i] 是该页的 (形状名, 文本) 列表。</summary>
        private static void WritePptx(string path, string[][][] slides)
        {
            var entries = new Dictionary<string, string>();
            for (var s = 0; s < slides.Length; s++)
            {
                var xml = new StringBuilder();
                xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                xml.Append("<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" ");
                xml.Append("xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><p:cSld><p:spTree>");
                foreach (var shape in slides[s])
                {
                    xml.Append("<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"" + shape[0] + "\"/><p:nvPr/></p:nvSpPr>");
                    xml.Append("<p:txBody><a:p><a:r><a:t>" + shape[1] + "</a:t></a:r></a:p></p:txBody></p:sp>");
                }
                xml.Append("</p:spTree></p:cSld></p:sld>");
                entries["ppt/slides/slide" + (s + 1) + ".xml"] = xml.ToString();
            }
            WriteZip(path, entries);
        }

        private static void TestSettingsForm(string root)
        {
            var doc = Path.Combine(root, "form", "报告.docx");
            using (var form = new SettingsForm("Word", doc))
            {
                Check(form.Text.Contains("Word"), "设置窗体标题包含应用名");
                Check(form.ClientSize.Width >= 500 && form.ClientSize.Height >= 400, "设置窗体尺寸合理");
                Check(form.Controls.Count > 0, "设置窗体包含内容");
            }
        }

        // ------------------------------------------------------------ 界面配色

        /// <summary>
        /// 配色必须跟着 Office 走：浅色/深色来自注册表里的 Office 主题，主色来自宿主品牌色；
        /// 同时保证任何主题下对比度都够（深色背景上直接用深蓝字会糊成一片）。
        /// </summary>
        private static void TestTheme(string root)
        {
            var previousPath = UiAppearance.OverrideFilePath;
            var themeFile = Path.Combine(root, "ui", "ui.json");
            var themeDir = Path.GetDirectoryName(themeFile);
            if (!string.IsNullOrEmpty(themeDir))
            {
                Directory.CreateDirectory(themeDir);
            }

            try
            {
                UiAppearance.OverrideFilePath = themeFile;

                // 注册表取值 → 主题（Office：0 彩色 / 3 白色 / 4 深灰 / 5 黑色）
                Check(!OfficeTheme.IsDarkKind(OfficeTheme.FromRegistryValue(0)), "注册表 0 对应彩色主题（浅色）");
                Check(!OfficeTheme.IsDarkKind(OfficeTheme.FromRegistryValue(3)), "注册表 3 对应白色主题（浅色）");
                Check(OfficeTheme.FromRegistryValue(4) == OfficeThemeKind.DarkGray, "注册表 4 对应深灰色主题");
                Check(OfficeTheme.FromRegistryValue(5) == OfficeThemeKind.Black, "注册表 5 对应黑色主题");
                Check(!OfficeTheme.IsDarkKind(OfficeTheme.FromRegistryValue(2)), "未知取值按浅色处理，不会把浅色界面误判成深色");

                // 宿主品牌色
                Check(OfficeTheme.AccentFor("WINWORD") == Color.FromArgb(0x18, 0x5A, 0xBD), "Word 主色 = Word 蓝 #185ABD");
                Check(OfficeTheme.AccentFor("EXCEL") == Color.FromArgb(0x10, 0x7C, 0x41), "Excel 主色 = Excel 绿 #107C41");
                Check(OfficeTheme.AccentFor("POWERPNT") == Color.FromArgb(0xC4, 0x3E, 0x1C), "PowerPoint 主色 = 橙红 #C43E1C");
                Check(OfficeTheme.AccentFor("notepad") == OfficeTheme.DefaultAccent, "非三大应用的宿主退回 Office 蓝");

                // 浅色：浅画布 + 深文字 + 品牌色按钮
                UiTheme.Apply(OfficeThemeKind.Colorful, OfficeTheme.AccentFor("WINWORD"));
                Check(!UiTheme.IsDark, "彩色主题标记为浅色");
                Check(Brightness(UiTheme.WindowBack) > 230,
                    "浅色画布接近 Office 的 F3F2F1（亮度 " + Brightness(UiTheme.WindowBack).ToString("0") + "）");
                Check(Contrast(UiTheme.TextPrimary, UiTheme.CardBack) >= 10, "浅色：正文与卡片对比度 ≥ 10:1");
                Check(Contrast(UiTheme.TextMuted, UiTheme.CardBack) >= 4.5, "浅色：说明文字对比度 ≥ 4.5:1");
                Check(Contrast(Color.White, UiTheme.Accent) >= 4.5, "浅色：主按钮白字对比度 ≥ 4.5:1");

                // 深色：深画布 + 浅文字，强调色要调亮，否则深蓝字在深底上看不清
                UiTheme.Apply(OfficeThemeKind.Black, OfficeTheme.AccentFor("EXCEL"));
                Check(UiTheme.IsDark, "黑色主题标记为深色");
                Check(Brightness(UiTheme.WindowBack) < 60,
                    "深色画布足够暗（亮度 " + Brightness(UiTheme.WindowBack).ToString("0") + "）");
                Check(Contrast(UiTheme.TextPrimary, UiTheme.CardBack) >= 10, "深色：正文与卡片对比度 ≥ 10:1");
                Check(Contrast(UiTheme.TextMuted, UiTheme.CardBack) >= 4.5, "深色：说明文字对比度 ≥ 4.5:1");
                Check(Contrast(UiTheme.AccentText, UiTheme.CardBack) >= 4.5,
                    "深色：强调色文字被调亮到可读（对比度 " + Contrast(UiTheme.AccentText, UiTheme.CardBack).ToString("0.0") + ":1）");
                Check(Contrast(Color.White, UiTheme.Accent) >= 4.5, "深色：主按钮白字仍然清楚");

                // 偏好持久化
                UiAppearance.Save(ThemePreference.Dark);
                Check(UiAppearance.Load() == ThemePreference.Dark, "配色偏好可写入并读回（深色）");
                Check(File.ReadAllText(themeFile).Contains("dark"), "偏好文件内容可读（appearance=dark）");
                UiAppearance.Save(ThemePreference.Light);
                Check(UiAppearance.Load() == ThemePreference.Light, "配色偏好可写入并读回（浅色）");
                UiAppearance.Save(ThemePreference.Automatic);
                Check(UiAppearance.Load() == ThemePreference.Automatic, "配色偏好可恢复为跟随 Office 主题");

                // 强制浅色 / 深色时忽略注册表
                UiAppearance.Apply(ThemePreference.Dark);
                Check(UiTheme.IsDark, "选择“始终深色”时忽略注册表里的浅色主题");
                UiAppearance.Apply(ThemePreference.Light);
                Check(!UiTheme.IsDark, "选择“始终浅色”时忽略注册表里的深色主题");

                // 窗体确实用了调色板，而不是控件默认色
                UiAppearance.Save(ThemePreference.Dark);
                UiAppearance.Apply(ThemePreference.Dark);
                using (var form = new SettingsForm("Word", Path.Combine(root, "form", "报告.docx")))
                {
                    Check(UiTheme.IsDark && form.BackColor == UiTheme.WindowBack, "深色偏好下设置窗体用深色底");
                    Check(form.ForeColor == UiTheme.TextPrimary, "深色偏好下设置窗体用浅色文字");
                    Check(FindComboBoxItem(form, "跟随 Office 主题（推荐）"), "设置窗体提供“界面配色”下拉选项");
                }
            }
            finally
            {
                UiAppearance.OverrideFilePath = previousPath;
                UiTheme.ApplyAutomatic();
                try
                {
                    UiAppearance.EnsureApplied();
                }
                catch
                {
                    // 恢复失败不影响结论
                }
            }
        }

        // ------------------------------------------------------------ 日志

        private static void TestLogger(string root)
        {
            var logPath = Path.Combine(root, "logs", "t.log");
            var logger = new Logger(logPath);
            logger.Info("测试日志");
            logger.Warn("警告");
            logger.Error("错误", new InvalidOperationException("测试异常"));

            Check(File.Exists(logPath), "日志文件已创建");
            var text = File.ReadAllText(logPath);
            Check(text.Contains("测试日志") && text.Contains("[INFO]") && text.Contains("[WARN]") && text.Contains("[ERROR]"),
                "日志包含各级别记录");
            Check(text.Contains("测试异常"), "日志包含异常详情");
        }

        // ------------------------------------------------------------ 加载项类

        private static void TestAddInClasses()
        {
            var word = new OfficeAutoVersion.Word.ThisAddIn(null, null);
            Check(word != null, "Word 加载项类可实例化（ThisAddIn + VSTO 基础设施代码编译通过）");
            Check(word.Application == null, "Word 加载项的 Application 宿主字段可用");

            var excel = new OfficeAutoVersion.Excel.ThisAddIn(null, null);
            Check(excel != null, "Excel 加载项类可实例化");
            Check(excel.Application == null, "Excel 加载项的 Application 宿主字段可用");

            var powerpoint = new OfficeAutoVersion.PowerPoint.ThisAddIn(null, null);
            Check(powerpoint != null, "PowerPoint 加载项类可实例化");
            Check(powerpoint.Application == null, "PowerPoint 加载项的 Application 宿主字段可用");

            // 事件委托签名核对（与当前版本 PIA 一致才能编译通过）：
            //   Word: (Document, ref bool, ref bool)  Excel: (Workbook, bool, ref bool)  PPT: (Presentation, ref bool)
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.DocumentBeforeSave += delegate (Microsoft.Office.Interop.Word.Document doc, ref bool saveAsUi, ref bool cancel) { };
            wordApp.DocumentBeforeSave -= delegate (Microsoft.Office.Interop.Word.Document doc, ref bool saveAsUi, ref bool cancel) { };
            Check(true, "Word DocumentBeforeSave 委托签名 (Document, ref bool, ref bool) 匹配");

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.WorkbookBeforeSave += delegate (Microsoft.Office.Interop.Excel.Workbook wb, bool saveAsUi, ref bool cancel) { cancel = false; };
            excelApp.WorkbookBeforeSave -= delegate (Microsoft.Office.Interop.Excel.Workbook wb, bool saveAsUi, ref bool cancel) { cancel = false; };
            Check(true, "Excel WorkbookBeforeSave 委托签名 (Workbook, bool, ref bool) 匹配");

            var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            pptApp.PresentationBeforeSave += delegate (Microsoft.Office.Interop.PowerPoint.Presentation pres, ref bool cancel) { cancel = false; };
            pptApp.PresentationBeforeSave -= delegate (Microsoft.Office.Interop.PowerPoint.Presentation pres, ref bool cancel) { cancel = false; };
            Check(true, "PowerPoint PresentationBeforeSave 委托签名 (Presentation, ref bool) 匹配");
        }

        /// <summary>
        /// Office 在触发 BeforeSave 时仍持有文件句柄（以宽松共享方式打开）。
        /// 该场景下去重必须依然有效，否则每次 Ctrl+S 都会多出一个内容完全相同的版本。
        /// </summary>
        private static void TestDedupeWhileLocked(string root)
        {
            var dir = Path.Combine(root, "locked");
            Directory.CreateDirectory(dir);
            var doc = Path.Combine(dir, "被占用.docx");
            File.WriteAllText(doc, "第一版内容");

            var options = new VersionOptions();
            options.MaxVersionsToKeep = 0;
            var manager = new VersionManager(options, new Logger(Path.Combine(dir, "selftest.log")));

            var first = manager.CreateVersion(doc);
            Check(first != null, "文件首次快照生成成功");

            using (var handle = new FileStream(doc, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                var second = manager.CreateVersion(doc);
                Check(second == null, "文件被其它进程占用且内容未变 -> 跳过，不产生重复版本");
            }

            Check(manager.ListVersions(doc).Count == 1, "被占用期间版本数量仍为 1 -> 实际 " + manager.ListVersions(doc).Count);

            File.WriteAllText(doc, "第二版内容");
            var third = manager.CreateVersion(doc);
            Check(third != null, "内容确实发生变化 -> 正常生成新版本");
            Check(manager.ListVersions(doc).Count == 2, "版本数量变为 2 -> 实际 " + manager.ListVersions(doc).Count);
        }

        // ------------------------------------------------------------ 断言辅助

        private static void Section(string name)
        {
            Console.WriteLine();
            Console.WriteLine("== " + name + " ==");
        }

        private static void Check(bool condition, string description)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("  [PASS] " + description);
            }
            else
            {
                Fail(description, string.Empty);
            }
        }

        private static void Fail(string description, string detail)
        {
            _failed++;
            Failures.Add(description);
            Console.WriteLine("  [FAIL] " + description);
            if (!string.IsNullOrEmpty(detail))
            {
                Console.WriteLine("         " + detail.Replace("\n", "\n         "));
            }
        }
    }
}
