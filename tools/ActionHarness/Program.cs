using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using OfficeAutoVersion.Common;
using OfficeAutoVersion.Core;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using WordInterop = Microsoft.Office.Interop.Word;

namespace ActionHarness
{
    /// <summary>
    /// 用真实的 Word/Excel/PowerPoint 验证三个加载项里“比较 / 恢复”的真实链路：
    /// 反射调用 internal 的 XxxVersionActions，因此测试的就是交付的代码。
    /// 用法：先编译 OfficeAutoVersion.sln，再编译并运行本工程；
    /// 只想跑某个应用时传参数 excel / word / ppt。
    /// </summary>
    internal static class Program
    {
        private static int _failed;
        private static int _passed;

        private static int Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = System.Text.Encoding.UTF8;
            }
            catch
            {
                // 控制台不支持时忽略
            }

            var only = args.Length > 0 ? args[0].ToLowerInvariant() : null;

            if (only == null || only == "excel")
            {
                Run("Excel", RunExcelSuite);
            }
            if (only == null || only == "word")
            {
                Run("Word", RunWordSuite);
            }
            if (only == null || only == "ppt")
            {
                Run("PowerPoint", RunPowerPointSuite);
            }
            if (only == "ui")
            {
                var uiThread = new Thread(RunUiSuite);
                uiThread.SetApartmentState(ApartmentState.STA);
                uiThread.Start();
                uiThread.Join();
            }

            Console.WriteLine();
            Console.WriteLine("=====================================");
            Console.WriteLine(_failed == 0
                ? "RESULT: ALL PASSED (" + _passed + " checks)"
                : "RESULT: " + _failed + " FAILED, " + _passed + " passed");
            return _failed == 0 ? 0 : 1;
        }

        private static void Run(string name, Action suite)
        {
            Console.WriteLine();
            Console.WriteLine("=== " + name + " ===");
            try
            {
                suite();
            }
            catch (Exception ex)
            {
                Console.WriteLine("[!] suite crashed: " + ex.Message);
                _failed++;
            }
        }

        // ------------------------------------------------------------------ Excel

        private static void RunExcelSuite()
        {
            var root = NewDirectory("excel");
            var doc = Path.Combine(root, "harness.xlsx");
            var excel = new ExcelInterop.Application { Visible = false, DisplayAlerts = false };
            WaitForAddIn(excel.COMAddIns, "OfficeAutoVersion.Excel", 25);

            try
            {
                var workbook = excel.Workbooks.Add();
                Cell(workbook).Value2 = "one";
                workbook.SaveAs(doc);
                Thread.Sleep(1200);
                Cell(workbook).Value2 = "two";
                workbook.Save();
                Thread.Sleep(1200);
                Cell(workbook).Value2 = "three";
                workbook.Save();

                var versions = ListVersions(doc);
                Console.WriteLine("versions: " + versions.Count);
                Expect(versions.Count >= 2, "Excel: 保存三次后生成至少 2 个版本");

                var actions = CreateActions("OfficeAutoVersion.Excel.dll",
                    "OfficeAutoVersion.Excel.ExcelVersionActions",
                    new[] { typeof(Func<ExcelInterop.Application>), typeof(Func<VersionManager>), typeof(Action<bool>) },
                    new object[] { (Func<ExcelInterop.Application>)delegate { return excel; }, ManagerFactory, SuppressFlag });

                // 比较：Excel 自带“并排查看”
                ActivateWorkbook(excel, doc);
                // 并排查看/同步滚动属于视图功能，不可见时会失效，这里临时显示窗口
                excel.Visible = true;
                Invoke(actions, "CompareWithCurrent", doc, versions[0]);
                Thread.Sleep(400);
                Console.WriteLine("windows=" + excel.Windows.Count + " sync=" + excel.Windows.SyncScrollingSideBySide);
                Expect(excel.Windows.Count >= 2, "Excel: 比较时版本以第二个窗口打开");
                Expect(excel.Windows.SyncScrollingSideBySide, "Excel: 已开启自带的“并排查看”同步滚动");

                CloseOtherWorkbooks(excel, doc);

                // 关键回归：VSTO 可能在加载项初始化之前就创建功能区，那一刻 Application 还是 null。
                // 动作类必须“调用时才取 Application”，并且在未就绪时给出明确提示，
                // 而不是把 null 存下来、之后每次比较都抛 NullReferenceException。
                ExcelInterop.Application delayedApp = null;
                var delayedActions = CreateActions("OfficeAutoVersion.Excel.dll",
                    "OfficeAutoVersion.Excel.ExcelVersionActions",
                    new[] { typeof(Func<ExcelInterop.Application>), typeof(Func<VersionManager>), typeof(Action<bool>) },
                    new object[] { (Func<ExcelInterop.Application>)delegate { return delayedApp; }, ManagerFactory, SuppressFlag });
                try
                {
                    Invoke(delayedActions, "CompareWithCurrent", doc, versions[0]);
                    Expect(false, "Excel: 宿主未就绪时比较应失败并给出提示");
                }
                catch (TargetInvocationException ex)
                {
                    var inner = ex.InnerException;
                    Console.WriteLine("delayed-app error: " +
                                      (inner == null ? "<none>" : inner.GetType().Name + ": " + inner.Message));
                    Expect(!(inner is NullReferenceException), "Excel: 宿主未就绪时不应抛 NullReferenceException");
                    Expect(inner is InvalidOperationException, "Excel: 宿主未就绪时给出可读的 InvalidOperationException");
                }

                // 同一动作实例，宿主稍后就绪 → 必须能正常工作（说明取的是委托，不是快照）
                delayedApp = excel;
                Invoke(delayedActions, "CompareWithCurrent", doc, versions[0]);
                Thread.Sleep(400);
                Console.WriteLine("windows(delayed)=" + excel.Windows.Count);
                Expect(excel.Windows.Count >= 2, "Excel: 宿主稍后就绪后比较依然可用");
                CloseOtherWorkbooks(excel, doc);

                // 恢复
                ActivateWorkbook(excel, doc);
                excel.Visible = false;

                // 关键回归：打开另一个工作簿并让它成为“活动工作簿”。
                // 恢复目标仍是 doc——旧实现靠 Application.ActiveWorkbook 定位，
                // 会误把 other 当成目标（这正是用户看到“当前没有打开的工作簿”的那类问题）。
                var otherPath = Path.Combine(root, "other.xlsx");
                var other = excel.Workbooks.Add();
                Cell(other).Value2 = "keep-me";
                other.SaveAs(otherPath);
                other.Activate();
                Thread.Sleep(300);

                var oldest = versions[versions.Count - 1];
                var before = ListVersions(doc).Count;
                var snapshot = (string)Invoke(actions, "RestoreVersion", doc, oldest);
                Thread.Sleep(600);

                var content = ReadCell(excel, doc);
                Console.WriteLine("content=" + content + " snapshot=" +
                                  (snapshot == null ? "<none>" : Path.GetFileName(snapshot)));
                Expect(content == "one", "Excel: 恢复后单元格内容等于所选版本");
                Expect(!string.IsNullOrEmpty(snapshot), "Excel: 恢复前建立了安全快照");
                Expect(ListVersions(doc).Count == before + 1, "Excel: 恢复只新增 1 个快照");

                // 活动工作簿不是目标时，其它工作簿必须毫发无损
                var otherReopened = FindWorkbook(excel, otherPath);
                Expect(otherReopened != null, "Excel: 活动工作簿不是目标时，其它工作簿没有被关闭");
                if (otherReopened != null)
                {
                    otherReopened.Activate();
                    var otherValue = Convert.ToString(Cell(otherReopened).Value2);
                    Expect(otherValue == "keep-me", "Excel: 活动工作簿不是目标时，其它工作簿内容未被改写");
                    otherReopened.Close(false);
                }
            }
            finally
            {
                QuitQuietly(delegate
                {
                    excel.DisplayAlerts = false;
                    excel.Quit();
                });
            }
        }

        // ------------------------------------------------------------------- Word

        private static void RunWordSuite()
        {
            var root = NewDirectory("word");
            var doc = Path.Combine(root, "harness.docx");
            var word = new WordInterop.Application { Visible = false };
            WaitForAddIn(word.COMAddIns, "OfficeAutoVersion.Word", 25);

            try
            {
                var document = word.Documents.Add();
                document.Content.Text = "one";
                object path = doc;
                document.SaveAs2(ref path);
                Thread.Sleep(1200);
                document.Content.Text = "two";
                document.Save();
                Thread.Sleep(1200);
                document.Content.Text = "three";
                document.Save();

                var versions = ListVersions(doc);
                Console.WriteLine("versions: " + versions.Count);
                Expect(versions.Count >= 2, "Word: 保存三次后生成至少 2 个版本");

                var actions = CreateActions("OfficeAutoVersion.Word.dll",
                    "OfficeAutoVersion.Word.WordVersionActions",
                    new[] { typeof(Func<WordInterop.Application>), typeof(Func<VersionManager>), typeof(Action<bool>) },
                    new object[] { (Func<WordInterop.Application>)delegate { return word; }, ManagerFactory, SuppressFlag });

                // 比较：Word 自带“比较文档”
                ActivateDocument(word, doc);
                Console.WriteLine("documents before compare: " + word.Documents.Count);
                Invoke(actions, "CompareWithCurrent", doc, versions[0]);
                DumpDocuments(word);

                var comparison = FindComparisonDocument(word);
                Console.WriteLine("comparison doc: " + (comparison == null ? "<none>" : comparison.Name) +
                                  " revisions=" + (comparison == null ? -1 : comparison.Revisions.Count));
                Expect(comparison != null && comparison.Revisions.Count > 0,
                    "Word: 生成了带修订标记的对比文档");
                Expect(FindDocumentByPath(word, doc) != null,
                    "Word: 原文档在比较后仍然打开（未被误关闭）");

                CloseOtherDocuments(word, doc);

                // 恢复
                ActivateDocument(word, doc);
                var oldest = versions[versions.Count - 1];
                var before = ListVersions(doc).Count;
                var snapshot = (string)Invoke(actions, "RestoreVersion", doc, oldest);
                Thread.Sleep(600);

                var text = ActiveDocumentText(word, doc);
                Console.WriteLine("text=" + text + " snapshot=" +
                                  (snapshot == null ? "<none>" : Path.GetFileName(snapshot)));
                Expect(text.StartsWith("one", StringComparison.Ordinal), "Word: 恢复后正文等于所选版本");
                Expect(!string.IsNullOrEmpty(snapshot), "Word: 恢复前建立了安全快照");
                Expect(ListVersions(doc).Count == before + 1, "Word: 恢复只新增 1 个快照");

                // 页码映射：差异摘要的“改动在第几页”靠宿主按排版算，这里验证它真能给出页码
                var pageDoc = word.Documents.Add();
                pageDoc.Content.Text = "第一页的内容";
                var tail = pageDoc.Content;
                tail.Collapse(WordInterop.WdCollapseDirection.wdCollapseEnd);
                tail.InsertBreak(WordInterop.WdBreakType.wdPageBreak);
                tail.InsertAfter("第二页的内容");
                var pagePath = Path.Combine(root, "pages.docx");
                object pagePathObject = pagePath;
                pageDoc.SaveAs2(ref pagePathObject);
                Thread.Sleep(400);

                var pageMap = Invoke(actions, "GetTextPageMap", pagePath)
                    as System.Collections.Generic.Dictionary<string, int>;
                var firstPage = 0;
                var secondPage = 0;
                if (pageMap != null)
                {
                    pageMap.TryGetValue("第一页的内容", out firstPage);
                    pageMap.TryGetValue("第二页的内容", out secondPage);
                }
                Console.WriteLine("page map: count=" + (pageMap == null ? -1 : pageMap.Count) +
                                  " 第一页=" + firstPage + " 第二页=" + secondPage);
                Expect(pageMap != null && firstPage == 1 && secondPage == 2,
                    "Word: 能按排版给出“段落 → 页码”（改动可定位到页）");

                CloseWithoutSaving(pageDoc);
                ActivateDocument(word, doc);
            }
            finally
            {
                QuitQuietly(delegate { word.Quit(); });
            }
        }

        // ------------------------------------------------------------- PowerPoint

        private static void RunPowerPointSuite()
        {
            var root = NewDirectory("ppt");
            var doc = Path.Combine(root, "harness.pptx");
            var ppt = new PowerPointInterop.Application { Visible = Microsoft.Office.Core.MsoTriState.msoTrue };
            WaitForAddIn(ppt.COMAddIns, "OfficeAutoVersion.PowerPoint", 25);
            try
            {
                ppt.WindowState = PowerPointInterop.PpWindowState.ppWindowMinimized;
            }
            catch
            {
                // 忽略
            }

            try
            {
                var presentation = ppt.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                var slide = presentation.Slides.Add(1, PowerPointInterop.PpSlideLayout.ppLayoutTitle);
                slide.Shapes.Title.TextFrame.TextRange.Text = "first";
                presentation.SaveAs(doc, PowerPointInterop.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                Thread.Sleep(1200);
                presentation.Slides[1].Shapes.Title.TextFrame.TextRange.Text = "second";
                presentation.Save();
                Thread.Sleep(1200);
                presentation.Slides[1].Shapes.Title.TextFrame.TextRange.Text = "third";
                presentation.Save();

                var versions = ListVersions(doc);
                Console.WriteLine("versions: " + versions.Count);
                Expect(versions.Count >= 2, "PPT: 保存三次后生成至少 2 个版本");

                var actions = CreateActions("OfficeAutoVersion.PowerPoint.dll",
                    "OfficeAutoVersion.PowerPoint.PowerPointVersionActions",
                    new[] { typeof(Func<PowerPointInterop.Application>), typeof(Func<VersionManager>), typeof(Action<bool>) },
                    new object[] { (Func<PowerPointInterop.Application>)delegate { return ppt; }, ManagerFactory, SuppressFlag });

                // 恢复（PPT 的“比较”走 Presentation.Merge，会打开交互式审阅窗格，自动化环境不测）
                var oldest = versions[versions.Count - 1];
                var before = ListVersions(doc).Count;
                var snapshot = (string)Invoke(actions, "RestoreVersion", doc, oldest);
                Thread.Sleep(800);

                var title = ActiveSlideTitle(ppt, doc);
                Console.WriteLine("title=" + (title ?? "<error>") + " snapshot=" +
                                  (snapshot == null ? "<none>" : Path.GetFileName(snapshot)));
                Expect(title == "first", "PPT: 恢复后首屏标题等于所选版本（first）");
                Expect(!string.IsNullOrEmpty(snapshot), "PPT: 恢复前建立了安全快照");
                Expect(ListVersions(doc).Count == before + 1, "PPT: 恢复只新增 1 个快照");
            }
            finally
            {
                QuitQuietly(delegate { ppt.Quit(); });
            }
        }

        // -------------------------------------------------------------------- UI

        /// <summary>
        /// 把设置窗口真正显示出来，逐个检查控件是否完整可见（不被裁掉），并截图留证。
        /// 用一条很长的中文路径——这正是“信息显示不全”最容易出现的场景。
        /// </summary>
        private static void RunUiSuite()
        {
            var longPath = @"D:\项目资料\2026年度\第三季度\市场部\华东大区\客户回访记录\季度经营分析报告（最终版）.docx";
            try
            {
                using (var form = new SettingsForm("Word", longPath))
                {
                    form.Show();
                    Application.DoEvents();
                    Thread.Sleep(600);
                    Application.DoEvents();

                    var buttons = new List<Button>();
                    CollectButtons(form, buttons);
                    Expect(buttons.Count >= 4, "设置窗口有 4 个按钮 -> 实际 " + buttons.Count);
                    foreach (var button in buttons)
                    {
                        CheckFullyVisible(form, button, "按钮");
                    }
                    CheckFullyVisible(form, FindControlByTextPrefix(form, "设置文件："), "设置/日志路径");
                    CheckFullyVisible(form, FindControlByTextPrefix(form, "保存位置预览"), "预览标题");
                    CheckFullyVisible(form, FindControlByTextPrefix(form, "D:\\项目资料"), "预览路径");

                    var shot = Path.Combine(Path.GetTempPath(), "oav_settings_form.png");
                    using (var bmp = new Bitmap(form.Width, form.Height))
                    {
                        form.DrawToBitmap(bmp, new Rectangle(0, 0, form.Width, form.Height));
                        bmp.Save(shot, ImageFormat.Png);
                    }
                    Console.WriteLine("screenshot=" + shot);
                    form.Close();
                }
            }
            catch (Exception ex)
            {
                Expect(false, "设置窗口显示测试异常：" + ex.Message);
            }
        }

        private static void CollectButtons(Control parent, List<Button> result)
        {
            foreach (Control child in parent.Controls)
            {
                var button = child as Button;
                if (button != null)
                {
                    result.Add(button);
                }
                CollectButtons(child, result);
            }
        }

        private static Control FindControlByTextPrefix(Control parent, string prefix)
        {
            foreach (Control child in parent.Controls)
            {
                if (!string.IsNullOrEmpty(child.Text) &&
                    child.Text.StartsWith(prefix, StringComparison.Ordinal))
                {
                    return child;
                }
                var nested = FindControlByTextPrefix(child, prefix);
                if (nested != null)
                {
                    return nested;
                }
            }
            return null;
        }

        private static void CheckFullyVisible(Form form, Control control, string description)
        {
            if (control == null)
            {
                Expect(false, "设置窗口：找到控件 - " + description);
                return;
            }
            var inClient = form.RectangleToClient(control.Parent.RectangleToScreen(control.Bounds));
            var visible = form.ClientRectangle.Contains(inClient);
            var text = (control.Text ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
            if (text.Length > 16)
            {
                text = text.Substring(0, 16) + "...";
            }
            Expect(visible, "设置窗口内容完整可见（" + description + "）：" + text);
        }

        // ------------------------------------------------------------------ 基础

        private static Func<VersionManager> ManagerFactory
        {
            get { return delegate { return new VersionManager(new VersionOptions()); }; }
        }

        private static Action<bool> SuppressFlag
        {
            get { return delegate { }; }
        }

        private static string NewDirectory(string name)
        {
            var root = Path.Combine(Path.GetTempPath(),
                "oav_harness_" + name + "_" + Guid.NewGuid().ToString("N").Substring(0, 6));
            Directory.CreateDirectory(root);
            return root;
        }

        private static object CreateActions(string assemblyFile, string typeName, Type[] parameters, object[] arguments)
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, assemblyFile);
            var assembly = Assembly.LoadFrom(path);
            var type = assembly.GetType(typeName, true);
            var ctor = type.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic, null, parameters, null);
            if (ctor == null)
            {
                throw new InvalidOperationException("构造函数未找到：" + typeName);
            }
            return ctor.Invoke(arguments);
        }

        private static object Invoke(object target, string methodName, params object[] arguments)
        {
            var method = target.GetType().GetMethod(methodName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (method == null)
            {
                throw new InvalidOperationException("方法未找到：" + methodName);
            }
            return method.Invoke(target, arguments);
        }

        private static List<string> ListVersions(string documentPath)
        {
            return new VersionManager(new VersionOptions()).ListVersions(documentPath);
        }

        private static void QuitQuietly(Action quit)
        {
            try
            {
                quit();
            }
            catch
            {
                // 退出失败不影响结论
            }
        }

        // ------------------------------------------------------------- Excel 辅助

        private static ExcelInterop.Range Cell(ExcelInterop.Workbook workbook)
        {
            var sheet = (ExcelInterop.Worksheet)workbook.Sheets[1];
            return sheet.get_Range("A1", Type.Missing);
        }

        /// <summary>
        /// 等 VSTO 加载项真正就绪。Office 启动后是**异步**加载加载项的：
        /// 太早开始保存，DocumentBeforeSave 还没挂上，一个版本都不会生成
        /// （日志里连事件记录都没有，很容易误判成插件失效）。
        /// </summary>
        private static void WaitForAddIn(Microsoft.Office.Core.COMAddIns addIns, string progId, int timeoutSeconds)
        {
            var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    foreach (Microsoft.Office.Core.COMAddIn addIn in addIns)
                    {
                        if (string.Equals(addIn.ProgId, progId, StringComparison.OrdinalIgnoreCase) && addIn.Connect)
                        {
                            // Connect 之后 VSTO 才执行 ThisAddIn_Startup
                            Thread.Sleep(2000);
                            Console.WriteLine("add-in ready: " + progId);
                            return;
                        }
                    }
                }
                catch
                {
                    // 集合尚未准备好
                }
                Thread.Sleep(300);
            }
            Console.WriteLine("add-in NOT ready (timeout): " + progId);
        }

        private static ExcelInterop.Workbook FindWorkbook(ExcelInterop.Application excel, string documentPath)
        {
            var name = Path.GetFileName(documentPath);
            foreach (ExcelInterop.Workbook workbook in excel.Workbooks)
            {
                if (string.Equals(workbook.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook;
                }
            }
            return null;
        }

        private static void ActivateWorkbook(ExcelInterop.Application excel, string documentPath)
        {
            var workbook = FindWorkbook(excel, documentPath);
            if (workbook != null)
            {
                workbook.Activate();
            }
        }

        private static void CloseOtherWorkbooks(ExcelInterop.Application excel, string documentPath)
        {
            var name = Path.GetFileName(documentPath);
            for (var i = excel.Workbooks.Count; i >= 1; i--)
            {
                try
                {
                    var workbook = excel.Workbooks[i];
                    if (!string.Equals(workbook.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook.Close(false);
                    }
                }
                catch
                {
                    // 单个工作簿关闭失败不影响后续步骤
                }
            }
        }

        private static string ReadCell(ExcelInterop.Application excel, string documentPath)
        {
            ActivateWorkbook(excel, documentPath);
            var value = Cell(excel.ActiveWorkbook).Value2;
            return value == null ? string.Empty : Convert.ToString(value);
        }

        // -------------------------------------------------------------- Word 辅助

        private static void ActivateDocument(WordInterop.Application word, string documentPath)
        {
            foreach (WordInterop.Document document in word.Documents)
            {
                if (string.Equals(document.FullName, documentPath, StringComparison.OrdinalIgnoreCase))
                {
                    document.Activate();
                    return;
                }
            }
        }

        private static void CloseOtherDocuments(WordInterop.Application word, string documentPath)
        {
            var missing = Type.Missing;
            object doNotSave = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
            for (var i = word.Documents.Count; i >= 1; i--)
            {
                try
                {
                    var document = word.Documents[i];
                    if (!string.Equals(document.FullName, documentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        document.Close(ref doNotSave, ref missing, ref missing);
                    }
                }
                catch
                {
                    // 单个文档关闭失败不影响后续步骤
                }
            }
        }

        private static string ActiveDocumentText(WordInterop.Application word, string documentPath)
        {
            ActivateDocument(word, documentPath);
            return word.ActiveDocument.Content.Text.Trim();
        }

        /// <summary>关闭单个文档且不保存（用于清理临时文档）。</summary>
        private static void CloseWithoutSaving(WordInterop.Document document)
        {
            try
            {
                object doNotSave = WordInterop.WdSaveOptions.wdDoNotSaveChanges;
                var missing = Type.Missing;
                document.Close(ref doNotSave, ref missing, ref missing);
            }
            catch
            {
                // 清理失败不影响测试结论
            }
        }

        /// <summary>按路径找到已打开的文档（用于确认“用户原文档还开着”）。</summary>
        private static WordInterop.Document FindDocumentByPath(WordInterop.Application word, string documentPath)
        {
            for (var i = 1; i <= word.Documents.Count; i++)
            {
                try
                {
                    var document = word.Documents[i];
                    if (string.Equals(document.FullName, documentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return document;
                    }
                }
                catch
                {
                    // 幽灵项，跳过
                }
            }
            return null;
        }

        /// <summary>找到 Word 比较生成的结果文档（标题形如“比较结果2”/“Compare2”）。</summary>
        private static WordInterop.Document FindComparisonDocument(WordInterop.Application word)
        {
            for (var i = 1; i <= word.Documents.Count; i++)
            {
                try
                {
                    var document = word.Documents[i];
                    var name = document.Name ?? string.Empty;
                    if ((name.StartsWith("比较结果", StringComparison.Ordinal) ||
                         name.StartsWith("Compare", StringComparison.OrdinalIgnoreCase)) &&
                        document.Revisions.Count > 0)
                    {
                        return document;
                    }
                }
                catch
                {
                    // 幽灵项，跳过
                }
            }
            return null;
        }

        private static void DumpDocuments(WordInterop.Application word)
        {
            for (var i = 1; i <= word.Documents.Count; i++)
            {
                try
                {
                    var d = word.Documents[i];
                    Console.WriteLine("    doc[" + i + "] name=" + d.Name + " path=" + d.FullName +
                                      " revisions=" + d.Revisions.Count);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("    doc[" + i + "] <error: " + ex.Message + ">");
                }
            }

            Console.WriteLine("    windows=" + word.Windows.Count);
            for (var i = 1; i <= word.Windows.Count; i++)
            {
                try
                {
                    Console.WriteLine("    win[" + i + "] caption=" + word.Windows[i].Caption +
                                      " visible=" + word.Windows[i].Visible);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("    win[" + i + "] <error: " + ex.Message + ">");
                }
            }
        }

        // --------------------------------------------------------------- PPT 辅助

        private static string ActiveSlideTitle(PowerPointInterop.Application ppt, string documentPath)
        {
            for (var i = 1; i <= ppt.Presentations.Count; i++)
            {
                try
                {
                    var presentation = ppt.Presentations[i];
                    if (string.Equals(presentation.FullName, documentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return presentation.Slides[1].Shapes.Title.TextFrame.TextRange.Text;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("    (read title failed: " + ex.Message + ")");
                }
            }
            return null;
        }

        private static void Expect(bool condition, string description)
        {
            if (condition)
            {
                _passed++;
                Console.WriteLine("  [PASS] " + description);
            }
            else
            {
                _failed++;
                Console.WriteLine("  [FAIL] " + description);
            }
        }
    }
}
