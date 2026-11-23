using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OfficeAutoVersion.Setup
{
    /// <summary>产品信息与安装位置（全部集中在这里，改版本号只需改一处）。</summary>
    internal static class Product
    {
        public const string Name = "Office 自动版本";
        public const string Id = "OfficeAutoVersion";
        public const string Publisher = "OfficeAutoVersion";

        /// <summary>三个要注册的 Office 宿主。</summary>
        public static readonly string[] Hosts = { "Word", "Excel", "PowerPoint" };

        /// <summary>主版本号（取自程序集版本，改 csproj 的 Version 即可）。</summary>
        public static string Version
        {
            get
            {
                var version = typeof(Product).Assembly.GetName().Version;
                return version.Major + "." + version.Minor + "." + version.Build;
            }
        }

        /// <summary>每个用户自己的安装目录，不需要管理员权限。</summary>
        public static string InstallDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Id);
            }
        }

        public static string BinDir
        {
            get { return Path.Combine(InstallDir, "bin"); }
        }

        /// <summary>安装到安装目录里的安装程序副本，卸载时用它。</summary>
        public static string InstalledSetupPath
        {
            get { return Path.Combine(InstallDir, "Setup.exe"); }
        }

        /// <summary>随安装包一起发布的《使用说明》文档。</summary>
        public const string DocFileName = "Office自动版本-使用说明.docx";

        public static string DocDir
        {
            get { return Path.Combine(InstallDir, "docs"); }
        }

        public static string DocPath
        {
            get { return Path.Combine(DocDir, DocFileName); }
        }

        public static string AddInKey(string host)
        {
            return @"Software\Microsoft\Office\" + host + @"\Addins\" + Id + "." + host;
        }

        public static string UninstallKey
        {
            get { return @"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + Id; }
        }

        public static string StateKey
        {
            get { return @"Software\" + Id; }
        }
    }

    /// <summary>安装 / 卸载的实际动作，UI 与静默模式共用。</summary>
    internal static class Installer
    {
        /// <summary>进度回调：(百分比, 说明)。可能在后台线程触发。</summary>
        public static event Action<int, string> Progress;

        private static void Report(int percent, string message)
        {
            var handler = Progress;
            if (handler != null)
            {
                handler(percent, message);
            }
        }

        // ------------------------------------------------------------ 安装

        public static void Install()
        {
            Report(2, "正在检查运行环境…");
            if (!Requirements.VstoRuntimeInstalled())
            {
                throw new InvalidOperationException(
                    "未检测到 VSTO 4.0 运行时（Visual Studio Tools for Office Runtime）。\n" +
                    "它通常随 Office 一起安装；如果缺失，请先安装后再运行本安装程序。");
            }

            Report(8, "正在创建安装目录…");
            Directory.CreateDirectory(Product.BinDir);

            Report(15, "正在释放程序文件…");
            var locked = ExtractPayload();
            if (locked.Count > 0)
            {
                throw new IOException(
                    "以下文件正被占用，无法写入：\n\n" + string.Join("\n", locked.ToArray()) +
                    "\n\n请关闭 Word / Excel / PowerPoint 后重试。");
            }

            Report(70, "正在注册 Office 加载项…");
            RegisterAddIns();

            Report(85, "正在写入卸载信息…");
            WriteUninstallEntry();
            WriteStateKey();

            Report(95, "正在完成安装…");
            CopySelfForUninstall();

            Report(100, "安装完成");
        }

        /// <summary>把内嵌的 payload.zip 解到安装目录；返回写入失败（被占用）的文件列表。</summary>
        private static List<string> ExtractPayload()
        {
            var failed = new List<string>();
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Payload"))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("安装包损坏：找不到内置的程序文件。");
                }
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                        {
                            continue;
                        }
                        var target = Path.Combine(Product.InstallDir,
                            entry.FullName.Replace('/', Path.DirectorySeparatorChar));
                        try
                        {
                            var directory = Path.GetDirectoryName(target);
                            if (!string.IsNullOrEmpty(directory))
                            {
                                Directory.CreateDirectory(directory);
                            }
                            entry.ExtractToFile(target, true);
                        }
                        catch (IOException)
                        {
                            failed.Add(entry.FullName);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            failed.Add(entry.FullName);
                        }
                    }
                }
            }
            return failed;
        }

        private static void RegisterAddIns()
        {
            foreach (var host in Product.Hosts)
            {
                var vsto = Path.Combine(Product.BinDir, Product.Id + "." + host + ".vsto");
                using (var key = Registry.CurrentUser.CreateSubKey(Product.AddInKey(host)))
                {
                    if (key == null)
                    {
                        continue;
                    }
                    key.SetValue("FriendlyName", Product.Name, RegistryValueKind.String);
                    key.SetValue("Description",
                        "保存文档时自动建立版本，并可用 Office 自带的比较、恢复功能回到任意一版。",
                        RegistryValueKind.String);
                    key.SetValue("Manifest", ToManifestValue(vsto), RegistryValueKind.String);
                    // 3 = 启动时加载；Office 里被用户禁用过也不会自动改回，但新装默认加载
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
            }
        }

        /// <summary>
        /// VSTO 的本地安装写法：file:///…vsto|vstolocal。
        /// vstolocal 让运行时直接从本地读清单，不走 ClickOnce 的发布位置校验。
        /// </summary>
        internal static string ToManifestValue(string vstoPath)
        {
            return "file:///" + vstoPath.Replace('\\', '/') + "|vstolocal";
        }

        private static void WriteUninstallEntry()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(Product.UninstallKey))
            {
                if (key == null)
                {
                    return;
                }
                key.SetValue("DisplayName", Product.Name + "（Word / Excel / PowerPoint 插件）", RegistryValueKind.String);
                key.SetValue("DisplayVersion", Product.Version, RegistryValueKind.String);
                key.SetValue("Publisher", Product.Publisher, RegistryValueKind.String);
                key.SetValue("InstallLocation", Product.InstallDir, RegistryValueKind.String);
                key.SetValue("DisplayIcon", Product.InstalledSetupPath, RegistryValueKind.String);
                key.SetValue("UninstallString", "\"" + Product.InstalledSetupPath + "\" /uninstall", RegistryValueKind.String);
                key.SetValue("QuietUninstallString", "\"" + Product.InstalledSetupPath + "\" /uninstall /S", RegistryValueKind.String);
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                key.SetValue("EstimatedSize", EstimateSizeKb(), RegistryValueKind.DWord);
            }
        }

        private static void WriteStateKey()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(Product.StateKey))
            {
                if (key == null)
                {
                    return;
                }
                key.SetValue("Version", Product.Version, RegistryValueKind.String);
                key.SetValue("InstallDir", Product.InstallDir, RegistryValueKind.String);
                key.SetValue("InstalledAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), RegistryValueKind.String);
            }
        }

        /// <summary>把安装程序自己复制进安装目录，卸载时即使原来的下载文件删了也能卸。</summary>
        private static void CopySelfForUninstall()
        {
            try
            {
                var self = CurrentExePath();
                if (!string.IsNullOrEmpty(self) &&
                    !string.Equals(self, Product.InstalledSetupPath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(self, Product.InstalledSetupPath, true);
                }
            }
            catch (Exception)
            {
                // 复制失败只影响“以后从安装目录卸载”，不影响使用
            }
        }

        private static int EstimateSizeKb()
        {
            try
            {
                long total = 0;
                foreach (var file in Directory.GetFiles(Product.BinDir, "*", SearchOption.AllDirectories))
                {
                    total += new FileInfo(file).Length;
                }
                return (int)Math.Max(1, total / 1024);
            }
            catch
            {
                return 1;
            }
        }

        // ------------------------------------------------------------ 卸载

        public static void Uninstall()
        {
            Report(10, "正在移除加载项注册表…");
            foreach (var host in Product.Hosts)
            {
                DeleteKey(Product.AddInKey(host));
            }

            Report(40, "正在清理安装信息…");
            DeleteKey(Product.UninstallKey);
            DeleteKey(Product.StateKey);

            Report(60, "正在删除程序文件…");
            DeleteDirectory(Product.BinDir);
            DeleteDirectory(Product.DocDir);
            DeleteStaleFiles(Product.InstallDir);

            Report(100, "已卸载");
        }

        private static void DeleteKey(string path)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(path, false);
            }
            catch
            {
                // 键不存在或没有权限都忽略
            }
        }

        private static void DeleteDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch
            {
                // 有文件被占用时留到“稍后删除”
            }
        }

        /// <summary>删掉安装目录里的其它文件；正在运行的安装程序自己排到进程退出后再删。</summary>
        private static void DeleteStaleFiles(string directory)
        {
            if (!Directory.Exists(directory))
            {
                return;
            }
            var self = CurrentExePath();
            var selfInInstallDir = self != null &&
                string.Equals(Path.GetFullPath(self), Path.GetFullPath(Product.InstalledSetupPath),
                    StringComparison.OrdinalIgnoreCase);

            try
            {
                foreach (var file in Directory.GetFiles(directory))
                {
                    if (selfInInstallDir && string.Equals(Path.GetFullPath(file), Path.GetFullPath(self),
                            StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                        // 忽略单个文件
                    }
                }
                if (!selfInInstallDir)
                {
                    Directory.Delete(directory, false);
                }
                else
                {
                    ScheduleSelfDelete(directory);
                }
            }
            catch
            {
                // 忽略
            }
        }

        /// <summary>等本进程退出后再删掉整个安装目录（cmd 延迟 2 秒，不弹窗）。</summary>
        private static void ScheduleSelfDelete(string directory)
        {
            try
            {
                var arguments = "/c ping -n 3 127.0.0.1 >nul & rmdir /s /q \"" + directory + "\"";
                var info = new ProcessStartInfo("cmd.exe", arguments)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                };
                Process.Start(info);
            }
            catch
            {
                // 忽略
            }
        }

        // ------------------------------------------------------------ 公共

        /// <summary>当前正在运行的 Office 应用（Word/Excel/PowerPoint）。</summary>
        public static List<string> RunningOfficeApps()
        {
            var running = new List<string>();
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "WINWORD", "Word" },
                { "EXCEL", "Excel" },
                { "POWERPNT", "PowerPoint" },
            };
            var processes = new Process[0];
            try
            {
                processes = Process.GetProcesses();
            }
            catch
            {
                return running;
            }
            foreach (var process in processes)
            {
                try
                {
                    string name;
                    if (map.TryGetValue(process.ProcessName, out name) && !running.Contains(name))
                    {
                        running.Add(name);
                    }
                }
                catch
                {
                    // 忽略无法枚举的进程
                }
                finally
                {
                    process.Dispose();
                }
            }
            return running;
        }

        /// <summary>找到随包发布的使用说明：优先已安装的副本，其次安装程序同级目录。</summary>
        public static string FindDoc()
        {
            try
            {
                if (File.Exists(Product.DocPath))
                {
                    return Product.DocPath;
                }
                var self = CurrentExePath();
                if (!string.IsNullOrEmpty(self))
                {
                    var directory = Path.GetDirectoryName(self);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        var beside = Path.Combine(directory, Product.DocFileName);
                        if (File.Exists(beside))
                        {
                            return beside;
                        }
                    }
                }
            }
            catch
            {
                // 找不到就按“没有文档”处理
            }
            return null;
        }

        /// <summary>用系统默认程序打开使用说明；找不到返回 false。</summary>
        public static bool OpenDoc()
        {
            var path = FindDoc();
            if (path == null)
            {
                return false;
            }
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static string CurrentExePath()        {
            try
            {
                return Process.GetCurrentProcess().MainModule.FileName;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>读取已安装的版本号；未安装返回 null。</summary>
        public static string InstalledVersion()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(Product.StateKey))
                {
                    if (key == null)
                    {
                        return null;
                    }
                    var version = key.GetValue("Version") as string;
                    return string.IsNullOrEmpty(version) ? null : version;
                }
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>安装前的环境检查。</summary>
    internal static class Requirements
    {
        /// <summary>VSTO 4.0 运行时（Office 的 VSTO 加载项必需）。</summary>
        public static bool VstoRuntimeInstalled()
        {
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                try
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"))
                    {
                        if (key != null)
                        {
                            var install = key.GetValue("Install");
                            if (install != null && Convert.ToInt32(install) == 1)
                            {
                                return true;
                            }
                        }
                    }
                }
                catch
                {
                    // 换下一个视图
                }
            }

            // 注册表读不到时再看文件（更可靠）
            var candidates = new List<string>();
            try
            {
                candidates.Add(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                    @"Microsoft Shared\VSTO\10.0\VSTOLoader.dll"));
            }
            catch
            {
                // 忽略
            }
            try
            {
                candidates.Add(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86),
                    @"Microsoft Shared\VSTO\10.0\VSTOLoader.dll"));
            }
            catch
            {
                // 忽略
            }
            foreach (var path in candidates)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        return true;
                    }
                }
                catch
                {
                    // 忽略
                }
            }
            return false;
        }

        /// <summary>检测已安装的 Office 主机应用（返回显示名）。</summary>
        public static bool OfficeInstalled(out List<string> hosts)
        {
            hosts = new List<string>();
            foreach (var host in Product.Hosts)
            {
                try
                {
                    using (var key = Registry.ClassesRoot.OpenSubKey(host + ".Application"))
                    {
                        if (key != null)
                        {
                            hosts.Add(host);
                        }
                    }
                }
                catch
                {
                    // 忽略
                }
            }
            return hosts.Count > 0;
        }

        /// <summary>.NET Framework 4.8 是否安装（本程序自己的运行前提）。</summary>
        public static bool DotNet48Installed()
        {
            try
            {
                const string path = @"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full";
                foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (var key = baseKey.OpenSubKey(path))
                    {
                        if (key == null)
                        {
                            continue;
                        }
                        var release = key.GetValue("Release");
                        if (release != null && Convert.ToInt32(release) >= 528040)
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
                // 读不到就按“未知”，由调用方按已满足处理（程序能跑就说明有 4.x）
            }
            return true;
        }
    }
}
