using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 版本管理器：在文档“保存之前”把磁盘上的当前状态复制为一个带时间戳的版本快照。
    /// 设计原则：任何失败都不能影响用户的保存操作，所有异常只写日志。
    /// </summary>
    public sealed class VersionManager
    {
        private readonly VersionOptions _options;
        private readonly Logger _logger;

        public VersionManager(VersionOptions options, Logger logger = null)
        {
            _options = options ?? new VersionOptions();
            _logger = logger ?? Logger.Default;
        }

        /// <summary>当前生效的配置。</summary>
        public VersionOptions Options
        {
            get { return _options; }
        }

        /// <summary>
        /// 为指定文件建立版本快照（由 Office 保存事件调用）。
        /// </summary>
        /// <param name="filePath">文档完整路径</param>
        /// <returns>版本文件路径；跳过或失败时返回 null。</returns>
        public string CreateVersion(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return null;
            }
            if (!_options.Enabled)
            {
                return null;
            }

            try
            {
                if (!File.Exists(filePath))
                {
                    _logger.Debug("跳过自动版本：文件尚未保存到磁盘 - " + filePath);
                    return null;
                }

                try
                {
                    return CreateSnapshot(filePath, _options);
                }
                catch (Exception primaryException)
                {
                    // 原文件旁边放不下（只读目录、无写入权限等）时，退回到统一目录再试一次
                    if (_options.StorageMode == VersionStorageMode.NextToFile)
                    {
                        _logger.Warn("在原文件旁创建版本失败，尝试改用统一目录 - " + primaryException.Message);
                        var fallback = _options.Clone();
                        fallback.StorageMode = VersionStorageMode.CentralStore;
                        return CreateSnapshot(filePath, fallback);
                    }
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.Error("创建版本失败 - " + filePath, ex);
                return null;
            }
        }

        /// <summary>获取某文档的版本文件夹路径。</summary>
        public string GetVersionsDirectory(string filePath)
        {
            return VersionPaths.GetVersionsDirectory(filePath, _options);
        }

        /// <summary>列出某文档的全部版本文件（按修改时间从新到旧）。</summary>
        public List<string> ListVersions(string filePath)
        {
            var result = new List<string>();
            try
            {
                var directory = VersionPaths.GetVersionsDirectory(filePath, _options);
                if (!Directory.Exists(directory))
                {
                    return result;
                }
                var name = Path.GetFileNameWithoutExtension(filePath);
                var ext = Path.GetExtension(filePath);
                foreach (var file in Directory.EnumerateFiles(directory))
                {
                    if (IsVersionOf(file, name, ext))
                    {
                        result.Add(file);
                    }
                }
                result.Sort(CompareByRecencyDesc);
            }
            catch (Exception ex)
            {
                _logger.Error("列出版本失败 - " + filePath, ex);
            }
            return result;
        }

        private string CreateSnapshot(string filePath, VersionOptions options)
        {
            var directory = VersionPaths.GetVersionsDirectory(filePath, options);
            Directory.CreateDirectory(directory);

            if (options.SkipIfUnchanged && IsUnchanged(filePath, directory))
            {
                _logger.Info("跳过自动版本：内容与最新版本一致 - " + filePath);
                return null;
            }

            var target = Path.Combine(directory, VersionPaths.BuildVersionFileName(filePath, options, DateTime.Now));
            target = EnsureUniquePath(target);
            CopyFileRobust(filePath, target);
            // 把版本文件的修改时间固定为“快照生成时间”，保证清理旧版本时按生成顺序排序
            File.SetLastWriteTime(target, DateTime.Now);

            _logger.Info("已创建版本 - " + target);

            PruneOldVersions(directory, filePath, options);
            return target;
        }

        private bool IsUnchanged(string filePath, string directory)
        {
            try
            {
                var name = Path.GetFileNameWithoutExtension(filePath);
                var ext = Path.GetExtension(filePath);
                var latest = FindLatest(directory, name, ext);
                if (latest == null)
                {
                    return false;
                }
                var currentHash = ComputeHash(filePath);
                var latestHash = ComputeHash(latest);
                return !string.IsNullOrEmpty(currentHash) && string.Equals(currentHash, latestHash, StringComparison.Ordinal);
            }
            catch (Exception ex)
            {
                _logger.Debug("内容比较失败，按“有变化”处理 - " + ex.Message);
                return false;
            }
        }

        private static string FindLatest(string directory, string name, string ext)
        {
            string best = null;
            var bestTime = DateTime.MinValue;
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                if (!IsVersionOf(file, name, ext))
                {
                    continue;
                }
                var time = File.GetLastWriteTime(file);
                if (time >= bestTime)
                {
                    bestTime = time;
                    best = file;
                }
            }
            return best;
        }

        private static string ComputeHash(string path)
        {
            using (var md5 = MD5.Create())
            // 保存瞬间 Office 仍持有文件句柄，用最宽松的共享模式打开才能读到内容，
            // 否则会误判为“读取失败 → 内容有变化”，导致无修改也重复生成版本。
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                return Convert.ToBase64String(md5.ComputeHash(stream));
            }
        }

        /// <summary>
        /// 用指定文件覆盖目标文件（“恢复版本”用）。
        /// Office 关闭文档后可能仍短暂持有文件句柄，直接 File.Copy 会报“正由另一进程使用”，
        /// 因此内部带重试。
        /// </summary>
        public static void ReplaceFileRobust(string source, string target)
        {
            CopyFileRobust(source, target, true);
        }

        /// <summary>
        /// 等待目标文件的独占句柄释放（Office 关闭文档后仍可能短暂占用几百毫秒）。
        /// 返回 true 表示已可写入，false 表示超时仍被占用。
        /// </summary>
        public static bool WaitForFileUnlocked(string path, int timeoutMilliseconds)
        {
            if (string.IsNullOrEmpty(path))
            {
                return true;
            }
            var watch = System.Diagnostics.Stopwatch.StartNew();
            var limit = Math.Max(0, timeoutMilliseconds);
            while (true)
            {
                if (!File.Exists(path))
                {
                    return true;
                }
                FileStream probe = null;
                try
                {
                    probe = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    return true;
                }
                catch (IOException)
                {
                    // 仍被占用，继续等
                }
                catch (UnauthorizedAccessException)
                {
                    // 只读属性或被拒绝，交给后续覆盖逻辑处理
                    return false;
                }
                finally
                {
                    if (probe != null)
                    {
                        probe.Dispose();
                    }
                }
                if (watch.ElapsedMilliseconds >= limit)
                {
                    return false;
                }
                System.Threading.Thread.Sleep(100);
            }
        }

        /// <summary>
        /// 复制文件。Office 保存过程中可能仍持有原文件句柄，直接 File.Copy 会报“正由另一进程使用”，
        /// 因此带重试，最后一次改用最宽松的共享模式手工复制。
        /// </summary>
        private static void CopyFileRobust(string source, string target, bool overwrite = false)
        {
            const int maxAttempts = 12;
            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    if (attempt < maxAttempts)
                    {
                        File.Copy(source, target, overwrite);
                    }
                    else
                    {
                        using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                        using (var output = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            input.CopyTo(output);
                        }
                    }
                    return;
                }
                catch (IOException)
                {
                    if (attempt >= maxAttempts)
                    {
                        throw;
                    }
                    try
                    {
                        if (File.Exists(target))
                        {
                            File.Delete(target);
                        }
                    }
                    catch
                    {
                        // 清理半成品失败就继续重试
                    }
                    System.Threading.Thread.Sleep(150);
                }
            }
        }

        /// <summary>同一秒内多次保存时，避免文件名冲突。</summary>
        private static string EnsureUniquePath(string path)
        {
            if (!File.Exists(path))
            {
                return path;
            }
            var directory = Path.GetDirectoryName(path);
            var name = Path.GetFileNameWithoutExtension(path);
            var ext = Path.GetExtension(path);
            for (var i = 1; i < 1000; i++)
            {
                var candidate = Path.Combine(directory, string.Format("{0}({1}){2}", name, i, ext));
                if (!File.Exists(candidate))
                {
                    return candidate;
                }
            }
            return Path.Combine(directory, string.Format("{0}({1}){2}", name, Guid.NewGuid().ToString("N").Substring(0, 6), ext));
        }

        private void PruneOldVersions(string directory, string filePath, VersionOptions options)
        {
            var max = options.MaxVersionsToKeep;
            if (max <= 0)
            {
                return;
            }
            var name = Path.GetFileNameWithoutExtension(filePath);
            var ext = Path.GetExtension(filePath);
            var files = new List<string>();
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                if (IsVersionOf(file, name, ext))
                {
                    files.Add(file);
                }
            }
            if (files.Count <= max)
            {
                return;
            }
            files.Sort(CompareByRecencyDesc);
            for (var i = max; i < files.Count; i++)
            {
                try
                {
                    // 旧版本先尝试移入回收站（可恢复），失败再直接删除
                    if (!RecycleBin.Delete(files[i]))
                    {
                        File.Delete(files[i]);
                    }
                    _logger.Info("清理旧版本 - " + files[i]);
                }
                catch (Exception ex)
                {
                    _logger.Error("清理旧版本失败 - " + files[i], ex);
                }
            }
        }

        /// <summary>版本排序：新的在前（依次比较修改时间、创建时间、文件名）。</summary>
        private static int CompareByRecencyDesc(string a, string b)
        {
            var byWriteTime = File.GetLastWriteTime(b).CompareTo(File.GetLastWriteTime(a));
            if (byWriteTime != 0)
            {
                return byWriteTime;
            }
            var byCreationTime = File.GetCreationTime(b).CompareTo(File.GetCreationTime(a));
            if (byCreationTime != 0)
            {
                return byCreationTime;
            }
            return string.CompareOrdinal(b, a);
        }

        /// <summary>判断某文件是否是“原文件名为 name、扩展名为 ext”的文档的版本快照。</summary>
        private static bool IsVersionOf(string candidate, string name, string ext)
        {
            var fileName = Path.GetFileName(candidate);
            if (!string.Equals(Path.GetExtension(fileName), ext, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            var bare = Path.GetFileNameWithoutExtension(fileName);
            return bare.Equals(name, StringComparison.OrdinalIgnoreCase)
                || bare.StartsWith(name + "_", StringComparison.OrdinalIgnoreCase)
                || bare.StartsWith(name + " ", StringComparison.OrdinalIgnoreCase)
                || bare.StartsWith(name + "-", StringComparison.OrdinalIgnoreCase);
        }
    }
}
