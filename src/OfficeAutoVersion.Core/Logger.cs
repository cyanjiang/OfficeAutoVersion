using System;
using System.IO;
using System.Text;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 极简的运行日志。日志失败一律静默忽略，绝不影响主流程。
    /// </summary>
    public sealed class Logger
    {
        private static readonly Lazy<Logger> DefaultInstance = new Lazy<Logger>(delegate ()
        {
            return new Logger(Path.Combine(OptionsStore.AppDataFolder, "logs", "autoversion.log"));
        });

        /// <summary>进程级默认日志实例（%APPDATA%\OfficeAutoVersion\logs\autoversion.log）。</summary>
        public static Logger Default
        {
            get { return DefaultInstance.Value; }
        }

        private const long MaxFileSize = 2 * 1024 * 1024;

        private readonly object _gate = new object();
        private readonly string _path;

        public Logger(string path)
        {
            _path = path;
        }

        /// <summary>日志文件完整路径。</summary>
        public string LogPath
        {
            get { return _path; }
        }

        public void Info(string message)
        {
            Write("INFO", message, null);
        }

        public void Warn(string message)
        {
            Write("WARN", message, null);
        }

        public void Debug(string message)
        {
            Write("DEBUG", message, null);
        }

        public void Error(string message, Exception exception = null)
        {
            Write("ERROR", message, exception);
        }

        private void Write(string level, string message, Exception exception)
        {
            try
            {
                lock (_gate)
                {
                    var directory = Path.GetDirectoryName(_path);
                    if (!string.IsNullOrEmpty(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                    if (File.Exists(_path) && new FileInfo(_path).Length > MaxFileSize)
                    {
                        var backup = _path + ".1";
                        try
                        {
                            if (File.Exists(backup))
                            {
                                File.Delete(backup);
                            }
                            File.Move(_path, backup);
                        }
                        catch
                        {
                            // 轮转失败就继续写当前文件
                        }
                    }
                    var line = string.Format("{0:yyyy-MM-dd HH:mm:ss.fff} [{1}] {2}", DateTime.Now, level, message ?? string.Empty);
                    if (exception != null)
                    {
                        line += Environment.NewLine + "    " + exception;
                    }
                    File.AppendAllText(_path, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // 日志写失败一律忽略
            }
        }
    }
}
