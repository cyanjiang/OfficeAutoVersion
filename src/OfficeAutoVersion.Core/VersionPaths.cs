using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 版本文件夹与版本文件名的计算规则（纯函数，便于单元测试）。
    /// </summary>
    public static class VersionPaths
    {
        /// <summary>默认版本文件夹名模板。</summary>
        public const string DefaultVersionFolderName = "{name}_Versions";

        /// <summary>默认版本文件名模板。</summary>
        public const string DefaultFileNamePattern = "{name}_{ts}";

        /// <summary>默认时间戳格式。</summary>
        public const string DefaultTimestampFormat = "yyyyMMdd-HHmmss";

        /// <summary>集中保存模式的默认根目录：文档\OfficeAutoVersions。</summary>
        public static string GetCentralRoot(VersionOptions options)
        {
            if (options != null && !string.IsNullOrWhiteSpace(options.CentralRoot))
            {
                return options.CentralRoot.Trim();
            }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "OfficeAutoVersions");
        }

        /// <summary>计算某个文档的版本文件夹完整路径。</summary>
        public static string GetVersionsDirectory(string filePath, VersionOptions options)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("文件路径不能为空。", "filePath");
            }
            options = options ?? new VersionOptions();

            var directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            var name = Path.GetFileNameWithoutExtension(filePath);

            var folderName = Expand(options.VersionFolderName, name, null, DateTime.MinValue, false);
            if (string.IsNullOrWhiteSpace(folderName))
            {
                folderName = Expand(DefaultVersionFolderName, name, null, DateTime.MinValue, false);
            }
            folderName = SanitizeFileName(folderName);
            if (string.IsNullOrWhiteSpace(folderName))
            {
                folderName = name + "_Versions";
            }

            if (options.StorageMode == VersionStorageMode.CentralStore)
            {
                return Path.Combine(GetCentralRoot(options), MirrorDirectory(directory), folderName);
            }
            return Path.Combine(directory, folderName);
        }

        /// <summary>计算版本文件名（不含目录）。</summary>
        public static string BuildVersionFileName(string filePath, VersionOptions options, DateTime now)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("文件路径不能为空。", "filePath");
            }
            options = options ?? new VersionOptions();

            var name = Path.GetFileNameWithoutExtension(filePath);
            var ext = Path.GetExtension(filePath);

            var pattern = string.IsNullOrWhiteSpace(options.FileNamePattern) ? DefaultFileNamePattern : options.FileNamePattern.Trim();
            var fileName = Expand(pattern, name, ext, now, true);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = Expand(DefaultFileNamePattern, name, ext, now, true);
            }

            fileName = SanitizeFileName(fileName);
            if (!string.IsNullOrEmpty(ext) && !fileName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                fileName += ext;
            }
            return fileName;
        }

        /// <summary>
        /// 展开占位符：{name}=文件名（不含扩展名），{ext}=扩展名（含点），{ts}=时间戳；
        /// 其它 {token} 视为 .NET 日期时间格式（如 {yyyyMMdd-HHmmss}）。文件夹名场景不展开时间。
        /// </summary>
        private static string Expand(string template, string name, string ext, DateTime now, bool allowTimeTokens)
        {
            if (string.IsNullOrEmpty(template))
            {
                return string.Empty;
            }
            return Regex.Replace(template, @"\{([^{}]*)\}", delegate (Match match)
            {
                var token = match.Groups[1].Value;
                if (token.Equals("name", StringComparison.OrdinalIgnoreCase))
                {
                    return name ?? string.Empty;
                }
                if (token.Equals("ext", StringComparison.OrdinalIgnoreCase))
                {
                    return ext ?? string.Empty;
                }
                if (token.Equals("ts", StringComparison.OrdinalIgnoreCase))
                {
                    return allowTimeTokens ? now.ToString(DefaultTimestampFormat) : string.Empty;
                }
                if (!allowTimeTokens)
                {
                    return string.Empty;
                }
                try
                {
                    return now.ToString(token);
                }
                catch
                {
                    return string.Empty;
                }
            });
        }

        /// <summary>把原文件所在目录镜像为集中目录下的相对路径（去掉盘符/UNC 前缀中的非法字符）。</summary>
        private static string MirrorDirectory(string directory)
        {
            var full = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var root = Path.GetPathRoot(full) ?? string.Empty;
            var rest = full.Length > root.Length ? full.Substring(root.Length) : string.Empty;
            var drive = root
                .Replace(":", string.Empty)
                .Replace(Path.DirectorySeparatorChar.ToString(), string.Empty)
                .Replace(Path.AltDirectorySeparatorChar.ToString(), string.Empty);
            var combined = (drive + Path.DirectorySeparatorChar + rest).Trim(Path.DirectorySeparatorChar);
            return SanitizeRelativePath(combined);
        }

        private static string SanitizeRelativePath(string relativePath)
        {
            var parts = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            for (var i = 0; i < parts.Length; i++)
            {
                parts[i] = SanitizeFileName(parts[i]);
                if (string.IsNullOrWhiteSpace(parts[i]))
                {
                    parts[i] = "_";
                }
            }
            return string.Join(Path.DirectorySeparatorChar.ToString(), parts);
        }

        /// <summary>替换文件名中的非法字符，并去掉首尾空格和结尾的小数点。</summary>
        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return string.Empty;
            }
            var invalid = Path.GetInvalidFileNameChars();
            var builder = new StringBuilder(name.Length);
            foreach (var ch in name)
            {
                builder.Append(Array.IndexOf(invalid, ch) >= 0 ? '_' : ch);
            }
            return builder.ToString().Trim().TrimEnd('.', ' ');
        }
    }
}
