using System;

namespace OfficeAutoVersion.Core
{
    /// <summary>
    /// 版本保存位置模式。
    /// </summary>
    public enum VersionStorageMode
    {
        /// <summary>保存在原文件同目录下的“&lt;文件名&gt;_Versions”文件夹中（默认，随文件走，拷贝/同步文件时版本一起带走）。</summary>
        NextToFile = 0,

        /// <summary>集中保存到统一目录（默认：文档\OfficeAutoVersions），并按原路径镜像目录结构，保持原文件夹干净。</summary>
        CentralStore = 1
    }

    /// <summary>
    /// 自动版本插件的配置项。每个 Office 应用各持有一份，按应用键（Word/Excel/PowerPoint）持久化到 JSON。
    /// </summary>
    public sealed class VersionOptions
    {
        /// <summary>是否启用自动版本（总开关）。</summary>
        public bool Enabled { get; set; }

        /// <summary>版本保存位置模式。</summary>
        public VersionStorageMode StorageMode { get; set; }

        /// <summary>版本文件夹名称，支持 {name} 占位符（原文件名，不含扩展名）。</summary>
        public string VersionFolderName { get; set; }

        /// <summary>集中保存模式的根目录；为空时使用“文档\OfficeAutoVersions”。</summary>
        public string CentralRoot { get; set; }

        /// <summary>最多保留的版本数量，0 表示不限制（超出的旧版本移入回收站）。</summary>
        public int MaxVersionsToKeep { get; set; }

        /// <summary>内容与最新版本完全一致（MD5 相同）时跳过，避免重复快照。</summary>
        public bool SkipIfUnchanged { get; set; }

        /// <summary>
        /// 版本文件名格式。支持占位符：
        /// {name}=文件名（不含扩展名），{ext}=扩展名（含点），{ts}=时间戳(yyyyMMdd-HHmmss)，
        /// 也可以直接写 .NET 日期时间格式，如 {yyyyMMdd-HHmmss}、{yyyy年MM月dd日}。
        /// </summary>
        public string FileNamePattern { get; set; }

        /// <summary>是否记录运行日志。</summary>
        public bool LogEnabled { get; set; }

        /// <summary>创建一份默认配置。</summary>
        public VersionOptions()
        {
            Enabled = true;
            StorageMode = VersionStorageMode.NextToFile;
            VersionFolderName = VersionPaths.DefaultVersionFolderName;
            CentralRoot = string.Empty;
            MaxVersionsToKeep = 20;
            SkipIfUnchanged = true;
            FileNamePattern = VersionPaths.DefaultFileNamePattern;
            LogEnabled = true;
        }

        /// <summary>复制一份配置（用于回退到集中目录等场景）。</summary>
        public VersionOptions Clone()
        {
            return new VersionOptions
            {
                Enabled = this.Enabled,
                StorageMode = this.StorageMode,
                VersionFolderName = this.VersionFolderName,
                CentralRoot = this.CentralRoot,
                MaxVersionsToKeep = this.MaxVersionsToKeep,
                SkipIfUnchanged = this.SkipIfUnchanged,
                FileNamePattern = this.FileNamePattern,
                LogEnabled = this.LogEnabled,
            };
        }
    }
}
