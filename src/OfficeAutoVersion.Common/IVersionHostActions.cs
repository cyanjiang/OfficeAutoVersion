using System;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// 宿主应用（Word/Excel/PowerPoint）为“版本历史”窗口提供的两个动作。
    /// 比较一律调用 Office 自带的比较能力，各应用方式不同，因此由加载项分别实现。
    /// </summary>
    public interface IVersionHostActions
    {
        /// <summary>比较按钮上显示的文字（不同应用的比较方式不同，如实标注）。</summary>
        string CompareButtonText { get; }

        /// <summary>比较按钮的提示文字。</summary>
        string CompareHint { get; }

        /// <summary>
        /// 用 Office 自带功能比较“当前文档”与指定版本文件。
        /// </summary>
        /// <param name="currentFilePath">
        /// 当前文档的完整路径。必须显式传入，不能靠应用自己的“活动文档”：
        /// 版本历史窗口是模态窗口，打开期间 Excel 的 ActiveWorkbook / Word 的 ActiveDocument
        /// 都可能返回 null（“当前没有打开的工作簿”就是这么来的）。
        /// </param>
        /// <param name="versionFilePath">要对比的版本快照。</param>
        void CompareWithCurrent(string currentFilePath, string versionFilePath);

        /// <summary>
        /// 把当前文档恢复到指定版本。
        /// </summary>
        /// <param name="currentFilePath">当前文档的完整路径（理由同 <see cref="CompareWithCurrent"/>）。</param>
        /// <param name="versionFilePath">要恢复到的版本快照。</param>
        /// <returns>恢复前为“当时状态”自动建立的快照路径；未建立时返回 null。</returns>
        string RestoreVersion(string currentFilePath, string versionFilePath);

        /// <summary>
        /// 为差异摘要提供“段落文本 → 页码”的映射，用来把改动定位到具体页。
        /// 页码是排版结果，文件里并不保存，只能由宿主应用按当前排版算出来；
        /// 不支持（Excel/PowerPoint 没有页的概念）或文档未打开时返回 null。
        /// 键必须是 <see cref="OfficeAutoVersion.Core.ContentDiff.NormalizeKey"/> 规范化后的文本。
        /// </summary>
        System.Collections.Generic.Dictionary<string, int> GetTextPageMap(string documentFilePath);
    }
}
