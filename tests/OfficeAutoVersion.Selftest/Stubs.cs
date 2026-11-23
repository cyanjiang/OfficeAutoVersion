using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

// 说明：这些接口只是为了在“没有安装 Office PIA”的开发机上编译并自测 Common 层代码，
// 成员签名与 Microsoft.Office.Core 中的真实接口保持一致（仅包含本插件用到的部分）。
// 真实加载项工程引用的是 Office PIA（office.dll），不会使用这里的定义。
namespace Microsoft.Office.Core
{
    public enum MsoTriState
    {
        msoTrue = -1,
        msoFalse = 0,
        msoCTrue = 1,
        msoTriStateToggle = -3,
        msoTriStateMixed = -2,
    }

    [ComVisible(true)]
    public interface IRibbonExtensibility
    {
        string GetCustomUI(string ribbonId);
    }

    [ComVisible(true)]
    public interface IRibbonControl
    {
        string Id { get; }
        string Tag { get; }
    }

    [ComVisible(true)]
    public interface IRibbonUI
    {
        void Invalidate();
        void InvalidateControl(string controlId);
        void InvalidateControlMso(string controlId);
    }
}

namespace OfficeAutoVersion.Selftest
{
    using Microsoft.Office.Core;
    using OfficeAutoVersion.Common;

    /// <summary>记录调用的宿主动作替身：验证窗体是否把“比较/恢复”正确转交给 Office 自带能力。</summary>
    internal sealed class FakeHostActions : IVersionHostActions
    {
        public string ComparedCurrentPath;
        public string ComparedPath;
        public string RestoredCurrentPath;
        public string RestoredPath;
        public string RestoreSnapshotToReturn;

        public string CompareButtonText
        {
            get { return "比较差异（测试）"; }
        }

        public string CompareHint
        {
            get { return "测试用提示"; }
        }

        public void CompareWithCurrent(string currentFilePath, string versionFilePath)
        {
            this.ComparedCurrentPath = currentFilePath;
            this.ComparedPath = versionFilePath;
        }

        public string RestoreVersion(string currentFilePath, string versionFilePath)
        {
            this.RestoredCurrentPath = currentFilePath;
            this.RestoredPath = versionFilePath;
            return this.RestoreSnapshotToReturn;
        }

        /// <summary>测试用页映射：可在用例里塞入“文本 → 页码”，验证摘要里的页码标注。</summary>
        public System.Collections.Generic.Dictionary<string, int> PageMap =
            new System.Collections.Generic.Dictionary<string, int>(StringComparer.Ordinal);

        public System.Collections.Generic.Dictionary<string, int> GetTextPageMap(string documentFilePath)
        {
            return PageMap.Count == 0 ? null : PageMap;
        }
    }

    internal sealed class FakeRibbonControl : IRibbonControl
    {
        public FakeRibbonControl(string id)
        {
            this.Id = id;
        }

        public string Id { get; private set; }

        public string Tag
        {
            get { return null; }
        }
    }

    internal sealed class FakeRibbonUi : IRibbonUI
    {
        public readonly List<string> InvalidatedControls = new List<string>();
        public int InvalidateAllCount;

        public void Invalidate()
        {
            this.InvalidateAllCount++;
        }

        public void InvalidateControl(string controlId)
        {
            this.InvalidatedControls.Add(controlId);
        }

        public void InvalidateControlMso(string controlId)
        {
            this.InvalidatedControls.Add(controlId);
        }
    }
}
