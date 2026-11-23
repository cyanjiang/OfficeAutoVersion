using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

// 说明：本文件中的类型只是为了在“没有安装 VSTO/Office PIA”的开发机上编译并自测
// 三个加载项工程（Word/Excel/PowerPoint）的 ThisAddIn 代码。
// 成员签名与真实的 VSTO 运行时、Office 主互操作程序集保持一致（仅包含用到的部分），
// 事件委托的签名均按当前版本 PIA 的官方文档核对过：
//   Word        ApplicationEvents4_DocumentBeforeSaveEventHandler(Document Doc, ref bool SaveAsUI)
//   Excel       AppEvents_WorkbookBeforeSaveEventHandler(Workbook Wb, bool SaveAsUI, out bool Cancel)
//   PowerPoint  EApplication_PresentationBeforeSaveEventHandler(Presentation Pres, out bool Cancel)
// 真实加载项工程引用的是 Office PIA 与 VSTO 运行时程序集，不会使用这里的定义。

namespace Microsoft.VisualStudio.Tools.Applications.Runtime
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public sealed class StartupObjectAttribute : Attribute
    {
        public StartupObjectAttribute(int index)
        {
            this.Index = index;
        }

        public int Index { get; private set; }
    }
}

namespace Microsoft.Office.Tools.Ribbon
{
    public sealed class RibbonFactory
    {
    }

    public abstract class RibbonCollectionBase
    {
        protected RibbonCollectionBase(RibbonFactory factory)
        {
        }
    }
}

namespace Microsoft.Office.Tools
{
    public sealed class Factory
    {
        internal Ribbon.RibbonFactory GetRibbonFactory()
        {
            return null;
        }
    }

    public sealed class StubDataHost
    {
        public bool IsCacheInitialized
        {
            get { return false; }
        }

        public void FillCachedData(object item)
        {
        }

        public void StartCaching(object item, string memberName)
        {
        }

        public void StopCaching(object item, string memberName)
        {
        }

        // 与真实 VSTO 运行时 ICachedDataProvider 一致：宿主对象 + 成员名
        public bool IsCached(object hostObject, string memberName)
        {
            return false;
        }

        public bool NeedsFill(object hostObject, string memberName)
        {
            return false;
        }
    }

    public abstract class AddInBase
    {
        private readonly StubDataHost _dataHost = new StubDataHost();

        protected AddInBase(Factory factory, IServiceProvider serviceProvider, string primaryCookie, string identifier)
        {
        }

        public event EventHandler Startup;

        public event EventHandler Shutdown;

        protected StubDataHost DataHost
        {
            get { return _dataHost; }
        }

        protected T GetHostItem<T>(Type type, string cookie)
        {
            return default(T);
        }

        protected virtual void Initialize()
        {
        }

        protected virtual void FinishInitialization()
        {
            // 由派生类（ThisAddIn.Designer.cs）重写，其中调用 InternalStartup 与 OnStartup
            this.OnStartup();
        }

        protected virtual void InitializeDataBindings()
        {
            // 由派生类（ThisAddIn.Designer.cs）重写
        }

        protected virtual void BeginInit()
        {
        }

        protected virtual void EndInit()
        {
        }

        protected virtual void OnStartup()
        {
            var handler = this.Startup;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        protected virtual void OnShutdown()
        {
            var handler = this.Shutdown;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        protected virtual Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return null;
        }
    }
}

namespace Microsoft.Office.Interop.Word
{
    public enum WdSaveOptions
    {
        wdDoNotSaveChanges = 0,
        wdSaveChanges = -1,
        wdPromptToSaveChanges = -2,
    }

    public enum WdCompareDestination
    {
        wdCompareDestinationOriginal = 0,
        wdCompareDestinationRevised = 1,
        wdCompareDestinationNew = 2,
    }

    public enum WdGranularity
    {
        wdGranularityCharLevel = 0,
        wdGranularityWordLevel = 1,
    }

    /// <summary>Information 的参数（真实 PIA 里是带参数的属性，C# 用索引器语法访问）。</summary>
    public enum WdInformation
    {
        wdActiveEndPageNumber = 3,
    }

    /// <summary>桩：模拟 Range.Information[WdInformation] 这种参数化属性的索引访问。</summary>
    public sealed class InformationStub
    {
        public object this[WdInformation index]
        {
            get { return 3; } // 桩里一律报第 3 页，便于断言“页码确实被标注”
        }
    }

    public sealed class Range
    {
        public string Text
        {
            get { return null; }
        }

        public InformationStub Information
        {
            get { return new InformationStub(); }
        }
    }

    public sealed class Paragraph
    {
        public Range Range
        {
            get { return null; }
        }
    }

    public sealed class Paragraphs
    {
        public int Count
        {
            get { return 0; }
        }

        public Paragraph this[int index]
        {
            get { return null; }
        }
    }

    public sealed class Document
    {
        public string FullName
        {
            get { return null; }
        }

        public Paragraphs Paragraphs
        {
            get { return null; }
        }

        public bool Saved
        {
            get { return true; }
            set { }
        }

        public void Save()
        {
        }

        public void Activate()
        {
        }

        public void Close(ref object saveChanges, ref object originalFormat, ref object routeDocument)
        {
        }
    }

    public sealed class Documents
    {
        public int Count
        {
            get { return 0; }
        }

        public Document this[int index]
        {
            get { return null; }
        }

        public Document Open(ref object fileName, ref object confirmConversions, ref object readOnly,
            ref object addToRecentFiles, ref object passwordDocument, ref object passwordTemplate,
            ref object revert, ref object writePasswordDocument, ref object writePasswordTemplate,
            ref object format, ref object encoding, ref object visible, ref object openAndRepair,
            ref object documentDirection, ref object noEncodingDialog, ref object xmlTransform)
        {
            return null;
        }
    }

    public delegate void ApplicationEvents4_DocumentBeforeSaveEventHandler(Document doc, ref bool saveAsUi, ref bool cancel);

    public sealed class Application
    {
        public event ApplicationEvents4_DocumentBeforeSaveEventHandler DocumentBeforeSave;

        public Document ActiveDocument
        {
            get { return null; }
        }

        public Documents Documents
        {
            get { return new Documents(); }
        }

        public Document CompareDocuments(Document originalDocument, Document revisedDocument,
            WdCompareDestination destination, WdGranularity granularity,
            bool compareFormatting, bool compareCaseChanges, bool compareWhitespace, bool compareTables,
            bool compareHeaders, bool compareFootnotes, bool compareTextboxes, bool compareFields,
            bool compareComments, bool compareMoves, string revisedAuthor, bool ignoreAllComparisonWarnings)
        {
            return null;
        }
    }
}

namespace Microsoft.Office.Interop.Excel
{
    public enum XlArrangeStyle
    {
        xlArrangeStyleTiled = 1,
        xlArrangeStyleCascade = 7,
        xlArrangeStyleVertical = 2,
        xlArrangeStyleHorizontal = 3,
    }

    public sealed class Window
    {
        public void Activate()
        {
        }
    }

    public sealed class Windows
    {
        public bool SyncScrollingSideBySide { get; set; }

        public void CompareSideBySideWith(string windowName)
        {
        }

        public void Arrange(XlArrangeStyle arrangeStyle, object activeWorkbook, object syncHorizontal, object syncVertical)
        {
        }
    }

    public sealed class Workbook
    {
        public string FullName
        {
            get { return null; }
        }

        public string Name
        {
            get { return null; }
        }

        public bool Saved
        {
            get { return true; }
            set { }
        }

        public void Save()
        {
        }

        public void Activate()
        {
        }

        public void Close(object saveChanges = null, object fileName = null, object routeWorkbook = null)
        {
        }
    }

    public sealed class Workbooks
    {
        public int Count
        {
            get { return 0; }
        }

        public Workbook this[int index]
        {
            get { return null; }
        }

        public Workbook Open(string fileName, object updateLinks = null, object readOnly = null,
            object format = null, object password = null, object writeResPassword = null,
            object ignoreReadOnlyRecommended = null, object origin = null, object delimiter = null,
            object editable = null, object notify = null, object converter = null, object addToMru = null,
            object local = null, object corruptLoad = null)
        {
            return null;
        }
    }

    public delegate void AppEvents_WorkbookBeforeSaveEventHandler(Workbook wb, bool saveAsUi, ref bool cancel);

    public sealed class Application
    {
        public event AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSave;

        public Workbook ActiveWorkbook
        {
            get { return null; }
        }

        public Workbooks Workbooks
        {
            get { return new Workbooks(); }
        }

        public Window ActiveWindow
        {
            get { return new Window(); }
        }

        public Windows Windows
        {
            get { return new Windows(); }
        }
    }
}

namespace Microsoft.Office.Interop.PowerPoint
{
    public sealed class DocumentWindow
    {
        public void Activate()
        {
        }
    }

    public sealed class DocumentWindows
    {
        public DocumentWindow this[int index]
        {
            get { return new DocumentWindow(); }
        }
    }

    public sealed class Presentation
    {
        public string FullName
        {
            get { return null; }
        }

        public Microsoft.Office.Core.MsoTriState Saved
        {
            get { return Microsoft.Office.Core.MsoTriState.msoTrue; }
        }

        public DocumentWindows Windows
        {
            get { return new DocumentWindows(); }
        }

        public void Save()
        {
        }

        public void Close()
        {
        }

        public void Merge(string path)
        {
        }
    }

    public sealed class Presentations
    {
        public int Count
        {
            get { return 0; }
        }

        public Presentation this[int index]
        {
            get { return null; }
        }

        public Presentation Open(string fileName, Microsoft.Office.Core.MsoTriState readOnly,
            Microsoft.Office.Core.MsoTriState untitled, Microsoft.Office.Core.MsoTriState withWindow)
        {
            return null;
        }
    }

    public delegate void EApplication_PresentationBeforeSaveEventHandler(Presentation pres, ref bool cancel);

    public sealed class Application
    {
        public event EApplication_PresentationBeforeSaveEventHandler PresentationBeforeSave;

        public Presentation ActivePresentation
        {
            get { return null; }
        }

        public Presentations Presentations
        {
            get { return new Presentations(); }
        }
    }
}
