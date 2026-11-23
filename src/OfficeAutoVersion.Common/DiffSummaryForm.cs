using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeAutoVersion.Core;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// 差异摘要窗口：展示“当前文档 vs 所选版本”的内容差异（统计 + 明细）。
    /// 数据由 VersionDiffer 纯文件解析得到，不依赖 Office 对象模型。
    /// </summary>
    public sealed class DiffSummaryForm : Form
    {
        private readonly DiffSummary _summary;
        private readonly string _currentPath;
        private readonly string _versionPath;
        private readonly Func<string, int> _pageOf;

        private TextBox _textBox;

        public DiffSummaryForm(DiffSummary summary, string currentPath, string versionPath)
            : this(summary, currentPath, versionPath, null)
        {
        }

        /// <param name="pageOf">
        /// 可选：文本 → 页码。由宿主应用按当前排版算出来（docx 里不存页码）。
        /// 为 null 或返回 0 时不标注页码。
        /// </param>
        public DiffSummaryForm(DiffSummary summary, string currentPath, string versionPath,
            Func<string, int> pageOf)
        {
            _summary = summary;
            _currentPath = currentPath;
            _versionPath = versionPath;
            _pageOf = pageOf;
            BuildUi();
        }

        private void BuildUi()
        {
            UiTheme.PrepareForm(this, "差异摘要 - " + Path.GetFileName(_currentPath));
            FormBorderStyle = FormBorderStyle.Sizable;
            ClientSize = new Size(980, 680);
            MinimumSize = new Size(820, 520);
            DoubleBuffered = true;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = UiTheme.WindowBack,
                Padding = new Padding(18, 14, 18, 14),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            root.Controls.Add(BuildHeaderCard(), 0, 0);
            root.Controls.Add(BuildBodyCard(), 0, 1);
            root.Controls.Add(BuildFooter(), 0, 2);

            Controls.Add(root);
        }

        private Control BuildHeaderCard()
        {
            var card = new UiTheme.CardPanel();
            var content = new TableLayoutPanel
            {
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                BackColor = Color.Transparent,
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var title = UiTheme.CreateLabel("差异摘要", UiTheme.Title(), UiTheme.TextPrimary);
            title.Margin = new Padding(0, 0, 0, 2);
            content.Controls.Add(title);

            var subtitle = UiTheme.CreateLabel(
                "当前文档：" + Path.GetFileName(_currentPath) + "    对比版本：" + Path.GetFileName(_versionPath),
                UiTheme.Body(), UiTheme.TextMuted);
            subtitle.Margin = new Padding(0, 4, 0, 6);
            content.Controls.Add(subtitle);

            var stats = UiTheme.CreateLabel(_summary == null
                    ? "（没有可展示的差异数据）"
                    : string.Format("新增 {0} 处    删除 {1} 处    修改 {2} 处    共 {3} 处差异",
                        _summary.AddedCount, _summary.RemovedCount, _summary.ChangedCount, _summary.TotalCount),
                UiTheme.BodyBold(), UiTheme.AccentText);
            content.Controls.Add(stats);

            content.RowCount = content.Controls.Count;
            for (var i = 0; i < content.RowCount; i++)
            {
                content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            card.AddContent(content);
            return card;
        }

        private Control BuildBodyCard()
        {
            var card = new UiTheme.CardPanel
            {
                AutoSize = false,
                Padding = new Padding(8, 8, 8, 8),
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 10),
            };

            _textBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                BorderStyle = BorderStyle.None,
                BackColor = UiTheme.InputBack,
                ForeColor = UiTheme.TextPrimary,
                Font = UiTheme.Mono(),
                Text = _summary != null ? _summary.ToText(_pageOf, false) : "",
            };
            NativeTheme.ApplyToControl(_textBox);
            // 只读多行框有时会整段选中（蓝色高亮），显示前取消选区
            _textBox.SelectionStart = 0;
            _textBox.SelectionLength = 0;
            card.Controls.Add(_textBox, 0, 0);
            return card;
        }

        private Control BuildFooter()
        {
            var footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 2, 0, 0),
            };

            var btnClose = UiTheme.CreateButton("关闭", false);
            btnClose.DialogResult = DialogResult.Cancel;
            footer.Controls.Add(btnClose);

            var btnCopy = UiTheme.CreateButton("复制全部", true);
            btnCopy.Click += delegate
            {
                if (!string.IsNullOrEmpty(_textBox.Text))
                {
                    Clipboard.SetText(_textBox.Text);
                }
            };
            footer.Controls.Add(btnCopy);

            var btnSave = UiTheme.CreateButton("保存为文本文件…", false);
            btnSave.Click += OnSave;
            footer.Controls.Add(btnSave);

            CancelButton = btnClose;
            return footer;
        }

        private void OnSave(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog
            {
                Title = "保存差异摘要",
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                FileName = "差异摘要_" + Path.GetFileNameWithoutExtension(_currentPath) + ".txt",
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }
                File.WriteAllText(dialog.FileName, _textBox.Text, Encoding.UTF8);
            }
        }
    }
}
