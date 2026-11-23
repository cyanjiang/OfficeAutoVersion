using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeAutoVersion.Core;

namespace OfficeAutoVersion.Common
{
    /// <summary>版本历史窗口结束后要执行的动作（由功能区在窗口关闭后执行）。</summary>
    public enum VersionHistoryAction
    {
        None = 0,
        Compare = 1,
        Restore = 2,
    }

    /// <summary>
    /// 版本历史窗口：列出一个文档的所有版本快照，并调用 Office 自带功能进行比较 / 恢复。
    ///
    /// 关键约束：本窗口是模态窗口，**不能在它显示期间操作 Office 对象模型**——
    /// 此时 Excel 的 ActiveWorkbook 会变成 null（“当前没有打开的工作簿”就是它导致的），
    /// 打开文档、并排查看等调用也可能被拒绝。
    /// 因此窗口只负责“选择”，把动作记在 PendingAction 里并关闭窗口，
    /// 由 VersioningRibbon 在 ShowDialog 返回之后执行真正的比较 / 恢复。
    /// </summary>
    public sealed class VersionHistoryForm : Form
    {
        private readonly string _appKey;
        private readonly string _documentPath;
        private readonly VersionManager _manager;
        private readonly IVersionHostActions _actions;

        private ListView _list;
        private Label _lblStatus;
        private Button _btnCompare;
        private Button _btnRestore;

        public VersionHistoryForm(string appKey, string documentPath, VersionManager manager, IVersionHostActions actions)
        {
            _appKey = appKey;
            _documentPath = documentPath;
            _manager = manager;
            _actions = actions;
            BuildUi();
            ReloadVersions();
        }

        /// <summary>窗口关闭后要执行的动作。</summary>
        public VersionHistoryAction PendingAction { get; private set; }

        /// <summary>动作作用于哪个版本文件。</summary>
        public string PendingVersionPath { get; private set; }

        /// <summary>
        /// 是否以模态方式显示（调用方在 ShowDialog 前设置）。
        /// 模态期间不能操作 Office 对象模型，此时只记录 PendingAction 并关闭窗口。
        /// </summary>
        public bool IsModal { get; set; }

        private void BuildUi()
        {
            UiTheme.PrepareForm(this, "版本历史 - " + Path.GetFileName(_documentPath));
            FormBorderStyle = FormBorderStyle.Sizable;
            ClientSize = new Size(940, 640);
            MinimumSize = new Size(860, 560);
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
            root.Controls.Add(BuildListCard(), 0, 1);
            root.Controls.Add(BuildFooter(), 0, 2);

            Controls.Add(root);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            AdjustColumnWidths();
        }

        /// <summary>顶部信息卡：标题 + 文件名 + 完整路径 + 版本数量。</summary>
        private Control BuildHeaderCard()
        {
            var card = new UiTheme.CardPanel();
            var content = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                BackColor = Color.Transparent,
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var title = UiTheme.CreateLabel("版本历史", UiTheme.Title(), UiTheme.TextPrimary);
            title.Margin = new Padding(0, 0, 0, 2);
            content.Controls.Add(title, 0, 0);

            _lblStatus = UiTheme.CreateLabel("", UiTheme.Body(), UiTheme.TextMuted);
            _lblStatus.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            _lblStatus.Margin = new Padding(12, 0, 0, 6);
            content.Controls.Add(_lblStatus, 1, 0);

            var fileName = UiTheme.CreateLabel(Path.GetFileName(_documentPath), UiTheme.BodyBold(), UiTheme.TextPrimary);
            fileName.Margin = new Padding(0, 6, 0, 2);
            content.Controls.Add(fileName, 0, 1);
            content.SetColumnSpan(fileName, 2);

            var path = new Label
            {
                Text = _documentPath,
                Font = UiTheme.Small(),
                ForeColor = UiTheme.TextMuted,
                AutoSize = false,
                AutoEllipsis = true,
                Height = 20,
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
            };
            content.Controls.Add(path, 0, 2);
            content.SetColumnSpan(path, 2);

            card.AddContent(content);
            return card;
        }

        /// <summary>中间列表卡：圆角白底 + 无边框列表，尽量接近现代列表的观感。</summary>
        private Control BuildListCard()
        {
            var card = new UiTheme.CardPanel
            {
                AutoSize = false,
                Padding = new Padding(8, 8, 8, 10),
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 10),
            };

            _list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false,
                GridLines = false,
                BorderStyle = BorderStyle.None,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                Font = UiTheme.Body(),
                BackColor = UiTheme.CardBack,
                ForeColor = UiTheme.TextPrimary,
                // 空图片列表只为了把行高撑到 28px：纯文字列表太挤
                SmallImageList = new ImageList { ImageSize = new Size(1, 28) },
            };
            _list.Columns.Add("生成时间", 180, HorizontalAlignment.Left);
            _list.Columns.Add("大小", 100, HorizontalAlignment.Right);
            _list.Columns.Add("版本文件", 560, HorizontalAlignment.Left);
            // 深色主题下列表头是系统画的白底，非常突兀；改成自绘（表头 + 单元格都接管）
            if (UiTheme.IsDark)
            {
                _list.OwnerDraw = true;
                _list.DrawColumnHeader += DrawDarkColumnHeader;
                _list.DrawSubItem += DrawDarkSubItem;
                _list.DrawItem += delegate(object s, DrawListViewItemEventArgs e) { e.DrawDefault = true; };
            }
            _list.DoubleClick += delegate { CompareSelected(); };
            _list.SelectedIndexChanged += delegate { UpdateButtonState(); };
            _list.Resize += delegate { AdjustColumnWidths(); };

            card.Controls.Add(_list, 0, 0);
            NativeTheme.ApplyToControl(_list);
            return card;
        }

        /// <summary>深色主题的自绘列表头：深底 + 浅字 + 底部一条分隔线。</summary>
        private void DrawDarkColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            using (var brush = new SolidBrush(UiTheme.CardBack))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }
            using (var pen = new Pen(UiTheme.CardBorder))
            {
                e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            }
            var alignment = e.Header.TextAlign == HorizontalAlignment.Right
                ? TextFormatFlags.Right
                : TextFormatFlags.Left;
            TextRenderer.DrawText(e.Graphics, e.Header.Text, _list.Font,
                new Rectangle(e.Bounds.Left + 4, e.Bounds.Top, e.Bounds.Width - 8, e.Bounds.Height),
                UiTheme.TextMuted, alignment | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        /// <summary>深色主题下交还默认单元格绘制（选中高亮等仍是系统行为）。</summary>
        private void DrawDarkSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            e.DrawDefault = true;
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

            _btnCompare = UiTheme.CreateButton(
                (_actions != null && !string.IsNullOrEmpty(_actions.CompareButtonText))
                    ? _actions.CompareButtonText
                    : "比较差异", true);
            _btnCompare.Click += delegate { CompareSelected(); };
            footer.Controls.Add(_btnCompare);
            if (_actions != null && !string.IsNullOrEmpty(_actions.CompareHint))
            {
                var tip = new ToolTip();
                tip.SetToolTip(_btnCompare, _actions.CompareHint);
            }

            // 差异摘要：纯文件解析（zip + XML），不碰 Office 对象模型，
            // 因此模态期间也可以安全执行，不需要 PendingAction 机制
            var btnDiff = UiTheme.CreateButton("差异摘要", false);
            btnDiff.Click += delegate { ShowDiffSummary(); };
            footer.Controls.Add(btnDiff);

            _btnRestore = UiTheme.CreateButton("恢复到此版本", false);
            _btnRestore.Click += delegate { RestoreSelected(); };
            footer.Controls.Add(_btnRestore);

            var btnOpen = UiTheme.CreateButton("打开", false);
            btnOpen.Click += OnOpen;
            footer.Controls.Add(btnOpen);

            var btnFolder = UiTheme.CreateButton("打开文件夹", false);
            btnFolder.Click += OnOpenFolder;
            footer.Controls.Add(btnFolder);

            CancelButton = btnClose;

            if (_actions == null)
            {
                _btnCompare.Enabled = false;
                _btnRestore.Enabled = false;
            }
            return footer;
        }

        /// <summary>窗口宽度变化时让“版本文件”列吃掉剩余空间，不出现横向空白。</summary>
        private void AdjustColumnWidths()
        {
            if (_list == null || _list.Columns.Count < 3)
            {
                return;
            }
            try
            {
                var available = Math.Max(240, _list.ClientSize.Width - 4);
                var time = 180;
                var size = 100;
                _list.Columns[0].Width = time;
                _list.Columns[1].Width = size;
                _list.Columns[2].Width = Math.Max(160, available - time - size);
            }
            catch
            {
                // 列宽只是观感，失败不影响功能
            }
        }

        private void ReloadVersions()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            var versions = _manager != null ? _manager.ListVersions(_documentPath) : new List<string>();
            foreach (var path in versions)
            {
                var item = new ListViewItem(FormatTime(path));
                item.SubItems.Add(FormatSize(path));
                item.SubItems.Add(Path.GetFileName(path));
                item.Tag = path;
                _list.Items.Add(item);
            }
            _list.EndUpdate();

            if (_list.Items.Count > 0)
            {
                _list.Items[0].Selected = true;
                _list.Items[0].Focused = true;
            }

            _lblStatus.Text = versions.Count == 0
                ? "还没有任何版本。保存文档后会自动生成。"
                : string.Format("共 {0} 个版本（最新的在最上面）", versions.Count);
            AdjustColumnWidths();
            UpdateButtonState();
        }

        private void UpdateButtonState()
        {
            var hasSelection = GetSelectedVersion() != null;
            if (_actions != null)
            {
                _btnCompare.Enabled = hasSelection;
                _btnRestore.Enabled = hasSelection;
            }
        }

        private string GetSelectedVersion()
        {
            if (_list.SelectedItems.Count > 0)
            {
                return _list.SelectedItems[0].Tag as string;
            }
            // 列表尚未创建窗口句柄时（例如自动化测试）SelectedItems 可能为空，
            // 此时按“默认选中最新版本”处理，与界面上的初始状态一致。
            if (_list.Items.Count > 0)
            {
                return _list.Items[0].Tag as string;
            }
            return null;
        }

        /// <summary>当前选中的版本文件路径（供界面与自动化测试使用）。</summary>
        public string SelectedVersionPath
        {
            get { return GetSelectedVersion(); }
        }

        /// <summary>版本数量。</summary>
        public int VersionCount
        {
            get { return _list.Items.Count; }
        }

        /// <summary>用 Office 自带功能比较“当前文档”与选中的版本。</summary>
        public void CompareSelected()
        {
            var version = GetSelectedVersion();
            if (version == null || _actions == null)
            {
                return;
            }

            // 模态窗口：只记下意图并关闭窗口，真正的比较交给关闭之后的代码执行
            if (IsModal)
            {
                PendingAction = VersionHistoryAction.Compare;
                PendingVersionPath = version;
                Close();
                return;
            }

            // 非模态（自动化测试）：直接执行，走同一条调用链
            try
            {
                _actions.CompareWithCurrent(_documentPath, version);
            }
            catch (Exception ex)
            {
                ShowError("比较失败", ex);
            }
        }

        /// <summary>显示“当前文档 vs 选中版本”的内容差异摘要（新增/删除/修改各多少处 + 明细）。</summary>
        public void ShowDiffSummary()
        {
            var version = GetSelectedVersion();
            if (version == null)
            {
                return;
            }
            try
            {
                var summary = VersionDiffer.Diff(_documentPath, version);
                using (var form = new DiffSummaryForm(summary, _documentPath, version, BuildPageLookup()))
                {
                    form.ShowDialog(this);
                }
            }
            catch (Exception ex)
            {
                ShowError("差异摘要失败", ex);
            }
        }

        /// <summary>
        /// 让宿主给出“段落文本 → 页码”的映射（页码是排版结果，文件里没有）。
        /// 拿不到就返回 null，摘要照常显示，只是不标页码。
        /// </summary>
        private Func<string, int> BuildPageLookup()
        {
            try
            {
                if (_actions == null)
                {
                    return null;
                }
                var map = _actions.GetTextPageMap(_documentPath);
                if (map == null || map.Count == 0)
                {
                    return null;
                }
                return delegate(string text)
                {
                    int page = 0;
                    return map.TryGetValue(ContentDiff.NormalizeKey(text), out page) ? page : 0;
                };
            }
            catch
            {
                return null; // 页码只是附加信息，取不到不能影响摘要
            }
        }

        /// <summary>把当前文档恢复到选中的版本（恢复前会先建安全快照）。</summary>
        public void RestoreSelected()
        {
            var version = GetSelectedVersion();
            if (version == null || _actions == null)
            {
                return;
            }

            var message =
                "将用下面这个版本的内容，覆盖当前文件在磁盘上的内容：\n\n" +
                Path.GetFileName(version) + "\n\n" +
                "恢复前会先保存当前文档，并把它的状态另存为一个新版本，\n" +
                "所以恢复这一步不会丢失你看得见的内容，也能随时再恢复回来。\n\n" +
                "确定继续？";
            if (MessageBox.Show(this, message, "恢复版本", MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK)
            {
                return;
            }

            // 模态窗口：记下意图并关闭窗口，真正的恢复交给关闭之后的代码执行
            if (IsModal)
            {
                PendingAction = VersionHistoryAction.Restore;
                PendingVersionPath = version;
                Close();
                return;
            }

            // 非模态（自动化测试）：直接执行
            try
            {
                var safetySnapshot = _actions.RestoreVersion(_documentPath, version);
                ShowRestoreResult(null, version, safetySnapshot);
                ReloadVersions();
            }
            catch (Exception ex)
            {
                ShowError("恢复失败", ex);
            }
        }

        /// <summary>恢复完成后的统一提示（窗口与功能区共用）。</summary>
        internal static void ShowRestoreResult(IWin32Window owner, string versionFilePath, string safetySnapshot)
        {
            var text = "已恢复到：" + Path.GetFileName(versionFilePath);
            if (!string.IsNullOrEmpty(safetySnapshot))
            {
                text += "\n\n恢复前的状态已保存为：\n" + Path.GetFileName(safetySnapshot) +
                        "\n如果恢复错了，可以再选中它恢复回来。";
            }
            MessageBox.Show(owner, text, "恢复完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void OnOpen(object sender, EventArgs e)
        {
            var version = GetSelectedVersion();
            if (version == null)
            {
                return;
            }
            try
            {
                Process.Start("explorer.exe", "\"" + version + "\"");
            }
            catch (Exception ex)
            {
                ShowError("打开失败", ex);
            }
        }

        private void OnOpenFolder(object sender, EventArgs e)
        {
            var version = GetSelectedVersion();
            var directory = version != null
                ? Path.GetDirectoryName(version)
                : VersionPaths.GetVersionsDirectory(_documentPath, OptionsStore.Load(_appKey));
            try
            {
                if (Directory.Exists(directory))
                {
                    Process.Start("explorer.exe", "\"" + directory + "\"");
                }
            }
            catch (Exception ex)
            {
                ShowError("打开文件夹失败", ex);
            }
        }

        private void ShowError(string title, Exception ex)
        {
            try
            {
                Logger.Default.Error(title, ex);
            }
            catch
            {
                // 忽略
            }
            MessageBox.Show(this, title + "：" + ex.Message, "自动版本",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static string FormatTime(string path)
        {
            try
            {
                return File.GetLastWriteTime(path).ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string FormatSize(string path)
        {
            try
            {
                var bytes = new FileInfo(path).Length;
                if (bytes < 1024)
                {
                    return bytes + " B";
                }
                if (bytes < 1024 * 1024)
                {
                    return (bytes / 1024.0).ToString("0.0") + " KB";
                }
                return (bytes / (1024.0 * 1024.0)).ToString("0.0") + " MB";
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
