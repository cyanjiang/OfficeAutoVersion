using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OfficeAutoVersion.Setup
{
    /// <summary>
    /// 安装程序配色：跟 Office 的“Office 主题”一致（浅色／深色），
    /// 主色用 Office 蓝——安装程序不属于某个具体应用，所以不取 Word/Excel 的品牌色。
    /// </summary>
    internal static class Palette
    {
        public static readonly bool Dark;
        public static readonly Color Window;
        public static readonly Color Card;
        public static readonly Color Border;
        public static readonly Color Text;
        public static readonly Color Muted;
        public static readonly Color Accent;
        public static readonly Color AccentHover;
        public static readonly Color Success;
        public static readonly Color Danger;

        static Palette()
        {
            Dark = DetectDark();
            if (Dark)
            {
                Window = Color.FromArgb(0x1B, 0x1A, 0x19);
                Card = Color.FromArgb(0x25, 0x24, 0x23);
                Border = Color.FromArgb(0x3B, 0x3A, 0x39);
                Text = Color.FromArgb(0xF3, 0xF2, 0xF1);
                Muted = Color.FromArgb(0xB3, 0xB0, 0xAD);
                Accent = Color.FromArgb(0x47, 0x9E, 0xF5);
                AccentHover = Color.FromArgb(0x2B, 0x88, 0xE0);
                Success = Color.FromArgb(0x6C, 0xCB, 0x5F);
                Danger = Color.FromArgb(0xE8, 0x74, 0x7C);
            }
            else
            {
                Window = Color.FromArgb(0xF3, 0xF2, 0xF1);
                Card = Color.White;
                Border = Color.FromArgb(0xE1, 0xDF, 0xDD);
                Text = Color.FromArgb(0x20, 0x1F, 0x1E);
                Muted = Color.FromArgb(0x60, 0x5E, 0x5C);
                Accent = Color.FromArgb(0x0F, 0x6C, 0xBD);
                AccentHover = Color.FromArgb(0x11, 0x5E, 0xA3);
                Success = Color.FromArgb(0x10, 0x7C, 0x41);
                Danger = Color.FromArgb(0xA4, 0x26, 0x2C);
            }
        }

        private static bool DetectDark()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common"))
                {
                    if (key == null)
                    {
                        return false;
                    }
                    var raw = key.GetValue("UI Theme");
                    if (raw == null)
                    {
                        return false;
                    }
                    return Convert.ToInt32(raw) >= 4;
                }
            }
            catch
            {
                return false;
            }
        }

        public static string FontFamily
        {
            get { return "Microsoft YaHei UI"; }
        }

        public static Font Font(float size, FontStyle style)
        {
            try
            {
                return new Font(FontFamily, size, style, GraphicsUnit.Point);
            }
            catch
            {
                return new Font(SystemFonts.MessageBoxFont.FontFamily, size, style, GraphicsUnit.Point);
            }
        }
    }

    /// <summary>
    /// 圆角卡片容器。继承 TableLayoutPanel 让高度真正由内容决定（AutoSize），
    /// 避免“固定高度的面板把内容裁掉”。用 <see cref="AddContent"/> 放内容。
    /// </summary>
    internal sealed class CardPanel : TableLayoutPanel
    {
        public CardPanel()
        {
            SetStyle(ControlStyles.UserPaint
                     | ControlStyles.AllPaintingInWmPaint
                     | ControlStyles.OptimizedDoubleBuffer
                     | ControlStyles.ResizeRedraw, true);
            ColumnCount = 1;
            RowCount = 1;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Dock = DockStyle.Fill;
            BackColor = Palette.Window;
            Padding = new Padding(20, 14, 20, 16);
            Margin = new Padding(0, 0, 0, 12);
            ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        /// <summary>放入内容（自动 Dock=Top，宽度跟随卡片）。</summary>
        public void AddContent(Control content)
        {
            content.Dock = DockStyle.Top;
            Controls.Add(content, 0, 0);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            using (var brush = new SolidBrush(Palette.Window))
            {
                e.Graphics.FillRectangle(brush, ClientRectangle);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }
            using (var path = Rounded(bounds, 10))
            using (var fill = new SolidBrush(Palette.Card))
            using (var border = new Pen(Palette.Border))
            {
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(border, path);
            }
        }

        private static GraphicsPath Rounded(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            var d = radius * 2;
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    /// <summary>安装／卸载窗口。</summary>
    internal sealed class SetupForm : Form
    {
        private readonly bool _uninstallMode;

        private Label _lblNet;
        private Label _lblVsto;
        private LinkLabel _lnkVsto;
        private Label _lblOffice;
        private Label _lblStatus;
        private ProgressBar _progress;
        private Button _btnPrimary;
        private Button _btnClose;

        private bool _busy;
        private bool _done;
        private Control _content;

        public SetupForm(bool uninstallMode)
        {
            _uninstallMode = uninstallMode;
            BuildUi();
            if (!_uninstallMode)
            {
                // 卸载模式下没有“环境检查”卡片，这些标签不存在
                UpdateRequirements();
            }
        }

        // ------------------------------------------------------------ 界面

        private void BuildUi()
        {
            Text = (_uninstallMode ? "卸载 " : "安装 ") + Product.Name;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(680, 596);
            BackColor = Palette.Window;
            ForeColor = Palette.Text;
            Font = Palette.Font(9.5f, FontStyle.Regular);
            AutoScaleMode = AutoScaleMode.Dpi;
            DoubleBuffered = true;

            var header = BuildHeader();
            var content = BuildContent();
            var footer = BuildFooter();

            Controls.Add(content);
            Controls.Add(footer);
            Controls.Add(header);

            NativeTheme.Apply(this);
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 108,
                BackColor = Palette.Accent,
            };

            var title = new Label
            {
                Text = _uninstallMode ? "卸载 " + Product.Name : "安装 " + Product.Name,
                Font = Palette.Font(18f, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(26, 22),
                BackColor = Color.Transparent,
            };

            var subtitle = new Label
            {
                Text = _uninstallMode
                    ? "将从 Office 中移除插件；已保存的历史版本文件不会被删除。"
                    : "保存文档时自动留下历史版本，随时比较、恢复。",
                Font = Palette.Font(9.5f, FontStyle.Regular),
                ForeColor = Color.FromArgb(235, 255, 255, 255),
                AutoSize = true,
                Location = new Point(28, 64),
                BackColor = Color.Transparent,
            };

            var version = new Label
            {
                Text = "v" + Product.Version,
                Font = Palette.Font(9f, FontStyle.Regular),
                ForeColor = Color.FromArgb(200, 255, 255, 255),
                AutoSize = true,
                Location = new Point(ClientSize.Width - 84, 30),
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
            };

            header.Controls.Add(title);
            header.Controls.Add(subtitle);
            header.Controls.Add(version);
            return header;
        }

        private Control BuildContent()
        {
            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Palette.Window,
                Padding = new Padding(26, 20, 26, 8),
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            content.Controls.Add(_uninstallMode ? BuildInfoCard() : BuildCheckCard());
            content.Controls.Add(BuildLocationCard());
            content.Controls.Add(BuildHintLabel());
            content.Controls.Add(BuildProgressArea());

            content.RowCount = content.Controls.Count;
            for (var i = 0; i < content.RowCount; i++)
            {
                content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }
            _content = content;
            return content;
        }

        /// <summary>首次显示时按内容实际高度收紧窗口，避免底部留一大块空白。</summary>
        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            try
            {
                var preferred = _content.GetPreferredSize(new Size(ClientSize.Width, 0));
                var target = 108 + preferred.Height + 78;
                ClientSize = new Size(ClientSize.Width, Math.Max(430, Math.Min(700, target)));
            }
            catch
            {
                // 尺寸兜底失败不影响功能
            }
        }

        /// <summary>“环境检查”卡片（安装模式）。</summary>
        private Control BuildCheckCard()
        {
            var card = new CardPanel { Dock = DockStyle.Top };
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            table.Controls.Add(SectionTitle("环境检查"), 0, 0);
            table.SetColumnSpan(table.GetControlFromPosition(0, 0), 3);

            _lblNet = StatusLabel();
            AddRow(table, 1, ".NET Framework 4.8", _lblNet, null);

            _lblVsto = StatusLabel();
            _lnkVsto = new LinkLabel
            {
                Text = "去下载",
                AutoSize = true,
                Font = Palette.Font(9f, FontStyle.Regular),
                LinkColor = Palette.Accent,
                ActiveLinkColor = Palette.AccentHover,
                VisitedLinkColor = Palette.Accent,
                BackColor = Color.Transparent,
                Margin = new Padding(12, 6, 0, 6),
                Visible = false,
            };
            _lnkVsto.LinkClicked += delegate
            {
                try
                {
                    System.Diagnostics.Process.Start(
                        "https://www.microsoft.com/download/details.aspx?id=56961");
                }
                catch
                {
                    MessageBox.Show(this,
                        "无法打开浏览器，请手动访问：\nhttps://www.microsoft.com/download/details.aspx?id=56961",
                        Product.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            AddRow(table, 2, "VSTO 运行时 4.0", _lblVsto, _lnkVsto);

            _lblOffice = StatusLabel();
            AddRow(table, 3, "Office 应用", _lblOffice, null);

            table.RowCount = 4;
            for (var i = 0; i < 4; i++)
            {
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            card.AddContent(table);
            return card;
        }

        /// <summary>“已安装版本”卡片（卸载模式）。</summary>
        private Control BuildInfoCard()
        {
            var card = new CardPanel { Dock = DockStyle.Top };
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            table.Controls.Add(SectionTitle("卸载信息"), 0, 0);
            table.SetColumnSpan(table.GetControlFromPosition(0, 0), 2);

            var version = Installer.InstalledVersion();
            AddRow(table, 1, "已安装版本", ValueLabel(version == null ? "（未记录）" : "v" + version), null);
            AddRow(table, 2, "说明", ValueLabel("只移除插件本身和历史版本的“注册”，版本文件保留在原处。"), null);

            table.RowCount = 3;
            for (var i = 0; i < 4; i++)
            {
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            card.AddContent(table);
            return card;
        }

        private Control BuildLocationCard()
        {
            var card = new CardPanel { Dock = DockStyle.Top };
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            table.Controls.Add(SectionTitle("安装位置"), 0, 0);
            table.SetColumnSpan(table.GetControlFromPosition(0, 0), 2);

            var path = new Label
            {
                Text = Product.InstallDir,
                Font = Palette.Font(9f, FontStyle.Regular),
                ForeColor = Palette.Accent,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 4, 0, 4),
            };
            AddRow(table, 1, "", path, null);
            AddRow(table, 2, "",
                ValueLabel("仅为当前用户安装，不需要管理员权限；卸载可在“设置 → 应用”中找到。"), null);

            table.RowCount = 3;
            for (var i = 0; i < 3; i++)
            {
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            card.AddContent(table);
            return card;
        }

        private Control BuildHintLabel()
        {
            var hint = new Label
            {
                Text = _uninstallMode
                    ? "建议先关闭 Word / Excel / PowerPoint，否则正在使用的程序文件需要重启后才能删除。历史版本文件不会被删除。"
                    : "安装前请关闭 Word / Excel / PowerPoint；装好后重新打开它们，功能区会多出“自动版本”选项卡。左下角可打开《使用说明》。",
                Font = Palette.Font(9f, FontStyle.Regular),
                ForeColor = Palette.Muted,
                AutoSize = true,
                MaximumSize = new Size(600, 0),
                BackColor = Color.Transparent,
                Margin = new Padding(4, 0, 0, 12),
            };
            return hint;
        }

        private Control BuildProgressArea()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
                Margin = new Padding(4, 0, 4, 0),
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            _progress = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 8,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Style = ProgressBarStyle.Continuous,
                Margin = new Padding(0, 0, 0, 8),
            };

            _lblStatus = new Label
            {
                Text = _uninstallMode ? "准备卸载" : "准备安装",
                Font = Palette.Font(9f, FontStyle.Regular),
                ForeColor = Palette.Muted,
                AutoSize = true,
                BackColor = Color.Transparent,
            };

            panel.Controls.Add(_progress);
            panel.Controls.Add(_lblStatus);
            panel.RowCount = 2;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return panel;
        }

        private Control BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 68,
                BackColor = Palette.Window,
                Padding = new Padding(26, 0, 26, 16),
            };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 12, 0, 0),
            };

            _btnPrimary = CreateButton(_uninstallMode ? "卸载" : "安装", true);
            _btnPrimary.Click += OnPrimaryClick;

            _btnClose = CreateButton("关闭", false);
            _btnClose.Click += delegate { Close(); };

            buttons.Controls.Add(_btnPrimary);
            buttons.Controls.Add(_btnClose);
            footer.Controls.Add(buttons);

            // 左下角：随时可打开随包发布的使用说明
            var link = new LinkLabel
            {
                Text = "查看使用说明",
                Dock = DockStyle.Left,
                AutoSize = false,
                Width = 160,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = Palette.Font(9f, FontStyle.Regular),
                LinkColor = Palette.Accent,
                ActiveLinkColor = Palette.AccentHover,
                LinkBehavior = LinkBehavior.HoverUnderline,
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand,
            };
            link.Click += delegate { OnOpenDoc(); };
            footer.Controls.Add(link);
            return footer;
        }

        /// <summary>打开使用说明；文档缺失时给出明确位置，避免用户以为安装坏了。</summary>
        private void OnOpenDoc()
        {
            if (Installer.OpenDoc())
            {
                return;
            }
            MessageBox.Show(this,
                "没有找到使用说明文档。\n\n它应当位于：\n" + Product.DocPath +
                "\n\n如果安装包里不含该文件，可以重新获取安装包再安装一次。",
                "使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static Button CreateButton(string text, bool primary)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = false,
                Size = new Size(primary ? 116 : 96, 36),
                FlatStyle = FlatStyle.Flat,
                Font = Palette.Font(9.5f, FontStyle.Regular),
                Margin = new Padding(10, 0, 0, 0),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
            };
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = Palette.Border;
            if (primary)
            {
                button.BackColor = Palette.Accent;
                button.ForeColor = Color.White;
                button.FlatAppearance.MouseOverBackColor = Palette.AccentHover;
                button.FlatAppearance.MouseDownBackColor = Palette.AccentHover;
            }
            else
            {
                button.BackColor = Palette.Card;
                button.ForeColor = Palette.Text;
                button.FlatAppearance.MouseOverBackColor = Palette.Window;
                button.FlatAppearance.MouseDownBackColor = Palette.Border;
            }
            return button;
        }

        private static Label SectionTitle(string text)
        {
            return new Label
            {
                Text = text,
                Font = Palette.Font(10.5f, FontStyle.Bold),
                ForeColor = Palette.Text,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 8),
            };
        }

        private static Label StatusLabel()
        {
            return new Label
            {
                Text = "",
                Font = Palette.Font(9.5f, FontStyle.Regular),
                ForeColor = Palette.Text,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 6, 0, 6),
            };
        }

        private static Label ValueLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = Palette.Font(9f, FontStyle.Regular),
                ForeColor = Palette.Muted,
                AutoSize = true,
                MaximumSize = new Size(540, 0),
                BackColor = Color.Transparent,
                Margin = new Padding(0, 4, 0, 4),
            };
        }

        private static void AddRow(TableLayoutPanel table, int row, string caption, Control value, Control extra)
        {
            var label = new Label
            {
                Text = caption,
                Font = Palette.Font(9.5f, FontStyle.Regular),
                ForeColor = Palette.Muted,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 6, 16, 6),
            };
            table.Controls.Add(label, 0, row);
            table.Controls.Add(value, 1, row);
            if (extra != null)
            {
                table.Controls.Add(extra, 2, row);
            }
        }

        // ------------------------------------------------------------ 状态

        private void UpdateRequirements()
        {
            if (_lblNet == null || _lblVsto == null || _lblOffice == null)
            {
                return;
            }
            var net = Requirements.DotNet48Installed();
            _lblNet.Text = net ? "✓ 已安装" : "✗ 未安装";
            _lblNet.ForeColor = net ? Palette.Success : Palette.Danger;

            var vsto = Requirements.VstoRuntimeInstalled();
            _lblVsto.Text = vsto ? "✓ 已安装" : "✗ 未安装";
            _lblVsto.ForeColor = vsto ? Palette.Success : Palette.Danger;
            _lnkVsto.Visible = !vsto;

            List<string> hosts;
            var office = Requirements.OfficeInstalled(out hosts);
            _lblOffice.Text = office ? "✓ " + Join(hosts) : "✗ 未检测到 Word / Excel / PowerPoint";
            _lblOffice.ForeColor = office ? Palette.Success : Palette.Danger;

            if (!vsto)
            {
                _btnPrimary.Enabled = false;
                _lblStatus.Text = "缺少 VSTO 4.0 运行时，安装无法继续。";
                _lblStatus.ForeColor = Palette.Danger;
            }
            else if (!office)
            {
                _btnPrimary.Enabled = false;
                _lblStatus.Text = "没有检测到 Office，请先安装 Word / Excel / PowerPoint。";
                _lblStatus.ForeColor = Palette.Danger;
            }
            else
            {
                var installed = Installer.InstalledVersion();
                if (installed != null)
                {
                    _btnPrimary.Text = "重新安装";
                    _lblStatus.Text = "已安装 v" + installed + "，再次安装会覆盖程序文件。";
                }
            }
        }

        private static string Join(List<string> items)
        {
            var text = new System.Text.StringBuilder();
            foreach (var item in items)
            {
                if (text.Length > 0)
                {
                    text.Append("、");
                }
                text.Append(item);
            }
            return text.ToString();
        }

        // ------------------------------------------------------------ 动作

        private void OnPrimaryClick(object sender, EventArgs e)
        {
            if (_done)
            {
                Close();
                return;
            }
            if (_busy)
            {
                return;
            }

            var running = Installer.RunningOfficeApps();
            if (running.Count > 0)
            {
                MessageBox.Show(this,
                    "请先保存并关闭以下应用，然后重试：\n\n" + Join(running),
                    Product.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_uninstallMode)
            {
                var confirm = MessageBox.Show(this,
                    "确定要卸载 " + Product.Name + " 吗？\n\n" +
                    "插件会从 Word / Excel / PowerPoint 中移除，\n但已经保存的历史版本文件不会被删除。",
                    "卸载", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);
                if (confirm != DialogResult.OK)
                {
                    return;
                }
            }

            RunWorker(_uninstallMode ? (Action)Installer.Uninstall : Installer.Install);
        }

        private void RunWorker(Action work)
        {
            _busy = true;
            _btnPrimary.Enabled = false;
            _btnClose.Enabled = false;
            _progress.Value = 0;
            _lblStatus.ForeColor = Palette.Muted;
            _lblStatus.Text = "正在处理…";

            Installer.Progress += OnProgress;

            var thread = new Thread(delegate()
            {
                Exception error = null;
                try
                {
                    work();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
                try
                {
                    if (!IsDisposed && IsHandleCreated)
                    {
                        BeginInvoke(new Action<Exception>(Finish), new object[] { error });
                    }
                }
                catch
                {
                    // 窗口已关闭
                }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void OnProgress(int percent, string message)
        {
            if (IsDisposed || !IsHandleCreated)
            {
                return;
            }
            if (InvokeRequired)
            {
                try
                {
                    BeginInvoke(new Action<int, string>(OnProgress), new object[] { percent, message });
                }
                catch
                {
                    // 忽略
                }
                return;
            }
            _progress.Value = Math.Max(_progress.Minimum, Math.Min(_progress.Maximum, percent));
            _lblStatus.Text = message;
        }

        private void Finish(Exception error)
        {
            Installer.Progress -= OnProgress;
            _busy = false;
            _btnClose.Enabled = true;
            _btnPrimary.Enabled = true;

            if (error != null)
            {
                _progress.Value = 0;
                _lblStatus.ForeColor = Palette.Danger;
                _lblStatus.Text = (_uninstallMode ? "卸载失败：" : "安装失败：") + error.Message;
                MessageBox.Show(this,
                    (_uninstallMode ? "卸载失败：" : "安装失败：") + "\n\n" + error.Message,
                    Product.Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _done = true;
            _progress.Value = 100;
            _lblStatus.ForeColor = Palette.Success;
            _btnPrimary.Text = "完成";

            if (_uninstallMode)
            {
                _lblStatus.Text = "已卸载。重新打开 Office 后功能区不再显示“自动版本”。";
                MessageBox.Show(this,
                    "已卸载 " + Product.Name + "。\n\n" +
                    "重新打开 Word / Excel / PowerPoint 后生效。\n" +
                    "已保存的历史版本文件仍保留在原处，不会被删除。",
                    Product.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                _lblStatus.Text = "安装完成。重新打开 Word / Excel / PowerPoint 即可使用。";
                MessageBox.Show(this,
                    "安装完成。\n\n" +
                    "重新打开 Word / Excel / PowerPoint，功能区会多出“自动版本”选项卡：\n" +
                    "  · 保存文档时自动建立版本\n" +
                    "  · “版本历史”里可以比较、恢复、查看差异摘要\n\n" +
                    "安装位置：" + Product.InstallDir,
                    Product.Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>深色主题下把标题栏也变深（失败不影响功能）。</summary>
    internal static class NativeTheme
    {
        private const int DwmUseImmersiveDarkMode = 20;
        private const int DwmUseImmersiveDarkModeLegacy = 19;

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        public static void Apply(Form form)
        {
            if (!Palette.Dark)
            {
                return;
            }
            try
            {
                if (form.IsHandleCreated)
                {
                    SetDark(form.Handle);
                }
                else
                {
                    form.HandleCreated += delegate { SetDark(form.Handle); };
                }
            }
            catch
            {
                // 外观优化，失败忽略
            }
        }

        private static void SetDark(IntPtr hwnd)
        {
            try
            {
                var enabled = 1;
                if (DwmSetWindowAttribute(hwnd, DwmUseImmersiveDarkMode, ref enabled, sizeof(int)) != 0)
                {
                    DwmSetWindowAttribute(hwnd, DwmUseImmersiveDarkModeLegacy, ref enabled, sizeof(int));
                }
            }
            catch
            {
                // 忽略
            }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            var silent = HasFlag(args, "/s", "/silent", "/quiet");
            var uninstall = HasFlag(args, "/uninstall", "/u");
            var check = HasFlag(args, "/check");
            // 开发用：把窗口渲染成 PNG 核对排版，不写注册表、不复制文件
            var shot = GetValue(args, "/shot:");

            if (shot != null)
            {
                return RenderShot(shot, uninstall);
            }
            if (check)
            {
                return RunCheck();
            }
            if (silent)
            {
                return RunSilent(uninstall);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SetupForm(uninstall));
            return 0;
        }

        /// <summary>把安装窗口画成 PNG（仅开发期使用：/shot:路径）。</summary>
        private static int RenderShot(string path, bool uninstall)
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using (var form = new SetupForm(uninstall))
                {
                    form.StartPosition = FormStartPosition.Manual;
                    form.Location = new Point(40, 40);
                    form.Show();
                    Application.DoEvents();
                    Thread.Sleep(400);
                    Application.DoEvents();
                    using (var bitmap = new Bitmap(form.Width, form.Height))
                    {
                        form.DrawToBitmap(bitmap, new Rectangle(0, 0, form.Width, form.Height));
                        bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    form.Close();
                }
                return 0;
            }
            catch (Exception ex)
            {
                // 开发模式：把异常写到目标文件旁边，方便排查（不弹窗）
                try
                {
                    File.WriteAllText(path + ".txt", ex.ToString(), Encoding.UTF8);
                }
                catch
                {
                    // 忽略
                }
                return 1;
            }
        }

        private static string GetValue(string[] args, string prefix)
        {
            if (args == null)
            {
                return null;
            }
            foreach (var arg in args)
            {
                if (arg != null && arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return arg.Substring(prefix.Length).Trim('"');
                }
            }
            return null;
        }

        private static bool HasFlag(string[] args, params string[] names)
        {
            if (args == null)
            {
                return false;
            }
            foreach (var arg in args)
            {
                foreach (var name in names)
                {
                    if (string.Equals(arg, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>静默安装 / 卸载：把过程写日志，返回 0 表示成功。</summary>
        private static int RunSilent(bool uninstall)
        {
            var log = new System.Text.StringBuilder();
            log.AppendLine("OfficeAutoVersion setup " + Product.Version);
            log.AppendLine("mode: " + (uninstall ? "uninstall" : "install"));
            Installer.Progress += delegate(int percent, string message)
            {
                log.AppendLine(percent.ToString().PadLeft(3) + "%  " + message);
            };
            try
            {
                if (uninstall)
                {
                    Installer.Uninstall();
                }
                else
                {
                    Installer.Install();
                }
                log.AppendLine("result: OK");
                return 0;
            }
            catch (Exception ex)
            {
                log.AppendLine("result: FAILED - " + ex.Message);
                return 1;
            }
            finally
            {
                try
                {
                    File.WriteAllText(
                        Path.Combine(Path.GetTempPath(), "OfficeAutoVersion-setup.log"),
                        log.ToString(), Encoding.UTF8);
                }
                catch
                {
                    // 忽略
                }
            }
        }

        /// <summary>环境自检：给用户/支持人员看的一屏结论。</summary>
        private static int RunCheck()
        {
            List<string> hosts;
            var office = Requirements.OfficeInstalled(out hosts);
            var vsto = Requirements.VstoRuntimeInstalled();
            var installed = Installer.InstalledVersion();

            var text = new System.Text.StringBuilder();
            text.AppendLine("产品版本：" + Product.Version);
            text.AppendLine(".NET Framework 4.8：" + (Requirements.DotNet48Installed() ? "已安装" : "未安装"));
            text.AppendLine("VSTO 4.0 运行时：" + (vsto ? "已安装" : "未安装"));
            text.AppendLine("Office 应用：" + (office ? "Word / Excel / PowerPoint 已检测到" : "未检测到"));
            text.AppendLine("插件状态：" + (installed == null ? "未安装" : "已安装 v" + installed));
            text.AppendLine("安装目录：" + Product.InstallDir);

            MessageBox.Show(text.ToString(), Product.Name + " 环境自检",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return vsto && office ? 0 : 1;
        }
    }
}
