using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// 一套调色板：浅色来自 Office「彩色/白色」主题，深色来自「深灰色/黑色」主题，
    /// 主色取宿主应用的品牌色（Word 蓝 / Excel 绿 / PowerPoint 橙红）。
    /// </summary>
    internal sealed class ThemePalette
    {
        public string Name;
        public bool Dark;

        public Color WindowBack;
        public Color CardBack;
        public Color CardBorder;
        public Color TextPrimary;
        public Color TextMuted;

        /// <summary>主色：按钮填充、强调条。</summary>
        public Color Accent;

        /// <summary>主色用于文字时（深色主题下需要更亮，保证对比度）。</summary>
        public Color AccentText;

        public Color AccentHover;
        public Color AccentPressed;
        public Color SubtleHover;
        public Color SubtlePressed;
        public Color InputBack;
        public Color InputBorder;
        public Color Danger;

        public static ThemePalette Create(OfficeThemeKind kind, Color accent)
        {
            var palette = new ThemePalette();
            palette.Accent = accent;

            switch (kind)
            {
                case OfficeThemeKind.Black:
                    palette.Name = "黑色（深色）";
                    palette.Dark = true;
                    palette.WindowBack = Color.FromArgb(0x1B, 0x1A, 0x19);
                    palette.CardBack = Color.FromArgb(0x25, 0x24, 0x23);
                    palette.CardBorder = Color.FromArgb(0x3B, 0x3A, 0x39);
                    palette.TextPrimary = Color.FromArgb(0xF3, 0xF2, 0xF1);
                    palette.TextMuted = Color.FromArgb(0xB3, 0xB0, 0xAD);
                    palette.AccentText = Blend(accent, Color.White, 0.46);
                    palette.AccentHover = Blend(accent, Color.White, 0.12);
                    palette.AccentPressed = Blend(accent, Color.Black, 0.16);
                    palette.SubtleHover = Color.FromArgb(0x2F, 0x2E, 0x2D);
                    palette.SubtlePressed = Color.FromArgb(0x3B, 0x3A, 0x39);
                    palette.InputBack = Color.FromArgb(0x20, 0x1F, 0x1E);
                    palette.InputBorder = Color.FromArgb(0x48, 0x46, 0x44);
                    palette.Danger = Color.FromArgb(0xE8, 0x74, 0x7C);
                    break;

                case OfficeThemeKind.DarkGray:
                    palette.Name = "深灰色（深色）";
                    palette.Dark = true;
                    palette.WindowBack = Color.FromArgb(0x2B, 0x2A, 0x29);
                    palette.CardBack = Color.FromArgb(0x32, 0x31, 0x30);
                    palette.CardBorder = Color.FromArgb(0x48, 0x46, 0x44);
                    palette.TextPrimary = Color.FromArgb(0xF3, 0xF2, 0xF1);
                    palette.TextMuted = Color.FromArgb(0xC8, 0xC6, 0xC4);
                    palette.AccentText = Blend(accent, Color.White, 0.46);
                    palette.AccentHover = Blend(accent, Color.White, 0.12);
                    palette.AccentPressed = Blend(accent, Color.Black, 0.16);
                    palette.SubtleHover = Color.FromArgb(0x3B, 0x3A, 0x39);
                    palette.SubtlePressed = Color.FromArgb(0x48, 0x46, 0x44);
                    palette.InputBack = Color.FromArgb(0x25, 0x24, 0x23);
                    palette.InputBorder = Color.FromArgb(0x54, 0x52, 0x50);
                    palette.Danger = Color.FromArgb(0xE8, 0x74, 0x7C);
                    break;

                case OfficeThemeKind.White:
                    palette.Name = "白色（浅色）";
                    palette.Dark = false;
                    palette.WindowBack = Color.FromArgb(0xF8, 0xF8, 0xF8);
                    palette.CardBack = Color.White;
                    palette.CardBorder = Color.FromArgb(0xE1, 0xDF, 0xDD);
                    palette.TextPrimary = Color.FromArgb(0x20, 0x1F, 0x1E);
                    palette.TextMuted = Color.FromArgb(0x60, 0x5E, 0x5C);
                    palette.AccentText = Darken(accent, 0.06);
                    palette.AccentHover = Darken(accent, 0.12);
                    palette.AccentPressed = Darken(accent, 0.28);
                    palette.SubtleHover = Color.FromArgb(0xF3, 0xF2, 0xF1);
                    palette.SubtlePressed = Color.FromArgb(0xED, 0xEB, 0xE9);
                    palette.InputBack = Color.White;
                    palette.InputBorder = Color.FromArgb(0xC8, 0xC6, 0xC4);
                    palette.Danger = Color.FromArgb(0xA4, 0x26, 0x2C);
                    break;

                default:
                    // 彩色主题（Office 默认）：画布是 Office 的浅灰 F3F2F1
                    palette.Name = "彩色（浅色）";
                    palette.Dark = false;
                    palette.WindowBack = Color.FromArgb(0xF3, 0xF2, 0xF1);
                    palette.CardBack = Color.White;
                    palette.CardBorder = Color.FromArgb(0xE1, 0xDF, 0xDD);
                    palette.TextPrimary = Color.FromArgb(0x20, 0x1F, 0x1E);
                    palette.TextMuted = Color.FromArgb(0x60, 0x5E, 0x5C);
                    palette.AccentText = Darken(accent, 0.06);
                    palette.AccentHover = Darken(accent, 0.12);
                    palette.AccentPressed = Darken(accent, 0.28);
                    palette.SubtleHover = Color.FromArgb(0xF3, 0xF2, 0xF1);
                    palette.SubtlePressed = Color.FromArgb(0xED, 0xEB, 0xE9);
                    palette.InputBack = Color.White;
                    palette.InputBorder = Color.FromArgb(0xC8, 0xC6, 0xC4);
                    palette.Danger = Color.FromArgb(0xA4, 0x26, 0x2C);
                    break;
            }
            return palette;
        }

        /// <summary>按比例向 white 混合（t=0 原色，t=1 纯白）。</summary>
        internal static Color Blend(Color color, Color target, double t)
        {
            if (t < 0)
            {
                t = 0;
            }
            if (t > 1)
            {
                t = 1;
            }
            return Color.FromArgb(
                (int)Math.Round(color.R + (target.R - color.R) * t),
                (int)Math.Round(color.G + (target.G - color.G) * t),
                (int)Math.Round(color.B + (target.B - color.B) * t));
        }

        internal static Color Darken(Color color, double t)
        {
            return Blend(color, Color.Black, t);
        }
    }

    /// <summary>
    /// 统一视觉：配色跟随 Office 主题（浅色/深色），主色取宿主应用品牌色，
    /// 保证对话框和 Office 自己的对话框、功能区是一个色系。
    /// 只依赖 System.Drawing / WinForms，不引入第三方库，VSTO 加载项里直接用。
    /// </summary>
    internal static class UiTheme
    {
        private static ThemePalette _palette =
            ThemePalette.Create(OfficeThemeKind.Colorful, OfficeTheme.DefaultAccent);

        static UiTheme()
        {
            // 进程启动后第一次用到配色的地方（通常是某个窗口的 PrepareForm）就把偏好落地
            try
            {
                UiAppearance.EnsureApplied();
            }
            catch
            {
                // 读偏好失败时保留“彩色主题 + Office 蓝”，不影响功能
            }
        }

        // ---------------- 调色板（属性而非常量：深浅主题可切换） ----------------

        public static bool IsDark
        {
            get { return _palette.Dark; }
        }

        /// <summary>当前主题名，例如“彩色（浅色）”。</summary>
        public static string ThemeName
        {
            get { return _palette.Name; }
        }

        public static Color WindowBack
        {
            get { return _palette.WindowBack; }
        }

        public static Color CardBack
        {
            get { return _palette.CardBack; }
        }

        public static Color CardBorder
        {
            get { return _palette.CardBorder; }
        }

        public static Color TextPrimary
        {
            get { return _palette.TextPrimary; }
        }

        public static Color TextMuted
        {
            get { return _palette.TextMuted; }
        }

        public static Color Accent
        {
            get { return _palette.Accent; }
        }

        /// <summary>主色当文字用时（深色主题会自动调亮）。</summary>
        public static Color AccentText
        {
            get { return _palette.AccentText; }
        }

        public static Color AccentHover
        {
            get { return _palette.AccentHover; }
        }

        public static Color AccentPressed
        {
            get { return _palette.AccentPressed; }
        }

        public static Color SubtleHover
        {
            get { return _palette.SubtleHover; }
        }

        public static Color SubtlePressed
        {
            get { return _palette.SubtlePressed; }
        }

        public static Color InputBack
        {
            get { return _palette.InputBack; }
        }

        public static Color InputBorder
        {
            get { return _palette.InputBorder; }
        }

        public static Color Danger
        {
            get { return _palette.Danger; }
        }

        /// <summary>直接指定主题与主色（设置界面的“立即预览”、自测、截图都用它）。</summary>
        public static void Apply(OfficeThemeKind kind, Color accent)
        {
            _palette = ThemePalette.Create(kind, accent);
        }

        /// <summary>恢复“跟随 Office”。</summary>
        public static void ApplyAutomatic()
        {
            Apply(OfficeTheme.Kind, OfficeTheme.HostAccent);
        }

        // ---------------- 字体 ----------------

        private static string _family;

        /// <summary>优先中文字体，保证中文界面不发虚（Segoe UI 无中文字形时会回退）。</summary>
        public static string FontFamily
        {
            get
            {
                if (_family == null)
                {
                    _family = PickFamily();
                }
                return _family;
            }
        }

        private static string PickFamily()
        {
            var candidates = new[] { "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "Tahoma" };
            foreach (var name in candidates)
            {
                try
                {
                    using (var family = new FontFamily(name))
                    {
                        if (string.Equals(family.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            return name;
                        }
                    }
                }
                catch
                {
                    // 该字体不存在，试下一个
                }
            }
            return SystemFonts.MessageBoxFont.FontFamily.Name;
        }

        public static Font Body()
        {
            return new Font(FontFamily, 9.5f, FontStyle.Regular, GraphicsUnit.Point);
        }

        public static Font BodyBold()
        {
            return new Font(FontFamily, 9.5f, FontStyle.Bold, GraphicsUnit.Point);
        }

        public static Font Title()
        {
            return new Font(FontFamily, 14f, FontStyle.Bold, GraphicsUnit.Point);
        }

        public static Font Subtitle()
        {
            return new Font(FontFamily, 10.5f, FontStyle.Regular, GraphicsUnit.Point);
        }

        public static Font Small()
        {
            return new Font(FontFamily, 8.5f, FontStyle.Regular, GraphicsUnit.Point);
        }

        /// <summary>等宽字体（差异摘要正文）。Consolas 不含中文字形，会由系统回退，这里给中文备用字体。</summary>
        public static Font Mono()
        {
            try
            {
                return new Font("Consolas", 9.5f, FontStyle.Regular, GraphicsUnit.Point);
            }
            catch
            {
                return new Font(FontFamily, 9.5f, FontStyle.Regular, GraphicsUnit.Point);
            }
        }

        // ---------------- 控件 ----------------

        /// <summary>窗体统一外观：按 Office 主题取底色、统一字体、按 DPI 缩放。</summary>
        public static void PrepareForm(Form form, string title)
        {
            try
            {
                UiAppearance.EnsureApplied();
            }
            catch
            {
                // 忽略：保持当前调色板
            }

            form.Text = title;
            form.Font = Body();
            form.BackColor = WindowBack;
            form.ForeColor = TextPrimary;
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.AutoScaleDimensions = new SizeF(96F, 96F);
            form.StartPosition = FormStartPosition.CenterParent;
            form.ShowInTaskbar = false;
            form.Padding = new Padding(0);

            NativeTheme.Apply(form);
        }

        /// <summary>主按钮（实心品牌色）或次按钮（卡片底 + 描边）。AutoSize + 内边距，文字永远不会被按钮框裁掉。</summary>
        public static Button CreateButton(string text, bool primary)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlatStyle = FlatStyle.Flat,
                Font = Body(),
                Padding = new Padding(16, 7, 16, 7),
                MinimumSize = new Size(92, 34),
                Margin = new Padding(8, 0, 0, 0),
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = false,
            };
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = primary ? Accent : InputBorder;
            if (primary)
            {
                button.BackColor = Accent;
                button.ForeColor = Color.White;
                button.FlatAppearance.MouseOverBackColor = AccentHover;
                button.FlatAppearance.MouseDownBackColor = AccentPressed;
            }
            else
            {
                button.BackColor = CardBack;
                button.ForeColor = TextPrimary;
                button.FlatAppearance.MouseOverBackColor = SubtleHover;
                button.FlatAppearance.MouseDownBackColor = SubtlePressed;
            }
            return button;
        }

        public static Label CreateLabel(string text, Font font, Color color)
        {
            return new Label
            {
                Text = text,
                Font = font,
                ForeColor = color,
                AutoSize = true,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
            };
        }

        /// <summary>输入控件统一字体与边距；深色主题下顺带把输入区也换成深色。</summary>
        public static void StyleInput(Control control)
        {
            if (control == null)
            {
                return;
            }
            control.Font = Body();
            control.Margin = new Padding(0, 6, 0, 6);
            if (IsDark)
            {
                ApplyDarkInput(control);
            }
        }

        /// <summary>
        /// WinForms 的标准控件默认是浅色外观，深色主题下直接铺上去会“白一块”。
        /// 这里按控件类型分别处理：文本框改底色字色，下拉框自己画列表项，微调框连内部编辑框一起改。
        /// </summary>
        private static void ApplyDarkInput(Control control)
        {
            var combo = control as ComboBox;
            if (combo != null)
            {
                combo.BackColor = InputBack;
                combo.ForeColor = TextPrimary;
                if (combo.DrawMode != DrawMode.OwnerDrawFixed)
                {
                    combo.DrawMode = DrawMode.OwnerDrawFixed;
                    combo.FlatStyle = FlatStyle.Flat;
                    combo.ItemHeight = Math.Max(18, (int)Math.Ceiling(combo.Font.GetHeight()) + 6);
                    combo.DrawItem += DrawDarkComboItem;
                }
                return;
            }

            control.BackColor = InputBack;
            control.ForeColor = TextPrimary;

            // NumericUpDown 的输入区是它内部的子控件
            foreach (Control child in control.Controls)
            {
                ApplyDarkInput(child);
            }
        }

        private static void DrawDarkComboItem(object sender, DrawItemEventArgs e)
        {
            var combo = sender as ComboBox;
            if (combo == null)
            {
                return;
            }
            var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using (var brush = new SolidBrush(selected ? SubtlePressed : InputBack))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }
            if (e.Index >= 0 && e.Index < combo.Items.Count)
            {
                var text = combo.GetItemText(combo.Items[e.Index]);
                TextRenderer.DrawText(e.Graphics, text, combo.Font, e.Bounds, TextPrimary,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }
        }

        public static CheckBox CreateCheckBox(string text)
        {
            return new CheckBox
            {
                Text = text,
                Font = Body(),
                ForeColor = TextPrimary,
                AutoSize = true,
                // 深色下系统主题复选框会带一块浅色底，改用扁平样式自己着色
                FlatStyle = IsDark ? FlatStyle.Flat : FlatStyle.System,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 7, 0, 7),
                Cursor = Cursors.Hand,
            };
        }

        /// <summary>顶部标题区：主标题 + 副标题（灰色说明）。</summary>
        public static Control CreateHeader(string title, string subtitle)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
                Margin = new Padding(2, 0, 2, 12),
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            panel.Controls.Add(CreateLabel(title, Title(), TextPrimary), 0, 0);
            var sub = CreateLabel(subtitle, Subtitle(), TextMuted);
            sub.Margin = new Padding(0, 4, 0, 0);
            sub.AutoEllipsis = false;
            panel.Controls.Add(sub, 0, 1);
            return panel;
        }

        /// <summary>
        /// 卡片容器：圆角卡片 + 细描边。继承 TableLayoutPanel 是为了让高度真正由内容决定（AutoSize），
        /// 避免“固定高度的面板把按钮/文字裁掉”。用 <see cref="AddContent"/> 放内容。
        /// </summary>
        public sealed class CardPanel : TableLayoutPanel
        {
            private int _cornerRadius = 10;

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
                BackColor = WindowBack;
                Padding = new Padding(18, 12, 18, 16);
                Margin = new Padding(0, 0, 0, 12);
                ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
                RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            public int CornerRadius
            {
                get { return _cornerRadius; }
                set
                {
                    _cornerRadius = value < 0 ? 0 : value;
                    Invalidate();
                }
            }

            /// <summary>放入内容（自动 Dock=Top，宽度跟随卡片）。</summary>
            public void AddContent(Control content)
            {
                if (content == null)
                {
                    return;
                }
                content.Dock = DockStyle.Top;
                Controls.Add(content, 0, 0);
            }

            protected override void OnPaintBackground(PaintEventArgs e)
            {
                // 卡片外的角落用窗体底色填满，圆角才干净
                using (var brush = new SolidBrush(WindowBack))
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
                using (var path = CreateRoundedRect(bounds, _cornerRadius))
                using (var fill = new SolidBrush(CardBack))
                using (var border = new Pen(CardBorder))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                }
            }

            private static GraphicsPath CreateRoundedRect(Rectangle bounds, int radius)
            {
                var path = new GraphicsPath();
                if (radius <= 0)
                {
                    path.AddRectangle(bounds);
                    return path;
                }
                var d = radius * 2;
                path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }
    }
}
