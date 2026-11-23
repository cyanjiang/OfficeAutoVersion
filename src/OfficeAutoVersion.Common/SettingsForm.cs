using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeAutoVersion.Core;

namespace OfficeAutoVersion.Common
{
    /// <summary>
    /// 自动版本设置对话框。配置按应用键保存到 %APPDATA%\OfficeAutoVersion\settings.json。
    /// 布局全部交给 TableLayoutPanel 的 AutoSize：窗口高度按内容算，按钮用 AutoSize + 内边距，
    /// 因此在高 DPI / 大字体 / 长路径下都不会出现“文字被遮住”“按钮显示不全”。
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly string _appKey;
        private readonly string _currentDocumentPath;
        private VersionOptions _options;

        private CheckBox _chkEnabled;
        private ComboBox _cmbStorageMode;
        private TextBox _txtFolderName;
        private TextBox _txtPattern;
        private NumericUpDown _numKeep;
        private CheckBox _chkSkipUnchanged;
        private CheckBox _chkLog;
        private ComboBox _cmbAppearance;
        private Label _lblPreview;
        private Label _lblPaths;

        private TableLayoutPanel _root;
        private bool _fitted;

        public SettingsForm(string appKey, string currentDocumentPath = null)
        {
            _appKey = appKey;
            _currentDocumentPath = currentDocumentPath;
            _options = OptionsStore.Load(appKey);
            BuildUi();
            BindToOptions();
        }

        private void BuildUi()
        {
            UiTheme.PrepareForm(this, "自动版本设置 - " + _appKey);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(780, 720);
            MinimumSize = new Size(700, 520);
            DoubleBuffered = true;

            var toolTip = new ToolTip
            {
                AutoPopDelay = 8000,
                InitialDelay = 250,
                ReshowDelay = 200,
                ShowAlways = true,
            };

            _root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = UiTheme.WindowBack,
                Padding = new Padding(18, 16, 18, 14),
            };
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            _root.Controls.Add(UiTheme.CreateHeader("自动版本设置",
                "保存文档时自动留下历史版本，可随时用 Office 自带的比较、恢复功能回到任意一版。"));
            _root.Controls.Add(BuildOptionsCard(toolTip));
            _root.Controls.Add(BuildPreviewCard());
            _root.Controls.Add(BuildInfoLabel());
            _root.Controls.Add(BuildFooter());
            // 显式声明每行都是 AutoSize：行样式缺省时高度可能按 0 处理，会把底部按钮压没
            _root.RowCount = _root.Controls.Count;
            for (var i = 0; i < _root.RowCount; i++)
            {
                _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            Controls.Add(_root);

            UpdatePreview();
        }

        /// <summary>“常规设置”卡片：所有输入项两列布局，标签列自适应宽度。</summary>
        private Control BuildOptionsCard(ToolTip toolTip)
        {
            var card = new UiTheme.CardPanel();
            var table = CreateFieldTable();

            _chkEnabled = UiTheme.CreateCheckBox("启用自动版本（每次保存时自动建立版本）");
            AddSpanningRow(table, _chkEnabled);

            _cmbStorageMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                FlatStyle = FlatStyle.System,
            };
            UiTheme.StyleInput(_cmbStorageMode);
            _cmbStorageMode.Items.AddRange(new object[]
            {
                "原文件旁边（文件夹随文件走，推荐）",
                "统一目录（文档\\OfficeAutoVersions，镜像原路径）",
            });
            _cmbStorageMode.SelectedIndexChanged += delegate { UpdatePreview(); };
            AddFieldRow(table, "保存位置", _cmbStorageMode);
            toolTip.SetToolTip(_cmbStorageMode, "“原文件旁边”便于随文档一起拷贝/同步；\n“统一目录”可以让原文件夹保持干净。");

            _txtFolderName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            UiTheme.StyleInput(_txtFolderName);
            _txtFolderName.TextChanged += delegate { UpdatePreview(); };
            AddFieldRow(table, "版本文件夹名称", _txtFolderName);
            toolTip.SetToolTip(_txtFolderName, "可用 {name} 表示原文件名（不含扩展名）。\n例如：{name}_Versions");

            _txtPattern = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right };
            UiTheme.StyleInput(_txtPattern);
            _txtPattern.TextChanged += delegate { UpdatePreview(); };
            AddFieldRow(table, "版本文件名格式", _txtPattern);
            toolTip.SetToolTip(_txtPattern,
                "{name} 原文件名；{ts} 时间戳(yyyyMMdd-HHmmss)；\n也可以直接写 .NET 日期格式，如 {yyyyMMdd-HHmmss}。\n例：{name}_{ts} →  报告_20261025-235501.docx");

            _numKeep = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 999,
                Width = 96,
                Anchor = AnchorStyles.Left,
            };
            UiTheme.StyleInput(_numKeep);
            AddFieldRow(table, "最多保留版本数", _numKeep);
            toolTip.SetToolTip(_numKeep, "0 表示不限制。\n超出的旧版本会被移入回收站（可恢复）。");

            _chkSkipUnchanged = UiTheme.CreateCheckBox("内容没有变化时跳过，避免生成重复快照");
            AddSpanningRow(table, _chkSkipUnchanged);

            _chkLog = UiTheme.CreateCheckBox("记录运行日志");
            AddSpanningRow(table, _chkLog);
            toolTip.SetToolTip(_chkLog, "日志位置：%APPDATA%\\OfficeAutoVersion\\logs\\autoversion.log");

            _cmbAppearance = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                FlatStyle = FlatStyle.System,
            };
            UiTheme.StyleInput(_cmbAppearance);
            _cmbAppearance.Items.AddRange(new object[]
            {
                "跟随 Office 主题（推荐）",
                "始终浅色",
                "始终深色",
            });
            AddFieldRow(table, "界面配色", _cmbAppearance);
            toolTip.SetToolTip(_cmbAppearance,
                "配色取自“文件 - 选项 - 常规 - Office 主题”，主色用本应用的品牌色：\n" +
                "Word 蓝、Excel 绿、PowerPoint 橙红。\n修改后重新打开窗口生效。");

            card.AddContent(table);
            return card;
        }

        /// <summary>“保存位置预览”卡片：长路径自动换行完整显示。</summary>
        private Control BuildPreviewCard()
        {
            var card = new UiTheme.CardPanel();
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var caption = UiTheme.CreateLabel("保存位置预览（基于当前文档计算）：", UiTheme.BodyBold(), UiTheme.TextPrimary);
            caption.Margin = new Padding(0, 0, 0, 6);
            table.Controls.Add(caption);

            _lblPreview = UiTheme.CreateLabel("", UiTheme.Body(), UiTheme.AccentText);
            _lblPreview.AutoSize = true;
            _lblPreview.MaximumSize = new Size(640, 0);
            table.Controls.Add(_lblPreview);
            table.RowCount = table.Controls.Count;
            for (var i = 0; i < table.RowCount; i++)
            {
                table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            card.AddContent(table);
            return card;
        }

        private Control BuildInfoLabel()
        {
            _lblPaths = UiTheme.CreateLabel(
                "设置文件：" + OptionsStore.SettingsFilePath + "\n日志目录：" + Logger.Default.LogPath,
                UiTheme.Small(), UiTheme.TextMuted);
            _lblPaths.AutoSize = true;
            _lblPaths.MaximumSize = new Size(720, 0);
            _lblPaths.Margin = new Padding(6, 2, 6, 10);
            return _lblPaths;
        }

        private Control BuildFooter()
        {
            var footer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 2, 0, 0),
            };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.Transparent,
            };

            var btnOk = UiTheme.CreateButton("确定", true);
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Click += OnOkClick;

            var btnCancel = UiTheme.CreateButton("取消", false);
            btnCancel.DialogResult = DialogResult.Cancel;

            var btnReset = UiTheme.CreateButton("恢复默认", false);
            btnReset.Click += delegate
            {
                _options = new VersionOptions();
                BindToOptions();
            };

            var btnOpenLog = UiTheme.CreateButton("打开日志", false);
            btnOpenLog.Click += delegate { OpenLogFolder(); };

            buttons.Controls.Add(btnOk);
            buttons.Controls.Add(btnCancel);
            buttons.Controls.Add(btnReset);
            buttons.Controls.Add(btnOpenLog);

            footer.Controls.Add(buttons);
            AcceptButton = btnOk;
            CancelButton = btnCancel;
            return footer;
        }

        private static TableLayoutPanel CreateFieldTable()
        {
            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.Transparent,
            };
            // 标签列用 AutoSize：中文标签按实际文字宽度分配，不会被固定的 140px 裁掉
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            return table;
        }

        private static void AddFieldRow(TableLayoutPanel table, string label, Control control)
        {
            var caption = UiTheme.CreateLabel(label, UiTheme.Body(), UiTheme.TextMuted);
            caption.Anchor = AnchorStyles.Left;
            caption.Margin = new Padding(0, 9, 18, 9);
            var row = table.RowCount;
            table.RowCount = row + 1;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(caption, 0, row);
            table.Controls.Add(control, 1, row);
        }

        private static void AddSpanningRow(TableLayoutPanel table, Control control)
        {
            var row = table.RowCount;
            table.RowCount = row + 1;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(control, 0, row);
            table.SetColumnSpan(control, 2);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            // 首个复选框默认带着虚线焦点框，观感差；把焦点交还给窗体
            ActiveControl = null;
            UpdateWrapWidth();
            if (!_fitted)
            {
                _fitted = true;
                FitToContent();
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateWrapWidth();
        }

        /// <summary>让两段可能很长的文字按当前窗口宽度换行，避免路径被截断看不全。</summary>
        private void UpdateWrapWidth()
        {
            if (_lblPreview == null || _lblPaths == null || _root == null)
            {
                return;
            }
            var width = Math.Max(240, ClientSize.Width - _root.Padding.Horizontal - 60);
            var size = new Size(width, 0);
            if (_lblPreview.MaximumSize != size)
            {
                _lblPreview.MaximumSize = size;
            }
            if (_lblPaths.MaximumSize != size)
            {
                _lblPaths.MaximumSize = size;
            }
        }

        /// <summary>
        /// 首次显示时按内容实际需要设置窗口大小，并把最小尺寸卡在“刚好放得下”。
        /// 这样按钮和底部信息永远在同一屏内，也不会留下大片空白。
        /// </summary>
        private void FitToContent()
        {
            try
            {
                var preferred = _root.GetPreferredSize(new Size(Math.Max(640, ClientSize.Width), 0));
                var chrome = Size - ClientSize;
                var width = Math.Max(ClientSize.Width, preferred.Width + chrome.Width);
                var height = preferred.Height + chrome.Height + 2;
                MinimumSize = new Size(width, height);
                ClientSize = new Size(width - chrome.Width, height - chrome.Height);
            }
            catch
            {
                // 尺寸兜底失败不应影响设置功能
            }
        }

        private void OpenLogFolder()
        {
            try
            {
                var directory = Path.GetDirectoryName(Logger.Default.LogPath);
                if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                {
                    Process.Start("explorer.exe", "\"" + directory + "\"");
                }
                else
                {
                    MessageBox.Show(this, "日志目录还不存在（尚未产生任何日志）。", "自动版本",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "无法打开日志目录：" + ex.Message, "自动版本",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BindToOptions()
        {
            _chkEnabled.Checked = _options.Enabled;
            _cmbStorageMode.SelectedIndex = _options.StorageMode == VersionStorageMode.CentralStore ? 1 : 0;
            _txtFolderName.Text = _options.VersionFolderName;
            _txtPattern.Text = _options.FileNamePattern;
            var keep = _options.MaxVersionsToKeep;
            if (keep < 0)
            {
                keep = 0;
            }
            if (keep > 999)
            {
                keep = 999;
            }
            _numKeep.Value = keep;
            _chkSkipUnchanged.Checked = _options.SkipIfUnchanged;
            _chkLog.Checked = _options.LogEnabled;
            _cmbAppearance.SelectedIndex = (int)LoadAppearancePreference();
            UpdatePreview();
        }

        /// <summary>读取配色偏好；读不到按“跟随 Office”。</summary>
        private static ThemePreference LoadAppearancePreference()
        {
            try
            {
                return UiAppearance.Load();
            }
            catch
            {
                return ThemePreference.Automatic;
            }
        }

        /// <summary>下拉框当前选择对应的配色偏好（索引与枚举值一致：0 跟随 / 1 浅色 / 2 深色）。</summary>
        private ThemePreference SelectedAppearance()
        {
            if (_cmbAppearance == null)
            {
                return ThemePreference.Automatic;
            }
            switch (_cmbAppearance.SelectedIndex)
            {
                case 1:
                    return ThemePreference.Light;
                case 2:
                    return ThemePreference.Dark;
                default:
                    return ThemePreference.Automatic;
            }
        }

        private VersionOptions ReadOptionsFromControls()
        {
            var options = _options.Clone();
            options.Enabled = _chkEnabled.Checked;
            options.StorageMode = _cmbStorageMode.SelectedIndex == 1 ? VersionStorageMode.CentralStore : VersionStorageMode.NextToFile;
            options.VersionFolderName = _txtFolderName.Text.Trim();
            options.FileNamePattern = _txtPattern.Text.Trim();
            options.MaxVersionsToKeep = (int)_numKeep.Value;
            options.SkipIfUnchanged = _chkSkipUnchanged.Checked;
            options.LogEnabled = _chkLog.Checked;
            return options;
        }

        private void UpdatePreview()
        {
            if (_lblPreview == null)
            {
                return;
            }
            try
            {
                if (string.IsNullOrEmpty(_currentDocumentPath))
                {
                    _lblPreview.Text = "（当前没有打开的文档，保存任意文档后将按上面的设置生成版本）";
                    return;
                }
                _lblPreview.Text = VersionPaths.GetVersionsDirectory(_currentDocumentPath, ReadOptionsFromControls());
            }
            catch
            {
                _lblPreview.Text = "（无法预览：当前设置包含无效字符）";
            }
        }

        private void OnOkClick(object sender, EventArgs e)
        {
            var folderName = _txtFolderName.Text.Trim();
            if (folderName.Length == 0)
            {
                MessageBox.Show(this, "版本文件夹名称不能为空。", "自动版本", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }
            if (folderName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                MessageBox.Show(this, "版本文件夹名称包含非法字符。", "自动版本", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }
            var pattern = _txtPattern.Text.Trim();
            if (pattern.Length == 0)
            {
                MessageBox.Show(this, "版本文件名格式不能为空。", "自动版本", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            _options = ReadOptionsFromControls();
            try
            {
                OptionsStore.Save(_appKey, _options);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "保存设置失败：" + ex.Message, "自动版本", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            // 配色偏好是全局的，单独存一份；写失败不影响已保存的版本设置
            try
            {
                UiAppearance.Save(SelectedAppearance());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "版本设置已保存，但配色偏好写入失败（下次打开窗口仍是原配色）：" + ex.Message,
                    "自动版本", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
