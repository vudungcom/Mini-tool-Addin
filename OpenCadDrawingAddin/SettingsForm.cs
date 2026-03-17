using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OpenCadDrawingAddin.Logic;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// Form Settings cho Open CAD Drawing Add-in
    /// Tab 1 - Settings: Chọn thư mục CAD + tùy chọn số sửa đổi
    /// Tab 2 - Language: Chọn ngôn ngữ hiển thị
    /// Version 1.0
    /// </summary>
    public class SettingsForm : Form
    {
        // === Controls ===
        private TabControl mainTabControl;
        private TabPage tabSettings;
        private TabPage tabLanguage;

        // Tab Settings
        private Label lblCadFolder;
        private TextBox txtCadFolder;
        private Button btnBrowse;
        private Label lblFolderHint;
        private CheckBox chkUseRevision;
        private Label lblRevisionHint;
        private Label lblExtension;
        private TextBox txtExtension;
        private Label lblExtensionHint;

        // Tab Language
        private ListBox lstLanguages;

        // Bottom buttons
        private Button btnSave;
        private Button btnCancel;

        private OpenCadSettings _settings;

        public SettingsForm()
        {
            _settings = OpenCadSettings.Load();
            InitializeComponent();
            LoadSettingsToUI();
        }

        private void InitializeComponent()
        {
            // DPI scaling
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi;

            this.Text = LanguageManager.L("SETTINGS_TITLE");
            this.Size = new Size(480, 380);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Font = new Font("Segoe UI", 9F);

            // ============ TAB CONTROL ============
            mainTabControl = new TabControl
            {
                Location = new Point(10, 10),
                Size = new Size(445, 290)
            };

            // --- Tab 1: Settings ---
            tabSettings = new TabPage(LanguageManager.L("TAB_SETTINGS"));
            BuildSettingsTab();
            mainTabControl.TabPages.Add(tabSettings);

            // --- Tab 2: Language ---
            tabLanguage = new TabPage(LanguageManager.L("TAB_LANGUAGE"));
            BuildLanguageTab();
            mainTabControl.TabPages.Add(tabLanguage);

            this.Controls.Add(mainTabControl);

            // ============ BOTTOM BUTTONS ============
            btnSave = new Button
            {
                Text = LanguageManager.L("BTN_SAVE"),
                Location = new Point(270, 310),
                Size = new Size(90, 30),
                BackColor = Color.LightGreen
            };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            btnCancel = new Button
            {
                Text = LanguageManager.L("BTN_CANCEL"),
                Location = new Point(370, 310),
                Size = new Size(85, 30)
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);

            this.CancelButton = btnCancel;
            this.AcceptButton = btnSave;
        }

        // ============================================================
        // TAB 1: SETTINGS
        // ============================================================
        private void BuildSettingsTab()
        {
            int x = 15, y = 15;

            // --- CAD Folder ---
            lblCadFolder = new Label
            {
                Text = LanguageManager.L("LBL_CAD_FOLDER"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabSettings.Controls.Add(lblCadFolder);
            y += 20;

            txtCadFolder = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 23),
                // PlaceholderText không có trong .NET 4.8 - bỏ qua
            };
            tabSettings.Controls.Add(txtCadFolder);

            btnBrowse = new Button
            {
                Text = LanguageManager.L("BTN_BROWSE"),
                Location = new Point(x + 325, y - 1),
                Size = new Size(85, 25)
            };
            btnBrowse.Click += BtnBrowse_Click;
            tabSettings.Controls.Add(btnBrowse);
            y += 25;

            lblFolderHint = new Label
            {
                Text = LanguageManager.L("LBL_FOLDER_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabSettings.Controls.Add(lblFolderHint);
            y += 30;

            // --- Revision Suffix ---
            chkUseRevision = new CheckBox
            {
                Text = LanguageManager.L("CHK_USE_REVISION"),
                Location = new Point(x, y),
                Size = new Size(400, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabSettings.Controls.Add(chkUseRevision);
            y += 25;

            lblRevisionHint = new Label
            {
                Text = LanguageManager.L("CHK_USE_REVISION_HINT"),
                Location = new Point(x + 20, y),
                Size = new Size(390, 40),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabSettings.Controls.Add(lblRevisionHint);
            y += 50;

            // --- File Extension ---
            lblExtension = new Label
            {
                Text = LanguageManager.L("LBL_EXTENSION"),
                Location = new Point(x, y),
                Size = new Size(200, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabSettings.Controls.Add(lblExtension);
            y += 20;

            txtExtension = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(80, 23),
                MaxLength = 10
            };
            tabSettings.Controls.Add(txtExtension);

            lblExtensionHint = new Label
            {
                Text = LanguageManager.L("LBL_EXTENSION_HINT"),
                Location = new Point(x + 90, y + 3),
                Size = new Size(200, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabSettings.Controls.Add(lblExtensionHint);
        }

        // ============================================================
        // TAB 2: LANGUAGE
        // ============================================================
        private void BuildLanguageTab()
        {
            var lblPrompt = new Label
            {
                Text = "Select display language:",
                Location = new Point(15, 15),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabLanguage.Controls.Add(lblPrompt);

            lstLanguages = new ListBox
            {
                Location = new Point(15, 38),
                Size = new Size(200, 200),
                Font = new Font("Segoe UI", 10F)
            };
            foreach (var kvp in LanguageManager.SupportedLanguages)
            {
                lstLanguages.Items.Add($"{kvp.Key} - {kvp.Value}");
            }
            tabLanguage.Controls.Add(lstLanguages);
        }

        // ============================================================
        // LOAD / SAVE
        // ============================================================
        private void LoadSettingsToUI()
        {
            txtCadFolder.Text = _settings.CadFolderPath;
            chkUseRevision.Checked = _settings.UseRevisionSuffix;
            txtExtension.Text = _settings.CadExtension.TrimStart('.');
            if (!txtExtension.Text.StartsWith("."))
                txtExtension.Text = "." + txtExtension.Text;

            // Select current language
            string currentLang = _settings.Language;
            for (int i = 0; i < lstLanguages.Items.Count; i++)
            {
                string item = lstLanguages.Items[i].ToString();
                if (item.StartsWith(currentLang + " "))
                {
                    lstLanguages.SelectedIndex = i;
                    break;
                }
            }
            if (lstLanguages.SelectedIndex < 0 && lstLanguages.Items.Count > 0)
                lstLanguages.SelectedIndex = 0;
        }

        private void SaveUIToSettings()
        {
            _settings.CadFolderPath = txtCadFolder.Text.Trim();

            // Normalize extension
            string ext = txtExtension.Text.Trim();
            if (!string.IsNullOrEmpty(ext) && !ext.StartsWith("."))
                ext = "." + ext;
            _settings.CadExtension = string.IsNullOrEmpty(ext) ? ".dwg" : ext;

            _settings.UseRevisionSuffix = chkUseRevision.Checked;

            // Language
            if (lstLanguages.SelectedItem != null)
            {
                string sel = lstLanguages.SelectedItem.ToString();
                string langCode = sel.Split('-')[0].Trim();
                _settings.Language = langCode;
                LanguageManager.CurrentLanguage = langCode;
            }
        }

        // ============================================================
        // EVENT HANDLERS
        // ============================================================
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select CAD Drawing Folder";
                dialog.ShowNewFolderButton = false;

                if (!string.IsNullOrEmpty(txtCadFolder.Text) && Directory.Exists(txtCadFolder.Text))
                    dialog.SelectedPath = txtCadFolder.Text;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtCadFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Validate folder
            string folder = txtCadFolder.Text.Trim();
            if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
            {
                var result = MessageBox.Show(
                    $"Folder does not exist:\n{folder}\n\nSave anyway?",
                    LanguageManager.L("TITLE_WARNING"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result == DialogResult.No) return;
            }

            SaveUIToSettings();
            _settings.Save();

            MessageBox.Show(
                LanguageManager.L("MSG_SETTINGS_SAVED"),
                LanguageManager.L("TITLE_INFO"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            this.Close();
        }
    }
}
