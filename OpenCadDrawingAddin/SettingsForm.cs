using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OpenCadDrawingAddin.Logic;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// Form Settings cho Mini Tool Add-in
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
        private TabPage tabBom;       // [NEW]
        private TabPage tabCheckRef;  // [NEW]
        private TabPage tabCreateDwg; // [NEW v1.2]

        // Tab Settings - [CHANGED v1.1] multi-folder
        private Label lblCadFolder;
        private ListBox lstCadFolders;   // [CHANGED] thay txtCadFolder
        private Button btnAddFolder;     // [NEW] thêm folder
        private Button btnRemoveFolder;  // [NEW] xóa folder
        private Label lblFolderHint;
        private CheckBox chkUseRevision;
        private Label lblRevisionHint;
        private Label lblExtension;
        private TextBox txtExtension;
        private Label lblExtensionHint;

        // Tab Bom format [NEW]
        private TextBox txtBomXmlPath;
        private Button btnBomBrowse;

        // Tab Check Reference [NEW]
        private TextBox txtCheckRefExcludePath;
        private Button btnCheckRefBrowse;

        // Tab Create DWG [NEW v1.2]
        private TextBox txtCreateDwgIniPath;
        private Button btnCreateDwgIniBrowse;
        private TextBox txtCreateDwgListPath;
        private Button btnCreateDwgListBrowse;
        private TextBox txtCreateDwgOutputPath;
        private Button btnCreateDwgOutputBrowse;

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

            // --- Tab 1: Settings (renamed to Cad drawing) ---
            tabSettings = new TabPage(LanguageManager.L("TAB_SETTINGS")); // [CHANGED] key TAB_SETTINGS giờ = "Cad drawing"
            BuildSettingsTab();
            mainTabControl.TabPages.Add(tabSettings);

            // --- Tab 2: Bom format ---
            tabBom = new TabPage(LanguageManager.L("TAB_BOM"));
            BuildBomTab();
            mainTabControl.TabPages.Add(tabBom);

            // --- Tab 3: Check Reference ---
            tabCheckRef = new TabPage(LanguageManager.L("TAB_CHECK_REF"));
            BuildCheckRefTab();
            mainTabControl.TabPages.Add(tabCheckRef);

            // --- Tab 4: Create DWG --- [NEW v1.2]
            tabCreateDwg = new TabPage(LanguageManager.L("TAB_CREATE_DWG"));
            BuildCreateDwgTab();
            mainTabControl.TabPages.Add(tabCreateDwg);

            // --- Tab 5: Language --- [MOVED] ra ngoài cùng phải
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

            // [CHANGED v1.1] ListBox thay cho TextBox - hỗ trợ nhiều folder
            lstCadFolders = new ListBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 70),
                Font = new Font("Segoe UI", 8.5F),
                SelectionMode = SelectionMode.One
            };
            tabSettings.Controls.Add(lstCadFolders);

            btnAddFolder = new Button
            {
                Text = "+",
                Location = new Point(x + 325, y),
                Size = new Size(85, 25),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                BackColor = Color.LightGreen
            };
            btnAddFolder.Click += BtnAddFolder_Click;
            tabSettings.Controls.Add(btnAddFolder);

            btnRemoveFolder = new Button
            {
                Text = "−",
                Location = new Point(x + 325, y + 30),
                Size = new Size(85, 25),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnRemoveFolder.Click += BtnRemoveFolder_Click;
            tabSettings.Controls.Add(btnRemoveFolder);
            y += 75;

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
        // TAB 2: BOM FORMAT [NEW]
        // ============================================================
        private void BuildBomTab()
        {
            int x = 15, y = 15;

            var lblXmlPath = new Label
            {
                Text = LanguageManager.L("LBL_BOM_XML_PATH"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabBom.Controls.Add(lblXmlPath);
            y += 20;

            txtBomXmlPath = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 23)
            };
            tabBom.Controls.Add(txtBomXmlPath);

            btnBomBrowse = new Button
            {
                Text = LanguageManager.L("BTN_BROWSE"),
                Location = new Point(x + 325, y - 1),
                Size = new Size(85, 25)
            };
            btnBomBrowse.Click += BtnBomBrowse_Click;
            tabBom.Controls.Add(btnBomBrowse);
            y += 25;

            var lblHint = new Label
            {
                Text = LanguageManager.L("LBL_BOM_XML_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabBom.Controls.Add(lblHint);
        }

        // ============================================================
        // TAB 3: CHECK REFERENCE [NEW]
        // ============================================================
        private void BuildCheckRefTab()
        {
            int x = 15, y = 15;

            var lblExclude = new Label
            {
                Text = LanguageManager.L("LBL_CHECK_REF_EXCLUDE"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabCheckRef.Controls.Add(lblExclude);
            y += 20;

            txtCheckRefExcludePath = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 23)
            };
            tabCheckRef.Controls.Add(txtCheckRefExcludePath);

            btnCheckRefBrowse = new Button
            {
                Text = LanguageManager.L("BTN_BROWSE"),
                Location = new Point(x + 325, y - 1),
                Size = new Size(85, 25)
            };
            btnCheckRefBrowse.Click += BtnCheckRefBrowse_Click;
            tabCheckRef.Controls.Add(btnCheckRefBrowse);
            y += 25;

            var lblHint = new Label
            {
                Text = LanguageManager.L("LBL_CHECK_REF_EXCLUDE_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabCheckRef.Controls.Add(lblHint);
        }

        // ============================================================
        // TAB 4: CREATE DWG [NEW v1.2]
        // ============================================================
        private void BuildCreateDwgTab()
        {
            int x = 15, y = 15;

            // --- DWGExport.ini path ---
            var lblIni = new Label
            {
                Text = LanguageManager.L("LBL_CREATE_DWG_INI"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabCreateDwg.Controls.Add(lblIni);
            y += 20;

            txtCreateDwgIniPath = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 23)
            };
            tabCreateDwg.Controls.Add(txtCreateDwgIniPath);

            btnCreateDwgIniBrowse = new Button
            {
                Text = LanguageManager.L("BTN_BROWSE"),
                Location = new Point(x + 325, y - 1),
                Size = new Size(85, 25)
            };
            btnCreateDwgIniBrowse.Click += BtnCreateDwgIniBrowse_Click;
            tabCreateDwg.Controls.Add(btnCreateDwgIniBrowse);
            y += 25;

            var lblIniHint = new Label
            {
                Text = LanguageManager.L("LBL_CREATE_DWG_INI_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabCreateDwg.Controls.Add(lblIniHint);
            y += 30;

            // --- List file path ---
            var lblList = new Label
            {
                Text = LanguageManager.L("LBL_CREATE_DWG_LIST"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabCreateDwg.Controls.Add(lblList);
            y += 20;

            txtCreateDwgListPath = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 23)
            };
            tabCreateDwg.Controls.Add(txtCreateDwgListPath);

            btnCreateDwgListBrowse = new Button
            {
                Text = LanguageManager.L("BTN_BROWSE"),
                Location = new Point(x + 325, y - 1),
                Size = new Size(85, 25)
            };
            btnCreateDwgListBrowse.Click += BtnCreateDwgListBrowse_Click;
            tabCreateDwg.Controls.Add(btnCreateDwgListBrowse);
            y += 25;

            var lblListHint = new Label
            {
                Text = LanguageManager.L("LBL_CREATE_DWG_LIST_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabCreateDwg.Controls.Add(lblListHint);
            y += 30;

            // --- Output folder ---
            var lblOutput = new Label
            {
                Text = LanguageManager.L("LBL_CREATE_DWG_OUTPUT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabCreateDwg.Controls.Add(lblOutput);
            y += 20;

            txtCreateDwgOutputPath = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(320, 23)
            };
            tabCreateDwg.Controls.Add(txtCreateDwgOutputPath);

            btnCreateDwgOutputBrowse = new Button
            {
                Text = LanguageManager.L("BTN_BROWSE"),
                Location = new Point(x + 325, y - 1),
                Size = new Size(85, 25)
            };
            btnCreateDwgOutputBrowse.Click += BtnCreateDwgOutputBrowse_Click;
            tabCreateDwg.Controls.Add(btnCreateDwgOutputBrowse);
            y += 25;

            var lblOutputHint = new Label
            {
                Text = LanguageManager.L("LBL_CREATE_DWG_OUTPUT_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabCreateDwg.Controls.Add(lblOutputHint);
        }

        // ============================================================
        // TAB 5: LANGUAGE
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
            // [CHANGED v1.1] Load multi-folder vào ListBox
            lstCadFolders.Items.Clear();
            foreach (var folder in _settings.CadFolderPaths)
                if (!string.IsNullOrEmpty(folder))
                    lstCadFolders.Items.Add(folder);
            chkUseRevision.Checked = _settings.UseRevisionSuffix;
            txtExtension.Text = _settings.CadExtension.TrimStart('.');
            if (!txtExtension.Text.StartsWith("."))
                txtExtension.Text = "." + txtExtension.Text;

            txtBomXmlPath.Text = _settings.BomXmlPath;
            txtCheckRefExcludePath.Text = _settings.CheckRefExcludeListPath;

            // [NEW v1.2] Create DWG
            txtCreateDwgIniPath.Text = _settings.CreateDwgIniPath;
            txtCreateDwgListPath.Text = _settings.CreateDwgListPath;
            txtCreateDwgOutputPath.Text = _settings.CreateDwgOutputPath;

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
            // [CHANGED v1.1] Save multi-folder từ ListBox
            _settings.CadFolderPaths.Clear();
            foreach (var item in lstCadFolders.Items)
                _settings.CadFolderPaths.Add(item.ToString());

            // Normalize extension
            string ext = txtExtension.Text.Trim();
            if (!string.IsNullOrEmpty(ext) && !ext.StartsWith("."))
                ext = "." + ext;
            _settings.CadExtension = string.IsNullOrEmpty(ext) ? ".dwg" : ext;

            _settings.UseRevisionSuffix = chkUseRevision.Checked;

            _settings.BomXmlPath = txtBomXmlPath.Text.Trim();
            _settings.CheckRefExcludeListPath = txtCheckRefExcludePath.Text.Trim();

            // [NEW v1.2] Create DWG
            _settings.CreateDwgIniPath = txtCreateDwgIniPath.Text.Trim();
            _settings.CreateDwgListPath = txtCreateDwgListPath.Text.Trim();
            _settings.CreateDwgOutputPath = txtCreateDwgOutputPath.Text.Trim();

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

        // [CHANGED v1.1] Thêm folder vào danh sách
        private void BtnAddFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select CAD Drawing Folder";
                dialog.ShowNewFolderButton = false;

                if (lstCadFolders.SelectedItem != null)
                    dialog.SelectedPath = lstCadFolders.SelectedItem.ToString();

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string path = dialog.SelectedPath;
                    if (!lstCadFolders.Items.Contains(path))
                        lstCadFolders.Items.Add(path);
                }
            }
        }

        // [CHANGED v1.1] Xóa folder được chọn khỏi danh sách
        private void BtnRemoveFolder_Click(object sender, EventArgs e)
        {
            if (lstCadFolders.SelectedIndex >= 0)
                lstCadFolders.Items.RemoveAt(lstCadFolders.SelectedIndex);
        }

        // [NEW] Browse for BOM XML file
        private void BtnBomBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select BOM XML File";
                dialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;

                string current = txtBomXmlPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);

                if (dialog.ShowDialog() == DialogResult.OK)
                    txtBomXmlPath.Text = dialog.FileName;
            }
        }

        // [NEW] Browse cho Check Reference exclude list
        private void BtnCheckRefBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select Exclude List File";
                dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;

                string current = txtCheckRefExcludePath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);

                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCheckRefExcludePath.Text = dialog.FileName;
            }
        }

        // [NEW v1.2] Browse cho DWGExport.ini
        private void BtnCreateDwgIniBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select DWGExport.ini File";
                dialog.Filter = "INI Files (*.ini)|*.ini|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;

                string current = txtCreateDwgIniPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);

                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCreateDwgIniPath.Text = dialog.FileName;
            }
        }

        // [NEW v1.2] Browse cho file list .txt
        private void BtnCreateDwgListBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select List File (.txt)";
                dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;

                string current = txtCreateDwgListPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);

                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCreateDwgListPath.Text = dialog.FileName;
            }
        }

        // [NEW v1.2] Browse cho output folder
        private void BtnCreateDwgOutputBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Output Folder for DWG Files";
                dialog.ShowNewFolderButton = true;

                string current = txtCreateDwgOutputPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && Directory.Exists(current))
                    dialog.SelectedPath = current;

                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCreateDwgOutputPath.Text = dialog.SelectedPath;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // [CHANGED v1.1] Validate tất cả folder trong list
            foreach (var item in lstCadFolders.Items)
            {
                string folder = item.ToString();
                if (!Directory.Exists(folder))
                {
                    var result = MessageBox.Show(
                        $"Folder does not exist:\n{folder}\n\nSave anyway?",
                        LanguageManager.L("TITLE_WARNING"),
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (result == DialogResult.No) return;
                }
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