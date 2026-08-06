using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OpenCadDrawingAddin.Logic;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// Form Settings cho Mini Tool Add-in
    /// Version 1.5 - Them tab Auto Hole Note [NEW v1.5]
    /// </summary>
    public class SettingsForm : Form
    {
        // === Controls ===
        private TabControl mainTabControl;
        private TabPage tabSettings;
        private TabPage tabLanguage;
        private TabPage tabBom;
        private TabPage tabCheckRef;
        private TabPage tabCreateDwg;
        private TabPage tabCopyPaste;
        private TabPage tabIdwCheck;
        private TabPage tabAutoHoleNote;   // [NEW v1.5]

        // Tab Settings
        private Label lblCadFolder;
        private ListBox lstCadFolders;
        private Button btnAddFolder;
        private Button btnRemoveFolder;
        private Label lblFolderHint;
        private CheckBox chkUseRevision;
        private Label lblRevisionHint;
        private Label lblExtension;
        private TextBox txtExtension;
        private Label lblExtensionHint;

        // Tab Bom
        private TextBox txtBomXmlPath;
        private Button btnBomBrowse;

        // Tab Check Reference
        private TextBox txtCheckRefExcludePath;
        private Button btnCheckRefBrowse;

        // Tab Create DWG
        private TextBox txtCreateDwgIniPath;
        private Button btnCreateDwgIniBrowse;
        private TextBox txtCreateDwgListPath;
        private Button btnCreateDwgListBrowse;
        private TextBox txtCreateDwgOutputPath;
        private Button btnCreateDwgOutputBrowse;

        // Tab Language
        private ListBox lstLanguages;

        // Tab Copy-Paste
        private CheckBox chkShowCopyNotify;
        private CheckBox chkShowPasteNotify;

        // Tab IDW Check
        private CheckBox chkAutoCheckDrawingName;
        private CheckBox chkAutoCheckAppearance;

        // Tab Auto Hole Note [NEW v1.5]
        private NumericUpDown nudTextHeight;
        private NumericUpDown nudClusterRadius;
        private NumericUpDown nudTolTap;
        private Button btnRunAutoHoleNote;
        // [AUTO-SIZE] CheckBox bat/tat che do tu dong tinh theo view.Scale
        private CheckBox chkHoleNoteAutoSize;

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

            tabSettings = new TabPage(LanguageManager.L("TAB_SETTINGS"));
            BuildSettingsTab();
            mainTabControl.TabPages.Add(tabSettings);

            tabBom = new TabPage(LanguageManager.L("TAB_BOM"));
            BuildBomTab();
            mainTabControl.TabPages.Add(tabBom);

            tabCheckRef = new TabPage(LanguageManager.L("TAB_CHECK_REF"));
            BuildCheckRefTab();
            mainTabControl.TabPages.Add(tabCheckRef);

            tabCreateDwg = new TabPage(LanguageManager.L("TAB_CREATE_DWG"));
            BuildCreateDwgTab();
            mainTabControl.TabPages.Add(tabCreateDwg);

            tabCopyPaste = new TabPage(LanguageManager.L("TAB_COPY_PASTE"));
            BuildCopyPasteTab();
            mainTabControl.TabPages.Add(tabCopyPaste);

            tabIdwCheck = new TabPage(LanguageManager.L("TAB_IDW_CHECK"));
            BuildIdwCheckTab();
            mainTabControl.TabPages.Add(tabIdwCheck);

            // [NEW v1.5] Tab Auto Hole Note
            tabAutoHoleNote = new TabPage("Auto Hole Note");
            BuildAutoHoleNoteTab();
            mainTabControl.TabPages.Add(tabAutoHoleNote);

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

            lblCadFolder = new Label
            {
                Text = LanguageManager.L("LBL_CAD_FOLDER"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabSettings.Controls.Add(lblCadFolder);
            y += 20;

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
                Text = "\u2212",
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
        // TAB 2: BOM FORMAT
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
        // TAB 3: CHECK REFERENCE
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
        // TAB 4: CREATE DWG
        // ============================================================
        private void BuildCreateDwgTab()
        {
            int x = 15, y = 15;

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
        // TAB 5: COPY-PASTE
        // ============================================================
        private void BuildCopyPasteTab()
        {
            int x = 15, y = 20;

            var lblTitle = new Label
            {
                Text = LanguageManager.L("LBL_COPY_PASTE_NOTIFY"),
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabCopyPaste.Controls.Add(lblTitle);
            y += 28;

            chkShowCopyNotify = new CheckBox
            {
                Text = LanguageManager.L("CHK_SHOW_COPY_NOTIFY"),
                Location = new Point(x, y),
                Size = new Size(400, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabCopyPaste.Controls.Add(chkShowCopyNotify);
            y += 28;

            chkShowPasteNotify = new CheckBox
            {
                Text = LanguageManager.L("CHK_SHOW_PASTE_NOTIFY"),
                Location = new Point(x, y),
                Size = new Size(400, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabCopyPaste.Controls.Add(chkShowPasteNotify);
            y += 36;

            var lblHint = new Label
            {
                Text = LanguageManager.L("LBL_COPY_PASTE_HINT"),
                Location = new Point(x, y),
                Size = new Size(400, 36),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabCopyPaste.Controls.Add(lblHint);
        }

        // ============================================================
        // TAB 6: IDW CHECK
        // ============================================================
        private void BuildIdwCheckTab()
        {
            int x = 15, y = 20;

            var lblTitle = new Label
            {
                Text = LanguageManager.L("LBL_IDW_CHECK_TITLE"),
                Location = new Point(x, y),
                Size = new Size(430, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabIdwCheck.Controls.Add(lblTitle);
            y += 30;

            chkAutoCheckDrawingName = new CheckBox
            {
                Text = LanguageManager.L("CHK_IDW_CHECK_NAME"),
                Location = new Point(x, y),
                Size = new Size(430, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabIdwCheck.Controls.Add(chkAutoCheckDrawingName);
            y += 28;

            chkAutoCheckAppearance = new CheckBox
            {
                Text = LanguageManager.L("CHK_IDW_CHECK_APPEARANCE"),
                Location = new Point(x, y),
                Size = new Size(430, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabIdwCheck.Controls.Add(chkAutoCheckAppearance);
            y += 36;

            var lblHint = new Label
            {
                Text = LanguageManager.L("LBL_IDW_CHECK_HINT"),
                Location = new Point(x, y),
                Size = new Size(430, 36),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabIdwCheck.Controls.Add(lblHint);
        }

        // ============================================================
        // TAB 7: AUTO HOLE NOTE [NEW v1.5]
        // ============================================================
        private void BuildAutoHoleNoteTab()
        {
            int x = 15, y = 15;

            // Title
            var lblTitle = new Label
            {
                Text = "Auto Hole Note Settings",
                Location = new Point(x, y),
                Size = new Size(400, 18),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            tabAutoHoleNote.Controls.Add(lblTitle);
            y += 30;

            // [AUTO-SIZE] --- CheckBox Auto Size ---
            chkHoleNoteAutoSize = new CheckBox
            {
                Text = "Auto size theo scale cua view (base @ view 1:4)",
                Location = new Point(x, y),
                Size = new Size(430, 22),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue
            };
            tabAutoHoleNote.Controls.Add(chkHoleNoteAutoSize);
            y += 28;

            // --- Text Height ---
            var lblTextHeight = new Label
            {
                Text = "Text height (mm):",
                Location = new Point(x, y),
                Size = new Size(155, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabAutoHoleNote.Controls.Add(lblTextHeight);

            nudTextHeight = new NumericUpDown
            {
                Location = new Point(x + 160, y - 2),
                Size = new Size(75, 24),
                Minimum = 0.5M,
                Maximum = 5.0M,
                DecimalPlaces = 1,
                Increment = 0.5M,
                Value = 3.0M,
                Font = new Font("Segoe UI", 9F)
            };
            tabAutoHoleNote.Controls.Add(nudTextHeight);

            // [AUTO-SIZE] Label default ben phai
            var lblTextDefault = new Label
            {
                Text = "(Default: 3.0)",
                Location = new Point(x + 240, y),
                Size = new Size(120, 22),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblTextDefault);
            y += 28;

            var lblTextHint = new Label
            {
                Text = "Chieu cao chu tren sheet (mm) @ view scale 1:4. Auto se scale tu day.",
                Location = new Point(x, y),
                Size = new Size(430, 16),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblTextHint);
            y += 28;

            // --- Cluster Radius ---
            var lblCluster = new Label
            {
                Text = "Cluster radius (mm):",
                Location = new Point(x, y),
                Size = new Size(155, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabAutoHoleNote.Controls.Add(lblCluster);

            nudClusterRadius = new NumericUpDown
            {
                Location = new Point(x + 160, y - 2),
                Size = new Size(75, 24),
                Minimum = 5,
                Maximum = 200,
                DecimalPlaces = 0,
                Increment = 5,
                Value = 30,
                Font = new Font("Segoe UI", 9F)
            };
            tabAutoHoleNote.Controls.Add(nudClusterRadius);

            // [AUTO-SIZE] Label default ben phai
            var lblClusterDefault = new Label
            {
                Text = "(Default: 30)",
                Location = new Point(x + 240, y),
                Size = new Size(120, 22),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblClusterDefault);
            y += 28;

            var lblClusterHint = new Label
            {
                Text = "Radius gom cum (mm) @ view scale 1:4. Auto se scale tu day.",
                Location = new Point(x, y),
                Size = new Size(430, 16),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblClusterHint);
            y += 28;

            // --- Tap Tolerance ---
            var lblTol = new Label
            {
                Text = "Tap tolerance (x0.01mm):",
                Location = new Point(x, y),
                Size = new Size(155, 22),
                Font = new Font("Segoe UI", 9F)
            };
            tabAutoHoleNote.Controls.Add(lblTol);

            nudTolTap = new NumericUpDown
            {
                Location = new Point(x + 160, y - 2),
                Size = new Size(75, 24),
                Minimum = 1,
                Maximum = 50,
                DecimalPlaces = 0,
                Increment = 1,
                Value = 8,  // [FIX v1.6] 15 -> 8 (0.08mm) de tranh Ø6.5 (lo tron) match nham M8 (pitch 6.647)
                Font = new Font("Segoe UI", 9F)
            };
            tabAutoHoleNote.Controls.Add(nudTolTap);

            // [AUTO-SIZE] Label default ben phai
            var lblTolDefault = new Label
            {
                Text = "(Default: 8)",  // [FIX v1.6] 15 -> 8
                Location = new Point(x + 240, y),
                Size = new Size(120, 22),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblTolDefault);
            y += 28;

            var lblTolHint = new Label
            {
                Text = "Sai so match lo ren (don vi 0.01mm). 8 = +-0.08mm. Default: 8",  // [FIX v1.6] 15 -> 8
                Location = new Point(x, y),
                Size = new Size(400, 16),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblTolHint);
            y += 30;

            // --- Run Button ---
            btnRunAutoHoleNote = new Button
            {
                Text = "▶  Run Auto Hole Note",
                Location = new Point(x, y),
                Size = new Size(200, 32),
                BackColor = Color.LightSkyBlue,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            btnRunAutoHoleNote.Click += BtnRunAutoHoleNote_Click;
            tabAutoHoleNote.Controls.Add(btnRunAutoHoleNote);
            y += 42;

            var lblRunHint = new Label
            {
                Text = "Click -> chon View trong IDW -> tu dong dien note.\nCtrl+Z sau khi chay de xoa toan bo va thu lai.",
                Location = new Point(x, y),
                Size = new Size(400, 34),
                ForeColor = Color.Gray,
                Font = new Font("Segoe UI", 8F, FontStyle.Italic)
            };
            tabAutoHoleNote.Controls.Add(lblRunHint);
        }

        // ============================================================
        // TAB 8: LANGUAGE
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
                lstLanguages.Items.Add($"{kvp.Key} - {kvp.Value}");
            tabLanguage.Controls.Add(lstLanguages);
        }

        // ============================================================
        // LOAD / SAVE SETTINGS <-> UI
        // ============================================================

        private void LoadSettingsToUI()
        {
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

            txtCreateDwgIniPath.Text = _settings.CreateDwgIniPath;
            txtCreateDwgListPath.Text = _settings.CreateDwgListPath;
            txtCreateDwgOutputPath.Text = _settings.CreateDwgOutputPath;

            chkShowCopyNotify.Checked = _settings.ShowCopyNotification;
            chkShowPasteNotify.Checked = _settings.ShowPasteNotification;

            chkAutoCheckDrawingName.Checked = _settings.AutoCheckDrawingName;
            chkAutoCheckAppearance.Checked = _settings.AutoCheckAppearance;

            // [NEW v1.5] Auto Hole Note
            nudTextHeight.Value = (decimal)Math.Max(0.5, Math.Min(5.0, _settings.HoleNoteTextHeightMm));
            nudClusterRadius.Value = (decimal)Math.Max(5, Math.Min(200, _settings.HoleNoteClusterRadiusMm));
            nudTolTap.Value = (decimal)Math.Max(1, Math.Min(50, _settings.HoleNoteTolTap * 100));

            // [AUTO-SIZE]
            chkHoleNoteAutoSize.Checked = _settings.HoleNoteAutoSize;

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
            _settings.CadFolderPaths.Clear();
            foreach (var item in lstCadFolders.Items)
                _settings.CadFolderPaths.Add(item.ToString());

            string ext = txtExtension.Text.Trim();
            if (!string.IsNullOrEmpty(ext) && !ext.StartsWith("."))
                ext = "." + ext;
            _settings.CadExtension = string.IsNullOrEmpty(ext) ? ".dwg" : ext;

            _settings.UseRevisionSuffix = chkUseRevision.Checked;
            _settings.BomXmlPath = txtBomXmlPath.Text.Trim();
            _settings.CheckRefExcludeListPath = txtCheckRefExcludePath.Text.Trim();

            _settings.CreateDwgIniPath = txtCreateDwgIniPath.Text.Trim();
            _settings.CreateDwgListPath = txtCreateDwgListPath.Text.Trim();
            _settings.CreateDwgOutputPath = txtCreateDwgOutputPath.Text.Trim();

            _settings.ShowCopyNotification = chkShowCopyNotify.Checked;
            _settings.ShowPasteNotification = chkShowPasteNotify.Checked;

            _settings.AutoCheckDrawingName = chkAutoCheckDrawingName.Checked;
            _settings.AutoCheckAppearance = chkAutoCheckAppearance.Checked;

            // [NEW v1.5] Auto Hole Note
            _settings.HoleNoteTextHeightMm = (double)nudTextHeight.Value;
            _settings.HoleNoteClusterRadiusMm = (double)nudClusterRadius.Value;
            _settings.HoleNoteTolTap = (double)nudTolTap.Value / 100.0;

            // [AUTO-SIZE]
            _settings.HoleNoteAutoSize = chkHoleNoteAutoSize.Checked;

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

        private void BtnRemoveFolder_Click(object sender, EventArgs e)
        {
            if (lstCadFolders.SelectedIndex >= 0)
                lstCadFolders.Items.RemoveAt(lstCadFolders.SelectedIndex);
        }

        private void BtnBomBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select BOM XML File";
                dialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
                string current = txtBomXmlPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);
                if (dialog.ShowDialog() == DialogResult.OK)
                    txtBomXmlPath.Text = dialog.FileName;
            }
        }

        private void BtnCheckRefBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select Exclude List File";
                dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                string current = txtCheckRefExcludePath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);
                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCheckRefExcludePath.Text = dialog.FileName;
            }
        }

        private void BtnCreateDwgIniBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select DWGExport.ini File";
                dialog.Filter = "INI Files (*.ini)|*.ini|All Files (*.*)|*.*";
                string current = txtCreateDwgIniPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);
                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCreateDwgIniPath.Text = dialog.FileName;
            }
        }

        private void BtnCreateDwgListBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select List File (.txt)";
                dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                string current = txtCreateDwgListPath.Text.Trim();
                if (!string.IsNullOrEmpty(current) && File.Exists(current))
                    dialog.InitialDirectory = Path.GetDirectoryName(current);
                if (dialog.ShowDialog() == DialogResult.OK)
                    txtCreateDwgListPath.Text = dialog.FileName;
            }
        }

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

        // [NEW v1.5] Run Auto Hole Note
        private void BtnRunAutoHoleNote_Click(object sender, EventArgs e)
        {
            var inventorApp = StandardAddInServer.InventorApplication;
            if (inventorApp == null)
            {
                MessageBox.Show("Khong tim thay Inventor Application.",
                    "Auto Hole Note", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Inventor.DrawingDocument drawDoc = null;
            try { drawDoc = inventorApp.ActiveDocument as Inventor.DrawingDocument; } catch { }
            if (drawDoc == null)
            {
                MessageBox.Show("Hay mo file IDW truoc khi chay.",
                    "Auto Hole Note", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Pick view - minimize form truoc de khong che view
            this.WindowState = FormWindowState.Minimized;

            Inventor.DrawingView view = null;
            try
            {
                view = inventorApp.CommandManager.Pick(
                    Inventor.SelectionFilterEnum.kDrawingViewFilter,
                    "Click vao view can quet lo:") as Inventor.DrawingView;
            }
            catch { }
            finally
            {
                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            }

            if (view == null)
            {
                MessageBox.Show("Chua chon view.",
                    "Auto Hole Note", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double textH = (double)nudTextHeight.Value;
            double clusterR = (double)nudClusterRadius.Value;
            double tolTap = (double)nudTolTap.Value / 100.0;

            // [AUTO-SIZE] Neu bat Auto -> textH & clusterR o UI la BASE @ 1:4.
            // ComputeAutoSize scale tu base theo view.Scale, lam tron & clamp.
            // Tap tolerance khong lien quan hien thi -> khong scale.
            if (chkHoleNoteAutoSize.Checked)
            {
                AutoHoleNoteLogic.ComputeAutoSize(
                    view.Scale,
                    (double)nudTextHeight.Value,        // base text @ 1:4
                    (double)nudClusterRadius.Value,     // base cluster @ 1:4
                    out textH,
                    out clusterR);
            }

            var logic = new AutoHoleNoteLogic(inventorApp);
            logic.Run(view, textH, clusterR, tolTap);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
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