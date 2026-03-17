using Inventor;
using ReplaceFileAddin.Logic;
using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReplaceFileAddin
{
    [Guid("64DB7978-5BCD-4388-BD7B-7EBA2C4B33D2")]
    [ProgId("ReplaceFileAddin.StandardAddInServer")]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_myButton;        // Replace File (trên)
        private ButtonDefinition m_settingsButton;  // Settings (giữa) [v1.1]
        private ButtonDefinition m_aboutButton;     // About (dưới)

        // --- PHẦN ĐA NGÔN NGỮ (MULTI-LANGUAGE) ---
        public static bool IsVietnamese = false;

        // Legacy method - kept for backward compatibility
        public static string _T(string en, string vn) => IsVietnamese ? vn : en;

        // --- KHAI BÁO CÁC CÂU THÔNG BÁO (now using LanguageManager) ---
        public static string Title_Info => LanguageManager.L("TITLE_INFO");
        public static string Title_Error => LanguageManager.L("TITLE_ERROR");
        public static string Title_Warning => LanguageManager.L("TITLE_WARNING");
        public static string Title_SelectError => LanguageManager.L("TITLE_SELECT_ERROR");

        public static string Msg_ProcessRunning => LanguageManager.L("MSG_PROCESS_RUNNING");
        public static string Msg_SelectPartWarning => LanguageManager.L("MSG_SELECT_PART_WARNING");
        public static string Msg_SelectTopPartError => LanguageManager.L("MSG_SELECT_TOP_PART");
        // ------------------------------------------

        public static LicenseHelper.UpdateInfo _cachedUpdateInfo = null;
        private static bool _isUpdateCheckRunning = false;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            m_inventorApplication = addInSiteObject.Application;

            // [v1.2] Load ngôn ngữ từ settings
            try { CleaningSettings.LoadAndApplyLanguage(); } catch { }

            try { CreateUserInterface(); } catch { }

            // --- KHẮC PHỤC LỖI CRASH KHI KHỞI ĐỘNG ---
            Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(8000);
                    LicenseHelper.StartBackgroundCheck();
                    await Task.Delay(3000);
                    LicenseHelper.ShowTrialReminderIfNeeded(IsVietnamese);
                    await CheckUpdateInBackground();
                    try
                    {
                        string addinVer = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                        string inventorVer = m_inventorApplication?.SoftwareVersion?.DisplayVersion ?? "Unknown";
                        LicenseHelper.SendTrackingAsync(addinVer, inventorVer);
                    }
                    catch { }
                    try { LicenseHelper.SendUsageIfNewDay(); } catch { }
                }
                catch { }
            });
        }

        public void Deactivate()
        {
            if (m_myButton != null) { m_myButton.Delete(); m_myButton = null; }
            if (m_settingsButton != null) { m_settingsButton.Delete(); m_settingsButton = null; }
            if (m_aboutButton != null) { m_aboutButton.Delete(); m_aboutButton = null; }
            if (m_inventorApplication != null) { Marshal.ReleaseComObject(m_inventorApplication); m_inventorApplication = null; }
            GC.Collect();
        }

        public object Automation => null;
        public void ExecuteCommand(int commandID) { }

        private void CreateUserInterface()
        {
            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

            stdole.IPictureDisp icon16 = null, icon32 = null;
            try
            {
                icon16 = IconConverter.ToIPictureDisp(Properties.Resources.ReplaceIcon16);
                icon32 = IconConverter.ToIPictureDisp(Properties.Resources.ReplaceIcon32);
            }
            catch { }

            // 1. Nút Replace File (trên cùng)
            m_myButton = controlDefs.AddButtonDefinition(
                "Replace File", "ReplaceFileCmd", CommandTypesEnum.kShapeEditCmdType,
                GenerateClientId("ReplaceFileCmd"), "Copy and replace file.", "Replace File Tool",
                icon16, icon32);
            m_myButton.OnExecute += m_myButton_OnExecute;

            // 2. [v1.1] Nút Settings (giữa)
            m_settingsButton = controlDefs.AddButtonDefinition(
                "Settings", "SettingsCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("SettingsCmd"),
                LanguageManager.L("ABOUT_CONFIGURE_CLEANING"),
                "Settings", null, null);
            m_settingsButton.OnExecute += m_settingsButton_OnExecute;

            // 3. Nút About / License (dưới cùng)
            m_aboutButton = controlDefs.AddButtonDefinition(
                "About / License", "AboutAuthorCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("AboutAuthorCmd"), "Information & Activation.", "About Author",
                null, null);
            m_aboutButton.OnExecute += m_aboutButton_OnExecute;

            Ribbon asmRibbon = m_inventorApplication.UserInterfaceManager.Ribbons["Assembly"];
            RibbonTab asmTab = asmRibbon.RibbonTabs["id_TabAssemble"];
            RibbonPanel myPanel = null;
            try { myPanel = asmTab.RibbonPanels.Add("Addin Tools", "id_Panel_AddinTools", "{64DB7978-5BCD-4388-BD7B-7EBA2C4B33D2}"); }
            catch { myPanel = asmTab.RibbonPanels["id_Panel_AddinTools"]; }

            // [v1.1] Xếp 3 nút dọc 1 cột: useLargeIcon = false để nút nhỏ xếp dọc
            // Thứ tự: Replace File (trên) -> Settings (giữa) -> About (dưới)
            if (myPanel != null)
            {
                if (!ButtonExists(myPanel, m_myButton))
                    myPanel.CommandControls.AddButton(m_myButton, false);      // false = small icon, xếp dọc
                if (!ButtonExists(myPanel, m_settingsButton))
                    myPanel.CommandControls.AddButton(m_settingsButton, false);
                if (!ButtonExists(myPanel, m_aboutButton))
                    myPanel.CommandControls.AddButton(m_aboutButton, false);
            }
        }

        private bool ButtonExists(RibbonPanel panel, ButtonDefinition buttonDef)
        {
            foreach (CommandControl cmd in panel.CommandControls)
                if (cmd.ControlDefinition == buttonDef) return true;
            return false;
        }

        private async Task CheckUpdateInBackground()
        {
            if (_isUpdateCheckRunning) return;
            _isUpdateCheckRunning = true;
            try
            {
                string currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                var updateInfo = await LicenseHelper.CheckForUpdateAsync(currentVersion);

                if (updateInfo.HasUpdate && updateInfo.ForceUpdate)
                {
                    await Task.Delay(500);
                    var retryInfo = await LicenseHelper.CheckForUpdateAsync(currentVersion);
                    if (!retryInfo.ForceUpdate) updateInfo = retryInfo;
                }

                _cachedUpdateInfo = updateInfo;
                if (updateInfo != null)
                    LicenseHelper.SetForceUpdate(updateInfo.ForceUpdate);
            }
            catch { }
            finally { _isUpdateCheckRunning = false; }
        }

        private void m_myButton_OnExecute(NameValueMap Context)
        {
            try
            {
                if (_cachedUpdateInfo != null && _cachedUpdateInfo.HasUpdate && _cachedUpdateInfo.ForceUpdate)
                {
                    string title = LanguageManager.L("ABOUT_UPDATE_REQUIRED");
                    string msg = LanguageManager.L("ABOUT_UPDATE_MSG", _cachedUpdateInfo.NewVersion);

                    MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    using (var frm = new AboutBoxForm(true)) { frm.ShowDialog(); }
                    return;
                }

                if (_cachedUpdateInfo == null)
                {
                    _ = CheckUpdateInBackground();
                }

                var logic = new ReplaceFileLogic(m_inventorApplication);
                logic.RunReplaceOperation();
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Lỗi khi chạy Replace File", ex);
            }
        }

        private void m_aboutButton_OnExecute(NameValueMap Context)
        {
            using (var frm = new AboutBoxForm(true)) { frm.ShowDialog(); }
        }

        // [v1.1] Xử lý khi click nút Settings
        private void m_settingsButton_OnExecute(NameValueMap Context)
        {
            using (var frm = new SettingsForm())
            {
                frm.ShowDialog();
            }
        }

        private string GenerateClientId(string cmdName) => "{" + this.GetType().GUID.ToString() + "}+" + cmdName;

        internal class IconConverter : AxHost
        {
            private IconConverter() : base(string.Empty) { }
            public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Image image) => (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        public class AboutBoxForm : Form
        {
            private Label lblInfo;
            private LinkLabel lnkUpdateNotify;
            private Button btnActive;
            private Button btnUpdate;
            private Button btnClose;
            private LinkLabel lnkCopyID;
            private LinkLabel lnkHelp;
            private LinkLabel lnkOtherAddin;
            private LinkLabel lnkViewLog;
            private Label lblAuthor;
            private LinkLabel lblEmail;
            private LinkLabel lnkCopyEmail;

#if SHOW_CONTACT_INFO
            private LinkLabel lnkTelegramMain;
            private LinkLabel lnkWhatsAppMain;
            private LinkLabel lnkZaloMain;
            private LinkLabel lnkTelegramQR;
            private LinkLabel lnkWhatsAppQR;
            private LinkLabel lnkZaloQR;
#endif

            private ToolTip toolTipContacts;

            private string downloadUrl = "";
            private string helpUrl = "";
            private string otherAddinUrl = "";
            private bool _isAutoCheck = false;

            private string _currentVersion = "";
            private string _hwId = "";
            private LicenseHelper.LicenseInfo _licInfo;
            private LicenseHelper.UpdateInfo _upInfo;

            public AboutBoxForm(bool isAutoCheck)
            {
                _isAutoCheck = isAutoCheck;
                InitializeComponent();
                this.Load += async (s, e) => await SafeLoadLicenseData();
            }

            private void InitializeComponent()
            {
                this.Size = new System.Drawing.Size(500, 440);
                this.Text = "About & License Manager";
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowIcon = false;
                this.ShowInTaskbar = false;

                toolTipContacts = new ToolTip();
                toolTipContacts.AutoPopDelay = 5000;
                toolTipContacts.InitialDelay = 500;
                toolTipContacts.ReshowDelay = 500;
                toolTipContacts.ShowAlways = true;

                lblInfo = new Label();
                lblInfo.Location = new System.Drawing.Point(20, 20);
                lblInfo.Size = new System.Drawing.Size(380, 160);
                lblInfo.Font = new System.Drawing.Font("Segoe UI", 10);
                this.Controls.Add(lblInfo);

                lnkHelp = new LinkLabel();
                lnkHelp.Location = new System.Drawing.Point(420, 20);
                lnkHelp.Size = new System.Drawing.Size(50, 25);
                lnkHelp.Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold);
                lnkHelp.LinkColor = System.Drawing.Color.Blue;
                lnkHelp.LinkClicked += (s, e) => {
                    if (!string.IsNullOrEmpty(helpUrl) && helpUrl.StartsWith("http")) OpenUrl(helpUrl);
                    else MessageBox.Show(LanguageManager.L("ABOUT_LOADING"), LanguageManager.L("TITLE_INFO"));
                };
                this.Controls.Add(lnkHelp);

                lnkCopyID = new LinkLabel();
                lnkCopyID.Location = new System.Drawing.Point(400, 70);
                lnkCopyID.Size = new System.Drawing.Size(70, 20);
                lnkCopyID.Font = new System.Drawing.Font("Segoe UI", 9);
                lnkCopyID.LinkColor = System.Drawing.Color.Green;
                lnkCopyID.LinkClicked += (s, e) => {
                    LicenseHelper.CopyHardwareIdToClipboard();
                    MessageBox.Show(LanguageManager.L("ABOUT_ID_COPIED"), LanguageManager.L("TITLE_INFO"));
                };
                this.Controls.Add(lnkCopyID);

                lnkUpdateNotify = new LinkLabel();
                lnkUpdateNotify.Location = new System.Drawing.Point(20, 185);
                lnkUpdateNotify.Size = new System.Drawing.Size(450, 20);
                lnkUpdateNotify.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold);
                lnkUpdateNotify.LinkColor = System.Drawing.Color.Gray;
                lnkUpdateNotify.LinkClicked += (s, e) => {
                    if (!string.IsNullOrEmpty(downloadUrl)) OpenUrl(downloadUrl);
                };
                this.Controls.Add(lnkUpdateNotify);

                int btnY = 215;

                btnActive = new Button();
                btnActive.Location = new System.Drawing.Point(20, btnY);
                btnActive.Size = new System.Drawing.Size(140, 35);
                btnActive.BackColor = System.Drawing.Color.LightYellow;
                btnActive.Click += async (s, e) => await OnActiveClick();
                this.Controls.Add(btnActive);

                btnUpdate = new Button();
                btnUpdate.Location = new System.Drawing.Point(170, btnY);
                btnUpdate.Size = new System.Drawing.Size(140, 35);
                btnUpdate.Click += async (s, e) => await OnUpdateClick();
                this.Controls.Add(btnUpdate);

                btnClose = new Button();
                btnClose.Location = new System.Drawing.Point(320, btnY);
                btnClose.Size = new System.Drawing.Size(140, 35);
                btnClose.Click += (s, e) => this.Close();
                this.Controls.Add(btnClose);

                // --- AUTHOR PANEL ---
                Panel pnlAuthor = new Panel();
                pnlAuthor.Location = new System.Drawing.Point(0, 275);
                pnlAuthor.Size = new System.Drawing.Size(500, 140);
                pnlAuthor.BackColor = System.Drawing.Color.WhiteSmoke;
                this.Controls.Add(pnlAuthor);

                lblAuthor = new Label();
                lblAuthor.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold);
                lblAuthor.ForeColor = System.Drawing.Color.DarkSlateGray;
                lblAuthor.Location = new System.Drawing.Point(20, 10);
                lblAuthor.Size = new System.Drawing.Size(200, 18);
                pnlAuthor.Controls.Add(lblAuthor);

                lblEmail = new LinkLabel();
                lblEmail.Text = "Email: replacefile.addin@gmail.com";
                lblEmail.Font = new System.Drawing.Font("Segoe UI", 9);
                lblEmail.Location = new System.Drawing.Point(20, 30);
                lblEmail.Size = new System.Drawing.Size(220, 18);
                lblEmail.LinkClicked += (s, e) => OpenUrl("mailto:replacefile.addin@gmail.com");
                pnlAuthor.Controls.Add(lblEmail);

                lnkCopyEmail = new LinkLabel();
                lnkCopyEmail.Font = new System.Drawing.Font("Segoe UI", 8);
                lnkCopyEmail.Location = new System.Drawing.Point(245, 30);
                lnkCopyEmail.Size = new System.Drawing.Size(50, 18);
                lnkCopyEmail.LinkColor = System.Drawing.Color.DimGray;
                lnkCopyEmail.LinkClicked += (s, e) => {
                    Clipboard.SetText("replacefile.addin@gmail.com");
                    MessageBox.Show(LanguageManager.L("ABOUT_EMAIL_COPIED"), LanguageManager.L("TITLE_INFO"));
                };
                pnlAuthor.Controls.Add(lnkCopyEmail);

#if SHOW_CONTACT_INFO
                int qrColumnX = 230;
                string qrTip = "Click to view QR Code";
                string linkTip = "Click to open Link";

                lnkTelegramMain = new LinkLabel();
                lnkTelegramMain.Text = "Telegram: t.me/replacefile";
                lnkTelegramMain.Font = new System.Drawing.Font("Segoe UI", 9);
                lnkTelegramMain.AutoSize = true;
                lnkTelegramMain.Location = new System.Drawing.Point(20, 50);
                lnkTelegramMain.LinkColor = System.Drawing.Color.FromArgb(0, 136, 204);
                lnkTelegramMain.LinkClicked += (s, e) => OpenUrl("https://t.me/replacefile");
                toolTipContacts.SetToolTip(lnkTelegramMain, linkTip);
                pnlAuthor.Controls.Add(lnkTelegramMain);

                lnkTelegramQR = new LinkLabel();
                lnkTelegramQR.Font = new System.Drawing.Font("Segoe UI", 9, FontStyle.Regular);
                lnkTelegramQR.AutoSize = true;
                lnkTelegramQR.Location = new System.Drawing.Point(qrColumnX, 50);
                lnkTelegramQR.LinkColor = System.Drawing.Color.Gray;
                lnkTelegramQR.LinkClicked += (s, e) => {
                    try { ShowQRCode(Properties.Resources.QR_Tele, "https://t.me/replacefile", "Telegram QR"); }
                    catch { ShowQRCode(null, "https://t.me/replacefile", "Telegram QR"); }
                };
                toolTipContacts.SetToolTip(lnkTelegramQR, qrTip);
                pnlAuthor.Controls.Add(lnkTelegramQR);

                lnkWhatsAppMain = new LinkLabel();
                lnkWhatsAppMain.Text = "WhatsApp: +84 963 058 858";
                lnkWhatsAppMain.Font = new System.Drawing.Font("Segoe UI", 9);
                lnkWhatsAppMain.AutoSize = true;
                lnkWhatsAppMain.Location = new System.Drawing.Point(20, 70);
                lnkWhatsAppMain.LinkColor = System.Drawing.Color.FromArgb(37, 211, 102);
                lnkWhatsAppMain.LinkClicked += (s, e) => OpenUrl("https://wa.me/84963058858");
                toolTipContacts.SetToolTip(lnkWhatsAppMain, linkTip);
                pnlAuthor.Controls.Add(lnkWhatsAppMain);

                lnkWhatsAppQR = new LinkLabel();
                lnkWhatsAppQR.AutoSize = true;
                lnkWhatsAppQR.Location = new System.Drawing.Point(qrColumnX, 70);
                lnkWhatsAppQR.LinkColor = System.Drawing.Color.Gray;
                lnkWhatsAppQR.LinkClicked += (s, e) => {
                    try { ShowQRCode(Properties.Resources.QR_WhatsApp, "https://wa.me/84963058858", "WhatsApp QR"); }
                    catch { ShowQRCode(null, "https://wa.me/84963058858", "WhatsApp QR"); }
                };
                toolTipContacts.SetToolTip(lnkWhatsAppQR, qrTip);
                pnlAuthor.Controls.Add(lnkWhatsAppQR);

                lnkZaloMain = new LinkLabel();
                lnkZaloMain.Text = "Zalo: 0963 058 858";
                lnkZaloMain.Font = new System.Drawing.Font("Segoe UI", 9);
                lnkZaloMain.AutoSize = true;
                lnkZaloMain.Location = new System.Drawing.Point(20, 90);
                lnkZaloMain.LinkColor = System.Drawing.Color.FromArgb(0, 104, 255);
                lnkZaloMain.LinkClicked += (s, e) => OpenUrl("https://zalo.me/0963058858");
                toolTipContacts.SetToolTip(lnkZaloMain, linkTip);
                pnlAuthor.Controls.Add(lnkZaloMain);

                lnkZaloQR = new LinkLabel();
                lnkZaloQR.AutoSize = true;
                lnkZaloQR.Location = new System.Drawing.Point(qrColumnX, 90);
                lnkZaloQR.LinkColor = System.Drawing.Color.Gray;
                lnkZaloQR.LinkClicked += (s, e) => {
                    try { ShowQRCode(Properties.Resources.QR_Zalo, "https://zalo.me/0963058858", "Zalo QR"); }
                    catch { ShowQRCode(null, "https://zalo.me/0963058858", "Zalo QR"); }
                };
                toolTipContacts.SetToolTip(lnkZaloQR, qrTip);
                pnlAuthor.Controls.Add(lnkZaloQR);
#endif

                lnkOtherAddin = new LinkLabel();
                lnkOtherAddin.Location = new System.Drawing.Point(350, 10);
                lnkOtherAddin.Size = new System.Drawing.Size(120, 18);
                lnkOtherAddin.Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold);
                lnkOtherAddin.LinkColor = System.Drawing.Color.DarkGreen;
                lnkOtherAddin.LinkClicked += (s, e) => {
                    if (!string.IsNullOrEmpty(otherAddinUrl)) OpenUrl(otherAddinUrl);
                    else MessageBox.Show(LanguageManager.L("ABOUT_NO_OTHER_ADDIN"), LanguageManager.L("TITLE_INFO"));
                };
                pnlAuthor.Controls.Add(lnkOtherAddin);

                lnkViewLog = new LinkLabel();
                lnkViewLog.Location = new System.Drawing.Point(350, 30);
                lnkViewLog.Size = new System.Drawing.Size(120, 18);
                lnkViewLog.Font = new System.Drawing.Font("Segoe UI", 9);
                lnkViewLog.LinkColor = System.Drawing.Color.Gray;
                lnkViewLog.LinkClicked += (s, e) => ShowErrorLog();
                pnlAuthor.Controls.Add(lnkViewLog);

                UpdateLanguage();
            }

            private void ShowQRCode(System.Drawing.Image qrImage, string url, string title)
            {
                if (qrImage == null) { OpenUrl(url); return; }

                Form qrForm = new Form();
                qrForm.Text = title;
                qrForm.Size = new System.Drawing.Size(300, 350);
                qrForm.StartPosition = FormStartPosition.CenterScreen;
                qrForm.FormBorderStyle = FormBorderStyle.FixedToolWindow;
                qrForm.ShowInTaskbar = false;

                PictureBox pb = new PictureBox();
                pb.Image = qrImage;
                pb.SizeMode = PictureBoxSizeMode.Zoom;
                pb.Dock = DockStyle.Fill;
                pb.BackColor = System.Drawing.Color.White;
                qrForm.Controls.Add(pb);

                Button btnOpenLink = new Button();
                btnOpenLink.Text = LanguageManager.L("ABOUT_OPEN_BROWSER");
                btnOpenLink.Height = 40;
                btnOpenLink.Dock = DockStyle.Bottom;
                btnOpenLink.Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold);
                btnOpenLink.BackColor = System.Drawing.Color.AliceBlue;
                btnOpenLink.Click += (s, e) => { OpenUrl(url); qrForm.Close(); };
                qrForm.Controls.Add(btnOpenLink);

                pb.Click += (s, e) => { OpenUrl(url); qrForm.Close(); };
                pb.Cursor = Cursors.Hand;

                qrForm.ShowDialog();
            }

            private void UpdateLanguage()
            {
                this.Text = LanguageManager.L("ABOUT_TITLE");
                lnkHelp.Text = LanguageManager.L("ABOUT_HELP");
                lnkCopyID.Text = LanguageManager.L("ABOUT_COPY_ID");
                lnkCopyEmail.Text = LanguageManager.L("ABOUT_COPY");
                lnkOtherAddin.Text = LanguageManager.L("ABOUT_OTHER_ADDIN");
                lnkViewLog.Text = LanguageManager.L("ABOUT_VIEW_LOG");

                btnUpdate.Text = LanguageManager.L("ABOUT_CHECK_UPDATE");
                btnClose.Text = LanguageManager.L("ABOUT_CLOSE");
                lblAuthor.Text = LanguageManager.L("ABOUT_AUTHOR");

#if SHOW_CONTACT_INFO
                lnkTelegramMain.Text = "Telegram: t.me/replacefile";
                lnkWhatsAppMain.Text = "WhatsApp: +84 963 058 858";
                lnkZaloMain.Text = "Zalo: 0963 058 858";

                string qrText = LanguageManager.L("ABOUT_QR_CLICK");
                lnkTelegramQR.Text = qrText;
                lnkWhatsAppQR.Text = qrText;
                lnkZaloQR.Text = qrText;

                string tipLink = LanguageManager.L("ABOUT_TIP_OPEN_LINK");
                string tipQR = LanguageManager.L("ABOUT_TIP_VIEW_QR");

                toolTipContacts.SetToolTip(lnkTelegramMain, tipLink);
                toolTipContacts.SetToolTip(lnkWhatsAppMain, tipLink);
                toolTipContacts.SetToolTip(lnkZaloMain, tipLink);

                toolTipContacts.SetToolTip(lnkTelegramQR, tipQR);
                toolTipContacts.SetToolTip(lnkWhatsAppQR, tipQR);
                toolTipContacts.SetToolTip(lnkZaloQR, tipQR);
#endif

                if (_licInfo != null) UpdateUI(_hwId, _currentVersion, _licInfo, _upInfo);
                else lblInfo.Text = LanguageManager.L("ABOUT_READING_INFO");
            }

            private void OpenUrl(string url)
            {
                if (string.IsNullOrEmpty(url)) return;
                try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true }); }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }

            private void ShowErrorLog()
            {
                string logPath = LicenseHelper.GetLogPath();
                if (!System.IO.File.Exists(logPath)) { MessageBox.Show(LanguageManager.L("ABOUT_NO_ERRORS"), LanguageManager.L("TITLE_INFO")); return; }
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{logPath}\"");
            }

            private async Task SafeLoadLicenseData()
            {
                try
                {
                    _hwId = LicenseHelper.GetHardwareId();
                    _currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();

                    if (_isAutoCheck)
                    {
                        _licInfo = await LicenseHelper.CheckLicenseAsync(true);
                        try { _upInfo = await LicenseHelper.CheckForUpdateAsync(_currentVersion); if (_upInfo != null) StandardAddInServer._cachedUpdateInfo = _upInfo; } catch { }
                    }
                    else
                    {
                        _licInfo = await LicenseHelper.CheckLicenseAsync(false);
                        try { _upInfo = await LicenseHelper.CheckForUpdateAsync(_currentVersion); if (_upInfo != null) StandardAddInServer._cachedUpdateInfo = _upInfo; } catch { }
                    }

                    if (_licInfo.LastCheckDate == DateTime.MinValue)
                        _licInfo.Message = LanguageManager.L("ABOUT_NOT_CHECKED");

                    UpdateUI(_hwId, _currentVersion, _licInfo, _upInfo);
                }
                catch (Exception ex)
                {
                    LicenseHelper.WriteLog("Lỗi khi load license data", ex);
                    lblInfo.Text = "Error: " + ex.Message;
                }
            }

            private void UpdateUI(string hwId, string version, LicenseHelper.LicenseInfo licInfo, LicenseHelper.UpdateInfo upInfo)
            {
                string statusText;
                if (licInfo.Status == LicenseHelper.LicenseStatus.Active)
                {
                    if (licInfo.IsTrial)
                    {
                        string timeLeft = !string.IsNullOrEmpty(licInfo.TrialTimeLeftDisplay)
                            ? licInfo.TrialTimeLeftDisplay
                            : LicenseHelper.FormatTimeLeft(licInfo.TrialMinutesLeft, StandardAddInServer.IsVietnamese);
                        statusText = LanguageManager.L("ABOUT_TRIAL", timeLeft);
                    }
                    else
                        statusText = LanguageManager.L("ABOUT_ACTIVE");
                }
                else
                {
                    if (licInfo.IsTrial && licInfo.TrialMinutesLeft <= 0)
                        statusText = LanguageManager.L("ABOUT_TRIAL_EXPIRED");
                    else
                        statusText = LanguageManager.L("ABOUT_NOT_ACTIVATED");
                }

                string checkInfo = (licInfo.LastCheckDate.Year > 2000) ? $"\nLast Check: {licInfo.LastCheckDate:dd/MM/yyyy HH:mm}" : "";

                string expirationDisplay = licInfo.ExpirationDate;
                if (!string.IsNullOrEmpty(licInfo.ExpirationDate) && DateTime.TryParse(licInfo.ExpirationDate, out DateTime expDate))
                {
                    double minutesLeft = (expDate - DateTime.Now).TotalMinutes;
                    string timeText;
                    if (minutesLeft >= 0)
                        timeText = LanguageManager.L("ABOUT_TIME_LEFT", LicenseHelper.FormatTimeLeft((int)minutesLeft, StandardAddInServer.IsVietnamese));
                    else
                        timeText = LanguageManager.L("ABOUT_EXPIRED");
                    expirationDisplay = $"{licInfo.ExpirationDate} ({timeText})";
                }
                else if (licInfo.ExpirationDate.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    expirationDisplay = "Lifetime";
                }

                string lblTitle = $"REPLACE FILE ADD-IN\n" +
                    LanguageManager.L("ABOUT_VERSION", version) + "\n--------------------\n" +
                    LanguageManager.L("ABOUT_HWID", hwId) + "\n" +
                    LanguageManager.L("ABOUT_STATUS", statusText) + "\n" +
                    LanguageManager.L("ABOUT_EXPIRES", expirationDisplay) + checkInfo;

                lblInfo.Text = lblTitle;

                if (upInfo != null)
                {
                    if (!string.IsNullOrEmpty(upInfo.HelpUrl)) this.helpUrl = upInfo.HelpUrl;
                    if (!string.IsNullOrEmpty(upInfo.OtherAddinUrl)) this.otherAddinUrl = upInfo.OtherAddinUrl;

                    if (upInfo.HasUpdate && !string.IsNullOrEmpty(upInfo.DownloadUrl))
                    {
                        string forceText = upInfo.ForceUpdate ? LanguageManager.L("ABOUT_REQUIRED") : "";
                        lnkUpdateNotify.Text = LanguageManager.L("ABOUT_NEW_VERSION", upInfo.NewVersion, forceText);

                        lnkUpdateNotify.LinkColor = upInfo.ForceUpdate ? System.Drawing.Color.Red : System.Drawing.Color.Blue;
                        this.downloadUrl = upInfo.DownloadUrl;
                        btnUpdate.Text = LanguageManager.L("ABOUT_DOWNLOAD");
                        btnUpdate.BackColor = upInfo.ForceUpdate ? System.Drawing.Color.OrangeRed : System.Drawing.Color.LightGreen;
                    }
                    else
                    {
                        lnkUpdateNotify.Text = LanguageManager.L("ABOUT_LATEST_VERSION");
                        lnkUpdateNotify.LinkColor = System.Drawing.Color.Green;
                        this.downloadUrl = "";
                        btnUpdate.Text = LanguageManager.L("ABOUT_CHECK_UPDATE");
                        btnUpdate.BackColor = System.Drawing.Color.WhiteSmoke;
                    }
                }
                else
                {
                    lnkUpdateNotify.Text = LanguageManager.L("ABOUT_CHECKING");
                }

                if (licInfo.Status == LicenseHelper.LicenseStatus.Active)
                {
                    btnActive.Text = LanguageManager.L("ABOUT_RECHECK");
                    btnActive.BackColor = System.Drawing.Color.LightGreen;
                    btnActive.Enabled = true;
                }
                else
                {
                    btnActive.Text = LanguageManager.L("ABOUT_ACTIVATE");
                    btnActive.BackColor = System.Drawing.Color.LightYellow;
                    btnActive.Enabled = true;
                }
            }

            private async Task OnUpdateClick()
            {
                if (!string.IsNullOrEmpty(downloadUrl)) { OpenUrl(downloadUrl); return; }
                btnUpdate.Text = LanguageManager.L("ABOUT_CHECKING");
                btnUpdate.Enabled = false;
                try
                {
                    _currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                    _upInfo = await LicenseHelper.CheckForUpdateAsync(_currentVersion);
                    _licInfo = await LicenseHelper.CheckLicenseAsync(false);

                    if (_upInfo != null) StandardAddInServer._cachedUpdateInfo = _upInfo;

                    UpdateUI(LicenseHelper.GetHardwareId(), _currentVersion, _licInfo, _upInfo);
                }
                catch (Exception ex) { lnkUpdateNotify.Text = "Error: " + ex.Message; }
                finally { btnUpdate.Enabled = true; if (string.IsNullOrEmpty(downloadUrl)) btnUpdate.Text = LanguageManager.L("ABOUT_CHECK_UPDATE"); }
            }

            private async Task OnActiveClick()
            {
                btnActive.Text = LanguageManager.L("ABOUT_CHECKING");
                btnActive.Enabled = false;
                try
                {
                    _licInfo = await LicenseHelper.CheckLicenseAsync(true);
                    _hwId = LicenseHelper.GetHardwareId();
                    _currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                    UpdateUI(_hwId, _currentVersion, _licInfo, null);

                    if (_licInfo.Status == LicenseHelper.LicenseStatus.Active)
                    {
                        LicenseHelper.ResetLicenseState();

                        MessageBox.Show(
                            LanguageManager.L("ABOUT_LICENSE_VALID_MSG", _licInfo.ExpirationDate),
                            LanguageManager.L("MSG_SUCCESS"),
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        LicenseHelper.CopyHardwareIdToClipboard();
                        MessageBox.Show(
                            LanguageManager.L("ABOUT_LICENSE_INVALID_MSG", _licInfo.Message, _hwId),
                            LanguageManager.L("TITLE_INFO"),
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally { btnActive.Enabled = true; }
            }
        }
    }
}