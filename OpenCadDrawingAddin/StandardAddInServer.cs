using Inventor;
using OpenCadDrawingAddin.Logic;
using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenCadDrawingAddin
{
    [Guid("A7F3B8D1-2C4E-4A9F-8E2B-6D1F5C3A9E72")]
    [ProgId("OpenCadDrawingAddin.StandardAddInServer")]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_myButton;
        private ButtonDefinition m_settingsButton;
        private ButtonDefinition m_aboutButton;

        public static bool IsVietnamese => LanguageManager.CurrentLanguage == "VN";

        public static string Title_Info => LanguageManager.L("TITLE_INFO");
        public static string Title_Error => LanguageManager.L("TITLE_ERROR");
        public static string Title_Warning => LanguageManager.L("TITLE_WARNING");

        public static LicenseHelper.UpdateInfo _cachedUpdateInfo = null;
        private static bool _isUpdateCheckRunning = false;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            m_inventorApplication = addInSiteObject.Application;

            try { OpenCadSettings.LoadAndApplyLanguage(); } catch { }
            try { CreateUserInterface(); } catch { }

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

            m_myButton = controlDefs.AddButtonDefinition(
                "Open CAD Drawing", "OpenCadCmd", CommandTypesEnum.kShapeEditCmdType,
                GenerateClientId("OpenCadCmd"),
                "Open corresponding CAD drawing file.",
                "Open CAD Drawing",
                null, null);
            m_myButton.OnExecute += m_myButton_OnExecute;

            m_settingsButton = controlDefs.AddButtonDefinition(
                "Settings", "OCDASettingsCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDASettingsCmd"),
                LanguageManager.L("ABOUT_CONFIGURE_CLEANING"),
                "Settings", null, null);
            m_settingsButton.OnExecute += m_settingsButton_OnExecute;

            m_aboutButton = controlDefs.AddButtonDefinition(
                "About / License", "OCDAAboutCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDAAboutCmd"),
                "Information & Activation.",
                "About Author",
                null, null);
            m_aboutButton.OnExecute += m_aboutButton_OnExecute;

            AddPanelToRibbon("Assembly", "id_TabAssemble", "{A7F3B8D1-2C4E-4A9F-8E2B-6D1F5C3A9E72}");
            AddPanelToRibbon("Part", "id_TabModel", "{A7F3B8D1-2C4E-4A9F-8E2B-6D1F5C3A9E73}");
        }

        private void AddPanelToRibbon(string ribbonName, string tabId, string panelGuid)
        {
            try
            {
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons[ribbonName];
                RibbonTab tab = ribbon.RibbonTabs[tabId];
                string panelId = "id_Panel_OpenCadTools_" + ribbonName;
                RibbonPanel panel = null;
                try { panel = tab.RibbonPanels.Add("CAD Drawing", panelId, panelGuid); }
                catch { try { panel = tab.RibbonPanels[panelId]; } catch { } }

                if (panel == null) return;
                if (!ButtonExists(panel, m_myButton)) panel.CommandControls.AddButton(m_myButton, true);
                if (!ButtonExists(panel, m_settingsButton)) panel.CommandControls.AddButton(m_settingsButton, false);
                if (!ButtonExists(panel, m_aboutButton)) panel.CommandControls.AddButton(m_aboutButton, false);
            }
            catch { }
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
                _cachedUpdateInfo = updateInfo;
                if (updateInfo != null) LicenseHelper.SetForceUpdate(updateInfo.ForceUpdate);
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
                    MessageBox.Show(LanguageManager.L("ABOUT_UPDATE_MSG", _cachedUpdateInfo.NewVersion),
                        LanguageManager.L("ABOUT_UPDATE_REQUIRED"), MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    using (var frm = new AboutBoxForm(true)) { frm.ShowDialog(); }
                    return;
                }
                var logic = new OpenCadLogic(m_inventorApplication);
                logic.RunOpenCad();
            }
            catch (Exception ex) { LicenseHelper.WriteLog("Error running Open CAD Drawing", ex); }
        }

        private void m_aboutButton_OnExecute(NameValueMap Context) { using (var frm = new AboutBoxForm(true)) { frm.ShowDialog(); } }
        private void m_settingsButton_OnExecute(NameValueMap Context) { using (var frm = new SettingsForm()) { frm.ShowDialog(); } }
        private string GenerateClientId(string cmdName) => "{" + this.GetType().GUID.ToString() + "}+" + cmdName;

        // ============================================================
        // AboutBoxForm - ĐÃ FIX LỖI AMBIGUOUS REFERENCE
        // ============================================================
        public class AboutBoxForm : Form
        {
            private Label lblInfo;
            private LinkLabel lnkUpdateNotify;
            private Button btnActive, btnUpdate, btnClose;
            private LinkLabel lnkCopyID, lnkHelp, lnkOtherAddin, lnkViewLog;
            private Label lblAuthor;
            private LinkLabel lblEmail, lnkCopyEmail;
            private LinkLabel lnkTelegramMain, lnkWhatsAppMain, lnkZaloMain;
            private ToolTip toolTipContacts;

            private string downloadUrl = "", helpUrl = "", otherAddinUrl = "";
            private bool _isAutoCheck = false;
            private string _currentVersion = "", _hwId = "";
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
                this.MaximizeBox = false; this.MinimizeBox = false; this.ShowIcon = false;

                toolTipContacts = new ToolTip { AutoPopDelay = 5000, InitialDelay = 500, ShowAlways = true };

                lblInfo = new Label { Location = new System.Drawing.Point(20, 20), Size = new System.Drawing.Size(380, 160), Font = new System.Drawing.Font("Segoe UI", 10) };
                this.Controls.Add(lblInfo);

                lnkHelp = new LinkLabel { Location = new System.Drawing.Point(420, 20), Size = new System.Drawing.Size(50, 25), Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold) };
                lnkHelp.LinkClicked += (s, e) => { if (!string.IsNullOrEmpty(helpUrl)) OpenUrl(helpUrl); };
                this.Controls.Add(lnkHelp);

                lnkCopyID = new LinkLabel { Location = new System.Drawing.Point(400, 70), Size = new System.Drawing.Size(70, 20), Font = new System.Drawing.Font("Segoe UI", 9), LinkColor = System.Drawing.Color.Green };
                lnkCopyID.LinkClicked += (s, e) => { LicenseHelper.CopyHardwareIdToClipboard(); MessageBox.Show(LanguageManager.L("ABOUT_ID_COPIED"), "Info"); };
                this.Controls.Add(lnkCopyID);

                lnkUpdateNotify = new LinkLabel { Location = new System.Drawing.Point(20, 185), Size = new System.Drawing.Size(450, 20), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), LinkColor = System.Drawing.Color.Gray };
                lnkUpdateNotify.LinkClicked += (s, e) => { if (!string.IsNullOrEmpty(downloadUrl)) OpenUrl(downloadUrl); };
                this.Controls.Add(lnkUpdateNotify);

                btnActive = new Button { Location = new System.Drawing.Point(20, 215), Size = new System.Drawing.Size(140, 35), BackColor = System.Drawing.Color.LightYellow };
                btnActive.Click += async (s, e) => await OnActiveClick();
                this.Controls.Add(btnActive);

                btnUpdate = new Button { Location = new System.Drawing.Point(170, 215), Size = new System.Drawing.Size(140, 35) };
                btnUpdate.Click += async (s, e) => await OnUpdateClick();
                this.Controls.Add(btnUpdate);

                btnClose = new Button { Location = new System.Drawing.Point(320, 215), Size = new System.Drawing.Size(140, 35) };
                btnClose.Click += (s, e) => this.Close();
                this.Controls.Add(btnClose);

                Panel pnlAuthor = new Panel { Location = new System.Drawing.Point(0, 275), Size = new System.Drawing.Size(500, 140), BackColor = System.Drawing.Color.WhiteSmoke };
                this.Controls.Add(pnlAuthor);

                lblAuthor = new Label { Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), ForeColor = System.Drawing.Color.DarkSlateGray, Location = new System.Drawing.Point(20, 10), Size = new System.Drawing.Size(200, 18) };
                pnlAuthor.Controls.Add(lblAuthor);

                lblEmail = new LinkLabel { Text = "Email: replacefile.addin@gmail.com", Font = new System.Drawing.Font("Segoe UI", 9), Location = new System.Drawing.Point(20, 30), Size = new System.Drawing.Size(220, 18) };
                lblEmail.LinkClicked += (s, e) => OpenUrl("mailto:replacefile.addin@gmail.com");
                pnlAuthor.Controls.Add(lblEmail);

                lnkCopyEmail = new LinkLabel { Font = new System.Drawing.Font("Segoe UI", 8), Location = new System.Drawing.Point(245, 30), Size = new System.Drawing.Size(50, 18), LinkColor = System.Drawing.Color.DimGray };
                lnkCopyEmail.LinkClicked += (s, e) => { Clipboard.SetText("replacefile.addin@gmail.com"); MessageBox.Show(LanguageManager.L("ABOUT_COPY_EMAIL"), "Info"); };
                pnlAuthor.Controls.Add(lnkCopyEmail);

                SetupSocialLinks(pnlAuthor);

                lnkOtherAddin = new LinkLabel { Location = new System.Drawing.Point(350, 10), Size = new System.Drawing.Size(120, 18), Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), LinkColor = System.Drawing.Color.DarkGreen };
                lnkOtherAddin.LinkClicked += (s, e) => { if (!string.IsNullOrEmpty(otherAddinUrl)) OpenUrl(otherAddinUrl); };
                pnlAuthor.Controls.Add(lnkOtherAddin);

                lnkViewLog = new LinkLabel { Location = new System.Drawing.Point(350, 30), Size = new System.Drawing.Size(120, 18), Font = new System.Drawing.Font("Segoe UI", 9), LinkColor = System.Drawing.Color.Gray };
                lnkViewLog.LinkClicked += (s, e) => ShowErrorLog();
                pnlAuthor.Controls.Add(lnkViewLog);

                UpdateLanguage();
            }

            private void SetupSocialLinks(Panel p)
            {
                lnkTelegramMain = new LinkLabel { Text = "Telegram: t.me/replacefile", Location = new System.Drawing.Point(20, 50), AutoSize = true, LinkColor = System.Drawing.Color.FromArgb(0, 136, 204) };
                lnkTelegramMain.LinkClicked += (s, e) => OpenUrl("https://t.me/replacefile");
                p.Controls.Add(lnkTelegramMain);

                lnkWhatsAppMain = new LinkLabel { Text = "WhatsApp: +84 963 058 858", Location = new System.Drawing.Point(20, 70), AutoSize = true, LinkColor = System.Drawing.Color.FromArgb(37, 211, 102) };
                lnkWhatsAppMain.LinkClicked += (s, e) => OpenUrl("https://wa.me/84963058858");
                p.Controls.Add(lnkWhatsAppMain);

                lnkZaloMain = new LinkLabel { Text = "Zalo: 0963 058 858", Location = new System.Drawing.Point(20, 90), AutoSize = true, LinkColor = System.Drawing.Color.FromArgb(0, 104, 255) };
                lnkZaloMain.LinkClicked += (s, e) => OpenUrl("https://zalo.me/0963058858");
                p.Controls.Add(lnkZaloMain);
            }

            private void UpdateLanguage()
            {
                this.Text = "About & License Manager";
                lnkHelp.Text = LanguageManager.L("ABOUT_HELP");
                lnkCopyID.Text = LanguageManager.L("ABOUT_COPY_ID");
                lnkCopyEmail.Text = LanguageManager.L("ABOUT_COPY_EMAIL");
                lnkOtherAddin.Text = LanguageManager.L("ABOUT_OTHER_ADDIN");
                lnkViewLog.Text = LanguageManager.L("ABOUT_VIEW_LOG");
                btnUpdate.Text = LanguageManager.L("ABOUT_CHECK_UPDATE");
                btnClose.Text = "Close";
                lblAuthor.Text = "Author: Kane";
                if (_licInfo != null) UpdateUI(_hwId, _currentVersion, _licInfo, _upInfo);
                else lblInfo.Text = LanguageManager.L("ABOUT_CHECKING");
            }

            private void OpenUrl(string url) { try { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true }); } catch { } }
            private void ShowErrorLog() { string p = LicenseHelper.GetLogPath(); if (System.IO.File.Exists(p)) System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{p}\""); }

            private async Task SafeLoadLicenseData()
            {
                try
                {
                    _hwId = LicenseHelper.GetHardwareId();
                    _currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                    _licInfo = await LicenseHelper.CheckLicenseAsync(_isAutoCheck);
                    _upInfo = await LicenseHelper.CheckForUpdateAsync(_currentVersion);
                    if (_upInfo != null) StandardAddInServer._cachedUpdateInfo = _upInfo;
                    UpdateUI(_hwId, _currentVersion, _licInfo, _upInfo);
                }
                catch { lblInfo.Text = "Error loading data"; }
            }

            private void UpdateUI(string hwId, string version, LicenseHelper.LicenseInfo licInfo, LicenseHelper.UpdateInfo upInfo)
            {
                string status = (licInfo.Status == LicenseHelper.LicenseStatus.Active) ? (licInfo.IsTrial ? LanguageManager.L("ABOUT_TRIAL", licInfo.TrialTimeLeftDisplay) : LanguageManager.L("ABOUT_ACTIVE")) : (licInfo.IsTrial && licInfo.TrialMinutesLeft <= 0 ? LanguageManager.L("ABOUT_TRIAL_EXPIRED") : LanguageManager.L("ABOUT_NOT_ACTIVATED"));
                string exp = licInfo.ExpirationDate;
                if (!string.IsNullOrEmpty(exp) && DateTime.TryParse(exp, out DateTime dt)) exp += (dt >= DateTime.Now) ? $" ({LanguageManager.L("ABOUT_TIME_LEFT", LicenseHelper.FormatTimeLeft((int)(dt - DateTime.Now).TotalMinutes, StandardAddInServer.IsVietnamese))})" : $" ({LanguageManager.L("ABOUT_EXPIRED")})";

                lblInfo.Text = $"OPEN CAD DRAWING ADD-IN\n{LanguageManager.L("ABOUT_VERSION", version)}\n--------------------\n{LanguageManager.L("ABOUT_HWID", hwId)}\n{LanguageManager.L("ABOUT_STATUS", status)}\n{LanguageManager.L("ABOUT_EXPIRES", exp)}";

                if (upInfo != null && upInfo.HasUpdate)
                {
                    lnkUpdateNotify.Text = LanguageManager.L("ABOUT_NEW_VERSION", upInfo.NewVersion, upInfo.ForceUpdate ? LanguageManager.L("ABOUT_REQUIRED") : "");
                    lnkUpdateNotify.LinkColor = upInfo.ForceUpdate ? System.Drawing.Color.Red : System.Drawing.Color.Blue;
                    downloadUrl = upInfo.DownloadUrl; btnUpdate.Text = LanguageManager.L("ABOUT_DOWNLOAD");
                    btnUpdate.BackColor = upInfo.ForceUpdate ? System.Drawing.Color.OrangeRed : System.Drawing.Color.LightGreen;
                }
                else
                {
                    lnkUpdateNotify.Text = LanguageManager.L("ABOUT_LATEST_VERSION"); lnkUpdateNotify.LinkColor = System.Drawing.Color.Green;
                    downloadUrl = ""; btnUpdate.Text = LanguageManager.L("ABOUT_CHECK_UPDATE"); btnUpdate.BackColor = System.Drawing.Color.WhiteSmoke;
                }

                btnActive.Text = (licInfo.Status == LicenseHelper.LicenseStatus.Active) ? LanguageManager.L("ABOUT_RECHECK") : LanguageManager.L("ABOUT_ACTIVATE");
                btnActive.BackColor = (licInfo.Status == LicenseHelper.LicenseStatus.Active) ? System.Drawing.Color.LightGreen : System.Drawing.Color.LightYellow;
            }

            private async Task OnUpdateClick() { if (!string.IsNullOrEmpty(downloadUrl)) OpenUrl(downloadUrl); else await SafeLoadLicenseData(); }
            private async Task OnActiveClick() { btnActive.Enabled = false; try { await SafeLoadLicenseData(); if (_licInfo.Status == LicenseHelper.LicenseStatus.Active) MessageBox.Show(LanguageManager.L("ABOUT_LICENSE_VALID_MSG", _licInfo.ExpirationDate), "Success"); else MessageBox.Show(LanguageManager.L("ABOUT_LICENSE_INVALID_MSG", _licInfo.Message, _hwId), "Info"); } finally { btnActive.Enabled = true; } }
        }
    }
}