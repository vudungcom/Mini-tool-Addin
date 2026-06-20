using Inventor;
using OpenCadDrawingAddin.Logic;
using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

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
        private ButtonDefinition m_bomButton;        // [NEW] Bom format button
        private ButtonDefinition m_nameUpdateButton; // [NEW] Name Update button
        private ButtonDefinition m_saveIdwButton;    // [NEW] Save IDW button
        private ButtonDefinition m_checkRefButton;   // [NEW] Check Reference button
        private ButtonDefinition m_createDwgButton;  // [NEW v1.2] Create DWG button
        private ButtonDefinition m_copyComponentButton;  // [NEW v1.3] Copy Component (cross-screen)
        private ButtonDefinition m_pasteComponentButton; // [NEW v1.3] Paste Component (cross-screen)
        private ButtonDefinition m_autoHoleNoteButton;   // [NEW v1.5] Auto Hole Note button
        private ApplicationEvents m_appEvents;           // [NEW v1.4] IDW Auto Check event
        public static Inventor.Application InventorApplication { get; private set; }

        public static bool IsVietnamese => LanguageManager.CurrentLanguage == "VN";

        public static string Title_Info => LanguageManager.L("TITLE_INFO");
        public static string Title_Error => LanguageManager.L("TITLE_ERROR");
        public static string Title_Warning => LanguageManager.L("TITLE_WARNING");

        // ============================================================
        // ICON HELPER
        // ============================================================

        private static object MakeEmojiIcon(string emoji, int size)
        {
            try
            {
                var bmp = new Bitmap(size, size);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(System.Drawing.Color.Transparent);
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    float fontSize = size * 0.65f;
                    using (var font = new Font("Segoe UI Emoji", fontSize, GraphicsUnit.Pixel))
                    {
                        var sf = new StringFormat
                        {
                            Alignment = StringAlignment.Center,
                            LineAlignment = StringAlignment.Center
                        };
                        g.DrawString(emoji, font, Brushes.Black, new RectangleF(0, 0, size, size), sf);
                    }
                }
                return AxHostConverter.GetIPictureDispFromPicture(bmp);
            }
            catch { return null; }
        }

        private static object _iconOpenCad16, _iconOpenCad32;
        private static object _iconBom16, _iconBom32;
        private static object _iconNameUpdate16, _iconNameUpdate32;
        private static object _iconSaveIdw16, _iconSaveIdw32;
        private static object _iconCheckRef16, _iconCheckRef32;
        private static object _iconCreateDwg16, _iconCreateDwg32;
        private static object _iconCopy16, _iconCopy32;
        private static object _iconPlace16, _iconPlace32;
        private static object _iconSettings16, _iconSettings32;
        private static object _iconAbout16, _iconAbout32;
        private static object _iconAutoHole16, _iconAutoHole32;  // [NEW v1.5]

        private static void EnsureIcons()
        {
            _iconOpenCad16 = MakeEmojiIcon("📐", 16); _iconOpenCad32 = MakeEmojiIcon("📐", 32);
            _iconBom16 = MakeEmojiIcon("📋", 16); _iconBom32 = MakeEmojiIcon("📋", 32);
            _iconNameUpdate16 = MakeEmojiIcon("🏷️", 16); _iconNameUpdate32 = MakeEmojiIcon("🏷️", 32);
            _iconSaveIdw16 = MakeEmojiIcon("💾", 16); _iconSaveIdw32 = MakeEmojiIcon("💾", 32);
            _iconCheckRef16 = MakeEmojiIcon("🔍", 16); _iconCheckRef32 = MakeEmojiIcon("🔍", 32);
            _iconCreateDwg16 = MakeEmojiIcon("📤", 16); _iconCreateDwg32 = MakeEmojiIcon("📤", 32);
            _iconCopy16 = MakeEmojiIcon("📌", 16); _iconCopy32 = MakeEmojiIcon("📌", 32);
            _iconPlace16 = MakeEmojiIcon("📍", 16); _iconPlace32 = MakeEmojiIcon("📍", 32);
            _iconSettings16 = MakeEmojiIcon("⚙️", 16); _iconSettings32 = MakeEmojiIcon("⚙️", 32);
            _iconAbout16 = MakeEmojiIcon("ℹ️", 16); _iconAbout32 = MakeEmojiIcon("ℹ️", 32);
            // [NEW v1.5] icon Auto Hole Note
            _iconAutoHole16 = MakeEmojiIcon("⊙", 16); _iconAutoHole32 = MakeEmojiIcon("⊙", 32);
        }

        private class AxHostConverter : System.Windows.Forms.AxHost
        {
            private AxHostConverter() : base("") { }
            public new static object GetIPictureDispFromPicture(Image image)
                => System.Windows.Forms.AxHost.GetIPictureDispFromPicture(image);
        }

        public static LicenseHelper.UpdateInfo _cachedUpdateInfo = null;
        private static bool _isUpdateCheckRunning = false;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            m_inventorApplication = addInSiteObject.Application;
            InventorApplication = m_inventorApplication; // [NEW v1.5]

            try { OpenCadSettings.LoadAndApplyLanguage(); } catch { }
            try { CreateUserInterface(); } catch { }

            try
            {
                m_appEvents = m_inventorApplication.ApplicationEvents;
                m_appEvents.OnOpenDocument += OnOpenDocument_IdwAutoCheck;
            }
            catch { }

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
            if (m_bomButton != null) { m_bomButton.Delete(); m_bomButton = null; }
            if (m_nameUpdateButton != null) { m_nameUpdateButton.Delete(); m_nameUpdateButton = null; }
            if (m_saveIdwButton != null) { m_saveIdwButton.Delete(); m_saveIdwButton = null; }
            if (m_checkRefButton != null) { m_checkRefButton.Delete(); m_checkRefButton = null; }
            if (m_createDwgButton != null) { m_createDwgButton.Delete(); m_createDwgButton = null; }
            if (m_copyComponentButton != null) { m_copyComponentButton.Delete(); m_copyComponentButton = null; }
            if (m_pasteComponentButton != null) { m_pasteComponentButton.Delete(); m_pasteComponentButton = null; }
            // [NEW v1.5]
            if (m_autoHoleNoteButton != null) { m_autoHoleNoteButton.Delete(); m_autoHoleNoteButton = null; }
            try { if (m_appEvents != null) { m_appEvents.OnOpenDocument -= OnOpenDocument_IdwAutoCheck; m_appEvents = null; } } catch { }
            if (m_inventorApplication != null) { Marshal.ReleaseComObject(m_inventorApplication); m_inventorApplication = null; }
            GC.Collect();
        }

        public object Automation => null;
        public void ExecuteCommand(int commandID) { }

        private void CreateUserInterface()
        {
            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

            EnsureIcons();

            m_myButton = controlDefs.AddButtonDefinition(
                "Open CAD", "OpenCadCmd", CommandTypesEnum.kShapeEditCmdType,
                GenerateClientId("OpenCadCmd"),
                "Open corresponding CAD drawing file.",
                "Open CAD", _iconOpenCad16, _iconOpenCad32);
            m_myButton.OnExecute += m_myButton_OnExecute;

            m_bomButton = controlDefs.AddButtonDefinition(
                "BOM Format", "OCDABomCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDABomCmd"),
                "Apply BOM customization from XML file.",
                "BOM Format", _iconBom16, _iconBom32);
            m_bomButton.OnExecute += m_bomButton_OnExecute;

            m_nameUpdateButton = controlDefs.AddButtonDefinition(
                "Name Update", "OCDANameUpdateCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDANameUpdateCmd"),
                "Update display names of parts/occurrences from file names.",
                "Name Update", _iconNameUpdate16, _iconNameUpdate32);
            m_nameUpdateButton.OnExecute += m_nameUpdateButton_OnExecute;

            m_saveIdwButton = controlDefs.AddButtonDefinition(
                "Save IDW", "OCDASaveIdwCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDASaveIdwCmd"),
                "Save drawing (.idw) to the same folder as the model file.",
                "Save IDW", _iconSaveIdw16, _iconSaveIdw32);
            m_saveIdwButton.OnExecute += m_saveIdwButton_OnExecute;

            m_checkRefButton = controlDefs.AddButtonDefinition(
                "Check Ref", "OCDACheckRefCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDACheckRefCmd"),
                "Scan assembly for occurrences with BOM Structure = Reference.",
                "Check Reference", _iconCheckRef16, _iconCheckRef32);
            m_checkRefButton.OnExecute += m_checkRefButton_OnExecute;

            m_createDwgButton = controlDefs.AddButtonDefinition(
                "Create DWG", "OCDACreateDwgCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDACreateDwgCmd"),
                "Export IDW files to DWG from Assembly.",
                "Create DWG", _iconCreateDwg16, _iconCreateDwg32);
            m_createDwgButton.OnExecute += m_createDwgButton_OnExecute;

            m_copyComponentButton = controlDefs.AddButtonDefinition(
                "Copy", "OCDACopyComponentCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDACopyComponentCmd"),
                "Copy current component path to clipboard (for cross-screen place).",
                "Copy Component", _iconCopy16, _iconCopy32);
            m_copyComponentButton.OnExecute += m_copyComponentButton_OnExecute;

            m_pasteComponentButton = controlDefs.AddButtonDefinition(
                "Place", "OCDAPasteComponentCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDAPasteComponentCmd"),
                "Insert copied component into the current Assembly (cross-screen).",
                "Place Component", _iconPlace16, _iconPlace32);
            m_pasteComponentButton.OnExecute += m_pasteComponentButton_OnExecute;

            // [NEW v1.5] Auto Hole Note button (Drawing ribbon)
            m_autoHoleNoteButton = controlDefs.AddButtonDefinition(
                "Auto Hole", "OCDAAutoHoleNoteCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDAAutoHoleNoteCmd"),
                "Auto fill hole notes (M-tap / Phi) from STEP drawing view.",
                "Auto Hole Note", _iconAutoHole16, _iconAutoHole32);
            m_autoHoleNoteButton.OnExecute += m_autoHoleNoteButton_OnExecute;

            m_settingsButton = controlDefs.AddButtonDefinition(
                "Settings", "OCDASettingsCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDASettingsCmd"),
                LanguageManager.L("ABOUT_CONFIGURE_CLEANING"),
                "Settings", _iconSettings16, _iconSettings32);
            m_settingsButton.OnExecute += m_settingsButton_OnExecute;

            m_aboutButton = controlDefs.AddButtonDefinition(
                "About / License", "OCDAAboutCmd", CommandTypesEnum.kQueryOnlyCmdType,
                GenerateClientId("OCDAAboutCmd"),
                "Information & Activation.",
                "About Author",
                _iconAbout16, _iconAbout32);
            m_aboutButton.OnExecute += m_aboutButton_OnExecute;

            AddPanelToRibbon("Assembly", "id_TabAssemble", "{A7F3B8D1-2C4E-4A9F-8E2B-6D1F5C3A9E72}");
            AddPanelToRibbon("Part", "id_TabModel", "{A7F3B8D1-2C4E-4A9F-8E2B-6D1F5C3A9E73}");
            AddDrawingRibbonPanel();
        }

        private void AddPanelToRibbon(string ribbonName, string tabId, string panelGuid, bool drawingOnly = false)
        {
            try
            {
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons[ribbonName];
                RibbonTab tab = ribbon.RibbonTabs[tabId];
                string panelId = "id_Panel_OpenCadTools_" + ribbonName;
                RibbonPanel panel = null;
                try { panel = tab.RibbonPanels.Add("Mini Tool", panelId, panelGuid); }
                catch { try { panel = tab.RibbonPanels[panelId]; } catch { } }

                if (panel == null) return;

                if (drawingOnly)
                {
                    if (!ButtonExists(panel, m_saveIdwButton)) panel.CommandControls.AddButton(m_saveIdwButton, false);
                    panel.CommandControls.AddSeparator();
                    if (!ButtonExists(panel, m_settingsButton)) panel.CommandControls.AddButton(m_settingsButton, false);
                    if (!ButtonExists(panel, m_aboutButton)) panel.CommandControls.AddButton(m_aboutButton, false);
                }
                else
                {
                    if (!ButtonExists(panel, m_myButton)) panel.CommandControls.AddButton(m_myButton, false);
                    if (!ButtonExists(panel, m_bomButton)) panel.CommandControls.AddButton(m_bomButton, false);
                    if (!ButtonExists(panel, m_nameUpdateButton)) panel.CommandControls.AddButton(m_nameUpdateButton, false);
                    if (!ButtonExists(panel, m_saveIdwButton)) panel.CommandControls.AddButton(m_saveIdwButton, false);
                    if (!ButtonExists(panel, m_checkRefButton)) panel.CommandControls.AddButton(m_checkRefButton, false);
                    if (!ButtonExists(panel, m_createDwgButton)) panel.CommandControls.AddButton(m_createDwgButton, false);

                    if (!ButtonExists(panel, m_copyComponentButton)) panel.CommandControls.AddButton(m_copyComponentButton, false);
                    if (!ButtonExists(panel, m_pasteComponentButton)) panel.CommandControls.AddButton(m_pasteComponentButton, false);

                    panel.CommandControls.AddSeparator();

                    if (!ButtonExists(panel, m_settingsButton)) panel.CommandControls.AddButton(m_settingsButton, false);
                    if (!ButtonExists(panel, m_aboutButton)) panel.CommandControls.AddButton(m_aboutButton, false);
                }
            }
            catch { }
        }

        private void AddDrawingRibbonPanel()
        {
            try
            {
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons["Drawing"];
                string panelId = "id_Panel_OpenCadTools_Drawing";
                string panelGuid = "{A7F3B8D1-2C4E-4A9F-8E2B-6D1F5C3A9E74}";

                string[] tabIds = {
                    "id_TabPlaceViews",
                    "id_TabAnnotate",
                    "id_TabTools"
                };

                RibbonPanel panel = null;
                foreach (string tabId in tabIds)
                {
                    try
                    {
                        RibbonTab tab = ribbon.RibbonTabs[tabId];
                        try { panel = tab.RibbonPanels.Add("Mini Tool", panelId, panelGuid); }
                        catch { try { panel = tab.RibbonPanels[panelId]; } catch { } }

                        if (panel != null) break;
                    }
                    catch { continue; }
                }

                if (panel == null) return;

                // Drawing panel: Save IDW + Create DWG + Auto Hole Note + separator + Settings/About
                if (!ButtonExists(panel, m_saveIdwButton)) panel.CommandControls.AddButton(m_saveIdwButton, false);
                if (!ButtonExists(panel, m_createDwgButton)) panel.CommandControls.AddButton(m_createDwgButton, false);
                // [NEW v1.5]
                if (!ButtonExists(panel, m_autoHoleNoteButton)) panel.CommandControls.AddButton(m_autoHoleNoteButton, false);
                panel.CommandControls.AddSeparator();
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

        private void m_createDwgButton_OnExecute(NameValueMap Context)
        {
            try
            {
                var logic = new CreateDwgLogic(m_inventorApplication);
                Document activeDoc = m_inventorApplication.ActiveDocument;

                if (activeDoc != null && activeDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    logic.RunCreateDwgFromIdw();
                else
                    logic.RunCreateDwg();
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error running Create DWG", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message,
                    LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void m_copyComponentButton_OnExecute(NameValueMap Context)
        {
            try
            {
                var logic = new CrossScreenCopyPasteLogic(m_inventorApplication);
                logic.CopyComponent();
            }
            catch (Exception ex) { LicenseHelper.WriteLog("Error running Copy Component", ex); }
        }

        private void m_pasteComponentButton_OnExecute(NameValueMap Context)
        {
            try
            {
                ComponentOccurrence preSelected = null;
                try
                {
                    var doc = m_inventorApplication.ActiveDocument;
                    if (doc?.SelectSet != null && doc.SelectSet.Count >= 1)
                        preSelected = CrossScreenCopyPasteLogic.GetOccurrenceFromObject(doc.SelectSet[1]);
                }
                catch { }

                var logic = new CrossScreenCopyPasteLogic(m_inventorApplication);
                logic.PasteComponent(preSelected);
            }
            catch (Exception ex) { LicenseHelper.WriteLog("Error running Paste Component", ex); }
        }

        // [NEW v1.5] Handler cho Auto Hole Note button
        private void m_autoHoleNoteButton_OnExecute(NameValueMap Context)
        {
            try
            {
                var drawDoc = m_inventorApplication.ActiveDocument as DrawingDocument;
                if (drawDoc == null)
                {
                    MessageBox.Show("Hay mo file IDW truoc khi chay Auto Hole Note.",
                        "Auto Hole Note", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DrawingView view = null;
                try
                {
                    view = m_inventorApplication.CommandManager.Pick(
                        SelectionFilterEnum.kDrawingViewFilter,
                        "Click vao view can quet lo:") as DrawingView;
                }
                catch { return; }

                if (view == null) return;

                // Load thong so user da chinh trong tab Settings
                var settings = OpenCadSettings.Load();
                var logic = new AutoHoleNoteLogic(m_inventorApplication);
                logic.Run(view,
                    settings.HoleNoteTextHeightMm,
                    settings.HoleNoteClusterRadiusMm,
                    settings.HoleNoteTolTap);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error running Auto Hole Note", ex);
                MessageBox.Show("Loi: " + ex.Message,
                    "Auto Hole Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void m_nameUpdateButton_OnExecute(NameValueMap Context)
        {
            try
            {
                Document activeDoc = m_inventorApplication.ActiveDocument;

                if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    PartDocument doc = (PartDocument)activeDoc;
                    string newName = doc.FullFileName.Substring(doc.FullFileName.LastIndexOf("\\") + 1);
                    newName = newName.Substring(0, newName.LastIndexOf("."));
                    doc.DisplayName = newName;
                    MessageBox.Show(LanguageManager.L("MSG_NAME_UPDATED", newName), LanguageManager.L("TITLE_NAME_UPDATE"),
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    AssemblyDocument adoc = (AssemblyDocument)activeDoc;
                    NameUpdateReplace(adoc.ComponentDefinition.Occurrences);
                    MessageBox.Show(LanguageManager.L("MSG_NAME_UPDATE_ALL_DONE"), LanguageManager.L("TITLE_NAME_UPDATE"),
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(LanguageManager.L("MSG_NAME_UPDATE_UNSUPPORTED"), LanguageManager.L("TITLE_NAME_UPDATE"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error running Name Update", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message, LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string[] _nameUpdatePaths = null;
        private int[] _nameUpdateIndex = null;
        private int _nameUpdateCount = 0;

        private void NameUpdateReplace(ComponentOccurrences occs)
        {
            bool isRoot = (_nameUpdatePaths == null);
            if (isRoot)
            {
                _nameUpdatePaths = new string[1000];
                _nameUpdateIndex = new int[1000];
                _nameUpdateCount = 0;
            }

            foreach (ComponentOccurrence occ in occs)
            {
                try
                {
                    Document doc = (Document)occ.Definition.Document;
                    string newName = doc.FullFileName.Substring(doc.FullFileName.LastIndexOf("\\") + 1);
                    newName = newName.Substring(0, newName.LastIndexOf("."));
                    if (string.IsNullOrEmpty(newName)) newName = "empty";

                    int idx = System.Array.IndexOf(_nameUpdatePaths, doc.FullFileName);
                    if (idx < 0)
                    {
                        _nameUpdatePaths[_nameUpdateCount] = doc.FullFileName;
                        _nameUpdateIndex[_nameUpdateCount] = 1;
                        occ.Name = newName + ":1";
                        _nameUpdateCount++;
                    }
                    else
                    {
                        _nameUpdateIndex[idx]++;
                        occ.Name = newName + ":" + _nameUpdateIndex[idx].ToString();
                    }

                    if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        AssemblyDocument subAdoc = (AssemblyDocument)(Document)m_inventorApplication.Documents.Open(doc.FullFileName, false);
                        NameUpdateReplace(subAdoc.ComponentDefinition.Occurrences);
                    }
                }
                catch { }
            }

            if (isRoot) { _nameUpdatePaths = null; _nameUpdateIndex = null; _nameUpdateCount = 0; }
        }

        private void m_saveIdwButton_OnExecute(NameValueMap Context)
        {
            try
            {
                Document activeDoc = m_inventorApplication.ActiveDocument;

                if (activeDoc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
                {
                    MessageBox.Show(LanguageManager.L("MSG_SAVE_IDW_DRAWING_ONLY"), LanguageManager.L("TITLE_SAVE_IDW"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DrawingDocument dwgDoc = (DrawingDocument)activeDoc;

                Document modelDoc;
                string modelPath;
                try
                {
                    modelDoc = (Document)dwgDoc.ActiveSheet.DrawingViews[1]
                                .ReferencedDocumentDescriptor.ReferencedDocument;
                    modelPath = modelDoc.FullFileName;
                }
                catch
                {
                    MessageBox.Show(LanguageManager.L("MSG_SAVE_IDW_NO_MODEL"), LanguageManager.L("TITLE_SAVE_IDW"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string savePath = modelPath.Substring(0, modelPath.Length - 4) + ".idw";

                try
                {
                    dwgDoc.SaveAs(savePath, false);
                    MessageBox.Show(LanguageManager.L("MSG_SAVE_IDW_SAVED", savePath), LanguageManager.L("TITLE_SAVE_IDW"),
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    MessageBox.Show(LanguageManager.L("MSG_SAVE_IDW_FAILED"), LanguageManager.L("TITLE_SAVE_IDW"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error running Save IDW", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message, LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void m_settingsButton_OnExecute(NameValueMap Context) { using (var frm = new SettingsForm()) { frm.ShowDialog(); } }
        private void m_aboutButton_OnExecute(NameValueMap Context) { using (var frm = new AboutBoxForm(true)) { frm.ShowDialog(); } }

        private void m_checkRefButton_OnExecute(NameValueMap Context)
        {
            try
            {
                Document activeDoc = m_inventorApplication.ActiveDocument;

                if (activeDoc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    MessageBox.Show(LanguageManager.L("MSG_CHECK_REF_ASM_ONLY"), LanguageManager.L("TITLE_CHECK_REF"),
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var settings = OpenCadSettings.Load();
                var excludeList = new System.Collections.Generic.List<string>();

                string excludeFile = settings.CheckRefExcludeListPath;
                if (!string.IsNullOrEmpty(excludeFile) && System.IO.File.Exists(excludeFile))
                {
                    foreach (string line in System.IO.File.ReadAllLines(excludeFile, System.Text.Encoding.UTF8))
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                            excludeList.Add(line.Trim().ToLower());
                    }
                }

                string refList = "";
                AssemblyDocument asmDoc = (AssemblyDocument)activeDoc;
                CheckRefOccurrences(asmDoc.ComponentDefinition.Occurrences, ref refList, excludeList);

                if (string.IsNullOrEmpty(refList))
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_CHECK_REF_NO_RESULT"),
                        LanguageManager.L("TITLE_CHECK_REF"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string asmFolder = System.IO.Path.GetDirectoryName(activeDoc.FullFileName);
                    string asmName = System.IO.Path.GetFileNameWithoutExtension(activeDoc.FullFileName);
                    string outPath = System.IO.Path.Combine(asmFolder, asmName + "_ReferenceCheck.txt");

                    using (var sw = new System.IO.StreamWriter(outPath, false, System.Text.Encoding.UTF8))
                    {
                        sw.WriteLine(LanguageManager.L("MSG_CHECK_REF_HEADER"));
                        sw.WriteLine(refList);
                    }

                    try { System.Diagnostics.Process.Start(outPath); }
                    catch
                    {
                        MessageBox.Show(LanguageManager.L("MSG_CHECK_REF_CANNOT_OPEN", outPath), LanguageManager.L("TITLE_CHECK_REF"),
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error running Check Reference", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message, LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckRefOccurrences(
            ComponentOccurrences occs,
            ref string refList,
            System.Collections.Generic.List<string> excludeList)
        {
            foreach (ComponentOccurrence occ in occs)
            {
                try
                {
                    if (occ.BOMStructure == BOMStructureEnum.kReferenceBOMStructure)
                    {
                        string occName = occ.Name.ToLower();
                        Document occDoc = (Document)occ.Definition.Document;
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(occDoc.FullFileName).ToLower();
                        string fileNameEx = System.IO.Path.GetFileName(occDoc.FullFileName).ToLower();

                        bool skip = false;
                        foreach (string ex in excludeList)
                        {
                            if (occName.Contains(ex) || fileName.Contains(ex) || fileNameEx.Contains(ex))
                            {
                                skip = true;
                                break;
                            }
                        }

                        if (!skip)
                            refList += "\r\n" + occ.Name + "   --> " + ((Document)occ.Definition.Document).FullFileName;
                    }

                    if (occ.DefinitionDocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        AssemblyDocument subAsm = (AssemblyDocument)(Document)occ.Definition.Document;
                        CheckRefOccurrences(subAsm.ComponentDefinition.Occurrences, ref refList, excludeList);
                    }
                }
                catch { }
            }
        }

        private void m_bomButton_OnExecute(NameValueMap Context)
        {
            try
            {
                var settings = OpenCadSettings.Load();
                string xmlPath = settings.BomXmlPath;

                if (string.IsNullOrEmpty(xmlPath) || !System.IO.File.Exists(xmlPath))
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_BOM_XML_NOT_SET"),
                        LanguageManager.L("TITLE_WARNING"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Document oADoc = m_inventorApplication.ActiveDocument;
                AssemblyDocument oDoc = null;

                if (oADoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    oDoc = (AssemblyDocument)oADoc;
                }
                else if (oADoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                {
                    try
                    {
                        var dwgDoc = (DrawingDocument)oADoc;
                        oDoc = (AssemblyDocument)dwgDoc.ActiveSheet.DrawingViews[1]
                               .ReferencedDocumentDescriptor.ReferencedDocument;
                    }
                    catch
                    {
                        MessageBox.Show(LanguageManager.L("MSG_BOM_LOD_ERROR"), LanguageManager.L("TITLE_ERROR"),
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show(LanguageManager.L("MSG_BOM_INVALID_DOC"), LanguageManager.L("TITLE_ERROR"),
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                BOM oBOM = oDoc.ComponentDefinition.BOM;
                oBOM.ImportBOMCustomization(xmlPath);
                oBOM.StructuredViewEnabled = true;
                oBOM.StructuredViewFirstLevelOnly = false;
                oBOM.PartsOnlyViewEnabled = true;

                MessageBox.Show(LanguageManager.L("MSG_BOM_SUCCESS"), LanguageManager.L("TITLE_INFO"),
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error running BOM format", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message, LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string GenerateClientId(string cmdName) => "{" + this.GetType().GUID.ToString() + "}+" + cmdName;

        // ============================================================
        // AboutBoxForm
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

                lblInfo.Text = $"Mini-Tool Add-in\n{LanguageManager.L("ABOUT_VERSION", version)}\n--------------------\n{LanguageManager.L("ABOUT_HWID", hwId)}\n{LanguageManager.L("ABOUT_STATUS", status)}\n{LanguageManager.L("ABOUT_EXPIRES", exp)}";

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

                if (upInfo != null)
                {
                    if (!string.IsNullOrEmpty(upInfo.HelpUrl)) helpUrl = upInfo.HelpUrl;
                    if (!string.IsNullOrEmpty(upInfo.OtherAddinUrl)) otherAddinUrl = upInfo.OtherAddinUrl;
                }

                lnkHelp.Text = "?";
                lnkHelp.Visible = !string.IsNullOrEmpty(helpUrl);
            }

            private async Task OnUpdateClick() { if (!string.IsNullOrEmpty(downloadUrl)) OpenUrl(downloadUrl); else await SafeLoadLicenseData(); }
            private async Task OnActiveClick() { btnActive.Enabled = false; try { await SafeLoadLicenseData(); if (_licInfo.Status == LicenseHelper.LicenseStatus.Active) MessageBox.Show(LanguageManager.L("ABOUT_LICENSE_VALID_MSG", _licInfo.ExpirationDate), "Success"); else MessageBox.Show(LanguageManager.L("ABOUT_LICENSE_INVALID_MSG", _licInfo.Message, _hwId), "Info"); } finally { btnActive.Enabled = true; } }
        }

        // [NEW v1.4] Handler tu dong check IDW khi mo document
        private void OnOpenDocument_IdwAutoCheck(
            Document documentObject,
            string fullDocumentName,
            EventTimingEnum beforeOrAfter,
            NameValueMap context,
            out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            try
            {
                if (beforeOrAfter != EventTimingEnum.kAfter) return;
                if (documentObject?.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 1500;
                var docRef = documentObject;
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    try
                    {
                        var logic = new IdwAutoCheckLogic(m_inventorApplication);
                        logic.RunChecks(docRef);
                    }
                    catch { }
                };
                timer.Start();
            }
            catch { }
        }
    }
}