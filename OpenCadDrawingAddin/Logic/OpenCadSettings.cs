using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Settings cho Open CAD Drawing Add-in
    /// Luu/Load qua XML tai AppData
    /// Version 1.5 - Them Auto Hole Note settings
    /// </summary>
    public class OpenCadSettings
    {
        // ============================================================
        // PROPERTIES
        // ============================================================

        /// <summary>
        /// [CHANGED v1.1] Danh sach thu muc chua file CAD (.dwg) - ho tro nhieu thu muc
        /// </summary>
        public List<string> CadFolderPaths { get; set; } = new List<string>();

        /// <summary>
        /// Compatibility: tra ve folder dau tien (dung cho logic cu)
        /// </summary>
        public string CadFolderPath
        {
            get => CadFolderPaths.Count > 0 ? CadFolderPaths[0] : "";
            set
            {
                if (!string.IsNullOrEmpty(value) && !CadFolderPaths.Contains(value))
                {
                    CadFolderPaths.Clear();
                    CadFolderPaths.Add(value);
                }
            }
        }

        public bool UseRevisionSuffix { get; set; } = false;
        public string CadExtension { get; set; } = ".dwg";
        public string Language { get; set; } = "EN";
        public string BomXmlPath { get; set; } = "";
        public string CheckRefExcludeListPath { get; set; } = "";

        // [NEW v1.2] Create DWG
        public string CreateDwgIniPath { get; set; } = "";
        public string CreateDwgListPath { get; set; } = "";
        public string CreateDwgOutputPath { get; set; } = "";

        // [NEW v1.3] Copy-Paste notifications
        public bool ShowCopyNotification { get; set; } = true;
        public bool ShowPasteNotification { get; set; } = true;

        // [NEW v1.4] IDW Auto Check
        public bool AutoCheckDrawingName { get; set; } = true;
        public bool AutoCheckAppearance { get; set; } = true;

        // [NEW v1.5] Auto Hole Note
        public double HoleNoteTextHeightMm { get; set; } = 3.0;
        public double HoleNoteClusterRadiusMm { get; set; } = 30.0;
        public double HoleNoteTolTap { get; set; } = 0.15;

        // [AUTO-SIZE] Neu true, textH & clusterR duoc tinh tu view.Scale luc Run
        // (bo qua gia tri manual o 2 o tren). Base: view scale 1:4 -> text 3mm, cluster 30mm.
        public bool HoleNoteAutoSize { get; set; } = true;

        // ============================================================
        // PATHS
        // ============================================================

        private static readonly string _appFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "InventorAddins");

        private static readonly string _settingsPath = Path.Combine(_appFolder, "ocda_settings.xml");

        // ============================================================
        // LOAD / SAVE
        // ============================================================

        public static OpenCadSettings Load()
        {
            var settings = new OpenCadSettings();
            try
            {
                if (!File.Exists(_settingsPath)) return settings;

                var xml = XDocument.Load(_settingsPath);
                var root = xml.Root;
                if (root == null) return settings;

                // [CHANGED v1.1] Load multi-folder, backward compat voi single folder cu
                var foldersEl = root.Element("CadFolderPaths");
                if (foldersEl != null)
                {
                    foreach (var el in foldersEl.Elements("Folder"))
                    {
                        string p = (string)el ?? "";
                        if (!string.IsNullOrEmpty(p) && !settings.CadFolderPaths.Contains(p))
                            settings.CadFolderPaths.Add(p);
                    }
                }
                else
                {
                    string old = (string)root.Element("CadFolderPath") ?? "";
                    if (!string.IsNullOrEmpty(old)) settings.CadFolderPaths.Add(old);
                }

                settings.UseRevisionSuffix = ParseBool(root.Element("UseRevisionSuffix"), false);
                settings.CadExtension = (string)root.Element("CadExtension") ?? ".dwg";
                settings.Language = (string)root.Element("Language") ?? "EN";
                settings.BomXmlPath = (string)root.Element("BomXmlPath") ?? "";
                settings.CheckRefExcludeListPath = (string)root.Element("CheckRefExcludeListPath") ?? "";

                // [NEW v1.2]
                settings.CreateDwgIniPath = (string)root.Element("CreateDwgIniPath") ?? "";
                settings.CreateDwgListPath = (string)root.Element("CreateDwgListPath") ?? "";
                settings.CreateDwgOutputPath = (string)root.Element("CreateDwgOutputPath") ?? "";

                // [NEW v1.3]
                settings.ShowCopyNotification = ParseBool(root.Element("ShowCopyNotification"), true);
                settings.ShowPasteNotification = ParseBool(root.Element("ShowPasteNotification"), true);

                // [NEW v1.4]
                settings.AutoCheckDrawingName = ParseBool(root.Element("AutoCheckDrawingName"), true);
                settings.AutoCheckAppearance = ParseBool(root.Element("AutoCheckAppearance"), true);

                // [NEW v1.5] Auto Hole Note
                settings.HoleNoteTextHeightMm = ParseDouble(root.Element("HoleNoteTextHeightMm"), 3.0);
                settings.HoleNoteClusterRadiusMm = ParseDouble(root.Element("HoleNoteClusterRadiusMm"), 30.0);
                settings.HoleNoteTolTap = ParseDouble(root.Element("HoleNoteTolTap"), 0.15);

                // [AUTO-SIZE]
                settings.HoleNoteAutoSize = ParseBool(root.Element("HoleNoteAutoSize"), true);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error loading OpenCadSettings", ex);
            }
            return settings;
        }

        public void Save()
        {
            try
            {
                if (!Directory.Exists(_appFolder))
                    Directory.CreateDirectory(_appFolder);

                var xml = new XDocument(
                    new XElement("OpenCadSettings",
                        new XElement("CadFolderPaths",
                            CadFolderPaths.Select(f => new XElement("Folder", f)).ToArray<object>()
                        ),
                        new XElement("UseRevisionSuffix", UseRevisionSuffix),
                        new XElement("CadExtension", CadExtension),
                        new XElement("Language", Language),
                        new XElement("BomXmlPath", BomXmlPath),
                        new XElement("CheckRefExcludeListPath", CheckRefExcludeListPath),
                        // [NEW v1.2]
                        new XElement("CreateDwgIniPath", CreateDwgIniPath),
                        new XElement("CreateDwgListPath", CreateDwgListPath),
                        new XElement("CreateDwgOutputPath", CreateDwgOutputPath),
                        // [NEW v1.3]
                        new XElement("ShowCopyNotification", ShowCopyNotification),
                        new XElement("ShowPasteNotification", ShowPasteNotification),
                        // [NEW v1.4]
                        new XElement("AutoCheckDrawingName", AutoCheckDrawingName),
                        new XElement("AutoCheckAppearance", AutoCheckAppearance),
                        // [NEW v1.5] Auto Hole Note
                        new XElement("HoleNoteTextHeightMm", HoleNoteTextHeightMm),
                        new XElement("HoleNoteClusterRadiusMm", HoleNoteClusterRadiusMm),
                        new XElement("HoleNoteTolTap", HoleNoteTolTap),
                        // [AUTO-SIZE]
                        new XElement("HoleNoteAutoSize", HoleNoteAutoSize)
                    )
                );
                xml.Save(_settingsPath);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error saving OpenCadSettings", ex);
            }
        }

        public static void LoadAndApplyLanguage()
        {
            try
            {
                var s = Load();
                if (!string.IsNullOrEmpty(s.Language))
                    LanguageManager.CurrentLanguage = s.Language;
            }
            catch { }
        }

        // ============================================================
        // HELPERS
        // ============================================================

        private static bool ParseBool(XElement el, bool defaultValue)
        {
            if (el == null) return defaultValue;
            return bool.TryParse((string)el, out bool result) ? result : defaultValue;
        }

        private static double ParseDouble(XElement el, double defaultValue)
        {
            if (el == null) return defaultValue;
            return double.TryParse((string)el,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out double result) ? result : defaultValue;
        }
    }
}