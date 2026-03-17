using System;
using System.IO;
using System.Xml.Linq;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Settings cho Open CAD Drawing Add-in
    /// Lưu/Load qua XML tại AppData
    /// Version 1.0
    /// </summary>
    public class OpenCadSettings
    {
        // ============================================================
        // PROPERTIES
        // ============================================================

        /// <summary>
        /// Thư mục chứa file CAD (.dwg)
        /// </summary>
        public string CadFolderPath { get; set; } = "";

        /// <summary>
        /// Có thêm số sửa đổi vào tên file không
        /// VD: true  → file-1.dwg (nếu revision = 1)
        ///     false → file.dwg (bỏ qua revision)
        /// </summary>
        public bool UseRevisionSuffix { get; set; } = false;

        /// <summary>
        /// Extension của file CAD cần mở (mặc định .dwg)
        /// </summary>
        public string CadExtension { get; set; } = ".dwg";

        /// <summary>
        /// Ngôn ngữ hiển thị
        /// </summary>
        public string Language { get; set; } = "EN";

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

        /// <summary>
        /// Load settings từ file XML. Trả về default nếu chưa có.
        /// </summary>
        public static OpenCadSettings Load()
        {
            var settings = new OpenCadSettings();
            try
            {
                if (!File.Exists(_settingsPath)) return settings;

                var xml = XDocument.Load(_settingsPath);
                var root = xml.Root;
                if (root == null) return settings;

                settings.CadFolderPath  = (string)root.Element("CadFolderPath")  ?? "";
                settings.UseRevisionSuffix = ParseBool(root.Element("UseRevisionSuffix"), false);
                settings.CadExtension   = (string)root.Element("CadExtension")   ?? ".dwg";
                settings.Language       = (string)root.Element("Language")        ?? "EN";
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error loading OpenCadSettings", ex);
            }
            return settings;
        }

        /// <summary>
        /// Lưu settings ra file XML
        /// </summary>
        public void Save()
        {
            try
            {
                if (!Directory.Exists(_appFolder))
                    Directory.CreateDirectory(_appFolder);

                var xml = new XDocument(
                    new XElement("OpenCadSettings",
                        new XElement("CadFolderPath",    CadFolderPath),
                        new XElement("UseRevisionSuffix", UseRevisionSuffix),
                        new XElement("CadExtension",     CadExtension),
                        new XElement("Language",         Language)
                    )
                );
                xml.Save(_settingsPath);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error saving OpenCadSettings", ex);
            }
        }

        /// <summary>
        /// Load language và áp dụng vào LanguageManager
        /// </summary>
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
            bool result;
            return bool.TryParse((string)el, out result) ? result : defaultValue;
        }
    }
}
