using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Settings cho Open CAD Drawing Add-in
    /// Lưu/Load qua XML tại AppData
    /// Version 1.2 - Thêm Create DWG settings
    /// </summary>
    public class OpenCadSettings
    {
        // ============================================================
        // PROPERTIES
        // ============================================================

        /// <summary>
        /// [CHANGED v1.1] Danh sách thư mục chứa file CAD (.dwg) - hỗ trợ nhiều thư mục
        /// </summary>
        public List<string> CadFolderPaths { get; set; } = new List<string>();

        /// <summary>
        /// Compatibility: trả về folder đầu tiên (dùng cho logic cũ)
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

        /// <summary>
        /// [NEW] Đường dẫn file .xml cho BOM customization
        /// </summary>
        public string BomXmlPath { get; set; } = "";

        /// <summary>
        /// [NEW] Đường dẫn file exclude_list.txt cho Check Reference
        /// </summary>
        public string CheckRefExcludeListPath { get; set; } = "";

        /// <summary>
        /// [NEW v1.2] Đường dẫn file DWGExport.ini dùng cho Create DWG
        /// </summary>
        public string CreateDwgIniPath { get; set; } = "";

        /// <summary>
        /// [NEW v1.2] Đường dẫn file danh sách (.txt) dùng khi export theo list
        /// </summary>
        public string CreateDwgListPath { get; set; } = "";

        /// <summary>
        /// [NEW v1.2] Đường dẫn thư mục lưu file DWG xuất ra
        /// </summary>
        public string CreateDwgOutputPath { get; set; } = "";

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

                // [CHANGED v1.1] Load multi-folder, backward compat với single folder cũ
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
                    // Đọc format cũ (single path) để migrate
                    string old = (string)root.Element("CadFolderPath") ?? "";
                    if (!string.IsNullOrEmpty(old)) settings.CadFolderPaths.Add(old);
                }
                settings.UseRevisionSuffix = ParseBool(root.Element("UseRevisionSuffix"), false);
                settings.CadExtension = (string)root.Element("CadExtension") ?? ".dwg";
                settings.Language = (string)root.Element("Language") ?? "EN";
                settings.BomXmlPath = (string)root.Element("BomXmlPath") ?? "";
                settings.CheckRefExcludeListPath = (string)root.Element("CheckRefExcludeListPath") ?? "";

                // [NEW v1.2] Create DWG settings
                settings.CreateDwgIniPath = (string)root.Element("CreateDwgIniPath") ?? "";
                settings.CreateDwgListPath = (string)root.Element("CreateDwgListPath") ?? "";
                settings.CreateDwgOutputPath = (string)root.Element("CreateDwgOutputPath") ?? "";
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
                        // [CHANGED v1.1] Lưu multi-folder
                        new XElement("CadFolderPaths",
                            CadFolderPaths.Select(f => new XElement("Folder", f)).ToArray<object>()
                        ),
                        new XElement("UseRevisionSuffix", UseRevisionSuffix),
                        new XElement("CadExtension", CadExtension),
                        new XElement("Language", Language),
                        new XElement("BomXmlPath", BomXmlPath),
                        new XElement("CheckRefExcludeListPath", CheckRefExcludeListPath),
                        // [NEW v1.2] Create DWG settings
                        new XElement("CreateDwgIniPath", CreateDwgIniPath),
                        new XElement("CreateDwgListPath", CreateDwgListPath),
                        new XElement("CreateDwgOutputPath", CreateDwgOutputPath)
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