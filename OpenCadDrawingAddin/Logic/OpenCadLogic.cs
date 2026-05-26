using System;
using System.IO;
using System.Windows.Forms;
using Inventor;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Logic chính: Tìm và mở file CAD (.dwg) tương ứng với file Inventor đang mở
    /// Port từ iLogic VB script → C# Add-in
    /// 
    /// Quy tắc:
    ///   - Part document   → lấy tên file Part
    ///   - Assembly doc    → phải chọn đúng 1 component, lấy tên file của component đó
    ///   - UseRevisionSuffix = true  → tên file CAD = {partName}-{revisionNumber}.dwg
    ///   - UseRevisionSuffix = false → tên file CAD = {partName}.dwg
    /// 
    /// Version 1.0
    /// </summary>
    public class OpenCadLogic
    {
        private readonly Inventor.Application _app;

        public OpenCadLogic(Inventor.Application app)
        {
            _app = app;
        }

        /// <summary>
        /// Điểm vào chính - gọi khi user click nút "Open CAD Drawing"
        /// </summary>
        public void RunOpenCad()
        {
            try
            {
                // 1. Load settings
                var settings = OpenCadSettings.Load();

                if (string.IsNullOrEmpty(settings.CadFolderPath) || !Directory.Exists(settings.CadFolderPath))
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_NO_FOLDER_CONFIGURED"),
                        LanguageManager.L("TITLE_ERROR"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 2. Lấy document hiện tại
                Document activeDoc = _app.ActiveDocument;
                if (activeDoc == null)
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_NO_ACTIVE_DOC"),
                        LanguageManager.L("TITLE_ERROR"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 3. Xác định partName và revision theo loại document
                string partName = "";
                string revision = "";

                if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    // === PART ===
                    partName = System.IO.Path.GetFileNameWithoutExtension(activeDoc.FullFileName);
                    revision = GetRevisionNumber(activeDoc);
                }
                else if (activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    // === ASSEMBLY ===
                    // Phải chọn đúng 1 component
                    if (activeDoc.SelectSet.Count != 1)
                    {
                        MessageBox.Show(
                            LanguageManager.L("MSG_SELECT_ONE_PART"),
                            LanguageManager.L("TITLE_WARNING"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    // Lấy component được chọn
                    var selectedObj = activeDoc.SelectSet[1]; // 1-based index trong Inventor
                    Document componentDoc = GetDocumentFromSelected(selectedObj);

                    if (componentDoc == null)
                    {
                        MessageBox.Show(
                            LanguageManager.L("MSG_CANNOT_GET_COMPONENT"),
                            LanguageManager.L("TITLE_ERROR"),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }

                    partName = System.IO.Path.GetFileNameWithoutExtension(componentDoc.FullFileName);
                    revision = GetRevisionNumber(componentDoc);
                }
                else
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_UNSUPPORTED_DOC_TYPE"),
                        LanguageManager.L("TITLE_WARNING"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // 4. Xây dựng tên file CAD
                string suffix = "";
                if (settings.UseRevisionSuffix)
                {
                    suffix = BuildRevisionSuffix(revision);
                }

                string extension = settings.CadExtension;
                if (!extension.StartsWith("."))
                    extension = "." + extension;

                string cadFileName = partName + suffix + extension;
                string cadFilePath = System.IO.Path.Combine(settings.CadFolderPath, cadFileName);

                // 5. Mở file
                if (System.IO.File.Exists(cadFilePath))
                {
                    System.Diagnostics.Process.Start(cadFilePath);
                }
                else
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_FILE_NOT_FOUND", cadFilePath),
                        LanguageManager.L("TITLE_ERROR"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in RunOpenCad", ex);
                MessageBox.Show(
                    LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message,
                    LanguageManager.L("TITLE_ERROR"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // PRIVATE HELPERS
        // ============================================================

        /// <summary>
        /// Lấy Revision Number từ Inventor Summary Information
        /// Trả về string rỗng nếu không có hoặc lỗi
        /// </summary>
        private string GetRevisionNumber(Document doc)
        {
            try
            {
                PropertySet propSet = doc.PropertySets["Inventor Summary Information"];
                object val = propSet["Revision Number"].Value;
                return val != null ? val.ToString().Trim() : "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Lấy Document từ object được chọn trong Assembly
        /// Hỗ trợ ComponentOccurrence
        /// </summary>
        private Document GetDocumentFromSelected(object selectedObj)
        {
            try
            {
                if (selectedObj is ComponentOccurrence occ)
                {
                    return occ.Definition.Document as Document;
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Xây dựng suffix từ revision number
        /// Chỉ chấp nhận revision là số tự nhiên dương chuẩn (1,2,3,...)
        /// Nếu revision không thỏa (ví dụ: "R01", "01", "A", "R1"...) thì không thêm suffix.
        /// </summary>
        private string BuildRevisionSuffix(string revision)
        {
            if (string.IsNullOrEmpty(revision)) return "";

            // Chỉ chấp nhận các số dương bắt đầu bằng 1-9 (không chấp nhận dẫn zero hoặc có chữ)
            // Ví dụ: "1", "2", "10" => chấp nhận; "0", "01", "R01", "A" => không chấp nhận
            try
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(revision.Trim(), "^[1-9][0-9]*$"))
                {
                    return "-" + revision.Trim();
                }
                return "";
            }
            catch
            {
                return "";
            }
        }
    }
}
