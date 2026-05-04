using System;
using System.Windows.Forms;
using Inventor;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Tính năng Cross-Screen Copy / Paste Component
    /// 
    /// Mục đích: Hỗ trợ làm việc 2 màn hình - copy component từ màn hình 1,
    ///           paste vào Assembly đang mở ở màn hình 2.
    ///
    /// Cách dùng:
    ///   1. Mở file IPT hoặc IAM (hoặc chọn component trong Assembly)
    ///      → Click nút [Copy Component]
    ///   2. Chuyển sang màn hình 2, mở Assembly đích
    ///      → Click nút [Paste Component]
    ///      → Component sẽ được insert vào Assembly tại vị trí camera hiện tại
    ///
    /// Version 1.0
    /// </summary>
    public class CrossScreenCopyPasteLogic
    {
        // Prefix đặc biệt để phân biệt với clipboard thông thường
        private const string ClipboardPrefix = "INVENTOR_COMPONENT_PATH::";

        private readonly Inventor.Application _app;

        public CrossScreenCopyPasteLogic(Inventor.Application app)
        {
            _app = app;
        }

        // ============================================================
        // COPY
        // ============================================================

        /// <summary>
        /// Copy: Lấy đường dẫn file của component đang active hoặc được chọn
        ///       → Lưu vào Clipboard
        ///
        /// Ưu tiên:
        ///   1. Nếu document hiện tại là Part hoặc Assembly → dùng luôn
        ///   2. Nếu đang trong Assembly và có chọn 1 component → lấy file của component đó
        /// </summary>
        public void CopyComponent()
        {
            try
            {
                Document activeDoc = _app.ActiveDocument;
                if (activeDoc == null)
                {
                    ShowError(LanguageManager.L("MSG_NO_ACTIVE_DOC"));
                    return;
                }

                string filePath = "";

                if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject ||
                    activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    // Trường hợp 1: File IPT hoặc IAM đang active trực tiếp
                    filePath = activeDoc.FullFileName;
                }
                else
                {
                    ShowError(LanguageManager.L("MSG_UNSUPPORTED_DOC_TYPE"));
                    return;
                }

                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                {
                    ShowError("Không tìm thấy file: " + filePath);
                    return;
                }

                // Lưu vào Clipboard với prefix đặc biệt
                Clipboard.SetText(ClipboardPrefix + filePath);

                // Thông báo ngắn gọn
                string shortName = System.IO.Path.GetFileName(filePath);
                MessageBox.Show(
                    $"Đã copy:\n{shortName}\n\nSang màn hình 2 → nhấn [Paste Component] để insert vào Assembly.",
                    "Copy Component",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in CopyComponent", ex);
                ShowError("Lỗi Copy: " + ex.Message);
            }
        }

        // ============================================================
        // PASTE
        // ============================================================

        /// <summary>
        /// Paste: Đọc đường dẫn từ Clipboard → Insert component vào Assembly hiện tại
        ///
        /// Vị trí: Đặt tại vị trí camera target (điểm nhìn hiện tại) để gần
        ///         vùng làm việc của user. User có thể kéo để chỉnh sau.
        /// </summary>
        public void PasteComponent()
        {
            try
            {
                // 1. Đọc clipboard
                string clipText = Clipboard.GetText();
                if (string.IsNullOrEmpty(clipText) || !clipText.StartsWith(ClipboardPrefix))
                {
                    MessageBox.Show(
                        "Clipboard không chứa component Inventor.\nHãy [Copy Component] trước.",
                        "Paste Component",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                string filePath = clipText.Substring(ClipboardPrefix.Length);

                if (!System.IO.File.Exists(filePath))
                {
                    ShowError($"Không tìm thấy file:\n{filePath}\n\nKiểm tra lại đường dẫn hoặc kết nối mạng.");
                    return;
                }

                // 2. Kiểm tra document đích phải là Assembly
                Document activeDoc = _app.ActiveDocument;
                if (activeDoc == null || activeDoc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    ShowError("Vui lòng mở một file Assembly (.iam) trước khi Paste.");
                    return;
                }

                AssemblyDocument asmDoc = activeDoc as AssemblyDocument;
                AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

                // 3. Xây dựng Matrix vị trí đặt component
                Matrix placementMatrix = BuildPlacementMatrix();

                // 4. Insert component
                ComponentOccurrence newOcc = asmDef.Occurrences.Add(filePath, placementMatrix);

                // 5. Thông báo thành công
                string shortName = System.IO.Path.GetFileName(filePath);
                MessageBox.Show(
                    $"Đã paste:\n{shortName}\n\nComponent được đặt tại vùng camera hiện tại.\nBạn có thể kéo để chỉnh vị trí.",
                    "Paste Component",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in PasteComponent", ex);
                ShowError("Lỗi Paste: " + ex.Message);
            }
        }

        // ============================================================
        // PRIVATE HELPERS
        // ============================================================

        /// <summary>
        /// Xây dựng Matrix vị trí đặt component khi paste.
        ///
        /// Chiến lược vị trí:
        ///   - Thử lấy Camera.Target của View hiện tại → đặt component tại đó
        ///   - Nếu lỗi → đặt tại origin (0, 0, 0)
        ///
        /// Lưu ý: Inventor dùng đơn vị cm nội bộ.
        /// </summary>
        private Matrix BuildPlacementMatrix()
        {
            TransientGeometry tg = _app.TransientGeometry;
            Matrix mat = tg.CreateMatrix(); // Identity matrix (origin)

            try
            {
                // Lấy vị trí Camera.Target của cửa sổ active
                // → đây là điểm trung tâm màn hình user đang nhìn
                Camera cam = _app.ActiveView.Camera;
                Point target = cam.Target;

                // Dịch matrix đến vị trí camera target
                mat.SetTranslation(tg.CreateVector(target.X, target.Y, target.Z), false);
            }
            catch
            {
                // Fallback: đặt tại origin nếu không lấy được camera
                // mat đã là identity, không cần làm gì thêm
            }

            return mat;
        }

        private void ShowError(string message)
        {
            MessageBox.Show(
                message,
                "Lỗi",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}
