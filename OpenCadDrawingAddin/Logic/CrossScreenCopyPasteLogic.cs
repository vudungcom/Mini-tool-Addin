using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Inventor;

namespace OpenCadDrawingAddin.Logic
{
    public class CrossScreenCopyPasteLogic
    {
        private const string ClipboardPrefix = "INVENTOR_COMPONENT_PATH::";
        private const string PathSeparator = "||";

        private readonly Inventor.Application _app;

        public CrossScreenCopyPasteLogic(Inventor.Application app)
        {
            _app = app;
        }

        // ============================================================
        // COPY
        // ============================================================

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

                var filePaths = new List<string>();

                // Ưu tiên: đọc selection hiện tại (user Ctrl+click nhiều component trong Assembly)
                if (activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    try
                    {
                        SelectSet sel = _app.ActiveView.SelectSet;
                        foreach (object item in sel)
                        {
                            if (item is ComponentOccurrence occ)
                            {
                                string path = GetPathFromOccurrence(occ);
                                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path) && !filePaths.Contains(path))
                                    filePaths.Add(path);
                            }
                        }
                    }
                    catch { }
                }

                // Fallback: dùng active document nếu không có selection
                if (filePaths.Count == 0)
                {
                    if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject ||
                        activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        string docPath = activeDoc.FullFileName;
                        if (!string.IsNullOrEmpty(docPath))
                            filePaths.Add(docPath);
                    }
                    else
                    {
                        ShowError(LanguageManager.L("MSG_UNSUPPORTED_DOC_TYPE"));
                        return;
                    }
                }

                // Validate
                foreach (string fp in filePaths)
                {
                    if (!System.IO.File.Exists(fp))
                    {
                        ShowError("Không tìm thấy file: " + fp);
                        return;
                    }
                }

                // Lưu vào Clipboard
                Clipboard.SetText(ClipboardPrefix + string.Join(PathSeparator, filePaths));

                // Thông báo
                if (filePaths.Count == 1)
                {
                    string shortName = System.IO.Path.GetFileName(filePaths[0]);
                    MessageBox.Show(
                        $"Đã copy:\n{shortName}\n\nSang màn hình 2 → nhấn [Paste Component] để insert vào Assembly.",
                        "Copy Component", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string names = string.Join("\n", filePaths.Select(f => "• " + System.IO.Path.GetFileName(f)));
                    MessageBox.Show(
                        $"Đã copy {filePaths.Count} component:\n{names}\n\nSang màn hình 2 → nhấn [Paste Component] để insert tất cả vào Assembly.",
                        "Copy Component", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
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
                        "Paste Component", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string[] filePaths = clipText
                    .Substring(ClipboardPrefix.Length)
                    .Split(new[] { PathSeparator }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string fp in filePaths)
                {
                    if (!System.IO.File.Exists(fp))
                    {
                        ShowError($"Không tìm thấy file:\n{fp}\n\nKiểm tra lại đường dẫn hoặc kết nối mạng.");
                        return;
                    }
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
                Matrix placementMatrix = BuildPlacementMatrix();

                // 3. Insert từng component
                var insertedNames = new List<string>();
                foreach (string filePath in filePaths)
                {
                    asmDef.Occurrences.Add(filePath, placementMatrix);
                    insertedNames.Add(System.IO.Path.GetFileName(filePath));
                }

                // 4. Thông báo
                if (insertedNames.Count == 1)
                {
                    MessageBox.Show(
                        $"Đã paste:\n{insertedNames[0]}\n\nComponent được đặt tại vùng camera hiện tại.\nBạn có thể kéo để chỉnh vị trí.",
                        "Paste Component", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string names = string.Join("\n", insertedNames.Select(n => "• " + n));
                    MessageBox.Show(
                        $"Đã paste {insertedNames.Count} component:\n{names}\n\nCác component được đặt tại vùng camera hiện tại.",
                        "Paste Component", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
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

        private string GetPathFromOccurrence(ComponentOccurrence occ)
        {
            // Thử ReferencedDocumentDescriptor trước (nhanh nhất)
            try
            {
                string path = occ.ReferencedDocumentDescriptor.FullDocumentName;
                if (!string.IsNullOrEmpty(path))
                    return path;
            }
            catch { }

            // Fallback qua Definition
            try
            {
                if (occ.Definition is PartComponentDefinition partDef)
                    return partDef.Document.FullFileName;
                if (occ.Definition is AssemblyComponentDefinition asmDef)
                    return asmDef.Document.FullFileName;
            }
            catch { }

            return null;
        }

        private Matrix BuildPlacementMatrix()
        {
            TransientGeometry tg = _app.TransientGeometry;
            Matrix mat = tg.CreateMatrix();

            try
            {
                Camera cam = _app.ActiveView.Camera;
                Point target = cam.Target;
                mat.SetTranslation(tg.CreateVector(target.X, target.Y, target.Z), false);
            }
            catch { }

            return mat;
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
