using System;
using System.Drawing;
using System.Windows.Forms;
using Inventor;
using DrawingPoint = System.Drawing.Point;
using InventorPoint = Inventor.Point;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Tính năng Cross-Screen Copy / Paste Component
    /// Version 1.2 - Thêm "Don't show again" + Settings integration
    /// </summary>
    public class CrossScreenCopyPasteLogic
    {
        private const string ClipboardPrefix = "INVENTOR_COMPONENT_PATH::";
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
                if (activeDoc == null) { ShowError(LanguageManager.L("MSG_NO_ACTIVE_DOC")); return; }

                var filePaths = new System.Collections.Generic.List<string>();

                if (activeDoc.SelectSet != null && activeDoc.SelectSet.Count >= 1)
                {
                    // Lấy tất cả các component được chọn
                    for (int i = 1; i <= activeDoc.SelectSet.Count; i++)
                    {
                        Document componentDoc = GetDocumentFromSelected(activeDoc.SelectSet[i]);
                        if (componentDoc != null)
                        {
                            string fp = componentDoc.FullFileName;
                            if (!filePaths.Contains(fp)) filePaths.Add(fp);
                        }
                    }

                    // Fallback: không parse được → dùng activeDoc
                    if (filePaths.Count == 0)
                    {
                        if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject ||
                            activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                            filePaths.Add(activeDoc.FullFileName);
                        else { ShowError(LanguageManager.L("MSG_COPY_COMP_CANNOT_GET")); return; }
                    }
                }
                else if (activeDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject ||
                         activeDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    filePaths.Add(activeDoc.FullFileName);
                }
                else
                {
                    MessageBox.Show(LanguageManager.L("MSG_COPY_COMP_NO_SELECTION"),
                        LanguageManager.L("TITLE_COPY_COMP"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Kiểm tra file tồn tại
                filePaths.RemoveAll(fp => !System.IO.File.Exists(fp));
                if (filePaths.Count == 0)
                {
                    ShowError(LanguageManager.L("MSG_COPY_COMP_FILE_NOT_FOUND", "")); return;
                }

                // Lưu vào clipboard - format: PREFIX + path1|path2 + "##" + mat1|mat2
                // Matrix serialize: 16 số double cách nhau bằng ","
                var matrices = new System.Collections.Generic.List<string>();
                if (activeDoc.SelectSet != null && activeDoc.SelectSet.Count >= 1)
                {
                    for (int i = 1; i <= activeDoc.SelectSet.Count; i++)
                    {
                        try
                        {
                            var occ = GetOccurrenceFromObject(activeDoc.SelectSet[i]);
                            if (occ != null)
                            {
                                double[] cells = new double[16];
                                occ.Transformation.GetMatrixData(ref cells);
                                matrices.Add(string.Join(",", System.Array.ConvertAll(cells, c => c.ToString("R"))));
                            }
                            else matrices.Add("");
                        }
                        catch { matrices.Add(""); }
                    }
                }
                // Pad matrices nếu ít hơn filePaths
                while (matrices.Count < filePaths.Count) matrices.Add("");

                string clipData = ClipboardPrefix
                    + string.Join("|", filePaths)
                    + "##"
                    + string.Join("|", matrices);
                Clipboard.SetText(clipData);

                var settings = OpenCadSettings.Load();
                if (settings.ShowCopyNotification)
                {
                    string names = string.Join("\n", filePaths.ConvertAll(System.IO.Path.GetFileName));
                    bool dontShowAgain = ShowNotifyWithCheckbox(
                        LanguageManager.L("MSG_COPY_COMP_SUCCESS", names),
                        LanguageManager.L("TITLE_COPY_COMP"));
                    if (dontShowAgain) { settings.ShowCopyNotification = false; settings.Save(); }
                }
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in CopyComponent", ex);
                ShowError(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message);
            }
        }

        // ============================================================
        // PASTE
        // ============================================================

        public void PasteComponent(ComponentOccurrence preSelected = null)
        {
            try
            {
                string clipText = Clipboard.GetText();
                if (string.IsNullOrEmpty(clipText) || !clipText.StartsWith(ClipboardPrefix))
                {
                    MessageBox.Show(LanguageManager.L("MSG_PASTE_COMP_NO_CLIPBOARD"),
                        LanguageManager.L("TITLE_PASTE_COMP"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Tách paths và matrices từ clipboard
                string payload = clipText.Substring(ClipboardPrefix.Length);
                string[] parts = payload.Split(new[] { "##" }, StringSplitOptions.None);
                string[] filePaths = parts[0].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                string[] matrixStrs = parts.Length > 1
                    ? parts[1].Split(new[] { '|' }, StringSplitOptions.None)
                    : new string[0];

                // Kiểm tra tất cả file tồn tại
                foreach (var fp in filePaths)
                {
                    if (!System.IO.File.Exists(fp))
                    {
                        ShowError(LanguageManager.L("MSG_PASTE_COMP_FILE_NOT_FOUND", fp)); return;
                    }
                }

                Document activeDoc = _app.ActiveDocument;
                if (activeDoc == null || activeDoc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    ShowError(LanguageManager.L("MSG_PASTE_COMP_ASM_ONLY")); return;
                }

                var asmDoc = activeDoc as AssemblyDocument;
                string notifyMsg;

                if (preSelected != null && filePaths.Length == 1)
                {
                    // Replace chỉ hoạt động khi chọn 1 component và 1 file
                    Matrix savedMatrix = preSelected.Transformation;
                    preSelected.Delete();
                    asmDoc.ComponentDefinition.Occurrences.Add(filePaths[0], savedMatrix);
                    notifyMsg = LanguageManager.L("MSG_PLACE_COMP_REPLACED",
                        System.IO.Path.GetFileName(filePaths[0]));
                }
                else
                {
                    TransientGeometry tg = _app.TransientGeometry;
                    for (int i = 0; i < filePaths.Length; i++)
                    {
                        Matrix mat = BuildPlacementMatrix();
                        // Dùng matrix gốc nếu có → giữ vị trí tương đối giữa các part
                        if (i < matrixStrs.Length && !string.IsNullOrEmpty(matrixStrs[i]))
                        {
                            try
                            {
                                double[] cells = System.Array.ConvertAll(
                                    matrixStrs[i].Split(','), double.Parse);
                                if (cells.Length == 16)
                                {
                                    mat = tg.CreateMatrix();
                                    mat.PutMatrixData(ref cells);
                                }
                            }
                            catch { }
                        }
                        asmDoc.ComponentDefinition.Occurrences.Add(filePaths[i], mat);
                    }

                    string names = string.Join(", ", System.Array.ConvertAll(filePaths, System.IO.Path.GetFileName));
                    notifyMsg = LanguageManager.L("MSG_PASTE_COMP_SUCCESS", names);
                }

                var settings = OpenCadSettings.Load();
                if (settings.ShowPasteNotification)
                {
                    bool dontShowAgain = ShowNotifyWithCheckbox(notifyMsg, LanguageManager.L("TITLE_PASTE_COMP"));
                    if (dontShowAgain) { settings.ShowPasteNotification = false; settings.Save(); }
                }
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in PasteComponent", ex);
                ShowError(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message);
            }
        }

        // ============================================================
        // PRIVATE HELPERS
        // ============================================================

        /// <summary>
        /// Hiện dialog thông báo có checkbox "Không hiển thị lại".
        /// Trả về true nếu user tick vào checkbox.
        /// </summary>
        private bool ShowNotifyWithCheckbox(string message, string title)
        {
            bool dontShowAgain = false;

            using (var form = new Form())
            {
                form.Text = title;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowInTaskbar = false;
                form.Width = 380;

                int pad = 16, y = pad;

                // Icon + message
                var iconBox = new PictureBox
                {
                    Image = SystemIcons.Information.ToBitmap(),
                    SizeMode = PictureBoxSizeMode.StretchImage,
                    Size = new Size(32, 32),
                    Location = new DrawingPoint(pad, y)
                };
                form.Controls.Add(iconBox);

                var lbl = new Label
                {
                    Text = message,
                    AutoSize = false,
                    Width = form.ClientSize.Width - pad * 2 - 40,
                    Location = new DrawingPoint(pad + 40, y),
                    Font = new Font("Segoe UI", 9f)
                };
                using (var g = lbl.CreateGraphics())
                {
                    var sz = g.MeasureString(message, lbl.Font, new SizeF(lbl.Width, float.MaxValue));
                    lbl.Height = (int)sz.Height + 4;
                }
                form.Controls.Add(lbl);
                y += Math.Max(iconBox.Height, lbl.Height) + pad;

                // Separator
                var sep = new Label
                {
                    BorderStyle = BorderStyle.Fixed3D,
                    Height = 2,
                    Width = form.ClientSize.Width - pad * 2,
                    Location = new DrawingPoint(pad, y)
                };
                form.Controls.Add(sep);
                y += sep.Height + 8;

                // Checkbox "Không hiển thị lại"
                var chk = new CheckBox
                {
                    Text = LanguageManager.L("CHK_DONT_SHOW_AGAIN"),
                    AutoSize = true,
                    Location = new DrawingPoint(pad, y),
                    Font = new Font("Segoe UI", 9f)
                };
                form.Controls.Add(chk);
                y += chk.PreferredSize.Height + pad;

                // Button OK
                var btn = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Width = 80,
                    Height = 28,
                    Font = new Font("Segoe UI", 9f)
                };
                btn.Location = new DrawingPoint(form.ClientSize.Width - btn.Width - pad, y);
                form.Controls.Add(btn);
                form.AcceptButton = btn;
                y += btn.Height + pad;

                form.ClientSize = new Size(form.ClientSize.Width, y);
                form.ShowDialog();
                dontShowAgain = chk.Checked;
            }

            return dontShowAgain;
        }

        private Document GetDocumentFromSelected(object selectedObj)
        {
            try
            {
                if (selectedObj is ComponentOccurrence occ)
                    return occ.Definition.Document as Document;
                if (selectedObj is FaceProxy fp)
                    return fp.ContainingOccurrence?.Definition?.Document as Document;
                if (selectedObj is EdgeProxy ep)
                    return ep.ContainingOccurrence?.Definition?.Document as Document;
                if (selectedObj is VertexProxy vp)
                    return vp.ContainingOccurrence?.Definition?.Document as Document;
                return null;
            }
            catch { return null; }
        }

        /// <summary>
        /// Lấy ComponentOccurrence từ bất kỳ object nào trong SelectSet.
        /// Public static để StandardAddInServer có thể gọi ngay trong OnExecute
        /// trước khi Inventor kịp clear SelectSet.
        /// </summary>
        public static ComponentOccurrence GetOccurrenceFromObject(object selectedObj)
        {
            try
            {
                if (selectedObj is ComponentOccurrence occ) return occ;
                if (selectedObj is FaceProxy fp) return fp.ContainingOccurrence;
                if (selectedObj is EdgeProxy ep) return ep.ContainingOccurrence;
                if (selectedObj is VertexProxy vp) return vp.ContainingOccurrence;
                return null;
            }
            catch { return null; }
        }

        private Matrix BuildPlacementMatrix()
        {
            TransientGeometry tg = _app.TransientGeometry;
            Matrix mat = tg.CreateMatrix();
            try
            {
                InventorPoint target = _app.ActiveView.Camera.Target;
                mat.SetTranslation(tg.CreateVector(target.X, target.Y, target.Z), false);
            }
            catch { }
            return mat;
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, LanguageManager.L("TITLE_ERROR"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}