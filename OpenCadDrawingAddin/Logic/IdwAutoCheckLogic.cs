using System;
using System.Windows.Forms;
using Inventor;
using IOPath = System.IO.Path;
using System.Collections.Generic;

namespace OpenCadDrawingAddin.Logic
{
    public class IdwAutoCheckLogic
    {
        private readonly Inventor.Application _app;

        public IdwAutoCheckLogic(Inventor.Application app)
        {
            _app = app;
        }

        public void RunChecks(Document doc)
        {
            try
            {
                if (doc == null || doc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject) return;

                // Load settings từ project của bạn
                var settings = OpenCadSettings.Load();
                if (!settings.AutoCheckDrawingName && !settings.AutoCheckAppearance) return;

                var drw = doc as DrawingDocument;
                List<string> issues = new List<string>();

                // 1. Kiểm tra Tên bản vẽ (Part_No)
                if (settings.AutoCheckDrawingName)
                {
                    string r = CheckDrawingName(drw);
                    if (!string.IsNullOrEmpty(r)) issues.Add(r.TrimEnd());
                }

                // 2. Kiểm tra Vật liệu & Bề mặt (Bắt cả lỗi PAINT vs Tiếng Hàn)
                if (settings.AutoCheckAppearance)
                {
                    string r = CheckMaterialAndAppearance(drw);
                    if (!string.IsNullOrEmpty(r))
                    {
                        foreach (string part in r.Split(new[] { "\n\n" }, StringSplitOptions.RemoveEmptyEntries))
                            issues.Add(part.TrimEnd());
                    }
                }

                if (issues.Count == 0) return;

                // Hiển thị thông báo
                string msg = "";
                for (int i = 0; i < issues.Count; i++)
                    msg += string.Format("[{0}] {1}\n\n", i + 1, issues[i]);

                msg += LanguageManager.L("MSG_IDW_CHECK_DISABLE_HINT");
                MessageBox.Show(msg, LanguageManager.L("TITLE_IDW_CHECK"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in IdwAutoCheckLogic.RunChecks", ex);
            }
        }

        private string CheckDrawingName(DrawingDocument drw)
        {
            try
            {
                string modelName = GetPrimaryModelName(drw);
                if (string.IsNullOrEmpty(modelName)) return "";

                string result = "";
                foreach (Sheet sheet in drw.Sheets)
                {
                    if (sheet.TitleBlock == null) continue;
                    string currentVal = ReadValueByPosition(sheet, "Part_No", true);

                    if (currentVal == null) continue;
                    string cleanVal = IOPath.GetFileNameWithoutExtension(currentVal).Trim();

                    if (string.Compare(cleanVal, modelName, StringComparison.OrdinalIgnoreCase) != 0)
                        result += LanguageManager.L("MSG_IDW_NAME_MISMATCH", currentVal, modelName, sheet.Name) + "\n";
                }
                return result;
            }
            catch { return ""; }
        }

        private string CheckMaterialAndAppearance(DrawingDocument drw)
        {
            try
            {
                string result = "";
                string matTb = null, appTb = null;

                if (drw.ActiveSheet.TitleBlock != null)
                {
                    var sheet = drw.ActiveSheet;
                    matTb = ReadValueByPosition(sheet, "Material", false);
                    appTb = ReadValueByPosition(sheet, "Appearance", false)
                          ?? ReadValueByPosition(sheet, "Treatment", false)
                          ?? ReadValueByPosition(sheet, "Finish", false);
                }

                foreach (Document refDoc in drw.ReferencedDocuments)
                {
                    if (!(refDoc is PartDocument pd)) continue;
                    var def = pd.ComponentDefinition;

                    string matModel = def.Material.Name ?? "";
                    string appModel = pd.ActiveAppearance != null ? pd.ActiveAppearance.DisplayName : "";

                    string matProp = "";
                    string appProp = "";
                    try
                    {
                        PropertySet designProp = pd.PropertySets["Design Tracking Properties"];
                        matProp = designProp["Material"].Value != null ? designProp["Material"].Value.ToString().Trim() : "";
                        appProp = designProp["Appearance"].Value != null ? designProp["Appearance"].Value.ToString().Trim() : "";
                    }
                    catch { }

                    // Check Material
                    bool matIntMismatch = !string.IsNullOrEmpty(matProp) && string.Compare(matModel, matProp, StringComparison.OrdinalIgnoreCase) != 0;
                    bool matTbMismatch = matTb != null && string.Compare(matModel, matTb.Trim(), StringComparison.OrdinalIgnoreCase) != 0;
                    if (matIntMismatch || matTbMismatch)
                    {
                        string detail = "Model=\"" + matModel + "\"";
                        if (matIntMismatch) detail += " | iProp=\"" + matProp + "\"";
                        if (matTbMismatch) detail += " | TB=\"" + matTb + "\"";
                        result += LanguageManager.L("MSG_IDW_MATERIAL_MISMATCH", pd.DisplayName, matModel, detail) + "\n\n";
                    }

                    // Check Appearance (Bắt lỗi PAINT vs Tiếng Hàn)
                    bool appIntMismatch = !string.IsNullOrEmpty(appProp) && string.Compare(appModel, appProp, StringComparison.OrdinalIgnoreCase) != 0;
                    bool appTbMismatch = appTb != null && string.Compare(appModel, appTb.Trim(), StringComparison.OrdinalIgnoreCase) != 0;
                    if (appIntMismatch || appTbMismatch)
                    {
                        string detail = "Model=\"" + appModel + "\"";
                        if (appIntMismatch) detail += " | iProp=\"" + appProp + "\"";
                        if (appTbMismatch) detail += " | TB=\"" + appTb + "\"";
                        result += LanguageManager.L("MSG_IDW_APPEARANCE_MISMATCH", pd.DisplayName, appModel, detail) + "\n\n";
                    }
                }
                return result;
            }
            catch { return ""; }
        }

        private string ReadValueByPosition(Sheet sheet, string labelText, bool searchRight)
        {
            try
            {
                var tbDef = sheet.TitleBlock.Definition;
                Inventor.TextBox labelBox = null;

                foreach (Inventor.TextBox tb in tbDef.Sketch.TextBoxes)
                {
                    if (string.Compare(tb.Text.Trim(), labelText, StringComparison.OrdinalIgnoreCase) == 0)
                    { labelBox = tb; break; }
                }

                if (labelBox == null) return null;

                Inventor.TextBox targetBox = null;
                double bestDist = 9999;

                foreach (Inventor.TextBox tb in tbDef.Sketch.TextBoxes)
                {
                    if (tb == labelBox) continue;
                    double xDiff = tb.Origin.X - labelBox.Origin.X;
                    double yDiff = labelBox.Origin.Y - tb.Origin.Y;

                    if (searchRight)
                    {
                        if (Math.Abs(yDiff) < 0.1 && xDiff > 0 && xDiff < 15 && xDiff < bestDist)
                        { bestDist = xDiff; targetBox = tb; }
                    }
                    else
                    {
                        if (yDiff > 0.4 && yDiff < 1.5 && Math.Abs(xDiff) < 0.1 && yDiff < bestDist)
                        { bestDist = yDiff; targetBox = tb; }
                    }
                }
                return targetBox != null ? sheet.TitleBlock.GetResultText(targetBox) : null;
            }
            catch { return null; }
        }

        private string GetPrimaryModelName(DrawingDocument drw)
        {
            foreach (Document refDoc in drw.ReferencedDocuments)
            {
                if (refDoc is PartDocument || refDoc is AssemblyDocument)
                    return IOPath.GetFileNameWithoutExtension(refDoc.FullFileName);
            }
            return "";
        }
    }
}