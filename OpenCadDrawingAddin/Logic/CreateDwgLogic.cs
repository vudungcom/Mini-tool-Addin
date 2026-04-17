using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Inventor;
// [FIX CS0104] Alias để tránh nhầm với Inventor.File / Inventor.Path
using File = System.IO.File;
using Path = System.IO.Path;
using Directory = System.IO.Directory;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Logic: Export IDW → DWG
    ///
    /// Hai chế độ:
    ///   1. RunCreateDwg()        — Assembly: hỏi Yes/No, xuất theo List trong Settings
    ///   2. RunCreateDwgFromIdw() — Drawing:  xuất thẳng file IDW hiện tại → DWG
    ///
    /// Version 1.1
    /// </summary>
    public class CreateDwgLogic
    {
        private readonly Inventor.Application _app;

        public CreateDwgLogic(Inventor.Application app)
        {
            _app = app;
        }

        // ============================================================
        // CHẾ ĐỘ 1: Assembly — xuất theo List
        // ============================================================

        /// <summary>
        /// Gọi khi click "Create DWG" trong môi trường Assembly
        /// </summary>
        public void RunCreateDwg()
        {
            try
            {
                var settings = OpenCadSettings.Load();

                if (string.IsNullOrEmpty(settings.CreateDwgIniPath) || !File.Exists(settings.CreateDwgIniPath))
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NO_INI"),
                        LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (string.IsNullOrEmpty(settings.CreateDwgOutputPath))
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NO_OUTPUT"),
                        LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Kiểm tra Assembly
                Document activeDoc = _app.ActiveDocument;
                if (activeDoc == null || activeDoc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject)
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NEED_ASSEMBLY"),
                        LanguageManager.L("TITLE_WARNING"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Hỏi xác nhận xuất theo danh sách
                if (MessageBox.Show(
                    LanguageManager.L("MSG_CREATE_DWG_ASK_LIST"),
                    LanguageManager.L("TITLE_CREATE_DWG_CONFIRM"),
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return;

                // Validate list file
                if (string.IsNullOrEmpty(settings.CreateDwgListPath) || !File.Exists(settings.CreateDwgListPath))
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NO_LIST"),
                        LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                AssemblyDocument asmDoc = (AssemblyDocument)activeDoc;

                // Tạo output folder nếu chưa có
                if (!Directory.Exists(settings.CreateDwgOutputPath))
                {
                    try { Directory.CreateDirectory(settings.CreateDwgOutputPath); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_CANNOT_CREATE_FOLDER") + ex.Message,
                            LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Quét model trong assembly
                var modelMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var duplicateMap = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                CollectAssemblyModelFiles(asmDoc, modelMap, duplicateMap);

                if (modelMap.Count == 0)
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NO_MODELS"),
                        LanguageManager.L("TITLE_WARNING"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Đọc danh sách cần xuất
                HashSet<string> wantedNames = ReadWantedBaseNamesFromTextFile(settings.CreateDwgListPath);
                if (wantedNames.Count == 0)
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_LIST_EMPTY"),
                        LanguageManager.L("TITLE_WARNING"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Khớp danh sách với IDW
                var selectedIdwFiles = new List<string>();
                var missingNameList = new List<string>();
                var noIdwNameList = new List<string>();
                var selectedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (string wantedName in wantedNames)
                {
                    if (modelMap.ContainsKey(wantedName))
                    {
                        string idwPath = FindIdwInSameFolderOnly(modelMap[wantedName]);
                        if (!string.IsNullOrEmpty(idwPath))
                            selectedSet.Add(idwPath);
                        else
                            noIdwNameList.Add(wantedName);
                    }
                    else
                    {
                        missingNameList.Add(wantedName);
                    }
                }
                selectedIdwFiles.AddRange(selectedSet);

                if (selectedIdwFiles.Count == 0)
                {
                    MessageBox.Show(
                        LanguageManager.L("MSG_CREATE_DWG_NO_MATCH") + "\n" +
                        LanguageManager.L("MSG_CREATE_DWG_MISSING_IN_ASM") + missingNameList.Count + "\n" +
                        LanguageManager.L("MSG_CREATE_DWG_NO_IDW_FOR_MODEL") + noIdwNameList.Count,
                        LanguageManager.L("TITLE_WARNING"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Preview + xác nhận
                var preview = new StringBuilder();
                preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_PREVIEW_COUNT") + selectedIdwFiles.Count);
                preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_PREVIEW_OUTPUT") + settings.CreateDwgOutputPath);
                preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_MISSING_IN_ASM") + missingNameList.Count);
                preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_NO_IDW_FOR_MODEL") + noIdwNameList.Count);
                if (duplicateMap.Count > 0)
                    preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_DUPLICATE_WARNING") + duplicateMap.Count);
                preview.AppendLine();
                preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_PREVIEW_FIRST5"));
                int showCount = Math.Min(5, selectedIdwFiles.Count);
                for (int i = 0; i < showCount; i++)
                    preview.AppendLine(" - " + Path.GetFileName(selectedIdwFiles[i]));
                preview.AppendLine();
                preview.AppendLine(LanguageManager.L("MSG_CREATE_DWG_CONFIRM_PROMPT"));

                if (MessageBox.Show(preview.ToString(), LanguageManager.L("TITLE_CREATE_DWG_CONFIRM"),
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK)
                    return;

                // Export
                int okCount = 0, failCount = 0;
                var log = new StringBuilder();
                log.AppendLine("===== EXPORT IDW TO DWG =====");
                log.AppendLine("Assembly: " + asmDoc.FullFileName);
                log.AppendLine("INI: " + settings.CreateDwgIniPath);
                log.AppendLine("Output: " + settings.CreateDwgOutputPath);
                log.AppendLine("List: " + settings.CreateDwgListPath);
                log.AppendLine();

                if (duplicateMap.Count > 0)
                {
                    log.AppendLine("===== CANH BAO TRUNG TEN =====");
                    foreach (var kvp in duplicateMap)
                    {
                        log.AppendLine("[DUPLICATE] " + kvp.Key);
                        foreach (string dp in kvp.Value) log.AppendLine("    " + dp);
                        if (modelMap.ContainsKey(kvp.Key)) log.AppendLine("    [USING] " + modelMap[kvp.Key]);
                    }
                    log.AppendLine();
                }
                if (missingNameList.Count > 0)
                {
                    log.AppendLine("===== KHONG TIM THAY TRONG ASSEMBLY =====");
                    foreach (string n in missingNameList) log.AppendLine(n);
                    log.AppendLine();
                }
                if (noIdwNameList.Count > 0)
                {
                    log.AppendLine("===== CO MODEL NHUNG KHONG CO IDW =====");
                    foreach (string n in noIdwNameList) log.AppendLine(n);
                    log.AppendLine();
                }

                foreach (string idwFile in selectedIdwFiles)
                {
                    try
                    {
                        string dwgPath = Path.Combine(settings.CreateDwgOutputPath,
                            Path.GetFileNameWithoutExtension(idwFile) + ".dwg");
                        ExportIdwToDwg(idwFile, dwgPath, settings.CreateDwgIniPath);
                        log.AppendLine("[OK]  " + idwFile + "\n      -> " + dwgPath);
                        okCount++;
                    }
                    catch (Exception ex)
                    {
                        log.AppendLine("[ERR] " + idwFile + "\n      " + ex.Message);
                        failCount++;
                    }
                }

                log.AppendLine("\nTong: " + selectedIdwFiles.Count + " | OK: " + okCount + " | Fail: " + failCount);
                string logPath = Path.Combine(settings.CreateDwgOutputPath, "Export_IDW_to_DWG_Log.txt");
                File.WriteAllText(logPath, log.ToString(), Encoding.UTF8);

                var finalMsg = new StringBuilder();
                finalMsg.AppendLine(LanguageManager.L("MSG_CREATE_DWG_DONE"));
                finalMsg.AppendLine(LanguageManager.L("MSG_CREATE_DWG_TOTAL") + selectedIdwFiles.Count);
                finalMsg.AppendLine(LanguageManager.L("MSG_CREATE_DWG_OK") + okCount);
                finalMsg.AppendLine(LanguageManager.L("MSG_CREATE_DWG_FAIL") + failCount);
                finalMsg.AppendLine("\nLog: " + logPath);

                // Hiển thị danh sách model trùng tên trực tiếp để user đọc luôn
                if (duplicateMap.Count > 0)
                {
                    finalMsg.AppendLine();
                    finalMsg.AppendLine("⚠ " + LanguageManager.L("MSG_CREATE_DWG_DUPLICATE_WARNING") + duplicateMap.Count);
                    foreach (var kvp in duplicateMap)
                    {
                        finalMsg.AppendLine("  • " + kvp.Key);
                        foreach (string dp in kvp.Value)
                            finalMsg.AppendLine("      - " + dp);
                        if (modelMap.ContainsKey(kvp.Key))
                            finalMsg.AppendLine("      → " + LanguageManager.L("MSG_CREATE_DWG_DUPLICATE_USED") + modelMap[kvp.Key]);
                    }
                }

                MessageBox.Show(finalMsg.ToString(),
                    LanguageManager.L("TITLE_CREATE_DWG_CONFIRM"), MessageBoxButtons.OK,
                    duplicateMap.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in RunCreateDwg", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message,
                    LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // CHẾ ĐỘ 2: Drawing (IDW) — xuất thẳng file hiện tại
        // ============================================================

        /// <summary>
        /// Gọi khi click "Create DWG" trong môi trường Drawing (IDW)
        /// Xuất file IDW đang mở → DWG vào output folder đã cấu hình trong Settings
        /// </summary>
        public void RunCreateDwgFromIdw()
        {
            try
            {
                var settings = OpenCadSettings.Load();

                if (string.IsNullOrEmpty(settings.CreateDwgIniPath) || !File.Exists(settings.CreateDwgIniPath))
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NO_INI"),
                        LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                if (string.IsNullOrEmpty(settings.CreateDwgOutputPath))
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_NO_OUTPUT"),
                        LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Kiểm tra đang mở IDW
                Document activeDoc = _app.ActiveDocument;
                if (activeDoc == null || activeDoc.DocumentType != DocumentTypeEnum.kDrawingDocumentObject)
                {
                    MessageBox.Show(LanguageManager.L("MSG_SAVE_IDW_DRAWING_ONLY"),
                        LanguageManager.L("TITLE_WARNING"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string idwPath = activeDoc.FullFileName;
                if (string.IsNullOrEmpty(idwPath) || !File.Exists(idwPath))
                {
                    MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_IDW_NOT_SAVED"),
                        LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Tạo output folder nếu chưa có
                if (!Directory.Exists(settings.CreateDwgOutputPath))
                {
                    try { Directory.CreateDirectory(settings.CreateDwgOutputPath); }
                    catch (Exception ex)
                    {
                        MessageBox.Show(LanguageManager.L("MSG_CREATE_DWG_CANNOT_CREATE_FOLDER") + ex.Message,
                            LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                string dwgFileName = Path.GetFileNameWithoutExtension(idwPath) + ".dwg";
                string dwgPath = Path.Combine(settings.CreateDwgOutputPath, dwgFileName);

                // Hỏi ghi đè nếu file đã tồn tại
                if (File.Exists(dwgPath))
                {
                    if (MessageBox.Show(
                        LanguageManager.L("MSG_CREATE_DWG_OVERWRITE", dwgFileName),
                        LanguageManager.L("TITLE_CREATE_DWG_CONFIRM"),
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                        return;
                }

                // Export
                ExportIdwToDwg(idwPath, dwgPath, settings.CreateDwgIniPath);

                MessageBox.Show(
                    LanguageManager.L("MSG_CREATE_DWG_IDW_DONE", dwgPath),
                    LanguageManager.L("TITLE_CREATE_DWG_CONFIRM"),
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LicenseHelper.WriteLog("Error in RunCreateDwgFromIdw", ex);
                MessageBox.Show(LanguageManager.L("MSG_PROCESSING_ERROR") + ex.Message,
                    LanguageManager.L("TITLE_ERROR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // PRIVATE HELPERS
        // ============================================================

        private void CollectAssemblyModelFiles(
            AssemblyDocument asmDoc,
            Dictionary<string, string> modelMap,
            Dictionary<string, List<string>> duplicateMap)
        {
            foreach (Document doc in asmDoc.AllReferencedDocuments)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(doc.FullFileName)) continue;
                    string ext = Path.GetExtension(doc.FullFileName).ToLower();
                    if (ext != ".ipt" && ext != ".iam") continue;

                    string baseName = Path.GetFileNameWithoutExtension(doc.FullFileName);
                    if (!modelMap.ContainsKey(baseName))
                    {
                        modelMap.Add(baseName, doc.FullFileName);
                    }
                    else
                    {
                        if (!duplicateMap.ContainsKey(baseName))
                        {
                            duplicateMap.Add(baseName, new List<string>());
                            duplicateMap[baseName].Add(modelMap[baseName]);
                        }
                        if (!duplicateMap[baseName].Contains(doc.FullFileName))
                            duplicateMap[baseName].Add(doc.FullFileName);
                    }
                }
                catch { }
            }
        }

        private string FindIdwInSameFolderOnly(string modelPath)
        {
            if (string.IsNullOrWhiteSpace(modelPath)) return "";
            string folder = Path.GetDirectoryName(modelPath);
            string idwPath = Path.Combine(folder, Path.GetFileNameWithoutExtension(modelPath) + ".idw");
            return File.Exists(idwPath) ? idwPath : "";
        }

        private HashSet<string> ReadWantedBaseNamesFromTextFile(string txtPath)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(txtPath, Encoding.UTF8))
            {
                string s = line.Trim().Replace("\"", "").Trim();
                if (string.IsNullOrWhiteSpace(s)) continue;
                s = Path.GetFileName(s);
                string ext = Path.GetExtension(s).ToLower();
                result.Add(ext == ".ipt" || ext == ".iam" || ext == ".idw"
                    ? Path.GetFileNameWithoutExtension(s) : s);
            }
            return result;
        }

        /// <summary>
        /// Export IDW → DWG dùng DWG Translator AddIn của Inventor
        /// GUID: {C24E3AC4-122E-11D5-8E91-0010B541CD80}
        /// </summary>
        private void ExportIdwToDwg(string idwPath, string dwgOutputPath, string iniFilePath)
        {
            DrawingDocument oDoc = null;
            bool wasAlreadyOpen = false;

            foreach (Document doc in _app.Documents)
            {
                if (string.Equals(doc.FullFileName, idwPath, StringComparison.OrdinalIgnoreCase))
                {
                    oDoc = doc as DrawingDocument;
                    if (oDoc != null) { wasAlreadyOpen = true; break; }
                }
            }

            if (oDoc == null)
                oDoc = (DrawingDocument)_app.Documents.Open(idwPath, false);

            // Inventor API dùng indexer [] không phải method ()
            TranslatorAddIn dwgAddIn = (TranslatorAddIn)_app.ApplicationAddIns.ItemById[
                "{C24E3AC4-122E-11D5-8E91-0010B541CD80}"];
            if (dwgAddIn == null)
                throw new Exception("Cannot get DWG Translator AddIn.");

            TranslationContext oContext = _app.TransientObjects.CreateTranslationContext();
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

            NameValueMap oOptions = _app.TransientObjects.CreateNameValueMap();
            if (dwgAddIn.HasSaveCopyAsOptions[oDoc, oContext, oOptions])
                oOptions.Value["Export_Acad_IniFile"] = iniFilePath;

            DataMedium oDataMedium = _app.TransientObjects.CreateDataMedium();
            oDataMedium.FileName = dwgOutputPath;

            dwgAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium);

            if (!wasAlreadyOpen)
                oDoc.Close(true);
        }
    }
}