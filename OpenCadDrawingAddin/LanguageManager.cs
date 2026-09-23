using System;
using System.Collections.Generic;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// Quản lý đa ngôn ngữ (i18n) cho Mini Tool Add-in
    /// Version 2.0 - Thêm BOM Compare + Auto Hole Note
    /// </summary>
    public static class LanguageManager
    {
        public static string CurrentLanguage { get; set; } = "EN";

        public static readonly Dictionary<string, string> SupportedLanguages = new Dictionary<string, string>
        {
            ["EN"] = "English",
            ["VN"] = "Tiếng Việt",
            ["JA"] = "日本語 (Japanese)",
            ["KO"] = "한국어 (Korean)",
            ["ZH"] = "中文 (Chinese)",
            ["ES"] = "Español (Spanish)",
            ["PT"] = "Português (Portuguese)"
        };

        private static readonly Dictionary<string, Dictionary<string, string>> _translations =
            new Dictionary<string, Dictionary<string, string>>
            {
                // ========================================
                // ENGLISH (Default/Fallback)
                // ========================================
                ["EN"] = new Dictionary<string, string>
                {
                    // === TITLES ===
                    ["TITLE_INFO"] = "Information",
                    ["TITLE_ERROR"] = "Error",
                    ["TITLE_WARNING"] = "Warning",

                    // === MESSAGES ===
                    ["MSG_SUCCESS"] = "Success!",
                    ["MSG_PROCESSING_ERROR"] = "Error during processing: ",
                    ["MSG_SETTINGS_SAVED"] = "Settings saved successfully!",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "CAD folder is not configured or does not exist.\nPlease open Settings and set the CAD folder path.",
                    ["MSG_NO_ACTIVE_DOC"] = "No active document found.",
                    ["MSG_SELECT_ONE_PART"] = "In Assembly, please select exactly 1 component before running.",
                    ["MSG_CANNOT_GET_COMPONENT"] = "Cannot get the selected component's document.",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "This function only supports Part or Assembly documents.",
                    ["MSG_FILE_NOT_FOUND"] = "CAD file not found.\nPath checked: {0}",
                    ["MSG_LICENSE_INVALID"] = "License is not valid or has expired.\nPlease activate through 'About / License' button.",

                    // === SETTINGS FORM ===
                    ["SETTINGS_TITLE"] = "Settings - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "Language",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",
                    ["TAB_CREATE_DWG"] = "Create DWG",  // [NEW v1.2]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",  // [NEW v1.3]
                    ["TAB_IDW_CHECK"] = "IDW Check",  // [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",  // [NEW v1.4]

                    // === CAD DRAWING TAB ===
                    ["LBL_CAD_FOLDER"] = "CAD Drawing Folder:",
                    ["LBL_FOLDER_HINT"] = "Folder containing .dwg files",
                    ["CHK_USE_REVISION"] = "Include revision suffix in filename",
                    ["CHK_USE_REVISION_HINT"] = "Example: part.ipt with revision=1 → part-1.dwg\r\n         part.ipt with revision=0 → part.dwg",
                    ["LBL_EXTENSION"] = "File extension:",
                    ["LBL_EXTENSION_HINT"] = "Default: .dwg",

                    // === BOM FORMAT TAB ===
                    ["LBL_BOM_XML_PATH"] = "BOM XML File:",
                    ["LBL_BOM_XML_HINT"] = "Path to .xml BOM customization file",

                    // === CHECK REFERENCE TAB ===
                    ["LBL_CHECK_REF_EXCLUDE"] = "Exclude List File:",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "Path to exclude_list.txt (one keyword per line)",

                    // === CREATE DWG TAB === [NEW v1.2]
                    ["LBL_CREATE_DWG_INI"] = "DWGExport.ini File:",
                    ["LBL_CREATE_DWG_INI_HINT"] = "Export configuration file for DWG format (from Inventor)",
                    ["LBL_CREATE_DWG_LIST"] = "List File (.txt):",
                    ["LBL_CREATE_DWG_LIST_HINT"] = "File names to export (used in 'Export by List' mode)",
                    ["LBL_CREATE_DWG_OUTPUT"] = "Output Folder:",
                    ["LBL_CREATE_DWG_OUTPUT_HINT"] = "Folder where exported DWG files will be saved",

                    // === BUTTONS ===
                    ["BTN_SAVE"] = "Save",
                    ["BTN_CANCEL"] = "Cancel",
                    ["BTN_BROWSE"] = "Browse...",

                    // === ABOUT / LICENSE ===
                    ["ABOUT_VERSION"] = "Version: {0}",
                    ["ABOUT_HWID"] = "Hardware ID: {0}",
                    ["ABOUT_STATUS"] = "Status: {0}",
                    ["ABOUT_EXPIRES"] = "Expires: {0}",
                    ["ABOUT_TRIAL"] = "Trial ({0} remaining)",
                    ["ABOUT_TRIAL_EXPIRED"] = "Trial expired",
                    ["ABOUT_ACTIVE"] = "Active",
                    ["ABOUT_NOT_ACTIVATED"] = "Not activated",
                    ["ABOUT_NOT_CHECKED"] = "Not checked yet",
                    ["ABOUT_CHECKING"] = "Checking...",
                    ["ABOUT_LOADING"] = "Loading information...",
                    ["ABOUT_NEW_VERSION"] = "New version available: {0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ You have the latest version",
                    ["ABOUT_DOWNLOAD"] = "Download Update",
                    ["ABOUT_CHECK_UPDATE"] = "Check Update",
                    ["ABOUT_ACTIVATE"] = "Activate",
                    ["ABOUT_RECHECK"] = "Re-check",
                    ["ABOUT_TIME_LEFT"] = "Time left: {0}",
                    ["ABOUT_EXPIRED"] = "Expired",
                    ["ABOUT_REQUIRED"] = " [REQUIRED]",
                    ["ABOUT_UPDATE_REQUIRED"] = "Update Required",
                    ["ABOUT_UPDATE_MSG"] = "Version {0} is required.\nPlease update to continue.",
                    ["ABOUT_ID_COPIED"] = "Hardware ID copied to clipboard!",
                    ["ABOUT_LICENSE_VALID_MSG"] = "License is valid!\nExpiration: {0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "License is not valid.\nReason: {0}\n\nYour Hardware ID:\n{1}\n\nPlease contact the developer to purchase a license.",
                    ["ABOUT_CONFIGURE_CLEANING"] = "Configure settings",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "Copy ID",
                    ["ABOUT_OTHER_ADDIN"] = "Other Add-ins",
                    ["ABOUT_VIEW_LOG"] = "View Log",
                    ["ABOUT_COPY_EMAIL"] = "Copy",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "Name Update",
                    ["MSG_NAME_UPDATED"] = "Name updated: {0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "Name update completed for all occurrences.",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Update only supports Part or Assembly documents.",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "Save IDW",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDW only works on Drawing documents.",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "Could not get model reference.",
                    ["MSG_SAVE_IDW_SAVED"] = "Drawing saved:\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "Drawing could not be saved for some reason.",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Reference only works in Assembly documents.",
                    ["MSG_CHECK_REF_NO_RESULT"] = "This assembly has no occurrences with BOM Structure = Reference\n(outside the exclude list).",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "Cannot open result file automatically:\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "Occurrences with BOM Structure = Reference (outside exclude list):",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "BOM XML file is not configured or does not exist.\nPlease open Settings > BOM Format tab to set the path.",
                    ["MSG_BOM_LOD_ERROR"] = "LOD in use in drawing; Macro failed!",
                    ["MSG_BOM_INVALID_DOC"] = "Invalid Document! Only Assembly or Drawing documents are supported.",
                    ["MSG_BOM_SUCCESS"] = "BOM format applied successfully!",

                    // === CREATE DWG === [NEW v1.2]
                    ["TITLE_CREATE_DWG_CONFIRM"] = "Create DWG",
                    ["MSG_CREATE_DWG_NO_INI"] = "DWGExport.ini not configured or not found.\nPlease set it in Settings > Create DWG tab.",
                    ["MSG_CREATE_DWG_NO_OUTPUT"] = "Output folder is not configured.\nPlease set it in Settings > Create DWG tab.",
                    ["MSG_CREATE_DWG_NEED_ASSEMBLY"] = "Create DWG only works in Assembly documents.",
                    ["MSG_CREATE_DWG_CANNOT_CREATE_FOLDER"] = "Cannot create output folder: ",
                    ["MSG_CREATE_DWG_NO_MODELS"] = "No model files (.ipt/.iam) found in this assembly.",
                    ["MSG_CREATE_DWG_CHOOSE_MODE"] = "Choose export mode:",
                    ["MSG_CREATE_DWG_MODE_ALL"] = "Export All",
                    ["MSG_CREATE_DWG_MODE_LIST"] = "Export by List",
                    ["MSG_CREATE_DWG_NO_IDW_FOUND"] = "No IDW files found in the same folder as assembly models.",
                    ["MSG_CREATE_DWG_NO_LIST"] = "List file not configured or not found.\nPlease set it in Settings > Create DWG tab.",
                    ["MSG_CREATE_DWG_LIST_EMPTY"] = "List file contains no valid file names.",
                    ["MSG_CREATE_DWG_NO_MATCH"] = "No IDW files matched the list within the assembly.",
                    ["MSG_CREATE_DWG_MISSING_IN_ASM"] = "Not found in assembly: ",
                    ["MSG_CREATE_DWG_NO_IDW_FOR_MODEL"] = "Missing IDW (model found, no .idw in folder): ",
                    ["MSG_CREATE_DWG_PREVIEW_COUNT"] = "IDW files to export: ",
                    ["MSG_CREATE_DWG_PREVIEW_OUTPUT"] = "Output folder: ",
                    ["MSG_CREATE_DWG_PREVIEW_MODE"] = "Mode: ",
                    ["MSG_CREATE_DWG_DUPLICATE_WARNING"] = "Warning - duplicate model names: ",
                    ["MSG_CREATE_DWG_DUPLICATE_USED"] = "Used: ",
                    ["MSG_CREATE_DWG_PREVIEW_FIRST5"] = "First 5 files:",
                    ["MSG_CREATE_DWG_CONFIRM_PROMPT"] = "Click OK to start export.",
                    ["MSG_CREATE_DWG_DONE"] = "DWG export completed.",
                    ["MSG_CREATE_DWG_TOTAL"] = "Total: ",
                    ["MSG_CREATE_DWG_OK"] = "Successful: ",
                    ["MSG_CREATE_DWG_FAIL"] = "Failed: ",
                    // [NEW v1.1] Assembly: hỏi xác nhận xuất theo list
                    ["MSG_CREATE_DWG_ASK_LIST"] = "Export DWG files according to the configured list?",
                    // [NEW v1.1] Drawing: xuất thẳng IDW hiện tại
                    ["MSG_CREATE_DWG_IDW_NOT_SAVED"] = "The current IDW file has not been saved to disk yet.",
                    ["MSG_CREATE_DWG_OVERWRITE"] = "File already exists:\n{0}\nOverwrite?",
                    ["MSG_CREATE_DWG_IDW_DONE"] = "DWG exported successfully:\n{0}",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "Copied:\n{0}\n\nGo to screen 2 → click [Paste Comp] to insert into Assembly.",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "Please click to select a Part or Sub-Assembly before pressing Copy.",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "Cannot get file from selected component.\nPlease select a Part or Sub-Assembly.",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "File not found: {0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "Clipboard does not contain an Inventor component.\nPlease [Copy Comp] first.",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "Please open an Assembly (.iam) file before using Paste.",
                    ["MSG_PASTE_COMP_SUCCESS"] = "Pasted:\n{0}\n\nComponent placed at current camera position.\nYou can drag to adjust.",
                    ["MSG_PLACE_COMP_REPLACED"] = "Replaced with:\n{0}\n\nConstraints preserved.",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "File not found:\n{0}\n\nCheck path or network connection.",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["LBL_COPY_PASTE_NOTIFY"] = "Notifications",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Show Copy notification",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Show Place notification",
                    ["LBL_COPY_PASTE_HINT"] = "Uncheck to skip the popup when working fast across screens.",
                    ["CHK_DONT_SHOW_AGAIN"] = "Don't show this again",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "Auto-check when opening IDW",
                    ["CHK_IDW_CHECK_NAME"] = "Check: Drawing file name matches model file name",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "Check: Material / Appearance matches iProperty",
                    ["LBL_IDW_CHECK_HINT"] = "Runs automatically each time an IDW is opened.",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "To disable this warning, go to Settings → IDW Check.",
                    ["MSG_IDW_NAME_MISMATCH"] = "DRAWING NAME MISMATCH\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "MATERIAL MISMATCH\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "APPEARANCE MISMATCH\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "BOM Compare",
                    ["BOMCMP_HINT"] = "Same window: select IAM/sub-asm then click Pick.\nOther window: select IAM there, click 'BOM Cmp' - auto fills here.",
                    ["BOMCMP_ASM1"] = "Assembly 1:",
                    ["BOMCMP_ASM2"] = "Assembly 2:",
                    ["BOMCMP_NOT_SELECTED"] = "(not selected)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (from other window)",
                    ["BOMCMP_SAME_FILE"] = "Same as Assembly {0}. Please choose a different file.",
                    ["BOMCMP_NO_BOM"] = "Cannot read BOM.",
                    ["BOMCMP_NO_DOC"] = "No document is open in this window.",
                    ["BOMCMP_NOT_ASM"] = "Active document is not an Assembly (.iam).\nSelect an IAM tab or click a sub-asm in the browser.",
                    ["BOMCMP_OCC_IS_PART"] = "Selected occurrence is a Part, not an Assembly.",
                    ["BOMCMP_COMPARING"] = "Comparing...",
                    ["BOMCMP_LOADING"] = "Loading...",
                    ["BOMCMP_SHEET_TREE"] = "Tree Compare",
                    ["BOMCMP_SHEET_PART"] = "Part Compare",
                    ["BOMCMP_COL_STATUS"] = "Status",
                    ["BOMCMP_COL_PART_NAME"] = "Part Name",
                    ["BOMCMP_COL_QTY"] = "Qty",
                    ["BOMCMP_COL_NOTE"] = "Note",
                    ["BOMCMP_NOTE_SAME"] = "Same",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "Different qty",
                    ["BOMCMP_NOTE_NOT_IN"] = "Not in table {0}",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "Text height (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "Cluster radius (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "Tap tolerance (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "Run",
                    ["MSG_HOLE_NEED_IDW"] = "Please open an IDW file before running Auto Hole Note.",
                    ["MSG_HOLE_PICK_VIEW"] = "Click on the view to scan holes:",
                    ["MSG_HOLE_ERROR"] = "Error: {0}",
                },

                // ========================================
                // VIETNAMESE
                // ========================================
                ["VN"] = new Dictionary<string, string>
                {
                    ["TITLE_INFO"] = "Thông báo",
                    ["TITLE_ERROR"] = "Lỗi",
                    ["TITLE_WARNING"] = "Cảnh báo",

                    ["MSG_SUCCESS"] = "Thành công!",
                    ["MSG_PROCESSING_ERROR"] = "Lỗi trong quá trình xử lý: ",
                    ["MSG_SETTINGS_SAVED"] = "Đã lưu cài đặt thành công!",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "Thư mục CAD chưa được cấu hình hoặc không tồn tại.\nVui lòng mở Settings và cài đặt đường dẫn thư mục CAD.",
                    ["MSG_NO_ACTIVE_DOC"] = "Không tìm thấy tài liệu đang mở.",
                    ["MSG_SELECT_ONE_PART"] = "Trong Assembly, vui lòng chọn đúng 1 component trước khi chạy.",
                    ["MSG_CANNOT_GET_COMPONENT"] = "Không thể lấy tài liệu của component được chọn.",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "Chức năng này chỉ hỗ trợ tài liệu Part hoặc Assembly.",
                    ["MSG_FILE_NOT_FOUND"] = "Không tìm thấy file CAD.\nĐường dẫn đã kiểm tra: {0}",
                    ["MSG_LICENSE_INVALID"] = "License không hợp lệ hoặc đã hết hạn.\nVui lòng kích hoạt qua nút 'About / License'.",

                    ["SETTINGS_TITLE"] = "Cài đặt - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "Ngôn ngữ",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",
                    ["TAB_CREATE_DWG"] = "Tạo DWG",  // [NEW v1.2]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",  // [NEW v1.3]
                    ["TAB_IDW_CHECK"] = "IDW Check",  // [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",  // [NEW v1.4]

                    ["LBL_CAD_FOLDER"] = "Thư mục file CAD:",
                    ["LBL_FOLDER_HINT"] = "Thư mục chứa các file .dwg",
                    ["CHK_USE_REVISION"] = "Thêm số sửa đổi vào tên file",
                    ["CHK_USE_REVISION_HINT"] = "Ví dụ: file.ipt với sửa đổi=1 → file-1.dwg\r\n        file.ipt với sửa đổi=0 → file.dwg",
                    ["LBL_EXTENSION"] = "Phần mở rộng file:",
                    ["LBL_EXTENSION_HINT"] = "Mặc định: .dwg",

                    ["LBL_BOM_XML_PATH"] = "File XML BOM:",
                    ["LBL_BOM_XML_HINT"] = "Đường dẫn đến file .xml tùy chỉnh BOM",

                    ["LBL_CHECK_REF_EXCLUDE"] = "File danh sách loại trừ:",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "Đường dẫn đến exclude_list.txt (mỗi dòng 1 từ khóa)",

                    // === TAB TẠO DWG === [NEW v1.2]
                    ["LBL_CREATE_DWG_INI"] = "File DWGExport.ini:",
                    ["LBL_CREATE_DWG_INI_HINT"] = "File cấu hình xuất DWG (lấy từ Inventor)",
                    ["LBL_CREATE_DWG_LIST"] = "File danh sách (.txt):",
                    ["LBL_CREATE_DWG_LIST_HINT"] = "Tên file cần xuất (dùng khi chọn 'Xuất theo danh sách')",
                    ["LBL_CREATE_DWG_OUTPUT"] = "Thư mục lưu DWG:",
                    ["LBL_CREATE_DWG_OUTPUT_HINT"] = "Thư mục lưu các file DWG được xuất ra",

                    ["BTN_SAVE"] = "Lưu",
                    ["BTN_CANCEL"] = "Hủy",
                    ["BTN_BROWSE"] = "Chọn...",

                    ["ABOUT_VERSION"] = "Phiên bản: {0}",
                    ["ABOUT_HWID"] = "Hardware ID: {0}",
                    ["ABOUT_STATUS"] = "Trạng thái: {0}",
                    ["ABOUT_EXPIRES"] = "Hết hạn: {0}",
                    ["ABOUT_TRIAL"] = "Dùng thử (còn {0})",
                    ["ABOUT_TRIAL_EXPIRED"] = "Hết hạn dùng thử",
                    ["ABOUT_ACTIVE"] = "Đã kích hoạt",
                    ["ABOUT_NOT_ACTIVATED"] = "Chưa kích hoạt",
                    ["ABOUT_NOT_CHECKED"] = "Chưa kiểm tra",
                    ["ABOUT_CHECKING"] = "Đang kiểm tra...",
                    ["ABOUT_LOADING"] = "Đang tải thông tin...",
                    ["ABOUT_NEW_VERSION"] = "Có phiên bản mới: {0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ Bạn đang dùng phiên bản mới nhất",
                    ["ABOUT_DOWNLOAD"] = "Tải bản cập nhật",
                    ["ABOUT_CHECK_UPDATE"] = "Kiểm tra cập nhật",
                    ["ABOUT_ACTIVATE"] = "Kích hoạt",
                    ["ABOUT_RECHECK"] = "Kiểm tra lại",
                    ["ABOUT_TIME_LEFT"] = "Còn lại: {0}",
                    ["ABOUT_EXPIRED"] = "Đã hết hạn",
                    ["ABOUT_REQUIRED"] = " [BẮT BUỘC]",
                    ["ABOUT_UPDATE_REQUIRED"] = "Yêu cầu cập nhật",
                    ["ABOUT_UPDATE_MSG"] = "Phiên bản {0} là bắt buộc.\nVui lòng cập nhật để tiếp tục.",
                    ["ABOUT_ID_COPIED"] = "Đã sao chép Hardware ID vào clipboard!",
                    ["ABOUT_LICENSE_VALID_MSG"] = "License hợp lệ!\nHết hạn: {0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "License không hợp lệ.\nLý do: {0}\n\nHardware ID của bạn:\n{1}\n\nVui lòng liên hệ tác giả để mua license.",
                    ["ABOUT_CONFIGURE_CLEANING"] = "Cấu hình cài đặt",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "Copy ID",
                    ["ABOUT_OTHER_ADDIN"] = "Add-in khác",
                    ["ABOUT_VIEW_LOG"] = "Xem Log",
                    ["ABOUT_COPY_EMAIL"] = "Sao chép",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "Cập nhật tên",
                    ["MSG_NAME_UPDATED"] = "Đã cập nhật tên: {0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "Đã cập nhật tên tất cả occurrences.",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Update chỉ hỗ trợ tài liệu Part hoặc Assembly.",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "Lưu IDW",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDW chỉ hoạt động với tài liệu Drawing.",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "Không thể lấy tham chiếu model.",
                    ["MSG_SAVE_IDW_SAVED"] = "Đã lưu bản vẽ:\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "Không thể lưu bản vẽ vì một lý do nào đó.",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Reference chỉ chạy được trong tài liệu Assembly.",
                    ["MSG_CHECK_REF_NO_RESULT"] = "Assembly này không có Occurrence nào BOM Structure = Reference\n(ngoài danh sách loại trừ).",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "Không mở được file tự động:\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "Các Occurrence BOM Structure = Reference (ngoài danh sách loại trừ):",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "File XML BOM chưa được cấu hình hoặc không tồn tại.\nVui lòng mở Settings > tab BOM Format để thiết lập.",
                    ["MSG_BOM_LOD_ERROR"] = "LOD đang được sử dụng trong bản vẽ; Macro thất bại!",
                    ["MSG_BOM_INVALID_DOC"] = "Tài liệu không hợp lệ! Chỉ hỗ trợ Assembly hoặc Drawing.",
                    ["MSG_BOM_SUCCESS"] = "Đã áp dụng định dạng BOM thành công!",

                    // === CREATE DWG === [NEW v1.2]
                    ["TITLE_CREATE_DWG_CONFIRM"] = "Tạo DWG",
                    ["MSG_CREATE_DWG_NO_INI"] = "File DWGExport.ini chưa được cấu hình hoặc không tồn tại.\nVui lòng thiết lập trong Settings > tab Tạo DWG.",
                    ["MSG_CREATE_DWG_NO_OUTPUT"] = "Thư mục lưu DWG chưa được cấu hình.\nVui lòng thiết lập trong Settings > tab Tạo DWG.",
                    ["MSG_CREATE_DWG_NEED_ASSEMBLY"] = "Tính năng Tạo DWG chỉ hoạt động trong tài liệu Assembly.",
                    ["MSG_CREATE_DWG_CANNOT_CREATE_FOLDER"] = "Không thể tạo thư mục output: ",
                    ["MSG_CREATE_DWG_NO_MODELS"] = "Không tìm thấy file model (.ipt/.iam) nào trong bản lắp này.",
                    ["MSG_CREATE_DWG_CHOOSE_MODE"] = "Chọn chế độ xuất:",
                    ["MSG_CREATE_DWG_MODE_ALL"] = "Xuất tất cả",
                    ["MSG_CREATE_DWG_MODE_LIST"] = "Xuất theo danh sách",
                    ["MSG_CREATE_DWG_NO_IDW_FOUND"] = "Không tìm thấy file IDW nào nằm đúng thư mục với các model trong bản lắp.",
                    ["MSG_CREATE_DWG_NO_LIST"] = "File danh sách chưa được cấu hình hoặc không tồn tại.\nVui lòng thiết lập trong Settings > tab Tạo DWG.",
                    ["MSG_CREATE_DWG_LIST_EMPTY"] = "File danh sách không chứa tên file hợp lệ nào.",
                    ["MSG_CREATE_DWG_NO_MATCH"] = "Không có file IDW nào khớp với danh sách trong phạm vi bản lắp.",
                    ["MSG_CREATE_DWG_MISSING_IN_ASM"] = "Không tìm thấy trong assembly: ",
                    ["MSG_CREATE_DWG_NO_IDW_FOR_MODEL"] = "Thiếu IDW (có model, không có .idw cùng thư mục): ",
                    ["MSG_CREATE_DWG_PREVIEW_COUNT"] = "Số IDW sẽ xuất: ",
                    ["MSG_CREATE_DWG_PREVIEW_OUTPUT"] = "Thư mục lưu: ",
                    ["MSG_CREATE_DWG_PREVIEW_MODE"] = "Chế độ: ",
                    ["MSG_CREATE_DWG_DUPLICATE_WARNING"] = "Cảnh báo - tên model trùng: ",
                    ["MSG_CREATE_DWG_DUPLICATE_USED"] = "Đang dùng: ",
                    ["MSG_CREATE_DWG_PREVIEW_FIRST5"] = "5 file đầu tiên:",
                    ["MSG_CREATE_DWG_CONFIRM_PROMPT"] = "Nhấn OK để bắt đầu xuất.",
                    ["MSG_CREATE_DWG_DONE"] = "Đã xuất DWG xong.",
                    ["MSG_CREATE_DWG_TOTAL"] = "Tổng số: ",
                    ["MSG_CREATE_DWG_OK"] = "Thành công: ",
                    ["MSG_CREATE_DWG_FAIL"] = "Thất bại: ",
                    // [NEW v1.1] Assembly: hỏi xác nhận xuất theo list
                    ["MSG_CREATE_DWG_ASK_LIST"] = "Xuất file DWG theo danh sách đã cấu hình không?",
                    // [NEW v1.1] Drawing: xuất thẳng IDW hiện tại
                    ["MSG_CREATE_DWG_IDW_NOT_SAVED"] = "File IDW hiện tại chưa được lưu xuống ổ đĩa.",
                    ["MSG_CREATE_DWG_OVERWRITE"] = "File đã tồn tại:\n{0}\nGhi đè?",
                    ["MSG_CREATE_DWG_IDW_DONE"] = "Xuất DWG thành công:\n{0}",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "Đã copy:\n{0}\n\nSang màn hình 2 → nhấn [Paste Comp] để insert vào Assembly.",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "Hãy click chọn một Part hoặc Sub-Assembly trước khi nhấn Copy.",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "Không lấy được file từ component đang chọn.\nVui lòng chọn một Part hoặc Sub-Assembly.",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "Không tìm thấy file: {0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "Clipboard không chứa component Inventor.\nHãy [Copy Comp] trước.",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "Vui lòng mở một file Assembly (.iam) trước khi Paste.",
                    ["MSG_PASTE_COMP_SUCCESS"] = "Đã paste:\n{0}\n\nComponent được đặt tại vùng camera hiện tại.\nBạn có thể kéo để chỉnh vị trí.",
                    ["MSG_PLACE_COMP_REPLACED"] = "Đã thay thế bằng:\n{0}\n\nRàng buộc được giữ nguyên.",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "Không tìm thấy file:\n{0}\n\nKiểm tra lại đường dẫn hoặc kết nối mạng.",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["LBL_COPY_PASTE_NOTIFY"] = "Thông báo",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Hiển thị thông báo sau khi Copy",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Hiển thị thông báo sau khi Place",
                    ["LBL_COPY_PASTE_HINT"] = "Bỏ tích để ẩn popup khi làm việc nhanh giữa 2 màn hình.",
                    ["CHK_DONT_SHOW_AGAIN"] = "Không hiển thị lại",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "Tự động kiểm tra khi mở IDW",
                    ["CHK_IDW_CHECK_NAME"] = "Kiểm tra: Tên file IDW phải khớp tên file model",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "Kiểm tra: Material / Appearance phải khớp iProperty",
                    ["LBL_IDW_CHECK_HINT"] = "Tự động chạy mỗi khi mở file IDW.",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "Để tắt cảnh báo này, vào Settings → IDW Check.",
                    ["MSG_IDW_NAME_MISMATCH"] = "TÊN FILE KHÔNG KHỚP\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "MATERIAL KHÔNG KHỚP\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "APPEARANCE KHÔNG KHỚP\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "So sanh BOM",
                    ["BOMCMP_HINT"] = "Cung cua so: chon IAM/sub-asm roi bam Pick.\nKhac cua so: chon IAM o cua so kia roi bam 'BOM Cmp' - o nay tu dien.",
                    ["BOMCMP_ASM1"] = "Ban lap 1:",
                    ["BOMCMP_ASM2"] = "Ban lap 2:",
                    ["BOMCMP_NOT_SELECTED"] = "(chua chon)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (tu cua so khac)",
                    ["BOMCMP_SAME_FILE"] = "Trung voi Ban lap {0}. Chon file khac.",
                    ["BOMCMP_NO_BOM"] = "Khong doc duoc BOM.",
                    ["BOMCMP_NO_DOC"] = "Khong co document nao dang mo trong cua so nay.",
                    ["BOMCMP_NOT_ASM"] = "Document hien tai khong phai Assembly (.iam).\nChon tab IAM hoac click sub-asm trong browser.",
                    ["BOMCMP_OCC_IS_PART"] = "Occurrence da chon la Part, khong phai Assembly.",
                    ["BOMCMP_COMPARING"] = "Dang so sanh...",
                    ["BOMCMP_LOADING"] = "Dang tai...",
                    ["BOMCMP_SHEET_TREE"] = "So sanh cum",
                    ["BOMCMP_SHEET_PART"] = "So sanh chi tiet",
                    ["BOMCMP_COL_STATUS"] = "Trang thai",
                    ["BOMCMP_COL_PART_NAME"] = "Ten Part",
                    ["BOMCMP_COL_QTY"] = "SL",
                    ["BOMCMP_COL_NOTE"] = "Ghi chu",
                    ["BOMCMP_NOTE_SAME"] = "Giong nhau",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "Khac so luong",
                    ["BOMCMP_NOTE_NOT_IN"] = "Khong co trong bang {0}",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "Chieu cao chu (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "Ban kinh gom cum (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "Dung sai ren (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "Chay",
                    ["MSG_HOLE_NEED_IDW"] = "Hay mo file IDW truoc khi chay Auto Hole Note.",
                    ["MSG_HOLE_PICK_VIEW"] = "Click vao view can quet lo:",
                    ["MSG_HOLE_ERROR"] = "Loi: {0}",
                },

                // ========================================
                // JAPANESE
                // ========================================
                ["JA"] = new Dictionary<string, string>
                {
                    ["TITLE_INFO"] = "情報",
                    ["TITLE_ERROR"] = "エラー",
                    ["TITLE_WARNING"] = "警告",

                    ["MSG_SUCCESS"] = "成功！",
                    ["MSG_PROCESSING_ERROR"] = "処理エラー: ",
                    ["MSG_SETTINGS_SAVED"] = "設定を保存しました！",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "CADフォルダが設定されていないか、存在しません。\n設定を開いてフォルダを設定してください。",
                    ["MSG_NO_ACTIVE_DOC"] = "アクティブドキュメントが見つかりません。",
                    ["MSG_SELECT_ONE_PART"] = "アセンブリでは、実行前に1つのコンポーネントを選択してください。",
                    ["MSG_CANNOT_GET_COMPONENT"] = "選択したコンポーネントのドキュメントを取得できません。",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "パートまたはアセンブリのドキュメントのみ対応しています。",
                    ["MSG_FILE_NOT_FOUND"] = "CADファイルが見つかりません。\n確認したパス: {0}",
                    ["MSG_LICENSE_INVALID"] = "ライセンスが無効または期限切れです。\n'About / License'ボタンからアクティベートしてください。",

                    ["SETTINGS_TITLE"] = "設定 - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "言語",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",

                    ["LBL_CAD_FOLDER"] = "CAD図面フォルダ:",
                    ["LBL_FOLDER_HINT"] = ".dwgファイルを含むフォルダ",
                    ["CHK_USE_REVISION"] = "ファイル名に改訂番号を追加する",
                    ["CHK_USE_REVISION_HINT"] = "例: part.ipt (改訂=1) → part-1.dwg\r\n    part.ipt (改訂=0) → part.dwg",
                    ["LBL_EXTENSION"] = "ファイル拡張子:",
                    ["LBL_EXTENSION_HINT"] = "デフォルト: .dwg",

                    ["LBL_BOM_XML_PATH"] = "BOM XMLファイル:",
                    ["LBL_BOM_XML_HINT"] = ".xml BOMカスタマイズファイルのパス",

                    ["LBL_CHECK_REF_EXCLUDE"] = "除外リストファイル:",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "exclude_list.txtのパス（1行に1キーワード）",

                    ["BTN_SAVE"] = "保存",
                    ["BTN_CANCEL"] = "キャンセル",
                    ["BTN_BROWSE"] = "参照...",

                    ["ABOUT_VERSION"] = "バージョン: {0}",
                    ["ABOUT_HWID"] = "ハードウェアID: {0}",
                    ["ABOUT_STATUS"] = "ステータス: {0}",
                    ["ABOUT_EXPIRES"] = "有効期限: {0}",
                    ["ABOUT_TRIAL"] = "トライアル ({0} 残り)",
                    ["ABOUT_TRIAL_EXPIRED"] = "トライアル期限切れ",
                    ["ABOUT_ACTIVE"] = "アクティブ",
                    ["ABOUT_NOT_ACTIVATED"] = "未アクティブ",
                    ["ABOUT_NOT_CHECKED"] = "未確認",
                    ["ABOUT_CHECKING"] = "確認中...",
                    ["ABOUT_LOADING"] = "情報を読み込み中...",
                    ["ABOUT_NEW_VERSION"] = "新バージョン: {0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ 最新バージョンです",
                    ["ABOUT_DOWNLOAD"] = "アップデートをダウンロード",
                    ["ABOUT_CHECK_UPDATE"] = "アップデートを確認",
                    ["ABOUT_ACTIVATE"] = "アクティベート",
                    ["ABOUT_RECHECK"] = "再確認",
                    ["ABOUT_TIME_LEFT"] = "残り: {0}",
                    ["ABOUT_EXPIRED"] = "期限切れ",
                    ["ABOUT_REQUIRED"] = " [必須]",
                    ["ABOUT_UPDATE_REQUIRED"] = "アップデートが必要",
                    ["ABOUT_UPDATE_MSG"] = "バージョン {0} が必要です。\n続けるにはアップデートしてください。",
                    ["ABOUT_ID_COPIED"] = "ハードウェアIDをコピーしました！",
                    ["ABOUT_LICENSE_VALID_MSG"] = "ライセンスは有効です！\n有効期限: {0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "ライセンスが無効です。\n理由: {0}\n\nハードウェアID:\n{1}\n\nライセンスの購入は開発者にお問い合わせください。",
                    ["ABOUT_CONFIGURE_CLEANING"] = "設定を構成",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "IDコピー",
                    ["ABOUT_OTHER_ADDIN"] = "他のアドイン",
                    ["ABOUT_VIEW_LOG"] = "ログ表示",
                    ["ABOUT_COPY_EMAIL"] = "コピー",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "名前更新",
                    ["MSG_NAME_UPDATED"] = "名前を更新しました: {0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "すべてのオカレンスの名前を更新しました。",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Updateはパートまたはアセンブリのみ対応しています。",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "IDW保存",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDWは図面ドキュメントのみ対応しています。",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "モデル参照を取得できませんでした。",
                    ["MSG_SAVE_IDW_SAVED"] = "図面を保存しました:\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "何らかの理由で図面を保存できませんでした。",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Referenceはアセンブリドキュメントのみ対応しています。",
                    ["MSG_CHECK_REF_NO_RESULT"] = "このアセンブリには BOM構造 = Reference のオカレンスはありません\n（除外リスト以外）。",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "結果ファイルを自動的に開けませんでした:\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "BOM構造 = Reference のオカレンス（除外リスト以外）:",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "BOM XMLファイルが設定されていないか存在しません。\n設定 > BOM Formatタブでパスを設定してください。",
                    ["MSG_BOM_LOD_ERROR"] = "LODが図面で使用中です。マクロが失敗しました！",
                    ["MSG_BOM_INVALID_DOC"] = "無効なドキュメントです！アセンブリまたは図面のみ対応しています。",
                    ["MSG_BOM_SUCCESS"] = "BOMフォーマットが正常に適用されました！",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "コピーしました:\n{0}\n\n画面2 → [Paste Comp]でアセンブリに挿入してください。",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "コピーする前にパーツまたはサブアセンブリを選択してください。",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "選択したコンポーネントのファイルを取得できません。\nパーツまたはサブアセンブリを選択してください。",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "ファイルが見つかりません: {0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "クリップボードにInventorコンポーネントがありません。\n先に[Copy Comp]を実行してください。",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "貼り付け前にアセンブリ(.iam)ファイルを開いてください。",
                    ["MSG_PASTE_COMP_SUCCESS"] = "貼り付けました:\n{0}\n\n現在のカメラ位置に配置されました。\nドラッグで位置を調整できます。",
                    ["MSG_PLACE_COMP_REPLACED"] = "置き換えました:\n{0}\n\n拘束は保持されています。",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "ファイルが見つかりません:\n{0}\n\nパスまたはネットワーク接続を確認してください。",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",
                    ["LBL_COPY_PASTE_NOTIFY"] = "通知",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Copyの後に通知を表示する",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Placeの後に通知を表示する",
                    ["LBL_COPY_PASTE_HINT"] = "チェックを外すと、2画面間の高速作業時にポップアップを非表示にします。",
                    ["CHK_DONT_SHOW_AGAIN"] = "次回から表示しない",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "IDWを開いたときに自動チェック",
                    ["CHK_IDW_CHECK_NAME"] = "チェック：図面ファイル名とモデルファイル名が一致するか",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "チェック：マテリアル/外観がiPropertyと一致するか",
                    ["LBL_IDW_CHECK_HINT"] = "IDWを開くたびに自動的に実行されます。",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "この警告を無効にするには、設定 → IDW Checkへ。",
                    ["MSG_IDW_NAME_MISMATCH"] = "ファイル名の不一致\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "マテリアルの不一致\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "外観の不一致\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "BOM比較",
                    ["BOMCMP_HINT"] = "同じウィンドウ: IAM/サブアセンブリを選択してPickをクリック。\n別ウィンドウ: そのウィンドウでIAMを選択し「BOM Cmp」ボタンをクリック — 自動入力されます。",
                    ["BOMCMP_ASM1"] = "アセンブリ1:",
                    ["BOMCMP_ASM2"] = "アセンブリ2:",
                    ["BOMCMP_NOT_SELECTED"] = "(未選択)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (他のウィンドウから)",
                    ["BOMCMP_SAME_FILE"] = "アセンブリ{0}と同じです。別のファイルを選択してください。",
                    ["BOMCMP_NO_BOM"] = "BOMを読み取れませんでした。",
                    ["BOMCMP_NO_DOC"] = "このウィンドウにドキュメントが開かれていません。",
                    ["BOMCMP_NOT_ASM"] = "アクティブドキュメントはアセンブリ(.iam)ではありません。\nIAMタブを選択するかブラウザでサブアセンブリをクリック。",
                    ["BOMCMP_OCC_IS_PART"] = "選択したオカレンスはアセンブリではなくパーツです。",
                    ["BOMCMP_COMPARING"] = "比較中...",
                    ["BOMCMP_LOADING"] = "読み込み中...",
                    ["BOMCMP_SHEET_TREE"] = "ツリー比較",
                    ["BOMCMP_SHEET_PART"] = "パーツ比較",
                    ["BOMCMP_COL_STATUS"] = "ステータス",
                    ["BOMCMP_COL_PART_NAME"] = "パーツ名",
                    ["BOMCMP_COL_QTY"] = "数量",
                    ["BOMCMP_COL_NOTE"] = "備考",
                    ["BOMCMP_NOTE_SAME"] = "一致",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "数量が異なる",
                    ["BOMCMP_NOTE_NOT_IN"] = "表{0}にありません",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "文字高さ (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "クラスタ半径 (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "タップ公差 (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "実行",
                    ["MSG_HOLE_NEED_IDW"] = "Auto Hole Noteを実行する前にIDWファイルを開いてください。",
                    ["MSG_HOLE_PICK_VIEW"] = "穴をスキャンするビューをクリック:",
                    ["MSG_HOLE_ERROR"] = "エラー: {0}",
                },

                // ========================================
                // KOREAN
                // ========================================
                ["KO"] = new Dictionary<string, string>
                {
                    ["TITLE_INFO"] = "정보",
                    ["TITLE_ERROR"] = "오류",
                    ["TITLE_WARNING"] = "경고",

                    ["MSG_SUCCESS"] = "성공!",
                    ["MSG_PROCESSING_ERROR"] = "처리 중 오류: ",
                    ["MSG_SETTINGS_SAVED"] = "설정이 저장되었습니다!",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "CAD 폴더가 설정되지 않았거나 존재하지 않습니다.\n설정을 열고 CAD 폴더 경로를 설정하세요.",
                    ["MSG_NO_ACTIVE_DOC"] = "활성 문서를 찾을 수 없습니다.",
                    ["MSG_SELECT_ONE_PART"] = "어셈블리에서 실행 전에 정확히 1개의 부품을 선택하세요.",
                    ["MSG_CANNOT_GET_COMPONENT"] = "선택한 부품의 문서를 가져올 수 없습니다.",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "이 기능은 부품 또는 어셈블리 문서만 지원합니다.",
                    ["MSG_FILE_NOT_FOUND"] = "CAD 파일을 찾을 수 없습니다.\n확인한 경로: {0}",
                    ["MSG_LICENSE_INVALID"] = "라이선스가 유효하지 않거나 만료되었습니다.\n'About / License' 버튼을 통해 활성화하세요.",

                    ["SETTINGS_TITLE"] = "설정 - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "언어",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",

                    ["LBL_CAD_FOLDER"] = "CAD 도면 폴더:",
                    ["LBL_FOLDER_HINT"] = ".dwg 파일이 있는 폴더",
                    ["CHK_USE_REVISION"] = "파일 이름에 개정 번호 포함",
                    ["CHK_USE_REVISION_HINT"] = "예: part.ipt (개정=1) → part-1.dwg\r\n    part.ipt (개정=0) → part.dwg",
                    ["LBL_EXTENSION"] = "파일 확장자:",
                    ["LBL_EXTENSION_HINT"] = "기본값: .dwg",

                    ["LBL_BOM_XML_PATH"] = "BOM XML 파일:",
                    ["LBL_BOM_XML_HINT"] = ".xml BOM 커스터마이즈 파일 경로",

                    ["LBL_CHECK_REF_EXCLUDE"] = "제외 목록 파일:",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "exclude_list.txt 경로 (한 줄에 키워드 하나)",

                    ["BTN_SAVE"] = "저장",
                    ["BTN_CANCEL"] = "취소",
                    ["BTN_BROWSE"] = "찾아보기...",

                    ["ABOUT_VERSION"] = "버전: {0}",
                    ["ABOUT_HWID"] = "하드웨어 ID: {0}",
                    ["ABOUT_STATUS"] = "상태: {0}",
                    ["ABOUT_EXPIRES"] = "만료일: {0}",
                    ["ABOUT_TRIAL"] = "평가판 ({0} 남음)",
                    ["ABOUT_TRIAL_EXPIRED"] = "평가판 만료",
                    ["ABOUT_ACTIVE"] = "활성",
                    ["ABOUT_NOT_ACTIVATED"] = "활성화되지 않음",
                    ["ABOUT_NOT_CHECKED"] = "아직 확인하지 않음",
                    ["ABOUT_CHECKING"] = "확인 중...",
                    ["ABOUT_LOADING"] = "정보 로드 중...",
                    ["ABOUT_NEW_VERSION"] = "새 버전 사용 가능: {0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ 최신 버전을 사용하고 있습니다",
                    ["ABOUT_DOWNLOAD"] = "업데이트 다운로드",
                    ["ABOUT_CHECK_UPDATE"] = "업데이트 확인",
                    ["ABOUT_ACTIVATE"] = "활성화",
                    ["ABOUT_RECHECK"] = "다시 확인",
                    ["ABOUT_TIME_LEFT"] = "남은 시간: {0}",
                    ["ABOUT_EXPIRED"] = "만료됨",
                    ["ABOUT_REQUIRED"] = " [필수]",
                    ["ABOUT_UPDATE_REQUIRED"] = "업데이트 필요",
                    ["ABOUT_UPDATE_MSG"] = "버전 {0}이(가) 필요합니다.\n계속하려면 업데이트하세요.",
                    ["ABOUT_ID_COPIED"] = "하드웨어 ID가 클립보드에 복사되었습니다!",
                    ["ABOUT_LICENSE_VALID_MSG"] = "라이선스가 유효합니다!\n만료일: {0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "라이선스가 유효하지 않습니다.\n이유: {0}\n\n하드웨어 ID:\n{1}\n\n라이선스 구매는 개발자에게 문의하세요.",
                    ["ABOUT_CONFIGURE_CLEANING"] = "설정 구성",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "ID 복사",
                    ["ABOUT_OTHER_ADDIN"] = "다른 애드인",
                    ["ABOUT_VIEW_LOG"] = "로그 보기",
                    ["ABOUT_COPY_EMAIL"] = "복사",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "이름 업데이트",
                    ["MSG_NAME_UPDATED"] = "이름이 업데이트되었습니다: {0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "모든 어커런스의 이름이 업데이트되었습니다.",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Update는 부품 또는 어셈블리 문서만 지원합니다.",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "IDW 저장",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDW는 도면 문서에서만 작동합니다.",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "모델 참조를 가져올 수 없습니다.",
                    ["MSG_SAVE_IDW_SAVED"] = "도면이 저장되었습니다:\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "어떤 이유로 도면을 저장할 수 없었습니다.",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Reference는 어셈블리 문서에서만 작동합니다.",
                    ["MSG_CHECK_REF_NO_RESULT"] = "이 어셈블리에는 BOM 구조 = Reference인 어커런스가 없습니다\n(제외 목록 외).",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "결과 파일을 자동으로 열 수 없습니다:\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "BOM 구조 = Reference인 어커런스 (제외 목록 외):",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "BOM XML 파일이 설정되지 않았거나 존재하지 않습니다.\n설정 > BOM Format 탭에서 경로를 설정하세요.",
                    ["MSG_BOM_LOD_ERROR"] = "도면에서 LOD가 사용 중입니다. 매크로가 실패했습니다!",
                    ["MSG_BOM_INVALID_DOC"] = "잘못된 문서입니다! 어셈블리 또는 도면만 지원됩니다.",
                    ["MSG_BOM_SUCCESS"] = "BOM 형식이 성공적으로 적용되었습니다!",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "복사되었습니다:\n{0}\n\n화면 2 → [Paste Comp]를 클릭하여 어셈블리에 삽입하세요.",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "복사하기 전에 부품 또는 하위 어셈블리를 클릭하여 선택하세요.",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "선택한 컴포넌트의 파일을 가져올 수 없습니다.\n부품 또는 하위 어셈블리를 선택하세요.",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "파일을 찾을 수 없습니다: {0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "클립보드에 Inventor 컴포넌트가 없습니다.\n먼저 [Copy Comp]를 실행하세요.",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "붙여넣기 전에 어셈블리(.iam) 파일을 여세요.",
                    ["MSG_PASTE_COMP_SUCCESS"] = "붙여넣기 완료:\n{0}\n\n현재 카메라 위치에 배치되었습니다.\n드래그하여 위치를 조정하세요.",
                    ["MSG_PLACE_COMP_REPLACED"] = "교체되었습니다:\n{0}\n\n구속 조건이 유지되었습니다.",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "파일을 찾을 수 없습니다:\n{0}\n\n경로 또는 네트워크 연결을 확인하세요.",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",
                    ["LBL_COPY_PASTE_NOTIFY"] = "알림",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Copy 후 알림 표시",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Place 후 알림 표시",
                    ["LBL_COPY_PASTE_HINT"] = "체크 해제 시 두 화면 간 빠른 작업 중 팝업을 숨깁니다.",
                    ["CHK_DONT_SHOW_AGAIN"] = "다시 표시하지 않음",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "IDW 열 때 자동 검사",
                    ["CHK_IDW_CHECK_NAME"] = "검사: 도면 파일명과 모델 파일명 일치 여부",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "검사: 재질/외관이 iProperty와 일치 여부",
                    ["LBL_IDW_CHECK_HINT"] = "IDW 파일을 열 때마다 자동으로 실행됩니다.",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "이 경고를 비활성화하려면 설정 → IDW Check로 이동하세요.",
                    ["MSG_IDW_NAME_MISMATCH"] = "파일명 불일치\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "재질 불일치\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "외관 불일치\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "BOM 비교",
                    ["BOMCMP_HINT"] = "같은 창: IAM/하위 어셈블리를 선택하고 Pick을 클릭.\n다른 창: 다른 창에서 IAM을 선택하고 'BOM Cmp' 버튼을 클릭 — 자동으로 채워집니다.",
                    ["BOMCMP_ASM1"] = "어셈블리 1:",
                    ["BOMCMP_ASM2"] = "어셈블리 2:",
                    ["BOMCMP_NOT_SELECTED"] = "(선택 안 됨)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (다른 창에서)",
                    ["BOMCMP_SAME_FILE"] = "어셈블리 {0}과(와) 동일합니다. 다른 파일을 선택하세요.",
                    ["BOMCMP_NO_BOM"] = "BOM을 읽을 수 없습니다.",
                    ["BOMCMP_NO_DOC"] = "이 창에 열려 있는 문서가 없습니다.",
                    ["BOMCMP_NOT_ASM"] = "활성 문서가 어셈블리(.iam)가 아닙니다.\nIAM 탭을 선택하거나 브라우저에서 하위 어셈블리를 클릭하세요.",
                    ["BOMCMP_OCC_IS_PART"] = "선택한 항목이 어셈블리가 아닌 부품입니다.",
                    ["BOMCMP_COMPARING"] = "비교 중...",
                    ["BOMCMP_LOADING"] = "로딩 중...",
                    ["BOMCMP_SHEET_TREE"] = "트리 비교",
                    ["BOMCMP_SHEET_PART"] = "부품 비교",
                    ["BOMCMP_COL_STATUS"] = "상태",
                    ["BOMCMP_COL_PART_NAME"] = "부품명",
                    ["BOMCMP_COL_QTY"] = "수량",
                    ["BOMCMP_COL_NOTE"] = "비고",
                    ["BOMCMP_NOTE_SAME"] = "동일",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "수량 다름",
                    ["BOMCMP_NOTE_NOT_IN"] = "표 {0}에 없음",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "텍스트 높이 (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "클러스터 반경 (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "탭 공차 (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "실행",
                    ["MSG_HOLE_NEED_IDW"] = "Auto Hole Note를 실행하기 전에 IDW 파일을 열어주세요.",
                    ["MSG_HOLE_PICK_VIEW"] = "홀을 스캔할 뷰를 클릭:",
                    ["MSG_HOLE_ERROR"] = "오류: {0}",
                },

                // ========================================
                // CHINESE SIMPLIFIED
                // ========================================
                ["ZH"] = new Dictionary<string, string>
                {
                    ["TITLE_INFO"] = "信息",
                    ["TITLE_ERROR"] = "错误",
                    ["TITLE_WARNING"] = "警告",

                    ["MSG_SUCCESS"] = "成功！",
                    ["MSG_PROCESSING_ERROR"] = "处理过程中出错：",
                    ["MSG_SETTINGS_SAVED"] = "设置已成功保存！",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "CAD文件夹未配置或不存在。\n请打开设置并设置CAD文件夹路径。",
                    ["MSG_NO_ACTIVE_DOC"] = "未找到活动文档。",
                    ["MSG_SELECT_ONE_PART"] = "在装配体中，请在运行前选择恰好1个零件。",
                    ["MSG_CANNOT_GET_COMPONENT"] = "无法获取所选零件的文档。",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "此功能仅支持零件或装配体文档。",
                    ["MSG_FILE_NOT_FOUND"] = "未找到CAD文件。\n已检查路径：{0}",
                    ["MSG_LICENSE_INVALID"] = "许可证无效或已过期。\n请通过'About / License'按钮激活。",

                    ["SETTINGS_TITLE"] = "设置 - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "语言",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",

                    ["LBL_CAD_FOLDER"] = "CAD图纸文件夹：",
                    ["LBL_FOLDER_HINT"] = "包含.dwg文件的文件夹",
                    ["CHK_USE_REVISION"] = "在文件名中包含修订后缀",
                    ["CHK_USE_REVISION_HINT"] = "例：part.ipt（修订=1）→ part-1.dwg\r\n    part.ipt（修订=0）→ part.dwg",
                    ["LBL_EXTENSION"] = "文件扩展名：",
                    ["LBL_EXTENSION_HINT"] = "默认：.dwg",

                    ["LBL_BOM_XML_PATH"] = "BOM XML文件：",
                    ["LBL_BOM_XML_HINT"] = ".xml BOM自定义文件路径",

                    ["LBL_CHECK_REF_EXCLUDE"] = "排除列表文件：",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "exclude_list.txt路径（每行一个关键词）",

                    ["BTN_SAVE"] = "保存",
                    ["BTN_CANCEL"] = "取消",
                    ["BTN_BROWSE"] = "浏览...",

                    ["ABOUT_VERSION"] = "版本：{0}",
                    ["ABOUT_HWID"] = "硬件ID：{0}",
                    ["ABOUT_STATUS"] = "状态：{0}",
                    ["ABOUT_EXPIRES"] = "到期日：{0}",
                    ["ABOUT_TRIAL"] = "试用版（剩余{0}）",
                    ["ABOUT_TRIAL_EXPIRED"] = "试用期已过",
                    ["ABOUT_ACTIVE"] = "已激活",
                    ["ABOUT_NOT_ACTIVATED"] = "未激活",
                    ["ABOUT_NOT_CHECKED"] = "尚未检查",
                    ["ABOUT_CHECKING"] = "检查中...",
                    ["ABOUT_LOADING"] = "加载信息中...",
                    ["ABOUT_NEW_VERSION"] = "有新版本：{0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ 您使用的是最新版本",
                    ["ABOUT_DOWNLOAD"] = "下载更新",
                    ["ABOUT_CHECK_UPDATE"] = "检查更新",
                    ["ABOUT_ACTIVATE"] = "激活",
                    ["ABOUT_RECHECK"] = "重新检查",
                    ["ABOUT_TIME_LEFT"] = "剩余时间：{0}",
                    ["ABOUT_EXPIRED"] = "已过期",
                    ["ABOUT_REQUIRED"] = " [必须]",
                    ["ABOUT_UPDATE_REQUIRED"] = "需要更新",
                    ["ABOUT_UPDATE_MSG"] = "需要版本 {0}。\n请更新后继续使用。",
                    ["ABOUT_ID_COPIED"] = "硬件ID已复制到剪贴板！",
                    ["ABOUT_LICENSE_VALID_MSG"] = "许可证有效！\n到期日：{0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "许可证无效。\n原因：{0}\n\n您的硬件ID：\n{1}\n\n请联系开发者购买许可证。",
                    ["ABOUT_CONFIGURE_CLEANING"] = "配置设置",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "复制ID",
                    ["ABOUT_OTHER_ADDIN"] = "其他插件",
                    ["ABOUT_VIEW_LOG"] = "查看日志",
                    ["ABOUT_COPY_EMAIL"] = "复制",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "名称更新",
                    ["MSG_NAME_UPDATED"] = "名称已更新：{0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "所有零件的名称已更新。",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Update仅支持零件或装配体文档。",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "保存IDW",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDW仅适用于图纸文档。",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "无法获取模型参考。",
                    ["MSG_SAVE_IDW_SAVED"] = "图纸已保存：\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "由于某种原因无法保存图纸。",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Reference仅适用于装配体文档。",
                    ["MSG_CHECK_REF_NO_RESULT"] = "此装配体中没有BOM结构 = Reference的零件\n（排除列表之外）。",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "无法自动打开结果文件：\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "BOM结构 = Reference的零件（排除列表之外）：",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "BOM XML文件未配置或不存在。\n请打开设置 > BOM Format选项卡设置路径。",
                    ["MSG_BOM_LOD_ERROR"] = "图纸中正在使用LOD，宏失败！",
                    ["MSG_BOM_INVALID_DOC"] = "无效文档！仅支持装配体或图纸。",
                    ["MSG_BOM_SUCCESS"] = "BOM格式已成功应用！",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "已复制：\n{0}\n\n切换到屏幕2 → 点击[Paste Comp]插入装配体。",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "请先点击选择一个零件或子装配体，再按复制。",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "无法获取所选零件的文件。\n请选择一个零件或子装配体。",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "未找到文件：{0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "剪贴板中没有Inventor零件。\n请先执行[Copy Comp]。",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "请先打开一个装配体(.iam)文件再粘贴。",
                    ["MSG_PASTE_COMP_SUCCESS"] = "粘贴成功：\n{0}\n\n零件已放置在当前相机位置。\n可拖动调整位置。",
                    ["MSG_PLACE_COMP_REPLACED"] = "已替换为：\n{0}\n\n约束已保留。",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "未找到文件：\n{0}\n\n请检查路径或网络连接。",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",
                    ["LBL_COPY_PASTE_NOTIFY"] = "通知",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Copy 后显示通知",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Place 后显示通知",
                    ["LBL_COPY_PASTE_HINT"] = "取消勾选可在两屏间快速操作时隐藏弹窗。",
                    ["CHK_DONT_SHOW_AGAIN"] = "不再显示",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "打开IDW时自动检查",
                    ["CHK_IDW_CHECK_NAME"] = "检查：图纸文件名与模型文件名是否一致",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "检查：材质/外观与iProperty是否一致",
                    ["LBL_IDW_CHECK_HINT"] = "每次打开IDW文件时自动运行。",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "要禁用此警告，请转到设置 → IDW Check。",
                    ["MSG_IDW_NAME_MISMATCH"] = "文件名不匹配\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "材质不匹配\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "外观不匹配\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "BOM对比",
                    ["BOMCMP_HINT"] = "同一窗口：选择IAM/子装配后点击Pick。\n不同窗口：在另一窗口选择IAM，点击'BOM Cmp'按钮 — 自动填入。",
                    ["BOMCMP_ASM1"] = "装配体 1:",
                    ["BOMCMP_ASM2"] = "装配体 2:",
                    ["BOMCMP_NOT_SELECTED"] = "(未选择)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (来自其他窗口)",
                    ["BOMCMP_SAME_FILE"] = "与装配体{0}相同。请选择不同的文件。",
                    ["BOMCMP_NO_BOM"] = "无法读取BOM。",
                    ["BOMCMP_NO_DOC"] = "此窗口中没有打开的文档。",
                    ["BOMCMP_NOT_ASM"] = "当前文档不是装配体(.iam)。\n请选择IAM选项卡或在浏览器中点击子装配。",
                    ["BOMCMP_OCC_IS_PART"] = "所选项是零件，不是装配体。",
                    ["BOMCMP_COMPARING"] = "对比中...",
                    ["BOMCMP_LOADING"] = "加载中...",
                    ["BOMCMP_SHEET_TREE"] = "树形对比",
                    ["BOMCMP_SHEET_PART"] = "零件对比",
                    ["BOMCMP_COL_STATUS"] = "状态",
                    ["BOMCMP_COL_PART_NAME"] = "零件名",
                    ["BOMCMP_COL_QTY"] = "数量",
                    ["BOMCMP_COL_NOTE"] = "备注",
                    ["BOMCMP_NOTE_SAME"] = "相同",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "数量不同",
                    ["BOMCMP_NOTE_NOT_IN"] = "不在表{0}中",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "文字高度 (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "聚类半径 (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "丝锥公差 (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "运行",
                    ["MSG_HOLE_NEED_IDW"] = "运行Auto Hole Note前请先打开IDW文件。",
                    ["MSG_HOLE_PICK_VIEW"] = "点击要扫描孔的视图:",
                    ["MSG_HOLE_ERROR"] = "错误: {0}",
                },

                // ========================================
                // SPANISH
                // ========================================
                ["ES"] = new Dictionary<string, string>
                {
                    ["TITLE_INFO"] = "Información",
                    ["TITLE_ERROR"] = "Error",
                    ["TITLE_WARNING"] = "Advertencia",

                    ["MSG_SUCCESS"] = "¡Éxito!",
                    ["MSG_PROCESSING_ERROR"] = "Error durante el procesamiento: ",
                    ["MSG_SETTINGS_SAVED"] = "¡Configuración guardada exitosamente!",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "La carpeta CAD no está configurada o no existe.\nPor favor, abra Configuración y establezca la ruta de la carpeta CAD.",
                    ["MSG_NO_ACTIVE_DOC"] = "No se encontró ningún documento activo.",
                    ["MSG_SELECT_ONE_PART"] = "En Ensamble, seleccione exactamente 1 componente antes de ejecutar.",
                    ["MSG_CANNOT_GET_COMPONENT"] = "No se puede obtener el documento del componente seleccionado.",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "Esta función solo admite documentos de Pieza o Ensamble.",
                    ["MSG_FILE_NOT_FOUND"] = "Archivo CAD no encontrado.\nRuta verificada: {0}",
                    ["MSG_LICENSE_INVALID"] = "La licencia no es válida o ha expirado.\nActive a través del botón 'About / License'.",

                    ["SETTINGS_TITLE"] = "Configuración - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "Idioma",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",

                    ["LBL_CAD_FOLDER"] = "Carpeta de dibujos CAD:",
                    ["LBL_FOLDER_HINT"] = "Carpeta que contiene archivos .dwg",
                    ["CHK_USE_REVISION"] = "Incluir sufijo de revisión en el nombre del archivo",
                    ["CHK_USE_REVISION_HINT"] = "Ejemplo: part.ipt con revisión=1 → part-1.dwg\r\n         part.ipt con revisión=0 → part.dwg",
                    ["LBL_EXTENSION"] = "Extensión de archivo:",
                    ["LBL_EXTENSION_HINT"] = "Predeterminado: .dwg",

                    ["LBL_BOM_XML_PATH"] = "Archivo XML de BOM:",
                    ["LBL_BOM_XML_HINT"] = "Ruta al archivo .xml de personalización de BOM",

                    ["LBL_CHECK_REF_EXCLUDE"] = "Archivo de lista de exclusión:",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "Ruta a exclude_list.txt (una palabra clave por línea)",

                    ["BTN_SAVE"] = "Guardar",
                    ["BTN_CANCEL"] = "Cancelar",
                    ["BTN_BROWSE"] = "Examinar...",

                    ["ABOUT_VERSION"] = "Versión: {0}",
                    ["ABOUT_HWID"] = "ID de hardware: {0}",
                    ["ABOUT_STATUS"] = "Estado: {0}",
                    ["ABOUT_EXPIRES"] = "Expira: {0}",
                    ["ABOUT_TRIAL"] = "Prueba ({0} restante)",
                    ["ABOUT_TRIAL_EXPIRED"] = "Prueba expirada",
                    ["ABOUT_ACTIVE"] = "Activo",
                    ["ABOUT_NOT_ACTIVATED"] = "No activado",
                    ["ABOUT_NOT_CHECKED"] = "Aún no verificado",
                    ["ABOUT_CHECKING"] = "Verificando...",
                    ["ABOUT_LOADING"] = "Cargando información...",
                    ["ABOUT_NEW_VERSION"] = "Nueva versión disponible: {0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ Tiene la última versión",
                    ["ABOUT_DOWNLOAD"] = "Descargar actualización",
                    ["ABOUT_CHECK_UPDATE"] = "Verificar actualización",
                    ["ABOUT_ACTIVATE"] = "Activar",
                    ["ABOUT_RECHECK"] = "Reverificar",
                    ["ABOUT_TIME_LEFT"] = "Tiempo restante: {0}",
                    ["ABOUT_EXPIRED"] = "Expirado",
                    ["ABOUT_REQUIRED"] = " [REQUERIDO]",
                    ["ABOUT_UPDATE_REQUIRED"] = "Actualización requerida",
                    ["ABOUT_UPDATE_MSG"] = "Se requiere la versión {0}.\nActualice para continuar.",
                    ["ABOUT_ID_COPIED"] = "¡ID de hardware copiado al portapapeles!",
                    ["ABOUT_LICENSE_VALID_MSG"] = "¡Licencia válida!\nExpiración: {0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "La licencia no es válida.\nRazón: {0}\n\nSu ID de hardware:\n{1}\n\nContacte al desarrollador para comprar una licencia.",
                    ["ABOUT_CONFIGURE_CLEANING"] = "Configurar ajustes",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "Copiar ID",
                    ["ABOUT_OTHER_ADDIN"] = "Otros complementos",
                    ["ABOUT_VIEW_LOG"] = "Ver registro",
                    ["ABOUT_COPY_EMAIL"] = "Copiar",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "Actualizar nombre",
                    ["MSG_NAME_UPDATED"] = "Nombre actualizado: {0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "Nombres de todas las ocurrencias actualizados.",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Update solo admite documentos de Pieza o Ensamble.",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "Guardar IDW",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDW solo funciona con documentos de Dibujo.",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "No se pudo obtener la referencia del modelo.",
                    ["MSG_SAVE_IDW_SAVED"] = "Dibujo guardado:\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "No se pudo guardar el dibujo por alguna razón.",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Reference solo funciona en documentos de Ensamble.",
                    ["MSG_CHECK_REF_NO_RESULT"] = "Este ensamble no tiene ocurrencias con Estructura BOM = Reference\n(fuera de la lista de exclusión).",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "No se puede abrir el archivo de resultados automáticamente:\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "Ocurrencias con Estructura BOM = Reference (fuera de la lista de exclusión):",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "El archivo XML de BOM no está configurado o no existe.\nAbra Configuración > pestaña BOM Format para establecer la ruta.",
                    ["MSG_BOM_LOD_ERROR"] = "LOD en uso en el dibujo; ¡Macro falló!",
                    ["MSG_BOM_INVALID_DOC"] = "¡Documento inválido! Solo se admiten Ensamble o Dibujo.",
                    ["MSG_BOM_SUCCESS"] = "¡Formato BOM aplicado exitosamente!",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "Copiado:\n{0}\n\nVaya a la pantalla 2 → haga clic en [Paste Comp] para insertar en el Ensamble.",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "Haga clic para seleccionar una Pieza o Subensamble antes de copiar.",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "No se puede obtener el archivo del componente seleccionado.\nSeleccione una Pieza o Subensamble.",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "Archivo no encontrado: {0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "El portapapeles no contiene un componente de Inventor.\nEjecute [Copy Comp] primero.",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "Abra un archivo de Ensamble (.iam) antes de pegar.",
                    ["MSG_PASTE_COMP_SUCCESS"] = "Pegado:\n{0}\n\nComponente colocado en la posición de cámara actual.\nPuede arrastrarlo para ajustar la posición.",
                    ["MSG_PLACE_COMP_REPLACED"] = "Reemplazado con:\n{0}\n\nLas restricciones se han conservado.",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "Archivo no encontrado:\n{0}\n\nVerifique la ruta o la conexión de red.",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",
                    ["LBL_COPY_PASTE_NOTIFY"] = "Notificaciones",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Mostrar notificación de Copy",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Mostrar notificación de Place",
                    ["LBL_COPY_PASTE_HINT"] = "Desmarque para ocultar el popup al trabajar rápido entre pantallas.",
                    ["CHK_DONT_SHOW_AGAIN"] = "No mostrar de nuevo",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "Verificación automática al abrir IDW",
                    ["CHK_IDW_CHECK_NAME"] = "Verificar: El nombre del archivo IDW coincide con el modelo",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "Verificar: Material/Apariencia coincide con iProperty",
                    ["LBL_IDW_CHECK_HINT"] = "Se ejecuta automáticamente cada vez que se abre un IDW.",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "Para desactivar esta advertencia, vaya a Configuración → IDW Check.",
                    ["MSG_IDW_NAME_MISMATCH"] = "Nombre no coincide\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "Material no coincide\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "Apariencia no coincide\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "Comparar BOM",
                    ["BOMCMP_HINT"] = "Misma ventana: seleccione IAM/sub-ensamblaje y haga clic en Pick.\nOtra ventana: seleccione IAM allí y haga clic en 'BOM Cmp' — se autocompleta aquí.",
                    ["BOMCMP_ASM1"] = "Ensamblaje 1:",
                    ["BOMCMP_ASM2"] = "Ensamblaje 2:",
                    ["BOMCMP_NOT_SELECTED"] = "(no seleccionado)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (de otra ventana)",
                    ["BOMCMP_SAME_FILE"] = "Igual que Ensamblaje {0}. Elija un archivo diferente.",
                    ["BOMCMP_NO_BOM"] = "No se puede leer BOM.",
                    ["BOMCMP_NO_DOC"] = "No hay documentos abiertos en esta ventana.",
                    ["BOMCMP_NOT_ASM"] = "El documento activo no es un Ensamblaje (.iam).\nSeleccione una pestaña IAM o haga clic en un sub-ensamblaje.",
                    ["BOMCMP_OCC_IS_PART"] = "La selección es una Pieza, no un Ensamblaje.",
                    ["BOMCMP_COMPARING"] = "Comparando...",
                    ["BOMCMP_LOADING"] = "Cargando...",
                    ["BOMCMP_SHEET_TREE"] = "Comparar árbol",
                    ["BOMCMP_SHEET_PART"] = "Comparar piezas",
                    ["BOMCMP_COL_STATUS"] = "Estado",
                    ["BOMCMP_COL_PART_NAME"] = "Nombre de pieza",
                    ["BOMCMP_COL_QTY"] = "Cant.",
                    ["BOMCMP_COL_NOTE"] = "Nota",
                    ["BOMCMP_NOTE_SAME"] = "Iguales",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "Cantidad diferente",
                    ["BOMCMP_NOTE_NOT_IN"] = "No está en tabla {0}",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "Altura de texto (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "Radio de agrupación (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "Tolerancia de rosca (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "Ejecutar",
                    ["MSG_HOLE_NEED_IDW"] = "Abra un archivo IDW antes de ejecutar Auto Hole Note.",
                    ["MSG_HOLE_PICK_VIEW"] = "Haga clic en la vista para escanear agujeros:",
                    ["MSG_HOLE_ERROR"] = "Error: {0}",
                },

                // ========================================
                // PORTUGUESE
                // ========================================
                ["PT"] = new Dictionary<string, string>
                {
                    ["TITLE_INFO"] = "Informação",
                    ["TITLE_ERROR"] = "Erro",
                    ["TITLE_WARNING"] = "Aviso",

                    ["MSG_SUCCESS"] = "Sucesso!",
                    ["MSG_PROCESSING_ERROR"] = "Erro durante o processamento: ",
                    ["MSG_SETTINGS_SAVED"] = "Configurações salvas com sucesso!",
                    ["MSG_NO_FOLDER_CONFIGURED"] = "A pasta CAD não está configurada ou não existe.\nAbra as Configurações e defina o caminho da pasta CAD.",
                    ["MSG_NO_ACTIVE_DOC"] = "Nenhum documento ativo encontrado.",
                    ["MSG_SELECT_ONE_PART"] = "No Conjunto, selecione exatamente 1 componente antes de executar.",
                    ["MSG_CANNOT_GET_COMPONENT"] = "Não é possível obter o documento do componente selecionado.",
                    ["MSG_UNSUPPORTED_DOC_TYPE"] = "Esta função suporta apenas documentos de Peça ou Conjunto.",
                    ["MSG_FILE_NOT_FOUND"] = "Arquivo CAD não encontrado.\nCaminho verificado: {0}",
                    ["MSG_LICENSE_INVALID"] = "A licença não é válida ou expirou.\nAtive através do botão 'About / License'.",

                    ["SETTINGS_TITLE"] = "Configurações - Mini Tool v2.0",
                    ["TAB_SETTINGS"] = "CAD Drawing",
                    ["TAB_LANGUAGE"] = "Idioma",
                    ["TAB_BOM"] = "BOM Format",
                    ["TAB_CHECK_REF"] = "Check Reference",

                    ["LBL_CAD_FOLDER"] = "Pasta de desenhos CAD:",
                    ["LBL_FOLDER_HINT"] = "Pasta contendo arquivos .dwg",
                    ["CHK_USE_REVISION"] = "Incluir sufixo de revisão no nome do arquivo",
                    ["CHK_USE_REVISION_HINT"] = "Exemplo: part.ipt com revisão=1 → part-1.dwg\r\n         part.ipt com revisão=0 → part.dwg",
                    ["LBL_EXTENSION"] = "Extensão do arquivo:",
                    ["LBL_EXTENSION_HINT"] = "Padrão: .dwg",

                    ["LBL_BOM_XML_PATH"] = "Arquivo XML do BOM:",
                    ["LBL_BOM_XML_HINT"] = "Caminho para o arquivo .xml de personalização do BOM",

                    ["LBL_CHECK_REF_EXCLUDE"] = "Arquivo de lista de exclusão:",
                    ["LBL_CHECK_REF_EXCLUDE_HINT"] = "Caminho para exclude_list.txt (uma palavra-chave por linha)",

                    ["BTN_SAVE"] = "Salvar",
                    ["BTN_CANCEL"] = "Cancelar",
                    ["BTN_BROWSE"] = "Procurar...",

                    ["ABOUT_VERSION"] = "Versão: {0}",
                    ["ABOUT_HWID"] = "ID de hardware: {0}",
                    ["ABOUT_STATUS"] = "Status: {0}",
                    ["ABOUT_EXPIRES"] = "Expira: {0}",
                    ["ABOUT_TRIAL"] = "Avaliação ({0} restante)",
                    ["ABOUT_TRIAL_EXPIRED"] = "Avaliação expirada",
                    ["ABOUT_ACTIVE"] = "Ativo",
                    ["ABOUT_NOT_ACTIVATED"] = "Não ativado",
                    ["ABOUT_NOT_CHECKED"] = "Ainda não verificado",
                    ["ABOUT_CHECKING"] = "Verificando...",
                    ["ABOUT_LOADING"] = "Carregando informações...",
                    ["ABOUT_NEW_VERSION"] = "Nova versão disponível: {0} {1}",
                    ["ABOUT_LATEST_VERSION"] = "✓ Você tem a versão mais recente",
                    ["ABOUT_DOWNLOAD"] = "Baixar atualização",
                    ["ABOUT_CHECK_UPDATE"] = "Verificar atualização",
                    ["ABOUT_ACTIVATE"] = "Ativar",
                    ["ABOUT_RECHECK"] = "Reverificar",
                    ["ABOUT_TIME_LEFT"] = "Tempo restante: {0}",
                    ["ABOUT_EXPIRED"] = "Expirado",
                    ["ABOUT_REQUIRED"] = " [OBRIGATÓRIO]",
                    ["ABOUT_UPDATE_REQUIRED"] = "Atualização necessária",
                    ["ABOUT_UPDATE_MSG"] = "A versão {0} é necessária.\nAtualize para continuar.",
                    ["ABOUT_ID_COPIED"] = "ID de hardware copiado para a área de transferência!",
                    ["ABOUT_LICENSE_VALID_MSG"] = "Licença válida!\nExpiração: {0}",
                    ["ABOUT_LICENSE_INVALID_MSG"] = "A licença não é válida.\nMotivo: {0}\n\nSeu ID de hardware:\n{1}\n\nContate o desenvolvedor para comprar uma licença.",
                    ["ABOUT_CONFIGURE_CLEANING"] = "Configurar ajustes",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "Copiar ID",
                    ["ABOUT_OTHER_ADDIN"] = "Outros suplementos",
                    ["ABOUT_VIEW_LOG"] = "Ver registro",
                    ["ABOUT_COPY_EMAIL"] = "Copiar",

                    // === NAME UPDATE ===
                    ["TITLE_NAME_UPDATE"] = "Atualizar nome",
                    ["MSG_NAME_UPDATED"] = "Nome atualizado: {0}",
                    ["MSG_NAME_UPDATE_ALL_DONE"] = "Nomes de todas as ocorrências atualizados.",
                    ["MSG_NAME_UPDATE_UNSUPPORTED"] = "Name Update suporta apenas documentos de Peça ou Conjunto.",

                    // === SAVE IDW ===
                    ["TITLE_SAVE_IDW"] = "Salvar IDW",
                    ["MSG_SAVE_IDW_DRAWING_ONLY"] = "Save IDW funciona apenas com documentos de Desenho.",
                    ["MSG_SAVE_IDW_NO_MODEL"] = "Não foi possível obter a referência do modelo.",
                    ["MSG_SAVE_IDW_SAVED"] = "Desenho salvo:\n{0}",
                    ["MSG_SAVE_IDW_FAILED"] = "Não foi possível salvar o desenho por algum motivo.",

                    // === CHECK REFERENCE ===
                    ["TITLE_CHECK_REF"] = "Check Reference",
                    ["MSG_CHECK_REF_ASM_ONLY"] = "Check Reference funciona apenas em documentos de Conjunto.",
                    ["MSG_CHECK_REF_NO_RESULT"] = "Este conjunto não tem ocorrências com Estrutura BOM = Reference\n(fora da lista de exclusão).",
                    ["MSG_CHECK_REF_CANNOT_OPEN"] = "Não é possível abrir o arquivo de resultados automaticamente:\n{0}",
                    ["MSG_CHECK_REF_HEADER"] = "Ocorrências com Estrutura BOM = Reference (fora da lista de exclusão):",

                    // === BOM FORMAT ===
                    ["TITLE_BOM_FORMAT"] = "BOM Format",
                    ["MSG_BOM_XML_NOT_SET"] = "O arquivo XML do BOM não está configurado ou não existe.\nAbra Configurações > aba BOM Format para definir o caminho.",
                    ["MSG_BOM_LOD_ERROR"] = "LOD em uso no desenho; Macro falhou!",
                    ["MSG_BOM_INVALID_DOC"] = "Documento inválido! Apenas Conjunto ou Desenho são suportados.",
                    ["MSG_BOM_SUCCESS"] = "Formato BOM aplicado com sucesso!",

                    // === CROSS-SCREEN COPY/PASTE === [NEW v1.3]
                    ["TITLE_COPY_COMP"] = "Copy",
                    ["TITLE_PASTE_COMP"] = "Place",
                    ["MSG_COPY_COMP_SUCCESS"] = "Copiado:\n{0}\n\nVá para a tela 2 → clique em [Paste Comp] para inserir no Conjunto.",
                    ["MSG_COPY_COMP_NO_SELECTION"] = "Clique para selecionar uma Peça ou Subconjunto antes de copiar.",
                    ["MSG_COPY_COMP_CANNOT_GET"] = "Não é possível obter o arquivo do componente selecionado.\nSelecione uma Peça ou Subconjunto.",
                    ["MSG_COPY_COMP_FILE_NOT_FOUND"] = "Arquivo não encontrado: {0}",
                    ["MSG_PASTE_COMP_NO_CLIPBOARD"] = "A área de transferência não contém um componente Inventor.\nExecute [Copy Comp] primeiro.",
                    ["MSG_PASTE_COMP_ASM_ONLY"] = "Abra um arquivo de Conjunto (.iam) antes de colar.",
                    ["MSG_PASTE_COMP_SUCCESS"] = "Colado:\n{0}\n\nComponente posicionado na posição atual da câmera.\nArraste para ajustar a posição.",
                    ["MSG_PLACE_COMP_REPLACED"] = "Substituído por:\n{0}\n\nAs restrições foram preservadas.",
                    ["MSG_PASTE_COMP_FILE_NOT_FOUND"] = "Arquivo não encontrado:\n{0}\n\nVerifique o caminho ou a conexão de rede.",

                    // === COPY-PASTE SETTINGS TAB === [NEW v1.3]
                    ["TAB_COPY_PASTE"] = "Copy-Paste",
                    ["LBL_COPY_PASTE_NOTIFY"] = "Notificações",
                    ["CHK_SHOW_COPY_NOTIFY"] = "Mostrar notificação de Copy",
                    ["CHK_SHOW_PASTE_NOTIFY"] = "Mostrar notificação de Place",
                    ["LBL_COPY_PASTE_HINT"] = "Desmarque para ocultar o popup ao trabalhar rapidamente entre telas.",
                    ["CHK_DONT_SHOW_AGAIN"] = "Não mostrar novamente",

                    // === IDW AUTO CHECK === [NEW v1.4]
                    ["TAB_IDW_CHECK"] = "IDW Check",
                    ["LBL_IDW_CHECK_TITLE"] = "Verificação automática ao abrir IDW",
                    ["CHK_IDW_CHECK_NAME"] = "Verificar: Nome do arquivo IDW coincide com o modelo",
                    ["CHK_IDW_CHECK_APPEARANCE"] = "Verificar: Material/Aparência coincide com iProperty",
                    ["LBL_IDW_CHECK_HINT"] = "Executado automaticamente cada vez que um IDW é aberto.",
                    ["TITLE_IDW_CHECK"] = "IDW Auto Check",
                    ["MSG_IDW_CHECK_DISABLE_HINT"] = "Para desativar este aviso, vá para Configurações → IDW Check.",
                    ["MSG_IDW_NAME_MISMATCH"] = "Nome não coincide\n    IDW   : \"{0}\"\n    Model : \"{1}\" (Sheet: {2})\n\n",
                    ["MSG_IDW_MATERIAL_MISMATCH"] = "Material não coincide\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",
                    ["MSG_IDW_APPEARANCE_MISMATCH"] = "Aparência não coincide\n    Part      : {0}\n    Model     : \"{1}\"\n    iProperty : \"{2}\"\n\n",

                    // === BOM COMPARE === [NEW v2.0]
                    ["BOMCMP_TITLE"] = "Comparar BOM",
                    ["BOMCMP_HINT"] = "Mesma janela: selecione IAM/sub-conjunto e clique em Pick.\nOutra janela: selecione IAM lá e clique no botão 'BOM Cmp' — preenchimento automático aqui.",
                    ["BOMCMP_ASM1"] = "Conjunto 1:",
                    ["BOMCMP_ASM2"] = "Conjunto 2:",
                    ["BOMCMP_NOT_SELECTED"] = "(não selecionado)",
                    ["BOMCMP_FROM_OTHER"] = "\u21BB {0}  (de outra janela)",
                    ["BOMCMP_SAME_FILE"] = "Igual ao Conjunto {0}. Escolha um arquivo diferente.",
                    ["BOMCMP_NO_BOM"] = "Não é possível ler BOM.",
                    ["BOMCMP_NO_DOC"] = "Nenhum documento aberto nesta janela.",
                    ["BOMCMP_NOT_ASM"] = "O documento ativo não é um Conjunto (.iam).\nSelecione uma aba IAM ou clique em um sub-conjunto no navegador.",
                    ["BOMCMP_OCC_IS_PART"] = "A seleção é uma Peça, não um Conjunto.",
                    ["BOMCMP_COMPARING"] = "Comparando...",
                    ["BOMCMP_LOADING"] = "Carregando...",
                    ["BOMCMP_SHEET_TREE"] = "Comparar árvore",
                    ["BOMCMP_SHEET_PART"] = "Comparar peças",
                    ["BOMCMP_COL_STATUS"] = "Status",
                    ["BOMCMP_COL_PART_NAME"] = "Nome da peça",
                    ["BOMCMP_COL_QTY"] = "Qtd.",
                    ["BOMCMP_COL_NOTE"] = "Nota",
                    ["BOMCMP_NOTE_SAME"] = "Iguais",
                    ["BOMCMP_NOTE_DIFF_QTY"] = "Quantidade diferente",
                    ["BOMCMP_NOTE_NOT_IN"] = "Não está na tabela {0}",

                    // === AUTO HOLE NOTE === [NEW v2.0]
                    ["TAB_AUTO_HOLE"] = "Auto Hole",
                    ["LBL_HOLE_TEXT_HEIGHT"] = "Altura do texto (mm):",
                    ["LBL_HOLE_CLUSTER_RADIUS"] = "Raio de agrupamento (mm):",
                    ["LBL_HOLE_TAP_TOL"] = "Tolerância de rosca (x0.01mm):",
                    ["BTN_HOLE_RUN"] = "Executar",
                    ["MSG_HOLE_NEED_IDW"] = "Abra um arquivo IDW antes de executar Auto Hole Note.",
                    ["MSG_HOLE_PICK_VIEW"] = "Clique na vista para escanear furos:",
                    ["MSG_HOLE_ERROR"] = "Erro: {0}",
                },
            };

        // ============================================================
        // PUBLIC API
        // ============================================================

        /// <summary>
        /// Lấy chuỗi dịch theo key. Fallback về EN nếu không có.
        /// </summary>
        public static string L(string key, params object[] args)
        {
            string text = GetRaw(key);
            if (args != null && args.Length > 0)
            {
                try { text = string.Format(text, args); }
                catch { }
            }
            return text;
        }

        private static string GetRaw(string key)
        {
            // Thử ngôn ngữ hiện tại
            if (_translations.TryGetValue(CurrentLanguage, out var dict) &&
                dict.TryGetValue(key, out var val))
                return val;

            // Fallback về EN
            if (_translations.TryGetValue("EN", out var enDict) &&
                enDict.TryGetValue(key, out var enVal))
                return enVal;

            // Trả về key nếu không tìm thấy
            return key;
        }
    }
}