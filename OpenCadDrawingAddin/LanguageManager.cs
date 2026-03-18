using System;
using System.Collections.Generic;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// Quản lý đa ngôn ngữ (i18n) cho Open CAD Drawing Add-in
    /// Version 1.0
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

                    // === SETTINGS FORM ===
                    ["SETTINGS_TITLE"] = "Settings - Mini Tool v1.0",
                    ["TAB_SETTINGS"] = "Cad drawing", // [CHANGED] "Settings" → "Cad drawing"
                    ["TAB_LANGUAGE"] = "Language",
                    ["TAB_BOM"] = "Bom format",  // [NEW]

                    // === BOM FORMAT TAB ===                                          // [NEW BLOCK]
                    ["LBL_BOM_XML_PATH"] = "BOM XML File:",
                    ["LBL_BOM_XML_HINT"] = "Path to .xml BOM customization file",

                    // === SETTINGS CONTROLS ===
                    ["LBL_CAD_FOLDER"] = "CAD Drawing Folder:",
                    ["LBL_FOLDER_HINT"] = "Folder containing .dwg files",
                    ["CHK_USE_REVISION"] = "Include revision suffix in filename",
                    ["CHK_USE_REVISION_HINT"] = "Example: part.ipt with revision=1 → part-1.dwg\r\n         part.ipt with revision=0 → part.dwg",
                    ["LBL_EXTENSION"] = "File extension:",
                    ["LBL_EXTENSION_HINT"] = "Default: .dwg",

                    // === BUTTONS ===
                    ["BTN_SAVE"] = "Save",
                    ["BTN_CANCEL"] = "Cancel",
                    ["BTN_BROWSE"] = "Browse...",

                    // === ABOUT ===
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
                    ["MSG_LICENSE_INVALID"] = "License is not valid or has expired.\nPlease activate through 'About / License' button.",
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

                    ["SETTINGS_TITLE"] = "Cài đặt - Mini Tool v1.0",
                    ["TAB_SETTINGS"] = "Cad drawing", // [CHANGED]
                    ["TAB_LANGUAGE"] = "Ngôn ngữ",
                    ["TAB_BOM"] = "Bom format",  // [NEW]

                    // [NEW BLOCK]
                    ["LBL_BOM_XML_PATH"] = "File XML BOM:",
                    ["LBL_BOM_XML_HINT"] = "Đường dẫn đến file .xml tùy chỉnh BOM",

                    ["LBL_CAD_FOLDER"] = "Thư mục file CAD:",
                    ["LBL_FOLDER_HINT"] = "Thư mục chứa các file .dwg",
                    ["CHK_USE_REVISION"] = "Thêm số sửa đổi vào tên file",
                    ["CHK_USE_REVISION_HINT"] = "Ví dụ: file.ipt với sửa đổi=1 → file-1.dwg\r\n        file.ipt với sửa đổi=0 → file.dwg",
                    ["LBL_EXTENSION"] = "Phần mở rộng file:",
                    ["LBL_EXTENSION_HINT"] = "Mặc định: .dwg",

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
                    ["MSG_LICENSE_INVALID"] = "License không hợp lệ hoặc đã hết hạn.\nVui lòng kích hoạt qua nút 'About / License'.",
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
                    ["SETTINGS_TITLE"] = "設定 - Mini Tool v1.0",
                    ["TAB_SETTINGS"] = "Cad drawing", // [CHANGED]
                    ["TAB_LANGUAGE"] = "言語",
                    ["TAB_BOM"] = "Bom format",  // [NEW]

                    // [NEW BLOCK]
                    ["LBL_BOM_XML_PATH"] = "BOM XMLファイル:",
                    ["LBL_BOM_XML_HINT"] = ".xml BOMカスタマイズファイルのパス",
                    ["LBL_CAD_FOLDER"] = "CAD図面フォルダ:",
                    ["LBL_FOLDER_HINT"] = ".dwgファイルを含むフォルダ",
                    ["CHK_USE_REVISION"] = "ファイル名に改訂番号を追加する",
                    ["CHK_USE_REVISION_HINT"] = "例: part.ipt (改訂=1) → part-1.dwg\r\n    part.ipt (改訂=0) → part.dwg",
                    ["LBL_EXTENSION"] = "ファイル拡張子:",
                    ["LBL_EXTENSION_HINT"] = "デフォルト: .dwg",
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
                    ["ABOUT_LICENSE_INVALID_MSG"] = "ライセンスが無効です。\n理由: {0}\n\nハードウェアID:\n{1}",
                    ["ABOUT_CONFIGURE_CLEANING"] = "設定を構成",
                    ["ABOUT_HELP"] = "?",
                    ["ABOUT_COPY_ID"] = "IDコピー",
                    ["ABOUT_OTHER_ADDIN"] = "他のアドイン",
                    ["ABOUT_VIEW_LOG"] = "ログ表示",
                    ["ABOUT_COPY_EMAIL"] = "コピー",
                    ["MSG_LICENSE_INVALID"] = "ライセンスが無効または期限切れです。\n'About / License'ボタンからアクティベートしてください。",
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