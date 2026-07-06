using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Management;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// LicenseHelper v1.3.2 - FIX: Tự động cập nhật static state sau mỗi lần CheckLicenseAsync
    /// </summary>
    public static class LicenseHelper
    {
        // ============================================================
        // PHẦN 1: APPS SCRIPT API CONFIG
        // ============================================================
        private static readonly string _eApiUrl = "OhEEHBJZSmkaDxdIIhFeCw4MAioMQgZOP0odDQIRCjVGH0pgGQMJDwMZDiA/JA9IBzc0PyYpXSEnWwNUJSoYJlNRJg8rAjpxHgMbCiIyBh4GNBVKAzodWwsCCQceWhN4FiMoDxdTLn8MGx0ONx0VDw==";
        private static readonly string _eApiKey = "AAAAAAAAAAAAAAATYldF";

        private static readonly byte[] _k1 = { 0x52, 0x65, 0x70, 0x6C };
        private static readonly byte[] _k2 = { 0x61, 0x63, 0x65, 0x46 };
        private static readonly byte[] _k3 = { 0x69, 0x6C, 0x65, 0x21 };

        private static string DecryptXor(string encrypted, byte[] key)
        {
            try
            {
                byte[] data = Convert.FromBase64String(encrypted);
                byte[] decoded = new byte[data.Length];
                for (int i = 0; i < data.Length; i++)
                    decoded[i] = (byte)(data[i] ^ key[i % key.Length]);
                return Encoding.UTF8.GetString(decoded);
            }
            catch { return ""; }
        }

        private static string GetApiUrl()
        {
            byte[] key = _k1.Concat(_k2).Concat(_k3).ToArray();
            return DecryptXor(_eApiUrl, key);
        }

        private static string GetApiKey()
        {
            byte[] key = _k1.Concat(_k2).Concat(_k3).ToArray();
            return DecryptXor(_eApiKey, key);
        }

        // ============================================================
        // PHẦN 2: CONFIG
        // ============================================================
        private const int TRIAL_MINUTES = 21600;
        private const int REMINDER_THRESHOLD_MINUTES = 2880;
        private const string REGISTRY_REMINDER_VALUE = "LastReminder";
        private const int OFFLINE_GRACE_DAYS = 7;

        private static readonly string _appFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "InventorAddins");

        private static readonly string CachePath = Path.Combine(_appFolder, "ocda_cfg.dat");
        private static readonly string TrialPath = Path.Combine(_appFolder, "ocda_init.dat");
        private static readonly string LogPath = Path.Combine(_appFolder, "ocda_log.txt");

        // ============================================================
        // PHẦN 3: MÃ HÓA AES
        // ============================================================
        private static byte[] GetEncryptionKey()
        {
            string hwid = GetHardwareId();
            string salt = "OCDA#2024$Inv!";
            using (var sha = SHA256.Create())
            {
                return sha.ComputeHash(Encoding.UTF8.GetBytes(hwid + salt));
            }
        }

        private static string EncryptString(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            try
            {
                byte[] key = GetEncryptionKey();
                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.GenerateIV();

                    using (var encryptor = aes.CreateEncryptor())
                    using (var ms = new MemoryStream())
                    {
                        ms.Write(aes.IV, 0, aes.IV.Length);

                        using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                        using (var sw = new StreamWriter(cs))
                        {
                            sw.Write(plainText);
                        }
                        return Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
            catch { return ""; }
        }

        private static string DecryptString(string cipherText)
        {
            if (string.IsNullOrEmpty(cipherText)) return "";
            try
            {
                byte[] key = GetEncryptionKey();
                byte[] fullCipher = Convert.FromBase64String(cipherText);

                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;

                    byte[] iv = new byte[16];
                    Array.Copy(fullCipher, 0, iv, 0, iv.Length);
                    aes.IV = iv;

                    using (var decryptor = aes.CreateDecryptor())
                    using (var ms = new MemoryStream(fullCipher, iv.Length, fullCipher.Length - iv.Length))
                    using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                    using (var sr = new StreamReader(cs))
                    {
                        return sr.ReadToEnd();
                    }
                }
            }
            catch { return ""; }
        }

        // ============================================================
        // PHẦN 4: CHECKSUM
        // ============================================================
        private static string ComputeChecksum(string data)
        {
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(data + GetHardwareId() + "OCDA_CHK"));
                return Convert.ToBase64String(hash).Substring(0, 16);
            }
        }

        // ============================================================
        // PHẦN 5: TRIAL
        // ============================================================
        private const string REGISTRY_PATH = @"Software\Autodesk\Inventor\Addins\OCDA";
        private const string REGISTRY_VALUE = "InitDate";

        private static DateTime GetTrialStartDate()
        {
            DateTime fileDate = DateTime.MinValue;
            DateTime regDate = DateTime.MinValue;

            try
            {
                if (File.Exists(TrialPath))
                {
                    string encrypted = File.ReadAllText(TrialPath);
                    string decrypted = DecryptString(encrypted);
                    if (DateTime.TryParse(decrypted, out DateTime fd))
                    {
                        if (fd <= DateTime.Now) fileDate = fd;
                    }
                }
            }
            catch { }

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(REGISTRY_PATH))
                {
                    if (key != null)
                    {
                        string regValue = key.GetValue(REGISTRY_VALUE) as string;
                        if (!string.IsNullOrEmpty(regValue))
                        {
                            string decrypted = DecryptString(regValue);
                            if (DateTime.TryParse(decrypted, out DateTime rd))
                            {
                                if (rd <= DateTime.Now) regDate = rd;
                            }
                        }
                    }
                }
            }
            catch { }

            DateTime oldest = DateTime.MinValue;
            if (fileDate != DateTime.MinValue && regDate != DateTime.MinValue)
                oldest = fileDate < regDate ? fileDate : regDate;
            else if (fileDate != DateTime.MinValue)
                oldest = fileDate;
            else if (regDate != DateTime.MinValue)
                oldest = regDate;

            if (oldest != DateTime.MinValue)
            {
                if (fileDate != oldest) SaveTrialToFile(oldest);
                if (regDate != oldest) SaveTrialToRegistry(oldest);
            }

            return oldest;
        }

        private static void SaveTrialToFile(DateTime startDate)
        {
            try
            {
                if (!Directory.Exists(_appFolder))
                    Directory.CreateDirectory(_appFolder);

                string encrypted = EncryptString(startDate.ToString("yyyy-MM-dd HH:mm:ss"));
                File.WriteAllText(TrialPath, encrypted);
                File.SetAttributes(TrialPath, FileAttributes.Hidden);
            }
            catch { }
        }

        private static void SaveTrialToRegistry(DateTime startDate)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH))
                {
                    if (key != null)
                    {
                        string encrypted = EncryptString(startDate.ToString("yyyy-MM-dd HH:mm:ss"));
                        key.SetValue(REGISTRY_VALUE, encrypted);
                    }
                }
            }
            catch { }
        }

        private static void InitializeTrial()
        {
            try
            {
                DateTime existing = GetTrialStartDate();
                if (existing != DateTime.MinValue) return;

                DateTime startDate = DateTime.Now;
                SaveTrialToFile(startDate);
                SaveTrialToRegistry(startDate);
            }
            catch { }
        }

        private static (bool isValid, int minutesLeft, DateTime startDate) GetTrialInfo()
        {
            DateTime startDate = GetTrialStartDate();

            if (startDate == DateTime.MinValue)
            {
                InitializeTrial();
                return (true, TRIAL_MINUTES, DateTime.Now);
            }

            DateTime expireDate = startDate.AddMinutes(TRIAL_MINUTES);
            int minutesLeft = (int)(expireDate - DateTime.Now).TotalMinutes;

            return (minutesLeft > 0, Math.Max(0, minutesLeft), startDate);
        }

        public static string FormatTimeLeft(int minutesLeft, bool vietnamese)
        {
            if (minutesLeft >= 1440)
            {
                int days = minutesLeft / 1440;
                return vietnamese ? $"{days} ngày" : $"{days} days";
            }
            else if (minutesLeft >= 60)
            {
                int hours = minutesLeft / 60;
                return vietnamese ? $"{hours} giờ" : $"{hours} hours";
            }
            else
            {
                return vietnamese ? $"{minutesLeft} phút" : $"{minutesLeft} minutes";
            }
        }

        public static (bool isTrialActive, int minutesRemaining, string timeLeftDisplay) GetTrialStatus(bool vietnamese = false)
        {
            var info = GetTrialInfo();
            string display = FormatTimeLeft(info.minutesLeft, vietnamese);
            return (info.isValid, info.minutesLeft, display);
        }

        public static void ShowTrialReminderIfNeeded(bool vietnamese)
        {
            try
            {
                if (!_isTrialMode) return;

                var trialInfo = GetTrialInfo();
                if (!trialInfo.isValid) return;
                if (trialInfo.minutesLeft > REMINDER_THRESHOLD_MINUTES) return;

                DateTime lastReminder = GetLastReminderDate();
                if (lastReminder.Date == DateTime.Now.Date) return;

                SaveLastReminderDate(DateTime.Now);

                string timeLeft = FormatTimeLeft(trialInfo.minutesLeft, vietnamese);
                string message = vietnamese
                    ? $"⚠️ Thời gian dùng thử sắp hết!\n\nBạn còn {timeLeft} để trải nghiệm Mini Tool Add-in.\n\nVui lòng kích hoạt license để tiếp tục sử dụng."
                    : $"⚠️ Trial period ending soon!\n\nYou have {timeLeft} left to try Mini Tool Add-in.\n\nPlease activate your license to continue using.";

                MessageBox.Show(message,
                    vietnamese ? "Nhắc nhở Trial" : "Trial Reminder",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            catch { }
        }

        private static DateTime GetLastReminderDate()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(REGISTRY_PATH))
                {
                    if (key != null)
                    {
                        string value = key.GetValue(REGISTRY_REMINDER_VALUE) as string;
                        if (!string.IsNullOrEmpty(value) && DateTime.TryParse(value, out DateTime dt))
                            return dt;
                    }
                }
            }
            catch { }
            return DateTime.MinValue;
        }

        private static void SaveLastReminderDate(DateTime date)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH))
                {
                    if (key != null)
                    {
                        key.SetValue(REGISTRY_REMINDER_VALUE, date.ToString("yyyy-MM-dd"));
                    }
                }
            }
            catch { }
        }

        // ============================================================
        // PHẦN 6: LOG
        // ============================================================
        public static void WriteLog(string message, Exception ex = null)
        {
            try
            {
                if (!Directory.Exists(_appFolder))
                    Directory.CreateDirectory(_appFolder);

                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
                if (ex != null)
                    logEntry += $"\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
                logEntry += "\n";

                File.AppendAllText(LogPath, logEntry);

                FileInfo fi = new FileInfo(LogPath);
                if (fi.Exists && fi.Length > 1024 * 1024)
                {
                    string[] lines = File.ReadAllLines(LogPath);
                    if (lines.Length > 100)
                        File.WriteAllLines(LogPath, lines.Skip(lines.Length - 100).ToArray());
                }
            }
            catch { }
        }

        public static string GetLogPath() => LogPath;

        // ============================================================
        // PHẦN 7: LICENSE CLASSES
        // ============================================================
        public enum LicenseStatus { Active, Inactive, Error }

        public class LicenseInfo
        {
            public LicenseStatus Status { get; set; } = LicenseStatus.Inactive;
            public string Message { get; set; } = "";
            public string ExpirationDate { get; set; } = "";
            public DateTime LastCheckDate { get; set; } = DateTime.MinValue;
            public string LicensedTo { get; set; } = "";

            public bool IsTrial { get; set; } = false;
            public DateTime TrialStartDate { get; set; } = DateTime.MinValue;
            public int TrialMinutesLeft { get; set; } = 0;
            public string TrialTimeLeftDisplay { get; set; } = "";

            public DateTime LastOnlineVerify { get; set; } = DateTime.MinValue;

            internal string _hwid { get; set; } = "";
            internal string _chk { get; set; } = "";
        }

        public class UpdateInfo
        {
            public bool HasUpdate { get; set; } = false;
            public bool ForceUpdate { get; set; } = false;
            public string NewVersion { get; set; } = "";
            public string DownloadUrl { get; set; } = "";
            public string Message { get; set; } = "";
            public string HelpUrl { get; set; } = "";
            public string OtherAddinUrl { get; set; } = "";
        }

        // ============================================================
        // PHẦN 8: HARDWARE ID
        // ============================================================
        private static string _cachedHwId = null;

        public static string GetHardwareId()
        {
            if (_cachedHwId != null) return _cachedHwId;

            StringBuilder sb = new StringBuilder();

            try
            {
                using (ManagementClass mc = new ManagementClass("Win32_ComputerSystemProduct"))
                {
                    foreach (ManagementObject mo in mc.GetInstances())
                    {
                        string uuid = mo.Properties["UUID"]?.Value?.ToString();
                        if (!string.IsNullOrEmpty(uuid) && uuid != "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF")
                            sb.Append(uuid);
                        break;
                    }
                }
            }
            catch { }

            try
            {
                using (ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
                {
                    foreach (ManagementObject mo in mos.Get())
                    {
                        sb.Append(mo["ProcessorId"]?.ToString() ?? "");
                        break;
                    }
                }
            }
            catch { }

            try
            {
                using (ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                {
                    foreach (ManagementObject mo in mos.Get())
                    {
                        sb.Append(mo["SerialNumber"]?.ToString() ?? "");
                        break;
                    }
                }
            }
            catch { }

            if (sb.Length == 0)
            {
                _cachedHwId = "UNKNOWN-" + Environment.MachineName;
                return _cachedHwId;
            }

            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
                string hex = BitConverter.ToString(hash).Replace("-", "").Substring(0, 16);
                _cachedHwId = $"{hex.Substring(0, 4)}-{hex.Substring(4, 4)}-{hex.Substring(8, 4)}-{hex.Substring(12, 4)}";
            }

            return _cachedHwId;
        }

        public static void CopyHardwareIdToClipboard()
        {
            try { Clipboard.SetText(GetHardwareId()); } catch { }
        }

        public static void PreloadHardwareId()
        {
            Task.Run(() =>
            {
                try { GetHardwareId(); } catch { }
            });
        }

        // ============================================================
        // PHẦN 9: CHECK LICENSE - [FIX v1.3.2]
        // TỰ ĐỘNG CẬP NHẬT STATIC STATE SAU MỖI LẦN CHECK!
        // ============================================================
        public static async Task<LicenseInfo> CheckLicenseAsync(bool forceCheckOnline)
        {
            LicenseInfo cachedInfo = new LicenseInfo();

            // 1. Đọc Cache
            if (File.Exists(CachePath))
            {
                try
                {
                    string encryptedData = File.ReadAllText(CachePath);
                    string decryptedJson = DecryptString(encryptedData);

                    if (!string.IsNullOrEmpty(decryptedJson))
                    {
                        cachedInfo = ParseJsonFromCache(decryptedJson);

                        // Verify integrity
                        if (!VerifyLicenseIntegrity(cachedInfo))
                        {
                            forceCheckOnline = true;
                        }
                    }
                }
                catch
                {
                    cachedInfo = new LicenseInfo();
                }
            }

            // [FIX v1.3.2] Nếu cache là TRIAL và trial đã EXPIRED → PHẢI check online
            if (cachedInfo.IsTrial)
            {
                var trialInfo = GetTrialInfo();
                if (!trialInfo.isValid)
                {
                    forceCheckOnline = true;
                }
            }

            // 2. Dùng cache nếu hợp lệ
            if (!forceCheckOnline && cachedInfo.LastCheckDate != DateTime.MinValue)
            {
                double daysSinceLastCheck = (DateTime.Now - cachedInfo.LastCheckDate).TotalDays;

                if (cachedInfo.Status == LicenseStatus.Active && !cachedInfo.IsTrial)
                {
                    bool isLifetime = !string.IsNullOrEmpty(cachedInfo.ExpirationDate) &&
                                      cachedInfo.ExpirationDate.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0;

                    double maxCacheDays = isLifetime ? 7 : 1;

                    if (daysSinceLastCheck < maxCacheDays)
                    {
                        // [FIX v1.3.2] CẬP NHẬT STATE TRƯỚC KHI RETURN!
                        UpdateStaticState(cachedInfo);
                        return cachedInfo;
                    }
                }
                else if (cachedInfo.IsTrial)
                {
                    var trialInfo = GetTrialInfo();
                    if (trialInfo.isValid)
                    {
                        cachedInfo.TrialMinutesLeft = trialInfo.minutesLeft;
                        cachedInfo.TrialTimeLeftDisplay = FormatTimeLeft(trialInfo.minutesLeft, false);

                        // [FIX v1.3.2] CẬP NHẬT STATE TRƯỚC KHI RETURN!
                        UpdateStaticState(cachedInfo);
                        return cachedInfo;
                    }
                }
            }

            // 3. Check Online
            try
            {
                string apiUrl = GetApiUrl();
                string apiKey = GetApiKey();

                if (string.IsNullOrEmpty(apiUrl))
                {
                    var errorInfo = ApplyTrialIfNeeded(new LicenseInfo { Status = LicenseStatus.Error, Message = "Config error" });
                    UpdateStaticState(errorInfo);
                    return errorInfo;
                }

                string myHwId = GetHardwareId();
                string requestUrl = $"{apiUrl}?action=check&key={apiKey}&hwid={Uri.EscapeDataString(myHwId)}";

                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(15);
                    client.DefaultRequestHeaders.Add("User-Agent", "OpenCadDrawingAddin/2.2");

                    string jsonResponse = await client.GetStringAsync(requestUrl);
                    var newInfo = ParseApiResponse(jsonResponse);

                    if (newInfo.Status != LicenseStatus.Active)
                    {
                        newInfo = ApplyTrialIfNeeded(newInfo);
                    }

                    newInfo.LastOnlineVerify = DateTime.Now;
                    newInfo._hwid = myHwId;
                    newInfo._chk = ComputeChecksum(newInfo.Status.ToString() + newInfo.ExpirationDate + newInfo.IsTrial.ToString());

                    SaveCacheSecure(newInfo);

                    // [FIX v1.3.2] CẬP NHẬT STATE SAU KHI CHECK ONLINE!
                    UpdateStaticState(newInfo);

                    return newInfo;
                }
            }
            catch (Exception ex)
            {
                WriteLog("License check failed (offline)", ex);

                // Offline fallback
                if (cachedInfo.LastCheckDate != DateTime.MinValue && cachedInfo.Status == LicenseStatus.Active && !cachedInfo.IsTrial)
                {
                    if (!string.IsNullOrEmpty(cachedInfo.ExpirationDate))
                    {
                        bool isLifetime = cachedInfo.ExpirationDate.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0;

                        if (!isLifetime && DateTime.TryParse(cachedInfo.ExpirationDate, out DateTime expDate))
                        {
                            if (expDate.Date < DateTime.Now.Date)
                            {
                                cachedInfo.Status = LicenseStatus.Inactive;
                                cachedInfo.Message = "License expired";
                            }
                        }
                    }

                    // [FIX v1.3.2] CẬP NHẬT STATE!
                    UpdateStaticState(cachedInfo);
                    return cachedInfo;
                }

                var trialFallback = ApplyTrialIfNeeded(new LicenseInfo { Status = LicenseStatus.Error, Message = "Network error" });
                UpdateStaticState(trialFallback);
                return trialFallback;
            }
        }

        /// <summary>
        /// [FIX v1.3.2] Cập nhật tất cả static variables từ LicenseInfo
        /// Được gọi TỰ ĐỘNG sau mỗi lần CheckLicenseAsync
        /// </summary>
        private static void UpdateStaticState(LicenseInfo info)
        {
            if (info == null) return;

            _licenseValid = (info.Status == LicenseStatus.Active);
            _isTrialMode = info.IsTrial;
            _lastOnlineVerify = info.LastOnlineVerify;
            _isLifetime = !string.IsNullOrEmpty(info.ExpirationDate) &&
                          info.ExpirationDate.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0;

            UpdateExpirationCache(info);

            _checkCompleted = true;

            WriteLog($"UpdateStaticState: valid={_licenseValid}, trial={_isTrialMode}, lifetime={_isLifetime}, exp={info.ExpirationDate}");
        }

        private static LicenseInfo ParseApiResponse(string json)
        {
            var info = new LicenseInfo
            {
                Status = LicenseStatus.Inactive,
                Message = "Parse error",
                LastCheckDate = DateTime.Now
            };

            try
            {
                var successMatch = Regex.Match(json, "\"success\"\\s*:\\s*(true|false)");
                if (!successMatch.Success || successMatch.Groups[1].Value != "true")
                {
                    var errorMatch = Regex.Match(json, "\"error\"\\s*:\\s*\"([^\"]+)\"");
                    info.Message = errorMatch.Success ? errorMatch.Groups[1].Value : "API error";
                    return info;
                }

                var statusMatch = Regex.Match(json, "\"status\"\\s*:\\s*\"([^\"]+)\"");
                if (statusMatch.Success)
                {
                    info.Status = statusMatch.Groups[1].Value == "Active"
                        ? LicenseStatus.Active
                        : LicenseStatus.Inactive;
                }

                var expMatch = Regex.Match(json, "\"expiration\"\\s*:\\s*\"([^\"]+)\"");
                if (expMatch.Success)
                    info.ExpirationDate = expMatch.Groups[1].Value;

                var msgMatch = Regex.Match(json, "\"message\"\\s*:\\s*\"([^\"]+)\"");
                if (msgMatch.Success)
                    info.Message = msgMatch.Groups[1].Value;

                var nameMatch = Regex.Match(json, "\"name\"\\s*:\\s*\"([^\"]+)\"");
                if (nameMatch.Success)
                    info.LicensedTo = nameMatch.Groups[1].Value;
            }
            catch { }

            return info;
        }

        private static LicenseInfo ApplyTrialIfNeeded(LicenseInfo info)
        {
            var trialInfo = GetTrialInfo();

            if (trialInfo.isValid)
            {
                info.Status = LicenseStatus.Active;
                info.IsTrial = true;
                info.TrialStartDate = trialInfo.startDate;
                info.TrialMinutesLeft = trialInfo.minutesLeft;
                info.TrialTimeLeftDisplay = FormatTimeLeft(trialInfo.minutesLeft, false);
                info.ExpirationDate = trialInfo.startDate.AddMinutes(TRIAL_MINUTES).ToString("yyyy-MM-dd HH:mm");
                info.Message = $"Trial ({info.TrialTimeLeftDisplay} left)";
            }
            else
            {
                info.Status = LicenseStatus.Inactive;
                info.IsTrial = true;
                info.TrialMinutesLeft = 0;
                info.TrialTimeLeftDisplay = "";
                info.Message = "Trial expired";
            }

            info.LastCheckDate = DateTime.Now;
            return info;
        }

        private static bool VerifyLicenseIntegrity(LicenseInfo info)
        {
            if (string.IsNullOrEmpty(info._chk)) return false;
            if (info._hwid != GetHardwareId()) return false;
            string expectedChk = ComputeChecksum(info.Status.ToString() + info.ExpirationDate + info.IsTrial.ToString());
            return info._chk == expectedChk;
        }

        private static void SaveCacheSecure(LicenseInfo info)
        {
            try
            {
                if (!Directory.Exists(_appFolder))
                    Directory.CreateDirectory(_appFolder);

                string json = CreateJson(info);
                string encrypted = EncryptString(json);
                File.WriteAllText(CachePath, encrypted);
                File.SetAttributes(CachePath, FileAttributes.Hidden);
            }
            catch { }
        }

        // ============================================================
        // PHẦN 10: BACKGROUND CHECK & IsBlocked
        // ============================================================
        private static bool _licenseValid = true;
        private static bool _forceUpdate = false;
        private static bool _checkCompleted = false;
        private static bool _isTrialMode = false;

        private static DateTime? _cachedExpiration = null;
        private static bool _expirationLoaded = false;
        private static DateTime _lastOnlineVerify = DateTime.MinValue;
        private static bool _isLifetime = false;

        public static bool IsBlocked
        {
            get
            {
                try
                {
                    if (_forceUpdate) return true;

                    // Chưa check xong → cho qua
                    if (!_checkCompleted) return false;

                    if (_isTrialMode)
                    {
                        var trialInfo = GetTrialInfo();
                        if (!trialInfo.isValid)
                        {
                            _licenseValid = false;
                            return true;
                        }
                    }
                    else
                    {
                        if (!VerifyLicenseExpiration())
                        {
                            _licenseValid = false;
                            return true;
                        }
                    }

                    return !_licenseValid;
                }
                catch
                {
                    return false;
                }
            }
        }

        public static bool IsForceUpdate => _forceUpdate;
        public static bool IsLicenseValid => _licenseValid;

        public static async void StartBackgroundCheck()
        {
            try
            {
                GetHardwareId();

                // [FIX v1.3.2] CheckLicenseAsync giờ tự động cập nhật state
                await CheckLicenseAsync(false);
            }
            catch
            {
                _licenseValid = true;
                _checkCompleted = true;
            }
        }

        public static void SetForceUpdate(bool force)
        {
            _forceUpdate = force;
        }

        public static void ResetLicenseState()
        {
            _licenseValid = true;
            _forceUpdate = false;
            _checkCompleted = false;
            _isTrialMode = false;
            _cachedExpiration = null;
            _expirationLoaded = false;
            _lastOnlineVerify = DateTime.MinValue;
            _isLifetime = false;
        }

        // ============================================================
        // PHẦN 10B: VERIFY EXPIRATION
        // ============================================================
        private static bool VerifyLicenseExpiration()
        {
            try
            {
                if (_expirationLoaded && _cachedExpiration.HasValue)
                {
                    if (_cachedExpiration.Value == DateTime.MaxValue) return true;
                    return _cachedExpiration.Value.Date >= DateTime.Now.Date;
                }

                if (!_expirationLoaded)
                {
                    _expirationLoaded = true;
                    LoadExpirationFromCache();
                }

                if (!_cachedExpiration.HasValue) return true;
                if (_cachedExpiration.Value == DateTime.MaxValue) return true;

                return _cachedExpiration.Value.Date >= DateTime.Now.Date;
            }
            catch
            {
                return true;
            }
        }

        private static void LoadExpirationFromCache()
        {
            try
            {
                if (!File.Exists(CachePath)) return;

                string encrypted = File.ReadAllText(CachePath);
                if (string.IsNullOrEmpty(encrypted)) return;

                string json = DecryptString(encrypted);
                if (string.IsNullOrEmpty(json)) return;

                var trialMatch = Regex.Match(json, "\"IsTrial\"\\s*:\\s*(true|false)", RegexOptions.IgnoreCase);
                if (trialMatch.Success && trialMatch.Groups[1].Value.ToLower() == "true")
                {
                    _cachedExpiration = DateTime.MaxValue;
                    return;
                }

                var expMatch = Regex.Match(json, "\"ExpirationDate\"\\s*:\\s*\"([^\"]+)\"");
                if (!expMatch.Success) return;

                string expStr = expMatch.Groups[1].Value;

                if (expStr.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _cachedExpiration = DateTime.MaxValue;
                    _isLifetime = true;
                    return;
                }

                if (DateTime.TryParse(expStr, out DateTime expDate))
                {
                    _cachedExpiration = expDate;
                }

                var onlineMatch = Regex.Match(json, "\"LastOnlineVerify\"\\s*:\\s*\"([^\"]+)\"");
                if (onlineMatch.Success && DateTime.TryParse(onlineMatch.Groups[1].Value, out DateTime onlineDate))
                {
                    _lastOnlineVerify = onlineDate;
                }
            }
            catch { }
        }

        private static void UpdateExpirationCache(LicenseInfo info)
        {
            try
            {
                if (info == null) return;

                if (info.IsTrial)
                {
                    _cachedExpiration = DateTime.MaxValue;
                }
                else if (!string.IsNullOrEmpty(info.ExpirationDate))
                {
                    if (info.ExpirationDate.IndexOf("Lifetime", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _cachedExpiration = DateTime.MaxValue;
                        _isLifetime = true;
                    }
                    else if (DateTime.TryParse(info.ExpirationDate, out DateTime expDate))
                    {
                        _cachedExpiration = expDate;
                    }
                }

                _lastOnlineVerify = info.LastOnlineVerify;
                _expirationLoaded = true;
            }
            catch { }
        }

        // ============================================================
        // PHẦN 11: UPDATE CHECK
        // ============================================================
        public static async Task<UpdateInfo> CheckForUpdateAsync(string currentVersion)
        {
            var updateInfo = new UpdateInfo { HasUpdate = false, Message = "" };

            try
            {
                string apiUrl = GetApiUrl();
                string apiKey = GetApiKey();
                if (string.IsNullOrEmpty(apiUrl)) return updateInfo;

                string requestUrl = $"{apiUrl}?action=update&key={apiKey}&version={Uri.EscapeDataString(currentVersion)}";

                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(15);
                    client.DefaultRequestHeaders.Add("User-Agent", "OpenCadDrawingAddin/2.2");

                    string jsonResponse = await client.GetStringAsync(requestUrl);

                    var hasUpdateMatch = Regex.Match(jsonResponse, "\"hasUpdate\"\\s*:\\s*(true|false)");
                    if (hasUpdateMatch.Success && hasUpdateMatch.Groups[1].Value == "true")
                        updateInfo.HasUpdate = true;

                    var versionMatch = Regex.Match(jsonResponse, "\"newVersion\"\\s*:\\s*\"([^\"]+)\"");
                    if (versionMatch.Success)
                        updateInfo.NewVersion = versionMatch.Groups[1].Value;

                    var downloadMatch = Regex.Match(jsonResponse, "\"downloadUrl\"\\s*:\\s*\"([^\"]+)\"");
                    if (downloadMatch.Success)
                        updateInfo.DownloadUrl = downloadMatch.Groups[1].Value;

                    var forceMatch = Regex.Match(jsonResponse, "\"forceUpdate\"\\s*:\\s*(true|false)");
                    if (forceMatch.Success)
                        updateInfo.ForceUpdate = forceMatch.Groups[1].Value == "true";

                    var helpMatch = Regex.Match(jsonResponse, "\"helpUrl\"\\s*:\\s*\"([^\"]+)\"");
                    if (helpMatch.Success)
                        updateInfo.HelpUrl = helpMatch.Groups[1].Value;

                    var otherMatch = Regex.Match(jsonResponse, "\"otherAddinUrl\"\\s*:\\s*\"([^\"]+)\"");
                    if (otherMatch.Success)
                        updateInfo.OtherAddinUrl = otherMatch.Groups[1].Value;
                }
            }
            catch (Exception ex)
            {
                updateInfo.HasUpdate = false;
                updateInfo.Message = "Update check error: " + ex.Message;
            }

            return updateInfo;
        }

        // ============================================================
        // PHẦN 12: HELPER FUNCTIONS
        // ============================================================
        private static LicenseInfo ParseJsonFromCache(string json)
        {
            var info = new LicenseInfo();
            try
            {
                var statusMatch = Regex.Match(json, "\"Status\"\\s*:\\s*\"([^\"]+)\"");
                if (statusMatch.Success && statusMatch.Groups[1].Value == "Active")
                    info.Status = LicenseStatus.Active;

                var msgMatch = Regex.Match(json, "\"Message\"\\s*:\\s*\"([^\"]+)\"");
                if (msgMatch.Success) info.Message = msgMatch.Groups[1].Value;

                var expMatch = Regex.Match(json, "\"ExpirationDate\"\\s*:\\s*\"([^\"]+)\"");
                if (expMatch.Success) info.ExpirationDate = expMatch.Groups[1].Value;

                var dateMatch = Regex.Match(json, "\"LastCheckDate\"\\s*:\\s*\"([^\"]+)\"");
                if (dateMatch.Success && DateTime.TryParse(dateMatch.Groups[1].Value, out DateTime dt))
                    info.LastCheckDate = dt;

                var hwidMatch = Regex.Match(json, "\"_hwid\"\\s*:\\s*\"([^\"]+)\"");
                if (hwidMatch.Success) info._hwid = hwidMatch.Groups[1].Value;

                var chkMatch = Regex.Match(json, "\"_chk\"\\s*:\\s*\"([^\"]+)\"");
                if (chkMatch.Success) info._chk = chkMatch.Groups[1].Value;

                var trialMatch = Regex.Match(json, "\"IsTrial\"\\s*:\\s*(true|false)", RegexOptions.IgnoreCase);
                if (trialMatch.Success) info.IsTrial = trialMatch.Groups[1].Value.ToLower() == "true";

                var trialStartMatch = Regex.Match(json, "\"TrialStartDate\"\\s*:\\s*\"([^\"]+)\"");
                if (trialStartMatch.Success && DateTime.TryParse(trialStartMatch.Groups[1].Value, out DateTime trialDt))
                    info.TrialStartDate = trialDt;

                var minutesLeftMatch = Regex.Match(json, "\"TrialMinutesLeft\"\\s*:\\s*(\\d+)");
                if (minutesLeftMatch.Success && int.TryParse(minutesLeftMatch.Groups[1].Value, out int mins))
                    info.TrialMinutesLeft = mins;

                var timeDisplayMatch = Regex.Match(json, "\"TrialTimeLeftDisplay\"\\s*:\\s*\"([^\"]+)\"");
                if (timeDisplayMatch.Success)
                    info.TrialTimeLeftDisplay = timeDisplayMatch.Groups[1].Value;

                var onlineMatch = Regex.Match(json, "\"LastOnlineVerify\"\\s*:\\s*\"([^\"]+)\"");
                if (onlineMatch.Success && DateTime.TryParse(onlineMatch.Groups[1].Value, out DateTime onlineDt))
                    info.LastOnlineVerify = onlineDt;
            }
            catch { }
            return info;
        }

        private static string CreateJson(LicenseInfo info)
        {
            return $@"{{
    ""Status"": ""{info.Status}"",
    ""Message"": ""{EscapeJson(info.Message)}"",
    ""ExpirationDate"": ""{EscapeJson(info.ExpirationDate)}"",
    ""LastCheckDate"": ""{info.LastCheckDate:yyyy-MM-dd HH:mm:ss}"",
    ""LastOnlineVerify"": ""{info.LastOnlineVerify:yyyy-MM-dd HH:mm:ss}"",
    ""IsTrial"": {info.IsTrial.ToString().ToLower()},
    ""TrialStartDate"": ""{info.TrialStartDate:yyyy-MM-dd HH:mm:ss}"",
    ""TrialMinutesLeft"": {info.TrialMinutesLeft},
    ""TrialTimeLeftDisplay"": ""{EscapeJson(info.TrialTimeLeftDisplay)}"",
    ""_hwid"": ""{info._hwid}"",
    ""_chk"": ""{info._chk}""
}}";
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }

        // ============================================================
        // PHẦN 13: TRACKING
        // ============================================================
        private static readonly string _eTrackUrl = "OhEEHBJZSmkaDxdIIhFeCw4MAioMQgZOP0odDQIRCjVGH0pgGQMJDwMaMzQcKB0QOzE1DiI2PWs/XgAZH1IpKSgREBwBGxwMNAwICzUNFQALCFJQAToXHwIaLRxeAAZYMCkhGlYEHS8DXB9CNUoVFAQA";
        private static readonly string _eTrackKey = "ACMxOBMCBi1bXFcU";

        private static string GetTrackUrl()
        {
            try
            {
                byte[] key = _k1.Concat(_k2).Concat(_k3).ToArray();
                return DecryptXor(_eTrackUrl, key);
            }
            catch { return ""; }
        }

        private static string GetTrackKey()
        {
            byte[] key = _k1.Concat(_k2).Concat(_k3).ToArray();
            return DecryptXor(_eTrackKey, key);
        }

        public static async void SendTrackingAsync(string addinVersion, string inventorVersion)
        {
            try
            {
                string trackUrl = GetTrackUrl();
                if (string.IsNullOrEmpty(trackUrl)) return;

                string hwid = GetHardwareId();
                string ip = await GetPublicIPAsync();
                string winVer = GetWindowsVersion();
                string trackKey = GetTrackKey();

                string url = $"{trackUrl}?key={Uri.EscapeDataString(trackKey)}" +
                             $"&hwid={Uri.EscapeDataString(hwid)}" +
                             $"&ip={Uri.EscapeDataString(ip)}" +
                             $"&av={Uri.EscapeDataString(addinVersion)}" +
                             $"&iv={Uri.EscapeDataString(inventorVersion)}" +
                             $"&wv={Uri.EscapeDataString(winVer)}";

                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    await client.GetStringAsync(url);
                }
            }
            catch { }
        }

        private static async Task<string> GetPublicIPAsync()
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(5);
                    string[] services = {
                        "https://api.ipify.org",
                        "https://icanhazip.com",
                        "https://ipinfo.io/ip"
                    };

                    foreach (var service in services)
                    {
                        try
                        {
                            string ip = await client.GetStringAsync(service);
                            ip = ip.Trim();
                            if (!string.IsNullOrEmpty(ip) && ip.Contains("."))
                                return ip;
                        }
                        catch { continue; }
                    }
                }
            }
            catch { }
            return "Unknown";
        }

        private static string GetWindowsVersion()
        {
            try
            {
                var os = Environment.OSVersion;
                string version = $"Windows {os.Version.Major}.{os.Version.Minor}";

                if (os.Version.Major == 10)
                {
                    if (os.Version.Build >= 22000)
                        version = "Windows 11";
                    else
                        version = "Windows 10";
                }

                return $"{version} (Build {os.Version.Build})";
            }
            catch { return "Unknown"; }
        }

        // ============================================================
        // PHẦN 14: USAGE TRACKING
        // ============================================================
        private const string REGISTRY_USAGE_DATE = "UsageDate";
        private const string REGISTRY_USAGE_COUNT = "UsageCount";
        private const string REGISTRY_LAST_SENT = "UsageLastSent";

        public static void IncrementUsageCounter()
        {
            try
            {
                string today = DateTime.Now.ToString("yyyy-MM-dd");

                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH))
                {
                    if (key == null) return;

                    string savedDate = key.GetValue(REGISTRY_USAGE_DATE) as string ?? "";
                    int count = 0;

                    if (savedDate == today)
                    {
                        int.TryParse(key.GetValue(REGISTRY_USAGE_COUNT)?.ToString(), out count);
                    }

                    count++;
                    key.SetValue(REGISTRY_USAGE_DATE, today);
                    key.SetValue(REGISTRY_USAGE_COUNT, count.ToString());
                }
            }
            catch { }
        }

        public static async void SendUsageIfNewDay()
        {
            try
            {
                string today = DateTime.Now.ToString("yyyy-MM-dd");
                string yesterday = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");

                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH))
                {
                    if (key == null) return;

                    string lastSent = key.GetValue(REGISTRY_LAST_SENT) as string ?? "";
                    if (lastSent == today) return;

                    string savedDate = key.GetValue(REGISTRY_USAGE_DATE) as string ?? "";
                    int count = 0;
                    int.TryParse(key.GetValue(REGISTRY_USAGE_COUNT)?.ToString(), out count);

                    if (savedDate == yesterday && count > 0)
                    {
                        bool success = await SendUsageAsync(yesterday, count);
                        if (success)
                        {
                            key.SetValue(REGISTRY_LAST_SENT, today);
                            key.SetValue(REGISTRY_USAGE_DATE, today);
                            key.SetValue(REGISTRY_USAGE_COUNT, "0");
                        }
                    }
                    else if (savedDate != today)
                    {
                        key.SetValue(REGISTRY_USAGE_DATE, today);
                        key.SetValue(REGISTRY_USAGE_COUNT, "0");
                        key.SetValue(REGISTRY_LAST_SENT, today);
                    }
                }
            }
            catch { }
        }

        private static async Task<bool> SendUsageAsync(string date, int count)
        {
            try
            {
                string trackUrl = GetTrackUrl();
                if (string.IsNullOrEmpty(trackUrl)) return false;

                string hwid = GetHardwareId();
                string trackKey = GetTrackKey();

                string url = $"{trackUrl}?key={Uri.EscapeDataString(trackKey)}" +
                             $"&action=usage" +
                             $"&hwid={Uri.EscapeDataString(hwid)}" +
                             $"&count={count}" +
                             $"&date={Uri.EscapeDataString(date)}";

                using (HttpClient client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    string result = await client.GetStringAsync(url);
                    return result.Contains("success\":true");
                }
            }
            catch { return false; }
        }
    }
}