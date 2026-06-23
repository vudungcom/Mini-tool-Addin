using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// BomCompareBridge - [NEW v1.7]
    /// Chia se BOM DATA (da doc san) giua 2 process Inventor qua file tam.
    ///
    /// QUAN TRONG: Khac voi version cu chi luu PATH, version nay luu CA NOI DUNG BOM
    /// (tree + parts da flatten). Ly do: 2 process Inventor co the doc CUNG 1 file ra
    /// data KHAC NHAU (do Vault cache / shared document khac trang thai). Vi vay moi
    /// process PHAI tu doc BOM cua file no dang mo (thay doi minh nhin thay), ghi ra
    /// bridge. Process kia doc lai data do thay vi tu mo file -> tranh doc nham.
    ///
    /// File: %TEMP%\ocda_bomcmp_data\slot_*.txt   (moi slot = 1 BOM data, giu toi da 2)
    /// Format text don gian (tu parse, khong can JSON):
    ///   NAME|&lt;name&gt;
    ///   PATH|&lt;fullpath&gt;
    ///   TIME|&lt;ticks&gt;
    ///   TREE|&lt;level&gt;|&lt;isAsm 0/1&gt;|&lt;name&gt;
    ///   PART|&lt;qty&gt;|&lt;name&gt;
    /// </summary>
    public static class BomCompareBridge
    {
        private static readonly string BridgeDir = Path.Combine(
            Path.GetTempPath(), "ocda_bomcmp_data");

        private const int ExpiryMinutes = 30;

        // ============================================================
        // PUSH: ghi 1 BOM data thanh 1 slot
        // ============================================================

        public static void Push(BomCompareLogic.BomData data)
        {
            if (data == null) return;

            try
            {
                Directory.CreateDirectory(BridgeDir);
                CleanupExpired();

                // Neu da co slot cung path -> ghi de (cap nhat data moi nhat)
                string targetSlot = FindSlotByPath(data.FullPath);

                if (targetSlot == null)
                {
                    // Tao slot moi. Neu da co >= 2 slot, xoa slot cu nhat
                    var slots = GetSlotFiles();
                    if (slots.Count >= 2)
                    {
                        slots.Sort((a, b) => File.GetLastWriteTime(a).CompareTo(File.GetLastWriteTime(b)));
                        try { File.Delete(slots[0]); } catch { }
                    }
                    targetSlot = Path.Combine(BridgeDir, "slot_" + Guid.NewGuid().ToString("N") + ".txt");
                }

                WriteData(targetSlot, data);
            }
            catch { }
        }

        // ============================================================
        // READ: doc tat ca BOM data con hieu luc, moi nhat truoc
        // ============================================================

        public static List<BomCompareLogic.BomData> ReadAll()
        {
            var result = new List<BomCompareLogic.BomData>();
            try
            {
                if (!Directory.Exists(BridgeDir)) return result;
                CleanupExpired();

                var slots = GetSlotFiles();
                slots.Sort((a, b) => File.GetLastWriteTime(b).CompareTo(File.GetLastWriteTime(a)));

                foreach (var slot in slots)
                {
                    var data = ReadData(slot);
                    if (data != null) result.Add(data);
                }
            }
            catch { }
            return result;
        }

        public static void Clear()
        {
            try
            {
                if (!Directory.Exists(BridgeDir)) return;
                foreach (var f in Directory.GetFiles(BridgeDir, "slot_*.txt"))
                {
                    try { File.Delete(f); } catch { }
                }
            }
            catch { }
        }

        // ============================================================
        // INTERNAL
        // ============================================================

        private static List<string> GetSlotFiles()
        {
            var list = new List<string>();
            try
            {
                if (Directory.Exists(BridgeDir))
                    list.AddRange(Directory.GetFiles(BridgeDir, "slot_*.txt"));
            }
            catch { }
            return list;
        }

        private static string FindSlotByPath(string fullPath)
        {
            foreach (var slot in GetSlotFiles())
            {
                try
                {
                    foreach (var line in File.ReadLines(slot))
                    {
                        if (line.StartsWith("PATH|"))
                        {
                            string p = line.Substring(5);
                            if (string.Equals(p, fullPath, StringComparison.OrdinalIgnoreCase))
                                return slot;
                            break;
                        }
                    }
                }
                catch { }
            }
            return null;
        }

        private static void CleanupExpired()
        {
            try
            {
                var cutoff = DateTime.Now.AddMinutes(-ExpiryMinutes);
                foreach (var slot in GetSlotFiles())
                {
                    try
                    {
                        if (File.GetLastWriteTime(slot) < cutoff)
                            File.Delete(slot);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static void WriteData(string path, BomCompareLogic.BomData data)
        {
            var sb = new StringBuilder();
            sb.AppendLine("NAME|" + data.Name);
            sb.AppendLine("PATH|" + data.FullPath);
            sb.AppendLine("TIME|" + DateTime.Now.Ticks);

            foreach (var n in data.TreeNodes)
                sb.AppendLine("TREE|" + n.Level + "|" + (n.IsAssembly ? "1" : "0") + "|" + n.Name);

            foreach (var p in data.Parts)
                sb.AppendLine("PART|" + p.Qty + "|" + p.Name);

            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                    using (var sw = new StreamWriter(fs, new UTF8Encoding(false)))
                    {
                        sw.Write(sb.ToString());
                    }
                    return;
                }
                catch
                {
                    System.Threading.Thread.Sleep(40);
                }
            }
        }

        private static BomCompareLogic.BomData ReadData(string path)
        {
            try
            {
                var data = new BomCompareLogic.BomData();
                string[] lines;
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var sr = new StreamReader(fs, Encoding.UTF8))
                {
                    lines = sr.ReadToEnd().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                }

                foreach (var line in lines)
                {
                    if (line.StartsWith("NAME|"))
                        data.Name = line.Substring(5);
                    else if (line.StartsWith("PATH|"))
                        data.FullPath = line.Substring(5);
                    else if (line.StartsWith("TREE|"))
                    {
                        var rest = line.Substring(5);
                        int s1 = rest.IndexOf('|');
                        int s2 = rest.IndexOf('|', s1 + 1);
                        if (s1 > 0 && s2 > s1)
                        {
                            int level = int.Parse(rest.Substring(0, s1));
                            bool isAsm = rest.Substring(s1 + 1, s2 - s1 - 1) == "1";
                            string name = rest.Substring(s2 + 1);
                            data.TreeNodes.Add(new BomCompareLogic.TreeNodeData
                            { Name = name, Level = level, IsAssembly = isAsm });
                        }
                    }
                    else if (line.StartsWith("PART|"))
                    {
                        var rest = line.Substring(5);
                        int s1 = rest.IndexOf('|');
                        if (s1 > 0)
                        {
                            int qty = int.Parse(rest.Substring(0, s1));
                            string name = rest.Substring(s1 + 1);
                            data.Parts.Add(new BomCompareLogic.PartCount { Name = name, Qty = qty });
                        }
                    }
                }

                if (string.IsNullOrEmpty(data.Name)) return null;
                return data;
            }
            catch { return null; }
        }
    }
}