using Inventor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// BOM Compare Logic - [NEW v1.6]
    /// So sanh 2 assembly:
    /// - Sheet 1 "Tree Compare": cay assembly cua 2 BOM canh nhau, indent theo level.
    ///   Vang neu node khac (DIFF), do neu chi co o 1 ben (ONLY).
    /// - Sheet 2 "Part Compare": flatten tat ca parts (IPT) de quy, dem tong quantity,
    ///   so sanh theo macro VBA format (Part/Qty/Note x2).
    ///
    /// XUAT FILE: viet thang .xlsx bang OpenXML (System.IO.Packaging) - KHONG can Excel COM.
    /// File .xlsx la 1 ZIP chua cac file XML. Cach nay khong phu thuoc Excel cai tren may,
    /// khong co late-bound COM nen khong gap loi DISP_E_MEMBERNOTFOUND.
    /// </summary>
    public class BomCompareLogic
    {
        private readonly Inventor.Application _app;

        public BomCompareLogic(Inventor.Application app)
        {
            _app = app;
        }

        // ============================================================
        // ENTRY POINT
        // ============================================================

        public void Run(AssemblyDocument asm1, AssemblyDocument asm2)
        {
            // Doc BOM data tu 2 assembly (trong RAM process hien tai)
            var data1 = ReadBomData(asm1);
            var data2 = ReadBomData(asm2);
            RunFromData(data1, data2, System.IO.Path.GetDirectoryName(asm1.FullFileName));
        }

        // ============================================================
        // [v1.7] DOC BOM DATA TU ASSEMBLY (trong RAM process hien tai)
        // Tach rieng de moi process tu doc BOM cua file no dang mo,
        // ghi ra bridge, process kia doc lai - tranh doc nham file cua process khac.
        // ============================================================

        public BomData ReadBomData(AssemblyDocument asm)
        {
            var data = new BomData();
            data.Name = System.IO.Path.GetFileNameWithoutExtension(asm.FullFileName);
            data.FullPath = asm.FullFileName;

            // Tree de quy -> luu flat list (Name, Level, IsAssembly) theo thu tu duyet
            var tree = BuildTree(asm.ComponentDefinition.Occurrences, 0);
            FlattenTreeToData(tree, data.TreeNodes);

            // Parts flatten
            var parts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            FlattenParts(asm.ComponentDefinition.Occurrences, parts, 1);
            foreach (var kv in parts)
                data.Parts.Add(new PartCount { Name = kv.Key, Qty = kv.Value });

            return data;
        }

        private void FlattenTreeToData(List<TreeNode> nodes, List<TreeNodeData> output)
        {
            foreach (var n in nodes)
            {
                output.Add(new TreeNodeData { Name = n.Name, Level = n.Level, IsAssembly = n.IsAssembly });
                FlattenTreeToData(n.Children, output);
            }
        }

        // Dung lai cay tu flat list (de so sanh)
        private List<TreeNode> RebuildTreeFromData(List<TreeNodeData> flat)
        {
            var roots = new List<TreeNode>();
            var stack = new Stack<TreeNode>();

            foreach (var nd in flat)
            {
                var node = new TreeNode { Name = nd.Name, Level = nd.Level, IsAssembly = nd.IsAssembly };

                // Pop ve dung cap cha
                while (stack.Count > 0 && stack.Peek().Level >= node.Level)
                    stack.Pop();

                if (stack.Count == 0)
                    roots.Add(node);
                else
                    stack.Peek().Children.Add(node);

                stack.Push(node);
            }
            return roots;
        }

        // ============================================================
        // [v1.7] SO SANH TU 2 BOM DATA (da doc san)
        // ============================================================

        public void RunFromData(BomData data1, BomData data2, string preferredDir)
        {
            string name1 = data1.Name;
            string name2 = data2.Name;

            // Dung lai cay tu data
            var tree1 = RebuildTreeFromData(data1.TreeNodes);
            var tree2 = RebuildTreeFromData(data2.TreeNodes);

            // Parts -> dict
            var parts1 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var parts2 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in data1.Parts) parts1[p.Name] = p.Qty;
            foreach (var p in data2.Parts) parts2[p.Name] = p.Qty;

            // Output path
            string outputDir = preferredDir;
            if (string.IsNullOrEmpty(outputDir) || !CanWriteToDirectory(outputDir))
                outputDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);

            string outputName = "BOMCompare_" + name1 + "_vs_" + name2 + ".xlsx";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                outputName = outputName.Replace(c.ToString(), "_");
            string outputPath = System.IO.Path.Combine(outputDir, outputName);

            EnsureFileNotLocked(outputPath);

            var treeRows = CompareTrees(tree1, tree2);
            var partRows = ComparePartsRows(parts1, parts2);

            XlsxWriter.Write(outputPath, name1, name2, treeRows, partRows);

            try { System.Diagnostics.Process.Start(outputPath); }
            catch { }
        }

        // ============================================================
        // PRIVATE: BUILD TREE (RECURSIVE)
        // ============================================================

        private class TreeNode
        {
            public string Name;
            public int Level;
            public bool IsAssembly;
            public List<TreeNode> Children = new List<TreeNode>();
        }

        private List<TreeNode> BuildTree(ComponentOccurrences occs, int level)
        {
            var list = new List<TreeNode>();
            foreach (ComponentOccurrence occ in occs)
            {
                try
                {
                    bool suppressed = false;
                    try { suppressed = occ.Suppressed; } catch { }
                    if (suppressed) continue;

                    Document occDoc = null;
                    try { occDoc = (Document)occ.Definition.Document; } catch { }
                    if (occDoc == null) continue;

                    string fileName = "";
                    try { fileName = System.IO.Path.GetFileNameWithoutExtension(occDoc.FullFileName); } catch { }
                    if (string.IsNullOrEmpty(fileName)) continue;

                    bool isAsm = (occDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject);

                    var node = new TreeNode { Name = fileName, Level = level, IsAssembly = isAsm };

                    if (isAsm)
                    {
                        // [FIX-TREE BOM-STRUCTURE] Ton trong BOM Structure giong Part Compare.
                        // Reference               -> bo ca cum (khong dua vao cay).
                        // Purchased / Inseparable -> hien 1 dong duy nhat, KHONG bung con (Children rong).
                        // Normal / Phantom / Varies -> de quy nhu cu.
                        BOMStructureEnum bomStruct = BOMStructureEnum.kNormalBOMStructure;
                        try { bomStruct = occ.BOMStructure; } catch { }

                        if (bomStruct == BOMStructureEnum.kReferenceBOMStructure)
                            continue;   // cum reference -> khong them node

                        if (bomStruct != BOMStructureEnum.kPurchasedBOMStructure &&
                            bomStruct != BOMStructureEnum.kInseparableBOMStructure)
                        {
                            var subAsm = (AssemblyDocument)occDoc;
                            node.Children = BuildTree(subAsm.ComponentDefinition.Occurrences, level + 1);
                        }
                        // Purchased/Inseparable: giu node, Children rong -> hien nhu 1 leaf .iam
                    }

                    list.Add(node);
                }
                catch { }
            }
            return list;
        }

        // ============================================================
        // PRIVATE: FLATTEN PARTS (RECURSIVE)
        // ============================================================

        private void FlattenParts(ComponentOccurrences occs, Dictionary<string, int> result, int multiplier)
        {
            foreach (ComponentOccurrence occ in occs)
            {
                try
                {
                    bool suppressed = false;
                    try { suppressed = occ.Suppressed; } catch { }
                    if (suppressed) continue;

                    Document occDoc = null;
                    try { occDoc = (Document)occ.Definition.Document; } catch { }
                    if (occDoc == null) continue;

                    string fileName = "";
                    try { fileName = System.IO.Path.GetFileNameWithoutExtension(occDoc.FullFileName); } catch { }
                    if (string.IsNullOrEmpty(fileName)) continue;

                    if (occDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        // [FIX-PART BOM-STRUCTURE] Doc BOM Structure cua cum truoc khi de quy.
                        // Reference               -> loai bo ca cum (giong part Reference ben duoi).
                        // Purchased / Inseparable -> coi nhu 1 leaf part (lay ten .iam), KHONG de quy.
                        //   => khac phuc loi no tung cum mua / khong tach roi thanh part con.
                        // Normal / Phantom / Varies -> de quy nhu cu (con duoc gom len).
                        BOMStructureEnum asmStruct = BOMStructureEnum.kNormalBOMStructure;
                        try { asmStruct = occ.BOMStructure; } catch { }

                        if (asmStruct == BOMStructureEnum.kReferenceBOMStructure)
                            continue;   // cum reference -> bo qua

                        if (asmStruct == BOMStructureEnum.kPurchasedBOMStructure ||
                            asmStruct == BOMStructureEnum.kInseparableBOMStructure)
                        {
                            // Ghi cum nhu 1 leaf part (dung ten .iam), khong bung con
                            if (result.ContainsKey(fileName))
                                result[fileName] += multiplier;
                            else
                                result[fileName] = multiplier;
                            continue;   // DUNG de quy
                        }

                        // Normal / Phantom / Varies -> de quy nhu cu
                        var subAsm = (AssemblyDocument)occDoc;
                        FlattenParts(subAsm.ComponentDefinition.Occurrences, result, multiplier);
                    }
                    else if (occDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                    {
                        bool isRef = false;
                        try { isRef = (occ.BOMStructure == BOMStructureEnum.kReferenceBOMStructure); } catch { }
                        if (isRef) continue;

                        if (result.ContainsKey(fileName))
                            result[fileName] += multiplier;
                        else
                            result[fileName] = multiplier;
                    }
                }
                catch { }
            }
        }

        // ============================================================
        // PRIVATE: TREE COMPARE
        // ============================================================

        private List<TreeRow> CompareTrees(List<TreeNode> t1, List<TreeNode> t2)
        {
            var rows = new List<TreeRow>();
            CompareTreesRecursive(t1, t2, rows);
            return rows;
        }

        private void CompareTreesRecursive(List<TreeNode> t1, List<TreeNode> t2, List<TreeRow> rows)
        {
            var matched2 = new HashSet<int>();

            for (int i = 0; i < t1.Count; i++)
            {
                var n1 = t1[i];
                int matchIdx = -1;
                for (int j = 0; j < t2.Count; j++)
                {
                    if (matched2.Contains(j)) continue;
                    if (string.Equals(t2[j].Name, n1.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        matchIdx = j;
                        break;
                    }
                }

                string indent = new string(' ', n1.Level * 2);
                if (matchIdx >= 0)
                {
                    var n2 = t2[matchIdx];
                    matched2.Add(matchIdx);

                    var childRows = new List<TreeRow>();
                    if (n1.IsAssembly || n2.IsAssembly)
                        CompareTreesRecursive(n1.Children, n2.Children, childRows);

                    bool hasDiffInChildren = childRows.Any(r => r.Status != "OK");
                    string status = hasDiffInChildren ? "DIFF" : "OK";

                    rows.Add(new TreeRow { Name1 = indent + n1.Name, Name2 = indent + n2.Name, Status = status });
                    rows.AddRange(childRows);
                }
                else
                {
                    rows.Add(new TreeRow { Name1 = indent + n1.Name, Name2 = "", Status = "ONLY-1" });
                    AddTreeAsOnlyOneSide(n1.Children, rows, true);
                }
            }

            for (int j = 0; j < t2.Count; j++)
            {
                if (matched2.Contains(j)) continue;
                var n2 = t2[j];
                string indent = new string(' ', n2.Level * 2);
                rows.Add(new TreeRow { Name1 = "", Name2 = indent + n2.Name, Status = "ONLY-2" });
                AddTreeAsOnlyOneSide(n2.Children, rows, false);
            }
        }

        private void AddTreeAsOnlyOneSide(List<TreeNode> nodes, List<TreeRow> rows, bool isLeft)
        {
            foreach (var n in nodes)
            {
                string indent = new string(' ', n.Level * 2);
                rows.Add(new TreeRow
                {
                    Name1 = isLeft ? indent + n.Name : "",
                    Name2 = isLeft ? "" : indent + n.Name,
                    Status = isLeft ? "ONLY-1" : "ONLY-2"
                });
                AddTreeAsOnlyOneSide(n.Children, rows, isLeft);
            }
        }

        // ============================================================
        // PRIVATE: PART COMPARE (theo macro VBA)
        // ============================================================

        private List<PartRow> ComparePartsRows(Dictionary<string, int> parts1, Dictionary<string, int> parts2)
        {
            var sorted1 = parts1.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase).ToList();
            var sorted2 = parts2.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase).ToList();

            int maxRows = Math.Max(sorted1.Count, sorted2.Count);
            var rows = new List<PartRow>();

            for (int i = 0; i < maxRows; i++)
            {
                var pr = new PartRow();

                // Cot trai (BOM 1)
                if (i < sorted1.Count)
                {
                    pr.Name1 = sorted1[i].Key;
                    pr.Qty1 = sorted1[i].Value;
                    if (!parts2.TryGetValue(pr.Name1, out int q2))
                    {
                        pr.Note1 = OpenCadDrawingAddin.LanguageManager.L("BOMCMP_NOTE_NOT_IN", "2");
                        pr.Note1Color = CellColor.Red;
                    }
                    else if (pr.Qty1 != q2)
                    {
                        pr.Note1 = OpenCadDrawingAddin.LanguageManager.L("BOMCMP_NOTE_DIFF_QTY");
                        pr.Note1Color = CellColor.Yellow;
                    }
                    else
                    {
                        pr.Note1 = OpenCadDrawingAddin.LanguageManager.L("BOMCMP_NOTE_SAME");
                        pr.Note1Color = CellColor.None;
                    }
                    pr.Has1 = true;
                }

                // Cot phai (BOM 2)
                if (i < sorted2.Count)
                {
                    pr.Name2 = sorted2[i].Key;
                    pr.Qty2 = sorted2[i].Value;
                    if (!parts1.TryGetValue(pr.Name2, out int q1))
                    {
                        pr.Note2 = OpenCadDrawingAddin.LanguageManager.L("BOMCMP_NOTE_NOT_IN", "1");
                        pr.Note2Color = CellColor.Red;
                    }
                    else if (pr.Qty2 != q1)
                    {
                        pr.Note2 = OpenCadDrawingAddin.LanguageManager.L("BOMCMP_NOTE_DIFF_QTY");
                        pr.Note2Color = CellColor.Yellow;
                    }
                    else
                    {
                        pr.Note2 = OpenCadDrawingAddin.LanguageManager.L("BOMCMP_NOTE_SAME");
                        pr.Note2Color = CellColor.None;
                    }
                    pr.Has2 = true;
                }

                rows.Add(pr);
            }

            return rows;
        }

        // ============================================================
        // PRIVATE: FILE HELPERS
        // ============================================================

        private bool CanWriteToDirectory(string dir)
        {
            try
            {
                string testFile = System.IO.Path.Combine(dir, "._ocda_write_test_" + Guid.NewGuid().ToString("N") + ".tmp");
                System.IO.File.WriteAllText(testFile, "test");
                System.IO.File.Delete(testFile);
                return true;
            }
            catch { return false; }
        }

        private void EnsureFileNotLocked(string path)
        {
            if (!System.IO.File.Exists(path)) return;
            try
            {
                using (var fs = System.IO.File.Open(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None))
                { }
            }
            catch
            {
                MessageBox.Show(
                    "File ket qua dang mo trong Excel. Hay dong file roi thu lai:\n" + path,
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                throw new OperationCanceledException();
            }
        }

        // ============================================================
        // DATA MODELS (shared voi XlsxWriter)
        // ============================================================

        public enum CellColor { None, Yellow, Red }

        public class TreeRow
        {
            public string Name1;
            public string Name2;
            public string Status;   // "OK" | "DIFF" | "ONLY-1" | "ONLY-2"
        }

        public class PartRow
        {
            public bool Has1, Has2;
            public string Name1 = "", Name2 = "";
            public int Qty1, Qty2;
            public string Note1 = "", Note2 = "";
            public CellColor Note1Color = CellColor.None;
            public CellColor Note2Color = CellColor.None;
        }

        // ============================================================
        // [v1.7] BOM DATA - serialize duoc de ghi/doc qua bridge file
        // ============================================================

        public class BomData
        {
            public string Name = "";
            public string FullPath = "";
            public List<TreeNodeData> TreeNodes = new List<TreeNodeData>();
            public List<PartCount> Parts = new List<PartCount>();
        }

        public class TreeNodeData
        {
            public string Name = "";
            public int Level;
            public bool IsAssembly;
        }

        public class PartCount
        {
            public string Name = "";
            public int Qty;
        }
    }
}