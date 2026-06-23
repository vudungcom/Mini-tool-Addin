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
            string name1 = System.IO.Path.GetFileNameWithoutExtension(asm1.FullFileName);
            string name2 = System.IO.Path.GetFileNameWithoutExtension(asm2.FullFileName);

            // === 1. Build tree de quy ===
            var tree1 = BuildTree(asm1.ComponentDefinition.Occurrences, 0);
            var tree2 = BuildTree(asm2.ComponentDefinition.Occurrences, 0);

            // === 2. Flatten parts de quy ===
            var parts1 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var parts2 = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            FlattenParts(asm1.ComponentDefinition.Occurrences, parts1, 1);
            FlattenParts(asm2.ComponentDefinition.Occurrences, parts2, 1);

            // === 3. Xac dinh output path ===
            string outputDir = System.IO.Path.GetDirectoryName(asm1.FullFileName);
            string outputName = "BOMCompare_" + name1 + "_vs_" + name2 + ".xlsx";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                outputName = outputName.Replace(c.ToString(), "_");
            string outputPath = System.IO.Path.Combine(outputDir, outputName);

            // Neu khong ghi duoc vao thu muc assembly (read-only Vault), fallback ra Desktop
            if (!CanWriteToDirectory(outputDir))
            {
                outputDir = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                outputPath = System.IO.Path.Combine(outputDir, outputName);
            }

            EnsureFileNotLocked(outputPath);

            // === 4. Build data cho 2 sheet ===
            var treeRows = CompareTrees(tree1, tree2);
            var partRows = ComparePartsRows(parts1, parts2);

            // === 5. Viet file .xlsx ===
            XlsxWriter.Write(outputPath, name1, name2, treeRows, partRows);

            // === 6. Mo file luon, khong hien thong bao ===
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
                        var subAsm = (AssemblyDocument)occDoc;
                        node.Children = BuildTree(subAsm.ComponentDefinition.Occurrences, level + 1);
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
                        pr.Note1 = "Khong co trong bang 2";
                        pr.Note1Color = CellColor.Red;
                    }
                    else if (pr.Qty1 != q2)
                    {
                        pr.Note1 = "Khac so luong";
                        pr.Note1Color = CellColor.Yellow;
                    }
                    else
                    {
                        pr.Note1 = "Giong nhau";
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
                        pr.Note2 = "Khong co trong bang 1";
                        pr.Note2Color = CellColor.Red;
                    }
                    else if (pr.Qty2 != q1)
                    {
                        pr.Note2 = "Khac so luong";
                        pr.Note2Color = CellColor.Yellow;
                    }
                    else
                    {
                        pr.Note2 = "Giong nhau";
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
    }
}