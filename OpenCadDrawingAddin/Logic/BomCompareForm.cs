using OpenCadDrawingAddin.Logic;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// BOM Compare Form - [NEW v1.7 - kien truc BOM data]
    ///
    /// Moi process khi pick assembly se TU DOC BOM cua file no dang mo (trong RAM cua no)
    /// roi push BOM DATA (khong phai path) vao bridge. Process kia doc lai data do.
    /// Nho vay tranh duoc loi 2 process doc cung file ra ket qua khac nhau.
    ///
    /// - Pick 1 / Pick 2: doc BOM assembly dang chon o process hien tai -> push bridge.
    /// - Auto-fill: Timer 3s + Activated doc bridge, dien BOM data tu cua so khac.
    /// - Compare: lay 2 BOM data (da doc san) -> so sanh -> xuat Excel.
    /// </summary>
    public class BomCompareForm : Form
    {
        private readonly Inventor.Application _app;

        private Label lblHint;
        private Label lblAsm1, lblAsm2;
        private Label lblPath1, lblPath2;
        private Button btnPick1, btnPick2;
        private Button btnCompare, btnCancel;
        private System.Windows.Forms.Timer _pollTimer;

        // BOM data da doc cho 2 slot (thay vi AssemblyDocument)
        private BomCompareLogic.BomData _data1;
        private BomCompareLogic.BomData _data2;

        public BomCompareForm(Inventor.Application app)
        {
            _app = app;
            InitializeComponent();

            this.Load += (s, e) => SyncFromBridge();
            this.Activated += (s, e) => SyncFromBridge();

            _pollTimer = new System.Windows.Forms.Timer { Interval = 3000 };
            _pollTimer.Tick += (s, e) => SyncFromBridge();
            _pollTimer.Start();

            this.FormClosed += (s, e) =>
            {
                try { _pollTimer.Stop(); _pollTimer.Dispose(); } catch { }
                try { BomCompareBridge.Clear(); } catch { }
            };
        }

        private void InitializeComponent()
        {
            this.Text = "BOM Compare";
            this.Size = new Size(620, 280);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.Font = new Font("Segoe UI", 9F);
            this.TopMost = true;
            this.ShowInTaskbar = false;

            int x = 15, y = 15;

            lblHint = new Label
            {
                Text = "Cung 1 cua so: chon IAM/sub-asm roi bam Pick.\n" +
                       "Khac cua so: chon IAM o cua so kia roi bam nut 'BOM Cmp' - o nay se tu dien.",
                Location = new Point(x, y),
                Size = new Size(585, 36),
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic)
            };
            this.Controls.Add(lblHint);
            y += 46;

            lblAsm1 = new Label
            {
                Text = "Assembly 1:",
                Location = new Point(x, y + 4),
                Size = new Size(85, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            this.Controls.Add(lblAsm1);

            btnPick1 = new Button
            {
                Text = "▶ Pick 1",
                Location = new Point(x + 90, y),
                Size = new Size(80, 26),
                BackColor = Color.LightYellow
            };
            btnPick1.Click += BtnPick1_Click;
            this.Controls.Add(btnPick1);

            lblPath1 = new Label
            {
                Text = "(chua chon)",
                Location = new Point(x + 180, y + 4),
                Size = new Size(410, 20),
                ForeColor = Color.Gray,
                AutoEllipsis = true
            };
            this.Controls.Add(lblPath1);
            y += 42;

            lblAsm2 = new Label
            {
                Text = "Assembly 2:",
                Location = new Point(x, y + 4),
                Size = new Size(85, 20),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold)
            };
            this.Controls.Add(lblAsm2);

            btnPick2 = new Button
            {
                Text = "▶ Pick 2",
                Location = new Point(x + 90, y),
                Size = new Size(80, 26),
                BackColor = Color.LightYellow
            };
            btnPick2.Click += BtnPick2_Click;
            this.Controls.Add(btnPick2);

            lblPath2 = new Label
            {
                Text = "(chua chon)",
                Location = new Point(x + 180, y + 4),
                Size = new Size(410, 20),
                ForeColor = Color.Gray,
                AutoEllipsis = true
            };
            this.Controls.Add(lblPath2);
            y += 50;

            btnCompare = new Button
            {
                Text = "Compare",
                Location = new Point(x + 300, y),
                Size = new Size(130, 32),
                BackColor = Color.LightGreen,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Enabled = false
            };
            btnCompare.Click += BtnCompare_Click;
            this.Controls.Add(btnCompare);

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(x + 440, y),
                Size = new Size(130, 32)
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);
            this.CancelButton = btnCancel;
        }

        // ============================================================
        // SYNC TU BRIDGE
        // ============================================================

        /// <summary>
        /// Doc cac BOM data trong bridge, dien data tu cua so khac vao o trong.
        /// Bo qua data co cung path voi cai da co trong form.
        /// </summary>
        private void SyncFromBridge()
        {
            try
            {
                var all = BomCompareBridge.ReadAll();
                if (all.Count == 0) return;

                foreach (var data in all)
                {
                    if (DataPathEquals(data, _data1) || DataPathEquals(data, _data2))
                        continue;

                    if (_data1 == null)
                        SetSlotFromBridge(1, data);
                    else if (_data2 == null)
                        SetSlotFromBridge(2, data);
                }

                UpdateCompareEnabled();
            }
            catch { }
        }

        private void SetSlotFromBridge(int slot, BomCompareLogic.BomData data)
        {
            string display = "\u21BB " + Path.GetFileName(data.FullPath) + "  (tu cua so khac)";

            if (slot == 1)
            {
                _data1 = data;
                lblPath1.Text = display;
                lblPath1.ForeColor = Color.MediumBlue;
                lblPath1.Font = new Font(lblPath1.Font, FontStyle.Bold);
            }
            else
            {
                _data2 = data;
                lblPath2.Text = display;
                lblPath2.ForeColor = Color.MediumBlue;
                lblPath2.Font = new Font(lblPath2.Font, FontStyle.Bold);
            }
        }

        // ============================================================
        // PICK THU CONG (doc BOM ngay tai process hien tai)
        // ============================================================

        private void BtnPick1_Click(object sender, EventArgs e)
        {
            var (data, err) = ReadSelectedBom();
            if (data == null)
            {
                MessageBox.Show(err ?? "Khong doc duoc BOM.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (DataPathEquals(data, _data2))
            {
                MessageBox.Show("Trung voi Assembly 2. Chon file khac.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _data1 = data;
            lblPath1.Text = Path.GetFileName(data.FullPath);
            lblPath1.ForeColor = Color.DarkGreen;
            lblPath1.Font = new Font(lblPath1.Font, FontStyle.Bold);

            // Push BOM DATA (da doc) vao bridge -> cua so kia doc lai
            BomCompareBridge.Push(data);

            UpdateCompareEnabled();
        }

        private void BtnPick2_Click(object sender, EventArgs e)
        {
            var (data, err) = ReadSelectedBom();
            if (data == null)
            {
                MessageBox.Show(err ?? "Khong doc duoc BOM.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (DataPathEquals(data, _data1))
            {
                MessageBox.Show("Trung voi Assembly 1. Chon file khac.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _data2 = data;
            lblPath2.Text = Path.GetFileName(data.FullPath);
            lblPath2.ForeColor = Color.DarkGreen;
            lblPath2.Font = new Font(lblPath2.Font, FontStyle.Bold);

            BomCompareBridge.Push(data);

            UpdateCompareEnabled();
        }

        /// <summary>
        /// Doc BOM cua assembly dang chon o process HIEN TAI (trong RAM cua no).
        /// Uu tien selection trong browser (sub-asm), fallback ActiveDocument.
        /// </summary>
        private (BomCompareLogic.BomData data, string err) ReadSelectedBom()
        {
            try
            {
                var asm = GetSelectedAssembly(out string err);
                if (asm == null) return (null, err);

                var logic = new BomCompareLogic(_app);
                var data = logic.ReadBomData(asm);
                return (data, null);
            }
            catch (Exception ex)
            {
                return (null, "Loi khi doc BOM: " + ex.Message);
            }
        }

        private Inventor.AssemblyDocument GetSelectedAssembly(out string err)
        {
            err = null;
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null)
                {
                    err = "Khong co document nao dang mo trong cua so nay.";
                    return null;
                }

                // UU TIEN: selection trong browser
                if (doc.SelectSet != null && doc.SelectSet.Count >= 1)
                {
                    for (int i = 1; i <= doc.SelectSet.Count; i++)
                    {
                        var sel = doc.SelectSet[i];
                        if (sel is Inventor.ComponentOccurrence occ)
                        {
                            try
                            {
                                var occDoc = (Inventor.Document)occ.Definition.Document;
                                if (occDoc.DocumentType == Inventor.DocumentTypeEnum.kAssemblyDocumentObject)
                                    return (Inventor.AssemblyDocument)occDoc;
                                err = "Occurrence da chon la Part, khong phai Assembly.";
                                return null;
                            }
                            catch { }
                        }
                        if (sel is Inventor.AssemblyComponentDefinition asmDef)
                        {
                            try { return (Inventor.AssemblyDocument)asmDef.Document; }
                            catch { }
                        }
                    }
                }

                // FALLBACK: ActiveDocument la IAM
                if (doc.DocumentType == Inventor.DocumentTypeEnum.kAssemblyDocumentObject)
                    return (Inventor.AssemblyDocument)doc;

                err = "Active document khong phai Assembly (.iam).\n" +
                      "Chon tab IAM hoac click sub-asm trong browser.";
                return null;
            }
            catch (Exception ex)
            {
                err = "Loi khi lay assembly: " + ex.Message;
                return null;
            }
        }

        // ============================================================
        // COMPARE
        // ============================================================

        private void UpdateCompareEnabled()
        {
            btnCompare.Enabled = (_data1 != null && _data2 != null);
        }

        private void BtnCompare_Click(object sender, EventArgs e)
        {
            if (_data1 == null || _data2 == null) return;

            btnCompare.Enabled = false;
            btnCompare.Text = "Comparing...";
            this.Cursor = Cursors.WaitCursor;
            try
            {
                // Thu muc luu: theo data1, fallback Desktop neu khong ghi duoc
                string preferredDir = "";
                try { preferredDir = Path.GetDirectoryName(_data1.FullPath); } catch { }

                var logic = new BomCompareLogic(_app);
                logic.RunFromData(_data1, _data2, preferredDir);

                BomCompareBridge.Clear();
                this.Close();
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                LogError(ex);
            }
            finally
            {
                this.Cursor = Cursors.Default;
                btnCompare.Enabled = true;
                btnCompare.Text = "Compare";
            }
        }

        // ============================================================
        // HELPERS
        // ============================================================

        private static bool DataPathEquals(BomCompareLogic.BomData a, BomCompareLogic.BomData b)
        {
            if (a == null || b == null) return false;
            if (string.IsNullOrEmpty(a.FullPath) || string.IsNullOrEmpty(b.FullPath)) return false;
            return string.Equals(a.FullPath, b.FullPath, StringComparison.OrdinalIgnoreCase);
        }

        private void LogError(Exception ex)
        {
            string logPath = Path.Combine(
                Path.GetTempPath(),
                "BOMCompare_Error_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== BOM COMPARE ERROR ===");
            sb.AppendLine("Time: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("Data1: " + (_data1?.FullPath ?? "(null)"));
            sb.AppendLine("Data2: " + (_data2?.FullPath ?? "(null)"));
            sb.AppendLine();

            int level = 0;
            var cur = ex;
            while (cur != null && level < 10)
            {
                sb.AppendLine("--- Level " + level + " ---");
                sb.AppendLine("Type: " + cur.GetType().FullName);
                sb.AppendLine("Message: " + cur.Message);
                sb.AppendLine("HResult: 0x" + cur.HResult.ToString("X8"));
                sb.AppendLine("Stack:");
                sb.AppendLine(cur.StackTrace ?? "(none)");
                sb.AppendLine();
                cur = cur.InnerException;
                level++;
            }

            try { File.WriteAllText(logPath, sb.ToString()); } catch { }

            var deep = ex;
            while (deep.InnerException != null) deep = deep.InnerException;

            var result = MessageBox.Show(
                "Loi: " + deep.Message + "\n\nLog: " + logPath + "\n\nMo log?",
                "BOM Compare Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
            if (result == DialogResult.Yes)
            {
                try { System.Diagnostics.Process.Start(logPath); } catch { }
            }
        }
    }
}