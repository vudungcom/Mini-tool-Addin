using OpenCadDrawingAddin.Logic;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace OpenCadDrawingAddin
{
    /// <summary>
    /// BOM Compare Form - [NEW v1.6]
    /// Dialog non-modal de cho user switch tab Inventor.
    /// Pick uu tien sub-assembly da selected trong browser tree.
    /// </summary>
    public class BomCompareForm : Form
    {
        private readonly Inventor.Application _app;

        private Label lblAsm1, lblAsm2;
        private Label lblPath1, lblPath2;
        private Button btnPick1, btnPick2;
        private Button btnCompare, btnCancel;

        private Inventor.AssemblyDocument _asm1;
        private Inventor.AssemblyDocument _asm2;

        public BomCompareForm(Inventor.Application app)
        {
            _app = app;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "BOM Compare";
            this.Size = new Size(600, 260);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow; // Tool window nhe, ko block
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.Font = new Font("Segoe UI", 9F);
            this.TopMost = true;  // Luon o tren de user thay
            this.ShowInTaskbar = false;

            int x = 15, y = 15;

            var lblHint = new Label
            {
                Text = "Cach dung:\n1. Click vao IAM (hoac sub-asm trong browser) ben Inventor\n2. Sau do click 'Pick' tuong ung",
                Location = new Point(x, y),
                Size = new Size(560, 48),
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8.5F, FontStyle.Italic)
            };
            this.Controls.Add(lblHint);
            y += 56;

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
                Size = new Size(390, 20),
                ForeColor = Color.Gray,
                AutoEllipsis = true
            };
            this.Controls.Add(lblPath1);
            y += 40;

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
                Size = new Size(390, 20),
                ForeColor = Color.Gray,
                AutoEllipsis = true
            };
            this.Controls.Add(lblPath2);
            y += 50;

            btnCompare = new Button
            {
                Text = "Compare",
                Location = new Point(x + 280, y),
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
                Location = new Point(x + 420, y),
                Size = new Size(130, 32)
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);
            this.CancelButton = btnCancel;
        }

        private void BtnPick1_Click(object sender, EventArgs e)
        {
            var asm = GetSelectedAssembly(out string errMsg);
            if (asm == null)
            {
                MessageBox.Show(errMsg ?? "Khong tim thay Assembly.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _asm1 = asm;
            lblPath1.Text = Path.GetFileName(asm.FullFileName);
            lblPath1.ForeColor = Color.DarkGreen;
            lblPath1.Font = new Font(lblPath1.Font, FontStyle.Bold);
            UpdateCompareEnabled();
        }

        private void BtnPick2_Click(object sender, EventArgs e)
        {
            var asm = GetSelectedAssembly(out string errMsg);
            if (asm == null)
            {
                MessageBox.Show(errMsg ?? "Khong tim thay Assembly.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_asm1 != null && string.Equals(asm.FullFileName, _asm1.FullFileName, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Trung file voi Assembly 1:\n" + asm.FullFileName +
                    "\n\nVui long chon file khac.",
                    "BOM Compare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _asm2 = asm;
            lblPath2.Text = Path.GetFileName(asm.FullFileName);
            lblPath2.ForeColor = Color.DarkGreen;
            lblPath2.Font = new Font(lblPath2.Font, FontStyle.Bold);
            UpdateCompareEnabled();
        }

        /// <summary>
        /// Lay assembly. UU TIEN selection trong browser:
        /// 1. Neu user da chon 1+ occurrence trong browser tree -> lay sub-asm cua occurrence dau tien
        /// 2. Neu khong chon gi -> lay ActiveDocument neu la IAM (root assembly)
        /// </summary>
        private Inventor.AssemblyDocument GetSelectedAssembly(out string errMsg)
        {
            errMsg = null;
            try
            {
                var doc = _app.ActiveDocument;
                if (doc == null)
                {
                    errMsg = "Khong co document nao dang mo trong Inventor.";
                    return null;
                }

                // UU TIEN 1: selection trong browser tree
                if (doc.SelectSet != null && doc.SelectSet.Count >= 1)
                {
                    for (int i = 1; i <= doc.SelectSet.Count; i++)
                    {
                        var sel = doc.SelectSet[i];

                        // Browser tree thuong chon dang ComponentOccurrence
                        if (sel is Inventor.ComponentOccurrence occ)
                        {
                            try
                            {
                                var occDoc = (Inventor.Document)occ.Definition.Document;
                                if (occDoc.DocumentType == Inventor.DocumentTypeEnum.kAssemblyDocumentObject)
                                    return (Inventor.AssemblyDocument)occDoc;
                                else
                                {
                                    errMsg = "Occurrence da chon khong phai Assembly (la Part).";
                                    return null;
                                }
                            }
                            catch { }
                        }

                        // Truong hop chon AssemblyComponentDefinition (root cua sub-asm trong browser)
                        if (sel is Inventor.AssemblyComponentDefinition asmDef)
                        {
                            try
                            {
                                var asmDoc = (Inventor.AssemblyDocument)asmDef.Document;
                                return asmDoc;
                            }
                            catch { }
                        }
                    }
                }

                // UU TIEN 2: ActiveDocument la IAM
                if (doc.DocumentType == Inventor.DocumentTypeEnum.kAssemblyDocumentObject)
                    return (Inventor.AssemblyDocument)doc;

                errMsg = "Active document khong phai Assembly (.iam).\nSwitch sang tab IAM hoac click vao sub-asm trong browser.";
                return null;
            }
            catch (Exception ex)
            {
                errMsg = "Loi khi lay assembly: " + ex.Message;
                return null;
            }
        }

        private void UpdateCompareEnabled()
        {
            btnCompare.Enabled = (_asm1 != null && _asm2 != null);
        }

        private void BtnCompare_Click(object sender, EventArgs e)
        {
            if (_asm1 == null || _asm2 == null) return;

            btnCompare.Enabled = false;
            btnCompare.Text = "Comparing...";
            this.Cursor = Cursors.WaitCursor;
            try
            {
                var logic = new BomCompareLogic(_app);
                logic.Run(_asm1, _asm2);
                this.Close();
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                // Ghi log day du ra file de debug
                string logPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "BOMCompare_Error_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

                var sb = new System.Text.StringBuilder();
                sb.AppendLine("=== BOM COMPARE ERROR LOG ===");
                sb.AppendLine("Time: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                sb.AppendLine("Assembly 1: " + (_asm1?.FullFileName ?? "(null)"));
                sb.AppendLine("Assembly 2: " + (_asm2?.FullFileName ?? "(null)"));
                sb.AppendLine();

                int level = 0;
                var cur = ex;
                while (cur != null)
                {
                    sb.AppendLine("--- Exception Level " + level + " ---");
                    sb.AppendLine("Type:    " + cur.GetType().FullName);
                    sb.AppendLine("Message: " + cur.Message);
                    sb.AppendLine("Source:  " + cur.Source);
                    sb.AppendLine("HResult: 0x" + cur.HResult.ToString("X8"));
                    sb.AppendLine("StackTrace:");
                    sb.AppendLine(cur.StackTrace ?? "(no stack)");
                    sb.AppendLine();

                    // Neu la TargetInvocationException -> dig sau hon
                    cur = cur.InnerException;
                    level++;
                    if (level > 10) break;
                }

                try { System.IO.File.WriteAllText(logPath, sb.ToString()); } catch { }

                // Hien dialog gon
                var deepEx = ex;
                while (deepEx.InnerException != null) deepEx = deepEx.InnerException;

                string msg = "Loi: " + deepEx.Message + "\n\n" +
                             "Da ghi log day du tai:\n" + logPath + "\n\n" +
                             "Mo file de gui debug?";

                var result = MessageBox.Show(msg, "BOM Compare Error",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                if (result == DialogResult.Yes)
                {
                    try { System.Diagnostics.Process.Start(logPath); } catch { }
                }
            }
            finally
            {
                this.Cursor = Cursors.Default;
                btnCompare.Enabled = true;
                btnCompare.Text = "Compare";
            }
        }
    }
}