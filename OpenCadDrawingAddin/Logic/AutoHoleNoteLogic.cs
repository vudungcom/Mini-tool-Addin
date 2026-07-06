using System;
using System.Collections.Generic;
using System.Reflection;
using Inventor;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// Auto Hole Note Logic - [NEW v1.5]
    /// Quet tat ca lo tron trong DrawingView tu file STEP/SolidEdge,
    /// phan loai M-tap (theo Solid Edge pitch diameter) hoac Phi tron,
    /// gom cum BFS, dat GeneralNote sat lo dai dien, tranh de bang bbox.
    /// Toan bo wrap trong Inventor Transaction -> Ctrl+Z 1 lan xoa het.
    /// </summary>
    public class AutoHoleNoteLogic
    {
        private readonly Inventor.Application _app;

        // === Bang ren Solid Edge convention (pitch diameter, mm) ===
        private static readonly (double Dia, string Name)[] TapTable =
        {
            (1.567,  "M2"),
            (2.013,  "M2.5"),
            (2.459,  "M3"),
            (3.242,  "M4"),
            (4.134,  "M5"),
            (4.917,  "M6"),
            (6.647,  "M8"),
            (8.376,  "M10"),
            (10.106, "M12"),
            (13.835, "M16"),
            (17.294, "M20"),
        };

        public AutoHoleNoteLogic(Inventor.Application app)
        {
            _app = app;
        }

        // ============================================================
        // [AUTO-SIZE-HELPER] Tinh textHeight & clusterRadius theo view scale
        // Base scale co dinh = 1:4 (0.25). Base value do CALLER truyen vao
        // (chinh la gia tri user set o UI - ho chinh 3.0 -> base 3.0, ho chinh 4.0 -> base 4.0).
        // Cong thuc: value = baseValue * (viewScale / 0.25).
        // Lam tron: text -> 0.5mm gan nhat, cluster -> 5mm gan nhat.
        // Clamp: text [0.5, 5.0], cluster [5, 200] (khop range cua NumericUpDown).
        // ============================================================
        public static void ComputeAutoSize(
            double viewScale,
            double baseTextMm,
            double baseClusterMm,
            out double textHeightMm,
            out double clusterRadiusMm)
        {
            const double BASE_SCALE = 0.25;   // 1:4

            if (viewScale <= 0) viewScale = BASE_SCALE;
            if (baseTextMm <= 0) baseTextMm = 3.0;
            if (baseClusterMm <= 0) baseClusterMm = 30.0;

            double factor = viewScale / BASE_SCALE;

            // Text: lam tron 0.5mm gan nhat, clamp [0.5, 5.0]
            double rawText = baseTextMm * factor;
            textHeightMm = Math.Round(rawText * 2.0) / 2.0;
            if (textHeightMm < 0.5) textHeightMm = 0.5;
            if (textHeightMm > 5.0) textHeightMm = 5.0;

            // Cluster: lam tron 5mm gan nhat, clamp [5, 200]
            double rawCluster = baseClusterMm * factor;
            clusterRadiusMm = Math.Round(rawCluster / 5.0) * 5.0;
            if (clusterRadiusMm < 5.0) clusterRadiusMm = 5.0;
            if (clusterRadiusMm > 200.0) clusterRadiusMm = 200.0;
        }

        /// <summary>
        /// Entry point: quet view, gom cum, tao note, wrap Transaction.
        /// </summary>
        public void Run(DrawingView view, double textHeightMm, double clusterRadiusMm, double tolTap)
        {
            // Fix CS1503: cast qua Document truoc, sau do sang DrawingDocument
            var doc = _app.ActiveDocument as DrawingDocument;
            if (doc == null) return;

            var sheet = doc.ActiveSheet;
            var tg = _app.TransientGeometry;

            const double TolRound = 0.05;
            double gapCm = 0.05;
            double fontSizeCm = textHeightMm / 10.0;
            double charW = fontSizeCm * 0.65;
            double textH = fontSizeCm;
            string fontStr = fontSizeCm.ToString("0.000",
                System.Globalization.CultureInfo.InvariantCulture);

            // === 1. Quet lo ===
            var holes = ScanHoles(view, tolTap, TolRound);
            if (holes.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Khong tim thay lo nao trong view.\n" +
                    "Dam bao da chon dung view va view co chua lo tron.",
                    "Auto Hole Note",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            // === 2. Sort deterministic: Label ASC, Y DESC, X ASC ===
            holes.Sort((a, b) =>
            {
                int c = string.Compare(a.Label, b.Label, StringComparison.Ordinal);
                if (c != 0) return c;
                c = b.Cy.CompareTo(a.Cy);
                if (c != 0) return c;
                return a.Cx.CompareTo(b.Cx);
            });

            // === 3. Cluster BFS ===
            double clusterRadiusCm = clusterRadiusMm / 10.0 * view.Scale;
            var clusters = BuildClusters(holes, clusterRadiusCm);

            // === 4. Tao note trong Transaction ===
            Transaction trans = null;
            try
            {
                trans = _app.TransactionManager.StartTransaction(
                    (Inventor._Document)doc,
                    "Auto Hole Note");

                // [ATTACH-VIEW] Tao overlay sketch cua view thay vi note tren sheet.
                // Overlay sketch nam trong he toa do MODEL cua view (khong phai sheet):
                //   1 unit sketch = 1 unit model = (1 / view.Scale) unit sheet.
                // Vi vay khi convert diem tu sheet -> sketch, phai:
                //   sketchPt = (sheetPt - view.Position) / view.Scale
                // Va font size trong sketch cung phai nhan (1 / view.Scale) de ra dung mm tren giay.
                // Neu view khong support overlay, fallback ve GeneralNotes tren sheet (khong dinh view).
                DrawingSketch sketch = null;
                double viewOx = 0, viewOy = 0;
                double invScale = 1.0;  // 1 / view.Scale, dung khi co sketch
                try
                {
                    sketch = view.Sketches.Add();
                    try { sketch.Name = "OCDA_HoleNotes"; } catch { }
                    sketch.Edit();
                    viewOx = view.Position.X;
                    viewOy = view.Position.Y;
                    invScale = (view.Scale > 0) ? (1.0 / view.Scale) : 1.0;
                }
                catch
                {
                    sketch = null;
                }

                var placedBbox = new List<(double x1, double y1, double x2, double y2)>();
                int noteCount = 0;

                foreach (var cluster in clusters)
                {
                    if (cluster.Count == 0) continue;

                    string label = holes[cluster[0]].Label;
                    string plainText = cluster.Count == 1
                        ? label
                        : $"{cluster.Count}x {label}";

                    string formatted =
                        $"<StyleOverride FontSize=\"{fontStr}\">{plainText}</StyleOverride>";

                    double textW = plainText.Length * charW;

                    // Centroid
                    double sumX = 0, sumY = 0;
                    foreach (int idx in cluster) { sumX += holes[idx].Cx; sumY += holes[idx].Cy; }
                    double ctrX = sumX / cluster.Count;
                    double ctrY = sumY / cluster.Count;

                    // Lo dai dien = gan centroid nhat
                    int repIdx = cluster[0];
                    double minD = double.MaxValue;
                    foreach (int idx in cluster)
                    {
                        double d = Dist(holes[idx].Cx - ctrX, holes[idx].Cy - ctrY);
                        if (d < minD) { minD = d; repIdx = idx; }
                    }

                    var rep = holes[repIdx];
                    double distFromCenter = rep.R + gapCm + Math.Max(textW, textH) / 2.0;

                    double anchorX = rep.Cx + rep.R + gapCm;
                    double anchorY = rep.Cy - textH / 2.0;
                    double bxMin = anchorX, byMin = anchorY;
                    double bxMax = anchorX + textW, byMax = anchorY + textH;
                    bool found = false;

                    for (int exp = 0; exp <= 10 && !found; exp++)
                    {
                        double extra = exp * (rep.R + gapCm) * 0.5;
                        for (int di = 0; di < 8 && !found; di++)
                        {
                            double angle = di * 45.0 * Math.PI / 180.0;
                            double tcx = rep.Cx + Math.Cos(angle) * (distFromCenter + extra);
                            double tcy = rep.Cy + Math.Sin(angle) * (distFromCenter + extra);

                            double nx1 = tcx - textW / 2.0;
                            double ny1 = tcy - textH / 2.0;
                            double nx2 = tcx + textW / 2.0;
                            double ny2 = tcy + textH / 2.0;

                            bool hit = false;
                            foreach (var bb in placedBbox)
                            {
                                if (!(nx2 < bb.x1 || nx1 > bb.x2 || ny2 < bb.y1 || ny1 > bb.y2))
                                { hit = true; break; }
                            }

                            if (!hit)
                            {
                                anchorX = nx1; anchorY = ny1;
                                bxMin = nx1; byMin = ny1;
                                bxMax = nx2; byMax = ny2;
                                found = true;
                            }
                        }
                    }

                    try
                    {
                        // [ATTACH-VIEW] Ghi text vao overlay sketch neu co, khong thi
                        // fallback ve GeneralNotes tren sheet nhu code cu.
                        if (sketch != null)
                        {
                            // [ATTACH-VIEW] He toa do overlay sketch la HE LAI (xac nhan qua 2 lan test):
                            //   - Toa do diem  : MODEL-space -> phai nhan (1/view.Scale).
                            //   - Font size    : don vi GIAY -> giu nguyen fontStr goc, KHONG nhan.
                            double sx = (anchorX - viewOx) * invScale;
                            double sy = (anchorY - viewOy) * invScale;
                            var ptView = tg.CreatePoint2d(sx, sy);

                            sketch.TextBoxes.AddFitted(ptView, formatted);
                        }
                        else
                        {
                            var pt = tg.CreatePoint2d(anchorX, anchorY);
                            sheet.DrawingNotes.GeneralNotes.AddFitted(pt, formatted);
                        }
                        placedBbox.Add((bxMin, byMin, bxMax, byMax));
                        noteCount++;
                    }
                    catch { }
                }

                // [ATTACH-VIEW] Thoat sketch edit truoc khi End Transaction
                if (sketch != null)
                {
                    try { sketch.ExitEdit(); } catch { }
                }

                trans.End();

                System.Windows.Forms.MessageBox.Show(
                    $"Da them {noteCount} note.\n" +
                    $"Tong: {holes.Count} lo, {clusters.Count} cum.\n" +
                    (sketch != null
                        ? "Text da gan vao view (dich view -> text di theo).\n"
                        : "Note tren sheet (fallback - khong dinh view).\n") +
                    "Ctrl+Z de xoa toan bo va chay lai.",
                    "Auto Hole Note",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                try { trans?.Abort(); } catch { }
                System.Windows.Forms.MessageBox.Show(
                    "Loi: " + ex.Message,
                    "Auto Hole Note",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // ============================================================
        // PRIVATE: SCAN HOLES
        // ============================================================

        private struct HoleInfo
        {
            public double Cx, Cy, R;
            public string Label;
        }

        /// <summary>
        /// Quet DrawingCurves, lay full-circle (ProjectedCurveType == kCircleCurve2d).
        /// Dung Reflection thay cho dynamic/Circle2d de tranh CS0656 + CS0246.
        /// </summary>
        private List<HoleInfo> ScanHoles(DrawingView view, double tolTap, double tolRound)
        {
            var list = new List<HoleInfo>();

            foreach (DrawingCurve curve in view.DrawingCurves)
            {
                if (curve.ProjectedCurveType != Curve2dTypeEnum.kCircleCurve2d) continue;
                if (curve.Segments.Count == 0) continue;

                double paperR, cx, cy;
                try
                {
                    // Dung Reflection thay dynamic (tranh yeu cau Microsoft.CSharp)
                    object geom = curve.Segments[1].Geometry;
                    Type t = geom.GetType();

                    object radiusVal = t.InvokeMember("Radius",
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, geom, null);
                    paperR = Convert.ToDouble(radiusVal);

                    object center = t.InvokeMember("Center",
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, geom, null);
                    Type tc = center.GetType();

                    object xVal = tc.InvokeMember("X",
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, center, null);
                    object yVal = tc.InvokeMember("Y",
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null, center, null);

                    cx = Convert.ToDouble(xVal);
                    cy = Convert.ToDouble(yVal);
                }
                catch { continue; }

                double diaMm = paperR * 2.0 * 10.0 / view.Scale;
                if (diaMm < 1.0 || diaMm > 30.0) continue;

                // Dedupe theo (tam, ban kinh)
                bool dup = false;
                foreach (var h in list)
                {
                    if (Math.Abs(h.Cx - cx) < 0.01 &&
                        Math.Abs(h.Cy - cy) < 0.01 &&
                        Math.Abs(h.R - paperR) < 0.01)
                    { dup = true; break; }
                }
                if (dup) continue;

                list.Add(new HoleInfo
                {
                    Cx = cx,
                    Cy = cy,
                    R = paperR,
                    Label = ClassifyHole(diaMm, tolTap, tolRound)
                });
            }

            return list;
        }

        /// <summary>
        /// Phan loai lo:
        /// - Duong kinh nguyen (+-tolRound) -> Phi (lo tron)
        /// - Duong kinh le + match bang tap -> Mx
        /// - Khong match -> Phi le
        /// </summary>
        private static string ClassifyHole(double diaMm, double tolTap, double tolRound)
        {
            double rounded = Math.Round(diaMm);
            if (Math.Abs(diaMm - rounded) < tolRound)
                return "\u00D8" + ((int)rounded).ToString();

            double bestDiff = double.MaxValue;
            string bestName = null;
            foreach (var (dia, name) in TapTable)
            {
                double d = Math.Abs(diaMm - dia);
                if (d <= tolTap && d < bestDiff) { bestDiff = d; bestName = name; }
            }
            if (bestName != null) return bestName;

            return "\u00D8" + diaMm.ToString("0.0",
                System.Globalization.CultureInfo.InvariantCulture);
        }

        // ============================================================
        // PRIVATE: CLUSTER BFS
        // ============================================================

        private static List<List<int>> BuildClusters(List<HoleInfo> holes, double radiusCm)
        {
            int n = holes.Count;
            var id = new int[n];
            for (int i = 0; i < n; i++) id[i] = -1;
            int nextId = 0;

            for (int i = 0; i < n; i++)
            {
                if (id[i] != -1) continue;
                id[i] = nextId;

                var queue = new Queue<int>();
                queue.Enqueue(i);
                while (queue.Count > 0)
                {
                    int cur = queue.Dequeue();
                    for (int j = 0; j < n; j++)
                    {
                        if (id[j] != -1) continue;
                        if (holes[j].Label != holes[cur].Label) continue;
                        if (Dist(holes[j].Cx - holes[cur].Cx,
                                 holes[j].Cy - holes[cur].Cy) <= radiusCm)
                        { id[j] = nextId; queue.Enqueue(j); }
                    }
                }
                nextId++;
            }

            var clusters = new List<List<int>>();
            for (int c = 0; c < nextId; c++) clusters.Add(new List<int>());
            for (int i = 0; i < n; i++) clusters[id[i]].Add(i);
            return clusters;
        }

        private static double Dist(double dx, double dy) => Math.Sqrt(dx * dx + dy * dy);
    }
}