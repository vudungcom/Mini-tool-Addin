using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Text;

namespace OpenCadDrawingAddin.Logic
{
    /// <summary>
    /// XlsxWriter - [NEW v1.6]
    /// Viet file .xlsx thuan OpenXML bang System.IO.Packaging.
    /// KHONG can Microsoft Excel cai tren may, KHONG dung COM.
    ///
    /// File .xlsx la 1 package ZIP gom:
    ///   [Content_Types].xml          - khai bao MIME types
    ///   _rels/.rels                  - root relationship -> workbook
    ///   xl/workbook.xml              - danh sach sheets
    ///   xl/_rels/workbook.xml.rels   - workbook -> worksheets + styles
    ///   xl/styles.xml                - dinh nghia fill mau + font
    ///   xl/worksheets/sheet1.xml     - du lieu sheet 1
    ///   xl/worksheets/sheet2.xml     - du lieu sheet 2
    ///
    /// Style index trong styles.xml:
    ///   s=0: default
    ///   s=1: bold (header)
    ///   s=2: fill vang
    ///   s=3: fill do
    /// </summary>
    public static class XlsxWriter
    {
        // Content type constants
        private const string CT_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        private const string CT_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

        private const string REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        private const string REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

        public static void Write(
            string path,
            string name1, string name2,
            List<BomCompareLogic.TreeRow> treeRows,
            List<BomCompareLogic.PartRow> partRows)
        {
            // Xoa file cu neu ton tai
            if (File.Exists(path)) File.Delete(path);

            using (var pkg = Package.Open(path, FileMode.Create, FileAccess.ReadWrite))
            {
                // --- workbook.xml ---
                var workbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
                var workbookPart = pkg.CreatePart(workbookUri, CT_WORKBOOK);
                WriteXml(workbookPart, BuildWorkbookXml());
                pkg.CreateRelationship(workbookUri, TargetMode.Internal, REL_OFFICE_DOC, "rId1");

                // --- styles.xml ---
                var stylesUri = new Uri("/xl/styles.xml", UriKind.Relative);
                var stylesPart = pkg.CreatePart(stylesUri, CT_STYLES);
                WriteXml(stylesPart, BuildStylesXml());
                workbookPart.CreateRelationship(
                    new Uri("/xl/styles.xml", UriKind.Relative),
                    TargetMode.Internal, REL_STYLES, "rId1");

                // --- sheet1.xml (Tree Compare) ---
                var sheet1Uri = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
                var sheet1Part = pkg.CreatePart(sheet1Uri, CT_WORKSHEET);
                WriteXml(sheet1Part, BuildTreeSheetXml(name1, name2, treeRows));
                workbookPart.CreateRelationship(
                    new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative),
                    TargetMode.Internal, REL_WORKSHEET, "rId2");

                // --- sheet2.xml (Part Compare) ---
                var sheet2Uri = new Uri("/xl/worksheets/sheet2.xml", UriKind.Relative);
                var sheet2Part = pkg.CreatePart(sheet2Uri, CT_WORKSHEET);
                WriteXml(sheet2Part, BuildPartSheetXml(name1, name2, partRows));
                workbookPart.CreateRelationship(
                    new Uri("/xl/worksheets/sheet2.xml", UriKind.Relative),
                    TargetMode.Internal, REL_WORKSHEET, "rId3");
            }
        }

        // ============================================================
        // XML BUILDERS
        // ============================================================

        private static string BuildWorkbookXml()
        {
            // 2 sheet: rId2 -> sheet1, rId3 -> sheet2
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ");
            sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            sb.Append("<sheets>");
            sb.Append("<sheet name=\"Tree Compare\" sheetId=\"1\" r:id=\"rId2\"/>");
            sb.Append("<sheet name=\"Part Compare\" sheetId=\"2\" r:id=\"rId3\"/>");
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        private static string BuildStylesXml()
        {
            // Fonts: 0=normal, 1=bold
            // Fills: 0=none(reserved), 1=gray125(reserved), 2=yellow, 3=red
            //   (Excel BAT BUOC fill index 0 va 1 la 2 fill mac dinh nay)
            // CellXfs (style index dung trong s="N"):
            //   0 = default (font 0, fill 0)
            //   1 = bold    (font 1, fill 0)
            //   2 = yellow  (font 0, fill 2)
            //   3 = red     (font 0, fill 3)
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            // Fonts
            sb.Append("<fonts count=\"2\">");
            sb.Append("<font><sz val=\"11\"/><name val=\"Times New Roman\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><name val=\"Times New Roman\"/></font>");
            sb.Append("</fonts>");

            // Fills (index 0 va 1 BAT BUOC theo spec)
            sb.Append("<fills count=\"4\">");
            sb.Append("<fill><patternFill patternType=\"none\"/></fill>");
            sb.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
            // Yellow FFFFFF00
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            // Red FFFF0000
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFF0000\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            sb.Append("</fills>");

            // Borders (1 default)
            sb.Append("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");

            // CellStyleXfs (1 default)
            sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

            // CellXfs
            sb.Append("<cellXfs count=\"4\">");
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");                          // 0 default
            sb.Append("<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>");          // 1 bold
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFill=\"1\"/>");          // 2 yellow
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFill=\"1\"/>");          // 3 red
            sb.Append("</cellXfs>");

            sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        // ---- SHEET 1: TREE COMPARE ----
        // Cot A=Name1, B=Status1, C=Name2, D=Status2
        private static string BuildTreeSheetXml(string name1, string name2, List<BomCompareLogic.TreeRow> rows)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            // Tinh do rong cot tu du lieu
            int maxA = ("BOM 1: " + name1).Length;
            int maxB = "Status".Length;
            int maxC = ("BOM 2: " + name2).Length;
            int maxD = "Status".Length;
            foreach (var r in rows)
            {
                if (r.Name1 != null && r.Name1.Length > maxA) maxA = r.Name1.Length;
                if (r.Status != null && r.Status.Length > maxB) maxB = r.Status.Length;
                if (r.Name2 != null && r.Name2.Length > maxC) maxC = r.Name2.Length;
            }
            sb.Append("<cols>");
            sb.Append(ColDef(1, maxA));
            sb.Append(ColDef(2, maxB));
            sb.Append(ColDef(3, maxC));
            sb.Append(ColDef(4, maxD));
            sb.Append("</cols>");

            sb.Append("<sheetData>");

            // Header row 1
            sb.Append("<row r=\"1\">");
            sb.Append(InlineStrCell("A1", "BOM 1: " + name1, 1));
            sb.Append(InlineStrCell("B1", "Status", 1));
            sb.Append(InlineStrCell("C1", "BOM 2: " + name2, 1));
            sb.Append(InlineStrCell("D1", "Status", 1));
            sb.Append("</row>");

            int r2 = 2;
            foreach (var row in rows)
            {
                int styleIdx = StatusStyle(row.Status);
                sb.Append("<row r=\"" + r2 + "\">");
                sb.Append(InlineStrCell("A" + r2, row.Name1, styleIdx));
                sb.Append(InlineStrCell("B" + r2, row.Status, styleIdx));
                sb.Append(InlineStrCell("C" + r2, row.Name2, styleIdx));
                sb.Append(InlineStrCell("D" + r2, row.Status, styleIdx));
                sb.Append("</row>");
                r2++;
            }

            sb.Append("</sheetData>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        // ---- SHEET 2: PART COMPARE ----
        // Cot A=Name1 B=Qty1 C=Note1 | D=Name2 E=Qty2 F=Note2
        private static string BuildPartSheetXml(string name1, string name2, List<BomCompareLogic.PartRow> rows)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

            // Tinh do rong cot
            int maxA = ("BOM 1: " + name1).Length;
            int maxB = 6; // "Qty" header + so
            int maxC = "Khong co trong bang 2".Length;
            int maxD = ("BOM 2: " + name2).Length;
            int maxE = 6;
            int maxF = maxC;
            foreach (var r in rows)
            {
                if (r.Has1 && r.Name1.Length > maxA) maxA = r.Name1.Length;
                if (r.Has2 && r.Name2.Length > maxD) maxD = r.Name2.Length;
            }
            sb.Append("<cols>");
            sb.Append(ColDef(1, maxA));
            sb.Append(ColDef(2, maxB));
            sb.Append(ColDef(3, maxC));
            sb.Append(ColDef(4, maxD));
            sb.Append(ColDef(5, maxE));
            sb.Append(ColDef(6, maxF));
            sb.Append("</cols>");

            sb.Append("<sheetData>");

            // Header row 1: ten BOM
            sb.Append("<row r=\"1\">");
            sb.Append(InlineStrCell("A1", "BOM 1: " + name1, 1));
            sb.Append(InlineStrCell("D1", "BOM 2: " + name2, 1));
            sb.Append("</row>");

            // Header row 2: column titles
            sb.Append("<row r=\"2\">");
            sb.Append(InlineStrCell("A2", "Part Name", 1));
            sb.Append(InlineStrCell("B2", "Qty", 1));
            sb.Append(InlineStrCell("C2", "Note", 1));
            sb.Append(InlineStrCell("D2", "Part Name", 1));
            sb.Append(InlineStrCell("E2", "Qty", 1));
            sb.Append(InlineStrCell("F2", "Note", 1));
            sb.Append("</row>");

            int r2 = 3;
            foreach (var row in rows)
            {
                sb.Append("<row r=\"" + r2 + "\">");

                if (row.Has1)
                {
                    sb.Append(InlineStrCell("A" + r2, row.Name1, 0));
                    sb.Append(NumberCell("B" + r2, row.Qty1, 0));
                    sb.Append(InlineStrCell("C" + r2, row.Note1, ColorStyle(row.Note1Color)));
                }
                if (row.Has2)
                {
                    sb.Append(InlineStrCell("D" + r2, row.Name2, 0));
                    sb.Append(NumberCell("E" + r2, row.Qty2, 0));
                    sb.Append(InlineStrCell("F" + r2, row.Note2, ColorStyle(row.Note2Color)));
                }

                sb.Append("</row>");
                r2++;
            }

            sb.Append("</sheetData>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        // ============================================================
        // CELL/SHEET HELPERS
        // ============================================================

        /// <summary>Dinh nghia do rong 1 cot. width = so ky tu * 1.3 (uoc luong).</summary>
        private static string ColDef(int colIdx, int maxChars)
        {
            double w = Math.Max(8, maxChars * 1.3);
            return "<col min=\"" + colIdx + "\" max=\"" + colIdx + "\" width=\""
                   + w.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)
                   + "\" bestFit=\"1\" customWidth=\"1\"/>";
        }

        /// <summary>Cell chua inline string (t="inlineStr"). Tranh phai dung sharedStrings.</summary>
        private static string InlineStrCell(string cellRef, string text, int styleIdx)
        {
            if (text == null) text = "";
            string s = styleIdx > 0 ? " s=\"" + styleIdx + "\"" : "";
            return "<c r=\"" + cellRef + "\"" + s + " t=\"inlineStr\"><is><t xml:space=\"preserve\">"
                   + XmlEscape(text) + "</t></is></c>";
        }

        /// <summary>Cell chua so (numeric).</summary>
        private static string NumberCell(string cellRef, int value, int styleIdx)
        {
            string s = styleIdx > 0 ? " s=\"" + styleIdx + "\"" : "";
            return "<c r=\"" + cellRef + "\"" + s + "><v>" + value + "</v></c>";
        }

        private static int StatusStyle(string status)
        {
            // DIFF -> vang (2), ONLY-1/ONLY-2 -> do (3), OK -> default (0)
            if (status == "DIFF") return 2;
            if (status == "ONLY-1" || status == "ONLY-2") return 3;
            return 0;
        }

        private static int ColorStyle(BomCompareLogic.CellColor c)
        {
            switch (c)
            {
                case BomCompareLogic.CellColor.Yellow: return 2;
                case BomCompareLogic.CellColor.Red: return 3;
                default: return 0;
            }
        }

        private static string XmlEscape(string s)
        {
            return s.Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\"", "&quot;")
                    .Replace("'", "&apos;");
        }

        private static void WriteXml(PackagePart part, string xml)
        {
            using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
            {
                // UTF-8 KHONG BOM (Excel strict reader khong thich BOM o giua package)
                var bytes = new UTF8Encoding(false).GetBytes(xml);
                stream.Write(bytes, 0, bytes.Length);
            }
        }
    }
}