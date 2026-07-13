// One-off inspector for the individual-student report template.
// Not part of the normal test run — invoked manually via `dotnet run -- inspect`.
using System;
using System.IO;
using ClosedXML.Excel;

namespace StudentSIS.IntegrationTests;

internal static class Inspect
{
    public static void RunPdf(string path)
    {
        if (!File.Exists(path)) { Console.WriteLine($"NOT FOUND: {path}"); return; }
        Console.WriteLine($"INSPECTING PDF: {path}  ({new FileInfo(path).Length:N0} bytes)");
        try
        {
            using var reader = new iTextSharp.text.pdf.PdfReader(path);
            Console.WriteLine($"  pages: {reader.NumberOfPages}");
            // Use LocationTextExtractionStrategy — handles CID/Identity-H fonts better.
            int totalChars = 0, laoChars = 0;
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                var strat = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
                string p = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, i, strat);
                totalChars += p.Length;
                foreach (var ch in p) if (ch >= '຀' && ch <= '໿') laoChars++;
                if (i == 1)
                {
                    Console.WriteLine($"  page 1 chars: {p.Length}");
                    Console.WriteLine($"  page 1 sample: {p.Substring(0, Math.Min(180, p.Length)).Replace("\n", " | ")}");
                }
            }
            Console.WriteLine($"  total chars (all pages): {totalChars:N0}");
            Console.WriteLine($"  Lao chars (U+0E80–U+0EFF): {laoChars:N0}");
        }
        catch (Exception ex) { Console.WriteLine($"  ERROR: {ex.Message}"); }
    }

    public static void RunDocx(string path)
    {
        if (!File.Exists(path)) { Console.WriteLine($"NOT FOUND: {path}"); return; }
        Console.WriteLine($"INSPECTING DOCX: {path}  ({new FileInfo(path).Length} bytes)");
        try
        {
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, false);
            if (doc.MainDocumentPart == null) { Console.WriteLine("  no main part"); return; }
            var texts = new System.Collections.Generic.List<string>();
            foreach (var t in doc.MainDocumentPart.RootElement!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                texts.Add(t.Text ?? "");
            string combined = string.Concat(texts);
            var rx = new System.Text.RegularExpressions.Regex(@"\{\{[A-Z_]+\}\}");
            var matches = rx.Matches(combined);
            Console.WriteLine($"  tokens found ({matches.Count}):");
            var distinct = new System.Collections.Generic.HashSet<string>();
            foreach (System.Text.RegularExpressions.Match m in matches) distinct.Add(m.Value);
            foreach (var t in distinct) Console.WriteLine($"    {t}");
        }
        catch (Exception ex) { Console.WriteLine($"  ERROR: {ex.Message}"); }
    }

    public static void Run(string path)
    {
        if (!File.Exists(path))
        {
            Console.WriteLine($"NOT FOUND: {path}");
            return;
        }
        Console.WriteLine($"INSPECTING: {path}");
        Console.WriteLine($"  size: {new FileInfo(path).Length} bytes");

        using var wb = new XLWorkbook(path);
        Console.WriteLine($"  sheets: {wb.Worksheets.Count}");
        int si = 0;
        foreach (var ws in wb.Worksheets)
        {
            si++;
            Console.WriteLine();
            Console.WriteLine($"── Sheet {si}: '{ws.Name}' ─────────────────────────────");
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
            Console.WriteLine($"  used range: A1:{(char)('A' + lastCol - 1)}{lastRow}  ({lastRow} rows × {lastCol} cols)");

            // Merged ranges.
            int mergeCount = 0;
            foreach (var m in ws.MergedRanges) mergeCount++;
            Console.WriteLine($"  merged ranges: {mergeCount}");
            int mi = 0;
            foreach (var m in ws.MergedRanges)
            {
                if (mi++ >= 20) { Console.WriteLine("    … (more merges)"); break; }
                Console.WriteLine($"    {m.RangeAddress}");
            }

            // Dump every used cell with its value/formula.
            Console.WriteLine("  cells (R/C : value / formula):");
            for (int r = 1; r <= lastRow; r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    if (cell.IsEmpty() && !cell.HasFormula) continue;
                    string addr = cell.Address.ToString();
                    if (cell.HasFormula)
                        Console.WriteLine($"    {addr} : [formula] = {cell.FormulaA1}");
                    else
                    {
                        string val = cell.GetString();
                        if (val.Length > 80) val = val.Substring(0, 77) + "...";
                        Console.WriteLine($"    {addr} : \"{val}\"");
                    }
                }
            }

            // Column widths.
            Console.WriteLine("  column widths:");
            for (int c = 1; c <= lastCol; c++)
            {
                Console.WriteLine($"    col {c}: width = {ws.Column(c).Width:F2}");
            }

            // Page setup.
            var ps = ws.PageSetup;
            Console.WriteLine($"  page setup: orientation={ps.PageOrientation}, paper={ps.PaperSize}, pagesWide={ps.PagesWide}, pagesTall={ps.PagesTall}");
        }
    }
}
