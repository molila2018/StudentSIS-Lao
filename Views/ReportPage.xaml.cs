using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using StudentSIS.Data;
using PdfDoc   = iTextSharp.text.Document;
using PdfWr    = iTextSharp.text.pdf.PdfWriter;
using PdfFont  = iTextSharp.text.Font;
using PdfBf    = iTextSharp.text.pdf.BaseFont;
using PdfPara  = iTextSharp.text.Paragraph;
using PdfTbl   = iTextSharp.text.pdf.PdfPTable;
using PdfCell  = iTextSharp.text.pdf.PdfPCell;
using PdfPhr   = iTextSharp.text.Phrase;
using PdfEl    = iTextSharp.text.Element;
using PdfPS    = iTextSharp.text.PageSize;
using PdfClr   = iTextSharp.text.BaseColor;

namespace StudentSIS.Views
{
    public partial class ReportPage : UserControl
    {
        private DataTable? _last;
        private string _lastTitle = "";

        public ReportPage()
        {
            InitializeComponent();

            foreach (var y in DB.AcademicYears()) CmbYear.Items.Add(y);
            CmbYear.SelectedIndex = 0;
            CmbSem.Items.Add("ທັງໝົດ"); CmbSem.Items.Add("ພາກ 1"); CmbSem.Items.Add("ພາກ 2");
            CmbSem.SelectedIndex = 0;

            foreach (var g in new[] { "ມ.1","ມ.2","ມ.3","ມ.4" })
                CmbGrade.Items.Add(g);
            CmbGrade.SelectedIndex = 3;
            foreach (var r in new[] { "1","2","3","4","5","6" }) CmbRoom.Items.Add(r);
            CmbRoom.SelectedIndex = 0;

            // Filter the student picker by the chosen grade + room. Both remaining
            // reports (Enrollment Agreement, Profile) operate on one student at a time
            // and use the class picker only to narrow the student dropdown.
            CmbGrade.SelectionChanged += Class_Changed;
            CmbRoom.SelectionChanged  += Class_Changed;
            // Year change refreshes the preview. (Sem/Month dropdowns previously
            // wired here were tied to the removed score reports.)
            CmbYear.SelectionChanged  += (_, _) => RefreshPreview();
            CmbSem.SelectionChanged   += (_, _) => RefreshPreview();
            ReloadStudents();

            // Student-picker status filter — default 'ກຳລັງຮຽນ'. Switching to 'ຈົບ'
            // surfaces graduates for an enrollment-agreement reprint.
            CmbStudentStatus.Items.Add("ກຳລັງຮຽນ");
            CmbStudentStatus.Items.Add("ຈົບ");
            CmbStudentStatus.Items.Add("ອອກ");
            CmbStudentStatus.Items.Add("ທັງໝົດ");
            CmbStudentStatus.SelectedIndex = 0;

            // Default to the first remaining report type (Enrollment Agreement).
            CmbType.SelectedIndex = 0;

            // NavContext handoff from ClassHubPage — pre-fill grade/room/year so the
            // destination view is ready. Score reports it used to deep-link into no
            // longer live here; the class context still pre-fills usefully for the
            // remaining administrative reports.
            if (!string.IsNullOrEmpty(DB.NavGrade))
            {
                for (int i = 0; i < CmbGrade.Items.Count; i++)
                    if (CmbGrade.Items[i]?.ToString() == DB.NavGrade) { CmbGrade.SelectedIndex = i; break; }
                if (!string.IsNullOrEmpty(DB.NavRoom))
                    for (int i = 0; i < CmbRoom.Items.Count; i++)
                        if (CmbRoom.Items[i]?.ToString() == DB.NavRoom) { CmbRoom.SelectedIndex = i; break; }
                if (!string.IsNullOrEmpty(DB.NavYear))
                    for (int i = 0; i < CmbYear.Items.Count; i++)
                        if (CmbYear.Items[i]?.ToString() == DB.NavYear) { CmbYear.SelectedIndex = i; break; }
                if (DB.NavSemester == 1) CmbSem.SelectedIndex = 1;       // "ພາກ 1"
                else if (DB.NavSemester == 2) CmbSem.SelectedIndex = 2;  // "ພາກ 2"
                DB.ClearNav();
            }
        }

        // ── Show/hide the right filter row when switching report type ──
        // Dropdown layout after the score-reports migration to ປະຫວັດຄະແນນນັກຮຽນ:
        //   0  📑 ໃບສັນຍາເຂົ້າຮຽນ (Word/PDF)        → needs STUDENT picker
        //   1  📋 ລາຍງານປະຫວັດນັກຮຽນ (Profile)      → needs STUDENT picker
        //
        // Both are per-student documents; both use the class picker only to narrow
        // the student dropdown. Class picker stays always visible for that purpose.
        private void CmbType_Changed(object s, SelectionChangedEventArgs e)
        {
            if (StudentPickerRow == null || ClassPickerRow == null) return;
            // Every remaining report is per-student.
            StudentPickerRow.Visibility = Visibility.Visible;
            ClassPickerRow.Visibility   = Visibility.Visible;
            ReloadStudents();
            RefreshPreview();
        }

        // Re-query the student dropdown using the currently-selected grade + room.
        // Called whenever CmbGrade or CmbRoom changes — keeps the Enrollment Agreement
        // student picker narrowed to one classroom at a time.  No harm calling this
        // even when the student picker is hidden (other report types just ignore it).
        private void ReloadStudents()
        {
            if (CmbStudent == null || CmbGrade == null || CmbRoom == null) return;
            string grade  = CmbGrade.SelectedItem?.ToString()  ?? "ມ.4";
            string room   = CmbRoom.SelectedItem?.ToString()   ?? "1";
            string status = CmbStudentStatus?.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            int? keepId = CmbStudent.SelectedValue is int i ? i : (int?)null;

            // Always narrow by grade + room so the student dropdown matches the class
            // filters. Status filter is layered on top (default 'ກຳລັງຮຽນ'; switching
            // to 'ຈົບ' surfaces graduates from that same class+room). This applies to
            // every report type that shows the student picker — Enrollment Agreement,
            // Transcript, and Individual — keeping behaviour consistent.
            var sb = new System.Text.StringBuilder(@"
                SELECT StudentID,
                       StudentCode || '  ·  ' || FirstName || ' ' || LastName ||
                       '  ·  ' || GradeLevel || '/' || IFNULL(ClassRoom,'-') AS D
                FROM Students
                WHERE GradeLevel=@g AND ClassRoom=@r");
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@g", grade), ("@r", room)
            };
            if (status != "ທັງໝົດ") { sb.Append(" AND Status=@st"); ps.Add(("@st", status)); }
            sb.Append(" ORDER BY StudentCode");
            var dt = DB.Query(sb.ToString(), null, ps.ToArray());

            CmbStudent.DisplayMemberPath = "D";
            CmbStudent.SelectedValuePath = "StudentID";
            CmbStudent.ItemsSource       = dt.DefaultView;
            if (keepId.HasValue) CmbStudent.SelectedValue = keepId.Value;
            // If the prior selection no longer matches the filter, pick the first row
            // so the user isn't left with an empty student picker.
            if (CmbStudent.SelectedValue == null && dt.Rows.Count > 0)
                CmbStudent.SelectedIndex = 0;

            if (TxtStudentCount != null)
                TxtStudentCount.Text = dt.Rows.Count > 0
                    ? $"({dt.Rows.Count} ຄົນ)"
                    : "(ບໍ່ມີນັກຮຽນທີ່ຕົງກັບຕົວກອງ)";
        }

        private void Class_Changed(object s, SelectionChangedEventArgs e)
        {
            ReloadStudents();
            RefreshPreview();
        }

        // Both remaining reports (Enrollment Agreement, Profile) are single-student
        // Word documents — no class-roster preview to render. Clear the preview pane
        // whenever the selection changes.
        private void RefreshPreview()
        {
            if (CmbType == null) return;
            ClearPreview();
        }

        private void ClearPreview()
        {
            _last = null; _lastTitle = "";
            if (TxtOut != null) TxtOut.Text = "";
            if (TblOut != null) TblOut.ItemsSource = null;
        }

        // Report export. Both remaining reports (Enrollment Agreement +
        // Student Profile) are Word-based; each generator opens its own
        // save dialog that lets the user pick .docx or .pdf. The old
        // "ສ້າງ Excel" button was removed — none of the current reports
        // produce an .xlsx. Score reports moved to ScoreHistoryPage.
        private void BtnPdf_Click(object s, RoutedEventArgs e)
        {
            // Both remaining reports produce their PDF directly from their own
            // save dialogs (Enrollment Agreement = Word→PDF; Student Profile =
            // Word→PDF). The xlsx→PDF-via-Excel-COM path used by the now-moved
            // score reports is unreachable from this UI but stays in
            // ConvertXlsxToPdfViaExcel (still called by score-history windows).
            switch (CmbType.SelectedIndex)
            {
                case 0: GenEnrollmentAgreement();    break;
                case 1: GenStudentProfileReport();   break;
            }
        }

        // Open the xlsx with Microsoft Excel via late-bound COM, then SaveAs PDF.
        // Late binding avoids needing a Microsoft.Office.Interop.Excel reference.
        // xlTypePDF == 0 in XlFixedFormatType.
        internal static void ConvertXlsxToPdfViaExcel(string xlsxPath, string pdfPath)
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Microsoft Excel ບໍ່ໄດ້ຕິດຕັ້ງໃນເຄື່ອງນີ້.");

            dynamic excel = Activator.CreateInstance(excelType)!;
            dynamic? wb = null;
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;
                // Open the file read-only — we just want to render & export.
                wb = excel.Workbooks.Open(Path.GetFullPath(xlsxPath),
                    /*UpdateLinks*/ 0, /*ReadOnly*/ true);
                // 0 = xlTypePDF; quality 0 = standard; openAfterPublish=false
                wb.ExportAsFixedFormat(0, Path.GetFullPath(pdfPath),
                    /*Quality*/ 0, /*IncludeDocProps*/ true,
                    /*IgnorePrintAreas*/ false, Type.Missing, Type.Missing,
                    /*OpenAfterPublish*/ false);
            }
            finally
            {
                try { if (wb != null) { wb.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(wb); } } catch { }
                try { excel.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // ── Generic table → landscape A4 PDF ───────────────────────────
        // Renders any DataTable using two fonts:
        //   Phetsarath OT  → any cell whose content contains Lao characters
        //   Times New Roman → numeric / Latin cells (student codes, scores, ranks)
        private static void ExportTablePdf(string path, string title, DataTable data)
        {
            // Landscape so the wide score grids fit; tight margins to save space.
            var doc = new PdfDoc(PdfPS.A4.Rotate(), 22f, 22f, 30f, 26f);
            using var fs = new FileStream(path, FileMode.Create);
            PdfWr.GetInstance(doc, fs);
            doc.Open();

            var laoBf = LoadFont("Phetsarath OT.ttf", "phetsarath_ot.ttf", "Phetsarath_ot.ttf");
            var tnrBf = LoadFont("times.ttf", "Times New Roman.ttf");
            var fTitle = new PdfFont(laoBf, 14, PdfFont.BOLD);
            var fSub   = new PdfFont(laoBf, 9);
            var fHdrLa = new PdfFont(laoBf, 9, PdfFont.BOLD, new PdfClr(255, 255, 255));
            var fHdrEn = new PdfFont(tnrBf, 9, PdfFont.BOLD, new PdfClr(255, 255, 255));
            var fLao   = new PdfFont(laoBf, 9);
            var fTnr   = new PdfFont(tnrBf, 9);
            var fFail  = new PdfFont(tnrBf, 9, PdfFont.BOLD, new PdfClr(185, 28, 28));
            var fFailLa= new PdfFont(laoBf, 9, PdfFont.BOLD, new PdfClr(185, 28, 28));
            var headerBg = new PdfClr(27, 79, 138);

            doc.Add(new PdfPara(DB.SchoolName, fTitle) { Alignment = PdfEl.ALIGN_CENTER });
            if (!string.IsNullOrWhiteSpace(title))
                doc.Add(new PdfPara(title, fTitle) { Alignment = PdfEl.ALIGN_CENTER });
            doc.Add(new PdfPara($"ສ້າງ: {DateTime.Now:dd/MM/yyyy HH:mm}    ໂດຍ: {DB.CurrentUser}",
                fSub) { Alignment = PdfEl.ALIGN_RIGHT });
            doc.Add(new PdfPara(" ", fSub));

            int cols = data.Columns.Count;
            var tbl = new PdfTbl(cols) { WidthPercentage = 100 };
            tbl.HeaderRows = 1;

            // Header row — pick font per column based on the column header text.
            foreach (DataColumn c in data.Columns)
            {
                var hf = ContainsLao(c.ColumnName) ? fHdrLa : fHdrEn;
                var cell = new PdfCell(new PdfPhr(c.ColumnName, hf))
                {
                    BackgroundColor = headerBg,
                    HorizontalAlignment = PdfEl.ALIGN_CENTER,
                    VerticalAlignment = PdfEl.ALIGN_MIDDLE,
                    Padding = 4f,
                    MinimumHeight = 22f
                };
                tbl.AddCell(cell);
            }

            // Data rows — Lao for Lao text (incl. "ຕົກ"/"ເສັງໃດ້"), TNR for the rest.
            // Cells holding the literal "ຕົກ" are bold-red to flag failures, matching the Excel.
            foreach (DataRow row in data.Rows)
            {
                foreach (DataColumn c in data.Columns)
                {
                    string text = row[c]?.ToString() ?? "";
                    bool isLao  = ContainsLao(text);
                    bool isFail = text == "ຕົກ";
                    PdfFont f = isFail ? (isLao ? fFailLa : fFail)
                                      : (isLao ? fLao    : fTnr);
                    var cell = new PdfCell(new PdfPhr(text, f))
                    {
                        HorizontalAlignment = isLao && c.ColumnName == "ຊື່ນັກຮຽນ"
                            ? PdfEl.ALIGN_LEFT : PdfEl.ALIGN_CENTER,
                        VerticalAlignment = PdfEl.ALIGN_MIDDLE,
                        Padding = 3f,
                        MinimumHeight = 16f
                    };
                    tbl.AddCell(cell);
                }
            }
            doc.Add(tbl);
            doc.Close();
        }

        // Locate a font by trying the bundled Templates\Fonts dir first, then the
        // Windows fonts dir. Falls back to Helvetica (Latin only) if nothing matches.
        private static PdfBf LoadFont(params string[] candidates)
        {
            string fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            string projectFontsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fonts");
            foreach (var name in candidates)
            {
                foreach (var dir in new[] { projectFontsDir, fontsDir })
                {
                    var fp = Path.Combine(dir, name);
                    if (!File.Exists(fp)) continue;
                    try { return PdfBf.CreateFont(fp, PdfBf.IDENTITY_H, PdfBf.EMBEDDED); }
                    catch { /* try next candidate */ }
                }
            }
            return PdfBf.CreateFont(PdfBf.HELVETICA, PdfBf.CP1252, false);
        }

        // Subjects sorted to match the school's printed logbook layout
        // (matches the template at report ຄະແນນ.xlsx):
        //   LAO · CIV · MATH · SCI · GEO · HIS · ENG · ICT · PE · MUS · ART · (rest)
        // CHA1 and LAB1 are split off as evaluation-only columns at the end.
        private static readonly string[] TemplateSubjectOrder = {
            "LAO1","CIV1","MATH1","SCI1","GEO1","HIS1","ENG1","ICT1","PE1","MUS1","ART1"
        };

        private static int TemplateOrderIndex(string code)
        {
            for (int i = 0; i < TemplateSubjectOrder.Length; i++)
                if (TemplateSubjectOrder[i] == code) return i;
            return int.MaxValue; // anything not in the list sorts to the end
        }


        // ═══════════════════════════════════════════════════════════
        //  Score templates — REPORT 6 (monthly) and REPORT 8 (yearly)
        //
        //  Both reports use the same Excel template that the school maintains:
        //      Templates/ໃບຄະແນນ.xlsx
        //  Row 1 (merged D1:O1) holds the title.
        //  Row 2 is the header (subject names + totals + character/labor).
        //  Row 3 is the FORMULA row — its SUM / AVERAGE / RANK / COUNTIF / IF
        //  formulas are written by the school in the template itself.  The
        //  code's job is only to: (1) rewrite the title, (2) ensure there are
        //  enough data rows by cloning row 3 down, and (3) overwrite the data
        //  columns (A, B, C, the subject columns, and ຄຸນສົມບັດ / ແຮງງານ).
        //
        //  Subject → column mapping is by HEADER TEXT in row 2, so the school
        //  can rearrange columns in the template without touching code.
        // ═══════════════════════════════════════════════════════════

        // SubjectCode → list of header substrings the template might use for it.
        // The first one found in row 2 wins. Keeps the mapping resilient to
        // tiny spelling variations between the template and the DB.
        private static readonly Dictionary<string, string[]> CodeHeaderMap = new()
        {
            ["LAO1"]  = new[] { "ພາສາລາວ" },
            ["CIV1"]  = new[] { "ສຶກສາພົນລະເມືອງ", "ພົນລະເມືອງ" },
            ["MATH1"] = new[] { "ຄະນິດສາດ" },
            ["SCI1"]  = new[] { "ວິທະຍາສາດທຳມະຊາດ", "ວິທະຍາສາດ" },
            ["GEO1"]  = new[] { "ພູມສາດ" },
            ["HIS1"]  = new[] { "ປະຫວັດສາດ" },
            ["ENG1"]  = new[] { "ພາສາອັງກິດ" },
            ["ICT1"]  = new[] { "ເຕັກໂນໂລຊີ", "ICT" },
            ["PE1"]   = new[] { "ພະລະສຶກສາ" },
            ["MUS1"]  = new[] { "ສິລະປະດົນຕີ", "ສີລະປະດົນຕີ", "ດົນຕີ" },
            ["ART1"]  = new[] { "ສິລະປະກຳ", "ສີລະປະກຳ" },
            ["VOC1"]  = new[] { "ພື້ນຖານວິຊາຊີບ", "ວິຊາຊີບ" },
            ["CHA1"]  = new[] { "ຄຸນສົມບັດ" },
            ["LAB1"]  = new[] { "ແຮງງານ", "ການອອກແຮງງານ" },
        };


        // Layout of one score sheet, learnt from the template at runtime.
        // The 4 template sheets (monthly / sem-1 / sem-2 / annual) have slightly
        // different column starts (Sheet1/4 start at A, Sheet2/3 at C), so we
        // detect everything by header text rather than hardcoded letters.
        // HeaderRow + DataStart are detected at runtime too — the template's
        // header row used to be row 2, but the school added 3 header lines
        // (national name + motto + school name) above the report title, so
        // headers now live at row 5 and data starts at row 6. Scanning for
        // the "ລຳດັບ" cell keeps the code resilient to future header edits.
        private class SheetLayout
        {
            public ClosedXML.Excel.IXLCell TitleCell = null!;
            public int OrderCol, CodeCol, NameCol;
            public int CharCol, LaborCol;
            public Dictionary<string, int> SubjectColByCode = new();
            public int MaxCol;
            public int HeaderRow;          // row containing "ລຳດັບ" (was 2, now 5)
            public int DataStart;          // first data row = HeaderRow + 1
        }

        // Find headers by scanning the first ~20 rows for the cell containing
        // "ລຳດັບ" — that anchors every other layout fact. Title cell = first
        // merged range in (HeaderRow - 1), which is where the dynamic report
        // title lives ("ສະຫຼຸບຄະແນນປະຈຳເດືອນ..."). Subject columns are matched
        // by Lao header substring (see CodeHeaderMap).
        private static SheetLayout MapSheet(ClosedXML.Excel.IXLWorksheet ws)
        {
            var L = new SheetLayout { MaxCol = ws.LastColumnUsed()?.ColumnNumber() ?? 24 };

            // Locate the header row by finding "ລຳດັບ" anywhere in rows 1-20.
            int headerRow = 0, orderCol = 0;
            for (int r = 1; r <= 20 && headerRow == 0; r++)
            {
                for (int c = 1; c <= L.MaxCol; c++)
                {
                    if (ws.Cell(r, c).GetString().Trim().Contains("ລຳດັບ"))
                    {
                        headerRow = r; orderCol = c; break;
                    }
                }
            }
            // Fall back to the old layout (headers in row 2) if detection fails —
            // mostly so unit tests on synthetic minimal sheets still parse.
            if (headerRow == 0) { headerRow = 2; orderCol = 1; }
            L.HeaderRow = headerRow;
            L.DataStart = headerRow + 1;
            L.OrderCol  = orderCol;

            // Title cell = first merged range in the row immediately above the
            // headers. That's where "ສະຫຼຸບ..." lives in the current template.
            ClosedXML.Excel.IXLCell? titleCell = null;
            int titleRow = Math.Max(1, headerRow - 1);
            foreach (var mr in ws.MergedRanges)
            {
                if (mr.FirstRow().RowNumber() == titleRow) { titleCell = mr.FirstCell(); break; }
            }
            // Also accept the very first merged range as a fallback (handles
            // legacy templates where the title was at row 1).
            if (titleCell == null)
            {
                foreach (var mr in ws.MergedRanges)
                {
                    if (mr.FirstRow().RowNumber() == 1) { titleCell = mr.FirstCell(); break; }
                }
            }
            L.TitleCell = titleCell ?? ws.Cell(titleRow, 4);

            for (int c = 1; c <= L.MaxCol; c++)
            {
                string h = ws.Cell(headerRow, c).GetString().Trim();
                if (string.IsNullOrEmpty(h)) continue;
                if (h.Contains("ລະຫັດ"))                             { L.CodeCol  = c; continue; }
                if (h.Contains("ຊື່") && h.Contains("ນາມສະກຸນ"))      { L.NameCol  = c; continue; }
                bool matched = false;
                foreach (var kv in CodeHeaderMap)
                {
                    foreach (var needle in kv.Value)
                    {
                        if (h.Contains(needle))
                        {
                            if (kv.Key == "CHA1")      L.CharCol  = c;
                            else if (kv.Key == "LAB1") L.LaborCol = c;
                            else                       L.SubjectColByCode[kv.Key] = c;
                            matched = true; break;
                        }
                    }
                    if (matched) break;
                }
            }
            // Sensible defaults if a template column is missing.
            if (L.CodeCol == 0) L.CodeCol = L.OrderCol + 1;
            if (L.NameCol == 0) L.NameCol = L.CodeCol  + 1;
            return L;
        }

        // Shift the row references in an A1 formula by a delta, while keeping
        // absolute refs (`$3`, `$A$3`, etc.) anchored.  Examples (delta = +1, oldRow=3):
        //   D3                    → D4
        //   $D3                   → $D4   (column absolute, row relative)
        //   D$3                   → D$3   (row absolute — unchanged)
        //   $P$3:$P$50            → $P$3:$P$50  (both absolute)
        //   D3:O3                 → D4:O4
        //   COUNTIF(D3:O3, "<5")  → COUNTIF(D4:O4, "<5")
        // Handles `[A-Z]+` column letters (no sheet refs in this template).
        private static readonly System.Text.RegularExpressions.Regex CellRefRx =
            new(@"(\$?)([A-Z]+)(\$?)(\d+)", System.Text.RegularExpressions.RegexOptions.Compiled);
        // Cached result of system-font detection — picks the best Lao-capable font
        // actually installed on this machine. We can't rely on "Phetsarath OT" being
        // system-installed because the WPF app loads it as a private font from its
        // own Fonts\ folder; Excel doesn't see private fonts and falls back to a
        // Latin-only font when exporting to PDF, producing blank Lao glyphs.
        private static string? _laoFontCache;
        private static string ResolveLaoFontName()
        {
            if (_laoFontCache != null) return _laoFontCache;
            string[] candidates = {
                "Phetsarath OT", "Phetsarath",
                "Saysettha OT",  "Saysettha",
                "Noto Sans Lao", "Noto Serif Lao",
                "Lao UI",        "Leelawadee UI", "Leelawadee"
            };
            var installed = new HashSet<string>(
                System.Windows.Media.Fonts.SystemFontFamilies.Select(f => f.Source),
                StringComparer.OrdinalIgnoreCase);
            foreach (var name in candidates)
                if (installed.Contains(name)) return _laoFontCache = name;
            // Final fallback — Leelawadee UI ships with Windows 8.1+ and covers Lao.
            return _laoFontCache = "Leelawadee UI";
        }

        // True if `s` contains any character in the Lao Unicode block (U+0E80–U+0EFF).
        private static bool ContainsLao(string? s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (var ch in s)
                if (ch >= '຀' && ch <= '໿') return true;
            return false;
        }

        private static string ShiftFormulaRows(string formula, int oldRow, int newRow)
        {
            int delta = newRow - oldRow;
            return CellRefRx.Replace(formula, m =>
            {
                string colAbs   = m.Groups[1].Value;
                string colLet   = m.Groups[2].Value;
                string rowAbs   = m.Groups[3].Value;
                int    rowNum   = int.Parse(m.Groups[4].Value);
                // Only relative row refs shift; absolute ($) stays put.
                int    adjusted = rowAbs == "$" ? rowNum : rowNum + delta;
                return $"{colAbs}{colLet}{rowAbs}{adjusted}";
            });
        }

        // Fill one worksheet with one row per student.
        //
        // Formula preservation strategy (replaces the earlier reliance on
        // ClosedXML's InsertRowsBelow, which in 0.102.x doesn't always carry
        // formulas with their refs into the inserted rows):
        //   1. Read each formula column's A1 formula from row 3 once.
        //   2. Strip every data row including the sample student in row 3.
        //   3. For each student, write data cells + write each formula explicitly,
        //      with row refs shifted to the current row (absolute refs anchored).
        //
        // This guarantees EVERY data row has all formula columns populated, not
        // just the first row.
        private static void FillSheet(
            ClosedXML.Excel.IXLWorksheet ws,
            string title,
            DataTable students,
            Func<int /*sid*/, string /*code*/, double?> getScore)
        {
            var L = MapSheet(ws);
            string laoFont = ResolveLaoFontName();
            // Page setup — applies to the saved .xlsx AND to the PDF that Excel
            // produces via ExportAsFixedFormat (uses the sheet's print settings).
            //   Landscape A4.
            //   Fit-to-1-page-wide: Excel scales the printout horizontally so the
            //     entire table fits within the page width — nothing overflows to
            //     a second page on the right. Page height is unconstrained, so
            //     long rosters can spill onto additional pages downward.
            //   Narrow margins maximise usable width.
            //   Centred horizontally so the table sits nicely on the page.
            ws.PageSetup.PageOrientation = ClosedXML.Excel.XLPageOrientation.Landscape;
            ws.PageSetup.PaperSize       = ClosedXML.Excel.XLPaperSize.A4Paper;
            ws.PageSetup.PagesWide       = 1;
            ws.PageSetup.PagesTall       = 0;   // unlimited tall
            ws.PageSetup.Margins.Left    = 0.25;
            ws.PageSetup.Margins.Right   = 0.25;
            ws.PageSetup.Margins.Top     = 0.5;
            ws.PageSetup.Margins.Bottom  = 0.5;
            ws.PageSetup.CenterHorizontally = true;
            // 1) Overwrite the title (works even when the cell is merged).
            L.TitleCell.Value = title;
            L.TitleCell.Style.Font.FontName = laoFont;

            int dataStart = L.DataStart;

            // 2) Snapshot the formula columns + their A1 formulas BEFORE we
            //    touch any rows.  This is the source of truth for what to
            //    write on every student row.
            var formulaTemplates = new Dictionary<int, string>();
            for (int c = 1; c <= L.MaxCol; c++)
            {
                if (ws.Cell(dataStart, c).HasFormula)
                    formulaTemplates[c] = ws.Cell(dataStart, c).FormulaA1;
            }

            // 2b) Snapshot the signature row (ອຳນວຍການ / ວິຊາການ / ຄູປະຈຳຫ້ອງ
            //     lives at row 10 in the static template, ABOVE where the
            //     student data will end up after expansion). Capture its
            //     non-empty cells + merge column spans now, before the strip
            //     in step 3 deletes it, then re-emit it below the students
            //     in step 5. Without this the signatures vanish completely
            //     in any generated class report.
            int lastUsed = ws.LastRowUsed()?.RowNumber() ?? dataStart;
            var sigLabels = new[] { "ອຳນວຍການ", "ວິຊາການ", "ຄູປະຈຳຫ້ອງ", "ຜູປົກຄອງ", "ຜູ້ປົກຄອງ" };
            int sigRow = 0;
            var sigCells = new List<(int col, string value)>();
            var sigMergeSpans = new List<(int firstCol, int lastCol)>();
            for (int r = dataStart; r <= lastUsed && sigRow == 0; r++)
            {
                for (int c = 1; c <= L.MaxCol; c++)
                {
                    string t = ws.Cell(r, c).GetString().Trim();
                    if (Array.IndexOf(sigLabels, t) >= 0) { sigRow = r; break; }
                }
            }
            if (sigRow > 0)
            {
                for (int c = 1; c <= L.MaxCol; c++)
                {
                    string t = ws.Cell(sigRow, c).GetString().Trim();
                    if (!string.IsNullOrEmpty(t)) sigCells.Add((c, t));
                }
                foreach (var mr in ws.MergedRanges.ToList())
                {
                    if (mr.FirstRow().RowNumber() == sigRow)
                        sigMergeSpans.Add((mr.FirstColumn().ColumnNumber(),
                                           mr.LastColumn().ColumnNumber()));
                }
            }

            // 3) Strip ALL sample rows (including the model row + signature row)
            //    so we start clean. Signature gets re-emitted at step 5.
            for (int r = lastUsed; r >= dataStart; r--) ws.Row(r).Delete();

            // 4) Write each student row from scratch: data + formulas.
            int idx = 0;
            foreach (DataRow s in students.Rows)
            {
                int rowIdx = dataStart + idx;
                int sid = Convert.ToInt32(s["StudentID"]);

                // 4a) Identity columns
                ws.Cell(rowIdx, L.OrderCol).Value = idx + 1;
                ws.Cell(rowIdx, L.CodeCol).Value  = s["StudentCode"].ToString();
                ws.Cell(rowIdx, L.NameCol).Value  = s["FullName"].ToString();

                // 4b) Subject scores (only columns the code recognises)
                foreach (var kv in L.SubjectColByCode)
                {
                    var v = getScore(sid, kv.Key);
                    if (v.HasValue) ws.Cell(rowIdx, kv.Value).Value = Math.Round(v.Value, 2);
                }
                if (L.CharCol > 0)
                {
                    var v = getScore(sid, "CHA1");
                    if (v.HasValue) ws.Cell(rowIdx, L.CharCol).Value = Math.Round(v.Value, 2);
                }
                if (L.LaborCol > 0)
                {
                    var v = getScore(sid, "LAB1");
                    if (v.HasValue) ws.Cell(rowIdx, L.LaborCol).Value = Math.Round(v.Value, 2);
                }

                // 4c) Formula columns — apply EVERY template formula to THIS row,
                //     with row references shifted from the source row (3) to rowIdx.
                foreach (var kv in formulaTemplates)
                {
                    string shifted = ShiftFormulaRows(kv.Value, dataStart, rowIdx);
                    ws.Cell(rowIdx, kv.Key).FormulaA1 = shifted;
                }
                idx++;
            }

            // 4d) Re-emit the signature row two rows below the last student row.
            //     We captured cells + merge spans in step 2b before they were
            //     stripped; recreate them now so the signature block stays at
            //     the bottom of every generated class report.
            if (sigRow > 0 && sigCells.Count > 0)
            {
                int newSigRow = dataStart + idx + 2;
                foreach (var (col, value) in sigCells)
                {
                    var c = ws.Cell(newSigRow, col);
                    c.Value = value;
                    c.Style.Font.FontName = laoFont;
                    c.Style.Font.Bold = true;
                    c.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    c.Style.Alignment.Vertical   = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                }
                foreach (var (firstCol, lastCol) in sigMergeSpans)
                {
                    if (lastCol > firstCol)
                        ws.Range(newSigRow, firstCol, newSigRow, lastCol).Merge();
                }
            }

            // 5) Apply thin borders to every data cell so newly added rows match the
            //    template's header gridlines. Deleting the model row wipes its style,
            //    so we re-apply borders here across the full written range.
            //    Set all four sides on each cell explicitly — ClosedXML's
            //    OutsideBorder/InsideBorder shortcuts can silently no-op on cells
            //    that have a default-row style override.
            // 5) Borders + centre-align on the data rows. Skip cells with no
            //    value/formula so blank columns stay borderless.
            if (idx > 0)
            {
                var thin = ClosedXML.Excel.XLBorderStyleValues.Thin;
                for (int r = dataStart; r < dataStart + idx; r++)
                {
                    for (int c = 1; c <= L.MaxCol; c++)
                    {
                        var cell = ws.Cell(r, c);
                        if (cell.IsEmpty()) continue;
                        cell.Style.Border.TopBorder    = thin;
                        cell.Style.Border.BottomBorder = thin;
                        cell.Style.Border.LeftBorder   = thin;
                        cell.Style.Border.RightBorder  = thin;
                        cell.Style.Alignment.Horizontal =
                            ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        cell.Style.Alignment.Vertical =
                            ClosedXML.Excel.XLAlignmentVerticalValues.Center;
                    }
                }
            }

            // 6) Font pass — override fonts on every non-empty cell in the
            //    whole used range (header rows + data rows). The template's
            //    pre-existing styles may reference a font name that isn't
            //    installed system-wide (e.g. "Phetsarath OT" registered only
            //    as a WPF private font), which causes Excel to fall back to
            //    a Latin-only font and produce blank Lao glyphs in PDF export.
            int lastRow = ws.LastRowUsed()?.RowNumber() ?? (dataStart + idx - 1);
            for (int r = 1; r <= lastRow; r++)
            {
                for (int c = 1; c <= L.MaxCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    if (cell.IsEmpty()) continue;
                    string probe = cell.HasFormula ? cell.FormulaA1 : cell.GetString();
                    cell.Style.Font.FontName = ContainsLao(probe)
                        ? laoFont
                        : "Times New Roman";
                }
            }
        }


        // Copy the template, delete every worksheet except keepSheetIndex (1-based),
        // fill the surviving sheet, save. Returns true on success.
        private bool ProduceTemplateReport(
            int keepSheetIndex, string outPath, string templatePath,
            string title, DataTable roster, Func<int, string, double?> getScore)
        {
            File.Copy(templatePath, outPath, overwrite: true);
            using (var wb = new ClosedXML.Excel.XLWorkbook(outPath))
            {
                if (keepSheetIndex < 1 || keepSheetIndex > wb.Worksheets.Count)
                {
                    MessageBox.Show($"ແມ່ແບບບໍ່ມີ Sheet {keepSheetIndex}", "ຜິດພາດ");
                    return false;
                }
                // Walk a snapshot of the worksheet list because deleting mutates it.
                var sheets = wb.Worksheets.ToList();
                for (int i = 0; i < sheets.Count; i++)
                {
                    if (i + 1 != keepSheetIndex) sheets[i].Delete();
                }
                FillSheet(wb.Worksheet(1), title, roster, getScore);
                wb.Save();
            }
            return true;
        }


        // Scope item for the Individual report (idx 7). Kind:
        //   "M" → monthly,        Value = month (2..12, excluding final-exam months 1,6)
        //   "S" → semester total, Value = 1 | 2
        //   "A" → annual,         Value = 0
        private class IndScope
        {
            public string Kind { get; }
            public int Value { get; }
            private readonly string _label;
            public IndScope(string kind, int value, string label) { Kind = kind; Value = value; _label = label; }
            public override string ToString() => _label;
        }

        // ═══════════════════════════════════════════════════════════
        //  REPORT 7 — Enrollment Agreement (ໃບສັນຍາເຂົ້າຮຽນ, .docx)
        //  Template-based: copies Templates/ໃບສັນຍາ.docx beside the EXE and
        //  replaces {{TOKEN}} placeholders with the student's data. All
        //  formatting (fonts, margins, alignment, school stamps if any) is
        //  inherited from the .docx the school prepared — this code only
        //  fills in the blanks. Edit the template in Word any time without
        //  recompiling.
        // ═══════════════════════════════════════════════════════════
        // ═══════════════════════════════════════════════════════════
        //  REPORT 7 — Individual student score report (ໃບຄະແນນບຸກຄົນ)
        //  Uses Templates/ລີພອດຄະແນນບຸກຄົນ.xlsx as the layout shell.
        //
        //  Layout (updated to match v2 of the school's template):
        //    A1, A2  — national header (merged A1:G1, A2:G2)
        //    C4      — title (merged C4:E4, changes per scope)
        //    A5      — name + class line (merged A5:G5)
        //    A7      — "ວິຊາ" header (merged A7:E7)
        //    F7      — "ຄະແນນ" header
        //    G7      — "ວັນຂາດ" (absence days; left blank — no per-subject attendance data)
        //    H7      — "ໝາຍເຫດ" (remarks; left blank)
        //    A8–A19  — 12 academic subject labels (template's ordering)
        //    F8–F19  — 12 academic subject scores (what this code writes)
        //    F20     — sum, F21 — average, F22 — rank
        //    F23     — CHA1, F24 — LAB1 (manual eval scores)
        //    A27     — ຄູປະຈຳຫ້ອງ signature, G27 — ຜູ້ປົກຄອງ signature
        //
        //  Scope is driven by CmbIndScope: monthly / semester / annual.
        // ═══════════════════════════════════════════════════════════

        // Template's display order for the 12 academic subjects (row → SubjectCode).
        // Distinct from SortOrder; matches the printed school logbook layout.
        private static readonly (int Row, string Code)[] TemplateSubjectRows = {
            (8,  "LAO1"),
            (9,  "GEO1"),
            (10, "HIS1"),
            (11, "CIV1"),
            (12, "SCI1"),
            (13, "MATH1"),
            (14, "ENG1"),
            (15, "ICT1"),
            (16, "VOC1"),
            (17, "PE1"),
            (18, "ART1"),
            (19, "MUS1"),
        };

        // Constants for the template's anchor cells. Centralised so any future
        // template revision only requires touching these.
        private const int IndScoreCol = 6;   // column F — the score column
        private const int IndSumRow   = 20;
        private const int IndAvgRow   = 21;
        private const int IndRankRow  = 22;
        private const int IndChaRow   = 23;
        private const int IndLabRow   = 24;


        // ── Score gatherer ─────────────────────────────────────────────
        private class IndScoreBundle
        {
            public Dictionary<string, double?> AcademicByCode = new();
            public double? Cha, Lab;
            public double Sum, Avg;
            public string RankLabel = "—";   // numeric "1"/"2"/… OR "ຕົກ"
        }

        private static IndScoreBundle ComputeIndividualScores(int sid, string year, IndScope scope, DataRow stu)
        {
            var bundle = new IndScoreBundle();
            string grade = stu["GradeLevel"].ToString()!;
            string room  = stu["ClassRoom"].ToString()!;

            // ── This student's per-subject score ──────────────────
            switch (scope.Kind)
            {
                case "M":
                    // Academic subjects: per-month sum from MonthlyAssessments (auto-calculated
                    // Activity + Discipline + Homework). CHA1/LAB1 are EXCLUDED from this query
                    // and read separately from EvaluationScores(Month{N}) — never derived.
                    var mDt = DB.Query(@"
                        SELECT sub.SubjectCode,
                               IFNULL(ma.ActivityScore,0)+IFNULL(ma.DisciplineScore,0)+IFNULL(ma.HomeworkScore,0) AS V,
                               (ma.ActivityScore IS NOT NULL OR ma.DisciplineScore IS NOT NULL OR ma.HomeworkScore IS NOT NULL) AS HasRow
                        FROM Enrollments e
                        JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                        LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@m
                        WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                          AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                        null, ("@s", sid), ("@y", year), ("@m", scope.Value),
                              ("@sm", DB.SemesterForMonth(scope.Value)));
                    foreach (DataRow r in mDt.Rows)
                    {
                        string code = r["SubjectCode"].ToString()!;
                        bool hasRow = Convert.ToInt32(r["HasRow"]) != 0;
                        double v = Convert.ToDouble(r["V"]);
                        bundle.AcademicByCode[code] = hasRow ? v : (double?)null;
                    }
                    // CHA1/LAB1 monthly — manual entry via EvaluationScores(Month1..Month8).
                    string monthCtx = DB.MonthContextName(scope.Value) ?? "";
                    if (!string.IsNullOrEmpty(monthCtx))
                    {
                        bundle.Cha = DB.GetEvaluationScore(sid, year, monthCtx, "CHA1");
                        bundle.Lab = DB.GetEvaluationScore(sid, year, monthCtx, "LAB1");
                    }
                    break;

                case "S":
                    // Semester academic score = Scores.TotalScore. CHA1/LAB1 from EvaluationScores.
                    var sDt = DB.Query(@"
                        SELECT sub.SubjectCode, sc.TotalScore AS V,
                               (sc.ScoreID IS NOT NULL) AS HasRow
                        FROM Enrollments e
                        JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                        LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                        WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                          AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                        null, ("@s", sid), ("@y", year), ("@sm", scope.Value));
                    foreach (DataRow r in sDt.Rows)
                    {
                        string code = r["SubjectCode"].ToString()!;
                        bool hasRow = Convert.ToInt32(r["HasRow"]) != 0;
                        double? v = hasRow && r["V"] != DBNull.Value ? Convert.ToDouble(r["V"]) : (double?)null;
                        bundle.AcademicByCode[code] = v;
                    }
                    bundle.Cha = DB.GetEvaluationScore(sid, year, $"SEM{scope.Value}", "CHA1");
                    bundle.Lab = DB.GetEvaluationScore(sid, year, $"SEM{scope.Value}", "LAB1");
                    break;

                case "A":
                    // Annual academic = average of Sem1+Sem2 TotalScore. CHA1/LAB1 from ANNUAL eval.
                    var aDt = DB.Query(@"
                        SELECT sub.SubjectCode, IFNULL(sc.TotalScore,0) AS V
                        FROM Enrollments e
                        JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                        LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                        WHERE e.StudentID=@s AND e.AcademicYear=@y
                          AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                          AND sc.ScoreID IS NOT NULL",
                        null, ("@s", sid), ("@y", year));
                    var bag = new Dictionary<string, List<double>>();
                    foreach (DataRow r in aDt.Rows)
                    {
                        string code = r["SubjectCode"].ToString()!;
                        double v = Convert.ToDouble(r["V"]);
                        if (v <= 0) continue;
                        if (!bag.TryGetValue(code, out var list)) bag[code] = list = new List<double>();
                        list.Add(v);
                    }
                    foreach (var kv in bag) bundle.AcademicByCode[kv.Key] = Math.Round(kv.Value.Average(), 2);
                    bundle.Cha = DB.GetEvaluationScore(sid, year, "ANNUAL", "CHA1");
                    bundle.Lab = DB.GetEvaluationScore(sid, year, "ANNUAL", "LAB1");
                    break;
            }

            // ── Sum + average over the 12 academic subjects (CHA1/LAB1 excluded). ──
            double sum = 0; int cnt = 0; int failed = 0;
            foreach (var (_, code) in TemplateSubjectRows)
            {
                if (bundle.AcademicByCode.TryGetValue(code, out var v) && v.HasValue)
                {
                    sum += v.Value; cnt++;
                    if (v.Value < DB.PassScore) failed++;
                }
            }
            bundle.Sum = Math.Round(sum, 2);
            bundle.Avg = cnt > 0 ? Math.Round(sum / cnt, 2) : 0.0;

            // ── Rank within class (passing students only; failed → "ຕົກ"). ──
            // Build per-classmate sum under the same scope. Uses the HISTORICAL
            // cohort (GetHistoricalClassRoster) so the rank is correct for past
            // years too — graduated + promoted students still rank against the
            // people they actually sat in class with. The previous query on
            // Students.GradeLevel/ClassRoom/AcademicYear was current-only and
            // returned 0 classmates for any past year, leaving RankLabel="—".
            var classmates = DB.GetHistoricalClassRoster(year, grade, room);
            var totals = new List<(int sid, double total, bool passed)>();
            foreach (DataRow cr in classmates.Rows)
            {
                int csid = Convert.ToInt32(cr["StudentID"]);
                var cb = ClassmateAcademicSum(csid, year, scope);
                totals.Add((csid, cb.sum, cb.passed));
            }
            // Failed students → "ຕົກ" (this student gets that label too if they failed).
            if (failed > 0 || cnt == 0)
            {
                bundle.RankLabel = "ຕົກ";
            }
            else
            {
                var passing = totals.Where(t => t.passed).OrderByDescending(t => t.total).ToList();
                int rank = 0; double prev = double.NaN;
                for (int i = 0; i < passing.Count; i++)
                {
                    if (passing[i].total != prev) { rank = i + 1; prev = passing[i].total; }
                    if (passing[i].sid == sid) { bundle.RankLabel = rank.ToString(); break; }
                }
                if (bundle.RankLabel == "—") bundle.RankLabel = "—"; // shouldn't happen, defensive
            }

            return bundle;
        }

        // Helper: compute one classmate's sum + pass-flag under the chosen scope.
        // Returns (sum-of-academic-scores, passed) — passed = no subject below PassScore
        // among the academic ones that DO have a value.
        private static (double sum, bool passed) ClassmateAcademicSum(int sid, string year, IndScope scope)
        {
            double sum = 0; int cnt = 0, failed = 0;
            void Add(double? v)
            {
                if (!v.HasValue) return;
                sum += v.Value; cnt++;
                if (v.Value < DB.PassScore) failed++;
            }
            switch (scope.Kind)
            {
                case "M":
                    var mDt = DB.Query(@"
                        SELECT sub.SubjectCode,
                               IFNULL(ma.ActivityScore,0)+IFNULL(ma.DisciplineScore,0)+IFNULL(ma.HomeworkScore,0) AS V,
                               (ma.MonthlyID IS NOT NULL) AS HasRow
                        FROM Enrollments e
                        JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                        LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@m
                        WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                          AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                        null, ("@s", sid), ("@y", year), ("@m", scope.Value),
                              ("@sm", DB.SemesterForMonth(scope.Value)));
                    foreach (DataRow r in mDt.Rows)
                    {
                        bool hasRow = Convert.ToInt32(r["HasRow"]) != 0;
                        Add(hasRow ? Convert.ToDouble(r["V"]) : (double?)null);
                    }
                    break;
                case "S":
                    var sDt = DB.Query(@"
                        SELECT sc.TotalScore AS V, (sc.ScoreID IS NOT NULL) AS HasRow
                        FROM Enrollments e
                        JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                        LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                        WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                          AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                        null, ("@s", sid), ("@y", year), ("@sm", scope.Value));
                    foreach (DataRow r in sDt.Rows)
                        Add(Convert.ToInt32(r["HasRow"]) != 0 && r["V"] != DBNull.Value
                              ? Convert.ToDouble(r["V"]) : (double?)null);
                    break;
                case "A":
                    var aDt = DB.Query(@"
                        SELECT sub.SubjectCode, IFNULL(sc.TotalScore,0) AS V
                        FROM Enrollments e
                        JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                        LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                        WHERE e.StudentID=@s AND e.AcademicYear=@y
                          AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                          AND sc.ScoreID IS NOT NULL",
                        null, ("@s", sid), ("@y", year));
                    var bag = new Dictionary<string, List<double>>();
                    foreach (DataRow r in aDt.Rows)
                    {
                        string code = r["SubjectCode"].ToString()!;
                        double v = Convert.ToDouble(r["V"]);
                        if (v <= 0) continue;
                        if (!bag.TryGetValue(code, out var list)) bag[code] = list = new List<double>();
                        list.Add(v);
                    }
                    foreach (var kv in bag) Add(Math.Round(kv.Value.Average(), 2));
                    break;
            }
            return (Math.Round(sum, 2), cnt > 0 && failed == 0);
        }

        // Cache subject names looked up by code (avoids repeated SQL inside FillIndividual).
        private static Dictionary<string, string>? _subjectNameCache;
        private static string SubjectNameByCode(string code)
        {
            if (_subjectNameCache == null)
            {
                _subjectNameCache = new Dictionary<string, string>();
                foreach (DataRow r in DB.Query("SELECT SubjectCode, SubjectName FROM Subjects").Rows)
                    _subjectNameCache[r["SubjectCode"].ToString()!] = r["SubjectName"].ToString()!;
            }
            return _subjectNameCache.TryGetValue(code, out var n) ? n : code;
        }

        // ── Excel writer — clone template, fill cells, save. ──
        private static void BuildIndividualXlsx(string outPath, DataRow stu, string year,
            string title, IndScope scope, IndScoreBundle scores)
        {
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Templates", "ລີພອດຄະແນນບຸກຄົນ.xlsx");
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("ບໍ່ພົບແມ່ແບບ Templates/ລີພອດຄະແນນບຸກຄົນ.xlsx", templatePath);

            File.Copy(templatePath, outPath, overwrite: true);
            using var wb = new ClosedXML.Excel.XLWorkbook(outPath);
            var ws = wb.Worksheet(1);
            string laoFont = ResolveLaoFontName();

            // Page setup — fit to 1 page wide so column I (the score column) never
            // gets clipped off the right edge when Excel renders to PDF. The shipped
            // template has PagesWide=0 (no fit-to-width), which made column I + the
            // right side of the title disappear in A4 portrait.
            ws.PageSetup.PageOrientation = ClosedXML.Excel.XLPageOrientation.Portrait;
            ws.PageSetup.PaperSize       = ClosedXML.Excel.XLPaperSize.A4Paper;
            ws.PageSetup.PagesWide       = 1;
            ws.PageSetup.PagesTall       = 0;
            ws.PageSetup.Margins.Left    = 0.4;
            ws.PageSetup.Margins.Right   = 0.4;
            ws.PageSetup.Margins.Top     = 0.5;
            ws.PageSetup.Margins.Bottom  = 0.5;
            ws.PageSetup.CenterHorizontally = true;

            // Title — row 4, wherever the merged title-row cell starts. Older
            // versions of the template merged C4:E4 (title at C4); the current
            // school version merges A4:H4 (title at A4). Detect the merged
            // range whose first cell sits on row 4 so future header edits
            // (moving the title cell around) don't require code changes.
            ClosedXML.Excel.IXLCell? titleCell = null;
            foreach (var mr in ws.MergedRanges)
                if (mr.FirstRow().RowNumber() == 4) { titleCell = mr.FirstCell(); break; }
            titleCell ??= ws.Cell("A4");
            titleCell.Value = title;
            titleCell.Style.Font.FontName = laoFont;
            titleCell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

            // Name + class line — A5 (merged A5:G5). Name only (no StudentCode):
            // the school's report convention shows only the student's display name,
            // not their internal code. Centred horizontally to match the title.
            string nameLine = $"ຊື່  {stu["FullName"]}    ຊັ້ນ  {stu["GradeLevel"]}/{stu["ClassRoom"]}    ປີ  {year}";
            ws.Cell("A5").Value = nameLine;
            ws.Cell("A5").Style.Font.FontName = laoFont;
            ws.Cell("A5").Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

            // National-name + motto (A1, A2) — centre defensively in case the
            // template was edited and lost its centre alignment. A3 deliberately
            // skipped: it holds "ໂຮງຮຽນ ... ເລກທີ" with text padded right via
            // spaces, so centring would push "ເລກທີ" into the middle.
            ws.Cell("A1").Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            ws.Cell("A2").Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

            // Academic subject scores → column F rows 8..19.
            foreach (var (row, code) in TemplateSubjectRows)
            {
                var cell = ws.Cell(row, IndScoreCol);
                if (scores.AcademicByCode.TryGetValue(code, out var v) && v.HasValue)
                    cell.Value = Math.Round(v.Value, 2);
                else
                    cell.Value = "—";
                ApplyScoreCellStyle(cell, laoFont);
            }

            // Summary — F20 sum, F21 avg, F22 rank
            ws.Cell(IndSumRow,  IndScoreCol).Value = scores.Sum;
            ws.Cell(IndAvgRow,  IndScoreCol).Value = scores.Avg;
            ws.Cell(IndRankRow, IndScoreCol).Value = scores.RankLabel;   // number or "ຕົກ"
            ApplyScoreCellStyle(ws.Cell(IndSumRow,  IndScoreCol), laoFont);
            ApplyScoreCellStyle(ws.Cell(IndAvgRow,  IndScoreCol), laoFont);
            ApplyScoreCellStyle(ws.Cell(IndRankRow, IndScoreCol), laoFont);

            // Eval scores — F23 CHA1, F24 LAB1
            if (scores.Cha.HasValue) ws.Cell(IndChaRow, IndScoreCol).Value = Math.Round(scores.Cha.Value, 2);
            else                     ws.Cell(IndChaRow, IndScoreCol).Value = "—";
            if (scores.Lab.HasValue) ws.Cell(IndLabRow, IndScoreCol).Value = Math.Round(scores.Lab.Value, 2);
            else                     ws.Cell(IndLabRow, IndScoreCol).Value = "—";
            ApplyScoreCellStyle(ws.Cell(IndChaRow, IndScoreCol), laoFont);
            ApplyScoreCellStyle(ws.Cell(IndLabRow, IndScoreCol), laoFont);

            // Ensure all subject-name cells use the Lao font (the template author set this
            // already but a defensive pass keeps the output consistent across environments).
            // Column A is the label column in the v2 template.
            for (int r = 7; r <= IndLabRow; r++)
            {
                var nameCell = ws.Cell(r, 1); // column A
                if (!nameCell.IsEmpty()) nameCell.Style.Font.FontName = laoFont;
            }
            // Lao header / fixed-text cells
            foreach (string addr in new[] { "A1", "A2", "A4", "A5", "A7", "F7", "G7", "H7", "A27", "G27" })
                ws.Cell(addr).Style.Font.FontName = laoFont;

            wb.Save();
        }

        // ════════════════════════════════════════════════════════════
        //  PUBLIC ENTRY POINTS for ScoreHistoryPage windows
        //
        //  Both call the SAME template-fill code path that ReportPage's
        //  built-in reports use (BuildIndividualXlsx + ProduceTemplateReport),
        //  so the Excel/PDF output is byte-identical between the
        //  Reports page and the Score-History windows.
        //
        //  Historical-data contract: callers pass the historical
        //  AcademicYear + GradeLevel + ClassRoom — these helpers NEVER
        //  read the current Students.GradeLevel/ClassRoom for layout.
        //  CHA1/LAB1: shown as manual values in their dedicated rows /
        //  columns; excluded from every aggregate (template formulas
        //  SUM(D3:O3) / AVERAGE(D3:O3) ignore CHA/LAB which sit at U/V).
        // ════════════════════════════════════════════════════════════

        /// <summary>Thin re-export of <see cref="DB.SafeFileName"/> so views can
        /// reach it without an extra using directive.</summary>
        internal static string SafeFileName(string raw) => DB.SafeFileName(raw);

        /// <summary>
        /// Render ONE student × ONE month using the individual template
        /// (Templates/ລີພອດຄະແນນບຸກຄົນ.xlsx). The displayed grade/room come
        /// from the caller (historical values), NOT from the current Students row.
        /// </summary>
        internal static void RenderIndividualMonthlyXlsx(
            int sid, string year, int month,
            string historicalGrade, string historicalRoom,
            string outPath, out string title)
        {
            // Pull this student's name + code only. Grade/room come from the caller
            // (historical reconstruction) and are stuffed into a synthetic DataRow so
            // the existing BuildIndividualXlsx can be reused unchanged.
            var stuDt = DB.Query(@"SELECT StudentCode, FirstName||' '||LastName AS FullName,
                                          IFNULL(Status,'') AS Status
                                   FROM Students WHERE StudentID=@id",
                null, ("@id", sid));
            if (stuDt.Rows.Count == 0)
                throw new InvalidOperationException("ບໍ່ພົບຂໍ້ມູນນັກຮຽນ");

            var synth = new DataTable();
            synth.Columns.Add("StudentCode");
            synth.Columns.Add("FullName");
            synth.Columns.Add("GradeLevel");
            synth.Columns.Add("ClassRoom");
            synth.Columns.Add("AcademicYear");
            synth.Columns.Add("Status");
            var r = synth.NewRow();
            r["StudentCode"]  = stuDt.Rows[0]["StudentCode"];
            r["FullName"]     = stuDt.Rows[0]["FullName"];
            r["GradeLevel"]   = historicalGrade;
            r["ClassRoom"]    = historicalRoom;
            r["AcademicYear"] = year;
            r["Status"]       = stuDt.Rows[0]["Status"];
            synth.Rows.Add(r);

            var scope = new IndScope("M", month, $"ປະຈຳເດືອນ {month}");
            var scores = ComputeIndividualScores(sid, year, scope, r);
            title = $"ໃບຄະແນນປະຈຳເດືອນ {month}";
            BuildIndividualXlsx(outPath, r, year, title, scope, scores);
        }

        /// <summary>
        /// Render an ENTIRE historical class roster × ONE month using the class template
        /// (Templates/ໃບຄະແນນ.xlsx, Sheet 1). The roster is supplied by the caller
        /// (must contain StudentID, StudentCode, FullName columns) so the cohort can be
        /// the historical one from GetHistoricalClassRoster — graduated and promoted
        /// students are included as long as the caller put them in the roster.
        /// Rows are sorted by descending academic monthly sum before writing, so the
        /// template's RANK.EQ formula renders rows in rank order. CHA1/LAB1 columns
        /// (U, V) carry their raw monthly totals and are EXCLUDED from the SUM/AVG/RANK
        /// formulas by virtue of their column position.
        /// </summary>
        internal static void RenderClassMonthlyXlsx(
            string year, string grade, string room, int month,
            DataTable roster, string outPath, out string title)
        {
            int sem = DB.SemesterForMonth(month);

            // One-shot bulk lookup: per (student, subject) monthly total for everyone
            // in the historical cohort. CHA1/LAB1 are included — they're rendered
            // raw in the template's U/V columns and excluded from the SUM/AVG range.
            var idList = new List<string>();
            foreach (DataRow rr in roster.Rows) idList.Add(rr["StudentID"].ToString()!);
            string idCsv = idList.Count > 0 ? string.Join(",", idList) : "0";

            var dt = DB.Query($@"
                SELECT e.StudentID, sub.SubjectCode,
                       (IFNULL(ma.ActivityScore,0)+IFNULL(ma.DisciplineScore,0)+IFNULL(ma.HomeworkScore,0)) AS Total,
                       (ma.MonthlyID IS NOT NULL) AS HasRow
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@m
                WHERE e.AcademicYear=@y AND e.Semester=@sm
                  AND e.StudentID IN ({idCsv})",
                null, ("@m", month), ("@y", year), ("@sm", sem));

            var scoreMap = new Dictionary<(int, string), double>();
            foreach (DataRow rr in dt.Rows)
            {
                bool hasRow = Convert.ToInt32(rr["HasRow"]) != 0;
                if (!hasRow) continue;
                double t = rr["Total"] == DBNull.Value ? 0 : Convert.ToDouble(rr["Total"]);
                scoreMap[(Convert.ToInt32(rr["StudentID"]), rr["SubjectCode"].ToString()!)] = t;
            }
            double? Score(int sid, string code) =>
                scoreMap.TryGetValue((sid, code), out var v) ? v : (double?)null;

            // Sort roster by descending academic monthly sum (CHA/LAB excluded from sum).
            // Failed students (any academic <PassScore) sort by sum like everyone else;
            // their "ຕົກ" rank is rendered by the template's RANK.EQ-with-COUNTIF formula.
            var academicCodes = new HashSet<string>();
            foreach (DataRow rr in DB.Query("SELECT SubjectCode FROM Subjects WHERE SubjectCode NOT IN ('CHA1','LAB1')").Rows)
                academicCodes.Add(rr["SubjectCode"].ToString()!);
            var sumBySid = new Dictionary<int, double>();
            foreach (DataRow rr in roster.Rows)
            {
                int sidv = Convert.ToInt32(rr["StudentID"]);
                double sum = 0;
                foreach (var c in academicCodes)
                    if (scoreMap.TryGetValue((sidv, c), out var v)) sum += v;
                sumBySid[sidv] = sum;
            }
            var sortedRoster = roster.Clone();
            foreach (DataRow rr in roster.Select("", "")
                .OrderByDescending(rw => sumBySid.TryGetValue(Convert.ToInt32(rw["StudentID"]), out var s) ? s : 0))
                sortedRoster.ImportRow(rr);

            string gradeLabel = grade.Replace(".", " ");
            string yearLabel  = year.Replace("-", " - ");
            title = $"ສະຫຼຸບຄະແນນປະຈຳເດືອນ {month} ຫ້ອງ {gradeLabel}/{room} ສົກຮຽນ {yearLabel}";

            string tpl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ໃບຄະແນນ.xlsx");
            if (!File.Exists(tpl))
                throw new FileNotFoundException("ບໍ່ພົບແມ່ແບບ Templates/ໃບຄະແນນ.xlsx", tpl);

            // ProduceTemplateReport keeps Sheet 1 (monthly), deletes the others, fills
            // the surviving sheet via FillSheet — which preserves the template's row-3
            // formulas (SUM, AVERAGE, RANK.EQ, COUNTIF<5, IF) and shifts their refs
            // per row. CHA/LAB columns are written but stay outside the formula range.
            File.Copy(tpl, outPath, overwrite: true);
            using var wb = new ClosedXML.Excel.XLWorkbook(outPath);
            var sheets = wb.Worksheets.ToList();
            for (int i = 0; i < sheets.Count; i++)
                if (i != 0) sheets[i].Delete();
            // Build academic-only map (CHA1/LAB1 already excluded by scoreMap construction).
            var academicForRank = new Dictionary<(int, string), double>();
            foreach (var kv in scoreMap)
                if (academicCodes.Contains(kv.Key.Item2)) academicForRank[kv.Key] = kv.Value;
            FillSheet(wb.Worksheet(1), title, sortedRoster, Score);
            WriteClassRanks(wb.Worksheet(1), sortedRoster, academicForRank);
            wb.Save();
        }

        /// <summary>
        /// Render ONE student × ONE semester's totals using the individual template.
        /// Identical code path to the Reports page's Individual report with scope=S.
        /// CHA1/LAB1 read from EvaluationScores(SEM1/SEM2), academic from Scores.TotalScore.
        /// </summary>
        internal static void RenderIndividualSemesterXlsx(
            int sid, string year, int sem,
            string historicalGrade, string historicalRoom,
            string outPath, out string title)
        {
            var stuDt = DB.Query(@"SELECT StudentCode, FirstName||' '||LastName AS FullName,
                                          IFNULL(Status,'') AS Status
                                   FROM Students WHERE StudentID=@id",
                null, ("@id", sid));
            if (stuDt.Rows.Count == 0)
                throw new InvalidOperationException("ບໍ່ພົບຂໍ້ມູນນັກຮຽນ");

            var synth = NewSynthStudentRow(stuDt.Rows[0], year, historicalGrade, historicalRoom);
            var scope  = new IndScope("S", sem, $"ສະຫຼຸບພາກຮຽນ {sem}");
            var scores = ComputeIndividualScores(sid, year, scope, synth);
            title = $"ໃບຄະແນນສະຫຼຸບພາກຮຽນ {sem}";
            BuildIndividualXlsx(outPath, synth, year, title, scope, scores);
        }

        /// <summary>
        /// Render ONE student × ANNUAL totals using the individual template.
        /// Academic per-subject value = mean of Sem1 and Sem2 TotalScore.
        /// CHA1/LAB1 read from EvaluationScores(ANNUAL).
        /// </summary>
        internal static void RenderIndividualAnnualXlsx(
            int sid, string year,
            string historicalGrade, string historicalRoom,
            string outPath, out string title)
        {
            var stuDt = DB.Query(@"SELECT StudentCode, FirstName||' '||LastName AS FullName,
                                          IFNULL(Status,'') AS Status
                                   FROM Students WHERE StudentID=@id",
                null, ("@id", sid));
            if (stuDt.Rows.Count == 0)
                throw new InvalidOperationException("ບໍ່ພົບຂໍ້ມູນນັກຮຽນ");

            var synth = NewSynthStudentRow(stuDt.Rows[0], year, historicalGrade, historicalRoom);
            var scope  = new IndScope("A", 0, "ສະຫຼຸບປະຈຳປີ");
            var scores = ComputeIndividualScores(sid, year, scope, synth);
            title = "ໃບຄະແນນປະຈຳປີ";
            BuildIndividualXlsx(outPath, synth, year, title, scope, scores);
        }

        /// <summary>Synthesize a one-row DataTable matching BuildIndividualXlsx's
        /// expected stu DataRow shape — but with HISTORICAL grade/room substituted
        /// in for accurate name-line rendering on past years.</summary>
        private static DataRow NewSynthStudentRow(DataRow src, string year,
            string historicalGrade, string historicalRoom)
        {
            var synth = new DataTable();
            synth.Columns.Add("StudentCode");
            synth.Columns.Add("FullName");
            synth.Columns.Add("GradeLevel");
            synth.Columns.Add("ClassRoom");
            synth.Columns.Add("AcademicYear");
            synth.Columns.Add("Status");
            var r = synth.NewRow();
            r["StudentCode"]  = src["StudentCode"];
            r["FullName"]     = src["FullName"];
            r["GradeLevel"]   = historicalGrade;
            r["ClassRoom"]    = historicalRoom;
            r["AcademicYear"] = year;
            r["Status"]       = src["Status"];
            synth.Rows.Add(r);
            return r;
        }

        /// <summary>
        /// Render an entire historical cohort × ONE semester using Sheet 2 / Sheet 3
        /// of ໃບຄະແນນ.xlsx. Academic scores come from Scores.TotalScore
        /// (Mid×MidPct% + Final×FinalPct%). CHA1/LAB1 from EvaluationScores(SEM1/SEM2)
        /// — they sit in the template's W/X columns, outside the formula range.
        /// Rows pre-sorted by descending academic sum so the template's RANK.EQ
        /// formula renders rows in rank order.
        /// </summary>
        internal static void RenderClassSemesterXlsx(
            string year, string grade, string room, int sem,
            DataTable roster, string outPath, out string title)
        {
            if (sem != 1 && sem != 2)
                throw new ArgumentException("sem must be 1 or 2", nameof(sem));

            var idList = new List<string>();
            foreach (DataRow rr in roster.Rows) idList.Add(rr["StudentID"].ToString()!);
            string idCsv = idList.Count > 0 ? string.Join(",", idList) : "0";

            // Academic per (student, subject) — CHA1/LAB1 excluded by the join filter.
            var acDt = DB.Query($@"
                SELECT e.StudentID, sub.SubjectCode, IFNULL(sc.TotalScore,0) AS Total
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE e.AcademicYear=@y AND e.Semester=@sm
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                  AND e.StudentID IN ({idCsv})",
                null, ("@y", year), ("@sm", sem));
            var academic = new Dictionary<(int, string), double>();
            foreach (DataRow r in acDt.Rows)
            {
                double t = r["Total"] == DBNull.Value ? 0 : Convert.ToDouble(r["Total"]);
                if (t > 0) academic[(Convert.ToInt32(r["StudentID"]), r["SubjectCode"].ToString()!)] = t;
            }

            // CHA1/LAB1 manual scores for this context.
            string ctx = sem == 1 ? "SEM1" : "SEM2";
            var evDt = DB.Query($@"
                SELECT StudentID, SubjectCode, Score
                FROM EvaluationScores
                WHERE AcademicYear=@y AND Context=@c
                  AND SubjectCode IN ('CHA1','LAB1')
                  AND StudentID IN ({idCsv})",
                null, ("@y", year), ("@c", ctx));
            var evalMap = new Dictionary<(int, string), double>();
            foreach (DataRow r in evDt.Rows)
            {
                if (r["Score"] == DBNull.Value) continue;
                evalMap[(Convert.ToInt32(r["StudentID"]), r["SubjectCode"].ToString()!)]
                    = Convert.ToDouble(r["Score"]);
            }

            double? Score(int sid, string code) =>
                (code == "CHA1" || code == "LAB1")
                    ? (evalMap.TryGetValue((sid, code), out var ev) ? ev : (double?)null)
                    : (academic.TryGetValue((sid, code), out var v) ? v : (double?)null);

            // Sort by descending academic sum so RANK.EQ renders rows top-down.
            var sortedRoster = SortRosterByAcademicSum(roster, academic);

            string gradeLabel = grade.Replace(".", " ");
            string yearLabel  = year.Replace("-", " - ");
            title = $"ສະຫຼຸບສະເລ່ຍພາກຮຽນ {sem} ຫ້ອງ {gradeLabel}/{room} ສົກຮຽນ {yearLabel}";

            string tpl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ໃບຄະແນນ.xlsx");
            if (!File.Exists(tpl))
                throw new FileNotFoundException("ບໍ່ພົບແມ່ແບບ Templates/ໃບຄະແນນ.xlsx", tpl);

            // Sheet 2 = Sem 1 layout; Sheet 3 = Sem 2 layout.
            int keepIdx = sem == 1 ? 2 : 3;
            File.Copy(tpl, outPath, overwrite: true);
            using var wb = new ClosedXML.Excel.XLWorkbook(outPath);
            var sheets = wb.Worksheets.ToList();
            for (int i = 0; i < sheets.Count; i++)
                if (i + 1 != keepIdx) sheets[i].Delete();
            FillSheet(wb.Worksheet(1), title, sortedRoster, Score);
            WriteClassRanks(wb.Worksheet(1), sortedRoster, academic);
            wb.Save();
        }

        /// <summary>
        /// Render an entire historical cohort × ANNUAL totals using Sheet 4 of
        /// ໃບຄະແນນ.xlsx. Per-subject academic = mean of Sem1 + Sem2 TotalScore;
        /// CHA1/LAB1 from EvaluationScores(ANNUAL). Sort by descending sum.
        /// </summary>
        internal static void RenderClassAnnualXlsx(
            string year, string grade, string room,
            DataTable roster, string outPath, out string title)
        {
            var idList = new List<string>();
            foreach (DataRow rr in roster.Rows) idList.Add(rr["StudentID"].ToString()!);
            string idCsv = idList.Count > 0 ? string.Join(",", idList) : "0";

            // Both semesters' academic scores — collect, then average per (student, subject).
            var acDt = DB.Query($@"
                SELECT e.StudentID, sub.SubjectCode, IFNULL(sc.TotalScore,0) AS Total
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE e.AcademicYear=@y
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                  AND sc.ScoreID IS NOT NULL
                  AND e.StudentID IN ({idCsv})",
                null, ("@y", year));
            var bag = new Dictionary<(int, string), List<double>>();
            foreach (DataRow r in acDt.Rows)
            {
                double t = r["Total"] == DBNull.Value ? 0 : Convert.ToDouble(r["Total"]);
                if (t <= 0) continue;
                var k = (Convert.ToInt32(r["StudentID"]), r["SubjectCode"].ToString()!);
                if (!bag.TryGetValue(k, out var list)) bag[k] = list = new List<double>();
                list.Add(t);
            }
            var academic = new Dictionary<(int, string), double>();
            foreach (var kv in bag) academic[kv.Key] = Math.Round(kv.Value.Average(), 2);

            // CHA1/LAB1 annual manual scores.
            var evDt = DB.Query($@"
                SELECT StudentID, SubjectCode, Score
                FROM EvaluationScores
                WHERE AcademicYear=@y AND Context='ANNUAL'
                  AND SubjectCode IN ('CHA1','LAB1')
                  AND StudentID IN ({idCsv})",
                null, ("@y", year));
            var evalMap = new Dictionary<(int, string), double>();
            foreach (DataRow r in evDt.Rows)
            {
                if (r["Score"] == DBNull.Value) continue;
                evalMap[(Convert.ToInt32(r["StudentID"]), r["SubjectCode"].ToString()!)]
                    = Convert.ToDouble(r["Score"]);
            }

            double? Score(int sid, string code) =>
                (code == "CHA1" || code == "LAB1")
                    ? (evalMap.TryGetValue((sid, code), out var ev) ? ev : (double?)null)
                    : (academic.TryGetValue((sid, code), out var v) ? v : (double?)null);

            var sortedRoster = SortRosterByAcademicSum(roster, academic);

            string gradeLabel = grade.Replace(".", " ");
            string yearLabel  = year.Replace("-", " - ");
            title = $"ສະຫຼຸບຄະແນນປະຈຳປີ ຫ້ອງ {gradeLabel}/{room} ສົກຮຽນ {yearLabel}";

            string tpl = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ໃບຄະແນນ.xlsx");
            if (!File.Exists(tpl))
                throw new FileNotFoundException("ບໍ່ພົບແມ່ແບບ Templates/ໃບຄະແນນ.xlsx", tpl);

            // Sheet 4 = Annual layout.
            File.Copy(tpl, outPath, overwrite: true);
            using var wb = new ClosedXML.Excel.XLWorkbook(outPath);
            var sheets = wb.Worksheets.ToList();
            for (int i = 0; i < sheets.Count; i++)
                if (i + 1 != 4) sheets[i].Delete();
            FillSheet(wb.Worksheet(1), title, sortedRoster, Score);
            WriteClassRanks(wb.Worksheet(1), sortedRoster, academic);
            wb.Save();
        }

        // Sort a roster DataTable into a new DataTable ordered by descending academic
        // sum (CHA/LAB excluded). Used by every class report so the template's
        // RANK.EQ formula renders rows in rank order.
        private static DataTable SortRosterByAcademicSum(DataTable roster,
            Dictionary<(int, string), double> academic)
        {
            var sumBySid = new Dictionary<int, double>();
            foreach (DataRow rr in roster.Rows)
            {
                int sidv = Convert.ToInt32(rr["StudentID"]);
                double sum = 0;
                foreach (var kv in academic)
                    if (kv.Key.Item1 == sidv) sum += kv.Value;
                sumBySid[sidv] = sum;
            }
            var sorted = roster.Clone();
            foreach (DataRow rr in roster.Select("", "")
                .OrderByDescending(rw => sumBySid.TryGetValue(Convert.ToInt32(rw["StudentID"]), out var s) ? s : 0))
                sorted.ImportRow(rr);
            return sorted;
        }

        // Post-pass: compute per-student rank and write VALUES into the template's
        // rank column, overriding its RANK.EQ formula. This guarantees the rank
        // displays even if Excel's RANK.EQ evaluation behaves differently across
        // viewers / versions, and avoids the "blank column" the user reported.
        // Tie-aware 1224 ranking among passing students; failed students → "ຕົກ".
        // The rest of the template's formulas (SUM, AVERAGE, COUNTIF, IF) are left
        // intact. ‘roster’ MUST be in the same order that FillSheet wrote to the
        // worksheet (the sorted roster from SortRosterByAcademicSum).
        private static void WriteClassRanks(ClosedXML.Excel.IXLWorksheet ws,
            DataTable roster, Dictionary<(int, string), double> academic)
        {
            // Locate the rank column by header substring in row 2.
            int rankCol = 0;
            int maxCol  = ws.LastColumnUsed()?.ColumnNumber() ?? 24;
            for (int c = 1; c <= maxCol; c++)
            {
                string h = ws.Cell(2, c).GetString().Trim();
                if (h.Contains("ອັນດັບ")) { rankCol = c; break; }
            }
            if (rankCol == 0) return;   // template has no rank column — nothing to do

            // Compute per-student academic sum + pass flag.
            var sumBySid = new Dictionary<int, double>();
            var failBySid = new Dictionary<int, bool>();
            foreach (DataRow r in roster.Rows)
            {
                int sid = Convert.ToInt32(r["StudentID"]);
                double sum = 0; bool failed = false;
                foreach (var kv in academic)
                {
                    if (kv.Key.Item1 != sid) continue;
                    sum += kv.Value;
                    if (kv.Value < DB.PassScore) failed = true;
                }
                sumBySid[sid]  = sum;
                failBySid[sid] = failed;
            }

            // Tie-aware 1224 rank among passing students (descending sum).
            var passingOrdered = sumBySid
                .Where(kv => !failBySid[kv.Key])
                .OrderByDescending(kv => kv.Value)
                .ToList();
            var rankBySid = new Dictionary<int, string>();
            int currentRank = 0; double prev = double.NaN;
            for (int i = 0; i < passingOrdered.Count; i++)
            {
                if (i == 0 || Math.Abs(passingOrdered[i].Value - prev) > 0.0001)
                {
                    currentRank = i + 1;
                    prev = passingOrdered[i].Value;
                }
                rankBySid[passingOrdered[i].Key] = currentRank.ToString();
            }
            foreach (var kv in failBySid)
                if (kv.Value) rankBySid[kv.Key] = "ຕົກ";

            // Write rank values, overriding the row's RANK.EQ formula. Row order
            // in `roster` matches FillSheet's write order.
            const int dataStart = 3;
            int idx = 0;
            foreach (DataRow r in roster.Rows)
            {
                int sid = Convert.ToInt32(r["StudentID"]);
                if (rankBySid.TryGetValue(sid, out var rankStr))
                {
                    var cell = ws.Cell(dataStart + idx, rankCol);
                    cell.Clear(ClosedXML.Excel.XLClearOptions.Contents);
                    cell.Value = rankStr;
                    cell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                }
                idx++;
            }
        }

        private static void ApplyScoreCellStyle(ClosedXML.Excel.IXLCell cell, string laoFont)
        {
            string probe = cell.HasFormula ? cell.FormulaA1 : cell.GetString();
            cell.Style.Font.FontName = ContainsLao(probe) ? laoFont : "Times New Roman";
            cell.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
            cell.Style.Alignment.Vertical   = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
            var thin = ClosedXML.Excel.XLBorderStyleValues.Thin;
            cell.Style.Border.TopBorder    = thin;
            cell.Style.Border.BottomBorder = thin;
            cell.Style.Border.LeftBorder   = thin;
            cell.Style.Border.RightBorder  = thin;
        }

        private void GenEnrollmentAgreement()
        {
            if (CmbStudent.SelectedValue == null)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນ", "ບໍ່ມີຂໍ້ມູນ");
                return;
            }
            int sid = Convert.ToInt32(CmbStudent.SelectedValue);
            var sr = DB.Query(@"SELECT * FROM Students WHERE StudentID=@id", null, ("@id", sid));
            if (sr.Rows.Count == 0)
            {
                MessageBox.Show("ບໍ່ພົບຂໍ້ມູນນັກຮຽນ", "ຜິດພາດ");
                return;
            }
            var row = sr.Rows[0];
            string Get(string col) =>
                sr.Columns.Contains(col) && row[col] != DBNull.Value ? row[col].ToString()!.Trim() : "";
            // Empty fields render as a dotted blank, matching the printed form.
            string Fb(string v) => string.IsNullOrEmpty(v) ? "...................." : v;

            // Template-file location: shipped next to the EXE under Templates\.
            string templatePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Templates", "ໃບສັນຍາ.docx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show(
                    "ບໍ່ພົບແມ່ແບບ (template):\n" + templatePath +
                    "\n\nກະລຸນາວາງໄຟລ໌ ‘ໃບສັນຍາ.docx’ ໄວ້ໃນໂຟນເດີ Templates ຂ້າງໆ EXE",
                    "ບໍ່ມີແມ່ແບບ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Two output options. PDF generation works by:
            //   1) Filling the .docx template (always),
            //   2) Then converting → PDF via Word COM (if Office installed) or LibreOffice.
            var dlg = new SaveFileDialog
            {
                Filter   = "Word|*.docx|PDF|*.pdf",
                FileName = $"agreement_{Get("StudentCode")}_{DateTime.Now:yyyyMMdd}.docx"
            };
            if (dlg.ShowDialog() != true) return;
            bool wantPdf = dlg.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
            // When the user wants PDF, fill into a temp .docx then convert.
            string docxPath = wantPdf
                ? Path.Combine(Path.GetTempPath(), $"sis_agreement_{Guid.NewGuid():N}.docx")
                : dlg.FileName;

            // Build the token → replacement map.
            string fatherName = Get("FatherName");
            if (string.IsNullOrEmpty(fatherName)) fatherName = Get("ParentName");
            var tokens = new Dictionary<string, string>
            {
                ["{{SCHOOL_NAME}}"]       = Fb(DB.SchoolName),
                ["{{FULL_NAME}}"]         = Fb($"{Get("FirstName")} {Get("LastName")}".Trim()),
                ["{{BIRTH_DATE}}"]        = Fb(Get("BirthDate")),
                ["{{BIRTH_VILLAGE}}"]     = Fb(Get("BirthVillage")),
                ["{{BIRTH_DISTRICT}}"]    = Fb(Get("BirthDistrict")),
                ["{{BIRTH_PROVINCE}}"]    = Fb(Get("BirthProvince")),
                ["{{VILLAGE}}"]           = Fb(Get("Village")),
                ["{{DISTRICT}}"]          = Fb(Get("District")),
                ["{{PROVINCE}}"]          = Fb(Get("Province")),
                ["{{FATHER_NAME}}"]       = Fb(fatherName),
                ["{{FATHER_AGE}}"]        = Fb(Get("FatherAge")),
                ["{{FATHER_JOB}}"]        = Fb(Get("FatherJob")),
                ["{{FATHER_VILLAGE}}"]    = Fb(Get("FatherVillage")),
                ["{{FATHER_DISTRICT}}"]   = Fb(Get("FatherDistrict")),
                ["{{FATHER_PROVINCE}}"]   = Fb(Get("FatherProvince")),
                ["{{MOTHER_NAME}}"]       = Fb(Get("MotherName")),
                ["{{MOTHER_AGE}}"]        = Fb(Get("MotherAge")),
                ["{{MOTHER_JOB}}"]        = Fb(Get("MotherJob")),
                ["{{MOTHER_VILLAGE}}"]    = Fb(Get("MotherVillage")),
                ["{{MOTHER_DISTRICT}}"]   = Fb(Get("MotherDistrict")),
                ["{{MOTHER_PROVINCE}}"]   = Fb(Get("MotherProvince")),
                ["{{ACADEMIC_YEAR}}"]     = Fb(Get("AcademicYear")),
                ["{{GRADE}}"]             = Fb(Get("GradeLevel")),
            };

            try
            {
                // 1) Copy template to working .docx path (preserves all formatting).
                File.Copy(templatePath, docxPath, overwrite: true);

                // 2) Replace tokens — uses paragraph-level concat-and-replace, so it
                //    handles tokens Word may have split across runs after edits.
                FillDocxTokens(docxPath, tokens);

                // 3) If user picked .pdf, convert the filled .docx → PDF and delete the temp.
                if (wantPdf)
                {
                    if (!ConvertDocxToPdf(docxPath, dlg.FileName, out string convertError))
                    {
                        MessageBox.Show(
                            "ບໍ່ສາມາດແປງເປັນ PDF ໄດ້:\n" + convertError +
                            "\n\nຕ້ອງມີ Microsoft Word ຫຼື LibreOffice ຕິດຕັ້ງໄວ້.\n" +
                            "ໄຟລ໌ Word ຍັງຄົງຢູ່ທີ່:\n" + docxPath,
                            "ບໍ່ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    try { File.Delete(docxPath); } catch { /* keep silent — non-critical */ }
                }

                DB.Log("EnrollmentAgreement", $"{Get("StudentCode")} → {Path.GetFileName(dlg.FileName)}");
                MessageBox.Show($"ສ້າງໃບສັນຍາສຳເລັດ:\n{dlg.FileName}", "ສຳເລັດ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ສ້າງໃບສັນຍາບໍ່ສຳເລັດ:\n{ex.Message}", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Replace {{TOKEN}} placeholders in a .docx file.
        // Robust to Word splitting tokens across multiple <w:t> runs (a common
        // editing side-effect — e.g. "{{FATHER_NAME}}" stored as "{{"+"FATHER_NAME}}").
        // We reassemble each paragraph's text, do the replacement on the combined
        // string, write the result into the FIRST run, then empty the rest.
        // All runs that contain or surround a token share the same formatting in our
        // templates (verified at build time), so the merge is visually lossless.
        private static void FillDocxTokens(string docxPath, Dictionary<string, string> tokens)
        {
            using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(docxPath, true);
            var parts = new List<DocumentFormat.OpenXml.Packaging.OpenXmlPart>();
            if (doc.MainDocumentPart != null)
            {
                parts.Add(doc.MainDocumentPart);
                foreach (var hp in doc.MainDocumentPart.HeaderParts) parts.Add(hp);
                foreach (var fp in doc.MainDocumentPart.FooterParts) parts.Add(fp);
            }
            foreach (var part in parts)
            {
                if (part.RootElement == null) continue;
                foreach (var para in part.RootElement
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    var texts = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                    if (texts.Count == 0) continue;
                    string combined = string.Concat(texts.Select(t => t.Text ?? ""));
                    bool anyHit = false;
                    foreach (var kv in tokens) if (combined.Contains(kv.Key)) { anyHit = true; break; }
                    if (!anyHit) continue;
                    string replaced = combined;
                    foreach (var kv in tokens) replaced = replaced.Replace(kv.Key, kv.Value);
                    if (replaced == combined) continue;
                    // Write the result into the first text element, clear the rest.
                    texts[0].Text  = replaced;
                    texts[0].Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
                    for (int i = 1; i < texts.Count; i++) texts[i].Text = "";
                }
                part.RootElement.Save();
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  REPORT 6 — Student Profile / History (ລາຍງານປະຫວັດນັກຮຽນ)
        //  Template: Templates/ລາຍງານປະຫວັດນັກຮຽນ.docx
        //  Uses the same Word/PDF save-dialog flow as the Enrollment Agreement,
        //  but with a different template + token map.
        // ═══════════════════════════════════════════════════════════
        private void GenStudentProfileReport()
        {
            if (CmbStudent.SelectedValue == null)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            int sid = Convert.ToInt32(CmbStudent.SelectedValue);
            var sr = DB.Query(@"SELECT * FROM Students WHERE StudentID=@id", null, ("@id", sid));
            if (sr.Rows.Count == 0)
            {
                MessageBox.Show("ບໍ່ພົບຂໍ້ມູນນັກຮຽນ", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var row = sr.Rows[0];
            string Get(string col) =>
                sr.Columns.Contains(col) && row[col] != DBNull.Value ? row[col].ToString()!.Trim() : "";
            string Fb(string v) => string.IsNullOrEmpty(v) ? "...................." : v;

            string templatePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Templates", "ລາຍງານປະຫວັດນັກຮຽນ.docx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show(
                    "ບໍ່ພົບແມ່ແບບ:\n" + templatePath +
                    "\n\nກະລຸນາວາງໄຟລ໌ ‘ລາຍງານປະຫວັດນັກຮຽນ.docx’ ໄວ້ໃນໂຟນເດີ Templates ຂ້າງໆ EXE",
                    "ບໍ່ມີແມ່ແບບ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter   = "Word|*.docx|PDF|*.pdf",
                FileName = $"profile_{Get("StudentCode")}_{DateTime.Now:yyyyMMdd}.docx"
            };
            if (dlg.ShowDialog() != true) return;
            bool wantPdf = dlg.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
            string docxPath = wantPdf
                ? Path.Combine(Path.GetTempPath(), $"sis_profile_{Guid.NewGuid():N}.docx")
                : dlg.FileName;

            // Build the token → replacement map. Father/Mother names fall back to the
            // legacy ParentName column if the modern split columns are empty.
            string fatherName = Get("FatherName");
            if (string.IsNullOrEmpty(fatherName)) fatherName = Get("ParentName");
            var tokens = new Dictionary<string, string>
            {
                ["{{FULL_NAME}}"]         = Fb($"{Get("FirstName")} {Get("LastName")}".Trim()),
                ["{{BIRTH_DATE}}"]        = Fb(Get("BirthDate")),
                ["{{BIRTH_VILLAGE}}"]     = Fb(Get("BirthVillage")),
                ["{{BIRTH_DISTRICT}}"]    = Fb(Get("BirthDistrict")),
                ["{{BIRTH_PROVINCE}}"]    = Fb(Get("BirthProvince")),
                ["{{VILLAGE}}"]           = Fb(Get("Village")),
                ["{{DISTRICT}}"]          = Fb(Get("District")),
                ["{{PROVINCE}}"]          = Fb(Get("Province")),
                ["{{FATHER_NAME}}"]       = Fb(fatherName),
                ["{{FATHER_AGE}}"]        = Fb(Get("FatherAge")),
                ["{{FATHER_JOB}}"]        = Fb(Get("FatherJob")),
                ["{{FATHER_VILLAGE}}"]    = Fb(Get("FatherVillage")),
                ["{{FATHER_DISTRICT}}"]   = Fb(Get("FatherDistrict")),
                ["{{FATHER_PROVINCE}}"]   = Fb(Get("FatherProvince")),
                ["{{MOTHER_NAME}}"]       = Fb(Get("MotherName")),
                ["{{MOTHER_AGE}}"]        = Fb(Get("MotherAge")),
                ["{{MOTHER_JOB}}"]        = Fb(Get("MotherJob")),
                ["{{MOTHER_VILLAGE}}"]    = Fb(Get("MotherVillage")),
                ["{{MOTHER_DISTRICT}}"]   = Fb(Get("MotherDistrict")),
                ["{{MOTHER_PROVINCE}}"]   = Fb(Get("MotherProvince")),
            };

            try
            {
                File.Copy(templatePath, docxPath, overwrite: true);
                FillDocxTokens(docxPath, tokens);

                if (wantPdf)
                {
                    if (!ConvertDocxToPdf(docxPath, dlg.FileName, out string convertError))
                    {
                        MessageBox.Show(
                            "ບໍ່ສາມາດແປງເປັນ PDF ໄດ້:\n" + convertError +
                            "\n\nຕ້ອງມີ Microsoft Word ຫຼື LibreOffice ຕິດຕັ້ງໄວ້.\n" +
                            "ໄຟລ໌ Word ຍັງຄົງຢູ່ທີ່:\n" + docxPath,
                            "ບໍ່ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    try { File.Delete(docxPath); } catch { /* non-critical */ }
                }

                DB.Log("StudentProfile", $"{Get("StudentCode")} → {Path.GetFileName(dlg.FileName)}");
                MessageBox.Show($"ສ້າງລາຍງານປະຫວັດສຳເລັດ:\n{dlg.FileName}", "ສຳເລັດ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ສ້າງລາຍງານປະຫວັດບໍ່ສຳເລັດ:\n{ex.Message}", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── .docx → .pdf converter ─────────────────────────────────
        // Tries Microsoft Word via COM late-binding first (most schools have Office),
        // then falls back to LibreOffice (soffice.exe) if Word isn't installed.
        // No NuGet dependency: COM via Type.GetTypeFromProgID + dynamic; LibreOffice via Process.
        private static bool ConvertDocxToPdf(string docxPath, string pdfPath, out string error)
        {
            error = "";

            // Attempt 1 — Microsoft Word COM (Word.Application)
            try
            {
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType != null)
                {
                    dynamic? word = Activator.CreateInstance(wordType);
                    if (word != null)
                    {
                        word.Visible = false;
                        word.DisplayAlerts = 0; // wdAlertsNone
                        dynamic doc = word.Documents.Open(docxPath, ReadOnly: true, Visible: false);
                        // wdFormatPDF = 17
                        doc.SaveAs2(pdfPath, FileFormat: 17);
                        doc.Close(SaveChanges: false);
                        word.Quit();
                        return File.Exists(pdfPath);
                    }
                }
            }
            catch (Exception ex)
            {
                error = "Word COM: " + ex.Message;
            }

            // Attempt 2 — LibreOffice (soffice.exe) headless conversion
            string[] candidates = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };
            foreach (var soffice in candidates)
            {
                if (!File.Exists(soffice)) continue;
                try
                {
                    string outDir = Path.GetDirectoryName(pdfPath) ?? Path.GetTempPath();
                    var psi = new System.Diagnostics.ProcessStartInfo(soffice,
                        $"--headless --convert-to pdf --outdir \"{outDir}\" \"{docxPath}\"")
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true
                    };
                    using var p = System.Diagnostics.Process.Start(psi);
                    if (p == null) { error = "LibreOffice: failed to start"; continue; }
                    p.WaitForExit(60_000);
                    // LibreOffice names the output after the input file's base name.
                    string produced = Path.Combine(outDir,
                        Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
                    if (File.Exists(produced))
                    {
                        if (produced != pdfPath)
                        {
                            File.Copy(produced, pdfPath, overwrite: true);
                            try { File.Delete(produced); } catch { }
                        }
                        return true;
                    }
                    error = "LibreOffice: no output produced";
                }
                catch (Exception ex)
                {
                    error = "LibreOffice: " + ex.Message;
                }
            }

            if (string.IsNullOrEmpty(error))
                error = "Word ຫຼື LibreOffice ບໍ່ໄດ້ຕິດຕັ້ງ";
            return false;
        }
    }
}
