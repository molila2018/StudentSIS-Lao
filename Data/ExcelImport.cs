using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  EXCEL SCORE IMPORT — single source of truth for both pages
    //
    //  PER-SUBJECT workflow: teachers pick (Year, Grade, Room,
    //  Subject) plus (Month or Semester), download a template that
    //  contains ONLY that one subject's roster, fill scores, then
    //  re-upload. Templates are styled to match the official Lao
    //  monthly-score sheet: Phetsarath OT font throughout, gray
    //  header fills, thin borders. The parser refuses any workbook
    //  whose hidden metadata cells don't match — random external
    //  xlsx files cannot be imported, and a template generated for
    //  one (subject/year/grade/room/month-or-sem) tuple cannot be
    //  uploaded under another.
    //
    //  Layout — ROW 1 title (Phetsarath OT 14pt bold, no merge);
    //  ROW 2 blank; ROW 3 headers (Phetsarath OT 11pt bold, light
    //  gray fill, thin border, centered); ROW 4+ data (Phetsarath
    //  OT 11pt, thin border).
    //
    //    Monthly + academic subject:
    //      A=ລຳດັບ · B=ລະຫັດນັກຮຽນ · C=ຊື່ນັກຮຽນ ·
    //      D=ກິດຈະກຳ2 · E=ຮ່ວມຮຽນ3 · F=ກວດກາ5
    //      (per-column max 2/3/5 matches MonthlyAssessments columns
    //      directly — no split, no merging into a single /10 cell)
    //    Monthly + CHA1/LAB1 (eval-only):
    //      A=ລຳດັບ · B=ລະຫັດນັກຮຽນ · C=ຊື່ນັກຮຽນ · D=ຄະແນນ (/10)
    //      (EvaluationScores stores a single /10 value — no sub-scores)
    //    Semester (any subject):
    //      A=ລຳດັບ · B=ລະຫັດ · C=ຊື່ · D=ສະເລ່ຍປະຈຳເດືອນ · E=ສອບເສງ
    //
    //  Hidden column Z carries metadata for round-trip validation:
    //    Z1 = magic ("SIS_LAO_MONTHLY_v3" or "SIS_LAO_SEMESTER_v3")
    //    Z2 = AcademicYear · Z3 = Grade · Z4 = Room ·
    //    Z5 = Month|Sem  · Z6 = SubjectCode
    //
    //  Score-routing rules (mirrors the manual entry pages):
    //    Academic + Monthly  → MonthlyAssessments (a/d/h written direct)
    //    Academic + Semester → Scores (MidScore + FinalScore)
    //    CHA1/LAB1           → EvaluationScores (Context: Month{N} / SEM{N})
    //
    //  CHA1/LAB1 are NEVER averaged or auto-derived. RecomputeMid-
    //  FromMonthly only fires for academic monthly enrollments after
    //  the import commits, so Scores.MidScore stays in sync with the
    //  new totals.
    // ════════════════════════════════════════════════════════════

    public enum ImportKind { Monthly, Semester }

    public class ImportRow
    {
        // Every data member is a PROPERTY (not a field) because
        // ImportPreviewWindow binds them into a WPF DataGrid — Binding uses
        // reflection over properties only, so fields silently render blank.
        // That was the exact symptom: RowNo / StudentName / Discipline /
        // Activity / Homework all appeared empty in the preview even though
        // the parser had populated them correctly.
        public int    RowNo       { get; set; }   // 1-based ordinal in the parsed list
        public int    SheetRow    { get; set; }   // 1-based Excel row (for error messages)
        public string StudentCode { get; set; } = "";
        public string StudentName { get; set; } = "";
        public string SubjectCode { get; set; } = "";   // copied from result-level subject

        // Monthly + academic subject: three sub-scores (max 2/3/5).
        // Monthly + CHA1/LAB1:        Score only (max 10).
        // Semester:                   FinalScore only (max 10) — Mid stays owned
        //                             by RecomputeMidFromMonthly and is not imported.
        public int?   DisciplineScore { get; set; }   // ກິດຈະກຳ (/2) — monthly academic
        public int?   ActivityScore   { get; set; }   // ຮ່ວມຮຽນ (/3) — monthly academic
        public int?   HomeworkScore   { get; set; }   // ກວດກາ (/5) — monthly academic
        public int?   Score           { get; set; }   // /10 — monthly CHA1/LAB1
        public int?   MidScore        { get; set; }   // /10 — semester (retained for legacy callers)
        public int?   FinalScore      { get; set; }   // /10 — semester

        public bool   IsValid { get; set; }
        public string Status  { get; set; } = "";

        // Resolved during validation, used during save.
        public int    StudentID { get; set; }
        public int    SubjectID { get; set; }
        public int    EnrollID  { get; set; }         // 0 for CHA1/LAB1 (Evaluation table is per-student)
        public bool   IsEval    { get; set; }         // CHA1 or LAB1

        // Sum of sub-scores (monthly academic) — exposed for preview display.
        public int    SubScoreTotal =>
            (DisciplineScore ?? 0) + (ActivityScore ?? 0) + (HomeworkScore ?? 0);
    }

    public class ImportResult
    {
        public ImportKind Kind;
        public string Year        = "";
        public string Grade       = "";
        public string Room        = "";
        public int    Month;
        public int    Semester;
        public string SubjectCode = "";
        public string SubjectName = "";
        public bool   IsEvalSubject;       // CHA1 or LAB1 — drives preview column layout
        public List<ImportRow> Rows = new();

        public string? FatalError;
        public int ValidCount   => Rows.Count(r => r.IsValid);
        public int InvalidCount => Rows.Count(r => !r.IsValid);
    }

    public static class ExcelImport
    {
        public const string MagicMonthly  = "SIS_LAO_MONTHLY_v3";
        public const string MagicSemester = "SIS_LAO_SEMESTER_v4";

        // Font names — Phetsarath OT is the standard Lao font installed at the
        // school. Excel falls back to a system default if it's not present, but
        // the file still renders correctly (just in the wrong font).
        private const string LaoFont = "Phetsarath OT";

        // Per-sub-score column maximums for monthly academic — match the
        // MonthlyAssessments column constraints used by MonthlyScoresPage.
        private const int MaxDiscipline = 2;
        private const int MaxActivity   = 3;
        private const int MaxHomework   = 5;

        // ─── Template generation ──────────────────────────────────

        public static void BuildMonthlyTemplate(string outPath, string year, string grade, string room, int month, string subjectCode)
        {
            var (subId, subName) = ResolveSubject(grade, subjectCode);
            bool isEval = subjectCode == "CHA1" || subjectCode == "LAB1";
            var roster = LoadRoster(year, grade, room);

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("ນຳເຂົ້າຄະແນນປະຈຳເດືອນ");
            SetDefaultFont(ws);

            string title = $"ນຳເຂົ້າຄະແນນປະຈຳເດືອນ — ປີ {year} · {grade}/{room} · ເດືອນ {month:D2} · ວິຊາ {subjectCode} — {subName}";
            ApplyTitle(ws, "A1", title);

            // Academic monthly = 3 sub-score cols (D/E/F). The reference file
            // uses a slightly darker fill (#BFBFBF) + default-aligned header
            // for E/F to visually group the sub-score columns separately from
            // the identifying A-D group.
            // CHA1/LAB1 monthly = single /10 col (D) with no second group.
            int lastCol;
            if (isEval)
            {
                WriteHeader(ws, "A3", "ລຳດັບ",       HeaderStyle.Primary);
                WriteHeader(ws, "B3", "ລະຫັດນັກຮຽນ", HeaderStyle.Primary);
                WriteHeader(ws, "C3", "ຊື່ນັກຮຽນ",  HeaderStyle.Primary);
                WriteHeader(ws, "D3", "ຄະແນນ (/10)", HeaderStyle.Primary);
                lastCol = 4;
                var existing = LoadExistingMonthlyEvalForSubject(year, grade, room, month, subjectCode);
                int row = 4, ord = 1;
                foreach (DataRow s in roster.Rows)
                {
                    int    sid  = Convert.ToInt32(s["StudentID"]);
                    string code = s["StudentCode"].ToString()!;
                    string name = s["FullName"].ToString()!;
                    WriteDataCell(ws, row, 1, ord++, center: true);
                    WriteDataCell(ws, row, 2, code,  center: true);
                    WriteDataCell(ws, row, 3, name,  center: false);
                    if (existing.TryGetValue(sid, out double v))
                        WriteDataCell(ws, row, 4, (int)Math.Round(v), center: true);
                    else
                        WriteEmptyCell(ws, row, 4, center: true);
                    row++;
                }
            }
            else
            {
                WriteHeader(ws, "A3", "ລຳດັບ",       HeaderStyle.Primary);
                WriteHeader(ws, "B3", "ລະຫັດນັກຮຽນ", HeaderStyle.Primary);
                WriteHeader(ws, "C3", "ຊື່ນັກຮຽນ",  HeaderStyle.Primary);
                WriteHeader(ws, "D3", "ກິດຈະກຳ2",   HeaderStyle.Primary);
                WriteHeader(ws, "E3", "ຮ່ວມຮຽນ3",   HeaderStyle.Secondary);
                WriteHeader(ws, "F3", "ກວດກາ5",     HeaderStyle.Secondary);
                lastCol = 6;
                var existing = LoadExistingMonthlyForSubject(year, grade, room, month, subjectCode);
                int row = 4, ord = 1;
                foreach (DataRow s in roster.Rows)
                {
                    int    sid  = Convert.ToInt32(s["StudentID"]);
                    string code = s["StudentCode"].ToString()!;
                    string name = s["FullName"].ToString()!;
                    WriteDataCell(ws, row, 1, ord++, center: true);
                    WriteDataCell(ws, row, 2, code,  center: true);
                    WriteDataCell(ws, row, 3, name,  center: false);
                    // D is centered (matches header group A-D); E/F use the
                    // default left alignment (matches the secondary group).
                    if (existing.TryGetValue(sid, out var sub))
                    {
                        WriteDataCell(ws, row, 4, sub.disc, center: true);
                        WriteDataCell(ws, row, 5, sub.act,  center: false);
                        WriteDataCell(ws, row, 6, sub.home, center: false);
                    }
                    else
                    {
                        WriteEmptyCell(ws, row, 4, center: true);
                        WriteEmptyCell(ws, row, 5, center: false);
                        WriteEmptyCell(ws, row, 6, center: false);
                    }
                    row++;
                }
            }

            ApplyMonthlyLayout(ws, lastCol, roster.Rows.Count, isEval);
            WriteMetadata(ws, MagicMonthly, year, grade, room, month, subjectCode);
            wb.SaveAs(outPath);
        }

        public static void BuildSemesterTemplate(string outPath, string year, string grade, string room, int sem, string subjectCode)
        {
            var (subId, subName) = ResolveSubject(grade, subjectCode);
            bool isEval = subjectCode == "CHA1" || subjectCode == "LAB1";
            var roster = LoadRoster(year, grade, room);
            var existing = LoadExistingSemesterForSubject(year, grade, room, sem, subjectCode);

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("ນຳເຂົ້າຄະແນນພາກຮຽນ");
            SetDefaultFont(ws);

            string title = $"ນຳເຂົ້າຄະແນນພາກຮຽນ — ປີ {year} · {grade}/{room} · ພາກ {sem} · ວິຊາ {subjectCode} — {subName}";
            ApplyTitle(ws, "A1", title);

            // 4 visible columns only — Midterm is NOT imported here. Mid stays
            // auto-derived from MonthlyAssessments via RecomputeMidFromMonthly;
            // teachers only fill the final exam score in this template. All
            // headers use the single Primary style (matches the reference file).
            WriteHeader(ws, "A3", "ລຳດັບ",                       HeaderStyle.Primary);
            WriteHeader(ws, "B3", "ລະຫັດນັກຮຽນ",                 HeaderStyle.Primary);
            WriteHeader(ws, "C3", "ຊື່ນັກຮຽນ",                  HeaderStyle.Primary);
            WriteHeader(ws, "D3", "ຄະແນນສອບເສງພາກຮຽນ (/10)",   HeaderStyle.Primary);

            int row = 4, ord = 1;
            foreach (DataRow s in roster.Rows)
            {
                int    sid  = Convert.ToInt32(s["StudentID"]);
                string code = s["StudentCode"].ToString()!;
                string name = s["FullName"].ToString()!;
                WriteDataCell(ws, row, 1, ord++, center: true);
                WriteDataCell(ws, row, 2, code,  center: true);
                WriteDataCell(ws, row, 3, name,  center: false);
                if (existing.TryGetValue(sid, out var pair) && pair.final.HasValue)
                    WriteDataCell(ws, row, 4, (int)Math.Round(pair.final.Value), center: true);
                else
                    WriteEmptyCell(ws, row, 4, center: true);
                row++;
            }

            ApplySemesterLayout(ws, roster.Rows.Count);
            WriteSemesterMetadata(ws, year, grade, room, sem, subjectCode);
            wb.SaveAs(outPath);
        }

        // ─── Styling helpers ──────────────────────────────────────

        // Header has two visual groups — matches the reference file's split
        // between identifying columns (A-D in monthly academic / A-C elsewhere)
        // and score columns (E-F / D-E). Primary uses a slightly lighter gray
        // with centered text; Secondary uses a slightly darker gray with
        // default left alignment (matches Excel's "general" alignment for text).
        private enum HeaderStyle { Primary, Secondary }

        // Worksheet-wide default font so any unstyled cell still renders in
        // Phetsarath OT. ClosedXML's default would be Calibri otherwise.
        private static void SetDefaultFont(IXLWorksheet ws)
        {
            ws.Style.Font.FontName = LaoFont;
            ws.Style.Font.FontSize = 11;
        }

        private static void ApplyTitle(IXLWorksheet ws, string addr, string title)
        {
            var c = ws.Cell(addr);
            c.Value = title;
            c.Style.Font.FontName = LaoFont;
            c.Style.Font.Bold     = true;
            c.Style.Font.FontSize = 14;
            c.WorksheetRow().Height = 23.4;
        }

        private static void WriteHeader(IXLWorksheet ws, string addr, string text, HeaderStyle style)
        {
            var c = ws.Cell(addr);
            c.Value = text;
            c.Style.Font.FontName = LaoFont;
            c.Style.Font.Bold     = true;
            c.Style.Font.FontSize = 11;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            if (style == HeaderStyle.Primary)
            {
                c.Style.Fill.BackgroundColor = XLColor.FromHtml("#D3D3D3");
                c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            else
            {
                // Theme white tinted -0.25 → ≈ #BFBFBF, matching the secondary
                // header fill in the reference file. No alignment override:
                // Excel's "general" alignment leaves text left, numbers right.
                c.Style.Fill.BackgroundColor = XLColor.FromHtml("#BFBFBF");
            }
        }

        private static void WriteDataCell(IXLWorksheet ws, int row, int col, object value, bool center)
        {
            var c = ws.Cell(row, col);
            switch (value)
            {
                case int    i: c.Value = i; break;
                case double d: c.Value = d; break;
                case string s: c.Value = s; break;
                default:       c.Value = value?.ToString() ?? ""; break;
            }
            StyleDataCell(c, center);
        }

        private static void WriteEmptyCell(IXLWorksheet ws, int row, int col, bool center)
            => StyleDataCell(ws.Cell(row, col), center);

        private static void StyleDataCell(IXLCell c, bool center)
        {
            c.Style.Font.FontName = LaoFont;
            c.Style.Font.FontSize = 11;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            if (center)
                c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        // Column widths copied verbatim from the reference file so the layout
        // matches pixel-for-pixel when opened. AdjustToContents is deliberately
        // NOT called — the title in A1 would otherwise force col A absurdly
        // wide, defeating the narrow ordinal-only design the reference uses.
        // Column C (ຊື່ນັກຮຽນ) stays at the reference's 8.78 width; Excel will
        // visually truncate long Lao names in display only — the data itself
        // remains intact and teachers can expand the column manually if needed.
        private static void ApplyMonthlyLayout(IXLWorksheet ws, int lastCol, int rosterCount, bool isEval)
        {
            ws.Column(1).Width = 5.77734375;   // ລຳດັບ
            ws.Column(2).Width = 12.6640625;   // ລະຫັດນັກຮຽນ
            ws.Column(3).Width = 8.77734375;   // ຊື່ນັກຮຽນ
            ws.Column(4).Width = 9.44140625;   // ກິດຈະກຳ2 / ຄະແນນ (eval)
            if (!isEval)
            {
                ws.Column(5).Width = 10.109375; // ຮ່ວມຮຽນ3
                ws.Column(6).Width = 9.109375;  // ກວດກາ5
            }
            ws.Column(26).Width = 23.77734375; // Z (hidden metadata)
            ApplyCommonLayout(ws, lastCol, rosterCount);
        }

        private static void ApplySemesterLayout(IXLWorksheet ws, int rosterCount)
        {
            // Widths copied verbatim from the semester reference file. Note C is
            // wider (16.44) than monthly's 8.78 — the reference accommodates real
            // Lao names. D is very wide (26.78) because the header text is long.
            ws.Column(1).Width  = 6.33203125;   // ລຳດັບ
            ws.Column(2).Width  = 12.88671875;  // ລະຫັດນັກຮຽນ
            ws.Column(3).Width  = 16.44140625;  // ຊື່ນັກຮຽນ
            ws.Column(4).Width  = 26.77734375;  // ຄະແນນສອບເສງພາກຮຽນ (/10)
            ws.Column(25).Width = 9.109375;     // Y (hidden metadata)
            ApplyCommonLayout(ws, 4, rosterCount);
        }

        private static void ApplyCommonLayout(IXLWorksheet ws, int lastCol, int rosterCount)
        {
            // Default row height to match the reference file.
            ws.RowHeight = 18;

            // Margins match the reference file exactly. No PageSetup orientation
            // override — defaults to LetterPaper Portrait, same as the reference.
            ws.PageSetup.Margins.Left   = 0.75;
            ws.PageSetup.Margins.Right  = 0.75;
            ws.PageSetup.Margins.Top    = 0.75;
            ws.PageSetup.Margins.Bottom = 0.5;
            ws.PageSetup.Margins.Header = 0.5;
            ws.PageSetup.Margins.Footer = 0.75;
        }

        // Metadata is the only thing that marks a workbook as a valid template;
        // missing or mismatched == reject. The metadata column is hidden so
        // casual edits don't touch it. Monthly templates use Z (6 visible cols
        // A-F); semester templates use Y (only 4 visible cols A-D) — matches
        // the attached reference files exactly.
        private static void WriteMetadata(IXLWorksheet ws, string magic, string year, string grade, string room, int unit, string subjectCode)
            => WriteMetadataAt(ws, "Z", magic, year, grade, room, unit, subjectCode);

        private static void WriteSemesterMetadata(IXLWorksheet ws, string year, string grade, string room, int sem, string subjectCode)
            => WriteMetadataAt(ws, "Y", MagicSemester, year, grade, room, sem, subjectCode);

        private static void WriteMetadataAt(IXLWorksheet ws, string col, string magic, string year, string grade, string room, int unit, string subjectCode)
        {
            ws.Cell(col + "1").Value = magic;
            ws.Cell(col + "2").Value = year;
            ws.Cell(col + "3").Value = grade;
            ws.Cell(col + "4").Value = room;
            ws.Cell(col + "5").Value = unit;
            ws.Cell(col + "6").Value = subjectCode;
            ws.Column(col).Hide();
        }

        // ─── Parse + validate ────────────────────────────────────

        public static ImportResult ParseMonthly(string filePath, string expectedYear, string expectedGrade, string expectedRoom, int expectedMonth, string expectedSubjectCode)
            => Parse(filePath, ImportKind.Monthly, expectedYear, expectedGrade, expectedRoom, expectedMonth, expectedSubjectCode);

        public static ImportResult ParseSemester(string filePath, string expectedYear, string expectedGrade, string expectedRoom, int expectedSem, string expectedSubjectCode)
            => Parse(filePath, ImportKind.Semester, expectedYear, expectedGrade, expectedRoom, expectedSem, expectedSubjectCode);

        private static ImportResult Parse(string filePath, ImportKind kind, string year, string grade, string room, int unit, string subjectCode)
        {
            var result = new ImportResult { Kind = kind, Year = year, Grade = grade, Room = room, SubjectCode = subjectCode };
            if (kind == ImportKind.Monthly)  result.Month    = unit;
            else                             result.Semester = unit;

            try
            {
                var (subId, subName) = ResolveSubject(grade, subjectCode);
                result.SubjectName = subName;
            }
            catch (Exception ex)
            {
                result.FatalError = ex.Message;
                return result;
            }

            bool isEvalSubject = subjectCode == "CHA1" || subjectCode == "LAB1";
            result.IsEvalSubject = isEvalSubject;

            XLWorkbook wb;
            try { wb = new XLWorkbook(filePath); }
            catch (Exception ex)
            {
                result.FatalError = $"ບໍ່ສາມາດເປີດໄຟລ໌ Excel ໄດ້: {ex.Message}";
                return result;
            }

            using (wb)
            {
                var ws = wb.Worksheet(1);
                // Monthly templates put metadata in Z (col 26 — 6 visible cols).
                // Semester templates put it in Y (col 25 — only 4 visible cols).
                string metaCol = kind == ImportKind.Monthly ? "Z" : "Y";
                string magic = (ws.Cell(metaCol + "1").GetString() ?? "").Trim();
                string want  = kind == ImportKind.Monthly ? MagicMonthly : MagicSemester;
                if (magic != want)
                {
                    result.FatalError = "ໄຟລ໌ນີ້ບໍ່ແມ່ນແບບຟອມຂອງລະບົບ — ກະລຸນາໂຫຼດແບບຟອມໃໝ່ດ້ວຍ ‘📥 ດາວໂຫຼດແບບຟອມ’ ກ່ອນ.";
                    return result;
                }

                string fileYear     = (ws.Cell(metaCol + "2").GetString() ?? "").Trim();
                string fileGrade    = (ws.Cell(metaCol + "3").GetString() ?? "").Trim();
                string fileRoom     = (ws.Cell(metaCol + "4").GetString() ?? "").Trim();
                int    fileUnit     = (int)ws.Cell(metaCol + "5").GetDouble();
                string fileSubject  = (ws.Cell(metaCol + "6").GetString() ?? "").Trim();

                if (fileYear != year || fileGrade != grade || fileRoom != room || fileUnit != unit)
                {
                    string label = kind == ImportKind.Monthly ? "ເດືອນ" : "ພາກ";
                    result.FatalError =
                        $"ໄຟລ໌ນີ້ສ້າງມາສຳລັບ ປີ {fileYear} · {fileGrade}/{fileRoom} · {label} {fileUnit} " +
                        $"ແຕ່ປະຈຸບັນເລືອກ ປີ {year} · {grade}/{room} · {label} {unit}. " +
                        "ກະລຸນາໂຫຼດແບບຟອມໃໝ່ ຫຼື ປ່ຽນຕົວກອງໃຫ້ກົງກັນ.";
                    return result;
                }
                if (fileSubject != subjectCode)
                {
                    result.FatalError =
                        $"ໄຟລ໌ນີ້ສ້າງມາສຳລັບວິຊາ ‘{fileSubject}’ ແຕ່ປະຈຸບັນເລືອກວິຊາ ‘{subjectCode}’. " +
                        "ກະລຸນາໂຫຼດແບບຟອມໃໝ່ ຫຼື ປ່ຽນວິຊາໃຫ້ກົງກັນ.";
                    return result;
                }

                var studentByCode = LoadRoster(year, grade, room).AsEnumerable()
                    .ToDictionary(r => r["StudentCode"].ToString()!, r => r);

                int sem = kind == ImportKind.Monthly ? DB.SemesterForMonth(unit) : unit;
                int subjectId = ResolveSubject(grade, subjectCode).id;

                var enrollBySid = isEvalSubject
                    ? new Dictionary<int, int>()
                    : LoadEnrollmentsForSubject(year, sem, subjectId,
                        studentByCode.Values.Select(r => Convert.ToInt32(r["StudentID"])).ToList());

                var seen = new HashSet<string>();

                int firstDataRow = 4;
                int lastRow      = ws.LastRowUsed()?.RowNumber() ?? firstDataRow - 1;
                int ord = 0;

                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    string sCode = (ws.Cell(r, 2).GetString() ?? "").Trim();
                    string sName = (ws.Cell(r, 3).GetString() ?? "").Trim();

                    // Read every score column for this kind/subject so the parser
                    // can decide "any score present?" correctly. Monthly academic
                    // has 3, monthly eval and semester have 2 each.
                    string rawA = "", rawB = "", rawC = "";
                    bool monthlyAcademic = kind == ImportKind.Monthly && !isEvalSubject;
                    if (monthlyAcademic)
                    {
                        rawA = (ws.Cell(r, 4).GetString() ?? "").Trim();
                        rawB = (ws.Cell(r, 5).GetString() ?? "").Trim();
                        rawC = (ws.Cell(r, 6).GetString() ?? "").Trim();
                    }
                    else
                    {
                        // Monthly + eval OR semester: single score in col D.
                        rawA = (ws.Cell(r, 4).GetString() ?? "").Trim();
                    }

                    // Skip totally blank rows.
                    if (string.IsNullOrEmpty(sCode)
                        && string.IsNullOrEmpty(rawA) && string.IsNullOrEmpty(rawB) && string.IsNullOrEmpty(rawC))
                        continue;
                    // Skip rows with no score — protects existing data when only
                    // some rows are filled.
                    if (string.IsNullOrEmpty(rawA) && string.IsNullOrEmpty(rawB) && string.IsNullOrEmpty(rawC))
                        continue;

                    ord++;
                    var row = new ImportRow
                    {
                        RowNo       = ord,
                        SheetRow    = r,
                        StudentCode = sCode,
                        StudentName = sName,
                        SubjectCode = subjectCode,
                        SubjectID   = subjectId,
                        IsEval      = isEvalSubject,
                    };

                    if (!studentByCode.TryGetValue(sCode, out var sRow))
                    {
                        row.Status = $"❌ ບໍ່ພົບນັກຮຽນ ‘{sCode}’ ໃນ {grade}/{room}";
                    }
                    else if ((sRow["Status"].ToString() ?? "") != "ກຳລັງຮຽນ")
                    {
                        row.StudentID = Convert.ToInt32(sRow["StudentID"]);
                        // Backfill the display name from the DB. The Excel sheet's
                        // column C may be blank if the teacher deleted it or built
                        // the file from scratch; showing the DB-resolved name in
                        // the preview is clearer than an empty cell.
                        if (string.IsNullOrEmpty(row.StudentName))
                            row.StudentName = sRow["FullName"].ToString() ?? "";
                        row.Status = $"❌ ນັກຮຽນ ‘{sCode}’ ມີສະຖານະ ‘{sRow["Status"]}’ (ບໍ່ສາມາດແກ້ໄຂໄດ້)";
                    }
                    else
                    {
                        row.StudentID = Convert.ToInt32(sRow["StudentID"]);
                        if (string.IsNullOrEmpty(row.StudentName))
                            row.StudentName = sRow["FullName"].ToString() ?? "";

                        if (!isEvalSubject && !enrollBySid.ContainsKey(row.StudentID))
                        {
                            row.Status = $"❌ ‘{sCode}’ ບໍ່ໄດ້ລົງທະບຽນວິຊາ ‘{subjectCode}’ ສຳລັບພາກ {sem}";
                        }
                        else
                        {
                            row.EnrollID = isEvalSubject ? 0 : enrollBySid[row.StudentID];

                            // Auto-implemented properties can't be passed as `out`
                            // parameters, so we parse into locals then copy across.
                            bool ok = true;
                            if (kind == ImportKind.Monthly && isEvalSubject)
                            {
                                if (!ParseScoreColumn(rawA, 10, out int? score))
                                { row.Status = $"❌ ຄະແນນຜິດ: ‘{rawA}’ ຕ້ອງເປັນຕົວເລກເຕັມ 0-10"; ok = false; }
                                else if (score == null)
                                { row.Status = "❌ ບໍ່ມີຄະແນນ"; ok = false; }
                                else row.Score = score;
                            }
                            else if (kind == ImportKind.Monthly)
                            {
                                // Academic monthly: three columns, per-column max.
                                if (!ParseScoreColumn(rawA, MaxDiscipline, out int? disc))
                                { row.Status = $"❌ ກິດຈະກຳຜິດ: ‘{rawA}’ ຕ້ອງເປັນຕົວເລກເຕັມ 0-{MaxDiscipline}"; ok = false; }
                                else if (!ParseScoreColumn(rawB, MaxActivity, out int? act))
                                { row.Status = $"❌ ຮ່ວມຮຽນຜິດ: ‘{rawB}’ ຕ້ອງເປັນຕົວເລກເຕັມ 0-{MaxActivity}"; ok = false; }
                                else if (!ParseScoreColumn(rawC, MaxHomework, out int? home))
                                { row.Status = $"❌ ກວດກາຜິດ: ‘{rawC}’ ຕ້ອງເປັນຕົວເລກເຕັມ 0-{MaxHomework}"; ok = false; }
                                else if (disc == null && act == null && home == null)
                                { row.Status = "❌ ບໍ່ມີຄະແນນ"; ok = false; }
                                else
                                {
                                    row.DisciplineScore = disc;
                                    row.ActivityScore   = act;
                                    row.HomeworkScore   = home;
                                }
                            }
                            else
                            {
                                // Semester: only the final-exam score is imported.
                                // Mid is auto-derived from MonthlyAssessments and stays untouched.
                                if (!ParseScoreColumn(rawA, 10, out int? finalS))
                                { row.Status = $"❌ ສອບເສງພາກຮຽນຜິດ: ‘{rawA}’ ຕ້ອງເປັນຕົວເລກເຕັມ 0-10"; ok = false; }
                                else if (finalS == null)
                                { row.Status = "❌ ບໍ່ມີຄະແນນ"; ok = false; }
                                else row.FinalScore = finalS;
                            }
                            if (ok)
                            {
                                if (!seen.Add(sCode))
                                    row.Status = $"❌ ຊ້ຳກັນ — ນັກຮຽນ ‘{sCode}’ ມີຫຼາຍກວ່າ 1 ແຖວໃນໄຟລ໌";
                                else
                                {
                                    row.IsValid = true;
                                    row.Status  = "✅ ຖືກຕ້ອງ";
                                }
                            }
                        }
                    }

                    result.Rows.Add(row);
                }
            }

            return result;
        }

        // Empty string → null (skip). Any non-empty input must satisfy
        // DB.TryParseIntScore(0, max) — same strict integer rule the manual
        // entry grids use, so the import path can't slip in fractional values.
        private static bool ParseScoreColumn(string raw, int max, out int? value)
        {
            value = null;
            if (string.IsNullOrWhiteSpace(raw)) return true;
            if (!DB.TryParseIntScore(raw, 0, max, out int v)) return false;
            value = v;
            return true;
        }

        // ─── Save ─────────────────────────────────────────────────

        public static int SaveImport(ImportResult result)
        {
            if (result.FatalError != null) return 0;

            int saved = 0;
            var touchedEnrollments = new HashSet<int>();

            using var conn = DB.Open();
            using var tx   = conn.BeginTransaction();
            try
            {
                foreach (var row in result.Rows)
                {
                    if (!row.IsValid) continue;

                    if (result.Kind == ImportKind.Monthly)
                    {
                        if (row.IsEval)
                        {
                            string ctx = DB.MonthContextName(result.Month) ?? "";
                            DB.SetEvaluationScore(row.StudentID, result.Year, ctx, row.SubjectCode,
                                row.Score, conn, tx);
                        }
                        else
                        {
                            // Direct write — the template now carries the three
                            // sub-scores explicitly, so no split is needed.
                            int d = row.DisciplineScore ?? 0;
                            int a = row.ActivityScore   ?? 0;
                            int h = row.HomeworkScore   ?? 0;
                            DB.ExecTx(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore,UpdatedAt)
                                        VALUES(@e,@m,@a,@d,@h,datetime('now','localtime'))
                                        ON CONFLICT(EnrollID,Month) DO UPDATE SET
                                          ActivityScore=excluded.ActivityScore,
                                          DisciplineScore=excluded.DisciplineScore,
                                          HomeworkScore=excluded.HomeworkScore,
                                          UpdatedAt=datetime('now','localtime')",
                                conn, tx,
                                ("@e", row.EnrollID), ("@m", result.Month),
                                ("@a", a), ("@d", d), ("@h", h));
                            touchedEnrollments.Add(row.EnrollID);
                        }
                    }
                    else // Semester
                    {
                        if (row.IsEval)
                        {
                            string ctx = $"SEM{result.Semester}";
                            DB.SetEvaluationScore(row.StudentID, result.Year, ctx, row.SubjectCode,
                                row.FinalScore, conn, tx);
                        }
                        else
                        {
                            DB.ExecTx(@"INSERT OR IGNORE INTO Scores(EnrollID,MidScore,FinalScore,TotalScore,Level)
                                        VALUES(@e,0,0,0,'')",
                                conn, tx, ("@e", row.EnrollID));
                            // Only FinalScore is imported. MidScore is preserved —
                            // it's owned by RecomputeMidFromMonthly (averaged from
                            // the monthly entries). Total + Level recomputed below.
                            DB.ExecTx("UPDATE Scores SET FinalScore=@f, UpdatedAt=datetime('now','localtime') WHERE EnrollID=@e",
                                conn, tx, ("@f", row.FinalScore!.Value), ("@e", row.EnrollID));

                            DB.ExecTx(@"UPDATE Scores
                                        SET TotalScore = ROUND(MidScore*(@mp/100.0) + FinalScore*(@fp/100.0), 2),
                                            Level = CASE
                                                WHEN MidScore*(@mp/100.0) + FinalScore*(@fp/100.0) >= 8 THEN 'ດີຫຼາຍ'
                                                WHEN MidScore*(@mp/100.0) + FinalScore*(@fp/100.0) >= 6 THEN 'ດີ'
                                                WHEN MidScore*(@mp/100.0) + FinalScore*(@fp/100.0) >= @ps THEN 'ຜ່ານ'
                                                ELSE 'ບໍ່ຜ່ານ' END,
                                            UpdatedAt = datetime('now','localtime')
                                        WHERE EnrollID=@e",
                                conn, tx,
                                ("@mp", DB.MidPct), ("@fp", DB.FinalPct),
                                ("@ps", DB.PassScore), ("@e", row.EnrollID));
                        }
                    }
                    saved++;
                }
                tx.Commit();
            }
            catch
            {
                tx.Rollback();
                throw;
            }

            if (touchedEnrollments.Count > 0)
            {
                using var rc = DB.Open();
                foreach (int eid in touchedEnrollments)
                    DB.RecomputeMidFromMonthly(eid, rc);
            }

            string detail = result.Kind == ImportKind.Monthly
                ? $"{saved} ແຖວ — ປີ {result.Year} {result.Grade}/{result.Room} ເດືອນ {result.Month:D2} ວິຊາ {result.SubjectCode}"
                : $"{saved} ແຖວ — ປີ {result.Year} {result.Grade}/{result.Room} ພາກ {result.Semester} ວິຊາ {result.SubjectCode}";
            DB.Log("ImportScores", detail);
            return saved;
        }

        // ─── Lookups ──────────────────────────────────────────────

        private static DataTable LoadRoster(string year, string grade, string room)
        {
            return DB.Query(@"
                SELECT StudentID, StudentCode, FirstName||' '||LastName AS FullName, Status
                FROM Students
                WHERE GradeLevel=@g AND ClassRoom=@r AND AcademicYear=@y
                ORDER BY StudentCode",
                null, ("@g", grade), ("@r", room), ("@y", year));
        }

        private static (int id, string name) ResolveSubject(string grade, string subjectCode)
        {
            var dt = DB.Query(@"
                SELECT SubjectID, SubjectName FROM Subjects
                WHERE SubjectCode=@c
                  AND (GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel='')",
                null, ("@c", subjectCode), ("@g", grade));
            if (dt.Rows.Count == 0)
                throw new InvalidOperationException(
                    $"ບໍ່ພົບວິຊາ ‘{subjectCode}’ ສຳລັບຊັ້ນ ‘{grade}’");
            return (Convert.ToInt32(dt.Rows[0]["SubjectID"]), dt.Rows[0]["SubjectName"].ToString()!);
        }

        private static Dictionary<int, int> LoadEnrollmentsForSubject(string year, int sem, int subjectId, List<int> studentIds)
        {
            var map = new Dictionary<int, int>();
            if (studentIds.Count == 0) return map;
            string inList = string.Join(",", studentIds);
            var dt = DB.Query($@"
                SELECT EnrollID, StudentID FROM Enrollments
                WHERE AcademicYear=@y AND Semester=@s AND SubjectID=@sub
                  AND StudentID IN ({inList})",
                null, ("@y", year), ("@s", sem), ("@sub", subjectId));
            foreach (DataRow r in dt.Rows)
                map[Convert.ToInt32(r["StudentID"])] = Convert.ToInt32(r["EnrollID"]);
            return map;
        }

        // Academic monthly prefill: read the three sub-scores stored in
        // MonthlyAssessments verbatim — the template column layout matches.
        private static Dictionary<int, (int disc, int act, int home)> LoadExistingMonthlyForSubject(string year, string grade, string room, int month, string subjectCode)
        {
            var map = new Dictionary<int, (int, int, int)>();
            int sem = DB.SemesterForMonth(month);
            var dt = DB.Query(@"
                SELECT s.StudentID,
                       IFNULL(ma.DisciplineScore,0) AS Disc,
                       IFNULL(ma.ActivityScore,0)   AS Act,
                       IFNULL(ma.HomeworkScore,0)   AS Home
                FROM Students s
                JOIN Enrollments e ON e.StudentID=s.StudentID AND e.AcademicYear=@y AND e.Semester=@sm
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID AND sub.SubjectCode=@sc
                JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@mo
                WHERE s.GradeLevel=@g AND s.ClassRoom=@r AND s.AcademicYear=@y",
                null, ("@y", year), ("@sm", sem), ("@mo", month), ("@sc", subjectCode),
                      ("@g", grade), ("@r", room));
            foreach (DataRow r in dt.Rows)
                map[Convert.ToInt32(r["StudentID"])] = (
                    Convert.ToInt32(r["Disc"]),
                    Convert.ToInt32(r["Act"]),
                    Convert.ToInt32(r["Home"]));
            return map;
        }

        // CHA1/LAB1 monthly prefill: single /10 from EvaluationScores(Month{N}).
        private static Dictionary<int, double> LoadExistingMonthlyEvalForSubject(string year, string grade, string room, int month, string subjectCode)
        {
            var map = new Dictionary<int, double>();
            string ctx = DB.MonthContextName(month) ?? "";
            var ev = DB.Query(@"
                SELECT ev.StudentID, ev.Score FROM EvaluationScores ev
                JOIN Students s ON s.StudentID=ev.StudentID
                WHERE ev.AcademicYear=@y AND ev.Context=@c AND ev.SubjectCode=@sc
                  AND s.GradeLevel=@g AND s.ClassRoom=@r AND s.AcademicYear=@y",
                null, ("@y", year), ("@c", ctx), ("@sc", subjectCode),
                      ("@g", grade), ("@r", room));
            foreach (DataRow r in ev.Rows)
                map[Convert.ToInt32(r["StudentID"])] = Convert.ToDouble(r["Score"]);
            return map;
        }

        // Semester prefill: (mid, final) for academic from Scores; (null, final)
        // for CHA1/LAB1 from EvaluationScores(SEM{N}).
        private static Dictionary<int, (double? mid, double? final)> LoadExistingSemesterForSubject(string year, string grade, string room, int sem, string subjectCode)
        {
            var map = new Dictionary<int, (double?, double?)>();
            bool isEval = subjectCode == "CHA1" || subjectCode == "LAB1";

            if (isEval)
            {
                string ctx = $"SEM{sem}";
                var ev = DB.Query(@"
                    SELECT ev.StudentID, ev.Score FROM EvaluationScores ev
                    JOIN Students s ON s.StudentID=ev.StudentID
                    WHERE ev.AcademicYear=@y AND ev.Context=@c AND ev.SubjectCode=@sc
                      AND s.GradeLevel=@g AND s.ClassRoom=@r AND s.AcademicYear=@y",
                    null, ("@y", year), ("@c", ctx), ("@sc", subjectCode),
                          ("@g", grade), ("@r", room));
                foreach (DataRow r in ev.Rows)
                    map[Convert.ToInt32(r["StudentID"])] = (null, Convert.ToDouble(r["Score"]));
            }
            else
            {
                var dt = DB.Query(@"
                    SELECT s.StudentID, sc.MidScore, sc.FinalScore
                    FROM Students s
                    JOIN Enrollments e ON e.StudentID=s.StudentID AND e.AcademicYear=@y AND e.Semester=@sm
                    JOIN Subjects sub ON sub.SubjectID=e.SubjectID AND sub.SubjectCode=@sc
                    JOIN Scores sc ON sc.EnrollID=e.EnrollID
                    WHERE s.GradeLevel=@g AND s.ClassRoom=@r AND s.AcademicYear=@y",
                    null, ("@y", year), ("@sm", sem), ("@sc", subjectCode),
                          ("@g", grade), ("@r", room));
                foreach (DataRow r in dt.Rows)
                    map[Convert.ToInt32(r["StudentID"])] = (
                        r["MidScore"]   == DBNull.Value ? null : Convert.ToDouble(r["MidScore"]),
                        r["FinalScore"] == DBNull.Value ? null : Convert.ToDouble(r["FinalScore"]));
            }
            return map;
        }
    }
}
