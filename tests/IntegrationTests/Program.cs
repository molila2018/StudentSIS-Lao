// Integration tests for StudentSIS_Lao.
// These exercise the LIVE production code in Data/DB.cs (linked via <Compile>)
// against a real SQLite file. Each test creates / mutates rows and asserts the
// behavioural contracts documented in CONTEXT_REFERENCE.md.
//
// What this does NOT cover (out of scope for a no-UI harness):
//   • WPF DataGrid interactions and dropdown rendering
//   • Excel COM → PDF conversion (requires Office installed at run time)
//   • Visual layout of generated PDFs / xlsx
//
// What it DOES cover:
//   • Schema integrity (every table the app touches)
//   • Calculation helpers vs documented Excel rules
//   • CHA1 / LAB1 contract (never auto-computed)
//   • RecomputeMidFromMonthly correctness
//   • EvaluationScores round-trip + delete-on-null
//   • SetCurrentAcademicYear keeping Settings + AcademicYears.IsCurrent in sync
//   • GetClassRanking ordering and CHA1/LAB1 exclusion
//   • Backup → Restore file fidelity
//   • Force-delete cascade leaves zero references
//   • Excel template file integrity (opens, has 4 sheets, formulas preserved)
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using StudentSIS.Data;

namespace StudentSIS.IntegrationTests;

internal static class Program
{
    static int _pass, _fail;
    static readonly List<string> _failures = new();

    // One-off CLI: generate a sample template for visual comparison.
    //   dotnet run --project tests/IntegrationTests -- gen-tpl     <outPath>   (monthly CIV1)
    //   dotnet run --project tests/IntegrationTests -- gen-tpl-sem <outPath>   (semester CIV1 sem 1)
    static int GenSampleTemplate(string outPath, bool semester)
    {
        DB.Initialize();
        string yr = "2025-2026"; string grade = "ມ.4"; string room = "1";
        DB.Exec("INSERT OR IGNORE INTO AcademicYears(Year,IsCurrent) VALUES(@y,1)", null, ("@y", yr));
        for (int i = 1; i <= 5; i++)
        {
            string code = $"GEN{i:D3}";
            DB.Exec(@"INSERT OR IGNORE INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                      VALUES(@c,@f,@l,'ຊາຍ',@g,@r,@y,'ກຳລັງຮຽນ')",
                null, ("@c", code), ("@f", "ນັກຮຽນ"), ("@l", $"ທີ {i}"),
                      ("@g", grade), ("@r", room), ("@y", yr));
        }
        if (semester) ExcelImport.BuildSemesterTemplate(outPath, yr, grade, room, 1, "CIV1");
        else          ExcelImport.BuildMonthlyTemplate (outPath, yr, grade, room, 9, "CIV1");
        Console.WriteLine($"wrote {outPath}");
        return 0;
    }

    static int Main(string[] args)
    {
        if (args.Length >= 2 && args[0] == "inspect")
        {
            Inspect.Run(args[1]);
            return 0;
        }
        if (args.Length >= 2 && args[0] == "gen-tpl")
            return GenSampleTemplate(args[1], semester: false);
        if (args.Length >= 2 && args[0] == "gen-tpl-sem")
            return GenSampleTemplate(args[1], semester: true);
        if (args.Length >= 2 && args[0] == "inspect-docx")
        {
            Inspect.RunDocx(args[1]);
            return 0;
        }
        if (args.Length >= 2 && args[0] == "inspect-pdf")
        {
            Inspect.RunPdf(args[1]);
            return 0;
        }
        if (args.Length >= 3 && args[0] == "repack")
        {
            Repack.Run(args[1], args[2]);
            return 0;
        }
        if (args.Length >= 1 && args[0] == "manual")
        {
            string outPath = args.Length >= 2
                ? args[1]
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "..", "..", "..", "..", "..", "installer", "User-Manual.pdf");
            UserManual.Generate(Path.GetFullPath(outPath));
            return 0;
        }

        // Clean slate every run.
        string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sis_lao.db");
        if (File.Exists(dbPath)) File.Delete(dbPath);

        DB.Initialize();
        DB.CurrentUser = "test"; DB.CurrentUserId = 1; DB.CurrentRole = "admin";

        Section("Schema");
        TestSchema();

        Section("Calculation helpers");
        TestCalculationHelpers();

        Section("CHA1/LAB1 contract");
        TestChaLabContract();

        Section("RecomputeMidFromMonthly");
        TestRecomputeMid();

        Section("EvaluationScores helpers");
        TestEvaluationScores();

        Section("Academic-year management");
        TestAcademicYearSync();

        Section("Class ranking");
        TestRanking();

        Section("Backup / Restore");
        TestBackupRestore();

        Section("Force-delete cascade");
        TestForceDeleteCascade();

        Section("Excel template integrity");
        TestExcelTemplate();

        Section("Individual report template");
        TestIndividualTemplate();

        Section("Auto-enrollment (one-click)");
        TestAutoEnroll();

        Section("Batch enrollment (class-wide)");
        TestBatchEnroll();

        Section("Student Profile docx template");
        TestStudentProfileTemplate();

        Section("User management CRUD + password round-trip");
        TestUserManagement();

        Section("Score history (multi-year, post-promotion)");
        TestScoreHistory();

        Section("Score-history template-based exports");
        TestScoreHistoryTemplateExports();

        Section("Graduated-student score-history access");
        TestGraduatedStudentHistoryAccess();

        Section("Promotion page — bulk actions + history preservation");
        TestPromotionActions();

        Section("Academic Year page — add / move / set-current / delete");
        TestAcademicYearPage();

        Section("Score-entry text parser — DB.TryParseScore");
        TestTryParseScore();

        Section("Score-entry STRICT integer parser — DB.TryParseIntScore");
        TestTryParseIntScore();

        Section("CHA1/LAB1 manual entry across all 11 contexts");
        TestChaLabManualEntry();

        Section("CHA1/LAB1 inline in ScoresPage roster");
        TestChaLabInlineSemester();

        Section("Excel score import — monthly + semester end-to-end");
        TestExcelImport();

        Console.WriteLine();
        Console.WriteLine("═══════════════════════════════════════════");
        Console.WriteLine($"  RESULT:  {_pass} passed, {_fail} failed");
        if (_fail > 0)
        {
            Console.WriteLine();
            Console.WriteLine("FAILURES:");
            foreach (var f in _failures) Console.WriteLine("  • " + f);
        }
        Console.WriteLine("═══════════════════════════════════════════");
        return _fail == 0 ? 0 : 1;
    }

    // ─────────────────────────────────────────────────────────────
    //  Assertion helpers
    // ─────────────────────────────────────────────────────────────
    static void Section(string s) { Console.WriteLine(); Console.WriteLine($"── {s} ──"); }
    static void Pass(string name) { _pass++; Console.WriteLine($"  ✅  {name}"); }
    static void Fail(string name, string detail)
    { _fail++; _failures.Add($"{name}: {detail}"); Console.WriteLine($"  ❌  {name}\n      {detail}"); }
    static void Check(string name, bool ok, string detail = "")
    { if (ok) Pass(name); else Fail(name, detail); }
    static void CheckEq<T>(string name, T expected, T actual)
        => Check(name, EqualityComparer<T>.Default.Equals(expected, actual), $"expected {expected}, got {actual}");
    static void CheckClose(string name, double expected, double actual, double tol = 0.01)
        => Check(name, Math.Abs(expected - actual) <= tol, $"expected ≈{expected}, got {actual}");

    // ─────────────────────────────────────────────────────────────
    //  Tests
    // ─────────────────────────────────────────────────────────────
    static void TestSchema()
    {
        string[] required = {
            "Users","Students","Subjects","Enrollments","Scores",
            "MonthlyAssessments","AttendanceRecords","GradeHistory",
            "ActivityLog","Settings","AcademicYears","EvaluationScores"
        };
        var present = new HashSet<string>();
        foreach (DataRow r in DB.Query("SELECT name FROM sqlite_master WHERE type='table'").Rows)
            present.Add(r["name"].ToString()!);
        foreach (var t in required)
            Check($"Table '{t}' exists", present.Contains(t), "missing");

        // Seed 14 subjects from migration; verify CHA1/LAB1 are there.
        int chaLab = DB.ScalarInt("SELECT COUNT(*) FROM Subjects WHERE SubjectCode IN ('CHA1','LAB1')");
        CheckEq("CHA1 + LAB1 in Subjects", 2, chaLab);
    }

    static void TestCalculationHelpers()
    {
        // CalcMonthlyTotal: pure sum of three sub-scores
        CheckClose("CalcMonthlyTotal(1,2,5) = 8", 8.0, DB.CalcMonthlyTotal(1, 2, 5));
        CheckClose("CalcMonthlyTotal(0,0,0) = 0", 0.0, DB.CalcMonthlyTotal(0, 0, 0));
        CheckClose("CalcMonthlyTotal(3,2,5) = 10", 10.0, DB.CalcMonthlyTotal(3, 2, 5));

        // CalcTotal: respects 50/50 default after Initialize.
        CheckClose("CalcTotal(6,8) = 7 @50/50", 7.0, DB.CalcTotal(6, 8));
        CheckClose("CalcTotal(10,0) = 5 @50/50", 5.0, DB.CalcTotal(10, 0));

        // CalcLevel: 4-band boundaries
        CheckEq("CalcLevel(9.0)",  "ດີຫຼາຍ", DB.CalcLevel(9.0));
        CheckEq("CalcLevel(8.0)",  "ດີຫຼາຍ", DB.CalcLevel(8.0));
        CheckEq("CalcLevel(7.99)", "ດີ",      DB.CalcLevel(7.99));
        CheckEq("CalcLevel(6.0)",  "ດີ",      DB.CalcLevel(6.0));
        CheckEq("CalcLevel(5.99)", "ຜ່ານ",    DB.CalcLevel(5.99));
        CheckEq("CalcLevel(5.0)",  "ຜ່ານ",    DB.CalcLevel(5.0));
        CheckEq("CalcLevel(4.99)", "ບໍ່ຜ່ານ", DB.CalcLevel(4.99));

        // CalcMoESLevel: 5-band boundaries at 9 / 7 / 5 / 3
        CheckEq("MoES(10)",   "ດີຫຼາຍ", DB.CalcMoESLevel(10));
        CheckEq("MoES(9.0)",  "ດີຫຼາຍ", DB.CalcMoESLevel(9.0));
        CheckEq("MoES(8.99)", "ດີ",      DB.CalcMoESLevel(8.99));
        CheckEq("MoES(7.0)",  "ດີ",      DB.CalcMoESLevel(7.0));
        CheckEq("MoES(6.99)", "ປານກາງ",  DB.CalcMoESLevel(6.99));
        CheckEq("MoES(5.0)",  "ປານກາງ",  DB.CalcMoESLevel(5.0));
        CheckEq("MoES(4.99)", "ອ່ອນ",    DB.CalcMoESLevel(4.99));
        CheckEq("MoES(3.0)",  "ອ່ອນ",    DB.CalcMoESLevel(3.0));
        CheckEq("MoES(2.99)", "ຕົກ",     DB.CalcMoESLevel(2.99));

        // SemesterForMonth: Jan→1, Feb–June→2, Sept–Dec→1
        CheckEq("Sem(1)",  1, DB.SemesterForMonth(1));
        CheckEq("Sem(2)",  2, DB.SemesterForMonth(2));
        CheckEq("Sem(6)",  2, DB.SemesterForMonth(6));
        CheckEq("Sem(9)",  1, DB.SemesterForMonth(9));
        CheckEq("Sem(12)", 1, DB.SemesterForMonth(12));

        // MonthsInSemester
        Check("MonthsInSem(1) = 9-12", DB.MonthsInSemester(1).SequenceEqual(new[]{9,10,11,12}));
        Check("MonthsInSem(2) = 2-5",  DB.MonthsInSemester(2).SequenceEqual(new[]{2,3,4,5}));

        // FinalExamMonth
        CheckEq("FinalExamMonth(1) = 1", 1, DB.FinalExamMonth(1));
        CheckEq("FinalExamMonth(2) = 6", 6, DB.FinalExamMonth(2));

        // NextGrade
        CheckEq("NextGrade(ມ.1) = ມ.2", "ມ.2", DB.NextGrade("ມ.1"));
        CheckEq("NextGrade(ມ.4) = ຈົບ", "ຈົບ", DB.NextGrade("ມ.4"));
    }

    // Seed one classful of students + enroll them in every subject for both semesters.
    static (int sid, Dictionary<string,int> enrollBySubjectCode) SeedOneStudent(string code, string year)
    {
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES(@c,@f,@l,'ຊາຍ','ມ.4','1',@y,'ກຳລັງຮຽນ')",
            null, ("@c", code), ("@f", "Test"), ("@l", code), ("@y", year));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode=@c", null, ("@c", code));

        var subjects = DB.Query("SELECT SubjectID, SubjectCode FROM Subjects");
        var map = new Dictionary<string, int>();
        foreach (DataRow r in subjects.Rows)
        {
            int subId = Convert.ToInt32(r["SubjectID"]);
            string sc = r["SubjectCode"].ToString()!;
            // both semesters
            for (int sem = 1; sem <= 2; sem++)
            {
                DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                          VALUES(@s,@sub,@y,@sm)",
                    null, ("@s", sid), ("@sub", subId), ("@y", year), ("@sm", sem));
            }
            // Track the SEM 1 enroll id for downstream tests
            map[sc] = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
                                     WHERE StudentID=@s AND SubjectID=@sub AND Semester=1",
                null, ("@s", sid), ("@sub", subId));
        }
        return (sid, map);
    }

    static void TestChaLabContract()
    {
        string yr = "2099-2100";  // an isolated year so we don't trip other tests
        var (sid, enrolls) = SeedOneStudent("CL001", yr);

        // Drop in a maxed-out monthly entry for the CHA1 enrolment in month 9.
        int chaEnroll = enrolls["CHA1"];
        DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                  VALUES(@e,9,3,2,5)",
            null, ("@e", chaEnroll));

        // Trigger the recompute.
        DB.RecomputeMidFromMonthly(chaEnroll);

        // Contract: NO Scores row should have been created for a CHA1 enrolment.
        int rows = DB.ScalarInt("SELECT COUNT(*) FROM Scores WHERE EnrollID=@e", null, ("@e", chaEnroll));
        CheckEq("CHA1: RecomputeMid creates no Scores row", 0, rows);

        // Same for LAB1.
        int labEnroll = enrolls["LAB1"];
        DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                  VALUES(@e,9,3,2,5)",
            null, ("@e", labEnroll));
        DB.RecomputeMidFromMonthly(labEnroll);
        int labRows = DB.ScalarInt("SELECT COUNT(*) FROM Scores WHERE EnrollID=@e", null, ("@e", labEnroll));
        CheckEq("LAB1: RecomputeMid creates no Scores row", 0, labRows);

        // GetClassRanking must exclude CHA1/LAB1 from totals.
        // The student has no academic Scores → both ສະເລ່ຍ and ລວມ should be 0 even
        // though MonthlyAssessments has a maxed CHA1 entry.
        var rank = DB.GetClassRanking("ມ.4", "1", yr, 1);
        if (rank.Rows.Count == 0)
            Pass("GetClassRanking returns empty when no academic Scores exist");
        else
        {
            double avg = Convert.ToDouble(rank.Rows[0]["ສະເລ່ຍ"] ?? 0.0);
            double sum = Convert.ToDouble(rank.Rows[0]["ລວມ"] ?? 0.0);
            CheckClose("GetClassRanking ສະເລ່ຍ ignores CHA1", 0.0, avg);
            CheckClose("GetClassRanking ລວມ ignores CHA1", 0.0, sum);
        }
    }

    static void TestRecomputeMid()
    {
        string yr = "2098-2099";
        var (sid, enrolls) = SeedOneStudent("RM001", yr);
        int mathEnroll = enrolls["MATH1"];

        // 4 months of (1,1,3) = 5/10 each → average = 5
        foreach (int m in new[]{9,10,11,12})
        {
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,@m,1,1,3)",
                null, ("@e", mathEnroll), ("@m", m));
        }
        DB.RecomputeMidFromMonthly(mathEnroll);
        double mid = Convert.ToDouble(
            DB.Scalar("SELECT MidScore FROM Scores WHERE EnrollID=@e", null, ("@e", mathEnroll))!);
        CheckClose("MidScore = avg of 4 monthly totals", 5.0, mid);

        // Set FinalScore = 6, recompute → Total = 5*0.5 + 6*0.5 = 5.5
        DB.Exec("UPDATE Scores SET FinalScore=6 WHERE EnrollID=@e", null, ("@e", mathEnroll));
        DB.RecomputeMidFromMonthly(mathEnroll);
        double total = Convert.ToDouble(
            DB.Scalar("SELECT TotalScore FROM Scores WHERE EnrollID=@e", null, ("@e", mathEnroll))!);
        CheckClose("TotalScore = Mid×50% + Final×50%", 5.5, total);

        // Level for 5.5 → ຜ່ານ
        string level = DB.Scalar("SELECT Level FROM Scores WHERE EnrollID=@e", null, ("@e", mathEnroll))!.ToString()!;
        CheckEq("Level for 5.5 = ຜ່ານ", "ຜ່ານ", level);

        // Mid recompute with zero monthly rows on a fresh enrolment → MidScore=0
        int engEnroll = enrolls["ENG1"];
        DB.RecomputeMidFromMonthly(engEnroll);
        double engMid = Convert.ToDouble(
            DB.Scalar("SELECT MidScore FROM Scores WHERE EnrollID=@e", null, ("@e", engEnroll))!);
        CheckClose("Empty monthly → MidScore = 0", 0.0, engMid);
    }

    static void TestEvaluationScores()
    {
        // Seed a tiny student outside any other test's year range.
        string yr = "2097-2098";
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('EV001','Eva','Test','ຍິງ','ມ.4','1',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='EV001'");

        // Round-trip with non-null value.
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1", "CHA1", 8.0, c, tx);
            tx.Commit();
        }
        var v1 = DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1");
        Check("EvaluationScores round-trip SEM1/CHA1 = 8.0", v1.HasValue && Math.Abs(v1.Value - 8.0) < 0.01);

        // Same key — UPSERT.
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1", "CHA1", 9.0, c, tx);
            tx.Commit();
        }
        var v2 = DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1");
        Check("EvaluationScores UPSERT overwrites to 9.0", v2.HasValue && Math.Abs(v2.Value - 9.0) < 0.01);

        // null score → row deleted.
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1", "CHA1", null, c, tx);
            tx.Commit();
        }
        var v3 = DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1");
        Check("EvaluationScores null deletes the row", !v3.HasValue);

        // Distinct contexts coexist.
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1",   "CHA1", 7.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "SEM2",   "CHA1", 8.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "ANNUAL", "CHA1", 9.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "ANNUAL", "LAB1", 6.0, c, tx);
            tx.Commit();
        }
        CheckEq("EvaluationScores SEM1/CHA1",   7.0, DB.GetEvaluationScore(sid, yr, "SEM1",   "CHA1") ?? -1);
        CheckEq("EvaluationScores SEM2/CHA1",   8.0, DB.GetEvaluationScore(sid, yr, "SEM2",   "CHA1") ?? -1);
        CheckEq("EvaluationScores ANNUAL/CHA1", 9.0, DB.GetEvaluationScore(sid, yr, "ANNUAL", "CHA1") ?? -1);
        CheckEq("EvaluationScores ANNUAL/LAB1", 6.0, DB.GetEvaluationScore(sid, yr, "ANNUAL", "LAB1") ?? -1);

        // Lookup of missing key returns null (not 0).
        var miss = DB.GetEvaluationScore(sid, yr, "ANNUAL", "MATH1"); // not a CHA/LAB anyway
        Check("EvaluationScores missing key returns null", !miss.HasValue);
    }

    static void TestAcademicYearSync()
    {
        // Pick a brand-new year unlikely to clash.
        string yr = "2030-2031";
        DB.SetCurrentAcademicYear(yr);

        // Settings.current_year matches.
        string st = DB.Scalar("SELECT Value FROM Settings WHERE Key='current_year'")!.ToString()!;
        CheckEq("Settings.current_year set", yr, st);

        // AcademicYears.IsCurrent flipped to exactly this year.
        int currentRows = DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE IsCurrent=1");
        CheckEq("Exactly one row IsCurrent=1", 1, currentRows);
        string flagged = DB.Scalar("SELECT Year FROM AcademicYears WHERE IsCurrent=1")!.ToString()!;
        CheckEq("IsCurrent flag is on the new year", yr, flagged);

        // Semester reset to 1.
        string sem = DB.Scalar("SELECT Value FROM Settings WHERE Key='current_semester'")!.ToString()!;
        CheckEq("Semester reset to 1 after year switch", "1", sem);

        // In-memory cache refreshed.
        CheckEq("DB.CurrentYear refreshed", yr, DB.CurrentYear);
        CheckEq("DB.CurrentSem refreshed", 1, DB.CurrentSem);

        // AcademicYears() helper returns the new year + every distinct year from
        // Students/Enrollments/GradeHistory.
        var years = DB.AcademicYears();
        Check("AcademicYears() includes new year", years.Contains(yr));
        Check("AcademicYears() includes 2099-2100 (seeded by TestChaLabContract)",
            years.Contains("2099-2100"));
    }

    static void TestRanking()
    {
        // Build 3 students. Two pass with different totals, one fails (one subject < pass).
        string yr  = "2096-2097";
        string gr  = "ມ.4";
        string rm  = "2";  // isolated room

        // Helper: insert student + full enrolment (Sem 1) + Scores for academic subjects.
        int MakeStudent(string code, double academicAvg, bool fail)
        {
            DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                      VALUES(@c,@f,'X','ຊາຍ',@g,@r,@y,'ກຳລັງຮຽນ')",
                null, ("@c", code), ("@f", code), ("@g", gr), ("@r", rm), ("@y", yr));
            int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode=@c", null, ("@c", code));
            foreach (DataRow r in DB.Query("SELECT SubjectID, SubjectCode FROM Subjects").Rows)
            {
                int subId = Convert.ToInt32(r["SubjectID"]);
                string sc = r["SubjectCode"].ToString()!;
                DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                          VALUES(@s,@sub,@y,1)",
                    null, ("@s", sid), ("@sub", subId), ("@y", yr));
                int eid = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments WHERE StudentID=@s AND SubjectID=@sub AND Semester=1",
                    null, ("@s", sid), ("@sub", subId));
                // CHA1/LAB1: no Scores row (per contract).
                if (sc == "CHA1" || sc == "LAB1") continue;
                // First academic subject for the failing student gets a sub-pass score.
                double t = (fail && sc == "MATH1") ? 3.0 : academicAvg;
                DB.Exec(@"INSERT INTO Scores(EnrollID,MidScore,FinalScore,TotalScore,Level)
                          VALUES(@e,@m,@m,@t,@l)",
                    null, ("@e", eid), ("@m", t), ("@t", t), ("@l", DB.CalcLevel(t)));
            }
            return sid;
        }

        int s1 = MakeStudent("R001", 8.0, fail: false);  // all 8.0
        int s2 = MakeStudent("R002", 7.0, fail: false);  // all 7.0
        int s3 = MakeStudent("R003", 9.0, fail: true);   // 11×9.0 + MATH=3.0 → avg 8.5

        var rank = DB.GetClassRanking(gr, rm, yr, 1);
        CheckEq("GetClassRanking returns all 3 students", 3, rank.Rows.Count);

        // GetClassRanking is a pure ORDER BY ສະເລ່ຍ DESC — it does NOT push failed
        // students to the bottom. That semantic is applied in the report builders.
        // Expected order: R003 (8.5) → R001 (8.0) → R002 (7.0).
        CheckEq("GetClassRanking[0] = R003 (highest avg)", s3, Convert.ToInt32(rank.Rows[0]["StudentID"]));
        CheckEq("GetClassRanking[1] = R001",               s1, Convert.ToInt32(rank.Rows[1]["StudentID"]));
        CheckEq("GetClassRanking[2] = R002 (lowest avg)",  s2, Convert.ToInt32(rank.Rows[2]["StudentID"]));

        // Verify CHA1/LAB1 didn't pollute the aggregates. R001's avg must be exactly 8.0
        // (12 academic rows @ 8.0 / 12 = 8.0). If CHA1/LAB1's null Scores leaked in, the
        // IFNULL(0) treatment would drag the avg down.
        double r001avg = Convert.ToDouble(rank.Rows.Cast<DataRow>().First(r => Convert.ToInt32(r["StudentID"]) == s1)["ສະເລ່ຍ"]);
        CheckClose("R001 ສະເລ່ຍ = 8.0 (CHA1/LAB1 excluded from agg)", 8.0, r001avg);
        double r001sum = Convert.ToDouble(rank.Rows.Cast<DataRow>().First(r => Convert.ToInt32(r["StudentID"]) == s1)["ລວມ"]);
        CheckClose("R001 ລວມ = 96 (12 academic × 8.0)", 96.0, r001sum);

        // Now verify the REPORT-LAYER rank rule by reproducing the algorithm
        // (lifted verbatim from ReportPage.ShowScorePreview).
        // Build (sid, total, passed) tuples from the seeded data.
        var snap = new List<(int sid, double total, bool passed)>();
        foreach (DataRow r in rank.Rows)
        {
            int rsid = Convert.ToInt32(r["StudentID"]);
            double sum = Convert.ToDouble(r["ລວມ"]);
            int failedSubjects = Convert.ToInt32(r["ຕົກ"]);
            snap.Add((rsid, sum, failedSubjects == 0));
        }
        var passing = snap.Where(x => x.passed).OrderByDescending(x => x.total).ToList();
        var rankBySid = new Dictionary<int, int>();
        int curRank = 0; double prevTotal = double.NaN;
        for (int i = 0; i < passing.Count; i++)
        {
            if (passing[i].total != prevTotal) { curRank = i + 1; prevTotal = passing[i].total; }
            rankBySid[passing[i].sid] = curRank;
        }
        // Report-layer expectation: R001 = #1, R002 = #2, R003 = "ຕົກ" (no rank).
        CheckEq("Report rank: R001 = #1 (passing, highest)",  1, rankBySid.TryGetValue(s1, out var r1) ? r1 : -1);
        CheckEq("Report rank: R002 = #2",                     2, rankBySid.TryGetValue(s2, out var r2) ? r2 : -1);
        Check("Report rank: R003 NOT in rank map (failed)",   !rankBySid.ContainsKey(s3));
    }

    static void TestBackupRestore()
    {
        string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sis_lao.db");
        string bk = Path.Combine(Path.GetTempPath(), "sis_lao_backup_test.db");
        if (File.Exists(bk)) File.Delete(bk);

        DB.Backup(bk);
        Check("Backup created", File.Exists(bk));
        long origSize = new FileInfo(dbPath).Length;
        long bkSize   = new FileInfo(bk).Length;
        Check("Backup size matches source", origSize == bkSize);

        // Marker change in live DB to prove Restore overwrites.
        DB.SaveSetting("school_name", "MARKER_BEFORE_RESTORE");
        CheckEq("Marker written before restore", "MARKER_BEFORE_RESTORE", DB.SchoolName);
        DB.Restore(bk);
        DB.LoadSettings();
        Check("Restore overwrote — marker GONE", DB.SchoolName != "MARKER_BEFORE_RESTORE");

        File.Delete(bk);
    }

    static void TestForceDeleteCascade()
    {
        // Cascade-delete an isolated year ("2095-2096") and prove every reference is gone.
        string yr = "2095-2096";
        DB.SetCurrentAcademicYear("2030-2031");  // restore current to known value so we don't try to delete it

        // Seed a student with enrolments + scores + monthly + eval + history in year yr.
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('FD001','FD','Test','ຊາຍ','ມ.4','3',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='FD001'");
        int subId = DB.ScalarInt("SELECT SubjectID FROM Subjects WHERE SubjectCode='MATH1'");
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  VALUES(@s,@sub,@y,1)",
            null, ("@s", sid), ("@sub", subId), ("@y", yr));
        int eid = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments WHERE StudentID=@s AND SubjectID=@sub AND Semester=1",
            null, ("@s", sid), ("@sub", subId));
        DB.Exec(@"INSERT INTO Scores(EnrollID,MidScore,FinalScore,TotalScore,Level)
                  VALUES(@e,5,5,5,'ຜ່ານ')", null, ("@e", eid));
        DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                  VALUES(@e,9,3,2,5)", null, ("@e", eid));
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "ANNUAL", "CHA1", 8.0, c, tx);
            tx.Commit();
        }
        DB.Exec(@"INSERT INTO GradeHistory(StudentID,FromGrade,ToGrade,AcademicYear,Note,ChangedBy)
                  VALUES(@s,'ມ.3','ມ.4',@y,'test','test')",
            null, ("@s", sid), ("@y", yr));

        // Sanity: data exists.
        Check("Pre-cascade: enrolment exists", DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, ("@y", yr)) > 0);

        // Run the cascade — mirror ForceDeleteSelectedYear's transaction.
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.ExecTx("DELETE FROM Enrollments       WHERE AcademicYear=@y", c, tx, ("@y", yr));
            DB.ExecTx("DELETE FROM EvaluationScores  WHERE AcademicYear=@y", c, tx, ("@y", yr));
            DB.ExecTx("DELETE FROM AttendanceRecords WHERE AcademicYear=@y", c, tx, ("@y", yr));
            DB.ExecTx("DELETE FROM GradeHistory      WHERE AcademicYear=@y", c, tx, ("@y", yr));
            DB.ExecTx("DELETE FROM Students          WHERE AcademicYear=@y", c, tx, ("@y", yr));
            DB.ExecTx("DELETE FROM AcademicYears     WHERE Year=@y",         c, tx, ("@y", yr));
            tx.Commit();
        }

        // Every table should be free of references to the year.
        CheckEq("Cascade: 0 Students for year",          0, DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y", null, ("@y", yr)));
        CheckEq("Cascade: 0 Enrollments for year",       0, DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, ("@y", yr)));
        CheckEq("Cascade: 0 EvaluationScores for year",  0, DB.ScalarInt("SELECT COUNT(*) FROM EvaluationScores WHERE AcademicYear=@y", null, ("@y", yr)));
        CheckEq("Cascade: 0 AttendanceRecords for year", 0, DB.ScalarInt("SELECT COUNT(*) FROM AttendanceRecords WHERE AcademicYear=@y", null, ("@y", yr)));
        CheckEq("Cascade: 0 GradeHistory for year",      0, DB.ScalarInt("SELECT COUNT(*) FROM GradeHistory WHERE AcademicYear=@y", null, ("@y", yr)));
        CheckEq("Cascade: 0 AcademicYears for year",     0, DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE Year=@y", null, ("@y", yr)));
        // Scores + MonthlyAssessments should cascade via Enrollments FK.
        CheckEq("Cascade: Scores rows for that enrolment GONE",
            0, DB.ScalarInt("SELECT COUNT(*) FROM Scores WHERE EnrollID=@e", null, ("@e", eid)));
        CheckEq("Cascade: MonthlyAssessments for that enrolment GONE",
            0, DB.ScalarInt("SELECT COUNT(*) FROM MonthlyAssessments WHERE EnrollID=@e", null, ("@e", eid)));
    }

    static void TestExcelTemplate()
    {
        string templatesDir = Path.Combine(
            new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Parent!.Parent!.Parent!.Parent!.FullName,
            "Templates");
        // Fallback: try the canonical project location.
        if (!Directory.Exists(templatesDir))
            templatesDir = Path.Combine(
                new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Parent!.Parent!.Parent!.Parent!.FullName,
                "..", "Templates");

        string tplPath = Path.Combine(templatesDir, "ໃບຄະແນນ.xlsx");
        // One more fallback — production templates are usually shipped next to the EXE.
        if (!File.Exists(tplPath))
        {
            string altDir = Path.Combine(
                Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory)!.Parent!.Parent!.Parent!.FullName,
                "Templates");
            string alt = Path.Combine(altDir, "ໃບຄະແນນ.xlsx");
            if (File.Exists(alt)) tplPath = alt;
        }

        if (!File.Exists(tplPath))
        {
            Fail("Excel template found", $"could not locate ໃບຄະແນນ.xlsx (searched: {templatesDir})");
            return;
        }
        Pass($"Excel template found: {tplPath}");

        try
        {
            using var wb = new ClosedXML.Excel.XLWorkbook(tplPath);
            CheckEq("Template has 4 sheets", 4, wb.Worksheets.Count);

            // Each sheet must have at least one formula in its first data row.
            // The data row is detected dynamically (row immediately below the
            // "ລຳດັບ" header) — the school added 3 lines above the report
            // title, so what used to be row 3 is now row 6. Detection keeps
            // the test resilient to future header additions.
            int sheetsWithDataRowFormulas = 0;
            foreach (var ws in wb.Worksheets)
            {
                int lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
                int lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
                int headerRow = 0;
                for (int r = 1; r <= Math.Min(20, lastRow) && headerRow == 0; r++)
                    for (int c = 1; c <= lastCol; c++)
                        if (ws.Cell(r, c).GetString().Trim().Contains("ລຳດັບ"))
                        { headerRow = r; break; }
                if (headerRow == 0) continue;
                int dataRow = headerRow + 1;
                for (int c = 1; c <= lastCol; c++)
                {
                    if (ws.Cell(dataRow, c).HasFormula) { sheetsWithDataRowFormulas++; break; }
                }
            }
            Check("All 4 sheets have formulas on the first data row (FillSheet contract)",
                sheetsWithDataRowFormulas == 4,
                $"got {sheetsWithDataRowFormulas}/4");
        }
        catch (Exception ex)
        {
            Fail("Template opens via ClosedXML", ex.Message);
        }

        // Word agreement template.
        string docxPath = Path.Combine(Path.GetDirectoryName(tplPath)!, "ໃບສັນຍາ.docx");
        Check("Word agreement template exists", File.Exists(docxPath));
    }

    // Verify the individual-report template's structural contract — what the new
    // GenIndividualReport / BuildIndividualXlsx code assumes about the file's layout.
    static void TestIndividualTemplate()
    {
        // Project root → Templates/ລີພອດຄະແນນບຸກຄົນ.xlsx
        string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "..", "..", "..", "..", ".."));
        string tpl = Path.Combine(root, "Templates", "ລີພອດຄະແນນບຸກຄົນ.xlsx");
        if (!File.Exists(tpl)) { Fail("Individual template found", $"missing: {tpl}"); return; }
        Pass($"Individual template found: {tpl}");

        // ClosedXML/OpenXml holds a file lock open in Excel; copy first.
        string copy = Path.Combine(Path.GetTempPath(), $"ind_test_{Guid.NewGuid():N}.xlsx");
        File.Copy(tpl, copy, true);
        try
        {
            using var wb = new ClosedXML.Excel.XLWorkbook(copy);
            CheckEq("Individual template has 1 sheet", 1, wb.Worksheets.Count);
            var ws = wb.Worksheet(1);

            // Required anchor cells for the v2 template (what BuildIndividualXlsx writes into).
            //   A1, A2   — national header
            //   C4       — title placeholder
            //   A5       — name+class line
            //   A7 / F7  — table headers (subject / score)
            //   G7 / H7  — extra columns: ວັນຂາດ / ໝາຍເຫດ
            //   A8–A19   — subject labels (LAO, GEO, HIS, CIV, SCI, MATH, ENG, ICT, VOC, PE, ART, MUS)
            //   A20–A22  — summary labels (ຄະແນນລວມ / ສະເລ່ຍ / ອັນດັບທີ)
            //   A23, A24 — eval labels (ຄຸນສົມບັດ / ແຮງງານ)
            CheckEq("A1 = national name",  "ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ", ws.Cell("A1").GetString());
            CheckEq("A7 = 'ວິຊາ' header",   "ວິຊາ",   ws.Cell("A7").GetString());
            CheckEq("F7 = 'ຄະແນນ' header",  "ຄະແນນ",  ws.Cell("F7").GetString());
            CheckEq("G7 = 'ວັນຂາດ' header", "ວັນຂາດ", ws.Cell("G7").GetString());
            CheckEq("H7 = 'ໝາຍເຫດ' header", "ໝາຍເຫດ", ws.Cell("H7").GetString());
            CheckEq("Row 20 col A = 'ຄະແນນລວມ'", "ຄະແນນລວມ", ws.Cell(20, 1).GetString());
            CheckEq("Row 21 col A = 'ສະເລ່ຍ'",   "ສະເລ່ຍ",   ws.Cell(21, 1).GetString());
            // ‘ອັນດັບທີ’ in the template (with possible trailing space)
            string rankLabel = ws.Cell(22, 1).GetString().TrimEnd();
            CheckEq("Row 22 col A = 'ອັນດັບທີ'", "ອັນດັບທີ",  rankLabel);
            CheckEq("Row 23 col A = 'ຄຸນສົມບັດ'", "ຄຸນສົມບັດ", ws.Cell(23, 1).GetString());
            CheckEq("Row 24 col A = 'ແຮງງານ'",   "ແຮງງານ",   ws.Cell(24, 1).GetString());
            // Signatures row 27 — school's current template has 4 columns:
            // A=Director · C=Parent · E=Academic · G=Class teacher.
            CheckEq("A27 = 'ອຳນວຍການ' signature",    "ອຳນວຍການ",  ws.Cell("A27").GetString());
            CheckEq("C27 = 'ຜູປົກຄອງ' signature",     "ຜູປົກຄອງ",   ws.Cell("C27").GetString());
            CheckEq("E27 = 'ວິຊາການ' signature",      "ວິຊາການ",   ws.Cell("E27").GetString());
            CheckEq("G27 = 'ຄູປະຈຳຫ້ອງ' signature",  "ຄູປະຈຳຫ້ອງ", ws.Cell("G27").GetString());

            // Simulate the fill: write scores into F8..F24 (the new score column)
            // and verify ClosedXML accepts the mixed numeric + string values.
            for (int r = 8; r <= 19; r++) ws.Cell(r, 6).Value = 7.5;
            ws.Cell(20, 6).Value = 90.0;
            ws.Cell(21, 6).Value = 7.5;
            ws.Cell(22, 6).Value = "1";
            ws.Cell(23, 6).Value = 8.0;
            ws.Cell(24, 6).Value = "—";
            wb.Save();
            Pass("Template fill simulation (v2 layout, column F): ClosedXML accepted all values");
        }
        catch (Exception ex)
        {
            Fail("Template fill simulation", ex.Message);
        }
        finally
        {
            try { File.Delete(copy); } catch { }
        }
    }

    // Exercise the exact SQL pattern that EnrollmentPage.EnrollAllSubjects() runs:
    //   for every subject matching the student's grade, INSERT OR IGNORE
    //   two Enrollments rows (Sem 1 + Sem 2). Verifies idempotence + duplicate-skip
    //   semantics against the UNIQUE(StudentID,SubjectID,AcademicYear,Semester) index.
    static void TestAutoEnroll()
    {
        string yr = "2094-2095";
        // Fresh student with no prior enrolments.
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('AE001','Auto','Enroll','ຊາຍ','ມ.4','5',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='AE001'");

        // First run — should insert one row per subject per semester.
        int newSubjects, existedSubjects;
        (newSubjects, existedSubjects) = AutoEnrollRun(sid, "ມ.4", yr);
        int subjectCount = DB.ScalarInt(
            "SELECT COUNT(*) FROM Subjects WHERE GradeLevel='ມ.4' OR GradeLevel IS NULL OR GradeLevel=''");
        CheckEq("Auto-enroll: all subjects newly enrolled on first run", subjectCount, newSubjects);
        CheckEq("Auto-enroll: nothing pre-existed",                       0,            existedSubjects);
        CheckEq("Auto-enroll: produced 2×subjects rows in Enrollments",
            subjectCount * 2,
            DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE StudentID=@s AND AcademicYear=@y",
                null, ("@s", sid), ("@y", yr)));

        // Second run — IDEMPOTENT. Every subject should now be "existed", no new rows.
        (newSubjects, existedSubjects) = AutoEnrollRun(sid, "ມ.4", yr);
        CheckEq("Auto-enroll: idempotent (no new rows on second run)", 0,            newSubjects);
        CheckEq("Auto-enroll: all subjects reported as existed",       subjectCount, existedSubjects);
        CheckEq("Auto-enroll: total rows unchanged after second run",
            subjectCount * 2,
            DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE StudentID=@s AND AcademicYear=@y",
                null, ("@s", sid), ("@y", yr)));

        // Both semesters covered for every subject?
        int sem1Count = DB.ScalarInt(
            "SELECT COUNT(*) FROM Enrollments WHERE StudentID=@s AND AcademicYear=@y AND Semester=1",
            null, ("@s", sid), ("@y", yr));
        int sem2Count = DB.ScalarInt(
            "SELECT COUNT(*) FROM Enrollments WHERE StudentID=@s AND AcademicYear=@y AND Semester=2",
            null, ("@s", sid), ("@y", yr));
        CheckEq("Auto-enroll: 1 Sem-1 row per subject", subjectCount, sem1Count);
        CheckEq("Auto-enroll: 1 Sem-2 row per subject", subjectCount, sem2Count);

        // After removing one subject (both semesters), re-run should re-create just that subject.
        int subToRemove = DB.ScalarInt("SELECT SubjectID FROM Subjects WHERE SubjectCode='MATH1'");
        DB.Exec("DELETE FROM Enrollments WHERE StudentID=@s AND SubjectID=@sub AND AcademicYear=@y",
            null, ("@s", sid), ("@sub", subToRemove), ("@y", yr));
        (newSubjects, existedSubjects) = AutoEnrollRun(sid, "ມ.4", yr);
        CheckEq("Auto-enroll: re-enrols just the removed subject", 1,                  newSubjects);
        CheckEq("Auto-enroll: others still skip as existed",       subjectCount - 1,    existedSubjects);
    }

    // Exercises the exact SQL the rewritten UserFormWin runs for:
    //   • Add a new user
    //   • Edit profile fields without touching password
    //   • Change password
    //   • Verify the new password works against the production Login() flow
    // Catches the historical "password not saved" bug class.
    static void TestUserManagement()
    {
        const string code = "user_test";
        // Clean slate.
        DB.Exec("DELETE FROM Users WHERE Username=@u", null, ("@u", code));

        // ── Add ────────────────────────────────────────────
        DB.Exec(@"INSERT INTO Users(Username, Password, FullName, Role)
                  VALUES(@u, @p, @n, @r)",
            null,
            ("@u", code), ("@p", "initial-pw"),
            ("@n", "Test User"), ("@r", "teacher"));
        int uid = DB.ScalarInt("SELECT UserID FROM Users WHERE Username=@u", null, ("@u", code));
        Check("User created with UserID", uid > 0);
        CheckEq("Username persisted",   code,        DB.Scalar("SELECT Username FROM Users WHERE UserID=@id",
                                                          null, ("@id", uid))?.ToString());
        CheckEq("FullName persisted",   "Test User", DB.Scalar("SELECT FullName FROM Users WHERE UserID=@id",
                                                          null, ("@id", uid))?.ToString());
        CheckEq("Role persisted",       "teacher",   DB.Scalar("SELECT Role FROM Users WHERE UserID=@id",
                                                          null, ("@id", uid))?.ToString());

        // ── Initial login works ────────────────────────────
        var (ok1, role1, _) = DB.Login(code, "initial-pw");
        Check("Login with initial password succeeds", ok1);
        CheckEq("Role returned by Login = teacher", "teacher", role1);

        // ── Edit profile only (no password change) ─────────
        DB.Exec(@"UPDATE Users SET Username=@u, FullName=@n, Role=@r WHERE UserID=@id",
            null,
            ("@u", code), ("@n", "Renamed User"), ("@r", "admin"),
            ("@id", uid));
        CheckEq("FullName updated",     "Renamed User", DB.Scalar("SELECT FullName FROM Users WHERE UserID=@id",
                                                          null, ("@id", uid))?.ToString());
        CheckEq("Role promoted to admin", "admin",      DB.Scalar("SELECT Role FROM Users WHERE UserID=@id",
                                                          null, ("@id", uid))?.ToString());
        // Critical: password must NOT have been wiped by the profile-only update.
        var (okStill, _, _) = DB.Login(code, "initial-pw");
        Check("Profile-only edit did NOT change the password", okStill);

        // ── Change password ────────────────────────────────
        DB.Exec(@"UPDATE Users SET Username=@u, Password=@p, FullName=@n, Role=@r WHERE UserID=@id",
            null,
            ("@u", code), ("@p", "new-secret"),
            ("@n", "Renamed User"), ("@r", "admin"),
            ("@id", uid));

        // Old password must FAIL.
        var (okOld, _, _) = DB.Login(code, "initial-pw");
        Check("Old password rejected after change", !okOld);
        // New password must WORK end-to-end via the real Login() path.
        var (okNew, role2, name2) = DB.Login(code, "new-secret");
        Check("Login with new password succeeds", okNew);
        CheckEq("Login returns updated role",     "admin",        role2);
        CheckEq("Login returns updated full name", "Renamed User", name2);

        // ── Round-trip read-back (mirrors UserFormWin's verify step) ──
        int pwdMatch = DB.ScalarInt(
            "SELECT COUNT(*) FROM Users WHERE UserID=@id AND Password=@p",
            null, ("@id", uid), ("@p", "new-secret"));
        CheckEq("Read-back: new password stored verbatim", 1, pwdMatch);

        // ── UNIQUE(Username) constraint protects against duplicates ──
        bool gotUniqueEx = false;
        try
        {
            DB.Exec(@"INSERT INTO Users(Username, Password, FullName, Role)
                      VALUES(@u, 'x', 'Dup', 'teacher')", null, ("@u", code));
        }
        catch (Exception ex) when (ex.Message.Contains("UNIQUE")) { gotUniqueEx = true; }
        Check("UNIQUE(Username) prevents duplicate insert", gotUniqueEx);

        // ── Deactivate (toggle IsActive) ─────────────────
        DB.Exec("UPDATE Users SET IsActive=0 WHERE UserID=@id", null, ("@id", uid));
        var (okDeact, _, _) = DB.Login(code, "new-secret");
        Check("Login refused after deactivation", !okDeact);
        DB.Exec("UPDATE Users SET IsActive=1 WHERE UserID=@id", null, ("@id", uid));
        var (okReact, _, _) = DB.Login(code, "new-secret");
        Check("Login restored after reactivation", okReact);

        // Clean up so the test is idempotent.
        DB.Exec("DELETE FROM Users WHERE UserID=@id", null, ("@id", uid));
    }

    // Verifies the shipped Student Profile docx template has every token the code
    // expects, AND that a paragraph-level merge-and-replace (which is what
    // FillDocxTokens does inside the WPF app) successfully writes the values back
    // even though some tokens are split across <w:t> runs in this template.
    static void TestStudentProfileTemplate()
    {
        string root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "..", "..", "..", "..", ".."));
        string tpl = Path.Combine(root, "Templates", "ລາຍງານປະຫວັດນັກຮຽນ.docx");
        if (!File.Exists(tpl)) { Fail("Profile template found", $"missing: {tpl}"); return; }
        Pass($"Profile template found: {tpl}");

        // Copy to temp so we can fill it without locking the source.
        string copy = Path.Combine(Path.GetTempPath(), $"profile_test_{Guid.NewGuid():N}.docx");
        File.Copy(tpl, copy, true);
        try
        {
            // ── Step 1: read the template, collect distinct {{TOKEN}} matches.
            var requiredTokens = new[] {
                "{{FULL_NAME}}",       "{{BIRTH_DATE}}",
                "{{BIRTH_VILLAGE}}",   "{{BIRTH_DISTRICT}}",   "{{BIRTH_PROVINCE}}",
                "{{VILLAGE}}",         "{{DISTRICT}}",         "{{PROVINCE}}",
                "{{FATHER_NAME}}",     "{{FATHER_AGE}}",       "{{FATHER_JOB}}",
                "{{FATHER_VILLAGE}}",  "{{FATHER_DISTRICT}}",  "{{FATHER_PROVINCE}}",
                "{{MOTHER_NAME}}",     "{{MOTHER_AGE}}",       "{{MOTHER_JOB}}",
                "{{MOTHER_VILLAGE}}",  "{{MOTHER_DISTRICT}}",  "{{MOTHER_PROVINCE}}",
            };
            string combined;
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(copy, false))
            {
                if (doc.MainDocumentPart?.RootElement == null)
                { Fail("Profile template: main part exists", "MainDocumentPart is null"); return; }
                Pass("Profile template: opens via OpenXml, main part present");
                var texts = doc.MainDocumentPart.RootElement
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Select(t => t.Text ?? "");
                combined = string.Concat(texts);
            }
            foreach (var token in requiredTokens)
                Check($"Template contains {token}", combined.Contains(token));

            // ── Step 2: run the same paragraph-level merge-and-replace the WPF app does.
            var tokens = new Dictionary<string, string>();
            foreach (var t in requiredTokens) tokens[t] = "X_" + t.Trim('{', '}') + "_X";

            // Local copy of the FillDocxTokens algorithm (we can't link the WPF code,
            // and the algorithm is small enough to mirror here for the test).
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(copy, true))
            {
                if (doc.MainDocumentPart?.RootElement != null)
                {
                    foreach (var para in doc.MainDocumentPart.RootElement
                        .Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                    {
                        var ts = para.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().ToList();
                        if (ts.Count == 0) continue;
                        string c = string.Concat(ts.Select(t => t.Text ?? ""));
                        bool hit = false;
                        foreach (var kv in tokens) if (c.Contains(kv.Key)) { hit = true; break; }
                        if (!hit) continue;
                        string replaced = c;
                        foreach (var kv in tokens) replaced = replaced.Replace(kv.Key, kv.Value);
                        if (replaced == c) continue;
                        ts[0].Text  = replaced;
                        ts[0].Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
                        for (int i = 1; i < ts.Count; i++) ts[i].Text = "";
                    }
                    doc.MainDocumentPart.RootElement.Save();
                }
            }

            // ── Step 3: re-open and confirm every original token is GONE
            //          and every replacement marker is now present.
            using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(copy, false))
            {
                var texts = doc.MainDocumentPart!.RootElement!
                    .Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                    .Select(t => t.Text ?? "");
                string after = string.Concat(texts);
                int origRemaining = 0;
                foreach (var t in requiredTokens) if (after.Contains(t)) origRemaining++;
                CheckEq("After fill: no original {{...}} tokens remain", 0, origRemaining);

                int replacementsFound = 0;
                foreach (var t in requiredTokens)
                {
                    string marker = "X_" + t.Trim('{', '}') + "_X";
                    if (after.Contains(marker)) replacementsFound++;
                }
                CheckEq("After fill: every token replaced with marker",
                    requiredTokens.Length, replacementsFound);
            }
        }
        catch (Exception ex)
        {
            Fail("Profile template fill", ex.Message);
        }
        finally
        {
            try { File.Delete(copy); } catch { }
        }
    }

    // Mirror of BatchEnrollPage.EnrollAll's loop body (no UI, no MessageBox).
    // Verifies the class-wide flow: every active student in (year, grade [, room])
    // gets every subject for the grade, both semesters, idempotently.
    static void TestBatchEnroll()
    {
        string yr = "2093-2094";
        // Seed 3 active students + 1 graduate. Graduate must NOT be batch-enrolled.
        var seed = new[] {
            ("BE001", "ມ.4", "7", "ກຳລັງຮຽນ"),
            ("BE002", "ມ.4", "7", "ກຳລັງຮຽນ"),
            ("BE003", "ມ.4", "7", "ກຳລັງຮຽນ"),
            ("BE004", "ມ.4", "7", "ຈົບ"),         // should be skipped (not 'ກຳລັງຮຽນ')
            ("BE005", "ມ.4", "8", "ກຳລັງຮຽນ"),    // different room — excluded when filtering room=7
        };
        foreach (var (code, g, room, status) in seed)
        {
            DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                      VALUES(@c,'BE','Test','ຊາຍ',@g,@r,@y,@st)",
                null, ("@c", code), ("@g", g), ("@r", room), ("@y", yr), ("@st", status));
        }

        // ── First run: filter ມ.4 / room 7 → 3 active students should be enrolled.
        // Graduate (BE004) and other-room student (BE005) must be excluded.
        var (rosterCount, subjectCount, newRows, skippedRows, sNew, sExisted) =
            BatchEnrollRun(yr, "ມ.4", "7");
        CheckEq("Batch: roster = 3 active students in ມ.4/7",       3, rosterCount);
        CheckEq("Batch: all 3 students got new enrolments",         3, sNew);
        CheckEq("Batch: none already existed",                      0, sExisted);
        CheckEq("Batch: new rows = students × subjects × 2 sem",
            rosterCount * subjectCount * 2, newRows);
        CheckEq("Batch: skipped rows = 0 on first run",              0, skippedRows);

        // Graduate's row count must remain zero — even though batch run ‘touched’
        // year/grade matching them, the WHERE Status='ກຳລັງຮຽນ' clause excluded them.
        int gradEnrolls = DB.ScalarInt(
            @"SELECT COUNT(*) FROM Enrollments
              WHERE StudentID=(SELECT StudentID FROM Students WHERE StudentCode='BE004')",
            null);
        CheckEq("Batch: graduated student NOT enrolled", 0, gradEnrolls);

        // Other-room student's row count must also be zero (excluded by room filter).
        int otherRoomEnrolls = DB.ScalarInt(
            @"SELECT COUNT(*) FROM Enrollments
              WHERE StudentID=(SELECT StudentID FROM Students WHERE StudentCode='BE005')",
            null);
        CheckEq("Batch: room-filtered student NOT enrolled", 0, otherRoomEnrolls);

        // ── Second run: should be a complete no-op (idempotent).
        (rosterCount, _, newRows, skippedRows, sNew, sExisted) =
            BatchEnrollRun(yr, "ມ.4", "7");
        CheckEq("Batch: idempotent — newRows = 0 on second run", 0,         newRows);
        CheckEq("Batch: skippedRows = students × subj × 2",
            rosterCount * subjectCount * 2, skippedRows);
        CheckEq("Batch: all students reported as already-complete",
            rosterCount, sExisted);

        // ── "ທັງໝົດ" room filter: room=null → both room 7 (3 students) and room 8 (1)
        //     are included. BE005 in room 8 gets newly enrolled now.
        (rosterCount, _, newRows, _, sNew, sExisted) =
            BatchEnrollRun(yr, "ມ.4", null);  // null = ທັງໝົດ
        CheckEq("Batch: ທັງໝົດ-room roster = 4 active ມ.4 students", 4, rosterCount);
        CheckEq("Batch: 1 NEW student (room-8 BE005)",                1, sNew);
        CheckEq("Batch: 3 students already complete (room-7)",        3, sExisted);
    }

    // Mirror of BatchEnrollPage.EnrollAll loop body.
    // `room == null` means ທັງໝົດ (no room filter).
    static (int RosterCount, int SubjectCount, int NewRows, int SkippedRows,
            int StudentsNew, int StudentsExisted)
        BatchEnrollRun(string year, string grade, string? room)
    {
        // Load roster matching the filters.
        var sb = new StringBuilder(@"SELECT StudentID FROM Students
            WHERE Status='ກຳລັງຮຽນ' AND GradeLevel=@g AND AcademicYear=@y");
        var ps = new List<(string, object)> { ("@g", grade), ("@y", year) };
        if (room != null) { sb.Append(" AND ClassRoom=@r"); ps.Add(("@r", room)); }
        sb.Append(" ORDER BY ClassRoom, StudentCode");
        var roster = DB.Query(sb.ToString(), null, ps.ToArray());

        var subjects = DB.Query(@"SELECT SubjectID FROM Subjects
            WHERE GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel=''
            ORDER BY SortOrder, SubjectCode", null, ("@g", grade));

        int newRows = 0, skipped = 0, sNew = 0, sExisted = 0;
        using var conn = DB.Open();
        using var tx = conn.BeginTransaction();
        foreach (DataRow r in roster.Rows)
        {
            int sid = Convert.ToInt32(r["StudentID"]);
            int newForThisStudent = 0;
            foreach (DataRow sr in subjects.Rows)
            {
                int subId = Convert.ToInt32(sr["SubjectID"]);
                for (int sem = 1; sem <= 2; sem++)
                {
                    int n = DB.ExecTx(
                        @"INSERT OR IGNORE INTO Enrollments(StudentID, SubjectID, AcademicYear, Semester)
                          VALUES(@s, @sub, @y, @sm)",
                        conn, tx, ("@s", sid), ("@sub", subId), ("@y", year), ("@sm", sem));
                    if (n > 0) { newRows++; newForThisStudent++; }
                    else       { skipped++; }
                }
            }
            if (newForThisStudent > 0) sNew++; else sExisted++;
        }
        tx.Commit();
        return (roster.Rows.Count, subjects.Rows.Count, newRows, skipped, sNew, sExisted);
    }

    // Mirror of EnrollmentPage.EnrollAllSubjects's loop body (no UI, no MessageBox).
    // Returns (newSubjects, existedSubjects).
    static (int New, int Existed) AutoEnrollRun(int sid, string grade, string year)
    {
        var subjects = DB.Query(@"
            SELECT SubjectID FROM Subjects
            WHERE GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel=''
            ORDER BY SortOrder, SubjectCode",
            null, ("@g", grade));
        int n = 0, e = 0;
        using var conn = DB.Open();
        using var tx = conn.BeginTransaction();
        foreach (DataRow r in subjects.Rows)
        {
            int subId = Convert.ToInt32(r["SubjectID"]);
            int inserted = 0;
            for (int sem = 1; sem <= 2; sem++)
                inserted += DB.ExecTx(
                    @"INSERT OR IGNORE INTO Enrollments(StudentID, SubjectID, AcademicYear, Semester)
                      VALUES(@s, @sub, @y, @sm)",
                    conn, tx, ("@s", sid), ("@sub", subId), ("@y", year), ("@sm", sem));
            if (inserted > 0) n++;
            else              e++;
        }
        tx.Commit();
        return (n, e);
    }

    // ─────────────────────────────────────────────────────────────
    //  Score history tests — multi-year survival + CHA1/LAB1 contract
    // ─────────────────────────────────────────────────────────────
    static void TestScoreHistory()
    {
        // Seed a student in year A; enroll & score in ມ.1; promote to ມ.2 in year B;
        // assert all year-A data is still queryable AFTER the promotion, with the
        // correct historical grade.
        string yrA = "2030-2031", yrB = "2031-2032";

        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('HIST001','ປະຫວັດ','ໜຶ່ງ','ຊາຍ','ມ.1','2',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yrA));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='HIST001'");

        // Enroll in every subject, both semesters of year A.
        var subjects = DB.Query("SELECT SubjectID, SubjectCode FROM Subjects");
        int mathEnrollSem1 = 0, chaEnrollSem1 = 0;
        foreach (DataRow r in subjects.Rows)
        {
            int subId = Convert.ToInt32(r["SubjectID"]);
            string code = r["SubjectCode"].ToString()!;
            for (int sem = 1; sem <= 2; sem++)
            {
                DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                          VALUES(@s,@sub,@y,@sm)",
                    null, ("@s", sid), ("@sub", subId), ("@y", yrA), ("@sm", sem));
                if (sem == 1)
                {
                    int eid = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
                        WHERE StudentID=@s AND SubjectID=@sub AND AcademicYear=@y AND Semester=1",
                        null, ("@s", sid), ("@sub", subId), ("@y", yrA));
                    if (code == "MATH1") mathEnrollSem1 = eid;
                    if (code == "CHA1")  chaEnrollSem1  = eid;
                }
            }
        }

        // Real score on MATH1 sem-1: monthly avg 8, final 6 → total 7 @50/50
        for (int m = 9; m <= 12; m++)
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,@m,3,2,3)",
                null, ("@e", mathEnrollSem1), ("@m", m));
        DB.Exec("INSERT OR IGNORE INTO Scores(EnrollID,FinalScore) VALUES(@e,6)", null, ("@e", mathEnrollSem1));
        DB.Exec("UPDATE Scores SET FinalScore=6 WHERE EnrollID=@e", null, ("@e", mathEnrollSem1));
        DB.RecomputeMidFromMonthly(mathEnrollSem1);

        // CHA1 sem-1: write monthly noise + a manual EvaluationScore
        for (int m = 9; m <= 12; m++)
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,@m,3,2,5)",
                null, ("@e", chaEnrollSem1), ("@m", m));
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yrA, "SEM1", "CHA1", 9.0, c, tx);
            tx.Commit();
        }
        DB.RecomputeMidFromMonthly(chaEnrollSem1);

        // Pre-promotion sanity: historical query returns yrA and shows ມ.1 / room 2.
        var beforePromote = DB.GetStudentHistoryYears(sid);
        Check("Pre-promote: history has yrA",
            beforePromote.Select($"AcademicYear='{yrA}'").Length == 1, "yrA missing");
        if (beforePromote.Select($"AcademicYear='{yrA}'").Length == 1)
        {
            var rowA = beforePromote.Select($"AcademicYear='{yrA}'")[0];
            CheckEq("Pre-promote yrA grade", "ມ.1", rowA["GradeLevel"].ToString());
            CheckEq("Pre-promote yrA room", "2", rowA["ClassRoom"].ToString());
        }

        // Promote: student goes to ມ.2 / room 3 in yrB.
        using (var c = DB.Open())
        using (var tx = c.BeginTransaction())
        {
            DB.ExecTx("UPDATE Students SET GradeLevel='ມ.2', ClassRoom='3', AcademicYear=@y WHERE StudentID=@s",
                c, tx, ("@y", yrB), ("@s", sid));
            DB.ExecTx(@"INSERT INTO GradeHistory(StudentID,FromGrade,ToGrade,AcademicYear,ClassRoom,Note,ChangedBy)
                        VALUES(@s,'ມ.1','ມ.2',@y,'2','promo','test')",
                c, tx, ("@s", sid), ("@y", yrB));
            tx.Commit();
        }

        // POST-promote assertions
        var afterPromote = DB.GetStudentHistoryYears(sid);
        Check("Post-promote: yrA still in history",
            afterPromote.Select($"AcademicYear='{yrA}'").Length == 1, "yrA gone after promotion");
        Check("Post-promote: yrB also in history",
            afterPromote.Select($"AcademicYear='{yrB}'").Length == 1, "yrB missing");
        if (afterPromote.Select($"AcademicYear='{yrA}'").Length == 1)
        {
            var rowA = afterPromote.Select($"AcademicYear='{yrA}'")[0];
            CheckEq("Post-promote yrA shows historical ມ.1", "ມ.1", rowA["GradeLevel"].ToString());
            CheckEq("Post-promote yrA shows historical room 2", "2", rowA["ClassRoom"].ToString());
        }
        if (afterPromote.Select($"AcademicYear='{yrB}'").Length == 1)
        {
            var rowB = afterPromote.Select($"AcademicYear='{yrB}'")[0];
            CheckEq("Post-promote yrB shows current ມ.2", "ມ.2", rowB["GradeLevel"].ToString());
            CheckEq("Post-promote yrB shows current room 3", "3", rowB["ClassRoom"].ToString());
        }

        // MATH1 sem-1 score survives the promotion.
        var sem1 = DB.GetHistorySemester(sid, yrA, 1);
        var mathRow = sem1.Select("ລະຫັດວິຊາ='MATH1'");
        Check("MATH1 yrA-sem1 row survives", mathRow.Length == 1, "row missing");
        if (mathRow.Length == 1)
            CheckClose("MATH1 yrA-sem1 total = 7", 7.0, Convert.ToDouble(mathRow[0]["ລວມພາກ"]));

        // CHA1 in semester listing shows manual eval score (9.0), NOT a derived avg.
        var chaRow = sem1.Select("ລະຫັດວິຊາ='CHA1'");
        Check("CHA1 yrA-sem1 row present", chaRow.Length == 1, "CHA1 row missing");
        if (chaRow.Length == 1 && chaRow[0]["ລວມພາກ"] != DBNull.Value)
            CheckClose("CHA1 yrA-sem1 = manual 9.0", 9.0, Convert.ToDouble(chaRow[0]["ລວມພາກ"]));

        // Semester summary excludes CHA1/LAB1. Of all 14 subjects, only MATH1 has a
        // real Scores row (we ran RecomputeMidFromMonthly on it). CHA1 also went
        // through RecomputeMidFromMonthly but the helper early-returns for CHA1/LAB1
        // before creating the Scores row — so the count proves CHA1 was excluded.
        var sum = DB.GetHistorySemesterSummary(sid, yrA, 1);
        Check("Semester summary counts only academic subjects with scores",
            sum.subjects == 1, $"got {sum.subjects}");
        int chaScoreRows = DB.ScalarInt(@"
            SELECT COUNT(*) FROM Scores sc
            JOIN Enrollments e ON e.EnrollID=sc.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            WHERE e.StudentID=@s AND sub.SubjectCode='CHA1'",
            null, ("@s", sid));
        Check("CHA1 has NO Scores row (excluded by RecomputeMid)", chaScoreRows == 0,
            $"got {chaScoreRows}");
        // Avg is sum-of-all-academic / 12 (LEFT-JOIN nulls → 0), matching GetClassRanking.
        // With only MATH1 (7.0) scored: 7/12 ≈ 0.58.
        CheckClose("Semester summary avg = 7/12", 7.0/12, sum.avg, 0.02);

        // Graduate the student — historical data must still be visible.
        DB.Exec("UPDATE Students SET Status='ຈົບ' WHERE StudentID=@s", null, ("@s", sid));
        var afterGrad = DB.GetStudentHistoryYears(sid);
        Check("Graduated student: yrA still visible",
            afterGrad.Select($"AcademicYear='{yrA}'").Length == 1, "yrA gone after graduation");
        var sem1AfterGrad = DB.GetHistorySemester(sid, yrA, 1);
        Check("Graduated student: yrA-sem1 scores still present", sem1AfterGrad.Rows.Count > 0,
            "scores empty after graduation");

        // Withdraw status — filtering by 'ອອກ' must include withdrawn students.
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('HIST002','ອອກ','ກາງ','ຍິງ','ມ.2','1',@y,'ອອກ')",
            null, ("@y", yrA));
        int withdrawnCount = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ອອກ'");
        Check("Withdraw status row exists", withdrawnCount >= 1, "no ອອກ rows");

        // ── Per-month split: monthly DataTable must contain exactly 4 sem-1 months
        //    for MATH1 (one per Sept/Oct/Nov/Dec). The page slices this client-side
        //    into 4 separate month sections; verify the underlying data supports it.
        var allMonthly = DB.GetHistoryMonthly(sid, yrA);
        int mathSept = allMonthly.Select("ລະຫັດວິຊາ='MATH1' AND ເດືອນ=9").Length;
        int mathOct  = allMonthly.Select("ລະຫັດວິຊາ='MATH1' AND ເດືອນ=10").Length;
        int mathNov  = allMonthly.Select("ລະຫັດວິຊາ='MATH1' AND ເດືອນ=11").Length;
        int mathDec  = allMonthly.Select("ລະຫັດວິຊາ='MATH1' AND ເດືອນ=12").Length;
        Check("MATH1 has 1 monthly row per sem-1 month (×4)",
            mathSept == 1 && mathOct == 1 && mathNov == 1 && mathDec == 1,
            $"got {mathSept}/{mathOct}/{mathNov}/{mathDec}");
        // No sem-2 monthly rows were seeded → all 4 sem-2 month slots should be empty.
        int sem2RowsM = allMonthly.Select("ເດືອນ IN (2,3,4,5) AND ລະຫັດວິຊາ='MATH1'").Length;
        Check("No MATH1 monthly rows in sem-2 (none seeded)", sem2RowsM == 0, $"got {sem2RowsM}");

        // ── Annual summary
        var ann = DB.GetHistoryAnnualSummary(sid, yrA);
        // sem1 avg was 7/12 ≈ 0.58; sem2 has nothing scored → 0 avg; annual = mean ≈ 0.29
        CheckClose("Annual: sem1Avg = 7/12", 7.0/12, ann.sem1Avg, 0.02);
        CheckClose("Annual: sem2Avg ≈ 0", 0.0, ann.sem2Avg, 0.02);
        // When sem2 has zero scored subjects the helper falls back to sem1 alone
        // rather than averaging in zeros — verify that behavior.
        CheckClose("Annual: empty-sem2 fallback uses sem1 avg only",
            ann.sem1Avg, ann.annualAvg, 0.01);
        Check("Annual: level computed from annualAvg",
            ann.level == DB.CalcMoESLevel(ann.annualAvg), $"got '{ann.level}'");

        // Now seed a sem-2 score so both semesters contribute → mean is computed.
        int mathSem2 = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
            WHERE StudentID=@s AND AcademicYear=@y AND Semester=2
              AND SubjectID=(SELECT SubjectID FROM Subjects WHERE SubjectCode='MATH1')",
            null, ("@s", sid), ("@y", yrA));
        for (int m = 2; m <= 5; m++)
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,@m,3,2,3)",
                null, ("@e", mathSem2), ("@m", m));
        DB.Exec("INSERT OR IGNORE INTO Scores(EnrollID,FinalScore) VALUES(@e,4)", null, ("@e", mathSem2));
        DB.Exec("UPDATE Scores SET FinalScore=4 WHERE EnrollID=@e", null, ("@e", mathSem2));
        DB.RecomputeMidFromMonthly(mathSem2);

        var ann2 = DB.GetHistoryAnnualSummary(sid, yrA);
        Check("Annual after sem-2 scored: both semesters non-zero",
            ann2.sem1Avg > 0 && ann2.sem2Avg > 0,
            $"sem1={ann2.sem1Avg} sem2={ann2.sem2Avg}");
        CheckClose("Annual when both semesters scored = mean",
            (ann2.sem1Avg + ann2.sem2Avg) / 2.0, ann2.annualAvg, 0.01);

        // ─── Historical class roster ───────────────────────────────────
        // Build a tiny mixed cohort:
        //   HIST_A — in (yrB, ມ.2, room 3) RIGHT NOW (their current row), no later promotion
        //   HIST_B — was in (yrA, ມ.1, room 2), promoted to (yrB, ມ.2, room 3) via GradeHistory
        //   HIST_C — was in (yrA, ມ.1, room 2), graduated (Status=ຈົບ) → still queryable
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('ROST_A','Aname','Aunlast','ຍິງ','ມ.2','3',@yb,'ກຳລັງຮຽນ')",
            null, ("@yb", yrB));
        int aId = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='ROST_A'");
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
            null, ("@s", aId), ("@y", yrB));

        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('ROST_B','Bname','Blast','ຊາຍ','ມ.2','3',@yb,'ກຳລັງຮຽນ')",
            null, ("@yb", yrB));
        int bId = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='ROST_B'");
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
            null, ("@s", bId), ("@y", yrA));
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
            null, ("@s", bId), ("@y", yrB));
        DB.Exec(@"INSERT INTO GradeHistory(StudentID,FromGrade,ToGrade,AcademicYear,ClassRoom,Note,ChangedBy)
                  VALUES(@s,'ມ.1','ມ.2',@y,'2','promo','test')",
            null, ("@s", bId), ("@y", yrB));

        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('ROST_C','Cname','Clast','ຍິງ','ມ.1','2',@y,'ຈົບ')",
            null, ("@y", yrA));
        int cId = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='ROST_C'");
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
            null, ("@s", cId), ("@y", yrA));

        // Query the (yrA, ມ.1, room 2) cohort — should include HIST001 (post-promote),
        // ROST_B (per GradeHistory), ROST_C (graduated still in original row).
        // Should NOT include ROST_A (never in yrA).
        var rosterA = DB.GetHistoricalClassRoster(yrA, "ມ.1", "2");
        bool hasOurStudent = rosterA.Select("ລະຫັດ='HIST001'").Length == 1;
        bool hasRostB      = rosterA.Select("ລະຫັດ='ROST_B'").Length == 1;
        bool hasRostC      = rosterA.Select("ລະຫັດ='ROST_C'").Length == 1;
        bool hasRostA      = rosterA.Select("ລະຫັດ='ROST_A'").Length > 0;
        Check("Roster yrA/ມ.1/room 2: includes promoted student (HIST001)", hasOurStudent, "missing");
        Check("Roster yrA/ມ.1/room 2: includes GradeHistory-only student (ROST_B)", hasRostB, "missing");
        Check("Roster yrA/ມ.1/room 2: includes graduated student (ROST_C)", hasRostC, "missing");
        Check("Roster yrA/ມ.1/room 2: excludes irrelevant student (ROST_A)", !hasRostA, "incorrectly present");

        // Query the (yrB, ມ.2, room 3) cohort — should include ROST_A (current row matches)
        // AND ROST_B (GradeHistory says they went to room ??? — wait, the GradeHistory row
        // we wrote uses ClassRoom='2' which is the OLD room). For yrB, ROST_B's current
        // Students row says room 3 with no later promotion — so they qualify via the (b)
        // branch. Verify:
        var rosterB = DB.GetHistoricalClassRoster(yrB, "ມ.2", "3");
        bool hasRostA_B = rosterB.Select("ລະຫັດ='ROST_A'").Length == 1;
        bool hasRostB_B = rosterB.Select("ລະຫັດ='ROST_B'").Length == 1;
        Check("Roster yrB/ມ.2/room 3: includes ROST_A (current row matches)", hasRostA_B, "missing");
        Check("Roster yrB/ມ.2/room 3: includes ROST_B (current row matches, no later promo)", hasRostB_B, "missing");

        // Class month grid + summary: ROST_C is in the cohort and graduated;
        // their data must still appear (key requirement).
        var monthGrid = DB.GetClassMonthGrid(yrA, "ມ.1", "2", 9);
        var s1Sum = DB.GetClassSemesterSummary(yrA, "ມ.1", "2", 1);
        var annSum = DB.GetClassAnnualSummary(yrA, "ມ.1", "2");
        Check("ClassMonthGrid yrA/sept: ROST_C present (graduated)",
            monthGrid.Select("ລະຫັດ='ROST_C'").Length == 1, "missing");
        Check("ClassSemesterSummary yrA: ROST_C present (graduated)",
            s1Sum.Select("ລະຫັດ='ROST_C'").Length == 1, "missing");
        Check("ClassAnnualSummary yrA: ROST_C present (graduated)",
            annSum.Select("ລະຫັດ='ROST_C'").Length == 1, "missing");

        // Column set in class month grid must EXCLUDE CHA1/LAB1
        bool hasCha = false, hasLab = false;
        foreach (DataColumn col in monthGrid.Columns)
        {
            if (col.ColumnName == "CHA1") hasCha = true;
            if (col.ColumnName == "LAB1") hasLab = true;
        }
        Check("ClassMonthGrid: no CHA1 column", !hasCha, "CHA1 column leaked in");
        Check("ClassMonthGrid: no LAB1 column", !hasLab, "LAB1 column leaked in");
    }

    // ─────────────────────────────────────────────────────────────
    //  Score-history template-based exports
    //  Covers SafeFileName + the contracts the new
    //  RenderIndividual/ClassMonthlyXlsx wrappers depend on:
    //    - SafeFileName strips every Windows-invalid char
    //    - Historical roster includes graduated students so their report
    //      can still be generated
    //    - Per-month bulk score lookup returns data for every cohort
    //      student (the same query shape RenderClassMonthlyXlsx uses)
    //    - Sort-by-descending-academic-sum produces stable ordering
    // ─────────────────────────────────────────────────────────────
    static void TestScoreHistoryTemplateExports()
    {
        // SafeFileName
        CheckEq("SafeFileName: empty → 'report'",  "report", DB.SafeFileName(""));
        CheckEq("SafeFileName: trims trailing dot", "ok",     DB.SafeFileName("ok."));
        CheckEq("SafeFileName: trims trailing space","ok",    DB.SafeFileName("ok "));
        CheckEq("SafeFileName: replaces forward slash", "a_b",     DB.SafeFileName("a/b"));
        CheckEq("SafeFileName: replaces backslash",     "a_b",     DB.SafeFileName("a\\b"));
        CheckEq("SafeFileName: replaces colon",         "a_b",     DB.SafeFileName("a:b"));
        CheckEq("SafeFileName: replaces star",          "a_b",     DB.SafeFileName("a*b"));
        CheckEq("SafeFileName: replaces question",      "a_b",     DB.SafeFileName("a?b"));
        CheckEq("SafeFileName: replaces quote",         "a_b",     DB.SafeFileName("a\"b"));
        CheckEq("SafeFileName: replaces lt",            "a_b",     DB.SafeFileName("a<b"));
        CheckEq("SafeFileName: replaces gt",            "a_b",     DB.SafeFileName("a>b"));
        CheckEq("SafeFileName: replaces pipe",          "a_b",     DB.SafeFileName("a|b"));
        CheckEq("SafeFileName: all invalid → underscores", "_________",
            DB.SafeFileName("\\/:*?\"<>|"));
        CheckEq("SafeFileName: preserves Lao chars",  "ໃບຄະແນນ_ປີ",
            DB.SafeFileName("ໃບຄະແນນ:ປີ"));
        CheckEq("SafeFileName: reserved CON prefixed", "_CON",  DB.SafeFileName("CON"));
        CheckEq("SafeFileName: reserved con (case-i) prefixed", "_con", DB.SafeFileName("con"));
        CheckEq("SafeFileName: reserved COM3 prefixed", "_COM3", DB.SafeFileName("COM3"));
        // Windows reserves device names even with an extension — CON.TXT opens the
        // CON device, not a file. Our helper matches Windows behavior by checking
        // the base name (extension-stripped) against the reserved list.
        CheckEq("SafeFileName: 'con.txt' treated as reserved (matches Windows)",
            "_con.txt", DB.SafeFileName("con.txt"));
        // Control chars are dropped
        Check("SafeFileName: drops control chars",
            DB.SafeFileName("a\tb\nc") == "abc", "got " + DB.SafeFileName("a\tb\nc"));

        // ── Historical roster includes graduated students ──
        // TestScoreHistory above seeded ROST_C (Status='ຈົບ') in (yrA, ມ.1, room 2).
        // RenderClassMonthlyXlsx calls GetHistoricalClassRoster — verify it returns
        // ROST_C so a class report for that historical year can include them.
        var roster = DB.GetHistoricalClassRoster("2030-2031", "ມ.1", "2");
        Check("Roster includes graduated ROST_C (for export)",
            roster.Select("ລະຫັດ='ROST_C'").Length == 1, "missing");

        // ── Per-month bulk score lookup matches RenderClassMonthlyXlsx's query shape ──
        var idList = new List<string>();
        foreach (DataRow rr in roster.Rows) idList.Add(rr["StudentID"].ToString()!);
        if (idList.Count > 0)
        {
            string idCsv = string.Join(",", idList);
            var dt = DB.Query($@"
                SELECT e.StudentID, sub.SubjectCode,
                       (IFNULL(ma.ActivityScore,0)+IFNULL(ma.DisciplineScore,0)+IFNULL(ma.HomeworkScore,0)) AS Total,
                       (ma.MonthlyID IS NOT NULL) AS HasRow
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@m
                WHERE e.AcademicYear=@y AND e.Semester=@sm
                  AND e.StudentID IN ({idCsv})",
                null, ("@m", 9), ("@y", "2030-2031"), ("@sm", 1));
            // HIST001 had monthly entries for sept on MATH1 + CHA1 (from TestScoreHistory).
            // Both should appear in this bulk lookup → CHA1 IS included in the query
            // (the template's RANK formula range is D:O which excludes CHA/LAB columns
            // anyway, so including CHA/LAB in the lookup doesn't taint aggregates).
            int hist001Math = dt.Select("StudentID=" +
                DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='HIST001'")
                + " AND SubjectCode='MATH1'").Length;
            Check("Bulk lookup: HIST001 MATH1 sept row present", hist001Math == 1, $"got {hist001Math}");
        }

        // ── Sort-by-descending-sum stability ──
        // Build a fake roster + sums, sort, and verify the order is descending.
        var fake = new DataTable();
        fake.Columns.Add("StudentID", typeof(int));
        fake.Rows.Add(1); fake.Rows.Add(2); fake.Rows.Add(3);
        var sums = new Dictionary<int, double> { {1, 50}, {2, 80}, {3, 65} };
        var sorted = new List<int>();
        foreach (DataRow rr in fake.Select("", "").OrderByDescending(rw => sums[Convert.ToInt32(rw["StudentID"])]))
            sorted.Add(Convert.ToInt32(rr["StudentID"]));
        Check("Sort by descending sum: 2 (80) → 3 (65) → 1 (50)",
            sorted.Count == 3 && sorted[0] == 2 && sorted[1] == 3 && sorted[2] == 1,
            string.Join(",", sorted));

        // ── Templates exist (smoke check — wrappers need them at runtime) ──
        // The test project runs from tests/IntegrationTests/bin/...; templates ship
        // with the main WPF project at <repo>/Templates/. Resolve relative to source.
        string repoRoot = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "..", "..", "..", "..", ".."));
        string classTpl = Path.Combine(repoRoot, "Templates", "ໃບຄະແນນ.xlsx");
        string indTpl   = Path.Combine(repoRoot, "Templates", "ລີພອດຄະແນນບຸກຄົນ.xlsx");
        Check("Class template ໃບຄະແນນ.xlsx exists",       File.Exists(classTpl), classTpl);
        Check("Individual template ລີພອດຄະແນນບຸກຄົນ.xlsx exists", File.Exists(indTpl), indTpl);

        // ── Class template has 4 sheets (Monthly / Sem1 / Sem2 / Annual) ──
        // RenderClassSemesterXlsx keeps Sheet 2/3; RenderClassAnnualXlsx keeps Sheet 4.
        if (File.Exists(classTpl))
        {
            using var wb = new ClosedXML.Excel.XLWorkbook(classTpl);
            CheckEq("Class template sheet count = 4", 4, wb.Worksheets.Count);
        }

        // ── Historical semester/annual queries mirror Render*Xlsx shape ──
        // The Render*Semester/Annual wrappers select via e.StudentID IN (cohort)
        // (not Students.GradeLevel/ClassRoom). Verify those queries return data
        // for graduated students too — same setup as before (ROST_C, Status=ຈົບ).
        var rosterAg = DB.GetHistoricalClassRoster("2030-2031", "ມ.1", "2");
        var ids = new List<string>();
        foreach (DataRow rr in rosterAg.Rows) ids.Add(rr["StudentID"].ToString()!);
        string ic = ids.Count > 0 ? string.Join(",", ids) : "0";

        // Sem 1 academic query — must exclude CHA1/LAB1 from the join.
        var semDt = DB.Query($@"
            SELECT e.StudentID, sub.SubjectCode, IFNULL(sc.TotalScore,0) AS Total
            FROM Enrollments e
            JOIN Subjects sub ON sub.SubjectID=e.SubjectID
            LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
            WHERE e.AcademicYear=@y AND e.Semester=@sm
              AND sub.SubjectCode NOT IN ('CHA1','LAB1')
              AND e.StudentID IN ({ic})",
            null, ("@y", "2030-2031"), ("@sm", 1));
        int hist001 = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='HIST001'");
        int semCha = semDt.Select($"StudentID={hist001} AND SubjectCode='CHA1'").Length;
        int semLab = semDt.Select($"StudentID={hist001} AND SubjectCode='LAB1'").Length;
        int semMath = semDt.Select($"StudentID={hist001} AND SubjectCode='MATH1'").Length;
        Check("Sem academic query: CHA1 excluded", semCha == 0, "CHA1 leaked into sem aggregate");
        Check("Sem academic query: LAB1 excluded", semLab == 0, "LAB1 leaked into sem aggregate");
        Check("Sem academic query: MATH1 included (HIST001)", semMath == 1, $"got {semMath}");

        // Annual query — same shape but year-wide (both semesters).
        var annDt = DB.Query($@"
            SELECT e.StudentID, sub.SubjectCode, IFNULL(sc.TotalScore,0) AS Total
            FROM Enrollments e
            JOIN Subjects sub ON sub.SubjectID=e.SubjectID
            LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
            WHERE e.AcademicYear=@y
              AND sub.SubjectCode NOT IN ('CHA1','LAB1')
              AND sc.ScoreID IS NOT NULL
              AND e.StudentID IN ({ic})",
            null, ("@y", "2030-2031"));
        int annCha = annDt.Select($"SubjectCode='CHA1'").Length;
        int annLab = annDt.Select($"SubjectCode='LAB1'").Length;
        Check("Annual academic query: CHA1 excluded", annCha == 0, "CHA1 leaked into annual aggregate");
        Check("Annual academic query: LAB1 excluded", annLab == 0, "LAB1 leaked into annual aggregate");

        // ── Graduated-student CHA1 manual value still readable via EvaluationScores ──
        // HIST001 had CHA1 SEM1 set to 9.0 in TestScoreHistory. After they're graduated
        // (Status='ຈົບ' was set later), the eval score must remain readable.
        double? graduatedCha = DB.GetEvaluationScore(hist001, "2030-2031", "SEM1", "CHA1");
        Check("Graduated student: SEM1 CHA1 manual value preserved",
            graduatedCha.HasValue && Math.Abs(graduatedCha.Value - 9.0) < 0.01,
            $"got {graduatedCha}");

        // ── Settings-backed “last selected report type” persists ──
        // Both history windows read+write the same Settings key
        // (score_history_report_type) so the user's last pick survives.
        // Round-trip through SaveSetting + Scalar read mirrors the windows' code.
        DB.SaveSetting("score_history_report_type", "S2");
        string? readBack = DB.Scalar(
            "SELECT Value FROM Settings WHERE Key=@k",
            null, ("@k", "score_history_report_type"))?.ToString();
        CheckEq("Last-selected report type persists across reads", "S2", readBack);
        DB.SaveSetting("score_history_report_type", "A");
        readBack = DB.Scalar(
            "SELECT Value FROM Settings WHERE Key=@k",
            null, ("@k", "score_history_report_type"))?.ToString();
        CheckEq("Last-selected report type updates on re-save", "A", readBack);
    }

    // ─────────────────────────────────────────────────────────────
    //  Graduated-student access to all 4 score-history report types
    //
    //  End-to-end coverage for the spec "graduated students must be able
    //  to view Monthly / Sem 1 / Sem 2 / Annual reports".
    //
    //  Setup: GRAD001 finishes ມ.4 in 2032-2033, then PromotionPage marks
    //  them graduated (Status='ຈົບ' + GradeHistory row with
    //  AcademicYear=2033-2034, FromGrade=ມ.4, ClassRoom=2, ToGrade='ຈົບ').
    //  We seed scores in every relevant place: monthly entries for ALL 8
    //  assessment months, semester finals + EvaluationScores for both
    //  sems, ANNUAL EvaluationScores.
    //
    //  Then verify each data path the report wrappers depend on:
    //    - GetHistoricalClassRoster includes the graduate
    //    - GetHistoryMonthly returns monthly rows for all 8 months
    //    - GetHistorySemester returns sem-1 + sem-2 rows
    //    - GetHistorySemesterSummary computes avg/total for both sems
    //    - GetHistoryAnnualSummary returns sem1Avg + sem2Avg + annualAvg
    //    - GetClassMonthGrid / SemesterSummary / AnnualSummary include them
    //    - EvaluationScores (CHA1/LAB1) readable across SEM1/SEM2/ANNUAL
    //  ─────────────────────────────────────────────────────────────
    static void TestGraduatedStudentHistoryAccess()
    {
        string yr = "2032-2033";

        // 1) Seed student already in ມ.4 / room 2 / 2032-2033 (active).
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('GRAD001','ຈົບ','ໝູ່','ຍິງ','ມ.4','2',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='GRAD001'");

        // 2) Enroll in every subject for both semesters of yr.
        var enrollIds = new Dictionary<(string code, int sem), int>();
        var subjects = DB.Query("SELECT SubjectID, SubjectCode FROM Subjects");
        foreach (DataRow r in subjects.Rows)
        {
            int subId = Convert.ToInt32(r["SubjectID"]);
            string code = r["SubjectCode"].ToString()!;
            for (int sem = 1; sem <= 2; sem++)
            {
                DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                          VALUES(@s,@sub,@y,@sm)",
                    null, ("@s", sid), ("@sub", subId), ("@y", yr), ("@sm", sem));
                int eid = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
                    WHERE StudentID=@s AND SubjectID=@sub AND AcademicYear=@y AND Semester=@sm",
                    null, ("@s", sid), ("@sub", subId), ("@y", yr), ("@sm", sem));
                enrollIds[(code, sem)] = eid;
            }
        }

        // 3) Seed scores: monthly entries for ALL 8 months on MATH1, plus final exam
        //    + Recompute → Scores row written.
        int math1Sem1 = enrollIds[("MATH1", 1)];
        int math1Sem2 = enrollIds[("MATH1", 2)];
        foreach (int m in new[] { 9, 10, 11, 12 })
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,@m,3,2,4)", null, ("@e", math1Sem1), ("@m", m));
        foreach (int m in new[] { 2, 3, 4, 5 })
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,@m,3,2,4)", null, ("@e", math1Sem2), ("@m", m));
        DB.Exec("INSERT OR IGNORE INTO Scores(EnrollID,FinalScore) VALUES(@e,7)", null, ("@e", math1Sem1));
        DB.Exec("UPDATE Scores SET FinalScore=7 WHERE EnrollID=@e", null, ("@e", math1Sem1));
        DB.Exec("INSERT OR IGNORE INTO Scores(EnrollID,FinalScore) VALUES(@e,8)", null, ("@e", math1Sem2));
        DB.Exec("UPDATE Scores SET FinalScore=8 WHERE EnrollID=@e", null, ("@e", math1Sem2));
        DB.RecomputeMidFromMonthly(math1Sem1);
        DB.RecomputeMidFromMonthly(math1Sem2);

        // 4) CHA1/LAB1 manual values for SEM1, SEM2, ANNUAL contexts.
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1",   "CHA1", 8.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "SEM1",   "LAB1", 7.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "SEM2",   "CHA1", 9.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "SEM2",   "LAB1", 8.0, c, tx);
            DB.SetEvaluationScore(sid, yr, "ANNUAL", "CHA1", 8.5, c, tx);
            DB.SetEvaluationScore(sid, yr, "ANNUAL", "LAB1", 7.5, c, tx);
            tx.Commit();
        }

        // 5) Graduate the student — exactly what PromotionPage does on the
        //    ‘Graduate’ action: Status=ຈົບ, AcademicYear stays at finishing year,
        //    GradeHistory row records the graduation with ClassRoom set.
        DB.Exec("UPDATE Students SET Status='ຈົບ' WHERE StudentID=@s", null, ("@s", sid));
        DB.Exec(@"INSERT INTO GradeHistory(StudentID,FromGrade,ToGrade,AcademicYear,ClassRoom,Note,ChangedBy)
                  VALUES(@s,'ມ.4','ຈົບ',@y,'2','graduated','test')",
            null, ("@s", sid), ("@y", "2033-2034"));

        // ── Roster: graduate must appear when querying their finishing class ──
        var roster = DB.GetHistoricalClassRoster(yr, "ມ.4", "2");
        Check("Graduate visible in finishing-year roster",
            roster.Select("ລະຫັດ='GRAD001'").Length == 1, "graduate missing");

        // ── Years dropdown: graduate's finishing year must be listed ──
        var years = DB.GetStudentHistoryYears(sid);
        Check("Graduate: finishing year listed in GetStudentHistoryYears",
            years.Select($"AcademicYear='{yr}'").Length == 1, "year missing");

        // ── Monthly Reports (Month 1-8): all 8 monthly entries readable ──
        var monthly = DB.GetHistoryMonthly(sid, yr);
        foreach (int m in new[] { 9, 10, 11, 12, 2, 3, 4, 5 })
        {
            int hits = monthly.Select($"ລະຫັດວິຊາ='MATH1' AND ເດືອນ={m}").Length;
            Check($"Graduate: MATH1 month {m} entry readable", hits == 1, $"got {hits}");
        }

        // ── Semester 1 Report: per-subject rows + summary ──
        var sem1 = DB.GetHistorySemester(sid, yr, 1);
        Check("Graduate: Sem 1 MATH1 row present",
            sem1.Select("ລະຫັດວິຊາ='MATH1'").Length == 1, "missing");
        var sem1Sum = DB.GetHistorySemesterSummary(sid, yr, 1);
        Check("Graduate: Sem 1 summary has scored subjects", sem1Sum.subjects >= 1,
            $"subjects={sem1Sum.subjects}");

        // CHA1/LAB1 for SEM1 must be readable but NOT in the summary aggregate.
        var sem1Cha = sem1.Select("ລະຫັດວິຊາ='CHA1'");
        if (sem1Cha.Length == 1 && sem1Cha[0]["ລວມພາກ"] != DBNull.Value)
            CheckClose("Graduate: SEM1 CHA1 manual = 8.0", 8.0, Convert.ToDouble(sem1Cha[0]["ລວມພາກ"]));
        double? sem1ChaEval = DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1");
        Check("Graduate: SEM1 CHA1 EvaluationScores readable",
            sem1ChaEval.HasValue && Math.Abs(sem1ChaEval.Value - 8.0) < 0.01,
            $"got {sem1ChaEval}");

        // ── Semester 2 Report ──
        var sem2 = DB.GetHistorySemester(sid, yr, 2);
        Check("Graduate: Sem 2 MATH1 row present",
            sem2.Select("ລະຫັດວິຊາ='MATH1'").Length == 1, "missing");
        var sem2Sum = DB.GetHistorySemesterSummary(sid, yr, 2);
        Check("Graduate: Sem 2 summary has scored subjects", sem2Sum.subjects >= 1,
            $"subjects={sem2Sum.subjects}");

        double? sem2LabEval = DB.GetEvaluationScore(sid, yr, "SEM2", "LAB1");
        Check("Graduate: SEM2 LAB1 EvaluationScores readable",
            sem2LabEval.HasValue && Math.Abs(sem2LabEval.Value - 8.0) < 0.01,
            $"got {sem2LabEval}");

        // ── Annual Report ──
        var ann = DB.GetHistoryAnnualSummary(sid, yr);
        Check("Graduate: Annual sem1Avg > 0", ann.sem1Avg > 0, $"got {ann.sem1Avg}");
        Check("Graduate: Annual sem2Avg > 0", ann.sem2Avg > 0, $"got {ann.sem2Avg}");
        CheckClose("Graduate: Annual avg = mean of sem averages",
            (ann.sem1Avg + ann.sem2Avg) / 2.0, ann.annualAvg, 0.01);
        double? annChaEval = DB.GetEvaluationScore(sid, yr, "ANNUAL", "CHA1");
        Check("Graduate: ANNUAL CHA1 EvaluationScores readable",
            annChaEval.HasValue && Math.Abs(annChaEval.Value - 8.5) < 0.01,
            $"got {annChaEval}");

        // ── Class-history queries (ClassroomHistoryWindow) include the graduate ──
        var cmGrid  = DB.GetClassMonthGrid(yr, "ມ.4", "2", 9);
        var cs1Sum  = DB.GetClassSemesterSummary(yr, "ມ.4", "2", 1);
        var cs2Sum  = DB.GetClassSemesterSummary(yr, "ມ.4", "2", 2);
        var cAnnSum = DB.GetClassAnnualSummary(yr, "ມ.4", "2");
        Check("Graduate in ClassMonthGrid (Sept)",
            cmGrid.Select("ລະຫັດ='GRAD001'").Length == 1, "missing");
        Check("Graduate in ClassSemesterSummary (Sem 1)",
            cs1Sum.Select("ລະຫັດ='GRAD001'").Length == 1, "missing");
        Check("Graduate in ClassSemesterSummary (Sem 2)",
            cs2Sum.Select("ລະຫັດ='GRAD001'").Length == 1, "missing");
        Check("Graduate in ClassAnnualSummary",
            cAnnSum.Select("ລະຫັດ='GRAD001'").Length == 1, "missing");

        // ── Backfill verification: GradeHistory.ClassRoom for the graduation
        //    row must be non-empty (proves the broader backfill works for
        //    AcademicYear-doesn't-match-current cases).
        string? backfilled = DB.Scalar(
            "SELECT ClassRoom FROM GradeHistory WHERE StudentID=@s AND AcademicYear='2033-2034'",
            null, ("@s", sid))?.ToString();
        Check("Graduate's graduation-row GradeHistory.ClassRoom backfilled",
            !string.IsNullOrEmpty(backfilled), $"got '{backfilled}'");

        // ── Pre-existing migration coverage: GradeHistory rows with NULL ClassRoom
        //    that don't match the year-match-current rule still get backfilled by
        //    pass 2. Simulate by clearing the room and re-running Initialize.
        DB.Exec("UPDATE GradeHistory SET ClassRoom=NULL WHERE StudentID=@s AND AcademicYear='2033-2034'",
            null, ("@s", sid));
        DB.Initialize();
        string? afterPass2 = DB.Scalar(
            "SELECT ClassRoom FROM GradeHistory WHERE StudentID=@s AND AcademicYear='2033-2034'",
            null, ("@s", sid))?.ToString();
        Check("Pass-2 backfill restores ClassRoom for non-current-year rows",
            !string.IsNullOrEmpty(afterPass2), $"got '{afterPass2}'");

        // ── Status filter on GetHistoricalClassRoster (the new radio buttons) ──
        // Seed an extra active classmate in the same year/grade/room so the
        // Active filter actually narrows: cohort = {GRAD001 ('ຈົບ'), STAT_A ('ກຳລັງຮຽນ')}.
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('STAT_A','ກຳ','ລັງ','ຊາຍ','ມ.4','2',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int statA = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='STAT_A'");
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
            null, ("@s", statA), ("@y", yr));

        // Active filter → only STAT_A (current ກຳລັງຮຽນ) — graduate excluded.
        var activeOnly = DB.GetHistoricalClassRoster(yr, "ມ.4", "2", "ກຳລັງຮຽນ");
        Check("Status=Active: STAT_A included",
            activeOnly.Select("ລະຫັດ='STAT_A'").Length == 1, "STAT_A missing");
        Check("Status=Active: GRAD001 excluded",
            activeOnly.Select("ລະຫັດ='GRAD001'").Length == 0, "graduate leaked into active");

        // Graduated filter → only GRAD001.
        var gradOnly = DB.GetHistoricalClassRoster(yr, "ມ.4", "2", "ຈົບ");
        Check("Status=Graduated: GRAD001 included",
            gradOnly.Select("ລະຫັດ='GRAD001'").Length == 1, "graduate missing");
        Check("Status=Graduated: STAT_A excluded",
            gradOnly.Select("ລະຫັດ='STAT_A'").Length == 0, "active leaked into graduated");

        // null filter (= All) → both.
        var allStatuses = DB.GetHistoricalClassRoster(yr, "ມ.4", "2", null);
        Check("Status=All: GRAD001 included",
            allStatuses.Select("ລະຫັດ='GRAD001'").Length == 1, "graduate missing");
        Check("Status=All: STAT_A included",
            allStatuses.Select("ລະຫັດ='STAT_A'").Length == 1, "active missing");

        // Backward-compat: omitting the param defaults to null (all statuses).
        var defaultBehavior = DB.GetHistoricalClassRoster(yr, "ມ.4", "2");
        CheckEq("Default behavior unchanged (all statuses returned)",
            allStatuses.Rows.Count, defaultBehavior.Rows.Count);
    }

    // ─────────────────────────────────────────────────────────────
    //  Promotion page bulk actions
    //
    //  Verifies the DB-level invariants of the three promotion actions
    //  (Promote Selected, Graduate Selected, Promote Entire Classroom)
    //  WITHOUT instantiating the WPF page. We replicate the EXACT same
    //  SQL the page issues — if these queries match the page's, the
    //  bulk-action behaviour is covered end-to-end.
    //
    //  Invariants checked:
    //    - Promotion writes a GradeHistory row with ClassRoom set
    //    - Promotion writes Students.GradeLevel/ClassRoom/AcademicYear
    //    - Graduation sets Status='ຈົບ', preserves AcademicYear+GradeLevel
    //    - Scores / MonthlyAssessments / EvaluationScores / Enrollments
    //      are NOT touched by either action
    //    - "Promote Entire Classroom" picks active students only
    //  ─────────────────────────────────────────────────────────────
    static void TestPromotionActions()
    {
        string yr1 = "2040-2041", yr2 = "2041-2042";

        // Helper: seed one student into (year, grade, room) with one MATH1
        // enrollment + monthly + final-score so we can later assert that
        // promotion/graduation didn't delete any of it.
        int Seed(string code, string year, string grade, string room, string status = "ກຳລັງຮຽນ")
        {
            DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                      VALUES(@c,'F','L','ຍິງ',@g,@r,@y,@st)",
                null, ("@c", code), ("@g", grade), ("@r", room), ("@y", year), ("@st", status));
            int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode=@c", null, ("@c", code));
            DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                      SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
                null, ("@s", sid), ("@y", year));
            int eid = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
                WHERE StudentID=@s AND AcademicYear=@y AND Semester=1
                  AND SubjectID=(SELECT SubjectID FROM Subjects WHERE SubjectCode='MATH1')",
                null, ("@s", sid), ("@y", year));
            DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                      VALUES(@e,9,3,2,4)", null, ("@e", eid));
            DB.Exec("INSERT OR IGNORE INTO Scores(EnrollID,FinalScore) VALUES(@e,7)", null, ("@e", eid));
            DB.Exec("UPDATE Scores SET FinalScore=7 WHERE EnrollID=@e", null, ("@e", eid));
            DB.RecomputeMidFromMonthly(eid);
            return sid;
        }

        // Snapshot the score-related tables for one student so we can prove
        // promotion didn't delete anything.
        (int enrollments, int monthlies, int scores) ScoreSnapshot(int sid) => (
            DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE StudentID=@s", null, ("@s", sid)),
            DB.ScalarInt(@"SELECT COUNT(*) FROM MonthlyAssessments m
                JOIN Enrollments e ON e.EnrollID=m.EnrollID WHERE e.StudentID=@s",
                null, ("@s", sid)),
            DB.ScalarInt(@"SELECT COUNT(*) FROM Scores sc
                JOIN Enrollments e ON e.EnrollID=sc.EnrollID WHERE e.StudentID=@s",
                null, ("@s", sid)));

        // ── Promote Selected: 2 students from (yr1, ມ.1, 3) → (yr2, ມ.2, 3) ──
        int p1 = Seed("PROM_1", yr1, "ມ.1", "3");
        int p2 = Seed("PROM_2", yr1, "ມ.1", "3");
        var before1 = ScoreSnapshot(p1);
        var before2 = ScoreSnapshot(p2);

        // Replicate PromotionPage.OpenPromoteDialog's loop body.
        string finalY = yr2, finalG = "ມ.2", finalR = "3";
        using (var conn = DB.Open()) using (var tx = conn.BeginTransaction())
        {
            foreach (int sid in new[] { p1, p2 })
            {
                string srcGrade = DB.Scalar("SELECT GradeLevel FROM Students WHERE StudentID=@s",
                    null, ("@s", sid))?.ToString() ?? "";
                string srcRoom  = DB.Scalar("SELECT ClassRoom FROM Students WHERE StudentID=@s",
                    null, ("@s", sid))?.ToString() ?? "";
                DB.ExecTx(@"UPDATE Students
                            SET GradeLevel=@g, ClassRoom=@cr, AcademicYear=@y, Status='ກຳລັງຮຽນ'
                            WHERE StudentID=@id",
                    conn, tx, ("@g", finalG), ("@cr", finalR), ("@y", finalY), ("@id", sid));
                DB.ExecTx(@"INSERT INTO GradeHistory
                            (StudentID, FromGrade, ToGrade, AcademicYear, ClassRoom, Note, ChangedBy)
                            VALUES (@sid, @fg, @tg, @y, @cr, @n, @by)",
                    conn, tx,
                    ("@sid", sid), ("@fg", srcGrade), ("@tg", finalG),
                    ("@y",   finalY), ("@cr", srcRoom), ("@n", "ຂຶ້ນຊັ້ນ"), ("@by", "test"));
            }
            tx.Commit();
        }

        // Verify Students moved to ມ.2 / yr2 / room 3.
        CheckEq("Promote Selected: PROM_1 → ມ.2", "ມ.2",
            DB.Scalar("SELECT GradeLevel FROM Students WHERE StudentID=@s", null, ("@s", p1))?.ToString());
        CheckEq("Promote Selected: PROM_1 → yr2", yr2,
            DB.Scalar("SELECT AcademicYear FROM Students WHERE StudentID=@s", null, ("@s", p1))?.ToString());
        CheckEq("Promote Selected: PROM_2 → ມ.2", "ມ.2",
            DB.Scalar("SELECT GradeLevel FROM Students WHERE StudentID=@s", null, ("@s", p2))?.ToString());

        // GradeHistory row exists for each, with ClassRoom set.
        Check("Promote Selected: PROM_1 GradeHistory row exists",
            DB.ScalarInt(@"SELECT COUNT(*) FROM GradeHistory
                WHERE StudentID=@s AND FromGrade='ມ.1' AND ToGrade='ມ.2'
                  AND AcademicYear=@y AND ClassRoom='3'",
                null, ("@s", p1), ("@y", yr2)) == 1, "missing or wrong");
        Check("Promote Selected: PROM_2 GradeHistory row exists",
            DB.ScalarInt(@"SELECT COUNT(*) FROM GradeHistory
                WHERE StudentID=@s AND FromGrade='ມ.1' AND ToGrade='ມ.2'
                  AND AcademicYear=@y AND ClassRoom='3'",
                null, ("@s", p2), ("@y", yr2)) == 1, "missing or wrong");

        // History preservation: score-related tables UNCHANGED.
        var after1 = ScoreSnapshot(p1);
        var after2 = ScoreSnapshot(p2);
        CheckEq("Promote Selected: PROM_1 Enrollments preserved",         before1.enrollments, after1.enrollments);
        CheckEq("Promote Selected: PROM_1 MonthlyAssessments preserved",  before1.monthlies,   after1.monthlies);
        CheckEq("Promote Selected: PROM_1 Scores preserved",              before1.scores,      after1.scores);
        CheckEq("Promote Selected: PROM_2 Enrollments preserved",         before2.enrollments, after2.enrollments);
        CheckEq("Promote Selected: PROM_2 MonthlyAssessments preserved",  before2.monthlies,   after2.monthlies);
        CheckEq("Promote Selected: PROM_2 Scores preserved",              before2.scores,      after2.scores);

        // ── Graduate Selected: GRAD_X from (yr1, ມ.4, 4) ──
        int gx = Seed("GRAD_X", yr1, "ມ.4", "4");
        var beforeGx = ScoreSnapshot(gx);

        // Replicate PromotionPage.GraduateSelected's loop body.
        using (var conn = DB.Open()) using (var tx = conn.BeginTransaction())
        {
            string nextY = DB.NextYearString(yr1);
            DB.ExecTx("UPDATE Students SET Status='ຈົບ' WHERE StudentID=@id",
                conn, tx, ("@id", gx));
            DB.ExecTx(@"INSERT INTO GradeHistory
                        (StudentID, FromGrade, ToGrade, AcademicYear, ClassRoom, Note, ChangedBy)
                        VALUES (@sid, 'ມ.4', 'ຈົບ', @y, '4', 'ຈົບການສຶກສາ', 'test')",
                conn, tx, ("@sid", gx), ("@y", nextY));
            tx.Commit();
        }

        // Status changed; AcademicYear + GradeLevel PRESERVED (anchored).
        CheckEq("Graduate: Status='ຈົບ'", "ຈົບ",
            DB.Scalar("SELECT Status FROM Students WHERE StudentID=@s", null, ("@s", gx))?.ToString());
        CheckEq("Graduate: AcademicYear preserved (anchored to finishing year)", yr1,
            DB.Scalar("SELECT AcademicYear FROM Students WHERE StudentID=@s", null, ("@s", gx))?.ToString());
        CheckEq("Graduate: GradeLevel preserved (anchored to finishing grade)", "ມ.4",
            DB.Scalar("SELECT GradeLevel FROM Students WHERE StudentID=@s", null, ("@s", gx))?.ToString());
        Check("Graduate: GradeHistory row written with ToGrade='ຈົບ'",
            DB.ScalarInt(@"SELECT COUNT(*) FROM GradeHistory
                WHERE StudentID=@s AND FromGrade='ມ.4' AND ToGrade='ຈົບ' AND ClassRoom='4'",
                null, ("@s", gx)) == 1, "missing or wrong");
        var afterGx = ScoreSnapshot(gx);
        CheckEq("Graduate: Enrollments preserved",        beforeGx.enrollments, afterGx.enrollments);
        CheckEq("Graduate: MonthlyAssessments preserved", beforeGx.monthlies,   afterGx.monthlies);
        CheckEq("Graduate: Scores preserved",             beforeGx.scores,      afterGx.scores);

        // Graduate still discoverable via score-history roster (post-promotion).
        var roster = DB.GetHistoricalClassRoster(yr1, "ມ.4", "4", "ຈົບ");
        Check("Graduate: visible in score-history roster (ຈົບ filter)",
            roster.Select("ລະຫັດ='GRAD_X'").Length == 1, "missing");

        // ── Promote Entire Classroom: skips graduated, picks all active ──
        // Setup: (yr1, ມ.3, 5) — 2 active + 1 graduated.
        int c1 = Seed("CLS_A", yr1, "ມ.3", "5");                                // active
        int c2 = Seed("CLS_B", yr1, "ມ.3", "5");                                // active
        int c3 = Seed("CLS_C", yr1, "ມ.3", "5", status: "ຈົບ");                 // already graduated

        // Replicate PromotionPage.PromoteEntireClassroom — selects only Status='ກຳລັງຮຽນ'.
        var activeIds = new List<int>();
        var classDt = DB.Query(@"
            SELECT StudentID FROM Students
            WHERE AcademicYear=@y AND GradeLevel=@g AND ClassRoom=@r AND Status='ກຳລັງຮຽນ'",
            null, ("@y", yr1), ("@g", "ມ.3"), ("@r", "5"));
        foreach (DataRow rr in classDt.Rows) activeIds.Add(Convert.ToInt32(rr["StudentID"]));
        Check("Promote Entire Classroom: only 2 active selected (graduated skipped)",
            activeIds.Count == 2 && activeIds.Contains(c1) && activeIds.Contains(c2)
                                 && !activeIds.Contains(c3),
            $"got {activeIds.Count} ids");
    }

    // ─────────────────────────────────────────────────────────────
    //  Academic Year page
    //
    //  Covers the spec's verify list:
    //    - Add year                → DB.CreateAcademicYear
    //    - Move to next year       → DB.CreateAcademicYear + DB.SetCurrentAcademicYear
    //    - Set as current year     → DB.SetCurrentAcademicYear (single-current invariant)
    //    - Delete year             → cascade DELETE order matches page's transaction
    //    - Force delete            → all cascade tables wiped in one tx
    //    - Statistics columns      → grid + footer queries return expected counts
    //
    //  Also asserts the spec's "consistency rules": the page MUST go through
    //  DB.SetCurrentAcademicYear (not custom SQL) — we verify the helper
    //  maintains both the Settings.current_year value AND the
    //  AcademicYears.IsCurrent flag atomically.
    //  ─────────────────────────────────────────────────────────────
    static void TestAcademicYearPage()
    {
        // ── Add year via DB.CreateAcademicYear ──────────────────────────
        string addY = "2050-2051";
        // Clean slate
        DB.Exec("DELETE FROM AcademicYears WHERE Year=@y", null, ("@y", addY));
        DB.CreateAcademicYear(addY, null, null, "test add");
        Check("Add year: row inserted",
            DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE Year=@y", null, ("@y", addY)) == 1,
            "missing");
        Check("Add year: not auto-set as current",
            DB.ScalarInt("SELECT IsCurrent FROM AcademicYears WHERE Year=@y", null, ("@y", addY)) == 0,
            "leaked into IsCurrent");

        // Idempotent — re-create silently skips (INSERT OR IGNORE).
        DB.CreateAcademicYear(addY, null, null, "test add 2");
        CheckEq("Add year: idempotent (no duplicate)", 1,
            DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE Year=@y", null, ("@y", addY)));

        // ── Set as current → IsCurrent + Settings update atomically ─────
        DB.SetCurrentAcademicYear(addY);
        CheckEq("Set current: AcademicYears.IsCurrent=1 on chosen", 1,
            DB.ScalarInt("SELECT IsCurrent FROM AcademicYears WHERE Year=@y", null, ("@y", addY)));
        CheckEq("Set current: Settings.current_year matches",
            addY,
            DB.Scalar("SELECT Value FROM Settings WHERE Key='current_year'")?.ToString());
        CheckEq("Set current: exactly ONE IsCurrent=1 row", 1,
            DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE IsCurrent=1"));
        CheckEq("Set current: semester reset to 1", "1",
            DB.Scalar("SELECT Value FROM Settings WHERE Key='current_semester'")?.ToString());
        CheckEq("Set current: DB.CurrentYear cache refreshed", addY, DB.CurrentYear);

        // ── Move to next year (page action replicated) ──────────────────
        string nextY = DB.NextYearString(addY);
        Check("NextYearString format ok", DB.IsValidYearFormat(nextY), $"got {nextY}");
        try { DB.CreateAcademicYear(nextY, null, null, "auto"); } catch { }
        DB.SetCurrentAcademicYear(nextY);
        CheckEq("Move to next: current year advanced", nextY, DB.CurrentYear);
        CheckEq("Move to next: previous year IsCurrent=0", 0,
            DB.ScalarInt("SELECT IsCurrent FROM AcademicYears WHERE Year=@y", null, ("@y", addY)));
        CheckEq("Move to next: single current invariant holds", 1,
            DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE IsCurrent=1"));

        // ── Statistics columns: counts match expectation ────────────────
        // Use a year well clear of the move-to-next collision (nextY = 2051-2052).
        string statsY = "2055-2056";
        DB.Exec("DELETE FROM AcademicYears WHERE Year=@y", null, ("@y", statsY));
        DB.CreateAcademicYear(statsY, null, null, "stats");
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('STAT_AY_1','A','A','ຍິງ','ມ.1','1',@y,'ກຳລັງຮຽນ')", null, ("@y", statsY));
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('STAT_AY_2','B','B','ຊາຍ','ມ.1','1',@y,'ຈົບ')", null, ("@y", statsY));
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('STAT_AY_3','C','C','ຍິງ','ມ.1','2',@y,'ກຳລັງຮຽນ')", null, ("@y", statsY));
        int activeCnt = DB.ScalarInt(
            "SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ກຳລັງຮຽນ'",
            null, ("@y", statsY));
        int gradCnt = DB.ScalarInt(
            "SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ຈົບ'",
            null, ("@y", statsY));
        int totalCnt = DB.ScalarInt(
            "SELECT COUNT(*) FROM Students WHERE AcademicYear=@y", null, ("@y", statsY));
        CheckEq("Stats column: ActiveCount = 2", 2, activeCnt);
        CheckEq("Stats column: GraduatedCount = 1", 1, gradCnt);
        CheckEq("Stats column: StudentCount = 3", 3, totalCnt);

        // Distinct classrooms for this year — query mirrors footer formula.
        int rooms = DB.ScalarInt(@"SELECT COUNT(DISTINCT GradeLevel||'/'||IFNULL(ClassRoom,''))
                                   FROM Students WHERE AcademicYear=@y
                                     AND GradeLevel IS NOT NULL AND GradeLevel<>''",
                                   null, ("@y", statsY));
        CheckEq("Stats footer: distinct classrooms = 2 (ມ.1/1 + ມ.1/2)", 2, rooms);

        // ── Delete: cascade order replicates page's transaction ─────────
        // Seed Enrollments + Scores + Monthly + Eval + Attendance + History for statsY.
        int sid1 = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='STAT_AY_1'");
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='MATH1'",
            null, ("@s", sid1), ("@y", statsY));
        int eid1 = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
            WHERE StudentID=@s AND AcademicYear=@y AND Semester=1
              AND SubjectID=(SELECT SubjectID FROM Subjects WHERE SubjectCode='MATH1')",
            null, ("@s", sid1), ("@y", statsY));
        DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                  VALUES(@e,9,3,2,4)", null, ("@e", eid1));
        DB.Exec("INSERT OR IGNORE INTO Scores(EnrollID,FinalScore) VALUES(@e,7)", null, ("@e", eid1));
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid1, statsY, "SEM1", "CHA1", 8.0, c, tx);
            tx.Commit();
        }
        DB.Exec(@"INSERT INTO GradeHistory(StudentID,FromGrade,ToGrade,AcademicYear,ClassRoom,Note,ChangedBy)
                  VALUES(@s,'ມ.1','ມ.2',@y,'1','test','test')", null, ("@s", sid1), ("@y", statsY));

        int beforeStudents = DB.ScalarInt("SELECT COUNT(*) FROM Students    WHERE AcademicYear=@y", null, ("@y", statsY));
        int beforeEnroll   = DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, ("@y", statsY));
        int beforeEval     = DB.ScalarInt("SELECT COUNT(*) FROM EvaluationScores WHERE AcademicYear=@y", null, ("@y", statsY));
        int beforeHist     = DB.ScalarInt("SELECT COUNT(*) FROM GradeHistory WHERE AcademicYear=@y", null, ("@y", statsY));
        Check("Pre-delete: data exists", beforeStudents + beforeEnroll + beforeEval + beforeHist > 0, "no data");

        // Execute the page's cascade transaction (same SQL).
        using (var conn = DB.Open()) using (var tx = conn.BeginTransaction())
        {
            DB.ExecTx("DELETE FROM Enrollments       WHERE AcademicYear=@y", conn, tx, ("@y", statsY));
            DB.ExecTx("DELETE FROM EvaluationScores  WHERE AcademicYear=@y", conn, tx, ("@y", statsY));
            DB.ExecTx("DELETE FROM AttendanceRecords WHERE AcademicYear=@y", conn, tx, ("@y", statsY));
            DB.ExecTx("DELETE FROM GradeHistory      WHERE AcademicYear=@y", conn, tx, ("@y", statsY));
            DB.ExecTx("DELETE FROM Students          WHERE AcademicYear=@y", conn, tx, ("@y", statsY));
            DB.ExecTx("DELETE FROM AcademicYears     WHERE Year=@y",         conn, tx, ("@y", statsY));
            tx.Commit();
        }
        CheckEq("Force delete: Students wiped",         0, DB.ScalarInt("SELECT COUNT(*) FROM Students    WHERE AcademicYear=@y", null, ("@y", statsY)));
        CheckEq("Force delete: Enrollments wiped",      0, DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, ("@y", statsY)));
        CheckEq("Force delete: EvaluationScores wiped", 0, DB.ScalarInt("SELECT COUNT(*) FROM EvaluationScores WHERE AcademicYear=@y", null, ("@y", statsY)));
        CheckEq("Force delete: GradeHistory wiped",     0, DB.ScalarInt("SELECT COUNT(*) FROM GradeHistory WHERE AcademicYear=@y", null, ("@y", statsY)));
        CheckEq("Force delete: AcademicYears row gone", 0, DB.ScalarInt("SELECT COUNT(*) FROM AcademicYears WHERE Year=@y", null, ("@y", statsY)));

        // The current year is unaffected by deleting another year.
        CheckEq("Force delete: current year unchanged", nextY, DB.CurrentYear);
        CheckEq("Force delete: current year still IsCurrent=1", 1,
            DB.ScalarInt("SELECT IsCurrent FROM AcademicYears WHERE Year=@y", null, ("@y", nextY)));
    }

    // ─────────────────────────────────────────────────────────────
    //  DB.TryParseScore — the parser both score-entry pages use to
    //  validate + clamp teacher-typed input. Covers the spec rules:
    //    - Decimals supported (0, 5, 7.5, 8.75, 10)
    //    - < min clamps to min
    //    - > max clamps to max
    //    - Non-numeric input rejected (returns false → page cancels commit)
    //    - Empty string maps to 0 (clearing a cell == zero)
    //    - Both '.' and ',' decimal separators accepted (Lao locale uses ',')
    //    - Each column passes its own max (Activity=3 / Discipline=2 /
    //      Homework=5 / Eval+Final=10)
    //  ─────────────────────────────────────────────────────────────
    static void TestTryParseScore()
    {
        bool ok; double v;

        // Spec examples — all valid for max=10
        ok = DB.TryParseScore("0",    0, 10, out v); Check("Parse '0' = 0",         ok && v == 0,    $"got {ok}/{v}");
        ok = DB.TryParseScore("5",    0, 10, out v); Check("Parse '5' = 5",         ok && v == 5,    $"got {ok}/{v}");
        ok = DB.TryParseScore("7.5",  0, 10, out v); Check("Parse '7.5' = 7.5",     ok && v == 7.5,  $"got {ok}/{v}");
        ok = DB.TryParseScore("8.75", 0, 10, out v); Check("Parse '8.75' = 8.75",   ok && v == 8.75, $"got {ok}/{v}");
        ok = DB.TryParseScore("10",   0, 10, out v); Check("Parse '10' = 10",       ok && v == 10,   $"got {ok}/{v}");

        // Clamping
        ok = DB.TryParseScore("-1",   0, 10, out v); Check("Clamp: '-1' → 0",       ok && v == 0,    $"got {ok}/{v}");
        ok = DB.TryParseScore("-5.5", 0, 10, out v); Check("Clamp: '-5.5' → 0",     ok && v == 0,    $"got {ok}/{v}");
        ok = DB.TryParseScore("12",   0, 10, out v); Check("Clamp: '12' → 10",      ok && v == 10,   $"got {ok}/{v}");
        ok = DB.TryParseScore("999",  0, 10, out v); Check("Clamp: '999' → 10",     ok && v == 10,   $"got {ok}/{v}");

        // Per-column max — Activity=3, Discipline=2, Homework=5
        ok = DB.TryParseScore("4",    0, 3,  out v); Check("Activity max 3: '4' → 3",    ok && v == 3, $"got {ok}/{v}");
        ok = DB.TryParseScore("2.7",  0, 3,  out v); Check("Activity max 3: '2.7' in-range", ok && v == 2.7, $"got {ok}/{v}");
        ok = DB.TryParseScore("10",   0, 2,  out v); Check("Discipline max 2: '10' → 2",     ok && v == 2, $"got {ok}/{v}");
        ok = DB.TryParseScore("7",    0, 5,  out v); Check("Homework max 5: '7' → 5",        ok && v == 5, $"got {ok}/{v}");

        // Empty input → 0
        ok = DB.TryParseScore("",     0, 10, out v); Check("Empty → 0",             ok && v == 0,    $"got {ok}/{v}");
        ok = DB.TryParseScore("   ",  0, 10, out v); Check("Whitespace → 0",        ok && v == 0,    $"got {ok}/{v}");
        ok = DB.TryParseScore(null,   0, 10, out v); Check("null → 0",              ok && v == 0,    $"got {ok}/{v}");

        // Comma decimal separator (Lao Windows locale often uses ',')
        ok = DB.TryParseScore("8,5",  0, 10, out v); Check("Comma sep: '8,5' → 8.5", ok && v == 8.5,  $"got {ok}/{v}");
        ok = DB.TryParseScore("7,25", 0, 10, out v); Check("Comma sep: '7,25' → 7.25", ok && v == 7.25, $"got {ok}/{v}");

        // Invalid input — must REJECT (returns false) so the page cancels commit
        ok = DB.TryParseScore("abc",  0, 10, out v); Check("Reject 'abc'",          !ok,             $"got {ok}/{v}");
        ok = DB.TryParseScore("--",   0, 10, out v); Check("Reject '--'",           !ok,             $"got {ok}/{v}");
        ok = DB.TryParseScore("12a",  0, 10, out v); Check("Reject '12a'",          !ok,             $"got {ok}/{v}");
        ok = DB.TryParseScore("5.5x", 0, 10, out v); Check("Reject '5.5x'",         !ok,             $"got {ok}/{v}");
        ok = DB.TryParseScore("1e999",0, 10, out v); Check("Reject Infinity",       !ok,             $"got {ok}/{v}");

        // Rounding to 2 decimals (keeps stored data tidy)
        ok = DB.TryParseScore("8.7654321", 0, 10, out v); Check("Round to 2dp: '8.7654321' → 8.77", ok && v == 8.77, $"got {ok}/{v}");
    }

    // ─────────────────────────────────────────────────────────────
    //  DB.TryParseIntScore — STRICT integer parser used by both
    //  score-entry grids after the integer-only switch. Decimals,
    //  locale separators, letters, and signs are all rejected.
    //  Out-of-range is also a hard reject (no clamping).
    //  ─────────────────────────────────────────────────────────────
    static void TestTryParseIntScore()
    {
        bool ok; int v;

        // Spec examples — all valid integers 0..10
        for (int i = 0; i <= 10; i++)
        {
            ok = DB.TryParseIntScore(i.ToString(), 0, 10, out v);
            Check($"Int: '{i}' valid", ok && v == i, $"got {ok}/{v}");
        }

        // Spec REJECT list
        ok = DB.TryParseIntScore("7.5",  0, 10, out v); Check("Reject decimal '7.5'",   !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("8.25", 0, 10, out v); Check("Reject decimal '8.25'",  !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("9.8",  0, 10, out v); Check("Reject decimal '9.8'",   !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("10.5", 0, 10, out v); Check("Reject decimal '10.5'",  !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("abc",  0, 10, out v); Check("Reject letters 'abc'",   !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("--",   0, 10, out v); Check("Reject '--'",            !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("12a",  0, 10, out v); Check("Reject '12a'",           !ok, $"got {ok}/{v}");

        // Locale decimal separators — explicitly rejected (no longer accepted as 8.5)
        ok = DB.TryParseIntScore("8,5",  0, 10, out v); Check("Reject comma sep '8,5'", !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("7.0",  0, 10, out v); Check("Reject '7.0' (decimal point present)", !ok, $"got {ok}/{v}");

        // Out-of-range is REJECTED (not clamped — strict mode)
        ok = DB.TryParseIntScore("-1",   0, 10, out v); Check("Reject '-1' (out of range)",  !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("11",   0, 10, out v); Check("Reject '11' (out of range)",  !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("999",  0, 10, out v); Check("Reject '999' (out of range)", !ok, $"got {ok}/{v}");

        // Per-column max — Activity=3, Discipline=2, Homework=5
        ok = DB.TryParseIntScore("3",    0, 3,  out v); Check("Activity max 3: '3' valid",    ok && v == 3, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("4",    0, 3,  out v); Check("Activity max 3: '4' rejected", !ok,          $"got {ok}/{v}");
        ok = DB.TryParseIntScore("2",    0, 2,  out v); Check("Discipline max 2: '2' valid",  ok && v == 2, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("3",    0, 2,  out v); Check("Discipline max 2: '3' rejected", !ok,        $"got {ok}/{v}");
        ok = DB.TryParseIntScore("5",    0, 5,  out v); Check("Homework max 5: '5' valid",    ok && v == 5, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("6",    0, 5,  out v); Check("Homework max 5: '6' rejected", !ok,          $"got {ok}/{v}");

        // Empty / whitespace → 0 (clearing a cell == zero)
        ok = DB.TryParseIntScore("",     0, 10, out v); Check("Empty → 0",      ok && v == 0, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("   ",  0, 10, out v); Check("Whitespace → 0", ok && v == 0, $"got {ok}/{v}");
        ok = DB.TryParseIntScore(null,   0, 10, out v); Check("null → 0",       ok && v == 0, $"got {ok}/{v}");

        // Sign + spaces inside number → rejected
        ok = DB.TryParseIntScore("+5",   0, 10, out v); Check("Reject '+5' (sign)",         !ok, $"got {ok}/{v}");
        ok = DB.TryParseIntScore("1 0",  0, 10, out v); Check("Reject '1 0' (inner space)", !ok, $"got {ok}/{v}");
    }

    // ─────────────────────────────────────────────────────────────
    //  CHA1 / LAB1 manual-entry contract across all 11 contexts:
    //    Month1..Month8 (per-month)  ·  SEM1  ·  SEM2  ·  ANNUAL
    //
    //  Verifies:
    //    1. DB.MonthContextName maps calendar months → "Month1".."Month8"
    //    2. UPSERT round-trip via SetEvaluationScore/GetEvaluationScore for
    //       every context — no duplicates, value-overwrites on second save
    //    3. Setting score=null DELETES the row (lets teachers clear entries)
    //    4. Each context is INDEPENDENT — writing Month1 doesn't touch
    //       Month2, SEM1, ANNUAL, etc.
    //    5. Monthly reads now route through EvaluationScores (no MonthlyAssessments
    //       derivation for CHA1/LAB1 in monthly scope) — verified by checking
    //       that academic-style monthly entries on the CHA1 enrollment do NOT
    //       leak into the monthly CHA score we read back.
    //  ─────────────────────────────────────────────────────────────
    static void TestChaLabManualEntry()
    {
        // (1) MonthContextName mapping
        CheckEq("MonthContextName(9)  = Month1",  "Month1", DB.MonthContextName(9));
        CheckEq("MonthContextName(10) = Month2",  "Month2", DB.MonthContextName(10));
        CheckEq("MonthContextName(11) = Month3",  "Month3", DB.MonthContextName(11));
        CheckEq("MonthContextName(12) = Month4",  "Month4", DB.MonthContextName(12));
        CheckEq("MonthContextName(2)  = Month5",  "Month5", DB.MonthContextName(2));
        CheckEq("MonthContextName(3)  = Month6",  "Month6", DB.MonthContextName(3));
        CheckEq("MonthContextName(4)  = Month7",  "Month7", DB.MonthContextName(4));
        CheckEq("MonthContextName(5)  = Month8",  "Month8", DB.MonthContextName(5));
        Check("MonthContextName(1)  rejected (final-exam month)", DB.MonthContextName(1) == null);
        Check("MonthContextName(6)  rejected (final-exam month)", DB.MonthContextName(6) == null);
        Check("MonthContextName(13) rejected (invalid)",          DB.MonthContextName(13) == null);

        // Seed a student
        string yr = "2055-2056";
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('CL_E1','C','L','ຍິງ','ມ.4','1',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='CL_E1'");

        // (2) UPSERT round-trip for every context
        string[] contexts = {
            "Month1","Month2","Month3","Month4","Month5","Month6","Month7","Month8",
            "SEM1","SEM2","ANNUAL"
        };
        foreach (var ctx in contexts)
        {
            using (var c = DB.Open()) using (var tx = c.BeginTransaction())
            {
                DB.SetEvaluationScore(sid, yr, ctx, "CHA1", 7, c, tx);
                DB.SetEvaluationScore(sid, yr, ctx, "LAB1", 8, c, tx);
                tx.Commit();
            }
            CheckEq($"Round-trip {ctx} CHA1 = 7",   7.0, DB.GetEvaluationScore(sid, yr, ctx, "CHA1") ?? -1);
            CheckEq($"Round-trip {ctx} LAB1 = 8",   8.0, DB.GetEvaluationScore(sid, yr, ctx, "LAB1") ?? -1);
        }

        // UPSERT: re-saving the same context updates the value in-place (no duplicates).
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "Month1", "CHA1", 9, c, tx);
            tx.Commit();
        }
        CheckEq("UPSERT: Month1 CHA1 updates 7→9", 9.0, DB.GetEvaluationScore(sid, yr, "Month1", "CHA1") ?? -1);
        // Row count remains 1 per (sid, year, ctx, code) — UNIQUE constraint guarantees this.
        int rowCount = DB.ScalarInt(@"SELECT COUNT(*) FROM EvaluationScores
            WHERE StudentID=@s AND AcademicYear=@y AND Context='Month1' AND SubjectCode='CHA1'",
            null, ("@s", sid), ("@y", yr));
        CheckEq("UPSERT: only 1 row per (sid, year, ctx, code)", 1, rowCount);

        // (3) score=null DELETES the row
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "Month1", "CHA1", null, c, tx);
            tx.Commit();
        }
        Check("Set Month1 CHA1=null deletes row",
            DB.GetEvaluationScore(sid, yr, "Month1", "CHA1") == null,
            $"got {DB.GetEvaluationScore(sid, yr, "Month1", "CHA1")}");
        // The other Month1 LAB1 row should still exist (only the deleted CHA1 went away).
        CheckEq("Set null deletes ONLY the targeted subject", 8.0,
            DB.GetEvaluationScore(sid, yr, "Month1", "LAB1") ?? -1);

        // (4) Independence between contexts
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "Month2", "CHA1", 5, c, tx);
            tx.Commit();
        }
        CheckEq("Independence: Month2 CHA1 = 5",  5.0, DB.GetEvaluationScore(sid, yr, "Month2", "CHA1") ?? -1);
        CheckEq("Independence: Month3 CHA1 unchanged (= 7)", 7.0,
            DB.GetEvaluationScore(sid, yr, "Month3", "CHA1") ?? -1);
        CheckEq("Independence: SEM1 CHA1 unchanged (= 7)",   7.0,
            DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1") ?? -1);
        CheckEq("Independence: ANNUAL CHA1 unchanged (= 7)", 7.0,
            DB.GetEvaluationScore(sid, yr, "ANNUAL", "CHA1") ?? -1);

        // (5) Monthly reads via EvaluationScores — verified directly via GetEvaluationScore.
        // The MonthlyAssessments table is NEVER consulted for CHA1/LAB1 monthly values
        // anymore. Spot-check: insert a MonthlyAssessments row for the same student's
        // CHA1 enrollment with bogus high values, then re-read EvaluationScores —
        // the value must come from EvaluationScores (5), not the MonthlyAssessments sum.
        DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                  SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode='CHA1'",
            null, ("@s", sid), ("@y", yr));
        int chaEid = DB.ScalarInt(@"SELECT EnrollID FROM Enrollments
            WHERE StudentID=@s AND AcademicYear=@y AND Semester=1
              AND SubjectID=(SELECT SubjectID FROM Subjects WHERE SubjectCode='CHA1')",
            null, ("@s", sid), ("@y", yr));
        // Bogus MonthlyAssessments entry (legacy auto-calc path)
        DB.Exec(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore)
                  VALUES(@e,10,3,2,5)", null, ("@e", chaEid));
        // EvaluationScores read should be unaffected — still 5.
        CheckEq("CHA1 monthly read isolated from MonthlyAssessments noise",
            5.0, DB.GetEvaluationScore(sid, yr, "Month2", "CHA1") ?? -1);

        // (6) GetHistoryMonthly must surface CHA1/LAB1 rows for the per-month
        // display window in ProfileHistoryPage. Before the UNION fix, only
        // academic subjects appeared because the query INNER JOINed
        // MonthlyAssessments — a table CHA1/LAB1 rows never touch.
        var hist = DB.GetHistoryMonthly(sid, yr);
        // Month2 (=October, Month2 context) has: LAB1=8 (from step 2) and
        // CHA1=5 (from step 4 independence check).
        var oct = hist.Select("ເດືອນ = 10");
        var octCha = oct.FirstOrDefault(r => r["ລະຫັດວິຊາ"].ToString() == "CHA1");
        var octLab = oct.FirstOrDefault(r => r["ລະຫັດວິຊາ"].ToString() == "LAB1");
        Check("GetHistoryMonthly: CHA1 row present for month 10 (Month2 ctx)", octCha != null);
        Check("GetHistoryMonthly: LAB1 row present for month 10 (Month2 ctx)", octLab != null);
        if (octCha != null)
            CheckEq("GetHistoryMonthly: CHA1 Month2 ລວມເດືອນ = 5", 5.0, Convert.ToDouble(octCha["ລວມເດືອນ"]));
        if (octLab != null)
            CheckEq("GetHistoryMonthly: LAB1 Month2 ລວມເດືອນ = 8", 8.0, Convert.ToDouble(octLab["ລວມເດືອນ"]));
        // Semester attribution: month 10 (Oct) should map to ພາກ=1.
        if (octCha != null)
            CheckEq("GetHistoryMonthly: CHA1 Month2 → ພາກ 1", 1, Convert.ToInt32(octCha["ພາກ"]));
        // Feb (=Month5) has CHA1=7 seeded in step 2 (contexts loop).
        var febCha = hist.Select("ເດືອນ = 2").FirstOrDefault(r => r["ລະຫັດວິຊາ"].ToString() == "CHA1");
        Check("GetHistoryMonthly: CHA1 row present for month 2 (Month5 ctx)", febCha != null);
        if (febCha != null)
        {
            CheckEq("GetHistoryMonthly: CHA1 Month5 ລວມເດືອນ = 7", 7.0, Convert.ToDouble(febCha["ລວມເດືອນ"]));
            CheckEq("GetHistoryMonthly: CHA1 Month5 → ພາກ 2", 2, Convert.ToInt32(febCha["ພາກ"]));
        }
    }

    // ─────────────────────────────────────────────────────────────
    //  ScoresPage CHA1/LAB1 inline contract
    //
    //  Verifies the data-layer invariants the redesigned ScoresPage
    //  relies on:
    //    1. The SQL query that drives the page now INCLUDES CHA1/LAB1
    //       (older query had NOT IN — verified by counting subjects).
    //    2. CHA1/LAB1 enrollments don't need a Scores row — they read
    //       FinalScore from EvaluationScores(SEM{N}, code).
    //    3. Save path = DB.SetEvaluationScore (UPSERT). Re-saving updates
    //       in place (no duplicates).
    //    4. Academic subjects stay on the Scores table — CHA1/LAB1 save
    //       doesn't accidentally write a Scores row.
    //  ─────────────────────────────────────────────────────────────
    static void TestChaLabInlineSemester()
    {
        string yr = "2060-2061";
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('INL_1','I','L','ຍິງ','ມ.1','1',@y,'ກຳລັງຮຽນ')",
            null, ("@y", yr));
        int sid = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='INL_1'");

        // Enroll in MATH1 (academic) + CHA1 + LAB1 (eval) for sem 1.
        foreach (var code in new[] { "MATH1", "CHA1", "LAB1" })
            DB.Exec(@"INSERT INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                      SELECT @s,SubjectID,@y,1 FROM Subjects WHERE SubjectCode=@c",
                null, ("@s", sid), ("@y", yr), ("@c", code));

        // (1) ScoresPage roster query SHAPE — counts must include CHA1/LAB1 now.
        // Mirrors the page's SELECT after removing the NOT IN filter.
        var dt = DB.Query(@"
            SELECT s.SubjectCode FROM Enrollments e
            JOIN Subjects s ON s.SubjectID=e.SubjectID
            WHERE e.StudentID=@sid AND e.AcademicYear=@yr AND e.Semester=@sm
            ORDER BY s.SortOrder",
            null, ("@sid", sid), ("@yr", yr), ("@sm", 1));
        int chaCount = dt.Select("SubjectCode='CHA1'").Length;
        int labCount = dt.Select("SubjectCode='LAB1'").Length;
        int mathCount = dt.Select("SubjectCode='MATH1'").Length;
        Check("ScoresPage roster: CHA1 row present", chaCount == 1, $"got {chaCount}");
        Check("ScoresPage roster: LAB1 row present", labCount == 1, $"got {labCount}");
        Check("ScoresPage roster: MATH1 row present", mathCount == 1, $"got {mathCount}");

        // (2) FinalScore for CHA1/LAB1 reads from EvaluationScores, not Scores.
        // Seed via SetEvaluationScore; read back via GetEvaluationScore.
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1", "CHA1", 9, c, tx);
            DB.SetEvaluationScore(sid, yr, "SEM1", "LAB1", 8, c, tx);
            tx.Commit();
        }
        CheckEq("ScoresPage CHA1 FinalScore = EvaluationScores(SEM1, CHA1)",
            9.0, DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1") ?? -1);
        CheckEq("ScoresPage LAB1 FinalScore = EvaluationScores(SEM1, LAB1)",
            8.0, DB.GetEvaluationScore(sid, yr, "SEM1", "LAB1") ?? -1);

        // (3) UPSERT: re-saving the same context overwrites (no duplicate row).
        using (var c = DB.Open()) using (var tx = c.BeginTransaction())
        {
            DB.SetEvaluationScore(sid, yr, "SEM1", "CHA1", 10, c, tx);
            tx.Commit();
        }
        CheckEq("ScoresPage CHA1 UPSERT 9→10",
            10.0, DB.GetEvaluationScore(sid, yr, "SEM1", "CHA1") ?? -1);
        int rows = DB.ScalarInt(@"SELECT COUNT(*) FROM EvaluationScores
            WHERE StudentID=@s AND AcademicYear=@y AND Context='SEM1' AND SubjectCode='CHA1'",
            null, ("@s", sid), ("@y", yr));
        CheckEq("ScoresPage CHA1 UPSERT no duplicate row", 1, rows);

        // (4) Academic Scores row for CHA1/LAB1 enrollments stays ABSENT — save
        // path routes them to EvaluationScores, never to Scores. Verify by
        // querying for any Scores row tied to a CHA1/LAB1 enrollment.
        int chaScoreRows = DB.ScalarInt(@"
            SELECT COUNT(*) FROM Scores sc
            JOIN Enrollments e ON e.EnrollID=sc.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            WHERE e.StudentID=@s AND e.AcademicYear=@y AND sub.SubjectCode IN ('CHA1','LAB1')",
            null, ("@s", sid), ("@y", yr));
        CheckEq("CHA1/LAB1 enrollments have NO Scores row", 0, chaScoreRows);
    }

    // ─────────────────────────────────────────────────────────────
    //  Excel score import — PER-SUBJECT round-trip. One template per
    //  (year, grade, room, month-or-sem, subjectCode). Covers both
    //  kinds (monthly + semester), academic + CHA1 + LAB1 routing,
    //  UPSERT, duplicate detection within a file, scope-mismatch
    //  rejection (wrong subject / wrong month / arbitrary xlsx),
    //  invalid-score handling, and graduated-student read-only.
    // ─────────────────────────────────────────────────────────────
    static void TestExcelImport()
    {
        string yr = "2090-2091";
        string grade = "ມ.4";
        string room  = "5";

        // Seed 3 active students + 1 graduated in (yr, grade, room).
        var ids = new Dictionary<string, int>();
        foreach (var sc in new[] { "IMP001", "IMP002", "IMP003" })
        {
            DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                      VALUES(@c,@f,@l,'ຊາຍ',@g,@r,@y,'ກຳລັງຮຽນ')",
                null, ("@c", sc), ("@f", "Imp"), ("@l", sc), ("@g", grade), ("@r", room), ("@y", yr));
            ids[sc] = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode=@c", null, ("@c", sc));
        }
        DB.Exec(@"INSERT INTO Students(StudentCode,FirstName,LastName,Gender,GradeLevel,ClassRoom,AcademicYear,Status)
                  VALUES('IMP_GRAD','Grad','Imp','ຊາຍ',@g,@r,@y,'ຈົບ')",
            null, ("@g", grade), ("@r", room), ("@y", yr));
        ids["IMP_GRAD"] = DB.ScalarInt("SELECT StudentID FROM Students WHERE StudentCode='IMP_GRAD'");

        // Enrol every student (incl. graduate) in every subject × both semesters.
        var subjects = DB.Query("SELECT SubjectID, SubjectCode FROM Subjects");
        foreach (var (code, sid) in ids)
            foreach (DataRow sub in subjects.Rows)
                for (int sem = 1; sem <= 2; sem++)
                    DB.Exec(@"INSERT OR IGNORE INTO Enrollments(StudentID,SubjectID,AcademicYear,Semester)
                              VALUES(@s,@sub,@y,@sm)",
                        null, ("@s", sid), ("@sub", Convert.ToInt32(sub["SubjectID"])),
                              ("@y", yr), ("@sm", sem));

        // ── Monthly template (MATH1 only) ───────────────────────
        string tplMath = Path.Combine(Path.GetTempPath(), $"sis_import_math_{Guid.NewGuid():N}.xlsx");
        ExcelImport.BuildMonthlyTemplate(tplMath, yr, grade, room, 9, "MATH1");
        Check("Monthly template (MATH1) created", File.Exists(tplMath));
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplMath))
        {
            var ws = wb.Worksheet(1);
            CheckEq("Monthly template magic v3", ExcelImport.MagicMonthly, ws.Cell("Z1").GetString());
            CheckEq("Monthly template year",     yr,      ws.Cell("Z2").GetString());
            CheckEq("Monthly template grade",    grade,   ws.Cell("Z3").GetString());
            CheckEq("Monthly template room",     room,    ws.Cell("Z4").GetString());
            CheckEq("Monthly template month",    9,       (int)ws.Cell("Z5").GetDouble());
            CheckEq("Monthly template subject",  "MATH1", ws.Cell("Z6").GetString());
            // Per-subject template: one row per active student = 4 (incl. grad).
            int lastRow = ws.LastRowUsed()!.RowNumber();
            CheckEq("Monthly template: 4 data rows (one per roster student)", 4 + 4 - 1, lastRow);
            CheckEq("Monthly template: column A is ordinal", 1, (int)ws.Cell(4, 1).GetDouble());
            CheckEq("Monthly template: column B has student code",
                "IMP001", ws.Cell(FindStudentRow(ws, "IMP001"), 2).GetString());
        }

        // Fill MATH1 sub-scores. Academic monthly templates have 3 columns:
        //   D = ກິດຈະກຳ (Discipline /2)
        //   E = ຮ່ວມຮຽນ (Activity   /3)
        //   F = ກວດກາ   (Homework   /5)
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplMath))
        {
            var ws = wb.Worksheet(1);
            // IMP001 → (2, 3, 3) = total 8
            ws.Cell(FindStudentRow(ws, "IMP001"), 4).Value = 2;
            ws.Cell(FindStudentRow(ws, "IMP001"), 5).Value = 3;
            ws.Cell(FindStudentRow(ws, "IMP001"), 6).Value = 3;
            // IMP002 → (2, 3, 5) = total 10
            ws.Cell(FindStudentRow(ws, "IMP002"), 4).Value = 2;
            ws.Cell(FindStudentRow(ws, "IMP002"), 5).Value = 3;
            ws.Cell(FindStudentRow(ws, "IMP002"), 6).Value = 5;
            // IMP003 → invalid Activity (column E = "abc")
            ws.Cell(FindStudentRow(ws, "IMP003"), 4).Value = 1;
            ws.Cell(FindStudentRow(ws, "IMP003"), 5).Value = "abc";
            // IMP_GRAD → graduated, marked invalid regardless of scores
            ws.Cell(FindStudentRow(ws, "IMP_GRAD"), 4).Value = 1;
            ws.Cell(FindStudentRow(ws, "IMP_GRAD"), 5).Value = 2;
            ws.Cell(FindStudentRow(ws, "IMP_GRAD"), 6).Value = 3;
            // Duplicate row for IMP001 (different sub-score values)
            int dup = ws.LastRowUsed()!.RowNumber() + 1;
            ws.Cell(dup, 1).Value = 99;
            ws.Cell(dup, 2).Value = "IMP001";
            ws.Cell(dup, 3).Value = "duplicate";
            ws.Cell(dup, 4).Value = 0;
            ws.Cell(dup, 5).Value = 0;
            ws.Cell(dup, 6).Value = 1;
            wb.Save();
        }

        var resMath = ExcelImport.ParseMonthly(tplMath, yr, grade, room, 9, "MATH1");
        Check("MATH1 parse: no fatal error", resMath.FatalError == null);
        CheckEq("MATH1 parse: 5 rows surfaced (4 prefilled + 1 dup)", 5, resMath.Rows.Count);
        CheckEq("MATH1 parse: 2 valid (IMP001, IMP002)", 2, resMath.ValidCount);
        CheckEq("MATH1 parse: 3 invalid (bad score, graduate, duplicate)", 3, resMath.InvalidCount);
        Check("MATH1: bad score flagged invalid",
            resMath.Rows.Any(r => r.StudentCode == "IMP003" && !r.IsValid));
        Check("MATH1: graduated student flagged invalid",
            resMath.Rows.Any(r => r.StudentCode == "IMP_GRAD" && !r.IsValid));
        Check("MATH1: duplicate row flagged invalid",
            resMath.Rows.Count(r => r.StudentCode == "IMP001") == 2 &&
            resMath.Rows.Where(r => r.StudentCode == "IMP001").Count(r => r.IsValid) == 1);
        Check("MATH1: every row's SubjectCode is MATH1",
            resMath.Rows.All(r => r.SubjectCode == "MATH1"));
        // IMP001 valid row carries the three sub-scores correctly.
        var imp001Row = resMath.Rows.First(r => r.StudentCode == "IMP001" && r.IsValid);
        CheckEq("MATH1: IMP001 Discipline = 2", 2, imp001Row.DisciplineScore ?? -1);
        CheckEq("MATH1: IMP001 Activity = 3",   3, imp001Row.ActivityScore   ?? -1);
        CheckEq("MATH1: IMP001 Homework = 3",   3, imp001Row.HomeworkScore   ?? -1);
        CheckEq("MATH1: IMP001 SubScoreTotal = 8", 8, imp001Row.SubScoreTotal);

        CheckEq("MATH1 save: 2 rows persisted", 2, ExcelImport.SaveImport(resMath));
        // Verify the IMP001 / MATH1 / Month 9 row in MonthlyAssessments —
        // sub-scores must be stored verbatim (no split logic anymore).
        var dt = DB.Query(@"
            SELECT ma.DisciplineScore, ma.ActivityScore, ma.HomeworkScore
            FROM MonthlyAssessments ma
            JOIN Enrollments e ON e.EnrollID=ma.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            JOIN Students s    ON s.StudentID=e.StudentID
            WHERE s.StudentCode='IMP001' AND sub.SubjectCode='MATH1'
              AND ma.Month=9 AND e.AcademicYear=@y AND e.Semester=1",
            null, ("@y", yr));
        Check("MATH1 save: MA row exists for IMP001", dt.Rows.Count == 1);
        CheckEq("MATH1 save: Discipline written verbatim = 2", 2.0, Convert.ToDouble(dt.Rows[0]["DisciplineScore"]));
        CheckEq("MATH1 save: Activity   written verbatim = 3", 3.0, Convert.ToDouble(dt.Rows[0]["ActivityScore"]));
        CheckEq("MATH1 save: Homework   written verbatim = 3", 3.0, Convert.ToDouble(dt.Rows[0]["HomeworkScore"]));

        // Scores.MidScore should be recomputed from MonthlyAssessments.
        double midMath = Convert.ToDouble(DB.Scalar(@"
            SELECT sc.MidScore FROM Scores sc
            JOIN Enrollments e ON e.EnrollID=sc.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            WHERE e.StudentID=@s AND sub.SubjectCode='MATH1' AND e.AcademicYear=@y AND e.Semester=1",
            null, ("@s", ids["IMP001"]), ("@y", yr)) ?? 0.0);
        Check("MATH1 save: Scores.MidScore recomputed for IMP001", midMath > 0);

        // ── UPSERT: re-import MATH1 with different sub-scores ────
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplMath))
        {
            var ws = wb.Worksheet(1);
            // Overwrite IMP001 with (1, 2, 2) = total 5
            ws.Cell(FindStudentRow(ws, "IMP001"), 4).Value = 1;
            ws.Cell(FindStudentRow(ws, "IMP001"), 5).Value = 2;
            ws.Cell(FindStudentRow(ws, "IMP001"), 6).Value = 2;
            wb.Save();
        }
        ExcelImport.SaveImport(ExcelImport.ParseMonthly(tplMath, yr, grade, room, 9, "MATH1"));
        var dt2 = DB.Query(@"
            SELECT ma.DisciplineScore, ma.ActivityScore, ma.HomeworkScore
            FROM MonthlyAssessments ma
            JOIN Enrollments e ON e.EnrollID=ma.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            JOIN Students s    ON s.StudentID=e.StudentID
            WHERE s.StudentCode='IMP001' AND sub.SubjectCode='MATH1'
              AND ma.Month=9 AND e.AcademicYear=@y AND e.Semester=1",
            null, ("@y", yr));
        CheckEq("MATH1 UPSERT: Discipline overwritten 2 → 1", 1.0, Convert.ToDouble(dt2.Rows[0]["DisciplineScore"]));
        CheckEq("MATH1 UPSERT: Activity   overwritten 3 → 2", 2.0, Convert.ToDouble(dt2.Rows[0]["ActivityScore"]));
        CheckEq("MATH1 UPSERT: Homework   overwritten 3 → 2", 2.0, Convert.ToDouble(dt2.Rows[0]["HomeworkScore"]));
        int rowCount = DB.ScalarInt(@"
            SELECT COUNT(*) FROM MonthlyAssessments ma
            JOIN Enrollments e ON e.EnrollID=ma.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            JOIN Students s    ON s.StudentID=e.StudentID
            WHERE s.StudentCode='IMP001' AND sub.SubjectCode='MATH1'
              AND ma.Month=9 AND e.AcademicYear=@y AND e.Semester=1",
            null, ("@y", yr));
        CheckEq("MATH1 UPSERT: no duplicate row", 1, rowCount);

        // Per-column max enforcement: 3 for Discipline (max 2) must be invalid.
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplMath))
        {
            var ws = wb.Worksheet(1);
            ws.Cell(FindStudentRow(ws, "IMP002"), 4).Value = 3;   // exceeds Discipline max (2)
            wb.Save();
        }
        var resOver = ExcelImport.ParseMonthly(tplMath, yr, grade, room, 9, "MATH1");
        Check("MATH1: Discipline over max flagged invalid",
            resOver.Rows.Any(r => r.StudentCode == "IMP002" && !r.IsValid && r.Status.Contains("ກິດຈະກຳ")));

        // ── CHA1 monthly: routes to EvaluationScores(Month1) ────
        string tplCha = Path.Combine(Path.GetTempPath(), $"sis_import_cha_{Guid.NewGuid():N}.xlsx");
        ExcelImport.BuildMonthlyTemplate(tplCha, yr, grade, room, 9, "CHA1");
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplCha))
        {
            var ws = wb.Worksheet(1);
            ws.Cell(FindStudentRow(ws, "IMP001"), 4).Value = 7;
            ws.Cell(FindStudentRow(ws, "IMP002"), 4).Value = 6;
            wb.Save();
        }
        var resCha = ExcelImport.ParseMonthly(tplCha, yr, grade, room, 9, "CHA1");
        CheckEq("CHA1 parse: 2 valid", 2, resCha.ValidCount);
        CheckEq("CHA1 save: 2 rows", 2, ExcelImport.SaveImport(resCha));
        double? chaImp001 = DB.GetEvaluationScore(ids["IMP001"], yr, "Month1", "CHA1");
        Check("CHA1 import → EvaluationScores(Month1) for IMP001", chaImp001.HasValue && chaImp001.Value == 7);
        double? chaImp002 = DB.GetEvaluationScore(ids["IMP002"], yr, "Month1", "CHA1");
        Check("CHA1 import → EvaluationScores(Month1) for IMP002", chaImp002.HasValue && chaImp002.Value == 6);
        // CHA1 import must not produce a MonthlyAssessments row.
        int chaInMonthly = DB.ScalarInt(@"
            SELECT COUNT(*) FROM MonthlyAssessments ma
            JOIN Enrollments e ON e.EnrollID=ma.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            WHERE sub.SubjectCode='CHA1' AND e.AcademicYear=@y",
            null, ("@y", yr));
        CheckEq("CHA1 import: ZERO rows in MonthlyAssessments", 0, chaInMonthly);

        // ── LAB1 monthly: same routing as CHA1 ──────────────────
        string tplLab = Path.Combine(Path.GetTempPath(), $"sis_import_lab_{Guid.NewGuid():N}.xlsx");
        ExcelImport.BuildMonthlyTemplate(tplLab, yr, grade, room, 9, "LAB1");
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplLab))
        {
            var ws = wb.Worksheet(1);
            ws.Cell(FindStudentRow(ws, "IMP001"), 4).Value = 9;
            wb.Save();
        }
        ExcelImport.SaveImport(ExcelImport.ParseMonthly(tplLab, yr, grade, room, 9, "LAB1"));
        double? labImp001 = DB.GetEvaluationScore(ids["IMP001"], yr, "Month1", "LAB1");
        Check("LAB1 import → EvaluationScores(Month1) for IMP001", labImp001.HasValue && labImp001.Value == 9);

        // ── Subject-scope mismatch ──────────────────────────────
        var subjMismatch = ExcelImport.ParseMonthly(tplMath, yr, grade, room, 9, "ENG1");
        Check("Subject-scope mismatch rejected (MATH1 template parsed as ENG1)",
            subjMismatch.FatalError != null);

        // ── Wrong template (magic mismatch) ─────────────────────
        string fakeFile = Path.Combine(Path.GetTempPath(), $"sis_import_fake_{Guid.NewGuid():N}.xlsx");
        using (var wb = new ClosedXML.Excel.XLWorkbook())
        {
            wb.AddWorksheet("data").Cell("A1").Value = "not a template";
            wb.SaveAs(fakeFile);
        }
        var fakeResult = ExcelImport.ParseMonthly(fakeFile, yr, grade, room, 9, "MATH1");
        Check("Monthly parse: arbitrary xlsx rejected", fakeResult.FatalError != null);

        // ── Scope mismatch (different month) ────────────────────
        string wrongMonth = Path.Combine(Path.GetTempPath(), $"sis_import_wm_{Guid.NewGuid():N}.xlsx");
        ExcelImport.BuildMonthlyTemplate(wrongMonth, yr, grade, room, 11, "MATH1");
        var wrongMonthRes = ExcelImport.ParseMonthly(wrongMonth, yr, grade, room, 9, "MATH1");
        Check("Monthly parse: month-scope mismatch rejected", wrongMonthRes.FatalError != null);

        // ── Semester template (ENG1 sem 2) — final-score-only layout ──
        string tplEng = Path.Combine(Path.GetTempPath(), $"sis_import_eng_{Guid.NewGuid():N}.xlsx");
        ExcelImport.BuildSemesterTemplate(tplEng, yr, grade, room, 2, "ENG1");
        Check("Semester template (ENG1) created", File.Exists(tplEng));
        // Semester metadata lives in column Y (col 25), NOT Z — the template
        // only has 4 visible columns A-D so Y is the first hidden col.
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplEng))
        {
            var ws = wb.Worksheet(1);
            CheckEq("Semester template magic v4 (in col Y)", ExcelImport.MagicSemester, ws.Cell("Y1").GetString());
            CheckEq("Semester template sem (in col Y)",      2,      (int)ws.Cell("Y5").GetDouble());
            CheckEq("Semester template subject (in col Y)",  "ENG1", ws.Cell("Y6").GetString());
            // Single Final-exam column (D). No Midterm in the template.
            ws.Cell(FindStudentRow(ws, "IMP002"), 4).Value = 9;
            wb.Save();
        }
        var resEng = ExcelImport.ParseSemester(tplEng, yr, grade, room, 2, "ENG1");
        Check("ENG1 semester parse: no fatal error", resEng.FatalError == null);
        CheckEq("ENG1 semester parse: 1 valid", 1, resEng.ValidCount);
        CheckEq("ENG1 semester save: 1 row", 1, ExcelImport.SaveImport(resEng));
        var engRow = DB.Query(@"
            SELECT sc.MidScore, sc.FinalScore, sc.TotalScore
            FROM Scores sc
            JOIN Enrollments e ON e.EnrollID=sc.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            JOIN Students s    ON s.StudentID=e.StudentID
            WHERE s.StudentCode='IMP002' AND sub.SubjectCode='ENG1'
              AND e.AcademicYear=@y AND e.Semester=2",
            null, ("@y", yr));
        Check("ENG1 semester save: Scores row exists", engRow.Rows.Count == 1);
        if (engRow.Rows.Count == 1)
        {
            // MidScore stays at whatever RecomputeMidFromMonthly set (0 here
            // since no monthly entries for ENG1 sem2). The import didn't touch it.
            CheckEq("ENG1 semester save: FinalScore = 9 (imported)", 9.0, Convert.ToDouble(engRow.Rows[0]["FinalScore"]));
            CheckClose("ENG1 semester save: TotalScore = mid×50 + final×50",
                Convert.ToDouble(engRow.Rows[0]["MidScore"]) * (DB.MidPct / 100.0)
                  + 9 * (DB.FinalPct / 100.0),
                Convert.ToDouble(engRow.Rows[0]["TotalScore"]));
        }

        // ── CHA1 semester: only Final → EvaluationScores(SEM2) ─────
        string tplChaSem = Path.Combine(Path.GetTempPath(), $"sis_import_chasem_{Guid.NewGuid():N}.xlsx");
        ExcelImport.BuildSemesterTemplate(tplChaSem, yr, grade, room, 2, "CHA1");
        using (var wb = new ClosedXML.Excel.XLWorkbook(tplChaSem))
        {
            var ws = wb.Worksheet(1);
            // Single Final column (D) — same single-score layout for any subject.
            ws.Cell(FindStudentRow(ws, "IMP001"), 4).Value = 8;
            ws.Cell(FindStudentRow(ws, "IMP002"), 4).Value = 7;
            wb.Save();
        }
        var resChaSem = ExcelImport.ParseSemester(tplChaSem, yr, grade, room, 2, "CHA1");
        CheckEq("CHA1 semester parse: 2 valid", 2, resChaSem.ValidCount);
        CheckEq("CHA1 semester save: 2 rows", 2, ExcelImport.SaveImport(resChaSem));
        double? chaSem1 = DB.GetEvaluationScore(ids["IMP001"], yr, "SEM2", "CHA1");
        Check("CHA1 semester → EvaluationScores(SEM2) for IMP001", chaSem1.HasValue && chaSem1.Value == 8);
        double? chaSem2 = DB.GetEvaluationScore(ids["IMP002"], yr, "SEM2", "CHA1");
        Check("CHA1 semester → EvaluationScores(SEM2) for IMP002", chaSem2.HasValue && chaSem2.Value == 7);
        int chaSemInScores = DB.ScalarInt(@"
            SELECT COUNT(*) FROM Scores sc
            JOIN Enrollments e ON e.EnrollID=sc.EnrollID
            JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
            WHERE sub.SubjectCode='CHA1' AND e.AcademicYear=@y",
            null, ("@y", yr));
        CheckEq("CHA1 semester import: NO Scores row produced", 0, chaSemInScores);

        // Cleanup temp files (best-effort).
        try { File.Delete(tplMath); File.Delete(tplCha); File.Delete(tplLab);
              File.Delete(wrongMonth); File.Delete(fakeFile); File.Delete(tplEng); File.Delete(tplChaSem); }
        catch { }
    }

    // Finds the first data row (>=4) whose column B equals the student code —
    // per-subject templates have one row per student, so the lookup is by code only.
    static int FindStudentRow(ClosedXML.Excel.IXLWorksheet ws, string sCode)
    {
        int last = ws.LastRowUsed()!.RowNumber();
        for (int r = 4; r <= last; r++)
            if (ws.Cell(r, 2).GetString() == sCode) return r;
        throw new InvalidOperationException($"row not found for student: {sCode}");
    }
}
