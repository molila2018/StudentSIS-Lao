using System.Collections.Generic;
using System.Data;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Report-builder queries (ReportPage + history windows)
    //
    //  Two families:
    //    • per-student  — one student's per-subject scores under a scope
    //                     (month / semester / annual); also used to rank
    //                     classmates one by one
    //    • class bulk   — one query per (roster × subject) grid so the
    //                     class reports don't loop per student
    //
    //  All "academic" queries exclude CHA1/LAB1 (evaluation-only —
    //  never averaged); eval queries fetch ONLY CHA1/LAB1.
    // ════════════════════════════════════════════════════════════
    public static partial class DB
    {
        // ─── Per-student subject scores ───────────────────────────

        // Academic subjects × one month: V = activity+discipline+homework,
        // HasRow = a MonthlyAssessments row exists for that month.
        // Columns: SubjectCode · V · HasRow
        public static DataTable GetStudentMonthlySubjectTotals(int studentId, string year, int month) =>
            Query(@"
                SELECT sub.SubjectCode,
                       IFNULL(ma.ActivityScore,0)+IFNULL(ma.DisciplineScore,0)+IFNULL(ma.HomeworkScore,0) AS V,
                       (ma.MonthlyID IS NOT NULL) AS HasRow
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@m
                WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                null, ("@s", studentId), ("@y", year), ("@m", month),
                      ("@sm", SemesterForMonth(month)));

        // Academic subjects × one semester: V = Scores.TotalScore.
        // Columns: SubjectCode · V · HasRow
        public static DataTable GetStudentSemesterSubjectTotals(int studentId, string year, int sem) =>
            Query(@"
                SELECT sub.SubjectCode, sc.TotalScore AS V,
                       (sc.ScoreID IS NOT NULL) AS HasRow
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                null, ("@s", studentId), ("@y", year), ("@sm", sem));

        // Academic subjects × the whole year — one row per existing Scores row
        // (both semesters). The caller averages per subject for the annual mean.
        // Columns: SubjectCode · V
        public static DataTable GetStudentAnnualSubjectTotals(int studentId, string year) =>
            Query(@"
                SELECT sub.SubjectCode, IFNULL(sc.TotalScore,0) AS V
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE e.StudentID=@s AND e.AcademicYear=@y
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                  AND sc.ScoreID IS NOT NULL",
                null, ("@s", studentId), ("@y", year));

        // ─── Class-wide bulk lookups (one query per report) ───────

        // Per (student, subject) monthly totals for a set of students. ALL
        // subjects included — the class template renders CHA1/LAB1 raw in
        // their own columns, outside the SUM/AVG/RANK formula range.
        // Columns: StudentID · SubjectCode · Total · HasRow
        public static DataTable GetClassMonthlyTotalsBulk(
            IEnumerable<int> studentIds, string year, int sem, int month)
        {
            string idCsv = JoinIds(studentIds);
            return Query($@"
                SELECT e.StudentID, sub.SubjectCode,
                       (IFNULL(ma.ActivityScore,0)+IFNULL(ma.DisciplineScore,0)+IFNULL(ma.HomeworkScore,0)) AS Total,
                       (ma.MonthlyID IS NOT NULL) AS HasRow
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@m
                WHERE e.AcademicYear=@y AND e.Semester=@sm
                  AND e.StudentID IN ({idCsv})",
                null, ("@m", month), ("@y", year), ("@sm", sem));
        }

        // Per (student, subject) semester totals — academic only.
        // Columns: StudentID · SubjectCode · Total
        public static DataTable GetClassSemesterTotalsBulk(
            IEnumerable<int> studentIds, string year, int sem)
        {
            string idCsv = JoinIds(studentIds);
            return Query($@"
                SELECT e.StudentID, sub.SubjectCode, IFNULL(sc.TotalScore,0) AS Total
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE e.AcademicYear=@y AND e.Semester=@sm
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                  AND e.StudentID IN ({idCsv})",
                null, ("@y", year), ("@sm", sem));
        }

        // Per (student, subject) score rows across BOTH semesters — academic
        // only; the caller averages per (student, subject) for the annual mean.
        // Columns: StudentID · SubjectCode · Total
        public static DataTable GetClassAnnualTotalsBulk(
            IEnumerable<int> studentIds, string year)
        {
            string idCsv = JoinIds(studentIds);
            return Query($@"
                SELECT e.StudentID, sub.SubjectCode, IFNULL(sc.TotalScore,0) AS Total
                FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE e.AcademicYear=@y
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                  AND sc.ScoreID IS NOT NULL
                  AND e.StudentID IN ({idCsv})",
                null, ("@y", year));
        }

        // CHA1/LAB1 manual scores for a set of students under one context
        // (SEM1 / SEM2 / ANNUAL).
        // Columns: StudentID · SubjectCode · Score
        public static DataTable GetClassEvalScoresBulk(
            IEnumerable<int> studentIds, string year, string context)
        {
            string idCsv = JoinIds(studentIds);
            return Query($@"
                SELECT StudentID, SubjectCode, Score
                FROM EvaluationScores
                WHERE AcademicYear=@y AND Context=@c
                  AND SubjectCode IN ('CHA1','LAB1')
                  AND StudentID IN ({idCsv})",
                null, ("@y", year), ("@c", context));
        }

        // ─── Small lookups ────────────────────────────────────────

        // Student picker narrowed to one classroom (grade + room), with an
        // optional status filter layered on top (null = every status).
        // Unlike GetActiveStudentPicker this has no year filter — the report
        // picker browses the class as it currently stands.
        // Columns: StudentID · D ("CODE · Name · Grade/Room" display string)
        public static DataTable GetStudentPickerForClass(string grade, string room, string? status)
        {
            var sql = new System.Text.StringBuilder(@"
                SELECT StudentID,
                       StudentCode || '  ·  ' || FirstName || ' ' || LastName ||
                       '  ·  ' || GradeLevel || '/' || IFNULL(ClassRoom,'-') AS D
                FROM Students
                WHERE GradeLevel=@g AND ClassRoom=@r");
            var ps = new List<(string, object)> { ("@g", grade), ("@r", room) };
            if (status != null) { sql.Append(" AND Status=@st"); ps.Add(("@st", status)); }
            sql.Append(" ORDER BY StudentCode");
            return Query(sql.ToString(), null, ps.ToArray());
        }

        // Report header line: student code, full name, status.
        // Columns: StudentCode · FullName · Status
        public static DataTable GetStudentHeader(int studentId) =>
            Query(@"SELECT StudentCode, FirstName||' '||LastName AS FullName,
                           IFNULL(Status,'') AS Status
                    FROM Students WHERE StudentID=@id",
                null, ("@id", studentId));

        // Every academic subject code (CHA1/LAB1 excluded).
        public static List<string> GetAcademicSubjectCodes()
        {
            var codes = new List<string>();
            foreach (DataRow r in Query(
                "SELECT SubjectCode FROM Subjects WHERE SubjectCode NOT IN ('CHA1','LAB1')").Rows)
                codes.Add(r["SubjectCode"].ToString()!);
            return codes;
        }

        // SubjectCode → SubjectName for display labels.
        public static Dictionary<string, string> GetSubjectNameMap()
        {
            var map = new Dictionary<string, string>();
            foreach (DataRow r in Query("SELECT SubjectCode, SubjectName FROM Subjects").Rows)
                map[r["SubjectCode"].ToString()!] = r["SubjectName"].ToString()!;
            return map;
        }

        // ids are ints — string.Join over them has no injection surface.
        // "0" fallback keeps `IN ()` valid SQL when the roster is empty.
        private static string JoinIds(IEnumerable<int> ids)
        {
            string csv = string.Join(",", ids);
            return csv.Length == 0 ? "0" : csv;
        }
    }
}
