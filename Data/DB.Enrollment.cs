using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Enrolment queries
    //  (EnrollmentPage idx 3 · BatchEnrollPage idx 4)
    // ════════════════════════════════════════════════════════════
    public static partial class DB
    {
        // Active students for the per-student enrolment picker.
        // grade=null → every grade. Only ‘ກຳລັງຮຽນ’ appear — graduated /
        // withdrawn students should never get new subjects.
        // Columns: StudentID · D ("CODE · Name · Grade/Room" display string)
        public static DataTable GetActiveStudentPicker(string year, string? grade)
        {
            var sql = new System.Text.StringBuilder(@"
                SELECT StudentID,
                       StudentCode || '  ·  ' || FirstName || ' ' || LastName ||
                       '  ·  ' || GradeLevel || '/' || IFNULL(ClassRoom,'-') AS D
                FROM Students
                WHERE Status='ກຳລັງຮຽນ' AND AcademicYear=@yr");
            var ps = new List<(string, object)> { ("@yr", year) };
            if (grade != null) { sql.Append(" AND GradeLevel=@g"); ps.Add(("@g", grade)); }
            sql.Append(" ORDER BY GradeLevel, ClassRoom, StudentCode");
            return Query(sql.ToString(), null, ps.ToArray());
        }

        // One row per subject the student is enrolled in for the year
        // (Sem 1 + Sem 2 collapse into one row — same subject identity).
        // Columns: SubID · SubjectCode · SubjectName · Category · Teacher
        public static DataTable GetStudentEnrolledSubjects(int studentId, string year) =>
            Query(@"
                SELECT s.SubjectID    AS SubID,
                       s.SubjectCode,
                       s.SubjectName,
                       s.Category,
                       MAX(e.Teacher)  AS Teacher
                FROM Enrollments e
                JOIN Subjects s ON s.SubjectID = e.SubjectID
                WHERE e.StudentID    = @sid
                  AND e.AcademicYear = @yr
                GROUP BY s.SubjectID
                ORDER BY s.SortOrder, s.SubjectCode",
                null, ("@sid", studentId), ("@yr", year));

        public static string GetStudentGradeLevel(int studentId) =>
            Scalar("SELECT GradeLevel FROM Students WHERE StudentID=@i",
                null, ("@i", studentId))?.ToString() ?? "";

        public static int CountSubjectsForGrade(string grade) =>
            ScalarInt(@"
                SELECT COUNT(*) FROM Subjects
                WHERE GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel=''",
                null, ("@g", grade));

        // Enrol one (student, subject, year) for BOTH semesters inside the
        // caller's transaction. INSERT OR IGNORE against the
        // UNIQUE(StudentID,SubjectID,AcademicYear,Semester) constraint makes
        // this idempotent. Returns how many rows were actually inserted (0-2).
        public static int EnrollBothSemesters(
            int studentId, int subjectId, string year,
            SQLiteConnection conn, SQLiteTransaction tx)
        {
            int inserted = 0;
            for (int sem = 1; sem <= 2; sem++)
                inserted += ExecTx(
                    @"INSERT OR IGNORE INTO Enrollments(StudentID, SubjectID, AcademicYear, Semester)
                      VALUES(@s, @sub, @y, @sm)",
                    conn, tx,
                    ("@s", studentId), ("@sub", subjectId), ("@y", year), ("@sm", sem));
            return inserted;
        }

        // Withdraw one subject for the year (both semesters). Scores +
        // MonthlyAssessments cascade automatically via FK ON DELETE CASCADE.
        public static void UnenrollSubject(int studentId, int subjectId, string year) =>
            Exec("DELETE FROM Enrollments WHERE StudentID=@s AND SubjectID=@sub AND AcademicYear=@y",
                null, ("@s", studentId), ("@sub", subjectId), ("@y", year));

        // Active students in (year, grade[, room]) for the batch-enrol preview.
        // room=null → every room.
        // Columns: StudentID · StudentCode · FullName · ClassRoom · AcademicYear
        public static DataTable GetActiveClassRoster(string year, string grade, string? room)
        {
            var sql = new System.Text.StringBuilder(@"
                SELECT StudentID, StudentCode,
                       FirstName || ' ' || LastName AS FullName,
                       ClassRoom, AcademicYear
                FROM Students
                WHERE Status='ກຳລັງຮຽນ' AND GradeLevel=@g AND AcademicYear=@y");
            var ps = new List<(string, object)> { ("@g", grade), ("@y", year) };
            if (room != null) { sql.Append(" AND ClassRoom=@r"); ps.Add(("@r", room)); }
            sql.Append(" ORDER BY ClassRoom, StudentCode");
            return Query(sql.ToString(), null, ps.ToArray());
        }
    }
}
