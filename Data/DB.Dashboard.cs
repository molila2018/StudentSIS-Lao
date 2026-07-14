using System.Data;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Dashboard queries (DashboardPage, idx 0)
    // ════════════════════════════════════════════════════════════
    public static partial class DB
    {
        public static int CountAllStudents() =>
            ScalarInt("SELECT COUNT(*) FROM Students");

        public static int CountActiveStudents() =>
            ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ກຳລັງຮຽນ'");

        public static int CountSubjects() =>
            ScalarInt("SELECT COUNT(*) FROM Subjects");

        // Students with at least one failing subject in the current
        // (year, semester). CHA1/LAB1 excluded — evaluation-only subjects
        // never contribute to pass/fail.
        public static int CountFailingStudents() =>
            ScalarInt(@"
                SELECT COUNT(DISTINCT e.StudentID) FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE sc.TotalScore<@p AND e.AcademicYear=@y AND e.Semester=@sm
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                null, ("@p", PassScore), ("@y", CurrentYear), ("@sm", CurrentSem));

        // Per-grade totals for the bar chart + summary table.
        // Columns: GradeLevel · Count · Active
        public static DataTable GetStudentCountsByGrade() =>
            Query(@"
                SELECT GradeLevel, COUNT(*) AS Count,
                       SUM(CASE Status WHEN 'ກຳລັງຮຽນ' THEN 1 ELSE 0 END) AS Active
                FROM Students GROUP BY GradeLevel ORDER BY GradeLevel");

        // Latest announcements for the side panel.
        // Columns: Title · CreatedAt
        public static DataTable GetRecentAnnouncements(int limit = 5) =>
            Query("SELECT Title, CreatedAt FROM Announcements ORDER BY AnnID DESC LIMIT @n",
                null, ("@n", limit));
    }
}
