using System.Collections.Generic;
using System.Data;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Class-hub queries (ClassHubPage, idx 1)
    // ════════════════════════════════════════════════════════════
    public static partial class DB
    {
        // Counts for one grade-picker card on the landing view.
        public static (int total, int active, int graduated) GetGradeCardStats(string grade)
        {
            var p = ("@g", (object)grade);
            return (
                ScalarInt("SELECT COUNT(*) FROM Students WHERE GradeLevel=@g", null, p),
                ScalarInt("SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND Status='ກຳລັງຮຽນ'", null, p),
                ScalarInt("SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND Status='ຈົບ'", null, p));
        }

        // Distinct rooms currently in use by active students of a grade,
        // for the card's "ຫ້ອງ 1 · 2 · 3" line.
        public static List<string> GetActiveRoomsForGrade(string grade)
        {
            var dt = Query(@"
                SELECT DISTINCT ClassRoom FROM Students
                WHERE GradeLevel=@g AND ClassRoom IS NOT NULL AND ClassRoom<>''
                  AND Status='ກຳລັງຮຽນ'
                ORDER BY ClassRoom",
                null, ("@g", grade));
            var rooms = new List<string>();
            foreach (DataRow r in dt.Rows) rooms.Add(r[0]?.ToString() ?? "");
            return rooms;
        }

        // Live stats row for the per-class hub view: student totals for the
        // picked (grade, room, year) + how many distinct subjects that class
        // has enrolments in.
        public static (int total, int active, int subjects) GetHubStats(
            string grade, string room, string year)
        {
            int total = ScalarInt(
                "SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND ClassRoom=@r AND AcademicYear=@y",
                null, ("@g", grade), ("@r", room), ("@y", year));
            int active = ScalarInt(
                "SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND ClassRoom=@r AND AcademicYear=@y AND Status='ກຳລັງຮຽນ'",
                null, ("@g", grade), ("@r", room), ("@y", year));
            int subjects = ScalarInt(@"
                SELECT COUNT(DISTINCT e.SubjectID)
                FROM Enrollments e
                JOIN Students   s ON s.StudentID = e.StudentID
                WHERE s.GradeLevel=@g AND s.ClassRoom=@r AND e.AcademicYear=@y",
                null, ("@g", grade), ("@r", room), ("@y", year));
            return (total, active, subjects);
        }
    }
}
