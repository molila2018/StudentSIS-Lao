using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Admin-area queries
    //  SubjectsPage (idx 9) · UsersPage (idx 10) ·
    //  AcademicYearPage (idx 11) · SettingsPage (idx 12)
    // ════════════════════════════════════════════════════════════

    // Per-year record counts shown by both AcademicYearDeleteWindow
    // (what a cascade delete would remove) and AcademicYearStatsWindow.
    public class YearStats
    {
        public int Students, Active, Graduated, Withdrawn, Classrooms;
        public int Enrollments, Scores, Monthly, Evaluations, Attendance, History;
        public bool HasData =>
            Students + Enrollments + Scores + Monthly + Evaluations + Attendance + History > 0;
    }

    public static partial class DB
    {
        // ─── Subjects ─────────────────────────────────────────────

        // Filtered subject list for the admin grid. null = "ທັງໝົດ".
        // Grade-agnostic ('' / NULL) and semester-agnostic (0) subjects are
        // included in every specific filter — the 13 official MoES subjects
        // are stored that way.
        // Columns: ID · SortOrder · SubjectCode · SubjectName · GradeLevel ·
        //          Semester · Category
        public static DataTable SearchSubjects(string? grade, int? semester, string keyword)
        {
            var sql = new System.Text.StringBuilder(@"
                SELECT SubjectID AS ID, SortOrder, SubjectCode, SubjectName,
                       GradeLevel, Semester, Category
                FROM Subjects WHERE 1=1");
            var ps = new List<(string, object)>();
            if (grade != null)
            {
                sql.Append(" AND (GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel='')");
                ps.Add(("@g", grade));
            }
            if (semester != null)
            {
                sql.Append(" AND (Semester=@sm OR Semester=0)");
                ps.Add(("@sm", semester.Value));
            }
            if (!string.IsNullOrEmpty(keyword))
            {
                sql.Append(" AND (SubjectCode LIKE @k OR SubjectName LIKE @k)");
                ps.Add(("@k", $"%{keyword}%"));
            }
            sql.Append(" ORDER BY GradeLevel, SortOrder, SubjectCode");
            return Query(sql.ToString(), null, ps.ToArray());
        }

        public static string GetSubjectLabel(int subjectId)
        {
            var dt = Query("SELECT SubjectCode, SubjectName FROM Subjects WHERE SubjectID=@i",
                null, ("@i", subjectId));
            return dt.Rows.Count > 0
                ? $"{dt.Rows[0]["SubjectCode"]} — {dt.Rows[0]["SubjectName"]}"
                : $"ID {subjectId}";
        }

        // Dependent rows a subject delete would cascade away.
        public static (int enrolls, int scores, int monthly) GetSubjectCascadeCounts(int subjectId)
        {
            var p = ("@i", (object)subjectId);
            return (
                ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE SubjectID=@i", null, p),
                ScalarInt("SELECT COUNT(*) FROM Scores WHERE EnrollID IN (SELECT EnrollID FROM Enrollments WHERE SubjectID=@i)", null, p),
                ScalarInt("SELECT COUNT(*) FROM MonthlyAssessments WHERE EnrollID IN (SELECT EnrollID FROM Enrollments WHERE SubjectID=@i)", null, p));
        }

        public static void DeleteSubject(int subjectId) =>
            Exec("DELETE FROM Subjects WHERE SubjectID=@i", null, ("@i", subjectId));

        // Full row for the edit form.
        public static DataTable GetSubject(int subjectId) =>
            Query("SELECT * FROM Subjects WHERE SubjectID=@id", null, ("@id", subjectId));

        // Insert (id=0) or update one subject.
        public static void SaveSubject(int id, string code, string name,
            string gradeLevel, int semester, string category, int sortOrder)
        {
            string sql = id == 0
                ? @"INSERT INTO Subjects(SubjectCode,SubjectName,GradeLevel,Semester,Category,SortOrder)
                    VALUES(@c,@n,@g,@sm,@cat,@so)"
                : @"UPDATE Subjects
                      SET SubjectCode=@c, SubjectName=@n, GradeLevel=@g,
                          Semester=@sm, Category=@cat, SortOrder=@so
                    WHERE SubjectID=@id";
            var ps = new List<(string, object)> {
                ("@c", code), ("@n", name), ("@g", gradeLevel),
                ("@sm", semester), ("@cat", category), ("@so", sortOrder)
            };
            if (id > 0) ps.Add(("@id", id));
            Exec(sql, null, ps.ToArray());
        }

        // ─── Users ────────────────────────────────────────────────

        // Admin grid — Lao headers, bound via AutoGenerateColumns.
        public static DataTable GetUsersOverview() =>
            Query(@"SELECT UserID AS ID, Username AS ຊື່ຜູ້ໃຊ້, FullName AS ຊື່ເຕັມ,
                           Role AS ບົດບາດ,
                           CASE IsActive WHEN 1 THEN 'ໃຊ້ງານ' ELSE 'ປິດ' END AS ສະຖານະ,
                           IFNULL(LastLogin,'ຍັງບໍ່ເຄີຍ') AS ເຂົ້າໃຊ້ລ່າສຸດ
                    FROM Users ORDER BY Role, Username");

        // (username, isActive) — or null when the id doesn't exist.
        public static (string username, bool isActive)? GetUserStatus(int userId)
        {
            var dt = Query("SELECT Username, IsActive FROM Users WHERE UserID=@i", null, ("@i", userId));
            if (dt.Rows.Count == 0) return null;
            return (dt.Rows[0]["Username"].ToString()!,
                    System.Convert.ToInt32(dt.Rows[0]["IsActive"]) == 1);
        }

        public static void ToggleUserActive(int userId) =>
            Exec("UPDATE Users SET IsActive=CASE IsActive WHEN 1 THEN 0 ELSE 1 END WHERE UserID=@i",
                null, ("@i", userId));

        // Username / FullName / Role for the edit form + post-save verify.
        public static DataTable GetUserBasic(int userId) =>
            Query("SELECT Username, FullName, Role FROM Users WHERE UserID=@id",
                null, ("@id", userId));

        public static void InsertUser(string username, string password, string fullName, string role) =>
            Exec(@"INSERT INTO Users(Username, Password, FullName, Role)
                   VALUES(@u, @p, @n, @r)",
                null, ("@u", username), ("@p", password), ("@n", fullName), ("@r", role));

        public static void UpdateUserWithPassword(int userId, string username, string password, string fullName, string role) =>
            Exec(@"UPDATE Users SET Username=@u, Password=@p, FullName=@n, Role=@r
                   WHERE UserID=@id",
                null, ("@u", username), ("@p", password), ("@n", fullName), ("@r", role), ("@id", userId));

        public static void UpdateUserProfile(int userId, string username, string fullName, string role) =>
            Exec(@"UPDATE Users SET Username=@u, FullName=@n, Role=@r
                   WHERE UserID=@id",
                null, ("@u", username), ("@n", fullName), ("@r", role), ("@id", userId));

        public static int GetUserIdByUsername(string username) =>
            ScalarInt("SELECT UserID FROM Users WHERE Username=@u", null, ("@u", username));

        // Round-trip password check — same WHERE the production Login() uses,
        // so a successful verify means the credentials will actually work.
        public static bool VerifyUserPassword(int userId, string password) =>
            ScalarInt("SELECT COUNT(*) FROM Users WHERE UserID=@id AND Password=@p",
                null, ("@id", userId), ("@p", password)) == 1;

        // ─── Academic years ───────────────────────────────────────

        // Year-list grid: per-year student counts + current-year badge.
        public static DataTable GetAcademicYearsOverview() =>
            Query(@"
                SELECT y.Year,
                       (SELECT COUNT(*) FROM Students s WHERE s.AcademicYear=y.Year)                         AS StudentCount,
                       (SELECT COUNT(*) FROM Students s WHERE s.AcademicYear=y.Year AND s.Status='ຈົບ')      AS GraduatedCount,
                       (SELECT COUNT(*) FROM Students s WHERE s.AcademicYear=y.Year AND s.Status='ກຳລັງຮຽນ') AS ActiveCount,
                       IFNULL(SUBSTR(y.CreatedAt,1,10),'')                  AS CreatedDate,
                       CASE y.IsCurrent WHEN 1 THEN '🟢 ປະຈຸບັນ' ELSE '⚪' END AS StatusLabel,
                       y.IsCurrent                                          AS IsCurrent
                FROM AcademicYears y
                ORDER BY y.Year DESC");

        // System-wide totals for the page footer.
        public static (int total, int active, int graduated, int withdrawn, int classrooms)
            GetSystemWideStudentStats() =>
            (ScalarInt("SELECT COUNT(*) FROM Students"),
             ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ກຳລັງຮຽນ'"),
             ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ຈົບ'"),
             ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ອອກ'"),
             ScalarInt(@"SELECT COUNT(DISTINCT GradeLevel||'/'||IFNULL(ClassRoom,''))
                         FROM Students
                         WHERE GradeLevel IS NOT NULL AND GradeLevel<>''"));

        // Every record count tied to one academic year — drives both the
        // delete-confirmation summary and the read-only stats popup.
        public static YearStats GetYearStats(string year)
        {
            var p = ("@y", (object)year);
            return new YearStats {
                Students    = ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y", null, p),
                Active      = ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ກຳລັງຮຽນ'", null, p),
                Graduated   = ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ຈົບ'", null, p),
                Withdrawn   = ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ອອກ'", null, p),
                Classrooms  = ScalarInt(@"SELECT COUNT(DISTINCT GradeLevel||'/'||IFNULL(ClassRoom,''))
                                          FROM Students WHERE AcademicYear=@y
                                            AND GradeLevel IS NOT NULL AND GradeLevel<>''", null, p),
                Enrollments = ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, p),
                Scores      = ScalarInt(@"SELECT COUNT(*) FROM Scores sc
                                          JOIN Enrollments e ON e.EnrollID=sc.EnrollID
                                          WHERE e.AcademicYear=@y", null, p),
                Monthly     = ScalarInt(@"SELECT COUNT(*) FROM MonthlyAssessments ma
                                          JOIN Enrollments e ON e.EnrollID=ma.EnrollID
                                          WHERE e.AcademicYear=@y", null, p),
                Evaluations = ScalarInt("SELECT COUNT(*) FROM EvaluationScores WHERE AcademicYear=@y", null, p),
                Attendance  = ScalarInt("SELECT COUNT(*) FROM AttendanceRecords WHERE AcademicYear=@y", null, p),
                History     = ScalarInt("SELECT COUNT(*) FROM GradeHistory WHERE AcademicYear=@y", null, p),
            };
        }

        // Cascade-delete one academic year in a single transaction. Order
        // respects FK dependencies: Enrollments cascades Scores + Monthly via
        // FK; EvaluationScores / AttendanceRecords / GradeHistory go by year;
        // Students(year) cascades anything else they own; the registry row
        // goes last. Throws on failure (nothing is deleted — tx rolls back).
        public static void DeleteAcademicYearCascade(string year)
        {
            using var conn = Open();
            using var tx = conn.BeginTransaction();
            try
            {
                ExecTx("DELETE FROM Enrollments       WHERE AcademicYear=@y", conn, tx, ("@y", year));
                ExecTx("DELETE FROM EvaluationScores  WHERE AcademicYear=@y", conn, tx, ("@y", year));
                ExecTx("DELETE FROM AttendanceRecords WHERE AcademicYear=@y", conn, tx, ("@y", year));
                ExecTx("DELETE FROM GradeHistory      WHERE AcademicYear=@y", conn, tx, ("@y", year));
                ExecTx("DELETE FROM Students          WHERE AcademicYear=@y", conn, tx, ("@y", year));
                ExecTx("DELETE FROM AcademicYears     WHERE Year=@y",         conn, tx, ("@y", year));
                tx.Commit();
            }
            catch
            {
                tx.Rollback();
                throw;
            }
        }

        // ─── Activity log (SettingsPage → Log tab) ────────────────

        public static DataTable GetActivityLog(int limit = 500) =>
            Query(@"SELECT LoggedAt AS ເວລາ, Username AS ຜູ້ໃຊ້,
                           Action AS ການກະທຳ, Detail AS ລາຍລະອຽດ
                    FROM ActivityLog ORDER BY LogID DESC LIMIT @n",
                null, ("@n", limit));

        // Deletes log entries older than `days`; returns how many were removed.
        public static int PruneActivityLog(int days = 30) =>
            Exec($"DELETE FROM ActivityLog WHERE LoggedAt<datetime('now','-{days} days')");
    }
}
