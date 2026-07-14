using System;
using System.Collections.Generic;
using System.Data;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Student queries (StudentsPage idx 2 + StudentFormWindow)
    // ════════════════════════════════════════════════════════════

    // All the fields the add/edit form captures for one student.
    // The form maps its controls into this; DB.SaveStudent maps it
    // to SQL. Ages are int? — null stores as SQL NULL.
    public class StudentDto
    {
        public string Code = "", FirstName = "", LastName = "", Gender = "", BirthDate = "";
        public string BirthVillage = "", BirthDistrict = "", BirthProvince = "";
        public string Village = "", District = "", Province = "";
        public string FatherName = "", FatherJob = "", FatherVillage = "", FatherDistrict = "", FatherProvince = "", FatherPhone = "";
        public string MotherName = "", MotherJob = "", MotherVillage = "", MotherDistrict = "", MotherProvince = "", MotherPhone = "";
        public int? FatherAge, MotherAge;
        public string GradeLevel = "", ClassRoom = "", AcademicYear = "", Status = "", Note = "";
    }

    public static partial class DB
    {
        // Filtered roster for the students list. Any filter passed as null
        // means "ທັງໝົດ" (no restriction). Keyword matches name, code, or
        // any parent-name column.
        // Columns: StudentID · StudentCode · FullName · Gender · GradeLevel ·
        //          ClassRoom · AcademicYear · ParentName · Status
        public static DataTable SearchStudents(
            string keyword, string? grade, string? room, string? year, string? status)
        {
            var sql = new System.Text.StringBuilder(@"
                SELECT StudentID, StudentCode,
                       FirstName||' '||LastName AS FullName, Gender,
                       GradeLevel, ClassRoom, AcademicYear,
                       COALESCE(NULLIF(FatherName,''), NULLIF(MotherName,''), NULLIF(ParentName,''), '') AS ParentName,
                       Status
                FROM Students WHERE 1=1");
            var ps = new List<(string, object)>();

            if (!string.IsNullOrEmpty(keyword))
            {
                sql.Append(" AND (FirstName||LastName LIKE @s OR StudentCode LIKE @s OR FatherName LIKE @s OR MotherName LIKE @s OR ParentName LIKE @s)");
                ps.Add(("@s", $"%{keyword}%"));
            }
            if (grade  != null) { sql.Append(" AND GradeLevel=@g");   ps.Add(("@g", grade)); }
            if (room   != null) { sql.Append(" AND ClassRoom=@r");    ps.Add(("@r", room)); }
            if (year   != null) { sql.Append(" AND AcademicYear=@y"); ps.Add(("@y", year)); }
            if (status != null) { sql.Append(" AND Status=@st");      ps.Add(("@st", status)); }
            sql.Append(" ORDER BY GradeLevel,ClassRoom,StudentCode");

            return Query(sql.ToString(), null, ps.ToArray());
        }

        // Full row for the edit form (SELECT * — the form reads columns by name).
        public static DataTable GetStudent(int studentId) =>
            Query("SELECT * FROM Students WHERE StudentID=@id", null, ("@id", studentId));

        // "CODE — First Last" for confirmation dialogs; falls back to "ID n".
        public static string GetStudentLabel(int studentId)
        {
            var dt = Query("SELECT StudentCode, FirstName, LastName FROM Students WHERE StudentID=@i",
                null, ("@i", studentId));
            return dt.Rows.Count > 0
                ? $"{dt.Rows[0]["StudentCode"]} — {dt.Rows[0]["FirstName"]} {dt.Rows[0]["LastName"]}"
                : $"ID {studentId}";
        }

        // How many dependent rows a hard-delete would cascade away —
        // shown in the delete-confirmation warning.
        public static (int enrolls, int scores, int monthly, int history)
            GetStudentCascadeCounts(int studentId)
        {
            var p = ("@i", (object)studentId);
            return (
                ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE StudentID=@i", null, p),
                ScalarInt("SELECT COUNT(*) FROM Scores WHERE EnrollID IN (SELECT EnrollID FROM Enrollments WHERE StudentID=@i)", null, p),
                ScalarInt("SELECT COUNT(*) FROM MonthlyAssessments WHERE EnrollID IN (SELECT EnrollID FROM Enrollments WHERE StudentID=@i)", null, p),
                ScalarInt("SELECT COUNT(*) FROM GradeHistory WHERE StudentID=@i", null, p));
        }

        // Hard delete; FK cascades remove dependent rows.
        public static void DeleteStudent(int studentId) =>
            Exec("DELETE FROM Students WHERE StudentID=@i", null, ("@i", studentId));

        // Insert (id=0) or update one student from the form DTO.
        // ParentName/ParentPhone legacy mirrors are kept in sync with the
        // father fields so old queries and CSV exports still show a parent.
        public static void SaveStudent(int id, StudentDto s)
        {
            string sql = id == 0
                ? @"INSERT INTO Students(
                      StudentCode, FirstName, LastName, Gender, BirthDate,
                      BirthVillage, BirthDistrict, BirthProvince,
                      Village, District, Province,
                      FatherName, FatherAge, FatherJob, FatherVillage, FatherDistrict, FatherProvince, FatherPhone,
                      MotherName, MotherAge, MotherJob, MotherVillage, MotherDistrict, MotherProvince, MotherPhone,
                      ParentName, ParentPhone,
                      GradeLevel, ClassRoom, AcademicYear, Status, Note)
                   VALUES(
                      @code, @fn, @ln, @gd, @bd,
                      @bvi, @bdi, @bpv,
                      @vi, @di, @pv,
                      @faN, @faA, @faJ, @faV, @faC, @faP, @faT,
                      @maN, @maA, @maJ, @maV, @maC, @maP, @maT,
                      @pn, @pp,
                      @gl, @rm, @yr, @st, @nt)"
                : @"UPDATE Students SET
                      StudentCode=@code, FirstName=@fn, LastName=@ln, Gender=@gd,
                      BirthDate=@bd,
                      BirthVillage=@bvi, BirthDistrict=@bdi, BirthProvince=@bpv,
                      Village=@vi, District=@di, Province=@pv,
                      FatherName=@faN, FatherAge=@faA, FatherJob=@faJ,
                      FatherVillage=@faV, FatherDistrict=@faC, FatherProvince=@faP, FatherPhone=@faT,
                      MotherName=@maN, MotherAge=@maA, MotherJob=@maJ,
                      MotherVillage=@maV, MotherDistrict=@maC, MotherProvince=@maP, MotherPhone=@maT,
                      ParentName=@pn, ParentPhone=@pp,
                      GradeLevel=@gl, ClassRoom=@rm, AcademicYear=@yr, Status=@st, Note=@nt
                   WHERE StudentID=@id";

            var ps = new List<(string, object)>
            {
                ("@code", s.Code), ("@fn", s.FirstName), ("@ln", s.LastName),
                ("@gd", s.Gender), ("@bd", s.BirthDate),
                ("@bvi", s.BirthVillage), ("@bdi", s.BirthDistrict), ("@bpv", s.BirthProvince),
                ("@vi", s.Village), ("@di", s.District), ("@pv", s.Province),
                ("@faN", s.FatherName), ("@faA", (object?)s.FatherAge ?? DBNull.Value),
                ("@faJ", s.FatherJob), ("@faV", s.FatherVillage), ("@faC", s.FatherDistrict),
                ("@faP", s.FatherProvince), ("@faT", s.FatherPhone),
                ("@maN", s.MotherName), ("@maA", (object?)s.MotherAge ?? DBNull.Value),
                ("@maJ", s.MotherJob), ("@maV", s.MotherVillage), ("@maC", s.MotherDistrict),
                ("@maP", s.MotherProvince), ("@maT", s.MotherPhone),
                ("@pn", s.FatherName), ("@pp", s.FatherPhone),
                ("@gl", s.GradeLevel), ("@rm", s.ClassRoom), ("@yr", s.AcademicYear),
                ("@st", s.Status), ("@nt", s.Note),
            };
            if (id > 0) ps.Add(("@id", id));
            Exec(sql, null, ps.ToArray());
        }

        // CSV bulk import: INSERT OR IGNORE one student (duplicate StudentCode
        // is silently skipped — returns 0). Status/Note aren't in the CSV so
        // schema defaults apply. ParentName/Phone mirror falls back to the
        // mother when the father fields are empty.
        // Returns rows inserted (1 = new, 0 = code already existed).
        public static int ImportStudentRow(StudentDto s, System.Data.SQLite.SQLiteConnection conn)
        {
            string pn = string.IsNullOrWhiteSpace(s.FatherName)  ? s.MotherName  : s.FatherName;
            string pp = string.IsNullOrWhiteSpace(s.FatherPhone) ? s.MotherPhone : s.FatherPhone;
            return Exec(@"INSERT OR IGNORE INTO Students(
                  StudentCode, FirstName, LastName, Gender, BirthDate,
                  BirthVillage, BirthDistrict, BirthProvince,
                  Village, District, Province,
                  FatherName, FatherAge, FatherJob, FatherVillage, FatherDistrict, FatherProvince, FatherPhone,
                  MotherName, MotherAge, MotherJob, MotherVillage, MotherDistrict, MotherProvince, MotherPhone,
                  ParentName, ParentPhone,
                  GradeLevel, ClassRoom, AcademicYear)
               VALUES(
                  @code,@fn,@ln,@gd,@bd,
                  @bvi,@bdi,@bpv,
                  @vi,@di,@pv,
                  @faN,@faA,@faJ,@faV,@faC,@faP,@faT,
                  @maN,@maA,@maJ,@maV,@maC,@maP,@maT,
                  @pn,@pp,
                  @gl,@rm,@yr)",
                conn,
                ("@code", s.Code), ("@fn", s.FirstName), ("@ln", s.LastName),
                ("@gd", s.Gender), ("@bd", s.BirthDate),
                ("@bvi", s.BirthVillage), ("@bdi", s.BirthDistrict), ("@bpv", s.BirthProvince),
                ("@vi", s.Village), ("@di", s.District), ("@pv", s.Province),
                ("@faN", s.FatherName), ("@faA", (object?)s.FatherAge ?? DBNull.Value),
                ("@faJ", s.FatherJob), ("@faV", s.FatherVillage), ("@faC", s.FatherDistrict),
                ("@faP", s.FatherProvince), ("@faT", s.FatherPhone),
                ("@maN", s.MotherName), ("@maA", (object?)s.MotherAge ?? DBNull.Value),
                ("@maJ", s.MotherJob), ("@maV", s.MotherVillage), ("@maC", s.MotherDistrict),
                ("@maP", s.MotherProvince), ("@maT", s.MotherPhone),
                ("@pn", pn), ("@pp", pp),
                ("@gl", s.GradeLevel), ("@rm", s.ClassRoom), ("@yr", s.AcademicYear));
        }
    }
}
