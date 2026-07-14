using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Promotion queries (PromotionPage, idx 8)
    //  Promote (ຂຶ້ນຊັ້ນ) · Repeat (ຊ້ຳຊັ້ນ) · Graduate (ຈົບ)
    // ════════════════════════════════════════════════════════════
    public static partial class DB
    {
        // Current Students fields for a set of ids (the historical cohort from
        // GetHistoricalClassRoster) — enriches the promotion roster grid.
        // ids are ints joined internally, so no injection surface.
        // Columns: StudentID · StudentCode · FullName · Gender · GradeLevel ·
        //          ClassRoom · AcademicYear · Status
        public static DataTable GetStudentsByIds(IEnumerable<int> studentIds)
        {
            string idCsv = string.Join(",", studentIds);
            if (idCsv.Length == 0) return new DataTable();
            return Query($@"
                SELECT StudentID, StudentCode, FirstName||' '||LastName AS FullName,
                       Gender, GradeLevel, ClassRoom, AcademicYear, Status
                FROM Students WHERE StudentID IN ({idCsv})
                ORDER BY StudentCode");
        }

        // Mark one student graduated inside the caller's transaction.
        // Students row stays anchored to the finishing grade + year; the
        // GradeHistory row records ToGrade='ຈົບ' under the year AFTER the
        // finishing year (the year being "entered"), per the app convention.
        public static void GraduateStudent(
            int studentId, string fromGrade, string classRoom, string yearAfterGrad,
            SQLiteConnection conn, SQLiteTransaction tx)
        {
            ExecTx("UPDATE Students SET Status='ຈົບ' WHERE StudentID=@id",
                conn, tx, ("@id", studentId));
            ExecTx(@"INSERT INTO GradeHistory
                     (StudentID, FromGrade, ToGrade, AcademicYear, ClassRoom, Note, ChangedBy)
                     VALUES (@sid, @fg, 'ຈົບ', @y, @cr, @n, @by)",
                conn, tx,
                ("@sid", studentId), ("@fg", fromGrade),
                ("@y", yearAfterGrad),
                ("@cr", string.IsNullOrEmpty(classRoom) ? DBNull.Value : (object)classRoom),
                ("@n", "ຈົບການສຶກສາ"), ("@by", CurrentUser));
        }

        // Move one student to (grade, room, year, status) and log the change.
        // Used by both ຂຶ້ນຊັ້ນ and ຊ້ຳຊັ້ນ — `note` carries which one, and
        // `historyRoom` records the room the student was in BEFORE the move
        // (GradeHistory.ClassRoom describes the FromGrade year).
        public static void ApplyPromotion(
            int studentId, string fromGrade, string toGrade, string newStatus,
            string newRoom, string newYear, string historyYear, string historyRoom,
            string historyToGrade, string note,
            SQLiteConnection conn, SQLiteTransaction tx)
        {
            ExecTx(@"UPDATE Students
                     SET GradeLevel=@g, ClassRoom=@cr, AcademicYear=@y, Status=@st
                     WHERE StudentID=@id",
                conn, tx,
                ("@g", toGrade), ("@cr", newRoom),
                ("@y", newYear), ("@st", newStatus),
                ("@id", studentId));
            ExecTx(@"INSERT INTO GradeHistory
                     (StudentID, FromGrade, ToGrade, AcademicYear, ClassRoom, Note, ChangedBy)
                     VALUES (@sid, @fg, @tg, @y, @cr, @n, @by)",
                conn, tx,
                ("@sid", studentId), ("@fg", fromGrade), ("@tg", historyToGrade),
                ("@y", historyYear),
                ("@cr", string.IsNullOrEmpty(historyRoom) ? DBNull.Value : (object)historyRoom),
                ("@n", note), ("@by", CurrentUser));
        }

        // Recent GradeHistory entries for the page's audit panel.
        // Column headers are Lao — bound straight into an auto-generate grid.
        public static DataTable GetPromotionHistory(int limit = 50) =>
            Query(@"SELECT h.ChangedAt   AS ວັນທີ,
                           s.StudentCode AS ລະຫັດ,
                           s.FirstName||' '||s.LastName AS ຊື່ນັກຮຽນ,
                           h.FromGrade   AS ຈາກຊັ້ນ,
                           h.ToGrade     AS ຂຶ້ນຊັ້ນ,
                           h.AcademicYear AS ປີ,
                           IFNULL(h.ClassRoom,'') AS ຫ້ອງ,
                           h.ChangedBy   AS ດຳເນີນໂດຍ
                    FROM GradeHistory h JOIN Students s ON s.StudentID=h.StudentID
                    ORDER BY h.HistoryID DESC LIMIT @n",
                null, ("@n", limit));
    }
}
