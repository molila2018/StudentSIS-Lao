using System.Data;
using System.Data.SQLite;

namespace StudentSIS.Data
{
    // ════════════════════════════════════════════════════════════
    //  DB — Score-entry queries
    //
    //  Serves the two score-entry pages:
    //    ScoresPage (idx 5)        semester finals + CHA1/LAB1 SEM/ANNUAL
    //    MonthlyScoresPage (idx 6) monthly sub-scores + CHA1/LAB1 Month{N}
    //
    //  Views call these named methods only — no SQL in view files.
    // ════════════════════════════════════════════════════════════
    public static partial class DB
    {
        // Subjects valid for a grade (grade-specific + grade-agnostic),
        // ordered like the official score sheet. evalOnly=true narrows to
        // CHA1/LAB1 — used by ANNUAL mode where academic subjects have no
        // manually-entered score.
        // Columns: SubjectID · SubjectCode · Display ("CODE  Name")
        public static DataTable GetSubjectsForGrade(string grade, bool evalOnly = false)
        {
            string evalFilter = evalOnly ? "AND SubjectCode IN ('CHA1','LAB1')" : "";
            return Query($@"
                SELECT SubjectID, SubjectCode, SubjectCode||'  '||SubjectName AS Display
                FROM Subjects
                WHERE (GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel='') {evalFilter}
                ORDER BY SortOrder",
                null, ("@g", grade));
        }

        // Whole-class roster for one (subject, year, semester) joined to each
        // student's enrollment + score row. LEFT JOINs keep students without
        // an enrollment visible (EnrollID=0) so the page can flag rather than
        // silently drop them. status=null → every status.
        // Columns: StudentID · StudentCode · FullName · EnrollID · ScoreID ·
        //          MidScore · FinalScore
        public static DataTable GetSemesterScoreRoster(
            int subjectId, string year, int sem, string grade, string room, string? status)
        {
            string statusFilter = status == null ? "" : "AND s.Status=@st";
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@subId", subjectId), ("@year", year), ("@sem", sem),
                ("@grade", grade), ("@room", room)
            };
            if (status != null) ps.Add(("@st", status));

            return Query($@"
                SELECT s.StudentID, s.StudentCode, s.FirstName||' '||s.LastName AS FullName,
                       IFNULL(e.EnrollID,0)         AS EnrollID,
                       IFNULL(sc.ScoreID,0)         AS ScoreID,
                       IFNULL(sc.MidScore,0)        AS MidScore,
                       IFNULL(sc.FinalScore,0)      AS FinalScore
                FROM Students s
                LEFT JOIN Enrollments e ON e.StudentID=s.StudentID
                                       AND e.SubjectID=@subId
                                       AND e.AcademicYear=@year
                                       AND e.Semester=@sem
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE s.GradeLevel=@grade AND s.ClassRoom=@room {statusFilter}
                ORDER BY s.StudentCode",
                null, ps.ToArray());
        }

        // Insert-or-update one enrollment's Scores row inside the caller's
        // transaction. scoreId=0 → the row doesn't exist yet → INSERT (Mid is
        // written once here; afterwards it's owned by RecomputeMidFromMonthly).
        // Existing row → UPDATE Final/Total/Level only, preserving Mid.
        public static void SaveSemesterScore(
            int scoreId, int enrollId, double mid, double fin, double total, string level,
            SQLiteConnection conn, SQLiteTransaction tx)
        {
            if (scoreId == 0)
                ExecTx(@"INSERT INTO Scores(EnrollID,MidScore,FinalScore,TotalScore,Level)
                         VALUES(@e,@m,@f,@t,@l)",
                    conn, tx,
                    ("@e", enrollId), ("@m", mid), ("@f", fin), ("@t", total), ("@l", level));
            else
                ExecTx(@"UPDATE Scores
                         SET FinalScore=@f, TotalScore=@t, Level=@l,
                             UpdatedAt=datetime('now','localtime')
                         WHERE ScoreID=@id",
                    conn, tx,
                    ("@f", fin), ("@t", total), ("@l", level), ("@id", scoreId));
        }

        // Class roster for one (subject, month) on the monthly page — enrolled
        // students only (INNER JOIN Enrollments), each with that month's three
        // sub-scores (0 when not yet entered). status=null → every status.
        // Columns: StudentID · StudentCode · FullName · ClassRoom · EnrollID ·
        //          ActivityScore · DisciplineScore · HomeworkScore
        public static DataTable GetMonthlyScoreRoster(
            int subjectId, string year, int month, string grade, string room, string? status)
        {
            int sem = SemesterForMonth(month);
            string statusFilter = status == null ? "" : "AND s.Status=@st";
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@subId", subjectId), ("@year", year), ("@sem", sem),
                ("@month", month), ("@grade", grade), ("@room", room)
            };
            if (status != null) ps.Add(("@st", status));

            return Query($@"
                SELECT s.StudentID, s.StudentCode, s.FirstName||' '||s.LastName AS FullName, s.ClassRoom,
                       e.EnrollID,
                       IFNULL(ma.ActivityScore,0)   AS ActivityScore,
                       IFNULL(ma.DisciplineScore,0) AS DisciplineScore,
                       IFNULL(ma.HomeworkScore,0)   AS HomeworkScore
                FROM Students s
                JOIN Enrollments e ON e.StudentID=s.StudentID
                                  AND e.SubjectID=@subId
                                  AND e.AcademicYear=@year
                                  AND e.Semester=@sem
                LEFT JOIN MonthlyAssessments ma ON ma.EnrollID=e.EnrollID AND ma.Month=@month
                WHERE s.GradeLevel=@grade AND s.ClassRoom=@room {statusFilter}
                ORDER BY s.StudentCode",
                null, ps.ToArray());
        }

        // Class roster for CHA1/LAB1 on the monthly page — evaluation scores
        // are keyed by student (no enrollment), one /10 value per context.
        // status=null → every status.
        // Columns: StudentID · StudentCode · FullName · ClassRoom · EvalScore
        public static DataTable GetEvalScoreRoster(
            string year, string context, string subjectCode,
            string grade, string room, string? status)
        {
            string statusFilter = status == null ? "" : "AND s.Status=@st";
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@year", year), ("@ctx", context), ("@code", subjectCode),
                ("@grade", grade), ("@room", room)
            };
            if (status != null) ps.Add(("@st", status));

            return Query($@"
                SELECT s.StudentID, s.StudentCode, s.FirstName||' '||s.LastName AS FullName, s.ClassRoom,
                       ev.Score AS EvalScore
                FROM Students s
                LEFT JOIN EvaluationScores ev
                       ON ev.StudentID=s.StudentID AND ev.AcademicYear=@year
                      AND ev.Context=@ctx AND ev.SubjectCode=@code
                WHERE s.GradeLevel=@grade AND s.ClassRoom=@room {statusFilter}
                ORDER BY s.StudentCode",
                null, ps.ToArray());
        }

        // Upsert one student's monthly sub-scores inside the caller's
        // transaction. UNIQUE(EnrollID,Month) drives the ON CONFLICT path.
        public static void SaveMonthlyAssessment(
            int enrollId, int month, double activity, double discipline, double homework,
            SQLiteConnection conn, SQLiteTransaction tx) =>
            ExecTx(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore,UpdatedAt)
                     VALUES(@e,@mo,@a,@d,@h,datetime('now','localtime'))
                     ON CONFLICT(EnrollID,Month) DO UPDATE SET
                       ActivityScore=excluded.ActivityScore,
                       DisciplineScore=excluded.DisciplineScore,
                       HomeworkScore=excluded.HomeworkScore,
                       UpdatedAt=datetime('now','localtime')",
                conn, tx,
                ("@e", enrollId), ("@mo", month),
                ("@a", activity), ("@d", discipline), ("@h", homework));
    }
}
