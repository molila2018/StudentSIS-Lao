using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace StudentSIS.Data
{
    // ── Partial class layout ─────────────────────────────────────
    // DB is split across domain files so each area stays small and
    // findable. This file is the CORE: connection, schema/migrations,
    // settings cache, calc helpers, and cross-page query helpers.
    // Page-specific queries live in sibling DB.*.cs files:
    //   DB.ScoreEntry.cs   score-entry rosters + saves (idx 5, 6)
    //   DB.Students.cs     student CRUD + roster (idx 2)
    //   DB.Dashboard.cs    landing-page stats (idx 0)
    //   DB.ClassHub.cs     grade cards + hub stats (idx 1)
    //   DB.Enrollment.cs   per-student + batch enrolment (idx 3, 4)
    //   DB.Promotion.cs    promote / repeat / graduate (idx 8)
    //   DB.Admin.cs        Subjects · Users · Years · Settings (idx 9-12)
    //   DB.Reports.cs      ReportPage document data (idx 7)
    // RULE: View files must never contain SQL — add a named method
    // in the matching domain file instead.
    public static partial class DB
    {
        private static readonly string DbPath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sis_lao.db");
        public static string ConnectionString => $"Data Source={DbPath};Version=3;";

        public static int    CurrentUserId { get; set; }
        public static string CurrentUser   { get; set; } = "";
        public static string CurrentRole   { get; set; } = "";

        // ── NavContext: one-shot filter handoff between pages ──────
        // ClassHubPage sets these before triggering navigation to MonthlyScoresPage /
        // StudentsPage / etc. The destination page reads them in its constructor and
        // calls ClearNav() so the values don't leak into subsequent manual navigations.
        public static string NavGrade    = "";
        public static string NavRoom     = "";
        public static string NavYear     = "";
        public static int    NavMonth    = 0;
        public static int    NavSemester = 0;
        public static void ClearNav()
        {
            NavGrade = ""; NavRoom = ""; NavYear = "";
            NavMonth = 0;  NavSemester = 0;
        }
        public static string SchoolName    { get; private set; } = "ໂຮງຮຽນ";
        public static string CurrentYear   { get; private set; } = "2025-2026";
        public static int    CurrentSem    { get; private set; } = 1;
        public static double MidPct        { get; private set; } = 60.0;
        public static double FinalPct      { get; private set; } = 40.0;
        public static double PassScore     { get; private set; } = 5.0;

        public static void Initialize()
        {
            if (!File.Exists(DbPath)) SQLiteConnection.CreateFile(DbPath);
            using var conn = Open();
            Exec(@"
PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS Users(
    UserID INTEGER PRIMARY KEY AUTOINCREMENT, Username TEXT NOT NULL UNIQUE,
    Password TEXT NOT NULL, FullName TEXT NOT NULL, Role TEXT NOT NULL DEFAULT 'teacher',
    IsActive INTEGER NOT NULL DEFAULT 1, LastLogin TEXT,
    CreatedAt TEXT DEFAULT(datetime('now','localtime')));
CREATE TABLE IF NOT EXISTS Students(
    StudentID INTEGER PRIMARY KEY AUTOINCREMENT, StudentCode TEXT NOT NULL UNIQUE,
    FirstName TEXT NOT NULL, LastName TEXT NOT NULL, Gender TEXT DEFAULT 'ຊາຍ',
    BirthDate TEXT, NationalID TEXT, Address TEXT,
    -- Birth place (st_vill_born / st_city_born / st_prov_born)
    BirthVillage TEXT, BirthDistrict TEXT, BirthProvince TEXT,
    -- Current address (st_vill_now / st_city_now / st_prov_now)
    Village TEXT, District TEXT, Province TEXT,
    Phone TEXT,
    -- Father (st_fa_*)
    FatherName TEXT, FatherAge INTEGER, FatherJob TEXT,
    FatherVillage TEXT, FatherDistrict TEXT, FatherProvince TEXT, FatherPhone TEXT,
    -- Mother (st_ma_*)
    MotherName TEXT, MotherAge INTEGER, MotherJob TEXT,
    MotherVillage TEXT, MotherDistrict TEXT, MotherProvince TEXT, MotherPhone TEXT,
    -- Legacy (kept for backward compatibility with old CSVs / pre-migration data)
    ParentName TEXT, ParentPhone TEXT, PhotoPath TEXT,
    GradeLevel TEXT NOT NULL, ClassRoom TEXT, AcademicYear TEXT NOT NULL,
    Status TEXT NOT NULL DEFAULT 'ກຳລັງຮຽນ', Note TEXT,
    CreatedAt TEXT DEFAULT(datetime('now','localtime')));
CREATE TABLE IF NOT EXISTS Subjects(
    SubjectID INTEGER PRIMARY KEY AUTOINCREMENT, SubjectCode TEXT NOT NULL UNIQUE,
    SubjectName TEXT NOT NULL, GradeLevel TEXT, Semester INTEGER DEFAULT 1,
    Category TEXT DEFAULT 'ວິຊາສາມັນ', SortOrder INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS Enrollments(
    EnrollID INTEGER PRIMARY KEY AUTOINCREMENT,
    StudentID INTEGER NOT NULL REFERENCES Students(StudentID) ON DELETE CASCADE,
    SubjectID INTEGER NOT NULL REFERENCES Subjects(SubjectID) ON DELETE CASCADE,
    AcademicYear TEXT NOT NULL, Semester INTEGER NOT NULL, Teacher TEXT,
    UNIQUE(StudentID,SubjectID,AcademicYear,Semester));
CREATE TABLE IF NOT EXISTS Scores(
    ScoreID INTEGER PRIMARY KEY AUTOINCREMENT,
    EnrollID INTEGER NOT NULL UNIQUE REFERENCES Enrollments(EnrollID) ON DELETE CASCADE,
    MidScore REAL DEFAULT 0, FinalScore REAL DEFAULT 0,
    TotalScore REAL DEFAULT 0, Level TEXT DEFAULT '', Remarks TEXT,
    UpdatedAt TEXT DEFAULT(datetime('now','localtime')));
-- Monthly continuous assessment: Activity(/3) + Discipline(/2) + Homework(/5) = /10.
-- Per (enrollment × calendar month). Monthly assessment months only:
--   Sem 1 = {9,10,11,12}  (January is the final-exam month, recorded as Scores.FinalScore)
--   Sem 2 = {2,3,4,5}     (June  is the final-exam month, recorded as Scores.FinalScore)
CREATE TABLE IF NOT EXISTS MonthlyAssessments(
    MonthlyID INTEGER PRIMARY KEY AUTOINCREMENT,
    EnrollID INTEGER NOT NULL REFERENCES Enrollments(EnrollID) ON DELETE CASCADE,
    Month INTEGER NOT NULL,
    ActivityScore REAL DEFAULT 0,
    DisciplineScore REAL DEFAULT 0,
    HomeworkScore REAL DEFAULT 0,
    UpdatedAt TEXT DEFAULT(datetime('now','localtime')),
    UNIQUE(EnrollID, Month));
CREATE TABLE IF NOT EXISTS AttendanceRecords(
    AttendID INTEGER PRIMARY KEY AUTOINCREMENT,
    StudentID INTEGER NOT NULL REFERENCES Students(StudentID) ON DELETE CASCADE,
    SubjectID INTEGER NOT NULL REFERENCES Subjects(SubjectID) ON DELETE CASCADE,
    AttendDate TEXT NOT NULL, AcademicYear TEXT NOT NULL, Semester INTEGER NOT NULL,
    Status TEXT NOT NULL DEFAULT 'present', Note TEXT, RecordedBy TEXT,
    UNIQUE(StudentID,SubjectID,AttendDate));
-- Partial-key indexes for the lookups that the UNIQUE composite indexes don't help with.
-- Enrollments: lookup by (StudentID, AcademicYear, Semester) drives the score grid.
-- MonthlyAssessments: RecomputeMidFromMonthly scans by EnrollID alone.
-- AttendanceRecords: roster joins look up by (StudentID, AcademicYear, Semester).
CREATE INDEX IF NOT EXISTS idx_enroll_sys ON Enrollments(StudentID, AcademicYear, Semester);
CREATE INDEX IF NOT EXISTS idx_ma_enroll  ON MonthlyAssessments(EnrollID);
CREATE INDEX IF NOT EXISTS idx_att_sys    ON AttendanceRecords(StudentID, AcademicYear, Semester);
CREATE TABLE IF NOT EXISTS GradeHistory(
    HistoryID INTEGER PRIMARY KEY AUTOINCREMENT,
    StudentID INTEGER NOT NULL REFERENCES Students(StudentID) ON DELETE CASCADE,
    FromGrade TEXT NOT NULL, ToGrade TEXT NOT NULL, AcademicYear TEXT NOT NULL,
    Note TEXT, ChangedBy TEXT, ChangedAt TEXT DEFAULT(datetime('now','localtime')));
CREATE TABLE IF NOT EXISTS ActivityLog(
    LogID INTEGER PRIMARY KEY AUTOINCREMENT, UserID INTEGER, Username TEXT,
    Action TEXT NOT NULL, Detail TEXT, LoggedAt TEXT DEFAULT(datetime('now','localtime')));
CREATE TABLE IF NOT EXISTS Settings(Key TEXT PRIMARY KEY, Value TEXT NOT NULL);
-- ── Memory / Conversation subsystem ──────────────────────────────
-- Conversations are top-level threads; Messages live inside a conversation.
-- Role: 'user' / 'assistant' / 'system' / 'summary'. 'summary' rows are
-- generated automatically when a conversation exceeds N messages — when
-- building optimised context we send {pinned facts + latest summary +
-- last few raw messages} instead of the entire transcript.
CREATE TABLE IF NOT EXISTS Conversations(
    ConvID         INTEGER PRIMARY KEY AUTOINCREMENT,
    Title          TEXT NOT NULL,
    StartedBy      TEXT,
    StartedAt      TEXT DEFAULT(datetime('now','localtime')),
    LastMessageAt  TEXT,
    MessageCount   INTEGER NOT NULL DEFAULT 0,
    TotalTokensEst INTEGER NOT NULL DEFAULT 0,
    IsArchived     INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS Messages(
    MsgID     INTEGER PRIMARY KEY AUTOINCREMENT,
    ConvID    INTEGER NOT NULL REFERENCES Conversations(ConvID) ON DELETE CASCADE,
    Role      TEXT NOT NULL,
    Content   TEXT NOT NULL,
    TokenEst  INTEGER NOT NULL DEFAULT 0,
    CreatedBy TEXT,
    CreatedAt TEXT DEFAULT(datetime('now','localtime')));
CREATE INDEX IF NOT EXISTS idx_msg_conv ON Messages(ConvID, MsgID);
-- Academic-year registry. One row per academic year the school has ever used.
-- IsCurrent flags the active year that drives DB.CurrentYear. All other tables
-- continue to store year as a TEXT column (Students.AcademicYear, etc.) — this
-- table is a catalogue, not an FK target.
CREATE TABLE IF NOT EXISTS AcademicYears(
    Year      TEXT PRIMARY KEY,
    StartDate TEXT,
    EndDate   TEXT,
    IsCurrent INTEGER NOT NULL DEFAULT 0,
    Note      TEXT,
    CreatedBy TEXT,
    CreatedAt TEXT DEFAULT(datetime('now','localtime')));

INSERT OR IGNORE INTO Settings VALUES('school_name','ໂຮງຮຽນ');
INSERT OR IGNORE INTO Settings VALUES('current_year','2025-2026');
INSERT OR IGNORE INTO Settings VALUES('current_semester','1');
INSERT OR IGNORE INTO Settings VALUES('mid_pct','50');
INSERT OR IGNORE INTO Settings VALUES('final_pct','50');
INSERT OR IGNORE INTO Settings VALUES('pass_score','5');

INSERT OR IGNORE INTO Users(Username,Password,FullName,Role) VALUES('admin','admin1234','ຜູ້ດູແລລະບົບ','admin');
INSERT OR IGNORE INTO Users(Username,Password,FullName,Role) VALUES('teacher1','teacher1234','ຄູສົມໃຈ ໃຈດີ','teacher');

-- Official MoES Lao lower-secondary subject list (ມ.1 – ມ.4 / Grade 7–10).
-- Same 13 subjects across all four grades, both semesters:
--   GradeLevel = '' → grade-agnostic (any of ມ.1 – ມ.4)
--   Semester   = 0  → semester-agnostic (any of Sem 1 / Sem 2)
-- Order matches the printed logbook layout used by the class-summary report.
INSERT OR IGNORE INTO Subjects(SubjectCode,SubjectName,GradeLevel,Semester,Category,SortOrder) VALUES
('CIV1', 'ສຶກສາພົນລະເມືອງ',         '',0,'ວິຊາສາມັນ',1),
('SCI1', 'ວິທະຍາສາດທຳມະຊາດ',       '',0,'ວິຊາສາມັນ',2),
('LAO1', 'ພາສາລາວ-ວັນນະຄະດີ',      '',0,'ວິຊາສາມັນ',3),
('MATH1','ຄະນິດສາດ',                '',0,'ວິຊາສາມັນ',4),
('ENG1', 'ພາສາອັງກິດ',              '',0,'ວິຊາສາມັນ',5),
('ICT1', 'ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ', '',0,'ວິຊາສາມັນ',6),
('GEO1', 'ພູມສາດ',                  '',0,'ວິຊາສາມັນ',7),
('HIS1', 'ປະຫວັດສາດ',               '',0,'ວິຊາສາມັນ',8),
('ART1', 'ສິລະປະກຳ',                '',0,'ວິຊາສາມັນ',9),
('MUS1', 'ສິລະປະດົນຕີ',             '',0,'ວິຊາສາມັນ',10),
('PE1',  'ພະລະສຶກສາ',               '',0,'ວິຊາສາມັນ',11),
('VOC1', 'ພື້ນຖານວິຊາຊີບ',          '',0,'ວິຊາສາມັນ',12),
('CHA1', 'ຄຸນສົມບັດ',                '',0,'ວິຊາສາມັນ',13),
('LAB1', 'ການອອກແຮງງານ',            '',0,'ວິຊາສາມັນ',14);
", conn);
            // Migration: add tables that might be missing in old DBs
            try { new SQLiteCommand("CREATE TABLE IF NOT EXISTS Announcements (AnnID INTEGER PRIMARY KEY AUTOINCREMENT, Title TEXT NOT NULL, Body TEXT, Priority TEXT DEFAULT 'normal', CreatedBy TEXT, CreatedAt TEXT DEFAULT (datetime('now','localtime')), ExpiresAt TEXT)", conn).ExecuteNonQuery(); } catch { }
            try { new SQLiteCommand("INSERT OR IGNORE INTO Announcements(Title,Body,Priority,CreatedBy) VALUES('ຍິນດີຕ້ອນຮັບ','ລະບົບຄຸ້ມຄອງຂໍ້ມູນນັກຮຽນ','high','admin')", conn).ExecuteNonQuery(); } catch { }

            // GradeHistory: add ClassRoom column so the Score-History page can show which
            // room a student was in each year (Students.ClassRoom only ever reflects the
            // CURRENT room). Backfill in two passes:
            //   1) For rows whose AcademicYear matches the student's current AcademicYear,
            //      copy Students.ClassRoom in — this is exact.
            //   2) For any remaining NULL/empty rows, copy Students.ClassRoom anyway as a
            //      best-effort guess. This matters for GRADUATED students (whose last
            //      GradeHistory row has AcademicYear=year-after-grad, never matching their
            //      Students.AcademicYear=year-finished) and for multi-step promoted students
            //      (whose oldest rows pre-date the column). Without pass 2, graduates would
            //      be INVISIBLE in the score-history roster because their final GradeHistory
            //      row's NULL ClassRoom never matches the room filter in
            //      GetHistoricalClassRoster. Trade-off: a student who changed rooms across
            //      years gets the wrong room on older rows — but their grade (FromGrade) is
            //      correct, and from this point forward PromotionPage records the actual
            //      room per promotion, so this only affects legacy data.
            try { new SQLiteCommand("ALTER TABLE GradeHistory ADD COLUMN ClassRoom TEXT", conn).ExecuteNonQuery(); } catch { }
            try { new SQLiteCommand(@"
                UPDATE GradeHistory
                SET ClassRoom = (SELECT ClassRoom FROM Students s WHERE s.StudentID = GradeHistory.StudentID)
                WHERE (ClassRoom IS NULL OR ClassRoom='')
                  AND AcademicYear = (SELECT AcademicYear FROM Students s WHERE s.StudentID = GradeHistory.StudentID)",
                conn).ExecuteNonQuery(); } catch { }
            try { new SQLiteCommand(@"
                UPDATE GradeHistory
                SET ClassRoom = (SELECT ClassRoom FROM Students s WHERE s.StudentID = GradeHistory.StudentID)
                WHERE (ClassRoom IS NULL OR ClassRoom='')
                  AND EXISTS (SELECT 1 FROM Students s
                              WHERE s.StudentID = GradeHistory.StudentID
                                AND s.ClassRoom IS NOT NULL AND s.ClassRoom <> '')",
                conn).ExecuteNonQuery(); } catch { }

            // EvaluationScores: per-student, per-context manual scores for ຄຸນສົມບັດ (CHA1)
            // and ການອອກແຮງງານ (LAB1). Context is one of 'SEM1' / 'SEM2' / 'ANNUAL' — these
            // values appear on the corresponding summary reports INSTEAD of any auto-derived
            // monthly aggregate. The teacher enters them manually on MonthlyScoresPage by
            // picking the matching ປະເພດການປະເມີນ filter.
            try { new SQLiteCommand(@"
                CREATE TABLE IF NOT EXISTS EvaluationScores(
                    EvalID INTEGER PRIMARY KEY AUTOINCREMENT,
                    StudentID INTEGER NOT NULL REFERENCES Students(StudentID) ON DELETE CASCADE,
                    AcademicYear TEXT NOT NULL,
                    Context TEXT NOT NULL,        -- 'SEM1' | 'SEM2' | 'ANNUAL'
                    SubjectCode TEXT NOT NULL,    -- 'CHA1' | 'LAB1'
                    Score REAL,
                    UpdatedAt TEXT DEFAULT(datetime('now','localtime')),
                    UNIQUE(StudentID, AcademicYear, Context, SubjectCode));
                CREATE INDEX IF NOT EXISTS idx_eval_lookup
                    ON EvaluationScores(AcademicYear, Context, SubjectCode);",
                conn).ExecuteNonQuery(); } catch { }

            // Migration: add new Students columns for full identity / parent details.
            // Each ALTER TABLE is wrapped: SQLite errors if the column already exists.
            string[] addCols = {
                "BirthVillage TEXT", "BirthDistrict TEXT", "BirthProvince TEXT",
                "FatherName TEXT", "FatherAge INTEGER", "FatherJob TEXT",
                "FatherVillage TEXT", "FatherDistrict TEXT", "FatherProvince TEXT", "FatherPhone TEXT",
                "MotherName TEXT", "MotherAge INTEGER", "MotherJob TEXT",
                "MotherVillage TEXT", "MotherDistrict TEXT", "MotherProvince TEXT", "MotherPhone TEXT"
            };
            foreach (var col in addCols)
                try { new SQLiteCommand($"ALTER TABLE Students ADD COLUMN {col}", conn).ExecuteNonQuery(); } catch { }

            // One-time backfill: copy legacy ParentName/ParentPhone into FatherName/FatherPhone
            // when the new columns are still empty, so old records remain useful.
            try { new SQLiteCommand("UPDATE Students SET FatherName=ParentName WHERE (FatherName IS NULL OR FatherName='') AND ParentName IS NOT NULL AND ParentName<>''", conn).ExecuteNonQuery(); } catch { }
            try { new SQLiteCommand("UPDATE Students SET FatherPhone=ParentPhone WHERE (FatherPhone IS NULL OR FatherPhone='') AND ParentPhone IS NOT NULL AND ParentPhone<>''", conn).ExecuteNonQuery(); } catch { }

            // Migrate previous 60/40 default to the new 50/50 default. Custom values are preserved.
            try { new SQLiteCommand("UPDATE Settings SET Value='50' WHERE Key='mid_pct'   AND Value='60'", conn).ExecuteNonQuery(); } catch { }
            try { new SQLiteCommand("UPDATE Settings SET Value='50' WHERE Key='final_pct' AND Value='40'", conn).ExecuteNonQuery(); } catch { }

            // ── AcademicYears catalogue backfill ──
            // Bring every distinct year that ever appeared anywhere into the registry,
            // then mark Settings.current_year as IsCurrent=1 (and every other row =0).
            try
            {
                new SQLiteCommand(
                    @"INSERT OR IGNORE INTO AcademicYears(Year, IsCurrent)
                      SELECT DISTINCT AcademicYear, 0 FROM Students
                      WHERE AcademicYear IS NOT NULL AND AcademicYear<>''", conn).ExecuteNonQuery();
                new SQLiteCommand(
                    @"INSERT OR IGNORE INTO AcademicYears(Year, IsCurrent)
                      SELECT DISTINCT AcademicYear, 0 FROM Enrollments
                      WHERE AcademicYear IS NOT NULL AND AcademicYear<>''", conn).ExecuteNonQuery();
                new SQLiteCommand(
                    @"INSERT OR IGNORE INTO AcademicYears(Year, IsCurrent)
                      SELECT DISTINCT AcademicYear, 0 FROM GradeHistory
                      WHERE AcademicYear IS NOT NULL AND AcademicYear<>''", conn).ExecuteNonQuery();

                // Ensure the year referenced by Settings.current_year exists in the catalogue
                using (var ins = new SQLiteCommand(
                    "INSERT OR IGNORE INTO AcademicYears(Year, IsCurrent) VALUES(@y, 1)", conn))
                {
                    ins.Parameters.AddWithValue("@y", CurrentYear);
                    ins.ExecuteNonQuery();
                }
                // Make sure exactly one row has IsCurrent=1 — the one matching Settings.current_year
                using (var upd = new SQLiteCommand(
                    "UPDATE AcademicYears SET IsCurrent = CASE WHEN Year=@y THEN 1 ELSE 0 END", conn))
                {
                    upd.Parameters.AddWithValue("@y", CurrentYear);
                    upd.ExecuteNonQuery();
                }
            } catch { }

            // ── Migration: align Subjects with the official MoES logbook list ──
            // Renames legacy seed names to canonical forms — but only when the row
            // still holds the original seeded value, so user-customised names are
            // left untouched.
            string[][] renames = {
                new[]{ "HIS1", "ປະຫວັດສາດ-2",       "ປະຫວັດສາດ"                  },
                new[]{ "GEO1", "ພູມສາດວິໄຈ",        "ພູມສາດ"                      },
                new[]{ "ENG1", "ພາສາຕ່າງປະເທດ",     "ພາສາອັງກິດ"                  },
                new[]{ "ART1", "ສິລະປະ",            "ສິລະປະກຳ"                    },
                new[]{ "SCI1", "ວິທະຍາສາດ",         "ວິທະຍາສາດທຳມະຊາດ"            },
            };
            foreach (var r in renames)
            {
                try
                {
                    using var u = new SQLiteCommand(
                        "UPDATE Subjects SET SubjectName=@n WHERE SubjectCode=@c AND SubjectName=@o", conn);
                    u.Parameters.AddWithValue("@c", r[0]);
                    u.Parameters.AddWithValue("@o", r[1]);
                    u.Parameters.AddWithValue("@n", r[2]);
                    u.ExecuteNonQuery();
                } catch { }
            }

            // Insert any subjects from the official MoES list that don't already exist
            // (covers databases created before the canonical list was added).
            // Order matches the printed logbook layout for ມ.1 – ມ.4 (Grade 7–10).
            (string code, string name, int sort)[] officialSubjects = {
                ("CIV1",  "ສຶກສາພົນລະເມືອງ",          1),
                ("SCI1",  "ວິທະຍາສາດທຳມະຊາດ",       2),
                ("LAO1",  "ພາສາລາວ-ວັນນະຄະດີ",       3),
                ("MATH1", "ຄະນິດສາດ",                 4),
                ("ENG1",  "ພາສາອັງກິດ",                5),
                ("ICT1",  "ເຕັກໂນໂລຊີຂໍ້ມູນຂ່າວສານ",  6),
                ("GEO1",  "ພູມສາດ",                    7),
                ("HIS1",  "ປະຫວັດສາດ",                8),
                ("ART1",  "ສິລະປະກຳ",                  9),
                ("MUS1",  "ສິລະປະດົນຕີ",              10),
                ("PE1",   "ພະລະສຶກສາ",                11),
                ("VOC1",  "ພື້ນຖານວິຊາຊີບ",            12),
                ("CHA1",  "ຄຸນສົມບັດ",                13),
                ("LAB1",  "ການອອກແຮງງານ",             14),
            };
            foreach (var sub in officialSubjects)
            {
                try
                {
                    using var ins = new SQLiteCommand(
                        "INSERT OR IGNORE INTO Subjects(SubjectCode,SubjectName,GradeLevel,Semester,Category,SortOrder) " +
                        "VALUES(@c,@n,'',0,'ວິຊາສາມັນ',@so)", conn);
                    ins.Parameters.AddWithValue("@c", sub.code);
                    ins.Parameters.AddWithValue("@n", sub.name);
                    ins.Parameters.AddWithValue("@so", sub.sort);
                    ins.ExecuteNonQuery();
                } catch { }
            }

            // Realign SortOrder + clear grade/semester scoping for the official codes so they:
            //   • always display in the canonical logbook order, and
            //   • appear in every grade/semester filter (since the same 13 subjects are
            //     taught in ມ.1 – ມ.4 across both semesters).
            // User-customised codes outside this list are untouched.
            foreach (var sub in officialSubjects)
            {
                try
                {
                    using var u = new SQLiteCommand(
                        "UPDATE Subjects SET SortOrder=@so, GradeLevel='', Semester=0 WHERE SubjectCode=@c", conn);
                    u.Parameters.AddWithValue("@c",  sub.code);
                    u.Parameters.AddWithValue("@so", sub.sort);
                    u.ExecuteNonQuery();
                } catch { }
            }

            // VOC1 (ພື້ນຖານວິຊາຊີບ) backfill — added as the 12th academic subject in
            // this version. Existing students never had VOC1 enrolments. For every
            // distinct (year, semester) where any enrolment exists, auto-enrol every
            // active student in VOC1 so the score report has a row to fill.
            try
            {
                new SQLiteCommand(@"
                    INSERT OR IGNORE INTO Enrollments(StudentID, SubjectID, AcademicYear, Semester)
                    SELECT DISTINCT e.StudentID,
                           (SELECT SubjectID FROM Subjects WHERE SubjectCode='VOC1'),
                           e.AcademicYear, e.Semester
                    FROM Enrollments e
                    WHERE EXISTS (SELECT 1 FROM Subjects WHERE SubjectCode='VOC1')
                      AND NOT EXISTS (
                          SELECT 1 FROM Enrollments e2
                          WHERE e2.StudentID = e.StudentID
                            AND e2.AcademicYear = e.AcademicYear
                            AND e2.Semester = e.Semester
                            AND e2.SubjectID = (SELECT SubjectID FROM Subjects WHERE SubjectCode='VOC1'))",
                    conn).ExecuteNonQuery();
            } catch { }

            // Clean up subjects from earlier versions that aren't in the official list:
            // - DES1 (Design)        — never an official MoES subject; remove if unused.
            // - CAR1 (Career)        — not in this school's curriculum; remove if unused.
            // Only deletes when the subject has zero enrollments, so user data is safe.
            foreach (string oldCode in new[] { "DES1", "CAR1" })
            {
                try
                {
                    int enrolls = ScalarInt(
                        "SELECT COUNT(*) FROM Enrollments WHERE SubjectID IN (SELECT SubjectID FROM Subjects WHERE SubjectCode=@c)",
                        conn, ("@c", oldCode));
                    if (enrolls == 0)
                    {
                        using var del = new SQLiteCommand("DELETE FROM Subjects WHERE SubjectCode=@c", conn);
                        del.Parameters.AddWithValue("@c", oldCode);
                        del.ExecuteNonQuery();
                    }
                } catch { }
            }

            LoadSettings(conn);
        }

        public static void LoadSettings(SQLiteConnection? ext = null)
        {
            bool owns = ext == null; var c = ext ?? Open();
            try
            {
                string G(string k, string d) { try { return Scalar("SELECT Value FROM Settings WHERE Key=@k", c, ("@k", k))?.ToString()??d; } catch { return d; } }
                int    GI(string k, int    d) => int.TryParse(G(k, d.ToString(CultureInfo.InvariantCulture)), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : d;
                double GD(string k, double d) => double.TryParse(G(k, d.ToString(CultureInfo.InvariantCulture)), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : d;
                SchoolName  = G("school_name","ໂຮງຮຽນ");
                CurrentYear = G("current_year","2025-2026");
                CurrentSem  = GI("current_semester", 1);
                // Default ratio is 50/50; the migration block flips legacy 60/40 if present.
                MidPct      = GD("mid_pct",   50);
                FinalPct    = GD("final_pct", 50);
                PassScore   = GD("pass_score", 5);
            }
            finally { if (owns) c.Dispose(); }
        }

        public static void SaveSetting(string key, string value)
        { Exec("INSERT OR REPLACE INTO Settings(Key,Value) VALUES(@k,@v)",null,("@k",key),("@v",value)); LoadSettings(); }

        // Read one Settings value; null when the key doesn't exist.
        public static string? GetSetting(string key) =>
            Scalar("SELECT Value FROM Settings WHERE Key=@k", null, ("@k", key))?.ToString();

        public static SQLiteConnection Open()
        {
            var c = new SQLiteConnection(ConnectionString); c.Open();
            new SQLiteCommand("PRAGMA foreign_keys=ON;",c).ExecuteNonQuery();
            return c;
        }

        public static (bool ok, string role, string fullName) Login(string user, string pass)
        {
            using var c = Open();
            using var cmd = new SQLiteCommand("SELECT UserID,Role,FullName FROM Users WHERE Username=@u AND Password=@p AND IsActive=1",c);
            cmd.Parameters.AddWithValue("@u",user); cmd.Parameters.AddWithValue("@p",pass);
            using var r = cmd.ExecuteReader();
            if (!r.Read()) return (false,"","");
            CurrentUserId = Convert.ToInt32(r["UserID"]);
            CurrentRole   = r["Role"].ToString()!;
            CurrentUser   = r["FullName"].ToString()!;
            r.Close();
            Exec($"UPDATE Users SET LastLogin=datetime('now','localtime') WHERE UserID={CurrentUserId}",c);
            Log("Login","ເຂົ້າໃຊ້ງານ",c);
            return (true,CurrentRole,CurrentUser);
        }

        // ── Lao scoring: out of 10, no credits/GPA ─────────────────
        // Final semester score = MonthlyAvg × MidPct% + FinalExam × FinalPct% (default 50:50).
        public static double CalcTotal(double mid, double fin)
            => Math.Round(mid*(MidPct/100.0) + fin*(FinalPct/100.0), 2);

        public static string CalcLevel(double total)
        {
            if (total >= 8.0)       return "ດີຫຼາຍ";
            if (total >= 6.0)       return "ດີ";
            if (total >= PassScore) return "ຜ່ານ";
            return "ບໍ່ຜ່ານ";
        }

        // ── MoES 5-band scale ──────────────────────────────────────
        // Used in official Lao secondary report cards (ໃບປະເມີນຜົນ).
        // 9.0–10.0 ດີຫຼາຍ  ·  7.0–8.99 ດີ  ·  5.0–6.99 ປານກາງ  ·  3.0–4.99 ອ່ອນ  ·  <3.0 ຕົກ
        public static string CalcMoESLevel(double total)
        {
            if (total >= 9.0) return "ດີຫຼາຍ";
            if (total >= 7.0) return "ດີ";
            if (total >= 5.0) return "ປານກາງ";
            if (total >= 3.0) return "ອ່ອນ";
            return "ຕົກ";
        }

        // ── Subject grouping for grouped report cards ──────────────
        // Maps a subject code/name to one of three categories matching the
        // logbook layout: Languages / Sciences-Math / Social Studies.
        public enum SubjectGroupId { Language, Science, Social }
        public static SubjectGroupId SubjectGroup(string code, string name)
        {
            string c = (code ?? "").ToUpperInvariant();
            string n = name ?? "";
            // Languages: anything with the word ພາສາ in the Lao name, or codes LAO/ENG/FRE/CHI/VIE
            if (n.Contains("ພາສາ") || n.Contains("ວັນນະຄະດີ")
                || c.StartsWith("LAO") || c.StartsWith("ENG") || c.StartsWith("FRE")
                || c.StartsWith("CHI") || c.StartsWith("VIE"))
                return SubjectGroupId.Language;
            // Sciences/Math/Tech: math, physics, chemistry, biology, natural science, ICT
            if (c.StartsWith("MATH") || c.StartsWith("PHY") || c.StartsWith("CHEM")
                || c.StartsWith("BIO")  || c.StartsWith("SCI") || c.StartsWith("ICT")
                || n.Contains("ຄະນິດ")       || n.Contains("ຟິສິກ") || n.Contains("ເຄມີ")
                || n.Contains("ຊີວະ")        || n.Contains("ວິທະຍາສາດ")
                || n.Contains("ເຕັກໂນໂລຊີ"))
                return SubjectGroupId.Science;
            // Default: Social Studies (civic, history, geography, PE, arts, music,
            // career guidance, design, religion, etc.)
            return SubjectGroupId.Social;
        }

        public static string SubjectGroupLabel(SubjectGroupId g) => g switch
        {
            SubjectGroupId.Language => "ກຸ່ມພາສາສາດ",
            SubjectGroupId.Science  => "ກຸ່ມວິທະຍາສາດ-ຄະນິດສາດ",
            _                       => "ກຸ່ມສັງຄົມ"
        };

        // ── Default conduct grade based on academic performance ──────
        // Schools traditionally hand-write this; this is a sensible auto-fill the
        // teacher can override. Uses Lao 4-band wording matching CalcLevel().
        public static string DefaultConduct(double avg, int failedSubjects, int unexcusedAbsences)
        {
            if (failedSubjects == 0 && unexcusedAbsences == 0 && avg >= 8.0) return "ດີຫຼາຍ";
            if (failedSubjects <= 1 && unexcusedAbsences <= 2 && avg >= 6.0) return "ດີ";
            if (failedSubjects <= 2 && unexcusedAbsences <= 5)               return "ປານກາງ";
            return "ຕ້ອງປັບປຸງ";
        }

        // ── Monthly continuous assessment helpers ──────────────────
        // Activity(/3) + Discipline(/2) + Homework(/5) = monthly score (/10).
        public static double CalcMonthlyTotal(double activity, double discipline, double homework)
            => Math.Round(activity + discipline + homework, 2);

        // Lao academic year runs Sept–June with two grading periods:
        //   Sem 1 monthly assessment:  Sept · Oct · Nov · Dec   (4 months)
        //   Sem 1 final exam:          January                  (1 month)
        //   Sem 2 monthly assessment:  Feb  · Mar · Apr · May   (4 months)
        //   Sem 2 final exam:          June                     (1 month)
        // MonthsInSemester returns ONLY the monthly-assessment months — the final
        // exam month is excluded so the monthly average isn't diluted by it.
        public static int[] MonthsInSemester(int semester)
            => semester == 1 ? new[]{9,10,11,12} : new[]{2,3,4,5};

        // Inverse mapping: which semester does a calendar month belong to?
        // Sept–Jan → Sem 1, Feb–June → Sem 2. Final-exam months map to their own semester.
        public static int SemesterForMonth(int month)
            => (month >= 9 || month == 1) ? 1 : 2;

        // The month when the semester final exam is administered.
        public static int FinalExamMonth(int semester) => semester == 1 ? 1 : 6;

        // ── EvaluationScores helpers ──────────────────────────────────────
        // Manual CHA1/LAB1 scores per (Student, Year, Context, SubjectCode).
        // Context strings used:
        //   "Month1".."Month8" — one per monthly assessment window (Sept = Month1,
        //                        Oct = Month2, …, May = Month8). Stored manually,
        //                        NEVER derived from MonthlyAssessments sub-scores.
        //   "SEM1" / "SEM2"    — semester summary scores (manual).
        //   "ANNUAL"           — annual summary score (manual).

        /// <summary>Map a calendar month (9-12 / 2-5) to the EvaluationScores Context
        /// string used by the CHA1/LAB1 monthly entry workflow. Sept = "Month1",
        /// Oct = "Month2", …, May = "Month8". Returns null for the final-exam months
        /// (January / June) and any other invalid month.</summary>
        public static string? MonthContextName(int calendarMonth) => calendarMonth switch
        {
            9  => "Month1",  10 => "Month2", 11 => "Month3", 12 => "Month4",
            2  => "Month5",   3 => "Month6",  4 => "Month7",  5 => "Month8",
            _  => null
        };

        public static double? GetEvaluationScore(int studentId, string year, string context, string subjectCode)
        {
            var o = Scalar(@"SELECT Score FROM EvaluationScores
                             WHERE StudentID=@s AND AcademicYear=@y AND Context=@c AND SubjectCode=@sc",
                null, ("@s", studentId), ("@y", year), ("@c", context), ("@sc", subjectCode));
            return (o == null || o == DBNull.Value) ? null : Convert.ToDouble(o);
        }

        public static void SetEvaluationScore(int studentId, string year, string context, string subjectCode, double? score, SQLiteConnection c, SQLiteTransaction tx)
        {
            // null score → delete the row so it doesn't pollute reports as a 0.
            if (score == null)
            {
                ExecTx("DELETE FROM EvaluationScores WHERE StudentID=@s AND AcademicYear=@y AND Context=@c AND SubjectCode=@sc",
                    c, tx,
                    ("@s", studentId), ("@y", year), ("@c", context), ("@sc", subjectCode));
                return;
            }
            ExecTx(@"INSERT INTO EvaluationScores(StudentID,AcademicYear,Context,SubjectCode,Score,UpdatedAt)
                    VALUES(@s,@y,@c,@sc,@sv,datetime('now','localtime'))
                    ON CONFLICT(StudentID,AcademicYear,Context,SubjectCode) DO UPDATE SET
                      Score=excluded.Score, UpdatedAt=datetime('now','localtime')",
                c, tx,
                ("@s", studentId), ("@y", year), ("@c", context), ("@sc", subjectCode), ("@sv", score.Value));
        }

        // Recompute Scores.MidScore for an enrollment as the average of its monthly totals,
        // then refresh TotalScore + Level using the current FinalScore and the configured ratio.
        // Creates the Scores row on the fly if it doesn't exist yet.
        public static void RecomputeMidFromMonthly(int enrollId)
        {
            using var c = Open();
            RecomputeMidFromMonthly(enrollId, c);
        }

        // Connection-reusing overload — batch callers (MonthlyScoresPage.SaveAll) should
        // open one connection and pass it in to avoid the per-row connection churn that
        // dominated the previous implementation.
        public static void RecomputeMidFromMonthly(int enrollId, SQLiteConnection c)
        {
            // ຄຸນສົມບັດ (CHA1) and ແຮງງານ (LAB1) are evaluation-only — the school's
            // rule is that they MUST NOT be averaged, weighted, or computed in any
            // way. Their monthly entries are the only stored values; nothing is
            // derived. So skip recomputing Scores.MidScore / TotalScore for them.
            var code = Scalar(@"
                SELECT sub.SubjectCode FROM Enrollments e
                JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                WHERE e.EnrollID=@e", c, ("@e", enrollId))?.ToString();
            if (code == "CHA1" || code == "LAB1") return;

            var avgObj = Scalar(@"SELECT AVG(IFNULL(ActivityScore,0)+IFNULL(DisciplineScore,0)+IFNULL(HomeworkScore,0))
                                  FROM MonthlyAssessments WHERE EnrollID=@e", c, ("@e", enrollId));
            double monthlyAvg = (avgObj == null || avgObj == DBNull.Value) ? 0 : Math.Round(Convert.ToDouble(avgObj), 2);

            // Ensure a Scores row exists.
            Exec("INSERT OR IGNORE INTO Scores(EnrollID,MidScore,FinalScore,TotalScore,Level) VALUES(@e,0,0,0,'')",
                 c, ("@e", enrollId));

            var finObj = Scalar("SELECT FinalScore FROM Scores WHERE EnrollID=@e", c, ("@e", enrollId));
            double finalScore = (finObj == null || finObj == DBNull.Value) ? 0 : Convert.ToDouble(finObj);
            double total = CalcTotal(monthlyAvg, finalScore);
            string level = CalcLevel(total);

            Exec(@"UPDATE Scores SET MidScore=@m, TotalScore=@t, Level=@l,
                   UpdatedAt=datetime('now','localtime') WHERE EnrollID=@e",
                 c, ("@m", monthlyAvg), ("@t", total), ("@l", level), ("@e", enrollId));
        }

        public static DataTable GetClassRanking(string grade, string room, string year, int sem)
        {
            using var c = Open();
            // Parameterised — any apostrophe-bearing input (room labels, future schemas)
            // is safe. PassScore is a static float and injected via parameter too.
            // ຄຸນສົມບັດ (CHA1) and ແຮງງານ (LAB1) are EXCLUDED from the ranking aggregates
            // because they are evaluation-only — never averaged, never used in totals.
            return Query(@"
                SELECT s.StudentID,
                       s.StudentCode                                   AS ລະຫັດ,
                       s.FirstName||' '||s.LastName                    AS ຊື່ນັກຮຽນ,
                       s.ClassRoom                                     AS ຫ້ອງ,
                       COUNT(sc.ScoreID)                               AS ຈຳນວນວິຊາ,
                       ROUND(AVG(IFNULL(sc.TotalScore,0)),2)           AS ສະເລ່ຍ,
                       ROUND(SUM(IFNULL(sc.TotalScore,0)),2)           AS ລວມ,
                       SUM(CASE WHEN sc.TotalScore < @pass AND sc.ScoreID IS NOT NULL THEN 1 ELSE 0 END) AS ຕົກ
                FROM Students s
                JOIN  Enrollments e ON e.StudentID=s.StudentID AND e.AcademicYear=@yr AND e.Semester=@sm
                JOIN  Subjects sub ON sub.SubjectID=e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                WHERE s.GradeLevel=@g AND s.ClassRoom=@r AND s.Status='ກຳລັງຮຽນ'
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                GROUP BY s.StudentID
                ORDER BY ສະເລ່ຍ DESC, ລວມ DESC",
                c,
                ("@pass", PassScore), ("@yr", year), ("@sm", sem),
                ("@g", grade), ("@r", room));
        }

        // ════════════════════════════════════════════════════════════
        //  STUDENT SCORE HISTORY
        // ════════════════════════════════════════════════════════════
        //  Read-only queries that pivot on Enrollments.AcademicYear (the
        //  HISTORICAL year), NEVER on Students.GradeLevel/ClassRoom (which
        //  always reflect the student's CURRENT state). This lets graduated
        //  and promoted students view every year they ever attended.
        //
        //  GradeHistory semantics (existing PromotionPage):
        //    AcademicYear = the year the student is being PROMOTED INTO.
        //    FromGrade    = the grade they JUST COMPLETED (= grade DURING the
        //                   year immediately before AcademicYear).
        //    ClassRoom    = the room they were in WHILE in FromGrade.
        //  So to recover (grade, room) FOR year Y, take the GradeHistory row
        //  with the smallest AcademicYear strictly greater than Y — its
        //  FromGrade + ClassRoom describe year Y. If no such row exists,
        //  Y is the student's current year and Students.GradeLevel /
        //  Students.ClassRoom apply.
        //
        //  All score aggregates (averages, totals, rank) EXCLUDE CHA1/LAB1
        //  per the school's evaluation-only contract.

        /// <summary>Every (year, grade, room) the student has data for, newest first.</summary>
        public static DataTable GetStudentHistoryYears(int studentId)
        {
            // Union of every year referenced by Enrollments + GradeHistory + the
            // student's current AcademicYear, with a historical (grade, room) attached.
            return Query(@"
                WITH years AS (
                    SELECT DISTINCT AcademicYear AS Yr FROM Enrollments WHERE StudentID=@s
                    UNION
                    SELECT AcademicYear AS Yr FROM Students WHERE StudentID=@s
                )
                SELECT y.Yr AS AcademicYear,
                       IFNULL(
                           (SELECT FromGrade FROM GradeHistory
                            WHERE StudentID=@s AND AcademicYear > y.Yr
                            ORDER BY AcademicYear ASC LIMIT 1),
                           (SELECT GradeLevel FROM Students WHERE StudentID=@s)
                       ) AS GradeLevel,
                       IFNULL(
                           (SELECT ClassRoom FROM GradeHistory
                            WHERE StudentID=@s AND AcademicYear > y.Yr
                            ORDER BY AcademicYear ASC LIMIT 1),
                           (SELECT ClassRoom FROM Students WHERE StudentID=@s)
                       ) AS ClassRoom
                FROM years y
                WHERE y.Yr IS NOT NULL AND y.Yr <> ''
                ORDER BY y.Yr DESC",
                null, ("@s", studentId));
        }

        /// <summary>Historical monthly continuous-assessment scores for a (student, year).
        /// One row per (subject × month). Academic subjects come from
        /// MonthlyAssessments (Activity/Discipline/Homework sub-scores).
        /// CHA1/LAB1 come from EvaluationScores(Month1..Month8) — they don't
        /// live in MonthlyAssessments so the query UNIONs them in with NULL
        /// sub-scores + the eval score in ລວມເດືອນ. Without the UNION they
        /// simply vanished from the per-month display in ProfileHistoryPage.
        /// CHA1/LAB1 rows are ordered LAST within each month for visual clarity.</summary>
        public static DataTable GetHistoryMonthly(int studentId, string year)
        {
            return Query(@"
                SELECT * FROM (
                    SELECT sub.SubjectCode                         AS ລະຫັດວິຊາ,
                           sub.SubjectName                         AS ຊື່ວິຊາ,
                           e.Semester                              AS ພາກ,
                           m.Month                                 AS ເດືອນ,
                           ROUND(IFNULL(m.ActivityScore,0),2)     AS ຮ່ວມຮຽນ,
                           ROUND(IFNULL(m.DisciplineScore,0),2)   AS ກິດຈະກຳ,
                           ROUND(IFNULL(m.HomeworkScore,0),2)     AS ກວດກາ,
                           ROUND(IFNULL(m.ActivityScore,0)+IFNULL(m.DisciplineScore,0)+IFNULL(m.HomeworkScore,0),2) AS ລວມເດືອນ
                    FROM Enrollments e
                    JOIN Subjects sub          ON sub.SubjectID = e.SubjectID
                    JOIN MonthlyAssessments m  ON m.EnrollID    = e.EnrollID
                    WHERE e.StudentID=@s AND e.AcademicYear=@y
                      AND sub.SubjectCode NOT IN ('CHA1','LAB1')
                    UNION ALL
                    SELECT sub.SubjectCode  AS ລະຫັດວິຊາ,
                           sub.SubjectName  AS ຊື່ວິຊາ,
                           CASE ev.Context
                                WHEN 'Month1' THEN 1 WHEN 'Month2' THEN 1
                                WHEN 'Month3' THEN 1 WHEN 'Month4' THEN 1
                                WHEN 'Month5' THEN 2 WHEN 'Month6' THEN 2
                                WHEN 'Month7' THEN 2 WHEN 'Month8' THEN 2
                           END              AS ພາກ,
                           CASE ev.Context
                                WHEN 'Month1' THEN 9  WHEN 'Month2' THEN 10
                                WHEN 'Month3' THEN 11 WHEN 'Month4' THEN 12
                                WHEN 'Month5' THEN 2  WHEN 'Month6' THEN 3
                                WHEN 'Month7' THEN 4  WHEN 'Month8' THEN 5
                           END              AS ເດືອນ,
                           NULL             AS ຮ່ວມຮຽນ,
                           NULL             AS ກິດຈະກຳ,
                           NULL             AS ກວດກາ,
                           ROUND(ev.Score,2) AS ລວມເດືອນ
                    FROM EvaluationScores ev
                    JOIN Subjects sub ON sub.SubjectCode = ev.SubjectCode
                    WHERE ev.StudentID=@s AND ev.AcademicYear=@y
                      AND ev.Context IN ('Month1','Month2','Month3','Month4',
                                         'Month5','Month6','Month7','Month8')
                      AND ev.SubjectCode IN ('CHA1','LAB1')
                )
                ORDER BY ພາກ, ເດືອນ,
                         CASE WHEN ລະຫັດວິຊາ IN ('CHA1','LAB1') THEN 1 ELSE 0 END,
                         ລະຫັດວິຊາ",
                null, ("@s", studentId), ("@y", year));
        }

        /// <summary>Per-subject semester scorecard for a (student, year, sem).
        /// MidScore = averaged monthly; FinalScore = exam; TotalScore by weighted formula.
        /// CHA1/LAB1 rows carry the manually-entered EvaluationScores.Score in ‘ຄະແນນປະເມີນ’
        /// instead of mid/final — they don't contribute to ‘ສະເລ່ຍ’ further down.</summary>
        public static DataTable GetHistorySemester(int studentId, string year, int sem)
        {
            string ctx = sem == 1 ? "SEM1" : "SEM2";
            return Query(@"
                SELECT sub.SubjectCode                                            AS ລະຫັດວິຊາ,
                       sub.SubjectName                                            AS ຊື່ວິຊາ,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1')
                            THEN (SELECT ROUND(ev.Score,2) FROM EvaluationScores ev
                                  WHERE ev.StudentID=@s AND ev.AcademicYear=@y
                                    AND ev.Context=@ctx AND ev.SubjectCode=sub.SubjectCode)
                            ELSE ROUND(IFNULL(sc.MidScore,0),2)                 END AS ສະເລ່ຍເດືອນ,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1') THEN NULL
                            ELSE ROUND(IFNULL(sc.FinalScore,0),2)               END AS ຄະແນນເສັງ,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1')
                            THEN (SELECT ROUND(ev.Score,2) FROM EvaluationScores ev
                                  WHERE ev.StudentID=@s AND ev.AcademicYear=@y
                                    AND ev.Context=@ctx AND ev.SubjectCode=sub.SubjectCode)
                            ELSE ROUND(IFNULL(sc.TotalScore,0),2)               END AS ລວມພາກ,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1') THEN ''
                            ELSE IFNULL(sc.Level,'')                            END AS ລະດັບ
                FROM Enrollments e
                JOIN Subjects sub      ON sub.SubjectID = e.SubjectID
                LEFT JOIN Scores sc    ON sc.EnrollID   = e.EnrollID
                WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                ORDER BY sub.SortOrder, sub.SubjectCode",
                null, ("@s", studentId), ("@y", year), ("@sm", sem), ("@ctx", ctx));
        }

        /// <summary>One-row summary for a (student, year, sem): subject count + average
        /// + total + class rank. Rank is computed against every student whose
        /// Enrollments place them in the same (grade, year, sem) — read from
        /// GradeHistory.FromGrade (preferred) or Students.GradeLevel (fallback for the
        /// current year). CHA1/LAB1 are excluded from every aggregate.</summary>
        public static (int subjects, double avg, double total, int rank, int classSize, bool failed)
            GetHistorySemesterSummary(int studentId, string year, int sem)
        {
            var ag = Query(@"
                SELECT COUNT(sc.ScoreID)                          AS N,
                       ROUND(AVG(IFNULL(sc.TotalScore,0)),2)      AS Avg,
                       ROUND(SUM(IFNULL(sc.TotalScore,0)),2)      AS Sum,
                       SUM(CASE WHEN sc.TotalScore < @pass AND sc.ScoreID IS NOT NULL THEN 1 ELSE 0 END) AS Fails
                FROM Enrollments e
                JOIN Subjects sub      ON sub.SubjectID = e.SubjectID
                LEFT JOIN Scores sc    ON sc.EnrollID   = e.EnrollID
                WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                  AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                null, ("@s", studentId), ("@y", year), ("@sm", sem), ("@pass", PassScore));
            if (ag.Rows.Count == 0)
                return (0, 0, 0, 0, 0, false);
            int    n     = Convert.ToInt32  (ag.Rows[0]["N"]);
            double avg   = ag.Rows[0]["Avg"] == DBNull.Value ? 0 : Convert.ToDouble(ag.Rows[0]["Avg"]);
            double sum   = ag.Rows[0]["Sum"] == DBNull.Value ? 0 : Convert.ToDouble(ag.Rows[0]["Sum"]);
            int    fails = ag.Rows[0]["Fails"] == DBNull.Value ? 0 : Convert.ToInt32(ag.Rows[0]["Fails"]);

            // Reconstruct grade-at-the-time. GradeHistory rows record the year being
            // promoted INTO — so the row that DESCRIBES year Y is the one with the
            // smallest AcademicYear strictly greater than Y, and its FromGrade is the
            // grade the student was in during Y. Fall back to current Students for the
            // most recent year (no later promotion exists yet).
            string? gradeAtTime = Scalar(
                @"SELECT FromGrade FROM GradeHistory
                  WHERE StudentID=@s AND AcademicYear > @y
                  ORDER BY AcademicYear ASC LIMIT 1",
                null, ("@s", studentId), ("@y", year))?.ToString();
            if (string.IsNullOrEmpty(gradeAtTime))
                gradeAtTime = Scalar(
                    "SELECT GradeLevel FROM Students WHERE StudentID=@s",
                    null, ("@s", studentId))?.ToString() ?? "";

            // Per-student averages for every classmate in the same (grade, year, sem).
            // A classmate is any student whose own historical grade for year @y was
            // @g — i.e. their next GradeHistory row after @y has FromGrade=@g, OR (for
            // their current year) their Students row matches.
            var peers = Query(@"
                WITH cohort AS (
                    SELECT DISTINCT gh.StudentID
                      FROM GradeHistory gh
                     WHERE gh.AcademicYear > @y AND gh.FromGrade=@g
                       AND gh.AcademicYear = (
                           SELECT MIN(gh2.AcademicYear) FROM GradeHistory gh2
                            WHERE gh2.StudentID = gh.StudentID AND gh2.AcademicYear > @y)
                    UNION
                    SELECT s.StudentID FROM Students s
                     WHERE s.GradeLevel=@g
                       AND NOT EXISTS (SELECT 1 FROM GradeHistory gh3
                                       WHERE gh3.StudentID=s.StudentID AND gh3.AcademicYear > @y)
                       AND EXISTS (SELECT 1 FROM Enrollments e
                                   WHERE e.StudentID=s.StudentID AND e.AcademicYear=@y)
                )
                SELECT e.StudentID AS sid,
                       ROUND(AVG(IFNULL(sc.TotalScore,0)),2) AS PeerAvg,
                       SUM(CASE WHEN sc.TotalScore < @pass AND sc.ScoreID IS NOT NULL THEN 1 ELSE 0 END) AS PeerFails
                FROM cohort c
                JOIN Enrollments e ON e.StudentID = c.StudentID AND e.AcademicYear=@y AND e.Semester=@sm
                JOIN Subjects sub  ON sub.SubjectID = e.SubjectID
                LEFT JOIN Scores sc ON sc.EnrollID = e.EnrollID
                WHERE sub.SubjectCode NOT IN ('CHA1','LAB1')
                GROUP BY e.StudentID
                ORDER BY PeerAvg DESC",
                null, ("@y", year), ("@g", gradeAtTime), ("@sm", sem), ("@pass", PassScore));

            int rank = 0, classSize = peers.Rows.Count;
            // Tie-aware 1224 ranking, but only among PASSING peers; failed peers all rank last.
            int currentRank = 0, processed = 0; double? prev = null;
            foreach (DataRow r in peers.Rows)
            {
                processed++;
                int    peerFails = r["PeerFails"] == DBNull.Value ? 0 : Convert.ToInt32(r["PeerFails"]);
                double peerAvg   = r["PeerAvg"]  == DBNull.Value ? 0 : Convert.ToDouble(r["PeerAvg"]);
                int    sid       = Convert.ToInt32(r["sid"]);
                if (peerFails > 0) continue; // failed peers — treated as ‘ຕົກ’, no numeric rank
                if (prev == null || Math.Abs(peerAvg - prev.Value) > 0.0001)
                    currentRank = processed;
                prev = peerAvg;
                if (sid == studentId) { rank = currentRank; break; }
            }
            return (n, avg, sum, rank, classSize, fails > 0);
        }

        /// <summary>Every student who was in (year, grade, room) HISTORICALLY.
        /// Pivots on GradeHistory + the current Students row for the latest year,
        /// NEVER on Students.GradeLevel/ClassRoom alone (those are current-only).
        /// <para>
        /// <paramref name="statusFilter"/>: optional <c>Students.Status</c> match
        /// (e.g. <c>"ກຳລັງຮຽນ"</c> / <c>"ຈົບ"</c> / <c>"ອອກ"</c>). <c>null</c> = all.
        /// The score-history page exposes this as Active / Graduated / All radios.
        /// </para></summary>
        public static DataTable GetHistoricalClassRoster(string year, string grade, string room,
            string? statusFilter = null)
        {
            // Two sources, UNIONed:
            //   (a) GradeHistory rows that DESCRIBE year=@y for the student — i.e. the
            //       smallest AcademicYear > @y, with FromGrade=@g + ClassRoom=@r.
            //   (b) Students still active in @y with matching current grade+room AND no
            //       later GradeHistory exists (so the current values describe @y).
            // Either source qualifies — a student appears at most once because the outer
            // SELECT groups by StudentID. Status filter (when supplied) is applied at
            // the outer level so it cuts across both sources uniformly.
            string statusSql = string.IsNullOrEmpty(statusFilter) ? "" : " WHERE s.Status=@status";
            var ps = new List<(string, object)> {
                ("@y", year), ("@g", grade), ("@r", room)
            };
            if (!string.IsNullOrEmpty(statusFilter)) ps.Add(("@status", statusFilter));
            return Query($@"
                WITH described AS (
                    SELECT gh.StudentID
                      FROM GradeHistory gh
                     WHERE gh.AcademicYear > @y AND gh.FromGrade=@g AND IFNULL(gh.ClassRoom,'')=@r
                       AND gh.AcademicYear = (
                           SELECT MIN(gh2.AcademicYear) FROM GradeHistory gh2
                            WHERE gh2.StudentID = gh.StudentID AND gh2.AcademicYear > @y)
                    UNION
                    SELECT s.StudentID FROM Students s
                     WHERE s.GradeLevel=@g AND IFNULL(s.ClassRoom,'')=@r
                       AND NOT EXISTS (SELECT 1 FROM GradeHistory gh3
                                       WHERE gh3.StudentID=s.StudentID AND gh3.AcademicYear > @y)
                       AND EXISTS (SELECT 1 FROM Enrollments e
                                   WHERE e.StudentID=s.StudentID AND e.AcademicYear=@y)
                )
                SELECT s.StudentID,
                       s.StudentCode                       AS ລະຫັດ,
                       s.FirstName||' '||s.LastName        AS ຊື່ນັກຮຽນ,
                       s.Gender                            AS ເພດ,
                       s.Status                            AS ສະຖານະ
                FROM described d
                JOIN Students s ON s.StudentID = d.StudentID
                {statusSql}
                ORDER BY s.StudentCode",
                null, ps.ToArray());
        }

        /// <summary>For class-wide history: per-student monthly /10 totals for one
        /// (year, sem, month) of the historical cohort. Rows = students; columns are
        /// dynamically named after subject codes. CHA1/LAB1 columns are present but
        /// only carry whatever monthly entries exist (never aggregated).</summary>
        public static DataTable GetClassMonthGrid(string year, string grade, string room, int month)
        {
            // Subject columns (academic codes only — CHA1/LAB1 omitted from per-month
            // class view since they don't have month-tracked totals teachers care about)
            var subjectCodes = new List<string>();
            foreach (DataRow r in Query("SELECT SubjectCode FROM Subjects WHERE SubjectCode NOT IN ('CHA1','LAB1') ORDER BY SortOrder, SubjectCode").Rows)
                subjectCodes.Add(r["SubjectCode"].ToString()!);

            var roster = GetHistoricalClassRoster(year, grade, room);
            var dt = new DataTable();
            dt.Columns.Add("ລະຫັດ", typeof(string));
            dt.Columns.Add("ຊື່ນັກຮຽນ", typeof(string));
            foreach (var sc in subjectCodes) dt.Columns.Add(sc, typeof(string));

            foreach (DataRow s in roster.Rows)
            {
                int sid = Convert.ToInt32(s["StudentID"]);
                var row = dt.NewRow();
                row["ລະຫັດ"]      = s["ລະຫັດ"];
                row["ຊື່ນັກຮຽນ"] = s["ຊື່ນັກຮຽນ"];
                // Pull this student's monthly totals for the month, keyed by subject code.
                var marks = Query(@"
                    SELECT sub.SubjectCode AS code,
                           ROUND(IFNULL(m.ActivityScore,0)+IFNULL(m.DisciplineScore,0)+IFNULL(m.HomeworkScore,0),2) AS total
                    FROM Enrollments e
                    JOIN Subjects sub          ON sub.SubjectID=e.SubjectID
                    JOIN MonthlyAssessments m  ON m.EnrollID=e.EnrollID
                    WHERE e.StudentID=@s AND e.AcademicYear=@y AND m.Month=@m
                      AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                    null, ("@s", sid), ("@y", year), ("@m", month));
                foreach (DataRow mk in marks.Rows)
                {
                    string code = mk["code"].ToString()!;
                    if (dt.Columns.Contains(code))
                        row[code] = Convert.ToDouble(mk["total"]).ToString("F2");
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>Per-student sem-summary for class-wide history (one row per student
        /// in the cohort, columns: avg / total / rank / level). CHA1/LAB1 excluded.</summary>
        public static DataTable GetClassSemesterSummary(string year, string grade, string room, int sem)
        {
            var roster = GetHistoricalClassRoster(year, grade, room);
            var dt = new DataTable();
            dt.Columns.Add("ລະຫັດ", typeof(string));
            dt.Columns.Add("ຊື່ນັກຮຽນ", typeof(string));
            dt.Columns.Add("ສະເລ່ຍ", typeof(string));
            dt.Columns.Add("ລວມ", typeof(string));
            dt.Columns.Add("ອັນດັບ", typeof(string));
            dt.Columns.Add("ລະດັບ", typeof(string));

            foreach (DataRow s in roster.Rows)
            {
                int sid = Convert.ToInt32(s["StudentID"]);
                var sum = GetHistorySemesterSummary(sid, year, sem);
                string rankTxt = sum.failed ? "ຕົກ" : (sum.rank > 0 ? sum.rank.ToString() : "—");
                string level   = sum.failed ? "ຕົກ" : CalcMoESLevel(sum.avg);
                var row = dt.NewRow();
                row["ລະຫັດ"]      = s["ລະຫັດ"];
                row["ຊື່ນັກຮຽນ"] = s["ຊື່ນັກຮຽນ"];
                row["ສະເລ່ຍ"]     = sum.avg.ToString("F2");
                row["ລວມ"]         = sum.total.ToString("F2");
                row["ອັນດັບ"]     = rankTxt;
                row["ລະດັບ"]      = level;
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>Per-student annual summary for class-wide history. CHA1/LAB1 excluded.</summary>
        public static DataTable GetClassAnnualSummary(string year, string grade, string room)
        {
            var roster = GetHistoricalClassRoster(year, grade, room);
            var dt = new DataTable();
            dt.Columns.Add("ລະຫັດ", typeof(string));
            dt.Columns.Add("ຊື່ນັກຮຽນ", typeof(string));
            dt.Columns.Add("ສະເລ່ຍພາກ 1", typeof(string));
            dt.Columns.Add("ສະເລ່ຍພາກ 2", typeof(string));
            dt.Columns.Add("ສະເລ່ຍປະຈຳປີ", typeof(string));
            dt.Columns.Add("ອັນດັບ", typeof(string));
            dt.Columns.Add("ລະດັບ", typeof(string));

            foreach (DataRow s in roster.Rows)
            {
                int sid = Convert.ToInt32(s["StudentID"]);
                var ann = GetHistoryAnnualSummary(sid, year);
                string rankTxt  = ann.failed ? "ຕົກ" : (ann.rank > 0 ? ann.rank.ToString() : "—");
                string levelTxt = ann.failed ? "ຕົກ" : ann.level;
                var row = dt.NewRow();
                row["ລະຫັດ"]            = s["ລະຫັດ"];
                row["ຊື່ນັກຮຽນ"]       = s["ຊື່ນັກຮຽນ"];
                row["ສະເລ່ຍພາກ 1"]   = ann.sem1Avg.ToString("F2");
                row["ສະເລ່ຍພາກ 2"]   = ann.sem2Avg.ToString("F2");
                row["ສະເລ່ຍປະຈຳປີ"] = ann.annualAvg.ToString("F2");
                row["ອັນດັບ"]            = rankTxt;
                row["ລະດັບ"]            = levelTxt;
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>Per-student per-subject Total for a (year, sem) — same shape
        /// as GetClassMonthGrid but with Scores.TotalScore instead of monthly sum.
        /// CHA1/LAB1 columns deliberately omitted (matches per-month behaviour).
        /// Used by ClassHistoryWindow to render a subject-by-student grid when
        /// the teacher picks S1 or S2 from the report dropdown.</summary>
        public static DataTable GetClassSemesterGrid(string year, string grade, string room, int sem)
        {
            var subjectCodes = new List<string>();
            foreach (DataRow r in Query("SELECT SubjectCode FROM Subjects WHERE SubjectCode NOT IN ('CHA1','LAB1') ORDER BY SortOrder, SubjectCode").Rows)
                subjectCodes.Add(r["SubjectCode"].ToString()!);

            var roster = GetHistoricalClassRoster(year, grade, room);
            var dt = new DataTable();
            dt.Columns.Add("ລະຫັດ", typeof(string));
            dt.Columns.Add("ຊື່ນັກຮຽນ", typeof(string));
            foreach (var sc in subjectCodes) dt.Columns.Add(sc, typeof(string));

            foreach (DataRow s in roster.Rows)
            {
                int sid = Convert.ToInt32(s["StudentID"]);
                var row = dt.NewRow();
                row["ລະຫັດ"]      = s["ລະຫັດ"];
                row["ຊື່ນັກຮຽນ"] = s["ຊື່ນັກຮຽນ"];
                var marks = Query(@"
                    SELECT sub.SubjectCode AS code,
                           ROUND(IFNULL(sc.TotalScore,0),2) AS total
                    FROM Enrollments e
                    JOIN Subjects sub ON sub.SubjectID=e.SubjectID
                    LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                    WHERE e.StudentID=@s AND e.AcademicYear=@y AND e.Semester=@sm
                      AND sub.SubjectCode NOT IN ('CHA1','LAB1')",
                    null, ("@s", sid), ("@y", year), ("@sm", sem));
                foreach (DataRow mk in marks.Rows)
                {
                    string code = mk["code"].ToString()!;
                    if (dt.Columns.Contains(code))
                        row[code] = Convert.ToDouble(mk["total"]).ToString("F2");
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>Per-student per-subject annual score (mean of sem1+sem2
        /// TotalScore) — same shape as GetClassMonthGrid. CHA1/LAB1 columns
        /// deliberately omitted. Used by ClassHistoryWindow for the annual view.</summary>
        public static DataTable GetClassAnnualGrid(string year, string grade, string room)
        {
            var subjectCodes = new List<string>();
            foreach (DataRow r in Query("SELECT SubjectCode FROM Subjects WHERE SubjectCode NOT IN ('CHA1','LAB1') ORDER BY SortOrder, SubjectCode").Rows)
                subjectCodes.Add(r["SubjectCode"].ToString()!);

            var roster = GetHistoricalClassRoster(year, grade, room);
            var dt = new DataTable();
            dt.Columns.Add("ລະຫັດ", typeof(string));
            dt.Columns.Add("ຊື່ນັກຮຽນ", typeof(string));
            foreach (var sc in subjectCodes) dt.Columns.Add(sc, typeof(string));

            foreach (DataRow s in roster.Rows)
            {
                int sid = Convert.ToInt32(s["StudentID"]);
                var row = dt.NewRow();
                row["ລະຫັດ"]      = s["ລະຫັດ"];
                row["ຊື່ນັກຮຽນ"] = s["ຊື່ນັກຮຽນ"];
                // For every academic subject, average the sem1 + sem2 Scores.TotalScore.
                // Rows where a semester has no Scores row → treated as 0 (matches
                // GetHistoryAnnualSummary + BuildIndividualXlsx behaviour).
                var marks = Query(@"
                    SELECT sub.SubjectCode AS code,
                           ROUND((IFNULL(sc1.TotalScore,0) + IFNULL(sc2.TotalScore,0)) / 2.0, 2) AS annual
                    FROM Subjects sub
                    LEFT JOIN Enrollments e1 ON e1.SubjectID=sub.SubjectID
                        AND e1.StudentID=@s AND e1.AcademicYear=@y AND e1.Semester=1
                    LEFT JOIN Enrollments e2 ON e2.SubjectID=sub.SubjectID
                        AND e2.StudentID=@s AND e2.AcademicYear=@y AND e2.Semester=2
                    LEFT JOIN Scores sc1 ON sc1.EnrollID=e1.EnrollID
                    LEFT JOIN Scores sc2 ON sc2.EnrollID=e2.EnrollID
                    WHERE sub.SubjectCode NOT IN ('CHA1','LAB1')
                      AND (e1.EnrollID IS NOT NULL OR e2.EnrollID IS NOT NULL)",
                    null, ("@s", sid), ("@y", year));
                foreach (DataRow mk in marks.Rows)
                {
                    string code = mk["code"].ToString()!;
                    if (dt.Columns.Contains(code))
                        row[code] = Convert.ToDouble(mk["annual"]).ToString("F2");
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>Per-subject annual scorecard for an INDIVIDUAL student.
        /// Academic subjects show sem1/sem2/annual as Scores.TotalScore averages.
        /// CHA1/LAB1 rows show the manually-entered EvaluationScores.Score for
        /// each context (SEM1/SEM2/ANNUAL) — they never contribute to averages.
        /// Used by StudentHistoryWindow when the teacher picks 'ສະຫຼຸບປະຈຳປີ'.</summary>
        public static DataTable GetHistoryAnnual(int studentId, string year)
        {
            return Query(@"
                SELECT sub.SubjectCode  AS ລະຫັດວິຊາ,
                       sub.SubjectName  AS ຊື່ວິຊາ,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1')
                            THEN (SELECT ROUND(ev.Score,2) FROM EvaluationScores ev
                                  WHERE ev.StudentID=@s AND ev.AcademicYear=@y
                                    AND ev.Context='SEM1' AND ev.SubjectCode=sub.SubjectCode)
                            ELSE ROUND(IFNULL(sc1.TotalScore,0),2)
                       END AS ສະເລ່ຍພາກ1,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1')
                            THEN (SELECT ROUND(ev.Score,2) FROM EvaluationScores ev
                                  WHERE ev.StudentID=@s AND ev.AcademicYear=@y
                                    AND ev.Context='SEM2' AND ev.SubjectCode=sub.SubjectCode)
                            ELSE ROUND(IFNULL(sc2.TotalScore,0),2)
                       END AS ສະເລ່ຍພາກ2,
                       CASE WHEN sub.SubjectCode IN ('CHA1','LAB1')
                            THEN (SELECT ROUND(ev.Score,2) FROM EvaluationScores ev
                                  WHERE ev.StudentID=@s AND ev.AcademicYear=@y
                                    AND ev.Context='ANNUAL' AND ev.SubjectCode=sub.SubjectCode)
                            ELSE ROUND((IFNULL(sc1.TotalScore,0)+IFNULL(sc2.TotalScore,0))/2.0, 2)
                       END AS ສະເລ່ຍປະຈຳປີ
                FROM Subjects sub
                LEFT JOIN Enrollments e1 ON e1.SubjectID=sub.SubjectID
                    AND e1.StudentID=@s AND e1.AcademicYear=@y AND e1.Semester=1
                LEFT JOIN Enrollments e2 ON e2.SubjectID=sub.SubjectID
                    AND e2.StudentID=@s AND e2.AcademicYear=@y AND e2.Semester=2
                LEFT JOIN Scores sc1 ON sc1.EnrollID=e1.EnrollID
                LEFT JOIN Scores sc2 ON sc2.EnrollID=e2.EnrollID
                WHERE e1.EnrollID IS NOT NULL OR e2.EnrollID IS NOT NULL
                ORDER BY sub.SortOrder, sub.SubjectCode",
                null, ("@s", studentId), ("@y", year));
        }

        /// <summary>Annual summary for a (student, year): the two semester averages,
        /// the annual average (mean of the two), the student's annual rank against the
        /// historical cohort, the MoES level for the annual average, and whether the
        /// student failed any subject in either semester. CHA1/LAB1 are excluded from
        /// every aggregate.</summary>
        public static (double sem1Avg, double sem2Avg, double annualAvg, int rank, int classSize, string level, bool failed)
            GetHistoryAnnualSummary(int studentId, string year)
        {
            var s1 = GetHistorySemesterSummary(studentId, year, 1);
            var s2 = GetHistorySemesterSummary(studentId, year, 2);
            // Annual = mean of the two semester averages; if one semester is empty,
            // fall back to whichever has data so the page still shows something useful.
            double annual;
            if (s1.subjects > 0 && s2.subjects > 0) annual = Math.Round((s1.avg + s2.avg) / 2.0, 2);
            else if (s1.subjects > 0)               annual = s1.avg;
            else                                    annual = s2.avg;
            bool failed = s1.failed || s2.failed;
            string level = CalcMoESLevel(annual);

            // Reconstruct historical grade (same logic as the semester helper).
            string? gradeAtTime = Scalar(
                @"SELECT FromGrade FROM GradeHistory
                  WHERE StudentID=@s AND AcademicYear > @y
                  ORDER BY AcademicYear ASC LIMIT 1",
                null, ("@s", studentId), ("@y", year))?.ToString();
            if (string.IsNullOrEmpty(gradeAtTime))
                gradeAtTime = Scalar(
                    "SELECT GradeLevel FROM Students WHERE StudentID=@s",
                    null, ("@s", studentId))?.ToString() ?? "";

            // Cohort-wide annual averages: per classmate, the mean of their two semester
            // averages. CHA1/LAB1 excluded by the join filter.
            var peers = Query(@"
                WITH cohort AS (
                    SELECT DISTINCT gh.StudentID
                      FROM GradeHistory gh
                     WHERE gh.AcademicYear > @y AND gh.FromGrade=@g
                       AND gh.AcademicYear = (
                           SELECT MIN(gh2.AcademicYear) FROM GradeHistory gh2
                            WHERE gh2.StudentID = gh.StudentID AND gh2.AcademicYear > @y)
                    UNION
                    SELECT s.StudentID FROM Students s
                     WHERE s.GradeLevel=@g
                       AND NOT EXISTS (SELECT 1 FROM GradeHistory gh3
                                       WHERE gh3.StudentID=s.StudentID AND gh3.AcademicYear > @y)
                       AND EXISTS (SELECT 1 FROM Enrollments e
                                   WHERE e.StudentID=s.StudentID AND e.AcademicYear=@y)
                ),
                per_sem AS (
                    SELECT e.StudentID AS sid, e.Semester AS sm,
                           AVG(IFNULL(sc.TotalScore,0)) AS Avg,
                           SUM(CASE WHEN sc.TotalScore < @pass AND sc.ScoreID IS NOT NULL THEN 1 ELSE 0 END) AS Fails
                      FROM cohort c
                      JOIN Enrollments e ON e.StudentID=c.StudentID AND e.AcademicYear=@y
                      JOIN Subjects sub  ON sub.SubjectID=e.SubjectID
                      LEFT JOIN Scores sc ON sc.EnrollID=e.EnrollID
                     WHERE sub.SubjectCode NOT IN ('CHA1','LAB1')
                     GROUP BY e.StudentID, e.Semester
                )
                SELECT sid,
                       ROUND(AVG(Avg),2) AS AnnualAvg,
                       SUM(Fails)        AS TotalFails
                FROM per_sem
                GROUP BY sid
                ORDER BY AnnualAvg DESC",
                null, ("@y", year), ("@g", gradeAtTime), ("@pass", PassScore));

            int rank = 0, classSize = peers.Rows.Count;
            int currentRank = 0, processed = 0; double? prev = null;
            foreach (DataRow r in peers.Rows)
            {
                processed++;
                int    peerFails = r["TotalFails"] == DBNull.Value ? 0 : Convert.ToInt32(r["TotalFails"]);
                double peerAvg   = r["AnnualAvg"]  == DBNull.Value ? 0 : Convert.ToDouble(r["AnnualAvg"]);
                int    sid       = Convert.ToInt32(r["sid"]);
                if (peerFails > 0) continue;
                if (prev == null || Math.Abs(peerAvg - prev.Value) > 0.0001)
                    currentRank = processed;
                prev = peerAvg;
                if (sid == studentId) { rank = currentRank; break; }
            }
            return (s1.avg, s2.avg, annual, rank, classSize, level, failed);
        }

        // ── Academic-year list ─────────────────────────────────────
        // Single source of truth for every year-combo in the UI.
        // Returns the union of: CurrentYear + the AcademicYears catalogue + every
        // distinct year present in Students / Enrollments / GradeHistory.
        // Deduplicated and sorted newest-first.
        public static List<string> AcademicYears()
        {
            var set = new HashSet<string> { CurrentYear };
            try
            {
                foreach (DataRow r in Query("SELECT Year FROM AcademicYears").Rows)
                    set.Add(r["Year"].ToString()!);
                foreach (DataRow r in Query("SELECT DISTINCT AcademicYear FROM Students   WHERE AcademicYear IS NOT NULL AND AcademicYear<>''").Rows)
                    set.Add(r["AcademicYear"].ToString()!);
                foreach (DataRow r in Query("SELECT DISTINCT AcademicYear FROM Enrollments WHERE AcademicYear IS NOT NULL AND AcademicYear<>''").Rows)
                    set.Add(r["AcademicYear"].ToString()!);
                foreach (DataRow r in Query("SELECT DISTINCT AcademicYear FROM GradeHistory WHERE AcademicYear IS NOT NULL AND AcademicYear<>''").Rows)
                    set.Add(r["AcademicYear"].ToString()!);
            } catch { /* swallow — happens on first run before tables exist */ }
            var list = set.ToList();
            list.Sort((a, b) => string.Compare(b, a, StringComparison.Ordinal)); // descending → newest first
            return list;
        }

        // ── Academic-year management helpers ───────────────────────

        /// <summary>
        /// Returns the year string that comes immediately after <paramref name="year"/>.
        /// "2025-2026" → "2026-2027". If the input is malformed, returns it unchanged.
        /// </summary>
        public static string NextYearString(string year)
        {
            if (string.IsNullOrWhiteSpace(year) || year.Length < 9 || year[4] != '-') return year;
            if (!int.TryParse(year.Substring(0, 4), out int y1)) return year;
            if (!int.TryParse(year.Substring(5, 4), out int y2)) return year;
            return $"{y1 + 1:D4}-{y2 + 1:D4}";
        }

        /// <summary>True if <paramref name="year"/> matches the canonical "YYYY-YYYY" form
        /// where the second year is exactly one greater than the first.</summary>
        public static bool IsValidYearFormat(string year)
        {
            if (string.IsNullOrWhiteSpace(year) || year.Length != 9 || year[4] != '-') return false;
            if (!int.TryParse(year.Substring(0, 4), out int y1)) return false;
            if (!int.TryParse(year.Substring(5, 4), out int y2)) return false;
            return y2 == y1 + 1;
        }

        /// <summary>
        /// Adds a row to the AcademicYears catalogue. Idempotent — INSERT OR IGNORE on
        /// the Year primary key. Does NOT change which year is current.
        /// </summary>
        public static void CreateAcademicYear(string year, string? startDate, string? endDate, string? note)
        {
            if (!IsValidYearFormat(year))
                throw new ArgumentException($"Invalid year format: '{year}' (expected YYYY-YYYY)");
            Exec(@"INSERT OR IGNORE INTO AcademicYears(Year, StartDate, EndDate, Note, CreatedBy)
                   VALUES(@y, @s, @e, @n, @by)",
                null,
                ("@y", year),
                ("@s", (object?)startDate ?? DBNull.Value),
                ("@e", (object?)endDate ?? DBNull.Value),
                ("@n", (object?)note ?? DBNull.Value),
                ("@by", CurrentUser));
            Log("CreateAcademicYear", year);
        }

        // ════════════════════════════════════════════════════════════
        //  MEMORY / CONTEXT SUBSYSTEM
        // ════════════════════════════════════════════════════════════
        //  Tunables (kept simple — the user can change them later).
        public const int SummarizeAfterMessages = 30;  // trigger summary every 30 msgs
        public const int RecentRawMessagesKeep  = 8;   // raw messages kept in built context
        public const int CharsPerTokenEstimate  = 3;   // ~3 chars/token rough average

        /// <summary>Conservative token estimator. ~3 chars per token works for
        /// Lao + Latin mixed text. Replace with a tokenizer call if precision needed.</summary>
        public static int EstimateTokens(string text)
            => string.IsNullOrEmpty(text) ? 0 : (text.Length + CharsPerTokenEstimate - 1) / CharsPerTokenEstimate;

        /// <summary>Create a new conversation thread. Returns the new ConvID.</summary>
        public static int StartConversation(string title)
        {
            using var c = Open();
            Exec("INSERT INTO Conversations(Title, StartedBy) VALUES(@t, @u)",
                 c, ("@t", title ?? "Untitled"), ("@u", CurrentUser));
            int id = ScalarInt("SELECT last_insert_rowid()", c);
            Log("StartConversation", $"#{id} {title}");
            return id;
        }

        /// <summary>Append a message to a conversation. Updates the parent counters.
        /// Auto-generates a summary when the message count since the last summary
        /// exceeds <see cref="SummarizeAfterMessages"/>.</summary>
        public static int AddMessage(int convId, string role, string content)
        {
            int tokens = EstimateTokens(content);
            using var c = Open();
            Exec(@"INSERT INTO Messages(ConvID, Role, Content, TokenEst, CreatedBy)
                   VALUES(@c, @r, @t, @tok, @u)",
                c,
                ("@c", convId), ("@r", role ?? "user"), ("@t", content ?? ""),
                ("@tok", tokens), ("@u", CurrentUser));
            int msgId = ScalarInt("SELECT last_insert_rowid()", c);
            Exec(@"UPDATE Conversations
                   SET MessageCount   = MessageCount + 1,
                       TotalTokensEst = TotalTokensEst + @tok,
                       LastMessageAt  = datetime('now','localtime')
                   WHERE ConvID = @c",
                 c, ("@c", convId), ("@tok", tokens));

            // Auto-summarize trigger (skip when the message we just added IS the summary)
            if (role != "summary" && ShouldSummarize(convId))
                GenerateSummary(convId);

            return msgId;
        }

        /// <summary>True when there are at least <see cref="SummarizeAfterMessages"/>
        /// non-summary messages since the last summary row.</summary>
        public static bool ShouldSummarize(int convId)
        {
            int lastSum = ScalarInt(
                "SELECT IFNULL(MAX(MsgID), 0) FROM Messages WHERE ConvID=@c AND Role='summary'",
                null, ("@c", convId));
            int unsummarised = ScalarInt(
                "SELECT COUNT(*) FROM Messages WHERE ConvID=@c AND MsgID > @s AND Role <> 'summary'",
                null, ("@c", convId), ("@s", lastSum));
            return unsummarised >= SummarizeAfterMessages;
        }

        /// <summary>
        /// Generate a summary of every non-summary message after the most-recent summary
        /// row, then store it as a new 'summary' role message. Heuristic: bullet-style
        /// extraction of role + first 200 chars per message, capped to ~12k characters
        /// total. Replace with an LLM call where indicated for a real abstractive summary.
        /// Returns the new summary MsgID, or 0 if nothing to summarise.
        /// </summary>
        public static int GenerateSummary(int convId)
        {
            int lastSum = ScalarInt(
                "SELECT IFNULL(MAX(MsgID), 0) FROM Messages WHERE ConvID=@c AND Role='summary'",
                null, ("@c", convId));
            var dt = Query(@"SELECT MsgID, Role, Content
                             FROM Messages
                             WHERE ConvID=@c AND MsgID > @s AND Role <> 'summary'
                             ORDER BY MsgID",
                null, ("@c", convId), ("@s", lastSum));
            if (dt.Rows.Count == 0) return 0;

            var sb = new StringBuilder();
            sb.AppendLine($"[Summary of {dt.Rows.Count} messages, {dt.Rows[0]["MsgID"]} → {dt.Rows[dt.Rows.Count - 1]["MsgID"]}]");
            int firstId = Convert.ToInt32(dt.Rows[0]["MsgID"]);
            int lastId  = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1]["MsgID"]);
            foreach (DataRow r in dt.Rows)
            {
                string content = r["Content"].ToString() ?? "";
                content = content.Replace("\r", " ").Replace("\n", " ").Trim();
                if (content.Length > 200) content = content.Substring(0, 200) + "…";
                sb.AppendLine($"  • {r["Role"]}: {content}");
                if (sb.Length > 12000) { sb.AppendLine("  • [truncated — too long for heuristic summary]"); break; }
            }
            // === LLM hook point ===
            // To replace with a real abstractive summary, call your LLM here with the
            // concatenated messages as the prompt, and use the returned text as 'sb'.
            // ======================
            return AddMessage(convId, "summary", sb.ToString());
        }

        /// <summary>Compact snapshot of the school's current state. This is what
        /// ‘persist project settings, academic year, semester, student data’ resolves
        /// to: a few hundred tokens of canonical context guaranteed to be fresh.</summary>
        public static string BuildSystemContext()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== SCHOOL CONTEXT (auto-generated) ===");
            sb.AppendLine($"School        : {SchoolName}");
            sb.AppendLine($"Academic Year : {CurrentYear}");
            sb.AppendLine($"Semester      : {CurrentSem}  (months {string.Join("/", MonthsInSemester(CurrentSem))})");
            sb.AppendLine($"Final Exam    : Sem1 → Jan,  Sem2 → Jun");
            sb.AppendLine($"Grading       : monthly × {MidPct:F0}% + final × {FinalPct:F0}% = total /10  |  pass ≥ {PassScore}");
            sb.AppendLine($"User          : {CurrentUser}  ({CurrentRole})");
            sb.AppendLine();
            try
            {
                int active  = ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ກຳລັງຮຽນ'");
                int gradTot = ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ຈົບ'");
                int subjs   = ScalarInt("SELECT COUNT(*) FROM Subjects");
                int teach   = ScalarInt("SELECT COUNT(*) FROM Users WHERE IsActive=1");
                sb.AppendLine($"Active students : {active}");
                sb.AppendLine($"Graduated       : {gradTot}");
                sb.AppendLine($"Subjects        : {subjs}");
                sb.AppendLine($"Active users    : {teach}");

                // Per-grade breakdown — small and useful in context
                var perGrade = Query(@"SELECT GradeLevel, COUNT(*) AS N
                                       FROM Students
                                       WHERE Status='ກຳລັງຮຽນ' AND GradeLevel<>''
                                       GROUP BY GradeLevel ORDER BY GradeLevel");
                if (perGrade.Rows.Count > 0)
                {
                    var parts = new List<string>();
                    foreach (DataRow r in perGrade.Rows)
                        parts.Add($"{r["GradeLevel"]}={r["N"]}");
                    sb.AppendLine("By grade        : " + string.Join("  ", parts));
                }
            }
            catch { /* fresh DB without seeded data — fine */ }

            sb.AppendLine("=== END SCHOOL CONTEXT ===");
            return sb.ToString();
        }

        /// <summary>
        /// Token-optimised context for a conversation. Produces:
        ///   [system context]
        ///   [latest summary]            ← if any
        ///   [last N raw messages]       ← chronological
        /// Use this as the prompt to an external LLM instead of the full transcript.
        /// </summary>
        public static string LoadOptimizedContext(int convId, int? recentMessages = null)
        {
            int keep = recentMessages ?? RecentRawMessagesKeep;
            var sb = new StringBuilder();
            sb.AppendLine(BuildSystemContext());

            var latestSummary = Query(
                "SELECT Content FROM Messages WHERE ConvID=@c AND Role='summary' ORDER BY MsgID DESC LIMIT 1",
                null, ("@c", convId));
            if (latestSummary.Rows.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("=== CONVERSATION SUMMARY SO FAR ===");
                sb.AppendLine(latestSummary.Rows[0]["Content"].ToString());
                sb.AppendLine("=== END SUMMARY ===");
            }

            // Recent raw messages AFTER the latest summary (chronological)
            int lastSumId = ScalarInt(
                "SELECT IFNULL(MAX(MsgID), 0) FROM Messages WHERE ConvID=@c AND Role='summary'",
                null, ("@c", convId));
            var recent = Query(@"SELECT Role, Content, CreatedAt FROM Messages
                                 WHERE ConvID=@c AND MsgID > @s AND Role <> 'summary'
                                 ORDER BY MsgID DESC LIMIT @lim",
                null, ("@c", convId), ("@s", lastSumId), ("@lim", keep));
            if (recent.Rows.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine($"=== LAST {recent.Rows.Count} MESSAGES ===");
                for (int i = recent.Rows.Count - 1; i >= 0; i--)
                {
                    var r = recent.Rows[i];
                    sb.AppendLine($"[{r["CreatedAt"]}] {r["Role"]}: {r["Content"]}");
                }
                sb.AppendLine("=== END MESSAGES ===");
            }
            return sb.ToString();
        }

        /// <summary>Keyword search across all conversations. Returns the most recent
        /// matching messages — useful for "find what we previously said about X".</summary>
        public static DataTable SearchMessages(string keyword, int limit = 25)
        {
            return Query(@"SELECT c.ConvID, c.Title, m.MsgID, m.Role, m.Content, m.CreatedAt
                           FROM Messages m
                           JOIN Conversations c ON c.ConvID = m.ConvID
                           WHERE m.Content LIKE @k AND m.Role <> 'summary'
                           ORDER BY m.MsgID DESC
                           LIMIT @lim",
                null, ("@k", "%" + (keyword ?? "") + "%"), ("@lim", limit));
        }

        /// <summary>
        /// Switches the system's active academic year:
        ///   - Updates AcademicYears.IsCurrent so the new year is the only flagged row
        ///   - Inserts the year into the catalogue if missing
        ///   - Sets Settings.current_year to the new year
        ///   - Resets Settings.current_semester to 1 (new year always starts at Sem 1)
        ///   - Refreshes the static settings cache so DB.CurrentYear/CurrentSem are live
        ///
        /// All historical Students / Enrollments / Scores / MonthlyAssessments /
        /// AttendanceRecords are left untouched — student promotion is handled
        /// separately via PromotionPage.
        /// </summary>
        public static void SetCurrentAcademicYear(string year)
        {
            if (!IsValidYearFormat(year))
                throw new ArgumentException($"Invalid year format: '{year}' (expected YYYY-YYYY)");
            using var c = Open();
            // Ensure the year exists in the catalogue.
            Exec("INSERT OR IGNORE INTO AcademicYears(Year, IsCurrent) VALUES(@y, 0)",
                c, ("@y", year));
            // Flip IsCurrent: only the chosen year keeps it set.
            Exec("UPDATE AcademicYears SET IsCurrent = CASE WHEN Year=@y THEN 1 ELSE 0 END",
                c, ("@y", year));
            // Mirror into Settings (which is what the rest of the app actually reads).
            SaveSetting("current_year", year);
            SaveSetting("current_semester", "1");
            Log("SetCurrentYear", year);
        }

        // Lower-secondary only: ມ.1 (Grade 7) → ມ.2 → ມ.3 → ມ.4 (Grade 10) → ຈົບ.
        public static string NextGrade(string g) => g switch {
            "ມ.1" => "ມ.2",
            "ມ.2" => "ມ.3",
            "ມ.3" => "ມ.4",
            "ມ.4" => "ຈົບ",
            _     => g
        };

        public static void Backup(string dest) { File.Copy(DbPath,dest,true); Log("Backup",dest); }
        public static void Restore(string src) { File.Copy(src,DbPath,true);  Log("Restore",src); }

        /// <summary>
        /// Parse a teacher-typed score string into a clamped double.
        /// Returns <c>true</c> + the parsed value (always in <c>[min, max]</c>, rounded
        /// to 2 decimals) when the input is a valid number. Returns <c>false</c> when
        /// the input is non-numeric and the score-entry grids should reject it.
        /// <para>
        /// Empty / whitespace input maps to <c>0</c> (so clearing a cell == zero, not
        /// "invalid"). Both dot and comma decimal separators are accepted.
        /// </para>
        /// Used by both ScoresPage (Final) and MonthlyScoresPage (Activity / Discipline /
        /// Homework / Eval). Each column passes its own <paramref name="max"/>:
        /// Activity=3 · Discipline=2 · Homework=5 · Eval/Final=10.
        /// </summary>
        public static bool TryParseScore(string? raw, double min, double max, out double value)
        {
            value = 0;
            string s = (raw ?? "").Trim();
            if (s.Length == 0) return true;                  // empty → 0
            // Accept both "8.5" and "8,5" (Lao Windows locale often uses ',').
            s = s.Replace(',', '.');
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                return false;
            if (double.IsNaN(v) || double.IsInfinity(v))     return false;
            if (v < min) v = min;
            if (v > max) v = max;
            value = Math.Round(v, 2);
            return true;
        }

        /// <summary>
        /// STRICT integer-only score parser used by both score-entry grids.
        /// Returns <c>true</c> + the parsed value (always in <c>[min, max]</c>) ONLY
        /// when the input is a non-empty string of digits 0-9 that parses to an integer
        /// inside the allowed range. Rejects anything else — decimals (<c>"7.5"</c>),
        /// locale separators (<c>","</c>), letters (<c>"abc"</c>), signs, whitespace
        /// inside the number, etc. Empty/whitespace input maps to <c>0</c> (so clearing
        /// a cell == zero, never "invalid").
        /// </summary>
        public static bool TryParseIntScore(string? raw, int min, int max, out int value)
        {
            value = 0;
            string s = (raw ?? "").Trim();
            if (s.Length == 0) return true;             // empty → 0
            // Manual digit-only check — rejects "7.5", "8,5", "-1", "1e2", "1 2", etc.
            foreach (char c in s) if (c < '0' || c > '9') return false;
            // After the loop we know s is 1+ digit chars → safe int.TryParse.
            if (!int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out int v))
                return false;
            if (v < min || v > max) return false;        // strict range: NO silent clamping
            value = v;
            return true;
        }

        /// <summary>Strip Windows-invalid filename characters. Replaces each of
        /// <c>\ / : * ? &quot; &lt; &gt; |</c> with '_', drops control chars, trims trailing
        /// dots/spaces, and re-prefixes reserved DOS device names (CON, PRN, AUX,
        /// NUL, COM1-9, LPT1-9). Used by every report export so user-typed
        /// filenames never break SaveFileDialog or NTFS path rules.</summary>
        public static string SafeFileName(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "report";
            var sb = new StringBuilder(raw.Length);
            foreach (char c in raw)
            {
                if (char.IsControl(c)) continue;
                sb.Append("\\/:*?\"<>|".IndexOf(c) >= 0 ? '_' : c);
            }
            string s = sb.ToString().Trim().TrimEnd('.', ' ');
            if (s.Length == 0) return "report";
            string baseName = Path.GetFileNameWithoutExtension(s);
            string[] reserved = { "CON","PRN","AUX","NUL","COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9" };
            foreach (var r in reserved)
                if (string.Equals(baseName, r, StringComparison.OrdinalIgnoreCase))
                    return "_" + s;
            return s;
        }

        public static DataTable Query(string sql, SQLiteConnection? ext=null, params (string,object)[] ps)
        {
            bool owns=ext==null; var c=ext??Open();
            try { using var cmd=new SQLiteCommand(sql,c); foreach(var(k,v)in ps)cmd.Parameters.AddWithValue(k,v); var dt=new DataTable(); new SQLiteDataAdapter(cmd).Fill(dt); return dt; }
            finally { if(owns)c.Dispose(); }
        }
        public static int Exec(string sql, SQLiteConnection? ext=null, params (string,object)[] ps)
        {
            bool owns=ext==null; var c=ext??Open();
            try { using var cmd=new SQLiteCommand(sql,c); foreach(var(k,v)in ps)cmd.Parameters.AddWithValue(k,v); return cmd.ExecuteNonQuery(); }
            finally { if(owns)c.Dispose(); }
        }
        public static int ExecTx(string sql, SQLiteConnection c, SQLiteTransaction tx, params (string,object)[] ps)
        { using var cmd=new SQLiteCommand(sql,c,tx); foreach(var(k,v)in ps)cmd.Parameters.AddWithValue(k,v); return cmd.ExecuteNonQuery(); }
        public static object? Scalar(string sql, SQLiteConnection? ext=null, params (string,object)[] ps)
        {
            bool owns=ext==null; var c=ext??Open();
            try { using var cmd=new SQLiteCommand(sql,c); foreach(var(k,v)in ps)cmd.Parameters.AddWithValue(k,v); return cmd.ExecuteScalar(); }
            finally { if(owns)c.Dispose(); }
        }
        public static int ScalarInt(string sql, SQLiteConnection? ext=null, params (string,object)[] ps)
            => Convert.ToInt32(Scalar(sql,ext,ps)??0);
        public static void Log(string action, string detail="", SQLiteConnection? ext=null)
        {
            bool owns=ext==null; var c=ext??Open();
            try { using var cmd=new SQLiteCommand("INSERT INTO ActivityLog(UserID,Username,Action,Detail) VALUES(@uid,@u,@a,@d)",c); cmd.Parameters.AddWithValue("@uid",CurrentUserId==0?(object)DBNull.Value:CurrentUserId); cmd.Parameters.AddWithValue("@u",CurrentUser); cmd.Parameters.AddWithValue("@a",action); cmd.Parameters.AddWithValue("@d",detail); cmd.ExecuteNonQuery(); }
            catch{} finally{if(owns)c?.Dispose();}
        }
    }
}
