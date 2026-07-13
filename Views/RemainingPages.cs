using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Microsoft.Win32;
using StudentSIS.Data;

namespace StudentSIS.Views
{
    // ════════════════════════════════════════════════════════════
    //  ENROLLMENT PAGE (per student)
    // ════════════════════════════════════════════════════════════
    public class EnrollmentPage : UserControl
    {
        // Enrollment is now year-wide (no semester picker).
        // Each registration creates two Enrollments rows (Sem 1 + Sem 2) for the same subject,
        // so the student is automatically registered for both grading periods of the year.
        private ComboBox _cmbStu = null!, _cmbYear = null!, _cmbGrade = null!;
        private DataGrid _dg = null!;
        private TextBlock _lbl = null!;
        // Guard so programmatic repopulation of the student list doesn't fire LoadEnroll
        // mid-rebuild.
        private bool _loadingStudents = false;

        public EnrollmentPage()
        {
            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star));
            var tb   = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,10,14,10));
            var flow = new WrapPanel();
            // Year + grade filters narrow the student dropdown so teachers can pick a
            // student by school year and class level. The selected year is also the
            // year that subjects get registered under.
            _cmbYear  = H.MkCmb(DB.AcademicYears().ToArray(), 90);
            SelectComboValue(_cmbYear, DB.CurrentYear);
            _cmbGrade = H.MkCmb(new[]{ "ທັງໝົດ", "ມ.1", "ມ.2", "ມ.3", "ມ.4" }, 90);
            _cmbStu   = new ComboBox { Width = 300, Margin = new Thickness(0,0,12,0) };
            _cmbStu.SelectionChanged  += (s,e) => { if (!_loadingStudents) LoadEnroll(); };
            _cmbYear.SelectionChanged += (s,e) => ReloadStudents();
            _cmbGrade.SelectionChanged += (s,e) => ReloadStudents();
            var bAdd = H.Btn("⚡  ລົງທະບຽນທຸກວິຊາ", "SuccessButton"); bAdd.Click += (s,e) => EnrollAllSubjects();
            var bRem = H.Btn("🗑  ຖອນວິຊາ",          "DangerButton");  bRem.Click += (s,e) => RemoveSubject();
            _lbl = new TextBlock {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12,0,0,0),
                FontSize = 13,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))
            };
            flow.Children.Add(H.Lbl("ປີ:"));
            flow.Children.Add(_cmbYear);
            flow.Children.Add(H.Lbl("ຊັ້ນ:"));
            flow.Children.Add(_cmbGrade);
            flow.Children.Add(H.Lbl("ນັກຮຽນ:"));
            flow.Children.Add(_cmbStu);
            flow.Children.Add(bAdd);
            flow.Children.Add(bRem);
            flow.Children.Add(_lbl);
            tb.Child = flow; Grid.SetRow(tb, 0); root.Children.Add(tb);

            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White
            };
            _dg.Columns.Add(H.Col("ລະຫັດວິຊາ", "SubjectCode", 100));
            _dg.Columns.Add(H.ColStar("ຊື່ວິຊາ", "SubjectName"));
            _dg.Columns.Add(H.Col("ປະເພດ", "Category", 120));
            _dg.Columns.Add(H.Col("ຄູ", "Teacher", 160));
            var card = H.MkCard(); card.Child = _dg;
            Grid.SetRow(card, 1); root.Children.Add(card);
            Content = root;

            ReloadStudents();
        }

        // Select the item whose text matches `val`; no-op if not found.
        private static void SelectComboValue(ComboBox cb, string val)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == val) { cb.SelectedIndex = i; return; }
        }

        // Repopulate the student dropdown for the selected year + grade. Only active
        // students appear — graduated / withdrawn students should never have new
        // subjects added under their record.
        private void ReloadStudents()
        {
            string year  = _cmbYear.SelectedItem?.ToString()  ?? DB.CurrentYear;
            string grade = _cmbGrade.SelectedItem?.ToString() ?? "ທັງໝົດ";

            var sb = new StringBuilder(@"
                SELECT StudentID,
                       StudentCode || '  ·  ' || FirstName || ' ' || LastName ||
                       '  ·  ' || GradeLevel || '/' || IFNULL(ClassRoom,'-') AS D
                FROM Students
                WHERE Status='ກຳລັງຮຽນ' AND AcademicYear=@yr");
            var ps = new List<(string, object)> { ("@yr", year) };
            if (grade != "ທັງໝົດ")
            {
                sb.Append(" AND GradeLevel=@g");
                ps.Add(("@g", grade));
            }
            sb.Append(" ORDER BY GradeLevel, ClassRoom, StudentCode");

            var dt = DB.Query(sb.ToString(), null, ps.ToArray());

            _loadingStudents = true;
            _cmbStu.DisplayMemberPath = "D";
            _cmbStu.SelectedValuePath = "StudentID";
            _cmbStu.ItemsSource       = dt.DefaultView;
            if (dt.Rows.Count > 0) _cmbStu.SelectedIndex = 0;
            _loadingStudents = false;

            // Refresh the subject grid for the (new) first student, or clear it.
            if (dt.Rows.Count > 0) LoadEnroll();
            else { _dg.ItemsSource = null; _lbl.Text = "📚 ບໍ່ມີນັກຮຽນທີ່ຕົງກັບຕົວກອງ"; }
        }

        private void LoadEnroll()
        {
            if (_cmbStu.SelectedValue == null) return;
            int    sid  = Convert.ToInt32(_cmbStu.SelectedValue);
            string year = _cmbYear.SelectedItem?.ToString() ?? DB.CurrentYear;

            // One row per subject regardless of how many semester-enrollments exist.
            // Both Sem 1 + Sem 2 share the same subject identity.
            var dt = DB.Query(@"
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
                null, ("@sid", sid), ("@yr", year));
            _dg.ItemsSource = dt.DefaultView;
            _lbl.Text = $"📚 {dt.Rows.Count} ວິຊາ  (ປີ {year} — ທັງສອງພາກ)";
        }

        // One-click enrollment: auto-creates Enrollment rows for every subject defined
        // for the selected student's grade, both semesters, in the selected academic year.
        // Idempotent — uses INSERT OR IGNORE against the
        // UNIQUE(StudentID,SubjectID,AcademicYear,Semester) constraint so re-running is safe.
        private void EnrollAllSubjects()
        {
            if (_cmbStu.SelectedValue == null)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            int    sid   = Convert.ToInt32(_cmbStu.SelectedValue);
            string year  = _cmbYear.SelectedItem?.ToString() ?? DB.CurrentYear;
            string grade = DB.Scalar("SELECT GradeLevel FROM Students WHERE StudentID=@i",
                                      null, ("@i", sid))?.ToString() ?? "";
            string label = _cmbStu.Text;

            // Subjects valid for the student's grade — either explicitly tagged to this
            // grade or grade-agnostic (NULL / empty GradeLevel = applies to all grades).
            // In Buekthong's setup ມ.1–ມ.4 share one subject set, so this normally
            // matches the full 14-subject list maintained on the ‘ວິຊາ’ page.
            var subjects = DB.Query(@"
                SELECT SubjectID, SubjectCode, SubjectName FROM Subjects
                WHERE GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel=''
                ORDER BY SortOrder, SubjectCode",
                null, ("@g", grade));

            if (subjects.Rows.Count == 0)
            {
                MessageBox.Show($"ບໍ່ມີວິຊາທີ່ກຳນົດໄວ້ສຳລັບຊັ້ນ {grade}\n" +
                                "ກະລຸນາເພີ່ມວິຊາໃນໜ້າ ‘ວິຊາ’ ກ່ອນ",
                    "ບໍ່ມີຂໍ້ມູນ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Confirm before bulk-inserting — gives the teacher a chance to abort if
            // they picked the wrong student / wrong year combo by mistake.
            if (MessageBox.Show(
                    $"ລົງທະບຽນວິຊາທັງໝົດ ({subjects.Rows.Count} ວິຊາ × ສອງພາກ) ໃຫ້:\n\n" +
                    $"   ນັກຮຽນ: {label}\n" +
                    $"   ປີ:    {year}\n\n" +
                    "ດຳເນີນການຕໍ່?",
                    "ຢືນຢັນ", MessageBoxButton.YesNo, MessageBoxImage.Question)
                != MessageBoxResult.Yes) return;

            int newSubjects = 0, existedSubjects = 0;
            using var conn = DB.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                foreach (DataRow r in subjects.Rows)
                {
                    int subId = Convert.ToInt32(r["SubjectID"]);
                    int inserted = 0;
                    for (int sem = 1; sem <= 2; sem++)
                    {
                        int n = DB.ExecTx(
                            @"INSERT OR IGNORE INTO Enrollments(StudentID, SubjectID, AcademicYear, Semester)
                              VALUES(@s, @sub, @y, @sm)",
                            conn, tx,
                            ("@s", sid), ("@sub", subId), ("@y", year), ("@sm", sem));
                        inserted += n;
                    }
                    if (inserted > 0) newSubjects++;
                    else              existedSubjects++;
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ລົງທະບຽນບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DB.Log("AutoEnroll", $"sid={sid} grade={grade} year={year} new={newSubjects} existed={existedSubjects}");
            MessageBox.Show(
                $"ລົງທະບຽນສຳເລັດ:\n\n" +
                $"   ວິຊາທັງໝົດສຳລັບຊັ້ນ {grade}: {subjects.Rows.Count}\n" +
                $"   ບັນທຶກໃໝ່: {newSubjects} ວິຊາ\n" +
                $"   ມີຢູ່ແລ້ວ (ຂ້າມ): {existedSubjects} ວິຊາ\n\n" +
                $"(ແຕ່ລະວິຊາລົງທະບຽນທັງສອງພາກໂດຍອັດຕະໂນມັດ)",
                "ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Information);
            LoadEnroll();
        }

        private void RemoveSubject()
        {
            if (_dg.SelectedItem is not DataRowView drv) return;
            int    subId = Convert.ToInt32(drv["SubID"]);
            int    sid   = Convert.ToInt32(_cmbStu.SelectedValue);
            string year  = _cmbYear.SelectedItem?.ToString() ?? DB.CurrentYear;
            string label = $"{drv["SubjectCode"]} — {drv["SubjectName"]}";

            if (MessageBox.Show($"ຖອນວິຊາ ‘{label}’ ອອກຈາກປີ {year}? (ທັງສອງພາກ)",
                                "ຢືນຢັນ", MessageBoxButton.YesNo, MessageBoxImage.Warning)
                != MessageBoxResult.Yes) return;

            // Remove both Sem 1 + Sem 2 enrollment rows for this subject/year. Scores +
            // MonthlyAssessments cascade automatically via FK ON DELETE CASCADE.
            DB.Exec("DELETE FROM Enrollments WHERE StudentID=@s AND SubjectID=@sub AND AcademicYear=@y",
                    null, ("@s", sid), ("@sub", subId), ("@y", year));
            DB.Log("Unenroll", label);
            LoadEnroll();
        }
    }

    // EnrollPickWindow + PickItem were removed when the per-student enrollment flow
    // moved to one-click "ລົງທະບຽນທຸກວິຊາ" (auto-enroll every subject for the grade).
    // BatchEnrollPage still exists separately for class-wide bulk operations.

    // ════════════════════════════════════════════════════════════
    //  BATCH ENROLL
    //
    //  Class-wide one-click enrolment. The teacher picks a year + grade
    //  (and optionally a room) and clicks ⚡ ລົງທະບຽນທຸກນັກຮຽນ. Every
    //  active student in that filter gets auto-enrolled in every subject
    //  defined for their grade, both semesters, via INSERT OR IGNORE.
    //
    //  Replaces the previous two-list (students × subjects) checkbox UI.
    //  The subject set is always the full ‘ວິຊາ’ list for the grade —
    //  there's no subject-selection step, mirroring the per-student
    //  EnrollmentPage's simplified flow.
    // ════════════════════════════════════════════════════════════
    public class BatchEnrollPage : UserControl
    {
        private ComboBox  _grade = null!, _year = null!, _room = null!;
        private DataGrid  _dg = null!;
        private TextBlock _lbl = null!;
        private Button    _btnEnroll = null!;

        public BatchEnrollPage()
        {
            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // ── Filter bar ───────────────────────────────────────
            var tb   = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,10,14,10));
            var flow = new WrapPanel();
            _year  = H.MkCmb(DB.AcademicYears().ToArray(), 90);
            SelectComboValueBE(_year, DB.CurrentYear);
            _grade = H.MkCmb(new[] { "ມ.1", "ມ.2", "ມ.3", "ມ.4" }, 80);
            _grade.SelectedIndex = 3; // ມ.4 — most-used class
            _room  = H.MkCmb(new[] { "ທັງໝົດ", "1", "2", "3", "4", "5", "6" }, 90);
            _year.SelectionChanged  += (s,e) => LoadStudents();
            _grade.SelectionChanged += (s,e) => LoadStudents();
            _room.SelectionChanged  += (s,e) => LoadStudents();
            var bLoad = H.Btn("🔄  ໂຫຼດ", "PrimaryButton"); bLoad.Click += (s,e) => LoadStudents();
            flow.Children.Add(H.Lbl("ປີ:"));    flow.Children.Add(_year);
            flow.Children.Add(H.Lbl("ຊັ້ນ:")); flow.Children.Add(_grade);
            flow.Children.Add(H.Lbl("ຫ້ອງ:")); flow.Children.Add(_room);
            flow.Children.Add(bLoad);
            tb.Child = flow; Grid.SetRow(tb, 0); root.Children.Add(tb);

            // ── Roster (read-only) ───────────────────────────────
            // Shows exactly which students will be enrolled. No checkboxes —
            // "load" produces the universe; "enrol" processes that universe.
            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White
            };
            _dg.Columns.Add(H.Col("ລະຫັດ",     "StudentCode", 110));
            _dg.Columns.Add(H.ColStar("ຊື່ນັກຮຽນ", "FullName"));
            _dg.Columns.Add(H.Col("ຫ້ອງ",      "ClassRoom",   70));
            _dg.Columns.Add(H.Col("ປີ",        "AcademicYear", 100));
            var card = H.MkCard(); card.Child = _dg;
            Grid.SetRow(card, 1); root.Children.Add(card);

            // ── Footer: status + action button ───────────────────
            var foot = H.MkCard(new Thickness(0,10,0,0), new Thickness(14,10,14,10));
            var fp = new DockPanel();
            _lbl = new TextBlock {
                VerticalAlignment = VerticalAlignment.Center, FontSize = 13,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))
            };
            _btnEnroll = H.Btn("⚡  ລົງທະບຽນທຸກນັກຮຽນ", "SuccessButton");
            _btnEnroll.Width = 240;
            _btnEnroll.Click += (s,e) => EnrollAll();
            DockPanel.SetDock(_btnEnroll, Dock.Right);
            fp.Children.Add(_lbl); fp.Children.Add(_btnEnroll);
            foot.Child = fp; Grid.SetRow(foot, 2); root.Children.Add(foot);
            Content = root;

            // NavContext handoff from ClassHubPage (pre-fills grade + year).
            if (!string.IsNullOrEmpty(DB.NavGrade))
            {
                for (int i = 0; i < _grade.Items.Count; i++)
                    if (_grade.Items[i]?.ToString() == DB.NavGrade) { _grade.SelectedIndex = i; break; }
                if (!string.IsNullOrEmpty(DB.NavRoom))
                    for (int i = 0; i < _room.Items.Count; i++)
                        if (_room.Items[i]?.ToString() == DB.NavRoom) { _room.SelectedIndex = i; break; }
                if (!string.IsNullOrEmpty(DB.NavYear))
                    for (int i = 0; i < _year.Items.Count; i++)
                        if (_year.Items[i]?.ToString() == DB.NavYear) { _year.SelectedIndex = i; break; }
                DB.ClearNav();
            }

            LoadStudents();
        }

        // Select the combobox item whose text matches `val`; no-op if absent.
        private static void SelectComboValueBE(ComboBox cb, string val)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == val) { cb.SelectedIndex = i; return; }
        }

        // Refresh the roster preview for the current year + grade (+ optional room).
        // Status text updates with a subject count so the teacher sees the scope
        // BEFORE clicking ⚡ Enrol.
        private void LoadStudents()
        {
            if (_dg == null) return;
            string year  = _year.SelectedItem?.ToString()  ?? DB.CurrentYear;
            string grade = _grade.SelectedItem?.ToString() ?? "ມ.4";
            string room  = _room.SelectedItem?.ToString()  ?? "ທັງໝົດ";

            var sb = new StringBuilder(@"
                SELECT StudentID, StudentCode,
                       FirstName || ' ' || LastName AS FullName,
                       ClassRoom, AcademicYear
                FROM Students
                WHERE Status='ກຳລັງຮຽນ' AND GradeLevel=@g AND AcademicYear=@y");
            var ps = new List<(string, object)> { ("@g", grade), ("@y", year) };
            if (room != "ທັງໝົດ") { sb.Append(" AND ClassRoom=@r"); ps.Add(("@r", room)); }
            sb.Append(" ORDER BY ClassRoom, StudentCode");

            var dt = DB.Query(sb.ToString(), null, ps.ToArray());
            _dg.ItemsSource = dt.DefaultView;

            int subjects = DB.ScalarInt(@"
                SELECT COUNT(*) FROM Subjects
                WHERE GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel=''",
                null, ("@g", grade));

            string roomLabel = room == "ທັງໝົດ" ? "ທຸກຫ້ອງ" : $"ຫ້ອງ {room}";
            _lbl.Text = $"📚 {dt.Rows.Count} ຄົນ ({grade} · {roomLabel} · ປີ {year})   ·   {subjects} ວິຊາ × 2 ພາກ";
            _btnEnroll.IsEnabled = dt.Rows.Count > 0 && subjects > 0;
        }

        // One transaction; for every loaded student × every grade subject × Sem 1+2,
        // INSERT OR IGNORE. UNIQUE(StudentID,SubjectID,AcademicYear,Semester) makes
        // duplicate-skip free, so re-running this button is safe.
        private void EnrollAll()
        {
            if (_dg.ItemsSource is not DataView dv || dv.Count == 0)
            {
                MessageBox.Show("ບໍ່ມີນັກຮຽນທີ່ຕົງກັບຕົວກອງ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            string year  = _year.SelectedItem?.ToString()  ?? DB.CurrentYear;
            string grade = _grade.SelectedItem?.ToString() ?? "ມ.4";
            string room  = _room.SelectedItem?.ToString()  ?? "ທັງໝົດ";

            var subjects = DB.Query(@"
                SELECT SubjectID FROM Subjects
                WHERE GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel=''
                ORDER BY SortOrder, SubjectCode",
                null, ("@g", grade));
            if (subjects.Rows.Count == 0)
            {
                MessageBox.Show($"ບໍ່ມີວິຊາທີ່ກຳນົດໄວ້ສຳລັບຊັ້ນ {grade}\n" +
                                "ກະລຸນາເພີ່ມວິຊາໃນໜ້າ ‘ວິຊາ’ ກ່ອນ",
                    "ບໍ່ມີຂໍ້ມູນ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string roomLabel = room == "ທັງໝົດ" ? "ທຸກຫ້ອງ" : $"ຫ້ອງ {room}";
            int studentCount = dv.Count;
            if (MessageBox.Show(
                    $"ລົງທະບຽນທຸກນັກຮຽນ:\n\n" +
                    $"   ຊັ້ນ: {grade}   ·   {roomLabel}   ·   ປີ: {year}\n" +
                    $"   ນັກຮຽນ: {studentCount} ຄົນ\n" +
                    $"   ວິຊາ: {subjects.Rows.Count} ວິຊາ × 2 ພາກ\n\n" +
                    "ດຳເນີນການຕໍ່?",
                    "ຢືນຢັນ", MessageBoxButton.YesNo, MessageBoxImage.Question)
                != MessageBoxResult.Yes) return;

            int studentsEnrolled = 0, studentsAlreadyComplete = 0;
            int rowsInserted = 0, rowsSkipped = 0;
            using var conn = DB.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                foreach (DataRowView dr in dv)
                {
                    int sid = Convert.ToInt32(dr["StudentID"]);
                    int newForThisStudent = 0;
                    foreach (DataRow sr in subjects.Rows)
                    {
                        int subId = Convert.ToInt32(sr["SubjectID"]);
                        for (int sem = 1; sem <= 2; sem++)
                        {
                            int n = DB.ExecTx(
                                @"INSERT OR IGNORE INTO Enrollments(StudentID, SubjectID, AcademicYear, Semester)
                                  VALUES(@s, @sub, @y, @sm)",
                                conn, tx,
                                ("@s", sid), ("@sub", subId), ("@y", year), ("@sm", sem));
                            if (n > 0) { rowsInserted++; newForThisStudent++; }
                            else       { rowsSkipped++; }
                        }
                    }
                    if (newForThisStudent > 0) studentsEnrolled++;
                    else                       studentsAlreadyComplete++;
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ລົງທະບຽນບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DB.Log("BatchEnroll",
                $"{grade}/{roomLabel} {year} students={studentCount} newRows={rowsInserted} skipped={rowsSkipped}");
            _lbl.Text = $"✅ {studentsEnrolled} ໃໝ່ · {studentsAlreadyComplete} ມີຢູ່ແລ້ວ · " +
                        $"ບັນທຶກ {rowsInserted} ແຖວ — {DateTime.Now:HH:mm:ss}";
            MessageBox.Show(
                $"ລົງທະບຽນສຳເລັດ:\n\n" +
                $"   ນັກຮຽນທີ່ໄດ້ບັນທຶກວິຊາໃໝ່: {studentsEnrolled} ຄົນ\n" +
                $"   ມີຢູ່ແລ້ວທັງໝົດ (ບໍ່ໄດ້ບັນທຶກໃໝ່): {studentsAlreadyComplete} ຄົນ\n\n" +
                $"   ແຖວ Enrollments ໃໝ່: {rowsInserted}\n" +
                $"   ແຖວ Enrollments ທີ່ຂ້າມ (ຊ້ຳ): {rowsSkipped}\n\n" +
                $"(ແຕ່ລະວິຊາລົງທະບຽນທັງສອງພາກໂດຍອັດຕະໂນມັດ)",
                "ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    public class ComboItem{public string Tag{get;}public string Display{get;}public ComboItem(string t,string d){Tag=t;Display=d;}}
    public class SubComboItem{
        public int Id{get;}
        public string Code{get;}                  // SubjectCode — used to detect CHA1/LAB1 without a DB query
        private readonly string _d;
        public SubComboItem(int id,string d){ Id=id; Code=""; _d=d; }
        public SubComboItem(int id,string code,string d){ Id=id; Code=code; _d=d; }
        public override string ToString()=>_d;
    }

    // ════════════════════════════════════════════════════════════
    //  CLASS HUB — class-first navigation (idx 1)
    //
    //  Step 1: pick ມ.1 / ມ.2 / ມ.3 / ມ.4 (each card shows total /
    //          active / graduated counts + the rooms in use)
    //  Step 2: pick the ຫ້ອງ + ປີ filters, see a live stats row
    //          (students, active, subjects enrolled, current sem),
    //          then jump to score-entry / management / reports.
    //
    //  Every action button sets DB.NavGrade/Room/Year/Month/Semester
    //  and calls MainWindow.NavigateToIndex(...) so the destination
    //  page picks the context up on load (ClearNav wipes it after).
    //
    //  Styling: uses the app-wide Card / SectionTitle / PageTitle
    //  styles and the palette brushes (PrimaryBrush, SuccessBrush …)
    //  from Styles/Controls.xaml + Styles/Colors.xaml so the page
    //  matches every other page. No hard-coded RGB.
    // ════════════════════════════════════════════════════════════
    public class ClassHubPage : UserControl
    {
        private static readonly string[] AllGrades = { "ມ.1", "ມ.2", "ມ.3", "ມ.4" };
        private static readonly string[] AllRooms  = { "1", "2", "3", "4", "5", "6" };

        private Grid      _classSelectView = null!;
        private Grid      _hubView         = null!;
        private TextBlock _hubBreadcrumb   = null!;
        private ComboBox  _hubRoom         = null!;
        private ComboBox  _hubYear         = null!;
        private TextBlock _statTotal       = null!;
        private TextBlock _statActive      = null!;
        private TextBlock _statSubjects    = null!;
        private TextBlock _statSem         = null!;
        private string    _selectedGrade   = "";

        public ClassHubPage()
        {
            var root = new Grid();
            _classSelectView = BuildClassSelectView();
            _hubView         = BuildHubView();
            root.Children.Add(_classSelectView);
            root.Children.Add(_hubView);
            Content = root;
            ShowClassSelect();
        }

        // ─── Style + brush helpers ───────────────────────────────
        // Pull styles/brushes from the merged App resources so the page
        // matches every other page — no locally-defined colours.
        private static Border StyledCard(Thickness? margin = null, Thickness? padding = null)
        {
            var b = new Border();
            b.SetResourceReference(Border.StyleProperty, "Card");
            if (margin.HasValue)  b.Margin  = margin.Value;
            if (padding.HasValue) b.Padding = padding.Value;
            return b;
        }
        private static TextBlock Styled(string styleKey, string text)
        {
            var t = new TextBlock { Text = text };
            t.SetResourceReference(TextBlock.StyleProperty, styleKey);
            return t;
        }
        private static System.Windows.Media.SolidColorBrush Brush(string key) =>
            (System.Windows.Media.SolidColorBrush)Application.Current.FindResource(key);

        // ── View 1: Class-selection landing ──────────────────────
        private Grid BuildClassSelectView()
        {
            var g = new Grid();
            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var stack  = new StackPanel();

            // Header card: title + subtitle (left) · current-year badge (right).
            var headerCard = StyledCard(new Thickness(0, 0, 0, 16), new Thickness(20, 14, 20, 14));
            var headerDock = new DockPanel();
            var yearBadge  = new Border {
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(10, 4, 10, 4),
                Background   = Brush("PrimaryLightBrush"),
                VerticalAlignment = VerticalAlignment.Center
            };
            yearBadge.Child = new TextBlock {
                Text = $"📅  ປີການສຶກສາ {DB.CurrentYear}",
                FontSize = 12, FontWeight = FontWeights.SemiBold,
                Foreground = Brush("PrimaryDarkBrush")
            };
            DockPanel.SetDock(yearBadge, Dock.Right);
            headerDock.Children.Add(yearBadge);

            var headerText = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            headerText.Children.Add(Styled("PageTitle", "🏫  ໜ້າຫ້ອງຮຽນ"));
            headerText.Children.Add(new TextBlock {
                Text = "ເລືອກຊັ້ນ ເພື່ອເຂົ້າສູ່ໜ້າບໍລິຫານຫ້ອງຮຽນ ບັນທຶກຄະແນນ ແລະ ອອກລາຍງານ",
                FontSize = 12, Margin = new Thickness(0, 4, 0, 0),
                Foreground = Brush("TextSecondaryBrush")
            });
            headerDock.Children.Add(headerText);
            headerCard.Child = headerDock;
            stack.Children.Add(headerCard);

            // Grade cards — one column each, live counts per grade.
            var grid = new UniformGrid { Rows = 1, Columns = AllGrades.Length };
            foreach (var grade in AllGrades)
                grid.Children.Add(BuildGradeCard(grade));
            stack.Children.Add(grid);

            // Footer hint.
            stack.Children.Add(new TextBlock {
                Text = "💡  ໂຮງຮຽນຮອງຮັບສະເພາະຊັ້ນ ມ.1 – ມ.4",
                FontSize = 11, Margin = new Thickness(0, 20, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = Brush("TextMutedBrush")
            });

            scroll.Content = stack;
            g.Children.Add(scroll);
            return g;
        }

        // Grade card: total students, active + graduated pills, rooms in use.
        // Rendered as a Button so click-anywhere navigates; hover swaps the
        // border to PrimaryBrush.
        private Button BuildGradeCard(string grade)
        {
            int total     = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE GradeLevel=@g", null, ("@g", grade));
            int active    = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND Status='ກຳລັງຮຽນ'", null, ("@g", grade));
            int graduated = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND Status='ຈົບ'",       null, ("@g", grade));
            var rooms = DB.Query(@"
                SELECT DISTINCT ClassRoom FROM Students
                WHERE GradeLevel=@g AND ClassRoom IS NOT NULL AND ClassRoom<>''
                  AND Status='ກຳລັງຮຽນ'
                ORDER BY ClassRoom",
                null, ("@g", grade));
            string roomList = rooms.Rows.Count > 0
                ? "ຫ້ອງ  " + string.Join(" · ", rooms.AsEnumerable().Select(r => r[0]?.ToString() ?? ""))
                : "ຍັງບໍ່ມີຫ້ອງ";

            var card = new Border {
                Background       = Brush("BgCardBrush"),
                BorderBrush      = Brush("BorderBrush_"),
                BorderThickness  = new Thickness(1),
                CornerRadius     = new CornerRadius(4),
                Padding          = new Thickness(18, 18, 18, 16)
            };
            var stack = new StackPanel();

            // Top row: grade label + icon badge on the right.
            var top   = new DockPanel { LastChildFill = true };
            var badge = new Border {
                Width = 42, Height = 42, CornerRadius = new CornerRadius(21),
                Background = Brush("PrimaryLightBrush"),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            badge.Child = new TextBlock {
                Text = "🏫", FontSize = 20,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            DockPanel.SetDock(badge, Dock.Right);
            top.Children.Add(badge);
            top.Children.Add(new TextBlock {
                Text = grade, FontSize = 30, FontWeight = FontWeights.Bold,
                Foreground = Brush("PrimaryBrush"),
                VerticalAlignment = VerticalAlignment.Center
            });
            stack.Children.Add(top);

            // Big total count.
            stack.Children.Add(new TextBlock {
                Text = "ຈຳນວນນັກຮຽນ",
                FontSize = 11, Margin = new Thickness(0, 14, 0, 4),
                Foreground = Brush("TextSecondaryBrush")
            });
            var countRow = new StackPanel { Orientation = Orientation.Horizontal };
            countRow.Children.Add(new TextBlock {
                Text = total.ToString(), FontSize = 26, FontWeight = FontWeights.Bold,
                Foreground = Brush("TextPrimaryBrush"),
                VerticalAlignment = VerticalAlignment.Bottom
            });
            countRow.Children.Add(new TextBlock {
                Text = " ຄົນ", FontSize = 12, Margin = new Thickness(4, 0, 0, 4),
                Foreground = Brush("TextSecondaryBrush"),
                VerticalAlignment = VerticalAlignment.Bottom
            });
            stack.Children.Add(countRow);

            // Status pills.
            var pills = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 0) };
            pills.Children.Add(MkPill($"ກຳລັງຮຽນ  {active}",   "SuccessBrush"));
            if (graduated > 0)
                pills.Children.Add(MkPill($"ຈົບ  {graduated}", "NeutralBrush"));
            stack.Children.Add(pills);

            // Rooms line.
            stack.Children.Add(new TextBlock {
                Text = roomList, FontSize = 11, Margin = new Thickness(0, 12, 0, 0),
                Foreground = Brush("TextMutedBrush"), TextWrapping = TextWrapping.Wrap
            });

            // CTA line.
            stack.Children.Add(new TextBlock {
                Text = "ກົດເພື່ອເປີດ →", FontSize = 12,
                Margin = new Thickness(0, 14, 0, 0), FontWeight = FontWeights.SemiBold,
                Foreground = Brush("PrimaryBrush")
            });

            card.Child = stack;

            // Button wrapper — transparent chrome so the Border above owns the look.
            var btn = new Button {
                Margin          = new Thickness(6),
                Cursor          = System.Windows.Input.Cursors.Hand,
                BorderThickness = new Thickness(0),
                Background      = System.Windows.Media.Brushes.Transparent,
                Padding         = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalContentAlignment   = VerticalAlignment.Stretch,
                Content         = card,
                Template        = MkTransparentButtonTemplate()
            };
            btn.MouseEnter += (s, e) => card.BorderBrush = Brush("PrimaryBrush");
            btn.MouseLeave += (s, e) => card.BorderBrush = Brush("BorderBrush_");
            btn.Click      += (s, e) => ShowHub(grade);
            return btn;
        }

        // Bare button template so the card border isn't fought by the
        // default button chrome (blue focus rectangle, default background).
        private static ControlTemplate MkTransparentButtonTemplate()
        {
            var tpl = new ControlTemplate(typeof(Button));
            var cp  = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,   VerticalAlignment.Stretch);
            tpl.VisualTree = cp;
            return tpl;
        }

        private static Border MkPill(string text, string brushKey)
        {
            var brush = Brush(brushKey);
            var soft  = System.Windows.Media.Color.FromArgb(30, brush.Color.R, brush.Color.G, brush.Color.B);
            var pill  = new Border {
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(8, 2, 8, 2),
                Background   = new System.Windows.Media.SolidColorBrush(soft),
                Margin       = new Thickness(0, 0, 6, 0)
            };
            pill.Child = new TextBlock {
                Text = text, FontSize = 11, FontWeight = FontWeights.SemiBold, Foreground = brush
            };
            return pill;
        }

        // ── View 2: Per-class hub ────────────────────────────────
        private Grid BuildHubView()
        {
            var g      = new Grid();
            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var stack  = new StackPanel();

            // ── Header card: back button · breadcrumb · ຫ້ອງ+ປີ filters ──
            var headerCard = StyledCard(new Thickness(0, 0, 0, 12), new Thickness(14, 10, 14, 10));
            var headerDock = new DockPanel();
            var back = H.Btn("◀  ກັບໄປ", "NeutralButton");
            back.Click += (s, e) => ShowClassSelect();
            DockPanel.SetDock(back, Dock.Left);
            headerDock.Children.Add(back);

            var rightFilters = new StackPanel {
                Orientation         = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center
            };
            rightFilters.Children.Add(H.Lbl("ຫ້ອງ:"));
            _hubRoom = H.MkCmb(AllRooms, 70);
            _hubRoom.SelectionChanged += (s, e) => RefreshHubStats();
            rightFilters.Children.Add(_hubRoom);
            rightFilters.Children.Add(H.Lbl("ປີ:"));
            _hubYear = H.MkCmb(DB.AcademicYears().ToArray(), 110);
            SelectComboValue(_hubYear, DB.CurrentYear);
            _hubYear.SelectionChanged += (s, e) => RefreshHubStats();
            rightFilters.Children.Add(_hubYear);
            DockPanel.SetDock(rightFilters, Dock.Right);
            headerDock.Children.Add(rightFilters);

            _hubBreadcrumb = new TextBlock {
                FontSize = 15, FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0),
                Foreground = Brush("TextPrimaryBrush")
            };
            headerDock.Children.Add(_hubBreadcrumb);
            headerCard.Child = headerDock;
            stack.Children.Add(headerCard);

            // ── Stats row: live counts for the picked (grade, room, year) ──
            _statTotal    = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            _statActive   = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            _statSubjects = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            _statSem      = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            var statsRow  = new UniformGrid { Rows = 1, Columns = 4, Margin = new Thickness(0, 0, 0, 12) };
            statsRow.Children.Add(BuildStatCard("ນັກຮຽນທັງໝົດ", "👥", _statTotal,    "PrimaryBrush",   "#D6E4F7", new Thickness(0, 0, 8, 0)));
            statsRow.Children.Add(BuildStatCard("ກຳລັງຮຽນ",     "✅", _statActive,   "SuccessBrush",   "#DCFCE7", new Thickness(4, 0, 4, 0)));
            statsRow.Children.Add(BuildStatCard("ວິຊາລົງທະບຽນ", "📚", _statSubjects, "SecondaryBrush", "#EDE3FB", new Thickness(4, 0, 4, 0)));
            statsRow.Children.Add(BuildStatCard("ພາກປັດຈຸບັນ",  "📅", _statSem,      "InfoBrush",      "#CFFAFE", new Thickness(8, 0, 0, 0)));
            stack.Children.Add(statsRow);

            // ── Two-column action panels ──
            var twoCol = new Grid();
            twoCol.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            twoCol.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left column — score entry (monthly · semester · annual).
            var leftCard  = StyledCard(new Thickness(0, 0, 6, 0), new Thickness(16));
            var leftStack = new StackPanel();
            leftStack.Children.Add(Styled("SectionTitle", "📅  ບັນທຶກຄະແນນປະຈຳເດືອນ"));
            leftStack.Children.Add(SectionSubheader("ພາກ 1 — ກ.ຍ. ຫາ ທ.ວ."));
            AddMonthButton(leftStack, "ກ.ຍ. (09)",  9);
            AddMonthButton(leftStack, "ຕ.ລ. (10)", 10);
            AddMonthButton(leftStack, "ພ.ຍ. (11)", 11);
            AddMonthButton(leftStack, "ທ.ວ. (12)", 12);

            leftStack.Children.Add(new Border { Height = 8 });
            leftStack.Children.Add(SectionSubheader("ພາກ 2 — ກ.ພ. ຫາ ພ.ພ."));
            AddMonthButton(leftStack, "ກ.ພ. (02)", 2);
            AddMonthButton(leftStack, "ມີ.ນ. (03)", 3);
            AddMonthButton(leftStack, "ມ.ສ. (04)", 4);
            AddMonthButton(leftStack, "ພ.ພ. (05)", 5);

            leftStack.Children.Add(new Border { Height = 16 });
            leftStack.Children.Add(Styled("SectionTitle", "📝  ບັນທຶກຄະແນນພາກຮຽນ / ປະຈຳປີ"));
            AddSemButton(leftStack,    "📘  ຄະແນນສອບເສງ ພາກ 1", 1);
            AddSemButton(leftStack,    "📗  ຄະແນນສອບເສງ ພາກ 2", 2);
            AddAnnualButton(leftStack, "📕  ຄະແນນປະຈຳປີ (CHA1 / LAB1)");

            leftCard.Child = leftStack;
            Grid.SetColumn(leftCard, 0);
            twoCol.Children.Add(leftCard);

            // Right column — class management + reports/history.
            var rightCard  = StyledCard(new Thickness(6, 0, 0, 0), new Thickness(16));
            var rightStack = new StackPanel();
            rightStack.Children.Add(Styled("SectionTitle", "👥  ບໍລິຫານຫ້ອງຮຽນ"));
            AddActionButton(rightStack, "👥  ຂໍ້ມູນນັກຮຽນຂອງຫ້ອງ",  "PrimaryButton", GoToStudents);
            AddActionButton(rightStack, "📋  ລົງທະບຽນວິຊາ (Batch)", "PrimaryButton", GoToBatchEnroll);

            rightStack.Children.Add(new Border { Height = 16 });
            rightStack.Children.Add(Styled("SectionTitle", "📊  ລາຍງານ & ປະຫວັດ"));
            AddActionButton(rightStack, "📚  ປະຫວັດຄະແນນທັງຫ້ອງ",   "SecondaryButton", GoToScoreHistory);
            AddActionButton(rightStack, "📄  ໃບຄະແນນ ນັກຮຽນ",       "SecondaryButton", GoToScores);
            AddActionButton(rightStack, "📑  ໃບສັນຍາ / ປະຫວັດ",    "NeutralButton",   GoToProfileReports);

            rightCard.Child = rightStack;
            Grid.SetColumn(rightCard, 1);
            twoCol.Children.Add(rightCard);

            stack.Children.Add(twoCol);
            scroll.Content = stack;
            g.Children.Add(scroll);
            return g;
        }

        // Dashboard-style stat card: label · big number (accent brush) · icon badge.
        private static Border BuildStatCard(string label, string icon, TextBlock numberTb, string accentBrushKey, string iconBg, Thickness margin)
        {
            var card = StyledCard(margin, new Thickness(16, 14, 16, 14));
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var left = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            left.Children.Add(new TextBlock {
                Text = label, FontSize = 11, FontWeight = FontWeights.Medium,
                Foreground = Brush("TextSecondaryBrush")
            });
            numberTb.Foreground = Brush(accentBrushKey);
            numberTb.Margin     = new Thickness(0, 4, 0, 0);
            left.Children.Add(numberTb);
            Grid.SetColumn(left, 0);
            grid.Children.Add(left);

            var iconBadge = new Border {
                Width = 40, Height = 40, CornerRadius = new CornerRadius(20),
                Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter().ConvertFromString(iconBg)!,
                VerticalAlignment = VerticalAlignment.Top
            };
            iconBadge.Child = new TextBlock {
                Text = icon, FontSize = 18,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            Grid.SetColumn(iconBadge, 1);
            grid.Children.Add(iconBadge);

            card.Child = grid;
            return card;
        }

        private static TextBlock SectionSubheader(string text) =>
            new TextBlock {
                Text = text, FontSize = 11, FontWeight = FontWeights.SemiBold,
                Foreground = Brush("TextSecondaryBrush"),
                Margin = new Thickness(0, 0, 0, 6)
            };

        private void AddMonthButton(StackPanel host, string label, int month)
        {
            var btn = H.Btn($"📝  ຄະແນນ ເດືອນ {label}", "PrimaryButton");
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => GoToMonthly(month);
            host.Children.Add(btn);
        }

        private void AddSemButton(StackPanel host, string label, int sem)
        {
            var btn = H.Btn(label, "SuccessButton");
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => GoToSemesterEntry(sem);
            host.Children.Add(btn);
        }

        private void AddAnnualButton(StackPanel host, string label)
        {
            var btn = H.Btn(label, "WarningButton");
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => GoToAnnualEntry();
            host.Children.Add(btn);
        }

        private void AddActionButton(StackPanel host, string label, string style, Action onClick)
        {
            var btn = H.Btn(label, style);
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => onClick();
            host.Children.Add(btn);
        }

        private static void SelectComboValue(ComboBox cb, string value)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == value) { cb.SelectedIndex = i; return; }
        }

        // ── View state ───────────────────────────────────────────
        private void ShowClassSelect()
        {
            _classSelectView.Visibility = Visibility.Visible;
            _hubView.Visibility         = Visibility.Collapsed;
        }
        private void ShowHub(string grade)
        {
            _selectedGrade = grade;
            _classSelectView.Visibility = Visibility.Collapsed;
            _hubView.Visibility         = Visibility.Visible;
            RefreshHubStats();
        }

        // Refresh the 4 stat cards + breadcrumb whenever room/year changes.
        private void RefreshHubStats()
        {
            string grade = _selectedGrade;
            string room  = _hubRoom.SelectedItem?.ToString() ?? "1";
            string year  = _hubYear.SelectedItem?.ToString() ?? DB.CurrentYear;

            int total = DB.ScalarInt(
                "SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND ClassRoom=@r AND AcademicYear=@y",
                null, ("@g", grade), ("@r", room), ("@y", year));
            int active = DB.ScalarInt(
                "SELECT COUNT(*) FROM Students WHERE GradeLevel=@g AND ClassRoom=@r AND AcademicYear=@y AND Status='ກຳລັງຮຽນ'",
                null, ("@g", grade), ("@r", room), ("@y", year));
            int subjects = DB.ScalarInt(@"
                SELECT COUNT(DISTINCT e.SubjectID)
                FROM Enrollments e
                JOIN Students   s ON s.StudentID = e.StudentID
                WHERE s.GradeLevel=@g AND s.ClassRoom=@r AND e.AcademicYear=@y",
                null, ("@g", grade), ("@r", room), ("@y", year));

            _statTotal.Text    = total.ToString();
            _statActive.Text   = active.ToString();
            _statSubjects.Text = subjects.ToString();
            _statSem.Text      = DB.CurrentSem.ToString();

            _hubBreadcrumb.Text = $"ໜ້າຫ້ອງຮຽນ  /  ຊັ້ນ {grade}  ·  ຫ້ອງ {room}  ·  ປີ {year}";
        }

        // ── Nav-context handoff ──────────────────────────────────
        private void SetNavGradeRoomYear()
        {
            DB.NavGrade = _selectedGrade;
            DB.NavRoom  = _hubRoom.SelectedItem?.ToString() ?? "1";
            DB.NavYear  = _hubYear.SelectedItem?.ToString() ?? DB.CurrentYear;
        }

        private MainWindow? Main => Window.GetWindow(this) as MainWindow;

        private void GoToMonthly(int month)
        {
            SetNavGradeRoomYear();
            DB.NavMonth = month;
            Main?.NavigateToIndex(6, $"ຄະແນນ ເດືອນ {month:D2} — {_selectedGrade}/{DB.NavRoom}");
        }

        // Sem 1/2 exam entry → ScoresPage (idx 5). NavSemester picked up on load.
        private void GoToSemesterEntry(int sem)
        {
            SetNavGradeRoomYear();
            DB.NavSemester = sem;
            Main?.NavigateToIndex(5, $"ຄະແນນສອບເສງ ພາກ {sem} — {_selectedGrade}/{DB.NavRoom}");
        }

        // Annual CHA1/LAB1 → ScoresPage (idx 5). Sentinel value 3 tells the
        // page to pick CmbSem.SelectedIndex=2 (ປະຈຳປີ) on load.
        private void GoToAnnualEntry()
        {
            SetNavGradeRoomYear();
            DB.NavSemester = 3;
            Main?.NavigateToIndex(5, $"ຄະແນນປະຈຳປີ (CHA/LAB) — {_selectedGrade}/{DB.NavRoom}");
        }

        private void GoToStudents()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(2, $"ນັກຮຽນ — {_selectedGrade}/{DB.NavRoom}");
        }
        private void GoToBatchEnroll()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(4, $"ລົງທະບຽນ (Batch) — {_selectedGrade}");
        }
        // Score-history page (idx 13) — every class-wide score report + student
        // history browser lives here (all score reports migrated off ReportPage).
        private void GoToScoreHistory()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(13, $"ປະຫວັດຄະແນນ — {_selectedGrade}/{DB.NavRoom}");
        }
        private void GoToScores()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(5, $"ບັນທຶກຄະແນນ — {_selectedGrade}/{DB.NavRoom}");
        }
        // ReportPage (idx 7) now holds only Enrollment Agreement + Student
        // Profile — the score reports moved to Score History. Keep this button
        // so the two remaining reports are one click from the class hub.
        private void GoToProfileReports()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(7, $"ໃບສັນຍາ / ປະຫວັດ — {_selectedGrade}/{DB.NavRoom}");
        }
    }

    // ════════════════════════════════════════════════════════════
    //  MONTHLY ASSESSMENT (class roster: Activity 3 + Discipline 2 + Homework 5 = /10)
    // ════════════════════════════════════════════════════════════
    public class MonthlyScoresPage : UserControl
    {
        private ComboBox _grade=null!,_room=null!,_subject=null!,_year=null!,_month=null!,_status=null!;
        private DataGrid _dg=null!;
        private TextBlock _info=null!;
        private readonly ObservableCollection<MonthlyRow> _rows=new();
        // Grid columns we toggle when switching between academic and eval display.
        // _colEval (single /10 column) shows when the picked subject is CHA1/LAB1;
        // the 3 sub-score columns + _colTotal hide in eval display.
        private DataGridColumn _colActivity=null!,_colDiscipline=null!,_colHomework=null!,_colTotal=null!,_colEval=null!;
        private TextBlock _lblMonth=null!,_hint=null!;
        private Button   _btnSave=null!;
        private const string MonthlyHint =
            "ສູດ: ຮ່ວມຮຽນ (/3) + ກິດຈະກຳ (/2) + ກວດກາ (/5) = ຄະແນນປະຈຳເດືອນ (/10)   |   ພາກ 1: ກ.ຍ.–ທ.ວ. (ສອບເສັງພາກ1 ມ.ກ.)   |   ພາກ 2: ກ.ພ.–ພ.ພ. (ສອບເສັງພາກຮຽນ2 ມິ.ຖ.)   |   ຄະແນນປະຈຳປີ = ສະເລ່ຍ(4 ເດືອນ) × 50% + ສອບເສັງພາກຮຽນ × 50%";
        private const string EvalHint =
            "ຄຸນສົມບັດ (CHA1) ແລະ ການອອກແຮງງານ (LAB1): ປ້ອນດ້ວຍຕົນເອງ (0–10) ສຳລັບແຕ່ລະເດືອນ. ບໍ່ມີການຄິດໄລ່ອັດຕະໂນມັດ — ບໍ່ຖືກລວມເຂົ້າສະເລ່ຍ ຫຼື ການຈັດອັນດັບ.";

        // Selected SubjectCode (cached on SubComboItem so we don't re-query the DB).
        private string SelectedSubjectCode =>
            (_subject?.SelectedItem as SubComboItem)?.Code ?? "";
        // CHA1/LAB1 always route through EvaluationScores (never auto-calculated).
        // The grid swaps to the single /10 EvalScore column for these subjects.
        private bool IsEvalMode =>
            SelectedSubjectCode == "CHA1" || SelectedSubjectCode == "LAB1";
        // EvaluationScores context for the picked month — only meaningful when
        // the picked subject is CHA1/LAB1. Each Month{N} is its own context.
        private string EvalContext =>
            IsEvalMode && _month?.SelectedItem is MonthItem mi
                ? (DB.MonthContextName(mi.Value) ?? "")
                : "";
        // Track what the currently-loaded roster represents, so Save can refuse
        // to write monthly data into the wrong month/subject/year/class.
        private int    _loadedMonth       = 0;
        private int    _loadedSubjectId   = 0;
        private string _loadedSubjectCode = "";  // stamped so Save writes to the loaded subject
        private string _loadedYear        = "";
        private string _loadedGrade       = "";
        private string _loadedRoom        = "";
        private string _loadedContext     = "";  // "" academic-monthly | Month1..Month8 (CHA1/LAB1)
        // Auto-reload: once the user has done their first manual ໂຫຼດ, any further
        // change to the filter combos reloads automatically — avoids the common
        // mistake of editing the wrong roster after switching month/subject.
        private bool   _autoReload      = true;    // ໂຫຼດ button removed → always auto-reload
        private bool   _dirty           = false;  // unsaved edits indicator

        public MonthlyScoresPage()
        {
            var root=H.MkGrid(GridLength.Auto,GridLength.Auto,new GridLength(1,GridUnitType.Star),GridLength.Auto);

            // ── Filter bar ───────────────────────────────────────
            var bar=H.MkCard(new Thickness(0,0,0,10),new Thickness(14,10,14,10));
            var flow=new WrapPanel();
            _grade=H.MkCmb(new[]{"ມ.1","ມ.2","ມ.3","ມ.4"},80);
            _grade.SelectedIndex=3;
            _grade.SelectionChanged+=(s,e)=>{ RefreshSubjects(); OnFilterChanged(); };
            _room=H.MkCmb(new[]{"1","2","3","4","5","6"},70);
            _room.SelectionChanged+=(s,e)=>OnFilterChanged();
            _subject=new ComboBox{Width=240,Margin=new Thickness(0,0,8,0)};
            // Switching subject in monthly mode can toggle academic ⇄ eval display
            // (CHA1/LAB1 are eval, every other subject is academic). Re-apply
            // column visibility before invalidating the loaded roster.
            _subject.SelectionChanged+=(s,e)=>{ ApplyDisplayMode(); OnFilterChanged(); };
            _year=H.MkCmb(DB.AcademicYears().ToArray(),100);
            _year.SelectionChanged+=(s,e)=>OnFilterChanged();
            _month = new ComboBox { Width = 200, Margin = new Thickness(0,0,8,0) };
            // Monthly assessment windows ONLY — final-exam months (Jan / June) are
            // handled on the ‘ບັນທຶກຄະແນນ’ page via the FinalScore column.
            //   Sem 1 monthly: Sept · Oct · Nov · Dec
            //   Sem 2 monthly: Feb  · Mar · Apr · May
            string[] monthLabels = {
                "ກ.ຍ. (09) — ພາກ 1",
                "ຕ.ລ. (10) — ພາກ 1",
                "ພ.ຍ. (11) — ພາກ 1",
                "ທ.ວ. (12) — ພາກ 1",
                "ກ.ພ. (02) — ພາກ 2",
                "ມີ.ນ. (03) — ພາກ 2",
                "ມ.ສ. (04) — ພາກ 2",
                "ພ.ພ. (05) — ພາກ 2",
            };
            int[] monthValues = { 9, 10, 11, 12, 2, 3, 4, 5 };
            for (int i = 0; i < monthLabels.Length; i++)
                _month.Items.Add(new MonthItem(monthValues[i], monthLabels[i]));
            _month.SelectedIndex = 0;
            _month.SelectionChanged += (s,e) => OnFilterChanged();
            // Status filter — defaults to ‘ກຳລັງຮຽນ’ so today's behaviour is unchanged.
            // Switch to ‘ຈົບ’ or ‘ທັງໝົດ’ to look up a graduate's historical scores.
            _status = H.MkCmb(new[]{"ກຳລັງຮຽນ","ຈົບ","ອອກ","ທັງໝົດ"}, 110);
            _status.SelectedIndex = 0;
            _status.SelectionChanged += (s,e) => OnFilterChanged();
            _btnSave=H.Btn("💾  ບັນທຶກ (Ctrl+S)","SuccessButton"); _btnSave.Click+=(s,e)=>SaveAll();
            // Excel import flow — see ImportPreviewWindow for the validation UI.
            // Download generates a template scoped to (year, grade, room, month);
            // Import opens an .xlsx, validates against that same scope, then
            // shows the preview window where the teacher confirms before saving.
            var bDownload = H.Btn("📥  ດາວໂຫຼດແບບຟອມ", "SecondaryButton");
            bDownload.Click += (s,e) => DownloadTemplate();
            var bImport   = H.Btn("📤  ນຳເຂົ້າ Excel",  "PrimaryButton");
            bImport.Click   += (s,e) => ImportExcel();
            _info=new TextBlock{VerticalAlignment=VerticalAlignment.Center,Margin=new Thickness(12,0,0,0),FontSize=12,Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))};
            _lblMonth = H.Lbl("ເດືອນ:");
            flow.Children.Add(H.Lbl("ຊັ້ນ:"));   flow.Children.Add(_grade);
            flow.Children.Add(H.Lbl("ຫ້ອງ:"));  flow.Children.Add(_room);
            flow.Children.Add(H.Lbl("ສະຖານະ:")); flow.Children.Add(_status);
            flow.Children.Add(H.Lbl("ວິຊາ:"));   flow.Children.Add(_subject);
            flow.Children.Add(H.Lbl("ປີ:"));    flow.Children.Add(_year);
            flow.Children.Add(_lblMonth);        flow.Children.Add(_month);
            flow.Children.Add(_btnSave);
            flow.Children.Add(bDownload); flow.Children.Add(bImport);
            flow.Children.Add(_info);
            bar.Child=flow; Grid.SetRow(bar,0); root.Children.Add(bar);

            // ── Formula hint ──────────────────────────────────────
            var hint=new Border{Background=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239,246,255)),
                                BorderBrush=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(214,228,247)),
                                BorderThickness=new Thickness(1),
                                CornerRadius=new CornerRadius(4),
                                Padding=new Thickness(12,7,12,7),Margin=new Thickness(0,0,0,10)};
            _hint=new TextBlock{
                Text=MonthlyHint,
                FontSize=12,
                Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(27,79,138)),
                TextWrapping=TextWrapping.Wrap};
            hint.Child=_hint;
            Grid.SetRow(hint,1); root.Children.Add(hint);

            // ── Data grid ─────────────────────────────────────────
            _dg=new DataGrid{
                AutoGenerateColumns=false,IsReadOnly=false,CanUserAddRows=false,
                BorderThickness=new Thickness(0),Background=System.Windows.Media.Brushes.White
            };
            _dg.Columns.Add(H.Col("ລະຫັດ","StudentCode",110,true));
            _dg.Columns.Add(H.ColStar("ຊື່ນັກຮຽນ","FullName",true));
            _dg.Columns.Add(H.Col("ຫ້ອງ","ClassRoom",60,true));
            // Direct text-entry score columns — INTEGERS ONLY (0-10 per column max).
            // Decimal separators (.,), letters, and out-of-range values are blocked at
            // keystroke time by IntOnly_PreviewTextInput (wired in OnPreparingCellForEdit)
            // and validated on commit via DB.TryParseIntScore. Invalid input restores
            // the previous value (CancelEdit) and shows a friendly message.
            //   Activity/3 · Discipline/2 · Homework/5 · Eval/10 — per-column max.
            // Shared cell styles from Controls.xaml — bigger, centred, bold digits
            // that make the score cells easy to read/type while grading a class.
            // `editableCellStyle` also tints the cell background soft-blue so
            // teachers immediately see which columns expect input (vs read-only
            // ‘ລວມ’ which stays row-coloured).
            var displayStyle      = (Style)FindResource("ScoreDisplayText");
            var editStyle         = (Style)FindResource("ScoreEditBox");
            var editableCellStyle = (Style)FindResource("ScoreEditableCell");

            _colActivity = new DataGridTextColumn{
                Header="ຮ່ວມຮຽນ (/3)",
                Binding = new System.Windows.Data.Binding("ActivityScore"){
                    StringFormat="F0",
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged },
                Width = new DataGridLength(105),
                CellStyle = editableCellStyle,
                ElementStyle = displayStyle,
                EditingElementStyle = editStyle
            };
            _dg.Columns.Add(_colActivity);
            _colDiscipline = new DataGridTextColumn{
                Header="ກິດຂະກຳ (/2)",
                Binding = new System.Windows.Data.Binding("DisciplineScore"){
                    StringFormat="F0",
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged },
                Width = new DataGridLength(95),
                CellStyle = editableCellStyle,
                ElementStyle = displayStyle,
                EditingElementStyle = editStyle
            };
            _dg.Columns.Add(_colDiscipline);
            _colHomework = new DataGridTextColumn{
                Header="ກວດກາ (/5)",
                Binding = new System.Windows.Data.Binding("HomeworkScore"){
                    StringFormat="F0",
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged },
                Width = new DataGridLength(110),
                CellStyle = editableCellStyle,
                ElementStyle = displayStyle,
                EditingElementStyle = editStyle
            };
            _dg.Columns.Add(_colHomework);
            _colTotal = new DataGridTextColumn{
                Header="ລວມ (/10)",
                Binding=new System.Windows.Data.Binding("Total"){StringFormat="F0"},
                Width=new DataGridLength(95), IsReadOnly=true,
                ElementStyle = displayStyle
            };
            _dg.Columns.Add(_colTotal);
            // Evaluation-mode column: single manual /10 score for CHA1/LAB1. Hidden
            // by default; shown only when an ‘ສະຫຼຸບ…’ assessment type is selected.
            _colEval = new DataGridTextColumn{
                Header="ຄະແນນ (/10)",
                Binding = new System.Windows.Data.Binding("EvalScore"){
                    StringFormat="F0", TargetNullValue = "",
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged },
                Width = new DataGridLength(120),
                Visibility = Visibility.Collapsed,
                CellStyle = editableCellStyle,
                ElementStyle = displayStyle,
                EditingElementStyle = editStyle
            };
            _dg.Columns.Add(_colEval);
            _dg.CellEditEnding += OnScoreCellEditEnding;
            _dg.PreparingCellForEdit += OnPreparingCellForEdit;
            _dg.ItemsSource=_rows;
            var card=H.MkCard(); card.Child=_dg; Grid.SetRow(card,2); root.Children.Add(card);

            // ── Keyboard shortcuts: F5 reload, Ctrl+S save ──
            // Page-level handler so shortcuts fire whether or not focus is in the grid.
            PreviewKeyDown += (s, e) =>
            {
                if (e.Key == System.Windows.Input.Key.F5)
                {
                    LoadRoster(manual:true); e.Handled = true;
                }
                else if (e.Key == System.Windows.Input.Key.S
                    && (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control)
                       == System.Windows.Input.ModifierKeys.Control)
                {
                    SaveAll(); e.Handled = true;
                }
            };

            // ── Status footer ─────────────────────────────────────
            var foot=new Border{Background=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252)),
                                BorderBrush=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                                BorderThickness=new Thickness(1),
                                CornerRadius=new CornerRadius(4),
                                Padding=new Thickness(12,6,12,6),Margin=new Thickness(0,8,0,0)};
            foot.Child=new TextBlock{
                Text="💡 ບັນທຶກແລ້ວ: ຄະແນນປະຈຳພາກ (Mid) ໃນຕາຕະລາງ ‘ບັນທຶກຄະແນນ’ ຈະຖືກປັບປຸງເປັນສະເລ່ຍຂອງເດືອນທັງໝົດໃນພາກນັ້ນໂດຍອັດຕະໂນມັດ",
                FontSize=11,
                Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99)),
                TextWrapping=TextWrapping.Wrap};
            Grid.SetRow(foot,3); root.Children.Add(foot);

            Content=root;
            RefreshSubjects();

            // Apply NavContext if the user arrived here via ClassHubPage. The hub sets
            // grade/room/year/month; we pre-select those, then auto-load the roster so
            // the teacher lands directly in the data-entry view.
            if (!string.IsNullOrEmpty(DB.NavGrade))
            {
                ApplyNavContext();
                DB.ClearNav();
            }
            else if (_subject.SelectedItem != null)
            {
                // No nav-context: auto-load with the default filters so the
                // teacher sees the class grid immediately (ໂຫຼດ button was removed).
                Dispatcher.InvokeAsync(() => LoadRoster(manual: false));
            }
        }

        private void ApplyNavContext()
        {
            // Grade
            for (int i = 0; i < _grade.Items.Count; i++)
                if (_grade.Items[i]?.ToString() == DB.NavGrade) { _grade.SelectedIndex = i; break; }
            // Room
            if (!string.IsNullOrEmpty(DB.NavRoom))
                for (int i = 0; i < _room.Items.Count; i++)
                    if (_room.Items[i]?.ToString() == DB.NavRoom) { _room.SelectedIndex = i; break; }
            // Year
            if (!string.IsNullOrEmpty(DB.NavYear))
                for (int i = 0; i < _year.Items.Count; i++)
                    if (_year.Items[i]?.ToString() == DB.NavYear) { _year.SelectedIndex = i; break; }
            // Refresh subject list for the (now-selected) grade.
            RefreshSubjects();
            // Month
            if (DB.NavMonth > 0)
                for (int i = 0; i < _month.Items.Count; i++)
                    if (_month.Items[i] is MonthItem mi && mi.Value == DB.NavMonth)
                    { _month.SelectedIndex = i; break; }
            // Auto-load when arrival came from the hub — but only if a subject is available.
            if (_subject.SelectedItem != null)
                Dispatcher.InvokeAsync(() => LoadRoster(manual:true));
        }

        // Called from every filter combo's SelectionChanged. Auto-reload is
        // always on (the ໂຫຼດ button was removed) — if edits are pending, we
        // ask via the 3-button SaveConfirmDialog: ບັນທຶກ (save then reload) ·
        // ບໍ່ບັນທຶກ (discard then reload) · ຍົກເລີກ (stay put).
        private void OnFilterChanged()
        {
            if (_dirty)
            {
                var decision = SaveConfirmDialog.Ask(Window.GetWindow(this));
                switch (decision)
                {
                    case SaveConfirmResult.Save:
                        SaveAll();
                        // SaveAll clears _dirty on success — if it's still set
                        // the save failed (error dialog already shown), so keep
                        // the current filter selection to protect the edits.
                        if (_dirty) return;
                        break;
                    case SaveConfirmResult.DontSave:
                        _dirty = false;
                        break;
                    case SaveConfirmResult.Cancel:
                    default:
                        return;
                }
            }
            if (_autoReload && _subject.SelectedItem != null)
                Dispatcher.InvokeAsync(() => LoadRoster(manual:false));
            else
                InvalidateRoster();
        }

        private void MarkDirty()
        {
            if (_dirty) return;
            _dirty = true;
            if (_info != null)
                _info.Text = "✏ ມີຂໍ້ມູນຍັງບໍ່ໄດ້ບັນທຶກ — ກົດ Ctrl+S ເພື່ອບັນທຶກ";
        }

        // Keystroke-level digit-only filter. Attached fresh each time a cell enters
        // edit mode. Blocks any non-digit character (including `.`, `,`, `-`, letters,
        // symbols) from reaching the TextBox so decimal separators are physically
        // unenterable.
        private void OnPreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.Column == _colActivity || e.Column == _colDiscipline ||
                e.Column == _colHomework || e.Column == _colEval)
            {
                if (e.EditingElement is TextBox tb)
                {
                    tb.PreviewTextInput -= IntOnly_PreviewTextInput;
                    tb.PreviewTextInput += IntOnly_PreviewTextInput;
                }
            }
        }

        private static void IntOnly_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
                if (c < '0' || c > '9') { e.Handled = true; return; }
        }

        // STRICT integer validator on commit. Each column has its own max
        // (Activity/3 · Discipline/2 · Homework/5 · Eval/10). Invalid input
        // (decimals, out-of-range, non-numeric paste) cancels the commit and
        // calls CancelEdit so the cell RESTORES the previous value rather than
        // holding the bad text in edit mode.
        private void OnScoreCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            if (e.Row.Item is not MonthlyRow row) return;
            if (e.EditingElement is TextBox tb)
            {
                int max = 10;
                if      (e.Column == _colActivity)   max = 3;
                else if (e.Column == _colDiscipline) max = 2;
                else if (e.Column == _colHomework)   max = 5;
                else if (e.Column == _colEval)       max = 10;
                else return;  // Read-only columns (StudentCode/FullName/Total) — nothing to validate

                if (!DB.TryParseIntScore(tb.Text, 0, max, out int parsed))
                {
                    MessageBox.Show(
                        $"‘{tb.Text}’ ບໍ່ແມ່ນຄະແນນທີ່ຖືກຕ້ອງ.\n" +
                        $"ກະລຸນາໃສ່ຕົວເລກເຕັມລະຫວ່າງ 0 ຫາ {max} (ບໍ່ມີທົດສະນິຍົມ).",
                        "ຄະແນນຜິດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                    e.Cancel = true;
                    Dispatcher.InvokeAsync(() => _dg.CancelEdit(DataGridEditingUnit.Cell));
                    return;
                }
                tb.Text = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            Dispatcher.InvokeAsync(() =>
            {
                if (!IsEvalMode) row.Recalc();
                MarkDirty();
            });
        }

        private void RefreshSubjects()
        {
            string g = _grade.SelectedItem?.ToString() ?? "ມ.4";
            _subject.Items.Clear();
            // ALL subjects (academic + CHA1 + LAB1). Picking CHA1/LAB1 swaps the
            // grid to the single /10 _colEval column via ApplyDisplayMode.
            foreach (DataRow r in DB.Query(
                @"SELECT SubjectID, SubjectCode, SubjectCode||'  '||SubjectName AS D
                  FROM Subjects
                  WHERE (GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel='')
                  ORDER BY SortOrder",
                null, ("@g", g)).Rows)
                _subject.Items.Add(new SubComboItem(
                    Convert.ToInt32(r["SubjectID"]),
                    r["SubjectCode"].ToString()!,
                    r["D"].ToString()!));
            if (_subject.Items.Count > 0) _subject.SelectedIndex = 0;
        }

        // Toggle grid columns based on whether the picked subject is academic
        // (3 sub-scores) or evaluation-style (single /10 EvalScore for CHA1/LAB1).
        private void ApplyDisplayMode()
        {
            bool eval = IsEvalMode;
            if (_colActivity   != null) _colActivity.Visibility   = eval ? Visibility.Collapsed : Visibility.Visible;
            if (_colDiscipline != null) _colDiscipline.Visibility = eval ? Visibility.Collapsed : Visibility.Visible;
            if (_colHomework   != null) _colHomework.Visibility   = eval ? Visibility.Collapsed : Visibility.Visible;
            if (_colTotal      != null) _colTotal.Visibility      = eval ? Visibility.Collapsed : Visibility.Visible;
            if (_colEval       != null) _colEval.Visibility       = eval ? Visibility.Visible : Visibility.Collapsed;
            if (_hint          != null) _hint.Text                = eval ? EvalHint : MonthlyHint;
        }

        // Clear the roster + tracked state whenever a filter changes so that Save
        // cannot accidentally write data for the wrong (subject,month,year,class) tuple.
        private void InvalidateRoster()
        {
            if (_loadedSubjectId == 0) return; // nothing was loaded yet
            _rows.Clear();
            _loadedMonth = 0; _loadedSubjectId = 0; _loadedSubjectCode = "";
            _loadedYear = ""; _loadedGrade = ""; _loadedRoom = ""; _loadedContext = "";
            if (_info != null) _info.Text = "🔄 ຕົວກອງປ່ຽນແລ້ວ — ໂຫຼດຂໍ້ມູນໃໝ່ໃນອີກຄາວ";
        }

        private void LoadRoster(bool manual = true)
        {
            string grade=_grade.SelectedItem?.ToString()??"ມ.4";
            string room=_room.SelectedItem?.ToString()??"1";
            string year=_year.SelectedItem?.ToString()??DB.CurrentYear;

            if(_subject.SelectedItem==null){ if(manual) MessageBox.Show("ກະລຸນາເລືອກວິຊາ", "ຍັງບໍ່ໄດ້ເລືອກ", MessageBoxButton.OK, MessageBoxImage.Information); return; }
            int subId=((SubComboItem)_subject.SelectedItem).Id;

            // CHA1/LAB1 route to the single-subject evaluation roster (EvaluationScores).
            // Every other subject uses the academic monthly path (MonthlyAssessments).
            if (IsEvalMode) { LoadEvalRoster(subId, grade, room, year, manual); return; }

            int month=((MonthItem)_month.SelectedItem).Value;
            int sem=DB.SemesterForMonth(month);

            string status = _status?.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            string statusFilter = status == "ທັງໝົດ" ? "" : "AND s.Status=@st";
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@subId", subId), ("@year", year), ("@sem", sem),
                ("@month", month), ("@grade", grade), ("@room", room)
            };
            if (status != "ທັງໝົດ") ps.Add(("@st", status));
            var dt = DB.Query($@"
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
            _rows.Clear();
            foreach(DataRow r in dt.Rows)
            {
                var row=new MonthlyRow{
                    StudentID  = Convert.ToInt32(r["StudentID"]),
                    EnrollID   = Convert.ToInt32(r["EnrollID"]),
                    StudentCode= r["StudentCode"].ToString()!,
                    FullName   = r["FullName"].ToString()!,
                    ClassRoom  = r["ClassRoom"].ToString()!,
                    ActivityScore   = Convert.ToDouble(r["ActivityScore"]),
                    DisciplineScore = Convert.ToDouble(r["DisciplineScore"]),
                    HomeworkScore   = Convert.ToDouble(r["HomeworkScore"])
                };
                row.Recalc();
                _rows.Add(row);
            }

            // Stamp the filter state we just loaded — Save uses this tuple as
            // the source of truth (not the current combos) so mid-transition
            // edits from SaveConfirmDialog go to the right (subject, month) slot.
            _loadedMonth       = month;
            _loadedSubjectId   = subId;
            _loadedSubjectCode = SelectedSubjectCode;
            _loadedYear        = year;
            _loadedGrade       = grade;
            _loadedRoom        = room;
            _loadedContext     = "";

            string statusLabel = status == "ທັງໝົດ" ? "ທຸກສະຖານະ" : status;
            string lockHint = (status == "ກຳລັງຮຽນ") ? "" : "  ·  🔒 ອ່ານຢ່າງດຽວ";
            _info.Text=$"📚 ພາກ {sem}  |  ເດືອນ {month:D2}  |  {_rows.Count} ຄົນ  ({statusLabel}){lockHint}";
            ApplyReadOnly();
            _dirty = false;
            _autoReload = true;     // any subsequent filter change reloads automatically
            if(_rows.Count==0 && manual)
                MessageBox.Show("ບໍ່ມີນັກຮຽນທີ່ລົງທະບຽນວິຊານີ້ໃນຫ້ອງດັ່ງກ່າວ\n(ກະລຸນາລົງທະບຽນວິຊາໃຫ້ນັກຮຽນກ່ອນ)","ບໍ່ມີຂໍ້ມູນ",MessageBoxButton.OK,MessageBoxImage.Information);
        }

        // Edits are allowed only when the selected status is exactly ‘ກຳລັງຮຽນ’.
        // Looking at ‘ຈົບ’, ‘ອອກ’, or ‘ທັງໝົດ’ (mixed) is read-only — graduates' and
        // withdrawn students' records are frozen for historical reference.
        private void ApplyReadOnly()
        {
            string status = _status?.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            bool readOnly = status != "ກຳລັງຮຽນ";
            if (_dg     != null) _dg.IsReadOnly      = readOnly;
            if (_btnSave!= null) _btnSave.IsEnabled  = !readOnly;
        }

        // Single-subject evaluation roster — every student in the class plus the
        // student's existing manual eval score for (subject, Month{N}). Used when
        // the picked subject is CHA1/LAB1. No Enrollment/month JOIN needed —
        // evaluation scores are per (student, year, context, subject).
        private void LoadEvalRoster(int subId, string grade, string room, string year, bool manual)
        {
            string ctx  = EvalContext;
            string code = SelectedSubjectCode;
            if (string.IsNullOrEmpty(ctx) || string.IsNullOrEmpty(code))
            {
                if (manual)
                    MessageBox.Show("ກະລຸນາເລືອກວິຊາ + ເດືອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string status = _status?.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            string statusFilter = status == "ທັງໝົດ" ? "" : "AND s.Status=@st";
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@year", year), ("@ctx", ctx), ("@code", code),
                ("@grade", grade), ("@room", room)
            };
            if (status != "ທັງໝົດ") ps.Add(("@st", status));
            var dt = DB.Query($@"
                SELECT s.StudentID, s.StudentCode, s.FirstName||' '||s.LastName AS FullName, s.ClassRoom,
                       ev.Score AS EvalScore
                FROM Students s
                LEFT JOIN EvaluationScores ev
                       ON ev.StudentID=s.StudentID AND ev.AcademicYear=@year
                      AND ev.Context=@ctx AND ev.SubjectCode=@code
                WHERE s.GradeLevel=@grade AND s.ClassRoom=@room {statusFilter}
                ORDER BY s.StudentCode",
                null, ps.ToArray());
            _rows.Clear();
            foreach (DataRow r in dt.Rows)
            {
                _rows.Add(new MonthlyRow{
                    StudentID  = Convert.ToInt32(r["StudentID"]),
                    EnrollID   = 0,                       // not used in eval mode
                    StudentCode= r["StudentCode"].ToString()!,
                    FullName   = r["FullName"].ToString()!,
                    ClassRoom  = r["ClassRoom"].ToString()!,
                    EvalScore  = r["EvalScore"] == DBNull.Value ? (double?)null : Convert.ToDouble(r["EvalScore"])
                });
            }

            // Stamp loaded state so Save writes to this exact (subject, month,
            // context) even if the combos have moved on. Each Month{N} is
            // its own EvaluationScores context — subject code matters too.
            _loadedMonth       = ((MonthItem)_month.SelectedItem).Value;
            _loadedSubjectId   = subId;
            _loadedSubjectCode = code;
            _loadedYear        = year;
            _loadedGrade       = grade;
            _loadedRoom        = room;
            _loadedContext     = ctx;

            string lockHint = (status == "ກຳລັງຮຽນ") ? "" : "  ·  🔒 ອ່ານຢ່າງດຽວ";
            _info.Text = $"📋 {code}  |  ເດືອນ {_loadedMonth:D2}  |  {_rows.Count} ຄົນ{lockHint}";
            ApplyReadOnly();
            _dirty = false;
            _autoReload = true;
            if (_rows.Count == 0 && manual)
                MessageBox.Show("ບໍ່ມີນັກຮຽນໃນຫ້ອງດັ່ງກ່າວ","ບໍ່ມີຂໍ້ມູນ",MessageBoxButton.OK,MessageBoxImage.Information);
        }

        private void SaveAll()
        {
            // Block save when looking at non-active students (read-only mode).
            if (_dg != null && _dg.IsReadOnly)
            {
                MessageBox.Show("ບໍ່ສາມາດແກ້ໄຂຄະແນນຂອງນັກຮຽນທີ່ບໍ່ໄດ້ກຳລັງຮຽນ — ປ່ຽນຕົວກອງ ‘ສະຖານະ’ ກັບໄປເປັນ ‘ກຳລັງຮຽນ’ ກ່ອນ",
                    "ອ່ານຢ່າງດຽວ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if(_rows.Count==0)
            {
                MessageBox.Show("ບໍ່ມີຂໍ້ມູນທີ່ຈະບັນທຶກ — ກະລຸນາໂຫຼດກ່ອນ","ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            _dg.CommitEdit(DataGridEditingUnit.Row,true);

            // Route by what was LOADED, not by what the combos currently show
            // (the SaveConfirmDialog Save path can call us with drifted combos).
            // CHA1/LAB1 → EvaluationScores. Every other subject → MonthlyAssessments.
            if (!string.IsNullOrEmpty(_loadedContext)) { SaveEval(); return; }

            // Use the loaded tuple as the write context. This eliminates the
            // old filter-drift MessageBox — writes go to (loaded subject,
            // loaded month) regardless of combo state.
            int month = _loadedMonth;

            int saved=0;
            using var conn=DB.Open();
            using var tx=conn.BeginTransaction();
            try
            {
                foreach(var row in _rows)
                {
                    // Clamp into valid sub-score ranges before persisting.
                    row.ActivityScore   = Math.Max(0, Math.Min(3, row.ActivityScore));
                    row.DisciplineScore = Math.Max(0, Math.Min(2, row.DisciplineScore));
                    row.HomeworkScore   = Math.Max(0, Math.Min(5, row.HomeworkScore));
                    row.Recalc();

                    DB.ExecTx(@"INSERT INTO MonthlyAssessments(EnrollID,Month,ActivityScore,DisciplineScore,HomeworkScore,UpdatedAt)
                                VALUES(@e,@mo,@a,@d,@h,datetime('now','localtime'))
                                ON CONFLICT(EnrollID,Month) DO UPDATE SET
                                  ActivityScore=excluded.ActivityScore,
                                  DisciplineScore=excluded.DisciplineScore,
                                  HomeworkScore=excluded.HomeworkScore,
                                  UpdatedAt=datetime('now','localtime')",
                        conn, tx,
                        ("@e",row.EnrollID),("@mo",month),
                        ("@a",row.ActivityScore),("@d",row.DisciplineScore),("@h",row.HomeworkScore));
                    saved++;
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // After commit: recompute Scores.MidScore (semester-wide average) for each enrollment.
            // Reuse one connection across the whole batch so the overhead is O(1) instead of O(N).
            using (var rc = DB.Open())
                foreach (var row in _rows) DB.RecomputeMidFromMonthly(row.EnrollID, rc);
            DB.Log("MonthlyScores",$"{saved} ຄົນ ເດືອນ {month:D2}");
            _dirty = false;
            // Inline feedback — non-blocking, lets the user continue editing immediately.
            _info.Text = $"✅ ບັນທຶກ {saved} ຄົນ ສຳເລັດ — {DateTime.Now:HH:mm:ss}";
        }

        // Single-subject (CHA1/LAB1) eval save. UPSERT per row via DB.SetEvaluationScore —
        // null score deletes the row. Filter-drift guard: (subject, year, grade, room,
        // context) must all match what was loaded. No averaging / recompute — each
        // Month{N} context is stored verbatim.
        private void SaveEval()
        {
            // Write to the loaded (year, context, subject) tuple — the
            // combos may have drifted (SaveConfirmDialog Save path) but the
            // in-memory `_rows` still hold data for the loaded slot.
            string year = _loadedYear;
            string ctx  = _loadedContext;
            string code = _loadedSubjectCode;

            int saved = 0;
            using var conn = DB.Open();
            using var tx = conn.BeginTransaction();
            try
            {
                foreach (var row in _rows)
                {
                    double? v = row.EvalScore.HasValue
                        ? Math.Max(0, Math.Min(10, row.EvalScore.Value))
                        : (double?)null;
                    DB.SetEvaluationScore(row.StudentID, year, ctx, code, v, conn, tx);
                    saved++;
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DB.Log("EvalScores", $"{_loadedContext} {code} {saved} ຄົນ");
            _dirty = false;
            _info.Text = $"✅ ບັນທຶກ {saved} ຄົນ ສຳເລັດ — {DateTime.Now:HH:mm:ss}";
        }

        // Filter guard for the import buttons. Import is per-subject so we also
        // require a Subject pick. The (year, grade, room, month, subject) tuple
        // is the scope key stored in the workbook's hidden metadata cells.
        private bool RequireFilters(out string year, out string grade, out string room, out int month, out string subjectCode)
        {
            year  = _year.SelectedItem?.ToString() ?? "";
            grade = _grade.SelectedItem?.ToString() ?? "";
            room  = _room.SelectedItem?.ToString() ?? "";
            month = _month.SelectedItem is MonthItem mi ? mi.Value : 0;
            subjectCode = SelectedSubjectCode;
            if (string.IsNullOrEmpty(year) || string.IsNullOrEmpty(grade)
                || string.IsNullOrEmpty(room) || month == 0 || string.IsNullOrEmpty(subjectCode))
            {
                MessageBox.Show("ກະລຸນາເລືອກ ປີ · ຊັ້ນ · ຫ້ອງ · ວິຊາ · ເດືອນ ກ່ອນ",
                    "ຍັງບໍ່ໄດ້ເລືອກ", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            return true;
        }

        private void DownloadTemplate()
        {
            if (!RequireFilters(out string year, out string grade, out string room, out int month, out string subject)) return;

            var dlg = new SaveFileDialog {
                Filter   = "Excel|*.xlsx",
                FileName = DB.SafeFileName($"ນຳເຂົ້າຄະແນນປະຈຳເດືອນ_{year}_{grade}-{room}_M{month:D2}_{subject}.xlsx")
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                ExcelImport.BuildMonthlyTemplate(dlg.FileName, year, grade, room, month, subject);
                _info.Text = $"✅ ສ້າງແບບຟອມສຳເລັດ ({subject}) — {Path.GetFileName(dlg.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ສ້າງແບບຟອມບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ImportExcel()
        {
            if (!RequireFilters(out string year, out string grade, out string room, out int month, out string subject)) return;
            string status = _status?.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            if (status != "ກຳລັງຮຽນ")
            {
                MessageBox.Show("ການນຳເຂົ້າຮອງຮັບສະເພາະນັກຮຽນທີ່ ‘ກຳລັງຮຽນ’ — ກະລຸນາປ່ຽນຕົວກອງ ‘ສະຖານະ’ ກັບໄປຂ້າງເທິງ",
                    "ອ່ານຢ່າງດຽວ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dlg = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (dlg.ShowDialog() != true) return;

            ImportResult result;
            try { result = ExcelImport.ParseMonthly(dlg.FileName, year, grade, room, month, subject); }
            catch (Exception ex)
            {
                MessageBox.Show($"ອ່ານໄຟລ໌ບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var win = new ImportPreviewWindow(result) { Owner = Window.GetWindow(this) };
            bool? ok = win.ShowDialog();
            if (ok == true)
            {
                // Refresh only the affected subject's roster (which is exactly
                // what the page is showing — the import scope matches the page
                // scope by construction). Other subjects on disk are untouched
                // so we don't need to repaint anything else.
                LoadRoster(manual:false);
                _info.Text = $"✅ ນຳເຂົ້າ {win.SavedCount} ແຖວ ({subject}) ສຳເລັດ — {DateTime.Now:HH:mm:ss}";
            }
        }

    }

    public class MonthlyRow : INotifyPropertyChanged
    {
        public int StudentID{get;set;} public int EnrollID{get;set;}
        public string StudentCode{get;set;}=""; public string FullName{get;set;}="";
        public string ClassRoom{get;set;}="";
        private double _a, _d, _h, _t;
        public double ActivityScore  { get=>_a; set{ if(_a==value) return; _a=value; Notify(); Recalc(); } }
        public double DisciplineScore{ get=>_d; set{ if(_d==value) return; _d=value; Notify(); Recalc(); } }
        public double HomeworkScore  { get=>_h; set{ if(_h==value) return; _h=value; Notify(); Recalc(); } }
        public double Total          { get=>_t; private set{ _t=value; Notify(); } }
        // Evaluation mode only — single manual /10 score for CHA1/LAB1. Nullable so
        // an un-entered cell stays blank rather than reading as 0.
        private double? _eval;
        public double? EvalScore     { get=>_eval; set{ if(_eval==value) return; _eval=value; Notify(); } }
        public void Recalc() => Total = DB.CalcMonthlyTotal(_a,_d,_h);
        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify([CallerMemberName] string? n=null) => PropertyChanged?.Invoke(this,new PropertyChangedEventArgs(n));
    }

    public class MonthItem
    {
        public int Value { get; }
        public string Label { get; }
        public MonthItem(int v, string l) { Value = v; Label = l; }
        public override string ToString() => Label;
    }

    // ════════════════════════════════════════════════════════════
    //  IMPORT PREVIEW WINDOW
    //
    //  Opened from MonthlyScoresPage / ScoresPage after the teacher picks
    //  an .xlsx via Open dialog. Shows the parsed rows with a Status
    //  column (✅ ຖືກຕ້ອງ / ❌ <reason>) and a summary footer. The 💾
    //  Save button is disabled when ValidCount == 0. On confirm,
    //  ExcelImport.SaveImport runs in one transaction; DialogResult=true
    //  tells the caller to refresh affected rows only.
    // ════════════════════════════════════════════════════════════
    public class ImportPreviewWindow : Window
    {
        private readonly ImportResult _result;
        public int SavedCount { get; private set; }

        public ImportPreviewWindow(ImportResult result)
        {
            _result = result;
            Title = "👁  ກວດເບິ່ງຂໍ້ມູນກ່ອນບັນທຶກ";
            Width = 820; Height = 620;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252));

            var root = H.MkGrid(GridLength.Auto, GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // Header — every preview is per-subject, so the subject name is part
            // of the title (the data rows don't carry a Subject column anymore).
            string subjLabel = string.IsNullOrEmpty(result.SubjectName)
                ? result.SubjectCode
                : $"{result.SubjectCode} — {result.SubjectName}";
            string ctxLabel = result.Kind == ImportKind.Monthly
                ? $"ປະຈຳເດືອນ · ປີ {result.Year} · {result.Grade}/{result.Room} · ເດືອນ {result.Month:D2} · ວິຊາ {subjLabel}"
                : $"ພາກຮຽນ · ປີ {result.Year} · {result.Grade}/{result.Room} · ພາກ {result.Semester} · ວິຊາ {subjLabel}";
            var hdr = H.MkCard(new Thickness(12,12,12,8), new Thickness(14,10,14,10));
            hdr.Child = new TextBlock {
                Text = $"📥  ນຳເຂົ້າຄະແນນ — {ctxLabel}",
                FontSize = 15, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            };
            Grid.SetRow(hdr, 0); root.Children.Add(hdr);

            // Summary bar
            var sum = new Border {
                Margin = new Thickness(12,0,12,8), Padding = new Thickness(14,8,14,8),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239,246,255)),
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(214,228,247)),
                BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(4)
            };
            string summary;
            if (result.FatalError != null)
                summary = "❌  " + result.FatalError;
            else
                summary = $"📊  ທັງໝົດ {result.Rows.Count} ແຖວ   ·   ✅ ຖືກຕ້ອງ {result.ValidCount}   ·   ❌ ຜິດພາດ {result.InvalidCount}";
            sum.Child = new TextBlock {
                Text = summary, FontSize = 13, TextWrapping = TextWrapping.Wrap,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(27,79,138))
            };
            Grid.SetRow(sum, 1); root.Children.Add(sum);

            // Preview grid
            var card = H.MkCard(new Thickness(12,0,12,8), new Thickness(0));
            var dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White,
                HeadersVisibility = DataGridHeadersVisibility.Column
            };
            dg.Columns.Add(new DataGridTextColumn { Header = "ລຳດັບ",     Binding = new System.Windows.Data.Binding("RowNo"),       Width = new DataGridLength(60) });
            dg.Columns.Add(new DataGridTextColumn { Header = "ລະຫັດ",     Binding = new System.Windows.Data.Binding("StudentCode"), Width = new DataGridLength(110) });
            dg.Columns.Add(new DataGridTextColumn { Header = "ຊື່ນັກຮຽນ", Binding = new System.Windows.Data.Binding("StudentName"), Width = new DataGridLength(240) });
            if (result.Kind == ImportKind.Monthly && !result.IsEvalSubject)
            {
                // Academic monthly: 3 sub-scores + computed total — matches the
                // template's D/E/F columns one-to-one.
                dg.Columns.Add(new DataGridTextColumn {
                    Header = "ກິດຈະກຳ (/2)", Width = new DataGridLength(80),
                    Binding = new System.Windows.Data.Binding("DisciplineScore") { TargetNullValue = "" }
                });
                dg.Columns.Add(new DataGridTextColumn {
                    Header = "ຮ່ວມຮຽນ (/3)", Width = new DataGridLength(80),
                    Binding = new System.Windows.Data.Binding("ActivityScore") { TargetNullValue = "" }
                });
                dg.Columns.Add(new DataGridTextColumn {
                    Header = "ກວດກາ (/5)", Width = new DataGridLength(80),
                    Binding = new System.Windows.Data.Binding("HomeworkScore") { TargetNullValue = "" }
                });
                dg.Columns.Add(new DataGridTextColumn {
                    Header = "ລວມ (/10)", Width = new DataGridLength(75),
                    Binding = new System.Windows.Data.Binding("SubScoreTotal")
                });
            }
            else if (result.Kind == ImportKind.Monthly)
            {
                // CHA1/LAB1 monthly: single /10 column (EvaluationScores).
                dg.Columns.Add(new DataGridTextColumn {
                    Header = "ຄະແນນ (/10)", Width = new DataGridLength(90),
                    Binding = new System.Windows.Data.Binding("Score") { TargetNullValue = "" }
                });
            }
            else
            {
                // Semester: only the final-exam score is imported (Mid stays
                // auto-derived from MonthlyAssessments, so it's not in the file).
                dg.Columns.Add(new DataGridTextColumn {
                    Header = "ສອບເສງພາກຮຽນ (/10)", Width = new DataGridLength(140),
                    Binding = new System.Windows.Data.Binding("FinalScore") { TargetNullValue = "" }
                });
            }
            dg.Columns.Add(new DataGridTextColumn {
                Header = "ສະຖານະການກວດສອບ", Binding = new System.Windows.Data.Binding("Status"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            });
            // Light row tint so the eye finds invalid rows immediately.
            var rowStyle = new Style(typeof(DataGridRow));
            var trigInv = new System.Windows.DataTrigger {
                Binding = new System.Windows.Data.Binding("IsValid"), Value = false
            };
            trigInv.Setters.Add(new Setter(BackgroundProperty,
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(254, 242, 242))));
            rowStyle.Triggers.Add(trigInv);
            dg.RowStyle = rowStyle;
            dg.ItemsSource = _result.Rows;
            card.Child = dg;
            Grid.SetRow(card, 2); root.Children.Add(card);

            // Footer: Cancel + Save
            var foot = new Border {
                Margin = new Thickness(12,4,12,12), Padding = new Thickness(0)
            };
            var bar = new DockPanel { LastChildFill = false };
            var btnCancel = H.Btn("✖  ຍົກເລີກ", "SecondaryButton");
            btnCancel.Click += (s,e) => { DialogResult = false; Close(); };
            DockPanel.SetDock(btnCancel, Dock.Left);
            bar.Children.Add(btnCancel);

            var btnSave = H.Btn($"💾  ບັນທຶກ {result.ValidCount} ແຖວ", "SuccessButton");
            btnSave.IsEnabled = result.FatalError == null && result.ValidCount > 0;
            btnSave.Click += (s,e) =>
            {
                try
                {
                    SavedCount = ExcelImport.SaveImport(_result);
                    DialogResult = true;
                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ:\n{ex.Message}",
                        "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            };
            DockPanel.SetDock(btnSave, Dock.Right);
            bar.Children.Add(btnSave);
            foot.Child = bar;
            Grid.SetRow(foot, 3); root.Children.Add(foot);

            Content = root;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  PROMOTION
    // ════════════════════════════════════════════════════════════
    // ════════════════════════════════════════════════════════════
    //  PROMOTION PAGE — redesigned for safer + simpler bulk actions
    //
    //  Top: filter card (Year + Grade + Room) + status radios + action
    //  toolbar. Roster shows every student matching (year, grade, room,
    //  status). Four bulk actions:
    //    ⬆ Promote Selected     → opens PromotionConfirmWindow with
    //                              destination preview; per-student
    //                              From → To list; user confirms.
    //    🔁 Repeat Selected     → same dialog as Promote, but destination
    //       (ຊ້ຳຊັ້ນ)             grade defaults to the SAME grade (student
    //                              retained); AcademicYear advances. Room
    //                              editable if the school reshuffles
    //                              sections. GradeHistory.Note='ຊ້ຳຊັ້ນ',
    //                              ‘ຈົບ’ hidden from the dst-grade dropdown.
    //    🎓 Graduate Selected   → simple Yes/No confirm; sets Status='ຈົບ',
    //                              writes GradeHistory(ToGrade='ຈົບ').
    //    🔄 Promote Entire      → auto-selects every active row + funnels
    //       Classroom             through the same Promote-Selected flow.
    //
    //  Plus: 📋 Select All / ❌ Clear Selection for manual row picking.
    //
    //  History rules respected end-to-end:
    //    - Scores / MonthlyAssessments / EvaluationScores /
    //      AttendanceRecords / Enrollments NEVER deleted by promotion or
    //      repeat.
    //    - GradeHistory row written for every action (ToGrade='ຈົບ' for
    //      graduation, NextGrade for promotion, FromGrade=ToGrade for
    //      repeat). ClassRoom recorded. Note reflects the action.
    //    - CHA1/LAB1 untouched.
    //    - Graduates' Students.AcademicYear stays at finishing year; their
    //      Students.GradeLevel stays at grade they finished — historical
    //      reports keep reconstructing their cohort via GradeHistory.
    // ════════════════════════════════════════════════════════════
    public class PromotionPage : UserControl
    {
        private ComboBox _cmbYear = null!, _cmbGrade = null!, _cmbRoom = null!;
        private RadioButton _rdoActive = null!, _rdoGraduated = null!, _rdoAll = null!;
        private DataGrid _dg = null!, _hist = null!;
        private TextBlock _info = null!;
        private Button _btnPromote = null!, _btnGraduate = null!, _btnRepeat = null!, _btnPromoteClass = null!;
        private readonly ObservableCollection<PromRow> _rows = new();

        public PromotionPage()
        {
            // 4 rows: filter card · action toolbar · roster grid · history panel.
            var root = H.MkGrid(GridLength.Auto, GridLength.Auto,
                                new GridLength(1, GridUnitType.Star),
                                new GridLength(180));

            // ── Row 0: Filter card ───────────────────────────────────────
            var tb   = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,10,14,10));
            var flow = new WrapPanel();
            _cmbYear  = H.MkCmb(DB.AcademicYears().ToArray(), 110);
            SelectComboValuePR(_cmbYear, DB.CurrentYear);
            _cmbGrade = H.MkCmb(new[] { "ມ.1", "ມ.2", "ມ.3", "ມ.4" }, 80);
            _cmbRoom  = H.MkCmb(new[] { "1", "2", "3", "4", "5", "6" }, 80);
            _cmbYear.SelectionChanged  += (s,e) => ReloadRoster();
            _cmbGrade.SelectionChanged += (s,e) => ReloadRoster();
            _cmbRoom.SelectionChanged  += (s,e) => ReloadRoster();

            _rdoActive    = MkRadioPR(" ກຳລັງຮຽນ", isChecked: true);
            _rdoGraduated = MkRadioPR("🎓 ຈົບ",        isChecked: false);
            _rdoAll       = MkRadioPR("📋 ທັງໝົດ",    isChecked: false);
            _rdoActive.Checked    += (s,e) => ReloadRoster();
            _rdoGraduated.Checked += (s,e) => ReloadRoster();
            _rdoAll.Checked       += (s,e) => ReloadRoster();

            flow.Children.Add(H.Lbl("ປີ:"));    flow.Children.Add(_cmbYear);
            flow.Children.Add(H.Lbl("ຊັ້ນ:"));  flow.Children.Add(_cmbGrade);
            flow.Children.Add(H.Lbl("ຫ້ອງ:"));  flow.Children.Add(_cmbRoom);
            flow.Children.Add(new Border {
                Width = 1, Height = 22, Margin = new Thickness(6,0,12,0),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(209,213,219))
            });
            flow.Children.Add(_rdoActive);
            flow.Children.Add(_rdoGraduated);
            flow.Children.Add(_rdoAll);
            tb.Child = flow; Grid.SetRow(tb, 0); root.Children.Add(tb);

            // ── Row 1: Action toolbar ────────────────────────────────────
            var actBar = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,8,14,8));
            var actFlow = new WrapPanel();
            var bSelAll = H.Btn("📋  ເລືອກທັງໝົດ",  "NeutralButton");
            var bClear  = H.Btn("❌  ລ້າງການເລືອກ", "NeutralButton");
            _btnPromote      = H.Btn("⬆️  ຂຶ້ນຊັ້ນ (ທີ່ເລືອກ)",   "PrimaryButton");
            _btnRepeat       = H.Btn("🔁  ຊ້ຳຊັ້ນ (ທີ່ເລືອກ)",   "SecondaryButton");
            _btnGraduate     = H.Btn("🎓  ຈົບການສຶກສາ (ທີ່ເລືອກ)", "WarningButton");
            _btnPromoteClass = H.Btn("🔄  ຂຶ້ນຊັ້ນທັງຫ້ອງ",        "SuccessButton");

            bSelAll.Click          += (s,e) => SelectAllActive();
            bClear.Click           += (s,e) => ClearSelection();
            _btnPromote.Click      += (s,e) => PromoteSelected();
            _btnRepeat.Click       += (s,e) => RepeatSelected();
            _btnGraduate.Click     += (s,e) => GraduateSelected();
            _btnPromoteClass.Click += (s,e) => PromoteEntireClassroom();

            _info = new TextBlock {
                FontSize = 12, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))
            };
            actFlow.Children.Add(bSelAll); actFlow.Children.Add(bClear);
            actFlow.Children.Add(new Border {
                Width = 1, Height = 22, Margin = new Thickness(6,0,12,0),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(209,213,219))
            });
            actFlow.Children.Add(_btnPromote);
            actFlow.Children.Add(_btnRepeat);
            actFlow.Children.Add(_btnGraduate);
            actFlow.Children.Add(_btnPromoteClass);
            actFlow.Children.Add(_info);
            actBar.Child = actFlow; Grid.SetRow(actBar, 1); root.Children.Add(actBar);

            // ── Row 2: Roster DataGrid ───────────────────────────────────
            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = false, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White
            };
            _dg.Columns.Add(new DataGridCheckBoxColumn {
                Header = "☑",
                Binding = new System.Windows.Data.Binding("IsSelected") {
                    UpdateSourceTrigger = System.Windows.Data.UpdateSourceTrigger.PropertyChanged
                },
                Width = new DataGridLength(40)
            });
            _dg.Columns.Add(H.Col("ລະຫັດ",      "StudentCode",  100, true));
            _dg.Columns.Add(H.ColStar("ຊື່ນັກຮຽນ", "FullName",       true));
            _dg.Columns.Add(H.Col("ເພດ",         "Gender",        60, true));
            _dg.Columns.Add(H.Col("ຊັ້ນ",        "GradeLevel",    70, true));
            _dg.Columns.Add(H.Col("ຫ້ອງ",        "ClassRoom",     70, true));
            _dg.Columns.Add(H.Col("ປີ",          "AcademicYear",  90, true));
            _dg.Columns.Add(H.Col("ສະຖານະ",      "Status",        90, true));
            _dg.ItemsSource = _rows;
            var dcard = H.MkCard(new Thickness(0,0,0,10)); dcard.Child = _dg;
            Grid.SetRow(dcard, 2); root.Children.Add(dcard);

            // ── Row 3: History panel ─────────────────────────────────────
            _hist = new DataGrid {
                AutoGenerateColumns = true, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White
            };
            var hcard = H.MkCard(); hcard.Child = new DockPanel();
            var ht = new TextBlock {
                Text = "📜  ປະຫວັດການຂຶ້ນຊັ້ນ (50 ຄັ້ງລ່າສຸດ)",
                FontSize = 14, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39)),
                Margin = new Thickness(0,0,0,8)
            };
            DockPanel.SetDock(ht, Dock.Top);
            ((DockPanel)hcard.Child).Children.Add(ht);
            ((DockPanel)hcard.Child).Children.Add(_hist);
            Grid.SetRow(hcard, 3); root.Children.Add(hcard);

            Content = root;
            ReloadRoster();
            LoadHistory();
        }

        // Same GroupName → exclusive selection.
        private static RadioButton MkRadioPR(string label, bool isChecked) =>
            new RadioButton {
                Content = label, IsChecked = isChecked,
                GroupName = "PromotionStatus",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0,0,14,0), FontSize = 13
            };

        private static void SelectComboValuePR(ComboBox cb, string val)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == val) { cb.SelectedIndex = i; return; }
        }

        private string? CurrentStatusFilter()
        {
            if (_rdoGraduated?.IsChecked == true) return "ຈົບ";
            if (_rdoAll?.IsChecked == true)       return null;
            return "ກຳລັງຮຽນ";
        }

        // ── Roster reload ────────────────────────────────────────────────
        // Performance: nothing loads until Year + Grade + Room are all set.
        // Reuses DB.GetHistoricalClassRoster so the cohort logic is identical
        // to the score-history page (handles graduated + promoted students),
        // then enriches with the extra Students columns the promotion UI
        // needs (Current Grade / Classroom / Academic Year / Status).
        private void ReloadRoster()
        {
            _rows.Clear();
            string year  = _cmbYear?.SelectedItem?.ToString()  ?? "";
            string grade = _cmbGrade?.SelectedItem?.ToString() ?? "";
            string room  = _cmbRoom?.SelectedItem?.ToString()  ?? "";
            bool ready = year != "" && grade != "" && room != "";
            UpdateInfo(ready ? 0 : -1);
            if (!ready) return;

            string? status = CurrentStatusFilter();
            var cohort = DB.GetHistoricalClassRoster(year, grade, room, status);
            if (cohort.Rows.Count == 0) { UpdateInfo(0); return; }

            // Pull current Students fields for each StudentID in the cohort.
            var ids = new List<string>();
            foreach (DataRow rr in cohort.Rows) ids.Add(rr["StudentID"].ToString()!);
            string idCsv = string.Join(",", ids);
            var stuDt = DB.Query($@"
                SELECT StudentID, StudentCode, FirstName||' '||LastName AS FullName,
                       Gender, GradeLevel, ClassRoom, AcademicYear, Status
                FROM Students WHERE StudentID IN ({idCsv})
                ORDER BY StudentCode");
            foreach (DataRow r in stuDt.Rows)
            {
                _rows.Add(new PromRow {
                    StudentID    = Convert.ToInt32(r["StudentID"]),
                    StudentCode  = r["StudentCode"].ToString()!,
                    FullName     = r["FullName"].ToString()!,
                    Gender       = r["Gender"].ToString()!,
                    GradeLevel   = r["GradeLevel"].ToString()!,
                    ClassRoom    = r["ClassRoom"].ToString()!,
                    AcademicYear = r["AcademicYear"].ToString()!,
                    Status       = r["Status"].ToString()!,
                });
            }
            UpdateInfo(_rows.Count);
        }

        private void UpdateInfo(int count)
        {
            if (_info == null) return;
            _info.Text = count < 0
                ? "👉 ກະລຸນາເລືອກ ປີ + ຊັ້ນ + ຫ້ອງ ກ່ອນ"
                : $"👥 ນັກຮຽນ {count} ຄົນ";
        }

        // ── Bulk-selection actions ───────────────────────────────────────
        private void SelectAllActive()
        {
            foreach (var r in _rows)
                if (r.Status == "ກຳລັງຮຽນ") r.IsSelected = true;
        }
        private void ClearSelection()
        {
            foreach (var r in _rows) r.IsSelected = false;
        }

        // ── Action: Promote selected ─────────────────────────────────────
        // Filters to active rows only (graduates/withdrawn can't be promoted
        // again), opens PromotionConfirmWindow with destination preview, runs
        // DB updates on confirm.
        private void PromoteSelected()
        {
            var sel = new List<PromRow>();
            foreach (var r in _rows) if (r.IsSelected && r.Status == "ກຳລັງຮຽນ") sel.Add(r);
            if (sel.Count == 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນທີ່ກຳລັງຮຽນກ່ອນ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (!EnsureNextYearExists(sel)) return;
            OpenPromoteDialog(sel);
        }

        // ── Action: Repeat selected (ຊ້ຳຊັ້ນ) ─────────────────────────────
        // Same dialog as Promote, but destination grade defaults to the SAME
        // grade (student is retained) while AcademicYear advances. Room can
        // change if the school reshuffles sections. GradeHistory row is
        // logged with Note='ຊ້ຳຊັ້ນ'. Graduated/withdrawn students are
        // skipped (only ‘ກຳລັງຮຽນ’ students can repeat).
        private void RepeatSelected()
        {
            var sel = new List<PromRow>();
            foreach (var r in _rows) if (r.IsSelected && r.Status == "ກຳລັງຮຽນ") sel.Add(r);
            if (sel.Count == 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນທີ່ກຳລັງຮຽນກ່ອນ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (!EnsureNextYearExists(sel)) return;
            OpenPromoteDialog(sel, isRepeat: true);
        }

        // ── Guardrail: next academic year must exist before promotion/repeat ─
        // Both flows write Students.AcademicYear + GradeHistory.AcademicYear
        // to a NEW year (srcYear + 1). If that year hasn't been added to the
        // AcademicYears registry yet, refuse to proceed — offer to open the
        // AcademicYearFormWin prefilled so the admin creates it right here.
        //
        // Returns true when the target year exists (or was just created);
        // false to cancel the promotion.
        private bool EnsureNextYearExists(List<PromRow> selected)
        {
            string srcYear = MostCommon(selected.Select(r => r.AcademicYear));
            string dstYear = DB.NextYearString(srcYear);
            if (DB.AcademicYears().Contains(dstYear)) return true;

            var res = MessageBox.Show(
                $"ຍັງບໍ່ໄດ້ເພີ່ມ ‘ປີການສຶກສາ {dstYear}’ ໃນລະບົບ.\n\n" +
                "ຕ້ອງເພີ່ມປີການສຶກສາ ‘ປີໜ້າ’ ກ່ອນ ຈຶ່ງຈະດຳເນີນການຂຶ້ນຊັ້ນ / ຊ້ຳຊັ້ນໄດ້.\n\n" +
                "ຕ້ອງການເພີ່ມປີການສຶກສາດຽວນີ້ບໍ?",
                "ຕ້ອງເພີ່ມປີການສຶກສາກ່ອນ",
                MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (res != MessageBoxResult.Yes) return false;

            var dlg = new AcademicYearFormWin(dstYear) { Owner = Window.GetWindow(this) };
            if (dlg.ShowDialog() != true) return false;

            // Double-check the user actually saved the expected year — they
            // may have edited the textbox before hitting ບັນທຶກ.
            if (!DB.AcademicYears().Contains(dstYear))
            {
                MessageBox.Show(
                    $"ຍັງບໍ່ໄດ້ບັນທຶກ ‘ປີການສຶກສາ {dstYear}’ — ຍົກເລີກການດຳເນີນການ.",
                    "ບໍ່ສາມາດດຳເນີນຕໍ່", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        // ── Action: Graduate selected ────────────────────────────────────
        // Sets Status='ຈົບ'. Students row's GradeLevel + AcademicYear are
        // PRESERVED (anchored to finishing year/grade) so historical reports
        // can still reconstruct their cohort.
        private void GraduateSelected()
        {
            var sel = new List<PromRow>();
            foreach (var r in _rows) if (r.IsSelected && r.Status == "ກຳລັງຮຽນ") sel.Add(r);
            if (sel.Count == 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນທີ່ກຳລັງຮຽນກ່ອນ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (MessageBox.Show(
                    $"⚠ ກຳລັງຈະປ່ຽນສະຖານະຂອງ {sel.Count} ຄົນ ເປັນ 'ຈົບ'.\n\n" +
                    "ຄະແນນ + ປະຫວັດທັງໝົດຍັງຄົງຢູ່. ດຳເນີນການຕໍ່?",
                    "ຢືນຢັນການຈົບການສຶກສາ", MessageBoxButton.YesNo, MessageBoxImage.Warning)
                != MessageBoxResult.Yes) return;

            // The “year being entered” for a graduation event is the year AFTER
            // the finishing year — matches PromotionPage's existing convention
            // (GradeHistory.AcademicYear = year-after-grad for ToGrade='ຈົບ').
            // For mixed-year selections we record per-student.
            using var conn = DB.Open(); using var tx = conn.BeginTransaction();
            try
            {
                foreach (var row in sel)
                {
                    string nextY = DB.NextYearString(row.AcademicYear);
                    DB.ExecTx("UPDATE Students SET Status='ຈົບ' WHERE StudentID=@id",
                        conn, tx, ("@id", row.StudentID));
                    DB.ExecTx(@"INSERT INTO GradeHistory
                                (StudentID, FromGrade, ToGrade, AcademicYear, ClassRoom, Note, ChangedBy)
                                VALUES (@sid, @fg, 'ຈົບ', @y, @cr, @n, @by)",
                        conn, tx,
                        ("@sid", row.StudentID), ("@fg", row.GradeLevel),
                        ("@y",   nextY),         ("@cr", (object?)row.ClassRoom ?? DBNull.Value),
                        ("@n",   "ຈົບການສຶກສາ"), ("@by", DB.CurrentUser));
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ບໍ່ສຳເລັດ:\n{ex.Message}", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            DB.Log("Graduation", $"{sel.Count} ຄົນ");
            MessageBox.Show($"ຈົບການສຶກສາ {sel.Count} ຄົນ ສຳເລັດ", "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);
            ReloadRoster(); LoadHistory();
        }

        // ── Action: Promote entire classroom ─────────────────────────────
        // Auto-selects every active student in the current roster, then
        // funnels through the same confirm flow as Promote Selected.
        private void PromoteEntireClassroom()
        {
            int active = 0;
            foreach (var r in _rows)
            {
                bool isActive = r.Status == "ກຳລັງຮຽນ";
                r.IsSelected = isActive;
                if (isActive) active++;
            }
            if (active == 0)
            {
                MessageBox.Show("ບໍ່ມີນັກຮຽນທີ່ກຳລັງຮຽນໃນຫ້ອງນີ້", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            PromoteSelected();
        }

        // Open the destination + preview dialog. `isRepeat` toggles between
        // ຂຶ້ນຊັ້ນ (default — dstGrade = NextGrade) and ຊ້ຳຊັ້ນ (dstGrade = srcGrade,
        // student stays in the same grade next year). The dialog itself hides
        // the "ຈົບ" grade option when repeat is active, and the GradeHistory
        // note + log message reflect the chosen intent.
        private void OpenPromoteDialog(List<PromRow> selected, bool isRepeat = false)
        {
            // Defaults: NextYear of the most-common selected year, NextGrade of
            // the most-common selected grade (or same grade for repeat), same ClassRoom.
            string srcYear  = MostCommon(selected.Select(r => r.AcademicYear));
            string srcGrade = MostCommon(selected.Select(r => r.GradeLevel));
            string srcRoom  = MostCommon(selected.Select(r => r.ClassRoom));
            string dstYear  = DB.NextYearString(srcYear);
            string dstGrade = isRepeat ? srcGrade : DB.NextGrade(srcGrade);
            string dstRoom  = srcRoom;
            string action   = isRepeat ? "ຊ້ຳຊັ້ນ" : "ຂຶ້ນຊັ້ນ";

            var dlg = new PromotionConfirmWindow(
                selected, srcYear, srcGrade, srcRoom, dstYear, dstGrade, dstRoom, action)
            {
                Owner = Window.GetWindow(this)
            };
            if (dlg.ShowDialog() != true) return;

            // User confirmed — execute the promotion using the dialog's final
            // destination Year/Grade/Room (which they may have edited).
            string finalY = dlg.DstYear;
            string finalG = dlg.DstGrade;
            string finalR = dlg.DstRoom;
            bool graduating = finalG == "ຈົບ";

            using var conn = DB.Open(); using var tx = conn.BeginTransaction();
            try
            {
                foreach (var row in selected)
                {
                    string toG, newSt, histTo, studentYear, studentRoom;
                    if (graduating)
                    {
                        // Student stays anchored to their finishing grade+year.
                        toG         = row.GradeLevel;
                        newSt       = "ຈົບ";
                        histTo      = "ຈົບ";
                        studentYear = row.AcademicYear;
                        studentRoom = row.ClassRoom;   // unchanged
                    }
                    else
                    {
                        toG         = finalG;
                        newSt       = "ກຳລັງຮຽນ";
                        histTo      = finalG;
                        studentYear = finalY;
                        studentRoom = finalR;
                    }

                    DB.ExecTx(@"UPDATE Students
                                SET GradeLevel=@g, ClassRoom=@cr, AcademicYear=@y, Status=@st
                                WHERE StudentID=@id",
                        conn, tx,
                        ("@g",  toG),         ("@cr", studentRoom),
                        ("@y",  studentYear), ("@st", newSt),
                        ("@id", row.StudentID));
                    DB.ExecTx(@"INSERT INTO GradeHistory
                                (StudentID, FromGrade, ToGrade, AcademicYear, ClassRoom, Note, ChangedBy)
                                VALUES (@sid, @fg, @tg, @y, @cr, @n, @by)",
                        conn, tx,
                        ("@sid", row.StudentID), ("@fg", row.GradeLevel), ("@tg", histTo),
                        ("@y",   finalY),        ("@cr", (object?)row.ClassRoom ?? DBNull.Value),
                        ("@n",   action),        ("@by", DB.CurrentUser));
                }
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ບໍ່ສຳເລັດ:\n{ex.Message}", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            DB.Log(isRepeat ? "Repeat" : "Promotion", $"{selected.Count} ຄົນ → {finalY}/{finalG}/{finalR}");
            MessageBox.Show($"{action} {selected.Count} ຄົນ ສຳເລັດ", "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);
            ReloadRoster(); LoadHistory();
        }

        private static string MostCommon(IEnumerable<string> vals)
        {
            var d = new Dictionary<string, int>();
            foreach (var v in vals)
            {
                if (string.IsNullOrEmpty(v)) continue;
                d.TryGetValue(v, out int c); d[v] = c + 1;
            }
            string best = ""; int bestN = 0;
            foreach (var kv in d) if (kv.Value > bestN) { best = kv.Key; bestN = kv.Value; }
            return best;
        }

        private void LoadHistory()
        {
            var dt = DB.Query(@"SELECT h.ChangedAt   AS ວັນທີ,
                                       s.StudentCode AS ລະຫັດ,
                                       s.FirstName||' '||s.LastName AS ຊື່ນັກຮຽນ,
                                       h.FromGrade   AS ຈາກຊັ້ນ,
                                       h.ToGrade     AS ຂຶ້ນຊັ້ນ,
                                       h.AcademicYear AS ປີ,
                                       IFNULL(h.ClassRoom,'') AS ຫ້ອງ,
                                       h.ChangedBy   AS ດຳເນີນໂດຍ
                                FROM GradeHistory h JOIN Students s ON s.StudentID=h.StudentID
                                ORDER BY h.HistoryID DESC LIMIT 50");
            _hist.ItemsSource = dt.DefaultView;
        }
    }

    public class PromRow : INotifyPropertyChanged
    {
        public int    StudentID    { get; set; }
        public string StudentCode  { get; set; } = "";
        public string FullName     { get; set; } = "";
        public string Gender       { get; set; } = "";
        public string GradeLevel   { get; set; } = "";
        public string ClassRoom    { get; set; } = "";
        public string AcademicYear { get; set; } = ""; // preserved when graduating
        public string Status       { get; set; } = "";
        private bool _sel;
        public bool IsSelected { get => _sel; set { if (_sel==value) return; _sel=value; Notify(); } }
        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify([System.Runtime.CompilerServices.CallerMemberName] string? n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    // ════════════════════════════════════════════════════════════
    //  PROMOTION CONFIRM WINDOW
    //
    //  Modal dialog opened by PromotionPage's "Promote Selected" /
    //  "Promote Entire Classroom" actions. Shows:
    //    - Current  : year/grade/room of the selection (read-only)
    //    - Destination: editable year/grade/room dropdowns
    //                   (defaults to NextYear / NextGrade / same room)
    //    - Preview  : per-student "Code | Name | From → To"
    //  DialogResult=true on confirm; PromotionPage reads DstYear/Grade/Room
    //  to run the actual DB update.
    // ════════════════════════════════════════════════════════════
    public class PromotionConfirmWindow : Window
    {
        public string DstYear  { get; private set; } = "";
        public string DstGrade { get; private set; } = "";
        public string DstRoom  { get; private set; } = "";

        private ComboBox _cmbDstYear  = null!;
        private ComboBox _cmbDstGrade = null!;
        private ComboBox _cmbDstRoom  = null!;
        private DataGrid _preview     = null!;
        private readonly List<PromRow> _selected;

        // `action` = ຂຶ້ນຊັ້ນ (default, standard promotion) or ຊ້ຳຊັ້ນ (retention).
        // Used for the window title, the "ກຳລັງຈະ…" header, and to hide the ‘ຈົບ’
        // option from the destination-grade dropdown when repeating (a repeat
        // path shouldn't also let the user graduate a student).
        public PromotionConfirmWindow(
            List<PromRow> selected,
            string srcYear, string srcGrade, string srcRoom,
            string dstYear, string dstGrade, string dstRoom,
            string action = "ຂຶ້ນຊັ້ນ")
        {
            _selected = selected;
            DstYear = dstYear; DstGrade = dstGrade; DstRoom = dstRoom;

            Title = $"⚠  ຢືນຢັນການ{action}";
            Width = 760; Height = 600;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252));

            var root = H.MkGrid(GridLength.Auto, GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // Header card
            var hdr = H.MkCard(new Thickness(12,12,12,8), new Thickness(14,10,14,10));
            hdr.Child = new TextBlock {
                Text = $"⚠  ກຳລັງຈະ{action} {selected.Count} ຄົນ",
                FontSize = 16, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            };
            Grid.SetRow(hdr, 0); root.Children.Add(hdr);

            // Current + Destination card
            var cur = H.MkCard(new Thickness(12,4,12,8), new Thickness(14,12,14,12));
            var stack = new StackPanel();
            stack.Children.Add(new TextBlock {
                Text = "ປະຈຸບັນ:", FontSize = 13, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            });
            stack.Children.Add(new TextBlock {
                Text = $"   ປີ {srcYear}   ·   ຊັ້ນ {srcGrade}   ·   ຫ້ອງ {srcRoom}",
                FontSize = 14, Margin = new Thickness(0,2,0,12),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(31,41,55))
            });

            stack.Children.Add(new TextBlock {
                Text = "ປ່ຽນເປັນ:", FontSize = 13, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            });
            var dstFlow = new WrapPanel { Margin = new Thickness(0,4,0,0) };
            // Year dropdown lists ONLY registered academic years. The caller
            // (PromotionPage.EnsureNextYearExists) has already guaranteed
            // dstYear is in the registry — no phantom-year insertion here.
            var years = DB.AcademicYears();
            _cmbDstYear  = H.MkCmb(years.ToArray(), 110);
            SelectComboValueDST(_cmbDstYear, dstYear);
            // Repeat flow keeps the student in school — never a graduation.
            string[] gradeChoices = action == "ຊ້ຳຊັ້ນ"
                ? new[] { "ມ.1", "ມ.2", "ມ.3", "ມ.4" }
                : new[] { "ມ.1", "ມ.2", "ມ.3", "ມ.4", "ຈົບ" };
            _cmbDstGrade = H.MkCmb(gradeChoices, 80);
            SelectComboValueDST(_cmbDstGrade, dstGrade);
            _cmbDstRoom  = H.MkCmb(new[] { "1", "2", "3", "4", "5", "6" }, 80);
            SelectComboValueDST(_cmbDstRoom, dstRoom);
            _cmbDstYear.SelectionChanged  += (s,e) => { DstYear  = _cmbDstYear.SelectedItem?.ToString() ?? DstYear;  RefreshPreview(); };
            _cmbDstGrade.SelectionChanged += (s,e) => { DstGrade = _cmbDstGrade.SelectedItem?.ToString() ?? DstGrade; RefreshPreview(); };
            _cmbDstRoom.SelectionChanged  += (s,e) => { DstRoom  = _cmbDstRoom.SelectedItem?.ToString() ?? DstRoom;  RefreshPreview(); };

            dstFlow.Children.Add(H.Lbl("ປີ:"));   dstFlow.Children.Add(_cmbDstYear);
            dstFlow.Children.Add(H.Lbl("ຊັ້ນ:")); dstFlow.Children.Add(_cmbDstGrade);
            dstFlow.Children.Add(H.Lbl("ຫ້ອງ:")); dstFlow.Children.Add(_cmbDstRoom);
            stack.Children.Add(dstFlow);
            cur.Child = stack; Grid.SetRow(cur, 1); root.Children.Add(cur);

            // Preview grid
            _preview = new DataGrid {
                AutoGenerateColumns = true, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White
            };
            var pcard = H.MkCard(new Thickness(12,4,12,8)); pcard.Child = _preview;
            Grid.SetRow(pcard, 2); root.Children.Add(pcard);

            // Action bar
            var actions = H.MkCard(new Thickness(12,4,12,12), new Thickness(12,8,12,8));
            var bar = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var bOk     = H.Btn("✅  ດຳເນີນການ", "PrimaryButton"); bOk.Click     += (s,e) => { DialogResult = true;  Close(); };
            var bCancel = H.Btn("ຍົກເລີກ",       "NeutralButton"); bCancel.Click += (s,e) => { DialogResult = false; Close(); };
            bar.Children.Add(bOk); bar.Children.Add(bCancel);
            actions.Child = bar; Grid.SetRow(actions, 3); root.Children.Add(actions);

            Content = root;
            RefreshPreview();
        }

        private static void SelectComboValueDST(ComboBox cb, string val)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == val) { cb.SelectedIndex = i; return; }
        }

        private void RefreshPreview()
        {
            var dt = new DataTable();
            dt.Columns.Add("ລະຫັດ", typeof(string));
            dt.Columns.Add("ຊື່ນັກຮຽນ", typeof(string));
            dt.Columns.Add("ຈາກ", typeof(string));
            dt.Columns.Add("ໄປ",  typeof(string));
            foreach (var r in _selected)
            {
                string from = $"{r.GradeLevel}/{r.ClassRoom}  ({r.AcademicYear})";
                string to   = DstGrade == "ຈົບ"
                    ? "ຈົບການສຶກສາ"
                    : $"{DstGrade}/{DstRoom}  ({DstYear})";
                dt.Rows.Add(r.StudentCode, r.FullName, from, to);
            }
            _preview.ItemsSource = dt.DefaultView;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  SUBJECTS PAGE  — filter + search + cascade-aware delete
    // ════════════════════════════════════════════════════════════
    public class SubjectsPage : UserControl
    {
        private DataGrid  _dg     = null!;
        private ComboBox  _fGrade = null!, _fSem = null!;
        private TextBox   _fSearch= null!;
        private TextBlock _lblCount = null!;
        private static readonly string[] AllGrades = { "ມ.1","ມ.2","ມ.3","ມ.4" };

        public SubjectsPage() { Build(); Load(); }

        private void Build()
        {
            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // ── Filter / action bar ──────────────────────────────
            var bar  = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,10,14,10));
            var flow = new WrapPanel();

            var bAdd  = H.Btn("➕  ເພີ່ມວິຊາ",   "SuccessButton"); bAdd.Click  += (s,e) => OpenForm(0);
            var bEdit = H.Btn("✏️  ແກ້ໄຂ",       "PrimaryButton"); bEdit.Click += (s,e) => EditSelected();
            var bDel  = H.Btn("🗑  ລຶບ",          "DangerButton");  bDel.Click  += (s,e) => DeleteSelected();
            flow.Children.Add(bAdd); flow.Children.Add(bEdit); flow.Children.Add(bDel);

            // Visual divider
            flow.Children.Add(new Border {
                Width = 1, Height = 24, Margin = new Thickness(8,4,8,0),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235))
            });

            flow.Children.Add(H.Lbl("ຊັ້ນ:"));
            _fGrade = new ComboBox { Width = 90, Margin = new Thickness(0,0,8,0) };
            _fGrade.Items.Add("ທັງໝົດ");
            foreach (var g in AllGrades) _fGrade.Items.Add(g);
            _fGrade.SelectedIndex = 0;
            _fGrade.SelectionChanged += (s,e) => Load();
            flow.Children.Add(_fGrade);

            flow.Children.Add(H.Lbl("ພາກ:"));
            _fSem = new ComboBox { Width = 80, Margin = new Thickness(0,0,8,0) };
            _fSem.Items.Add("ທັງໝົດ"); _fSem.Items.Add("1"); _fSem.Items.Add("2");
            _fSem.SelectedIndex = 0;
            _fSem.SelectionChanged += (s,e) => Load();
            flow.Children.Add(_fSem);

            flow.Children.Add(H.Lbl("ຄົ້ນຫາ:"));
            _fSearch = new TextBox { Width = 200, Margin = new Thickness(0,0,8,0) };
            _fSearch.TextChanged += (s,e) => Load();
            flow.Children.Add(_fSearch);

            _lblCount = new TextBlock {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8,0,0,0),
                FontSize = 12, FontWeight = FontWeights.Medium,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99))
            };
            flow.Children.Add(_lblCount);

            bar.Child = flow; Grid.SetRow(bar, 0); root.Children.Add(bar);

            // ── Data grid ────────────────────────────────────────
            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0),
                Background = System.Windows.Media.Brushes.White
            };
            _dg.Columns.Add(H.Col("ລຳດັບ",   "SortOrder",   70));
            _dg.Columns.Add(H.Col("ລະຫັດ",   "SubjectCode", 110));
            _dg.Columns.Add(H.ColStar("ຊື່ວິຊາ", "SubjectName"));
            _dg.Columns.Add(H.Col("ຊັ້ນ",      "GradeLevel",  70));
            _dg.Columns.Add(H.Col("ພາກ",      "Semester",    60));
            _dg.Columns.Add(H.Col("ປະເພດ",   "Category",   140));
            _dg.MouseDoubleClick += (s,e) => EditSelected();
            var card = H.MkCard(); card.Child = _dg; Grid.SetRow(card, 1); root.Children.Add(card);

            // ── Help footer ──────────────────────────────────────
            var hint = new Border {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252)),
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12,7,12,7),
                Margin = new Thickness(0,8,0,0)
            };
            hint.Child = new TextBlock {
                Text = "💡 ຫຼັງເພີ່ມວິຊາໃໝ່ ກະລຸນາໄປໜ້າ ‘ລົງທະບຽນ (Batch)’ ເພື່ອລົງທະບຽນວິຊາໃຫ້ນັກຮຽນ ຈຶ່ງຈະບັນທຶກຄະແນນໄດ້",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99)),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(hint, 2); root.Children.Add(hint);

            Content = root;
        }

        private void Load()
        {
            string g  = _fGrade.SelectedItem?.ToString() ?? "ທັງໝົດ";
            string sm = _fSem.SelectedItem?.ToString()   ?? "ທັງໝົດ";
            string kw = _fSearch.Text.Trim();

            var sb = new StringBuilder(@"SELECT SubjectID AS ID, SortOrder, SubjectCode, SubjectName,
                                                GradeLevel, Semester, Category
                                         FROM Subjects WHERE 1=1");
            var ps = new List<(string, object)>();
            // 13 official MoES subjects are grade-agnostic (GradeLevel='') and
            // semester-agnostic (Semester=0) — include them in every specific filter.
            if (g  != "ທັງໝົດ")
            {
                sb.Append(" AND (GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel='')");
                ps.Add(("@g",  g));
            }
            if (sm != "ທັງໝົດ")
            {
                sb.Append(" AND (Semester=@sm OR Semester=0)");
                ps.Add(("@sm", int.Parse(sm)));
            }
            if (!string.IsNullOrEmpty(kw))
            {
                sb.Append(" AND (SubjectCode LIKE @k OR SubjectName LIKE @k)");
                ps.Add(("@k", $"%{kw}%"));
            }
            sb.Append(" ORDER BY GradeLevel, SortOrder, SubjectCode");

            var dt = DB.Query(sb.ToString(), null, ps.ToArray());
            _dg.ItemsSource = dt.DefaultView;

            int total = DB.ScalarInt("SELECT COUNT(*) FROM Subjects");
            _lblCount.Text = $"📚 {dt.Rows.Count} ຈາກ {total} ວິຊາ";
        }

        private int SelId() => _dg.SelectedItem is DataRowView d ? Convert.ToInt32(d["ID"]) : 0;

        private void OpenForm(int id)
        {
            var f = new SubjectFormWin(id) { Owner = Window.GetWindow(this) };
            if (f.ShowDialog() == true) Load();
        }

        private void EditSelected()
        {
            int id = SelId();
            if (id <= 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກວິຊາທີ່ຕ້ອງການແກ້ໄຂກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            OpenForm(id);
        }

        private void DeleteSelected()
        {
            int id = SelId();
            if (id <= 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກວິຊາທີ່ຕ້ອງການລຶບກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Show subject name + count of dependent records that will cascade.
            var info = DB.Query("SELECT SubjectCode, SubjectName FROM Subjects WHERE SubjectID=@i", null, ("@i", id));
            string subjectLabel = info.Rows.Count > 0
                ? $"{info.Rows[0]["SubjectCode"]} — {info.Rows[0]["SubjectName"]}"
                : $"ID {id}";

            int enrollCount = DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE SubjectID=@i", null, ("@i", id));
            int scoreCount  = DB.ScalarInt("SELECT COUNT(*) FROM Scores WHERE EnrollID IN (SELECT EnrollID FROM Enrollments WHERE SubjectID=@i)", null, ("@i", id));
            int monthCount  = DB.ScalarInt("SELECT COUNT(*) FROM MonthlyAssessments WHERE EnrollID IN (SELECT EnrollID FROM Enrollments WHERE SubjectID=@i)", null, ("@i", id));

            string warn = $"ຕ້ອງການລຶບວິຊານີ້ບໍ?\n\n  •  {subjectLabel}\n";
            if (enrollCount + scoreCount + monthCount > 0)
            {
                warn += $"\n⚠ ການລຶບຈະຕັດຂໍ້ມູນຕໍ່ໄປນີ້ດ້ວຍ:\n" +
                        $"   • ການລົງທະບຽນ:        {enrollCount}\n" +
                        $"   • ຄະແນນພາກຮຽນ:      {scoreCount}\n" +
                        $"   • ຄະແນນລາຍເດືອນ:    {monthCount}\n\n" +
                        $"ການກະທຳນີ້ບໍ່ສາມາດຍົກເລີກໄດ້.";
            }

            if (MessageBox.Show(warn, "ຢືນຢັນການລຶບ", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                DB.Exec("DELETE FROM Subjects WHERE SubjectID=@i", null, ("@i", id));
                DB.Log("DelSubject", subjectLabel);
                Load();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ລຶບບໍ່ສຳເລັດ: {ex.Message}", "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // ════════════════════════════════════════════════════════════
    //  SUBJECT FORM  — Add / Edit a subject
    // ════════════════════════════════════════════════════════════
    public class SubjectFormWin : Window
    {
        private readonly int _id;
        private TextBox _code = null!, _name = null!, _sort = null!;
        private ComboBox _grade = null!, _sem = null!, _cat = null!;
        private static readonly string[] AllGrades = { "ມ.1","ມ.2","ມ.3","ມ.4" };
        private static readonly string[] AllCats   = { "ວິຊາສາມັນ", "ວິຊາເລືອກ", "ວິຊາສ້າງສັນ" };

        public SubjectFormWin(int id)
        {
            _id   = id;
            Title = id == 0 ? "➕  ເພີ່ມວິຊາໃໝ່" : "✏️  ແກ້ໄຂວິຊາ";
            Width = 460; Height = 460;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));
            UseLayoutRounding = true; SnapsToDevicePixels = true;

            var root  = H.MkGrid(new GridLength(1, GridUnitType.Star), GridLength.Auto);
            var card  = H.MkCard(new Thickness(16,16,16,8), new Thickness(20));
            var stack = new StackPanel();

            stack.Children.Add(FL("ລະຫັດວິຊາ *  (ຫ້າມຊ້ຳ — ຕົວຢ່າງ: MATH1, LAO2)"));
            _code = FI(stack);
            stack.Children.Add(FL("ຊື່ວິຊາ *"));
            _name = FI(stack);
            _grade = FC(stack, "ຊັ້ນ", AllGrades, defaultValue: "ມ.4");
            _sem   = FC(stack, "ພາກ", new[]{"1","2"});
            _cat   = FC(stack, "ປະເພດ", AllCats);
            stack.Children.Add(FL("ລຳດັບການສະແດງຜົນ  (ນ້ອຍ → ໃຫຍ່)"));
            _sort = FI(stack, "0");

            card.Child = stack;
            Grid.SetRow(card, 0); root.Children.Add(card);

            // Footer with action buttons
            var foot = new Border {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(0,1,0,0),
                Padding = new Thickness(14,10,14,10)
            };
            var fp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var bOK = H.Btn("💾  ບັນທຶກ", "SuccessButton"); bOK.Width = 110; bOK.Click += Save;
            var bCancel = new Button { Content = "ຍົກເລີກ", Width = 90, Height = 34, Margin = new Thickness(8,0,0,0), IsCancel = true };
            bCancel.SetResourceReference(Button.StyleProperty, "NeutralButton");
            fp.Children.Add(bOK); fp.Children.Add(bCancel);
            foot.Child = fp;
            Grid.SetRow(foot, 1); root.Children.Add(foot);

            Content = root;
            if (id > 0) LoadData();
        }

        private void LoadData()
        {
            var dt = DB.Query("SELECT * FROM Subjects WHERE SubjectID=@id", null, ("@id", _id));
            if (dt.Rows.Count == 0) return;
            var r = dt.Rows[0];
            _code.Text = r["SubjectCode"].ToString() ?? "";
            _name.Text = r["SubjectName"].ToString() ?? "";
            SC(_grade, r["GradeLevel"].ToString() ?? "");
            SC(_sem,   r["Semester"].ToString()   ?? "1");
            SC(_cat,   r["Category"].ToString()   ?? "");
            _sort.Text = r["SortOrder"].ToString() ?? "0";
        }

        private void Save(object s, RoutedEventArgs e)
        {
            string code = _code.Text.Trim();
            string name = _name.Text.Trim();

            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(name))
            { MessageBox.Show("ກະລຸນາໃສ່ ລະຫັດ ແລະ ຊື່ວິຊາ", "ຂໍ້ມູນບໍ່ຄົບ", MessageBoxButton.OK, MessageBoxImage.Information); return; }
            if (code.Contains(' '))
            { MessageBox.Show("ລະຫັດວິຊາຫ້າມມີຊ່ອງວ່າງ", "ຂໍ້ມູນບໍ່ຖືກຕ້ອງ", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            int.TryParse(_sort.Text, out int so);

            try
            {
                string sql = _id == 0
                    ? @"INSERT INTO Subjects(SubjectCode,SubjectName,GradeLevel,Semester,Category,SortOrder)
                        VALUES(@c,@n,@g,@sm,@cat,@so)"
                    : @"UPDATE Subjects
                          SET SubjectCode=@c, SubjectName=@n, GradeLevel=@g,
                              Semester=@sm, Category=@cat, SortOrder=@so
                        WHERE SubjectID=@id";

                var ps = new List<(string, object)>
                {
                    ("@c",   code),
                    ("@n",   name),
                    ("@g",   GC(_grade)),
                    ("@sm",  int.Parse(GC(_sem) is { Length: > 0 } v ? v : "1")),
                    ("@cat", GC(_cat)),
                    ("@so",  so)
                };
                if (_id > 0) ps.Add(("@id", _id));
                DB.Exec(sql, null, ps.ToArray());
                DB.Log(_id == 0 ? "AddSubject" : "EditSubject", $"{code} — {name}");
                DialogResult = true; Close();
            }
            catch (SQLiteException ex) when (ex.Message.Contains("UNIQUE"))
            { MessageBox.Show($"ລະຫັດວິຊາ ‘{code}’ ມີຢູ່ແລ້ວ — ກະລຸນາໃຊ້ລະຫັດອື່ນ", "ລະຫັດຊ້ຳ", MessageBoxButton.OK, MessageBoxImage.Warning); }
            catch (Exception ex)
            { MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ: {ex.Message}", "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        // ── Layout helpers ─────────────────────────────────────────
        private static TextBlock FL(string t) => new TextBlock {
            Text = t, FontSize = 12, FontWeight = FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100,116,139)),
            Margin = new Thickness(0,0,0,6)
        };
        private static TextBox FI(StackPanel p, string v = "")
        {
            var tb = new TextBox { Text = v, Margin = new Thickness(0,0,0,12) };
            p.Children.Add(tb); return tb;
        }
        private static ComboBox FC(StackPanel p, string lbl, string[] items, string? defaultValue = null)
        {
            p.Children.Add(FL(lbl));
            var cb = new ComboBox { Margin = new Thickness(0,0,0,12) };
            foreach (var i in items) cb.Items.Add(i);
            int defaultIdx = 0;
            if (defaultValue != null)
                for (int i = 0; i < items.Length; i++)
                    if (items[i] == defaultValue) { defaultIdx = i; break; }
            cb.SelectedIndex = defaultIdx;
            p.Children.Add(cb);
            return cb;
        }
        private static void SC(ComboBox cb, string v)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == v) { cb.SelectedIndex = i; return; }
        }
        private static string GC(ComboBox cb) => cb.SelectedItem?.ToString() ?? "";
    }

    // ════════════════════════════════════════════════════════════
    //  USERS PAGE
    // ════════════════════════════════════════════════════════════
    public class UsersPage : UserControl
    {
        private DataGrid _dg=null!;
        public UsersPage(){Build();Load();}
        private void Build(){var root=H.MkGrid(GridLength.Auto,new GridLength(1,GridUnitType.Star));var tb=H.MkCard(new Thickness(0,0,0,10),new Thickness(14,10,14,10));var flow=new WrapPanel();var bAdd=H.Btn("➕  ເພີ່ມ","SuccessButton");bAdd.Click+=(s,e)=>{var f=new UserFormWin(0){Owner=Window.GetWindow(this)};if(f.ShowDialog()==true)Load();};var bEdit=H.Btn("✏️  ແກ້ໄຂ","PrimaryButton");bEdit.Click+=(s,e)=>{if(SelId()>0){var f=new UserFormWin(SelId()){Owner=Window.GetWindow(this)};if(f.ShowDialog()==true)Load();}};var bToggle=H.Btn("🔒  ເປີດ/ປິດ","WarningButton");bToggle.Click+=(s,e)=>{int uid=SelId();if(uid<=0){MessageBox.Show("ກະລຸນາເລືອກຜູ້ໃຊ້ກ່ອນ","ຍັງບໍ່ໄດ້ເລືອກ",MessageBoxButton.OK,MessageBoxImage.Information);return;}if(uid==DB.CurrentUserId){MessageBox.Show("ບໍ່ສາມາດປິດບັນຊີຕົນເອງ");return;}var info=DB.Query("SELECT Username,IsActive FROM Users WHERE UserID=@i",null,("@i",uid));if(info.Rows.Count==0)return;string uname=info.Rows[0]["Username"].ToString()!;bool wasActive=Convert.ToInt32(info.Rows[0]["IsActive"])==1;DB.Exec("UPDATE Users SET IsActive=CASE IsActive WHEN 1 THEN 0 ELSE 1 END WHERE UserID=@i",null,("@i",uid));DB.Log(wasActive?"DeactivateUser":"ActivateUser",uname);Load();};flow.Children.Add(bAdd);flow.Children.Add(bEdit);flow.Children.Add(bToggle);tb.Child=flow;Grid.SetRow(tb,0);root.Children.Add(tb);_dg=new DataGrid{AutoGenerateColumns=true,IsReadOnly=true,CanUserAddRows=false,BorderThickness=new Thickness(0),Background=System.Windows.Media.Brushes.White};var card=H.MkCard();card.Child=_dg;Grid.SetRow(card,1);root.Children.Add(card);Content=root;}
        private void Load(){var dt=DB.Query("SELECT UserID AS ID,Username AS ຊື່ຜູ້ໃຊ້,FullName AS ຊື່ເຕັມ,Role AS ບົດບາດ,CASE IsActive WHEN 1 THEN 'ໃຊ້ງານ' ELSE 'ປິດ' END AS ສະຖານະ,IFNULL(LastLogin,'ຍັງບໍ່ເຄີຍ') AS ເຂົ້າໃຊ້ລ່າສຸດ FROM Users ORDER BY Role,Username");_dg.ItemsSource=dt.DefaultView;}
        private int SelId()=>_dg.SelectedItem is DataRowView d?Convert.ToInt32(d["ID"]):0;
    }

    public class UserFormWin : Window
    {
        private const int MinPwdLength = 4;

        private readonly int _id;
        private TextBox _user = null!, _name = null!;
        private PasswordBox _pwdNew = null!, _pwdConfirm = null!;
        private ComboBox _role = null!;

        public UserFormWin(int id)
        {
            _id = id;
            Title = id == 0 ? "➕  ເພີ່ມຜູ້ໃຊ້" : "✏️  ແກ້ໄຂຜູ້ໃຊ້";
            Width = 420; Height = 480;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = System.Windows.Media.Brushes.White;

            var stack = new StackPanel { Margin = new Thickness(20) };

            // ── Section 1: User identity ──
            stack.Children.Add(SectionHeader("ຂໍ້ມູນຜູ້ໃຊ້"));
            stack.Children.Add(FieldLabel("ຊື່ຜູ້ໃຊ້ *"));
            _user = FieldInput(stack);
            stack.Children.Add(FieldLabel("ຊື່ເຕັມ *"));
            _name = FieldInput(stack);
            stack.Children.Add(FieldLabel("ບົດບາດ"));
            _role = new ComboBox { Margin = new Thickness(0, 0, 0, 14) };
            _role.Items.Add("admin");
            _role.Items.Add("teacher");
            _role.SelectedIndex = 1;
            stack.Children.Add(_role);

            // ── Section 2: Password ──
            // For Add (id==0): both fields required + min length.
            // For Edit (id>0): leave both blank to keep the existing password;
            // type into both to change it. Confirm field prevents typos that
            // would otherwise lock the user out of the system.
            string pwdHeader = id == 0
                ? "ລະຫັດຜ່ານ"
                : $"ປ່ຽນລະຫັດຜ່ານ (ເວັ້ນຫາກບໍ່ປ່ຽນ)";
            stack.Children.Add(SectionHeader(pwdHeader));
            stack.Children.Add(FieldLabel(id == 0
                ? $"ລະຫັດຜ່ານ * (ຢ່າງໜ້ອຍ {MinPwdLength} ຕົວ)"
                : $"ລະຫັດຜ່ານໃໝ່ (ຢ່າງໜ້ອຍ {MinPwdLength} ຕົວ)"));
            _pwdNew = new PasswordBox { Margin = new Thickness(0, 0, 0, 10) };
            stack.Children.Add(_pwdNew);
            stack.Children.Add(FieldLabel("ຢືນຢັນລະຫັດຜ່ານ"));
            _pwdConfirm = new PasswordBox { Margin = new Thickness(0, 0, 0, 16) };
            stack.Children.Add(_pwdConfirm);

            // ── Buttons ──
            var fp = new StackPanel {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            string okLabel = id == 0 ? "💾  ສ້າງຜູ້ໃຊ້" : "💾  ບັນທຶກການປ່ຽນແປງ";
            var bOK = H.Btn(okLabel, "SuccessButton");
            bOK.MinWidth = 180;
            bOK.Click += Save;
            var bCan = new Button {
                Content = "ຍົກເລີກ", Width = 80, Height = 34,
                Margin = new Thickness(8, 0, 0, 0), IsCancel = true
            };
            bCan.SetResourceReference(Button.StyleProperty, "NeutralButton");
            fp.Children.Add(bOK); fp.Children.Add(bCan);
            stack.Children.Add(fp);
            Content = stack;

            if (id > 0) LoadData();
            // Focus the username field so it's clear where to start typing.
            Loaded += (_, _) => _user.Focus();
        }

        private void LoadData()
        {
            var dt = DB.Query("SELECT Username, FullName, Role FROM Users WHERE UserID=@id",
                null, ("@id", _id));
            if (dt.Rows.Count == 0) return;
            var r = dt.Rows[0];
            _user.Text = r["Username"].ToString() ?? "";
            _name.Text = r["FullName"].ToString() ?? "";
            // Pick the matching role item explicitly (don't rely on .Text setter for non-editable combo).
            string role = r["Role"].ToString() ?? "teacher";
            for (int i = 0; i < _role.Items.Count; i++)
                if (string.Equals(_role.Items[i]?.ToString(), role, StringComparison.OrdinalIgnoreCase))
                { _role.SelectedIndex = i; break; }
        }

        private void Save(object s, RoutedEventArgs e)
        {
            // ── Validation ──
            string uname = _user.Text.Trim();
            string fname = _name.Text.Trim();
            string role  = _role.SelectedItem?.ToString() ?? "teacher";
            string pwdNew = _pwdNew.Password;
            string pwdConf = _pwdConfirm.Password;
            bool hasPwd  = !string.IsNullOrEmpty(pwdNew) || !string.IsNullOrEmpty(pwdConf);

            if (string.IsNullOrEmpty(uname))
            {
                MessageBox.Show("ກະລຸນາໃສ່ຊື່ຜູ້ໃຊ້", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _user.Focus(); return;
            }
            if (string.IsNullOrEmpty(fname))
            {
                MessageBox.Show("ກະລຸນາໃສ່ຊື່ເຕັມ", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _name.Focus(); return;
            }
            // Add: password mandatory. Edit: password optional but if any field
            // is filled both must be filled and must match.
            if (_id == 0 && !hasPwd)
            {
                MessageBox.Show("ກະລຸນາໃສ່ລະຫັດຜ່ານ", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _pwdNew.Focus(); return;
            }
            if (hasPwd)
            {
                if (pwdNew.Length < MinPwdLength)
                {
                    MessageBox.Show($"ລະຫັດຜ່ານຕ້ອງມີຢ່າງໜ້ອຍ {MinPwdLength} ຕົວ",
                        "ລະຫັດຜ່ານສັ້ນເກີນໄປ", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _pwdNew.Focus(); return;
                }
                if (pwdNew != pwdConf)
                {
                    MessageBox.Show("ລະຫັດຜ່ານໃໝ່ ແລະ ການຢືນຢັນບໍ່ກົງກັນ",
                        "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _pwdConfirm.Clear();
                    _pwdConfirm.Focus(); return;
                }
            }

            // ── Persist ──
            try
            {
                if (_id == 0)
                {
                    DB.Exec(@"INSERT INTO Users(Username, Password, FullName, Role)
                              VALUES(@u, @p, @n, @r)",
                        null,
                        ("@u", uname), ("@p", pwdNew),
                        ("@n", fname), ("@r", role));
                }
                else if (hasPwd)
                {
                    DB.Exec(@"UPDATE Users
                              SET Username=@u, Password=@p, FullName=@n, Role=@r
                              WHERE UserID=@id",
                        null,
                        ("@u", uname), ("@p", pwdNew),
                        ("@n", fname), ("@r", role),
                        ("@id", _id));
                }
                else
                {
                    DB.Exec(@"UPDATE Users
                              SET Username=@u, FullName=@n, Role=@r
                              WHERE UserID=@id",
                        null,
                        ("@u", uname), ("@n", fname), ("@r", role),
                        ("@id", _id));
                }
            }
            catch (SQLiteException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show($"ຊື່ຜູ້ໃຊ້ ‘{uname}’ ມີຢູ່ແລ້ວ — ກະລຸນາໃຊ້ຊື່ອື່ນ",
                    "ຊື່ຊ້ຳ", MessageBoxButton.OK, MessageBoxImage.Warning);
                _user.SelectAll(); _user.Focus();
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // ── Verify the write actually landed (round-trip read-back) ──
            // This catches silent failures (e.g. param binding bugs, file locks) and
            // gives the admin certainty that the credentials they just set will work.
            int verifyId = _id;
            if (verifyId == 0)
                verifyId = DB.ScalarInt("SELECT UserID FROM Users WHERE Username=@u",
                    null, ("@u", uname));

            var check = DB.Query(@"SELECT Username, FullName, Role FROM Users WHERE UserID=@id",
                null, ("@id", verifyId));
            bool ok = check.Rows.Count == 1
                   && check.Rows[0]["Username"].ToString() == uname
                   && check.Rows[0]["FullName"].ToString() == fname
                   && string.Equals(check.Rows[0]["Role"].ToString(), role,
                                    StringComparison.OrdinalIgnoreCase);

            // Password round-trip: try to authenticate-style match.
            bool pwdOk = !hasPwd ||
                DB.ScalarInt("SELECT COUNT(*) FROM Users WHERE UserID=@id AND Password=@p",
                    null, ("@id", verifyId), ("@p", pwdNew)) == 1;

            if (!ok || !pwdOk)
            {
                MessageBox.Show(
                    "ບັນທຶກສຳເລັດ ແຕ່ກວດສອບແລ້ວຂໍ້ມູນບໍ່ຕົງ — ກະລຸນາລອງໃໝ່ ຫຼື ກວດກາສິດການຂຽນຖານຂໍ້ມູນ.",
                    "ຄຳເຕືອນ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ── Log + success message ──
            string action;
            string detail;
            if (_id == 0)
            {
                action = "AddUser"; detail = uname;
            }
            else if (hasPwd)
            {
                action = "EditUserWithPwd";
                detail = $"{uname} (password updated)";
            }
            else
            {
                action = "EditUser"; detail = uname;
            }
            DB.Log(action, detail);

            string successMsg = _id == 0
                ? $"ສ້າງຜູ້ໃຊ້ ‘{uname}’ ສຳເລັດ"
                : (hasPwd
                    ? $"ບັນທຶກສຳເລັດ:\n   • ຂໍ້ມູນຜູ້ໃຊ້ປ່ຽນແປງແລ້ວ\n   • ລະຫັດຜ່ານປ່ຽນແປງແລ້ວ\n\n" +
                      "ຄຳແນະນຳ: ໃຫ້ຜູ້ໃຊ້ logout ແລ້ວ login ໃໝ່ດ້ວຍລະຫັດຜ່ານໃໝ່."
                    : $"ບັນທຶກຂໍ້ມູນ ‘{uname}’ ສຳເລັດ");
            MessageBox.Show(successMsg, "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);

            DialogResult = true;
            Close();
        }

        // ── Layout helpers ──
        private static TextBlock SectionHeader(string text) => new TextBlock {
            Text = text,
            FontSize = 13, FontWeight = FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(27, 79, 138)),
            Margin = new Thickness(0, 6, 0, 8),
        };
        private static TextBlock FieldLabel(string t) => new TextBlock {
            Text = t, FontSize = 12,
            Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107, 114, 128)),
            Margin = new Thickness(0, 0, 0, 4)
        };
        private static TextBox FieldInput(StackPanel p) {
            var tb = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            p.Children.Add(tb);
            return tb;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  ACADEMIC YEAR MANAGEMENT
    //  Admin-only catalogue of every academic year the school uses.
    //  Lets the admin: add a year in advance, switch which year is "current",
    //  or advance to the next year in one click. Never touches existing
    //  Students / Enrollments / Scores records.
    // ════════════════════════════════════════════════════════════
    // ════════════════════════════════════════════════════════════
    //  ACADEMIC YEAR PAGE — redesigned for safety + consistency
    //
    //  Layout (top → bottom):
    //    1. Current-year header card: 📅 year, 📘 semester, 🟢 active
    //       badge, plus 4 global actions (Add / Move-to-Next / Set-Current
    //       / Delete).
    //    2. Year-list DataGrid with 7 spec columns +
    //       1 per-row 📊 ສະຖິຕິ button (opens AcademicYearStatsWindow).
    //    3. Footer stats card: system-wide totals.
    //
    //  Consistency rules respected:
    //    - Year switching goes ONLY through DB.SetCurrentAcademicYear
    //      (no custom UPDATE Settings / AcademicYears in this file).
    //    - New years go through DB.CreateAcademicYear.
    //    - "Next year" derivation uses DB.NextYearString.
    //
    //  Delete flow (one consolidated window):
    //    - AcademicYearDeleteWindow shows row-by-row counts of every
    //      cascade-deletable table and requires the user to type "DELETE"
    //      before its 🔥 Force Delete button enables. Cancel is always
    //      available. Force Delete runs in ONE transaction; failure rolls
    //      back. Current year is refused outright.
    //
    //  No automatic student promotion — that's a separate workflow on the
    //  ການຂຶ້ນຊັ້ນ / ຈົບ page. "Move to Next" only creates the new year row
    //  + switches the current-year setting + resets semester to 1.
    // ════════════════════════════════════════════════════════════
    public class AcademicYearPage : UserControl
    {
        private DataGrid  _dg = null!;
        private TextBlock _curHeaderTxt = null!, _statsTxt = null!;

        public AcademicYearPage() { Build(); Load(); }

        private void Build()
        {
            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // ── Row 0: Current-year header card + global actions ────────
            var hdr = H.MkCard(new Thickness(0,0,0,10), new Thickness(16,12,16,12));
            hdr.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239,246,255));
            var hdrStack = new StackPanel();

            // Top line: 📅 year + ພາກ in one combined line + 🟢 badge
            var topLine = new WrapPanel();
            topLine.Children.Add(new TextBlock {
                Text = "📅  ປີການສຶກສາປະຈຸບັນ: ",
                FontSize = 14, VerticalAlignment = VerticalAlignment.Center,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            });
            _curHeaderTxt = new TextBlock {
                FontSize = 18, FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0,0,18,0),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(27,79,138))
            };
            topLine.Children.Add(_curHeaderTxt);
            var badge = new Border {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220,252,231)),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(10,3,10,3),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock {
                    Text = "🟢 ກຳລັງໃຊ້",
                    FontSize = 12, FontWeight = FontWeights.SemiBold,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(22,101,52))
                }
            };
            topLine.Children.Add(badge);
            hdrStack.Children.Add(topLine);

            // Action buttons
            var actFlow = new WrapPanel { Margin = new Thickness(0,12,0,0) };
            var bAdd     = H.Btn("➕  ເພີ່ມປີໃໝ່",       "SuccessButton");
            var bAdv     = H.Btn("➡️  ຂຶ້ນປີຕໍ່ໄປ",      "WarningButton");
            var bSetCurr = H.Btn("🔄  ຕັ້ງເປັນປະຈຸບັນ",  "PrimaryButton");
            var bDelete  = H.Btn("🗑️  ລຶບປີ",            "DangerButton");
            bAdd.Click     += (s,e) => OpenAddDialog();
            bAdv.Click     += (s,e) => MoveToNextYear();
            bSetCurr.Click += (s,e) => SetSelectedAsCurrent();
            bDelete.Click  += (s,e) => DeleteSelectedYear();
            actFlow.Children.Add(bAdd); actFlow.Children.Add(bAdv);
            actFlow.Children.Add(bSetCurr); actFlow.Children.Add(bDelete);
            hdrStack.Children.Add(actFlow);

            hdr.Child = hdrStack;
            Grid.SetRow(hdr, 0); root.Children.Add(hdr);

            // ── Row 1: Year-list DataGrid (7 cols + per-row stats button) ──
            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0),
                Background = System.Windows.Media.Brushes.White,
                SelectionMode = DataGridSelectionMode.Single
            };
            _dg.Columns.Add(H.Col("ປີການສຶກສາ",   "Year",           120));
            _dg.Columns.Add(H.Col("ນັກຮຽນ",       "StudentCount",    90));
            _dg.Columns.Add(H.Col("ຈົບ",          "GraduatedCount",  70));
            _dg.Columns.Add(H.Col("ກຳລັງຮຽນ",     "ActiveCount",     90));
            _dg.Columns.Add(H.Col("ວັນທີສ້າງ",     "CreatedDate",    110));
            _dg.Columns.Add(H.ColStar("ສະຖານະ",   "StatusLabel",   true));
            _dg.Columns.Add(MkStatsCol());
            _dg.AddHandler(Button.ClickEvent, new RoutedEventHandler(OnStatsClick));
            _dg.MouseDoubleClick += (s,e) => SetSelectedAsCurrent();
            var card = H.MkCard(); card.Child = _dg;
            Grid.SetRow(card, 1); root.Children.Add(card);

            // ── Row 2: Footer system-wide stats ─────────────────────────
            var statCard = new Border {
                Background  = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(254,252,232)),
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(252,211,77)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(14,10,14,10),
                Margin  = new Thickness(0,10,0,0)
            };
            _statsTxt = new TextBlock {
                FontSize = 13, FontWeight = FontWeights.Medium,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(120,53,15)),
                TextWrapping = TextWrapping.Wrap
            };
            statCard.Child = _statsTxt;
            Grid.SetRow(statCard, 2); root.Children.Add(statCard);

            Content = root;
        }

        // Per-row stats button column — single button, click bubbles to OnStatsClick.
        private static DataGridTemplateColumn MkStatsCol()
        {
            var col = new DataGridTemplateColumn { Header = "", Width = new DataGridLength(110) };
            var fac = new System.Windows.FrameworkElementFactory(typeof(Button));
            fac.SetValue(Button.ContentProperty, "📊 ສະຖິຕິ");
            fac.SetValue(Button.HeightProperty, 26.0);
            fac.SetValue(Button.MarginProperty, new Thickness(2));
            fac.SetValue(Button.CursorProperty, System.Windows.Input.Cursors.Hand);
            fac.SetValue(Button.TagProperty, "stats");
            fac.SetResourceReference(Button.StyleProperty, "NeutralButton");
            col.CellTemplate = new System.Windows.DataTemplate { VisualTree = fac };
            return col;
        }

        private void OnStatsClick(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is not Button btn) return;
            if (btn.Tag as string != "stats") return;
            if (btn.DataContext is not DataRowView drv) return;
            string y = drv["Year"].ToString() ?? "";
            if (string.IsNullOrEmpty(y)) return;
            new AcademicYearStatsWindow(y) { Owner = Window.GetWindow(this) }.ShowDialog();
        }

        // Reload — fires only on actual data-changing actions; the page does NOT
        // auto-refresh on every selection change.
        private void Load()
        {
            RefreshHeader();
            RefreshGrid();
            RefreshFooterStats();
        }

        private void RefreshHeader()
        {
            _curHeaderTxt.Text = DB.CurrentYear;
        }

        private void RefreshGrid()
        {
            // Semester column shows only for the current year (past/future years
            // don't have an active semester). Student counts are split by Status.
            // CreatedDate is the date portion of CreatedAt for compactness.
            var dt = DB.Query(@"
                SELECT y.Year,
                       (SELECT COUNT(*) FROM Students s WHERE s.AcademicYear=y.Year)                                AS StudentCount,
                       (SELECT COUNT(*) FROM Students s WHERE s.AcademicYear=y.Year AND s.Status='ຈົບ')             AS GraduatedCount,
                       (SELECT COUNT(*) FROM Students s WHERE s.AcademicYear=y.Year AND s.Status='ກຳລັງຮຽນ')        AS ActiveCount,
                       IFNULL(SUBSTR(y.CreatedAt,1,10),'')                  AS CreatedDate,
                       CASE y.IsCurrent WHEN 1 THEN '🟢 ປະຈຸບັນ' ELSE '⚪' END AS StatusLabel,
                       y.IsCurrent                                          AS IsCurrent
                FROM AcademicYears y
                ORDER BY y.Year DESC");
            _dg.ItemsSource = dt.DefaultView;
        }

        private void RefreshFooterStats()
        {
            int total      = DB.ScalarInt("SELECT COUNT(*) FROM Students");
            int active     = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ກຳລັງຮຽນ'");
            int graduated  = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ຈົບ'");
            int withdrawn  = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE Status='ອອກ'");
            int classrooms = DB.ScalarInt(@"SELECT COUNT(DISTINCT GradeLevel||'/'||IFNULL(ClassRoom,''))
                                            FROM Students
                                            WHERE GradeLevel IS NOT NULL AND GradeLevel<>''");
            _statsTxt.Text =
                $"📊  ນັກຮຽນທັງໝົດ: {total}   ·   " +
                $"ກຳລັງຮຽນ: {active}   ·   " +
                $"ຈົບ: {graduated}   ·   " +
                $"ອອກ: {withdrawn}   ·   " +
                $"ຫ້ອງຮຽນ: {classrooms}";
        }

        private string? SelectedYear() =>
            _dg.SelectedItem is DataRowView d ? d["Year"].ToString() : null;
        private bool SelectedIsCurrent() =>
            _dg.SelectedItem is DataRowView d && Convert.ToInt32(d["IsCurrent"]) == 1;

        // ── Actions ──────────────────────────────────────────────────────
        // Add: only creates a row in AcademicYears — does NOT auto-promote
        // students or switch the current year.
        private void OpenAddDialog()
        {
            var dlg = new AcademicYearFormWin { Owner = Window.GetWindow(this) };
            if (dlg.ShowDialog() == true) Load();
        }

        // Move to Next: auto-creates the next year if missing, switches it to current,
        // resets semester to 1. Routed through DB.SetCurrentAcademicYear so the
        // Settings + AcademicYears.IsCurrent invariants are honored.
        private void MoveToNextYear()
        {
            string current = DB.CurrentYear;
            string next    = DB.NextYearString(current);
            if (next == current)
            {
                MessageBox.Show("ບໍ່ສາມາດກຳນົດປີຕໍ່ໄປໄດ້ — ກວດສອບຮູບແບບປີປະຈຸບັນ (YYYY-YYYY)",
                                "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string msg =
                $"ຕ້ອງການເລື່ອນລະບົບໄປປີໃໝ່ບໍ?\n\n" +
                $"   ປະຈຸບັນ:   {current}\n" +
                $"   ໃໝ່:       {next}\n\n" +
                "✅ ຂໍ້ມູນນັກຮຽນ + ຄະແນນທັງໝົດຄົງຢູ່.\n" +
                "ℹ ການເລື່ອນຊັ້ນຂອງນັກຮຽນແມ່ນເຮັດແຍກຢູ່ໜ້າ ‘ຂຶ້ນຊັ້ນ / ຈົບ’.";
            if (MessageBox.Show(msg, "ຢືນຢັນຂຶ້ນປີໃໝ່",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;
            try { DB.CreateAcademicYear(next, null, null, "ສ້າງອັດຕະໂນມັດເມື່ອຂຶ້ນປີໃໝ່"); }
            catch { /* already exists — fine */ }
            DB.SetCurrentAcademicYear(next);
            Load();
            MessageBox.Show($"ສຳເລັດ — ປະຈຸບັນແມ່ນປີ {next}",
                            "ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Set Current: thin wrapper around DB.SetCurrentAcademicYear. The helper
        // updates Settings.current_year + AcademicYears.IsCurrent atomically and
        // guarantees exactly one current year exists.
        private void SetSelectedAsCurrent()
        {
            var y = SelectedYear();
            if (string.IsNullOrEmpty(y))
            {
                MessageBox.Show("ກະລຸນາເລືອກປີຈາກຕາຕະລາງກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (y == DB.CurrentYear)
            {
                MessageBox.Show($"ປີ {y} ເປັນປະຈຸບັນຢູ່ແລ້ວ", "ບໍ່ມີການປ່ຽນແປງ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (MessageBox.Show(
                    $"ຕັ້ງປີ {y} ເປັນປະຈຸບັນບໍ?\n\n" +
                    "ການກະທຳນີ້ປ່ຽນແຕ່ການກຳນົດປີໃນລະບົບ — ຂໍ້ມູນທີ່ບັນທຶກໄວ້ບໍ່ຖືກກະທົບ.",
                    "ຢືນຢັນ", MessageBoxButton.YesNo, MessageBoxImage.Question)
                != MessageBoxResult.Yes) return;
            DB.SetCurrentAcademicYear(y);
            Load();
        }

        // Delete: opens the unified AcademicYearDeleteWindow which handles both
        // safe-delete (no data → simple click) and force-delete (data present →
        // requires typing "DELETE"). Page just kicks off the cascade transaction
        // on confirm. Current year is always refused.
        private void DeleteSelectedYear()
        {
            var y = SelectedYear();
            if (string.IsNullOrEmpty(y))
            {
                MessageBox.Show("ກະລຸນາເລືອກປີຈາກຕາຕະລາງກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (SelectedIsCurrent())
            {
                MessageBox.Show("ບໍ່ສາມາດລຶບປີທີ່ເປັນປະຈຸບັນໄດ້ — ໃຫ້ປ່ຽນປະຈຸບັນຫາປີອື່ນກ່ອນ",
                                "ບໍ່ສາມາດລຶບ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dlg = new AcademicYearDeleteWindow(y) { Owner = Window.GetWindow(this) };
            if (dlg.ShowDialog() != true) return;

            // Cascade delete in one transaction. Order respects FK dependencies:
            //   Enrollments → cascades Scores + MonthlyAssessments via FK
            //   EvaluationScores / AttendanceRecords / GradeHistory by AcademicYear
            //   Students(year) → cascades anything else they own
            //   AcademicYears(year) last
            using var conn = DB.Open(); using var tx = conn.BeginTransaction();
            try
            {
                DB.ExecTx("DELETE FROM Enrollments       WHERE AcademicYear=@y", conn, tx, ("@y", y));
                DB.ExecTx("DELETE FROM EvaluationScores  WHERE AcademicYear=@y", conn, tx, ("@y", y));
                DB.ExecTx("DELETE FROM AttendanceRecords WHERE AcademicYear=@y", conn, tx, ("@y", y));
                DB.ExecTx("DELETE FROM GradeHistory      WHERE AcademicYear=@y", conn, tx, ("@y", y));
                DB.ExecTx("DELETE FROM Students          WHERE AcademicYear=@y", conn, tx, ("@y", y));
                DB.ExecTx("DELETE FROM AcademicYears     WHERE Year=@y",         conn, tx, ("@y", y));
                tx.Commit();
            }
            catch (Exception ex)
            {
                tx.Rollback();
                MessageBox.Show($"ລຶບບໍ່ສຳເລັດ — ຂໍ້ມູນຍັງຢູ່ຄົບ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            DB.Log("DelAcademicYear", y);
            MessageBox.Show($"ລຶບປີ {y} ສຳເລັດ", "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);
            Load();
        }
    }

    // ════════════════════════════════════════════════════════════
    //  ACADEMIC YEAR DELETE WINDOW
    //
    //  Unified replacement for the old two-step MessageBox + ConfirmTypeWindow
    //  flow. Shows a row-by-row count of every cascade-deletable table for
    //  the picked year, then either:
    //    - all counts == 0 → just type "DELETE" + click Confirm
    //    - any count > 0   → 🔥 Force Delete (requires typing "DELETE")
    //                        with a prominent warning
    //  Cancel is always available. DialogResult=true = caller should proceed
    //  with the cascade DELETE transaction.
    // ════════════════════════════════════════════════════════════
    public class AcademicYearDeleteWindow : Window
    {
        public AcademicYearDeleteWindow(string year)
        {
            Title = $"🗑️  ລຶບປີ {year}";
            Width = 540; Height = 540;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));

            int students = DB.ScalarInt("SELECT COUNT(*) FROM Students    WHERE AcademicYear=@y", null, ("@y", year));
            int enrolls  = DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, ("@y", year));
            int scores   = DB.ScalarInt(@"SELECT COUNT(*) FROM Scores sc
                                          JOIN Enrollments e ON e.EnrollID=sc.EnrollID
                                          WHERE e.AcademicYear=@y", null, ("@y", year));
            int monthly  = DB.ScalarInt(@"SELECT COUNT(*) FROM MonthlyAssessments ma
                                          JOIN Enrollments e ON e.EnrollID=ma.EnrollID
                                          WHERE e.AcademicYear=@y", null, ("@y", year));
            int evals    = DB.ScalarInt("SELECT COUNT(*) FROM EvaluationScores WHERE AcademicYear=@y", null, ("@y", year));
            int attend   = DB.ScalarInt("SELECT COUNT(*) FROM AttendanceRecords WHERE AcademicYear=@y", null, ("@y", year));
            int history  = DB.ScalarInt("SELECT COUNT(*) FROM GradeHistory     WHERE AcademicYear=@y", null, ("@y", year));
            bool hasData = (students + enrolls + scores + monthly + evals + attend + history) > 0;

            var root = H.MkGrid(new GridLength(1, GridUnitType.Star), GridLength.Auto);
            var card = H.MkCard(new Thickness(16,16,16,8), new Thickness(20));
            var st = new StackPanel();

            // Header
            st.Children.Add(new TextBlock {
                Text = $"ກຳລັງຈະລຶບປີ {year}",
                FontSize = 17, FontWeight = FontWeights.Bold,
                Margin = new Thickness(0,0,0,14),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });

            // Summary grid
            var sumCard = new Border {
                Background = hasData
                    ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(254,242,242))
                    : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(240,253,244)),
                BorderBrush = hasData
                    ? new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(254,202,202))
                    : new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(187,247,208)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(14,10,14,10),
                Margin  = new Thickness(0,0,0,14)
            };
            var sumStack = new StackPanel();
            sumStack.Children.Add(new TextBlock {
                Text = hasData ? "ສະຫຼຸບຂໍ້ມູນທີ່ຈະຖືກລຶບ:" : "✅ ບໍ່ມີຂໍ້ມູນກ່ຽວຂ້ອງ — ປອດໄພທີ່ຈະລຶບ",
                FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0,0,0,6),
                Foreground = new System.Windows.Media.SolidColorBrush(
                    hasData ? System.Windows.Media.Color.FromRgb(153,27,27)
                            : System.Windows.Media.Color.FromRgb(22,101,52))
            });
            if (hasData)
            {
                sumStack.Children.Add(SumRow("ນັກຮຽນ",                students));
                sumStack.Children.Add(SumRow("ການລົງທະບຽນ",          enrolls));
                sumStack.Children.Add(SumRow("ຄະແນນປະຈຳພາກ",          scores));
                sumStack.Children.Add(SumRow("ຄະແນນປະຈຳເດືອນ",        monthly));
                sumStack.Children.Add(SumRow("ຄະແນນຄຸນສົມບັດ/ແຮງງານ", evals));
                sumStack.Children.Add(SumRow("ປະຫວັດການເຂົ້າຮຽນ",     attend));
                sumStack.Children.Add(SumRow("ປະຫວັດການຂຶ້ນຊັ້ນ",     history));
            }
            sumCard.Child = sumStack;
            st.Children.Add(sumCard);

            // Warning + typing requirement
            if (hasData)
            {
                st.Children.Add(new TextBlock {
                    Text = "⚠ ການກະທຳນີ້ ກູ້ຄືນບໍ່ໄດ້. ປະຫວັດທັງໝົດຈະຖືກລຶບແບບຖາວອນ.",
                    FontSize = 12, FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0,0,0,10), TextWrapping = TextWrapping.Wrap,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(153,27,27))
                });
            }
            st.Children.Add(new TextBlock {
                Text = "ພິມ ‘DELETE’ ເພື່ອຢືນຢັນ:",
                FontSize = 12,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81)),
                Margin = new Thickness(0,0,0,4)
            });
            var input = new TextBox { Margin = new Thickness(0,0,0,4), FontSize = 14 };
            st.Children.Add(input);

            card.Child = st;
            Grid.SetRow(card, 0); root.Children.Add(card);

            // Footer buttons
            var foot = new Border {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(0,1,0,0),
                Padding = new Thickness(14,10,14,10)
            };
            var fp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var bCancel = H.Btn("🛑  ຍົກເລີກ", "NeutralButton");
            var bDel = H.Btn(hasData ? "🔥  ບັງຄັບລຶບ" : "🗑️  ລຶບ", "DangerButton");
            bDel.IsEnabled = false;
            input.TextChanged += (s,e) => bDel.IsEnabled = input.Text.Trim() == "DELETE";
            bCancel.Click += (s,e) => { DialogResult = false; Close(); };
            bDel.Click    += (s,e) => { DialogResult = true;  Close(); };
            fp.Children.Add(bCancel); fp.Children.Add(bDel);
            foot.Child = fp;
            Grid.SetRow(foot, 1); root.Children.Add(foot);

            Content = root;
            Loaded += (_,_) => input.Focus();
        }

        private static UIElement SumRow(string label, int count)
        {
            var dp = new DockPanel { Margin = new Thickness(0,2,0,2) };
            var lbl = new TextBlock {
                Text = label, FontSize = 12,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99))
            };
            var val = new TextBlock {
                Text = count.ToString("N0"),
                FontSize = 13, FontWeight = FontWeights.SemiBold,
                HorizontalAlignment = HorizontalAlignment.Right,
                Foreground = new System.Windows.Media.SolidColorBrush(
                    count > 0 ? System.Windows.Media.Color.FromRgb(153,27,27)
                              : System.Windows.Media.Color.FromRgb(75,85,99))
            };
            DockPanel.SetDock(val, Dock.Right);
            dp.Children.Add(val); dp.Children.Add(lbl);
            return dp;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  ACADEMIC YEAR STATS WINDOW
    //
    //  Opened by the per-row 📊 ສະຖິຕິ button. Shows everything we can
    //  count for the picked year. Read-only.
    // ════════════════════════════════════════════════════════════
    public class AcademicYearStatsWindow : Window
    {
        public AcademicYearStatsWindow(string year)
        {
            Title = $"📊  ສະຖິຕິ — ປີ {year}";
            Width = 460; Height = 520;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));

            int students  = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y", null, ("@y", year));
            int active    = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ກຳລັງຮຽນ'", null, ("@y", year));
            int graduated = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ຈົບ'", null, ("@y", year));
            int withdrawn = DB.ScalarInt("SELECT COUNT(*) FROM Students WHERE AcademicYear=@y AND Status='ອອກ'", null, ("@y", year));
            int rooms     = DB.ScalarInt(@"SELECT COUNT(DISTINCT GradeLevel||'/'||IFNULL(ClassRoom,''))
                                           FROM Students WHERE AcademicYear=@y
                                             AND GradeLevel IS NOT NULL AND GradeLevel<>''",
                                           null, ("@y", year));
            int enrolls   = DB.ScalarInt("SELECT COUNT(*) FROM Enrollments WHERE AcademicYear=@y", null, ("@y", year));
            int scores    = DB.ScalarInt(@"SELECT COUNT(*) FROM Scores sc
                                           JOIN Enrollments e ON e.EnrollID=sc.EnrollID
                                           WHERE e.AcademicYear=@y", null, ("@y", year));
            int monthly   = DB.ScalarInt(@"SELECT COUNT(*) FROM MonthlyAssessments ma
                                           JOIN Enrollments e ON e.EnrollID=ma.EnrollID
                                           WHERE e.AcademicYear=@y", null, ("@y", year));
            int evals     = DB.ScalarInt("SELECT COUNT(*) FROM EvaluationScores WHERE AcademicYear=@y", null, ("@y", year));
            int attend    = DB.ScalarInt("SELECT COUNT(*) FROM AttendanceRecords WHERE AcademicYear=@y", null, ("@y", year));
            int history   = DB.ScalarInt("SELECT COUNT(*) FROM GradeHistory     WHERE AcademicYear=@y", null, ("@y", year));

            var root = H.MkGrid(new GridLength(1, GridUnitType.Star), GridLength.Auto);
            var card = H.MkCard(new Thickness(16,16,16,8), new Thickness(20));
            var st = new StackPanel();

            st.Children.Add(new TextBlock {
                Text = $"📊  ສະຖິຕິປີ {year}",
                FontSize = 17, FontWeight = FontWeights.Bold,
                Margin = new Thickness(0,0,0,14),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });

            st.Children.Add(Section("ນັກຮຽນ", new (string, int)[] {
                ("ທັງໝົດ",     students),
                ("ກຳລັງຮຽນ",  active),
                ("ຈົບ",        graduated),
                ("ອອກ",        withdrawn),
                ("ຫ້ອງຮຽນ",   rooms),
            }));
            st.Children.Add(Section("ຄະແນນ", new (string, int)[] {
                ("ການລົງທະບຽນ",       enrolls),
                ("ຄະແນນປະຈຳພາກ",       scores),
                ("ຄະແນນປະຈຳເດືອນ",    monthly),
                ("ຄະແນນຄຸນສົມບັດ/ແຮງງານ", evals),
                ("ປະຫວັດການເຂົ້າຮຽນ",  attend),
                ("ປະຫວັດການຂຶ້ນຊັ້ນ",   history),
            }));

            card.Child = st; Grid.SetRow(card, 0); root.Children.Add(card);

            var foot = new Border {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(0,1,0,0),
                Padding = new Thickness(14,10,14,10)
            };
            var bClose = H.Btn("ປິດ", "NeutralButton");
            bClose.HorizontalAlignment = HorizontalAlignment.Right;
            bClose.Click += (s,e) => Close();
            foot.Child = bClose;
            Grid.SetRow(foot, 1); root.Children.Add(foot);

            Content = root;
        }

        private static UIElement Section(string heading, (string Label, int Value)[] rows)
        {
            var stack = new StackPanel { Margin = new Thickness(0,0,0,14) };
            stack.Children.Add(new TextBlock {
                Text = heading, FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0,0,0,6),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            });
            foreach (var r in rows)
            {
                var dp = new DockPanel { Margin = new Thickness(0,1,0,1) };
                var lbl = new TextBlock {
                    Text = r.Label, FontSize = 12,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99))
                };
                var val = new TextBlock {
                    Text = r.Value.ToString("N0"),
                    FontSize = 13, FontWeight = FontWeights.SemiBold,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
                };
                DockPanel.SetDock(val, Dock.Right);
                dp.Children.Add(val); dp.Children.Add(lbl);
                stack.Children.Add(dp);
            }
            return stack;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  SAVE-CONFIRM DIALOG
    //
    //  3-button dirty-state prompt for both score-entry pages (ScoresPage
    //  + MonthlyScoresPage). Replaces the plain YesNo `ຈະຍົກເລີກບໍ?`
    //  MessageBox — the old buttons were English-locale ("Yes"/"No") and
    //  the wording (`ຈະຍົກເລີກບໍ?`) conflated "cancel edits" with "cancel
    //  the filter change".
    //
    //  Three explicit outcomes:
    //    💾 ບັນທຶກ      → save the dirty edits first, then let the caller
    //                     proceed with whatever prompted the dialog
    //    ❌ ບໍ່ບັນທຶກ  → discard edits, proceed
    //    ↩ ຍົກເລີກ     → keep edits, block the caller's action (stay put)
    //
    //  Usage:
    //      var d = SaveConfirmDialog.Ask(Window.GetWindow(this));
    //      switch (d) { case SaveConfirmResult.Save: …; case … }
    // ════════════════════════════════════════════════════════════
    public enum SaveConfirmResult { Save, DontSave, Cancel }

    public class SaveConfirmDialog : Window
    {
        public SaveConfirmResult Decision { get; private set; } = SaveConfirmResult.Cancel;

        public SaveConfirmDialog()
        {
            Title = "ຢືນຢັນ";
            Width = 440; Height = 210;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));
            UseLayoutRounding = true; SnapsToDevicePixels = true;

            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var body = new Border {
                Background = System.Windows.Media.Brushes.White,
                Padding = new Thickness(22, 20, 22, 20)
            };
            var flow = new StackPanel { Orientation = Orientation.Horizontal };
            flow.Children.Add(new TextBlock {
                Text = "❓", FontSize = 34,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 18, 0)
            });
            flow.Children.Add(new TextBlock {
                Text = "ມີຂໍ້ມູນທີ່ຍັງບໍ່ໄດ້ບັນທຶກ\nຕ້ອງການບັນທຶກກ່ອນປ່ຽນຕົວກອງບໍ?",
                FontSize = 14, TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(31,41,55))
            });
            body.Child = flow;
            Grid.SetRow(body, 0); root.Children.Add(body);

            var foot = new Border {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(14, 10, 14, 10)
            };
            var fp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var bSave     = H.Btn("💾  ບັນທຶກ",   "SuccessButton"); bSave.Width     = 110;
            var bDontSave = H.Btn("ບໍ່ບັນທຶກ",    "DangerButton");  bDontSave.Width = 110; bDontSave.Margin = new Thickness(8, 0, 0, 0);
            var bCancel   = H.Btn("ຍົກເລີກ",      "NeutralButton"); bCancel.Width   = 90;  bCancel.Margin   = new Thickness(8, 0, 0, 0); bCancel.IsCancel = true;
            bSave.Click     += (s, e) => { Decision = SaveConfirmResult.Save;     DialogResult = true;  Close(); };
            bDontSave.Click += (s, e) => { Decision = SaveConfirmResult.DontSave; DialogResult = true;  Close(); };
            bCancel.Click   += (s, e) => { Decision = SaveConfirmResult.Cancel;   DialogResult = false; Close(); };
            fp.Children.Add(bSave); fp.Children.Add(bDontSave); fp.Children.Add(bCancel);
            foot.Child = fp;
            Grid.SetRow(foot, 1); root.Children.Add(foot);

            Content = root;
        }

        // Convenience: build + show + return the decision. Owner is optional
        // (falls back to Application.Current.MainWindow) so the dialog centres
        // properly and stays modal to the parent.
        public static SaveConfirmResult Ask(Window? owner)
        {
            var dlg = new SaveConfirmDialog();
            if (owner != null) dlg.Owner = owner;
            else if (Application.Current.MainWindow != null && Application.Current.MainWindow.IsLoaded)
                dlg.Owner = Application.Current.MainWindow;
            dlg.ShowDialog();
            return dlg.Decision;
        }
    }

    public class AcademicYearFormWin : Window
    {
        private TextBox _year = null!, _start = null!, _end = null!, _note = null!;

        // `prefillYear` lets callers (e.g. the Promotion page's ຂຶ້ນຊັ້ນ / ຊ້ຳຊັ້ນ
        // pre-check) open the dialog with the exact year they need created —
        // defaults to NextYearString(CurrentYear) which matches the plain
        // "➕ ເພີ່ມປີໃໝ່" button on AcademicYearPage.
        public AcademicYearFormWin(string? prefillYear = null)
        {
            Title = "➕  ເພີ່ມປີການສຶກສາ";
            Width = 480; Height = 440;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));
            UseLayoutRounding = true; SnapsToDevicePixels = true;

            var root = H.MkGrid(new GridLength(1, GridUnitType.Star), GridLength.Auto);
            var card = H.MkCard(new Thickness(16,16,16,8), new Thickness(20));
            var st = new StackPanel();

            string yearValue = prefillYear ?? DB.NextYearString(DB.CurrentYear);
            st.Children.Add(FL($"ປີການສຶກສາ *  (ຮູບແບບ YYYY-YYYY — ຕົວຢ່າງ {DB.NextYearString(DB.CurrentYear)})"));
            _year = FI(st, yearValue);

            st.Children.Add(FL("ວັນເລີ່ມ  (optional, ຮູບແບບ YYYY-MM-DD)"));
            _start = FI(st, "");

            st.Children.Add(FL("ວັນສິ້ນສຸດ  (optional)"));
            _end = FI(st, "");

            st.Children.Add(FL("ໝາຍເຫດ"));
            _note = FI(st, "");

            card.Child = st;
            Grid.SetRow(card, 0); root.Children.Add(card);

            var foot = new Border {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(0,1,0,0),
                Padding = new Thickness(14,10,14,10)
            };
            var fp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var bOK = H.Btn("💾  ບັນທຶກ", "SuccessButton"); bOK.Width = 110; bOK.Click += Save;
            var bCancel = new Button { Content = "ຍົກເລີກ", Width = 90, Height = 34, Margin = new Thickness(8,0,0,0), IsCancel = true };
            bCancel.SetResourceReference(Button.StyleProperty, "NeutralButton");
            fp.Children.Add(bOK); fp.Children.Add(bCancel);
            foot.Child = fp;
            Grid.SetRow(foot, 1); root.Children.Add(foot);

            Content = root;
        }

        private void Save(object s, RoutedEventArgs e)
        {
            string year = _year.Text.Trim();
            if (!DB.IsValidYearFormat(year))
            {
                MessageBox.Show("ກະລຸນາໃສ່ປີໃນຮູບແບບ YYYY-YYYY ໂດຍປີທີສອງຕ້ອງເທົ່າກັບປີທຳອິດ + 1\n(ຕົວຢ່າງ: 2026-2027)",
                                "ຮູບແບບປີຜິດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string? start = string.IsNullOrWhiteSpace(_start.Text) ? null : _start.Text.Trim();
            string? end   = string.IsNullOrWhiteSpace(_end.Text)   ? null : _end.Text.Trim();
            string? note  = string.IsNullOrWhiteSpace(_note.Text)  ? null : _note.Text.Trim();
            try
            {
                DB.CreateAcademicYear(year, start, end, note);
                DialogResult = true; Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ: {ex.Message}", "ຜິດພາດ",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static TextBlock FL(string t) => new TextBlock {
            Text = t, FontSize = 12, FontWeight = FontWeights.Medium,
            Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100,116,139)),
            Margin = new Thickness(0,0,0,6)
        };
        private static TextBox FI(StackPanel p, string v = "")
        {
            var tb = new TextBox { Text = v, Margin = new Thickness(0,0,0,12) };
            p.Children.Add(tb); return tb;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  SETTINGS PAGE
    // ════════════════════════════════════════════════════════════
    public class SettingsPage : UserControl
    {
        public SettingsPage()
        {
            var tabs=new TabControl{BorderThickness=new Thickness(0)};
            tabs.Items.Add(MkTab("⚙️  ທົ່ວໄປ",    BuildGeneral()));
            tabs.Items.Add(MkTab("📥  ນຳເຂົ້າ CSV",BuildImport()));
            tabs.Items.Add(MkTab("💾  Backup",      BuildBackup()));
            tabs.Items.Add(MkTab("📋  Log",         BuildLog()));
            Content=tabs;
        }

        private TextBox _school=null!;
        private DataGrid _logDg=null!;

        private UIElement BuildGeneral()
        {
            var st = new StackPanel { Margin = new Thickness(20) };
            _school = F(st, "ຊື່ໂຮງຮຽນ", DB.SchoolName);

            // Academic year + semester are managed on the dedicated ‘ປີການສຶກສາ’
            // admin page (which keeps Settings + the AcademicYears registry in sync),
            // so they are intentionally NOT editable here.
            // Pass-score threshold is kept at its database default and is not editable
            // from the UI — it's a school-wide grading rule that shouldn't change
            // mid-year, and exposing it caused confusion.
            st.Children.Add(new TextBlock {
                Text = "ℹ ປີການສຶກສາ ແລະ ພາກຮຽນ ຈັດການຢູ່ໜ້າ ‘ປີການສຶກສາ’",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),
                Margin = new Thickness(0,0,0,12)
            });

            var btn = B("💾  ບັນທຶກ", "SuccessButton", 180);
            btn.Click += (s,e) => SaveGeneral();
            st.Children.Add(btn);
            return new ScrollViewer { Content = st };
        }

        private void SaveGeneral()
        {
            string school = _school.Text.Trim();
            if (school.Length == 0)
            {
                MessageBox.Show("ກະລຸນາໃສ່ຊື່ໂຮງຮຽນ", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DB.SaveSetting("school_name", school);
            DB.Log("Settings", "ທົ່ວໄປ");
            MessageBox.Show("ບັນທຶກສຳເລັດ", "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private UIElement BuildImport()
        {
            // Full column list for the new schema (30 fields).
            // 28 columns — NationalID and student Phone removed.
            const string CsvHeader =
                "StudentCode,FirstName,LastName,Gender,BirthDate," +
                "BirthVillage,BirthDistrict,BirthProvince," +
                "Village,District,Province," +
                "FatherName,FatherAge,FatherJob,FatherVillage,FatherDistrict,FatherProvince,FatherPhone," +
                "MotherName,MotherAge,MotherJob,MotherVillage,MotherDistrict,MotherProvince,MotherPhone," +
                "GradeLevel,ClassRoom,AcademicYear";
            const string CsvSample =
                "S001,ສົມຊາຍ,ໃຈດີ,ຊາຍ,01/01/2009," +
                "ນາໂພ,ຫາດຊາຍຟອງ,ວຽງຈັນ," +
                "ດົງໂດກ,ໄຊທານີ,ວຽງຈັນ," +
                "ທ້າວດີ ໃຈດີ,45,ກະສິກຳ,ດົງໂດກ,ໄຊທານີ,ວຽງຈັນ,0201234567," +
                "ນາງດີ ໃຈດີ,42,ຄ້າຂາຍ,ດົງໂດກ,ໄຊທານີ,ວຽງຈັນ,0207654322," +
                "ມ.4,1,2025-2026";

            var st = new StackPanel { Margin = new Thickness(20) };
            st.Children.Add(new TextBlock {
                Text = "ນຳເຂົ້ານັກຮຽນຈາກໄຟລ໌ CSV\n\n" +
                       "ຄໍລຳ (28 ຂໍ້): " + CsvHeader + "\n\n" +
                       "• ຖ້າລະຫັດຊ້ຳ ຈະຂ້າມ\n" +
                       "• ໄຟລ໌ UTF-8\n" +
                       "• ຄໍລຳທີ່ຂາດ ຈະຖືກບັນທຶກເປັນຄ່າຫວ່າງ — ສົ່ງສະເພາະຄໍລຳທີ່ຕ້ອງການ",
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0,0,0,14)
            });
            var bTpl = B("📄  Template", "NeutralButton", 180);
            bTpl.Click += (s,e) => {
                var dlg = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "students_template.csv" };
                if (dlg.ShowDialog() != true) return;
                File.WriteAllText(dlg.FileName, CsvHeader + "\n" + CsvSample + "\n", Encoding.UTF8);
                MessageBox.Show("ບັນທຶກ Template ສຳເລັດ", "ສຳເລັດ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            };
            st.Children.Add(bTpl);
            st.Children.Add(new TextBlock { Height = 10 });

            var bImp = B("📂  ເລືອກໄຟລ໌ ແລະ ນຳເຂົ້າ", "PrimaryButton", 220);
            bImp.Click += (s,e) => {
                var dlg = new OpenFileDialog { Filter = "CSV|*.csv" };
                if (dlg.ShowDialog() != true) return;
                var lines = File.ReadAllLines(dlg.FileName, Encoding.UTF8);
                if (lines.Length < 2) { MessageBox.Show("ໄຟລ໌ບໍ່ມີຂໍ້ມູນ", "ບໍ່ມີຂໍ້ມູນ", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                var hdrs = lines[0].Split(',');
                int ins = 0, skip = 0, invalid = 0;
                using var conn = DB.Open();
                for (int i = 1; i < lines.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i])) continue;
                    var vals = lines[i].Split(',');
                    string G(string h)
                    {
                        int ix = Array.IndexOf(hdrs, h);
                        return ix >= 0 && ix < vals.Length ? vals[ix].Trim().Trim('"') : "";
                    }
                    object Age(string h)
                    {
                        var s2 = G(h);
                        return int.TryParse(s2, out int n) ? (object)n : DBNull.Value;
                    }
                    // Required, NOT NULL columns — a row missing any of these would
                    // either fail the INSERT or create an unusable record that never
                    // shows in the grade/year-filtered views. Skip + count separately.
                    if (string.IsNullOrWhiteSpace(G("StudentCode"))
                        || string.IsNullOrWhiteSpace(G("FirstName"))
                        || string.IsNullOrWhiteSpace(G("LastName"))
                        || string.IsNullOrWhiteSpace(G("GradeLevel"))
                        || string.IsNullOrWhiteSpace(G("AcademicYear")))
                    { invalid++; continue; }
                    try
                    {
                        int n = DB.Exec(@"INSERT OR IGNORE INTO Students(
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
                            ("@code", G("StudentCode")), ("@fn", G("FirstName")), ("@ln", G("LastName")),
                            ("@gd",   G("Gender")),      ("@bd", G("BirthDate")),
                            ("@bvi",  G("BirthVillage")),("@bdi",G("BirthDistrict")),("@bpv",G("BirthProvince")),
                            ("@vi",   G("Village")),     ("@di", G("District")),    ("@pv", G("Province")),
                            ("@faN",  G("FatherName")),  ("@faA",Age("FatherAge")), ("@faJ",G("FatherJob")),
                            ("@faV",  G("FatherVillage")),("@faC",G("FatherDistrict")),("@faP",G("FatherProvince")),
                            ("@faT",  G("FatherPhone")),
                            ("@maN",  G("MotherName")),  ("@maA",Age("MotherAge")), ("@maJ",G("MotherJob")),
                            ("@maV",  G("MotherVillage")),("@maC",G("MotherDistrict")),("@maP",G("MotherProvince")),
                            ("@maT",  G("MotherPhone")),
                            // Legacy mirror so older queries still see *some* parent
                            ("@pn",   string.IsNullOrWhiteSpace(G("FatherName")) ? G("MotherName") : G("FatherName")),
                            ("@pp",   string.IsNullOrWhiteSpace(G("FatherPhone")) ? G("MotherPhone") : G("FatherPhone")),
                            ("@gl",   G("GradeLevel")),  ("@rm", G("ClassRoom")),   ("@yr", G("AcademicYear")));
                        if (n > 0) ins++; else skip++;
                    }
                    catch { skip++; }
                }
                DB.Log("ImportCSV", $"{ins} ຄົນ ຂ້າມ {skip} ບໍ່ຄົບ {invalid}");
                MessageBox.Show(
                    $"ນຳເຂົ້າ {ins} ຄົນ\n" +
                    $"ຂ້າມ (ລະຫັດຊ້ຳ) {skip} ຄົນ\n" +
                    $"ບໍ່ຄົບຂໍ້ມູນຈຳເປັນ {invalid} ແຖວ",
                    "ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Information);
            };
            st.Children.Add(bImp);
            return new ScrollViewer { Content = st };
        }

        private UIElement BuildBackup(){var st=new StackPanel{Margin=new Thickness(20)};st.Children.Add(new TextBlock{Text=$"ໄຟລ໌: sis_lao.db\nທີ່ຢູ່: {AppDomain.CurrentDomain.BaseDirectory}",Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,16)});var bB=B("📦  Backup","PrimaryButton",200);bB.Click+=(s,e)=>{var dlg=new SaveFileDialog{Filter="DB|*.db",FileName=$"backup_{DateTime.Now:yyyyMMdd_HHmm}.db"};if(dlg.ShowDialog()!=true)return;DB.Backup(dlg.FileName);MessageBox.Show("Backup ສຳເລັດ","ສຳເລັດ",MessageBoxButton.OK,MessageBoxImage.Information);};st.Children.Add(bB);st.Children.Add(new TextBlock{Height=10});var bR=B("🔄  Restore","WarningButton",200);bR.Click+=(s,e)=>{if(MessageBox.Show("⚠ Restore ຈະແທນທີ່ຂໍ້ມູນທັງໝົດ?","ຄຳເຕືອນ",MessageBoxButton.YesNo,MessageBoxImage.Warning)!=MessageBoxResult.Yes)return;var dlg=new OpenFileDialog{Filter="DB|*.db"};if(dlg.ShowDialog()!=true)return;DB.Restore(dlg.FileName);DB.LoadSettings();MessageBox.Show("Restore ສຳເລັດ — ກະລຸນາເປີດໂປຣແກຣ ໃໝ່","ສຳເລັດ",MessageBoxButton.OK,MessageBoxImage.Information);Application.Current.Shutdown();System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(System.Diagnostics.Process.GetCurrentProcess().MainModule!.FileName!){UseShellExecute=true});};st.Children.Add(bR);return new ScrollViewer{Content=st};}

        private UIElement BuildLog(){var root=H.MkGrid(GridLength.Auto,new GridLength(1,GridUnitType.Star));var flow=new WrapPanel{Margin=new Thickness(0,0,0,10)};var bR=H.Btn("🔄  ໂຫຼດ","PrimaryButton"); bR.Click+=(s,e)=>ReloadLog();var bC=H.Btn("🗑  ລ້າງ Log (>30 ວັນ)","DangerButton"); bC.Click+=(s,e)=>{if(MessageBox.Show("ລຶບ Log ທີ່ເກົ່າກວ່າ 30 ວັນ? (ກູ້ຄືນບໍ່ໄດ້)","ຢືນຢັນ",MessageBoxButton.YesNo,MessageBoxImage.Warning)!=MessageBoxResult.Yes)return;int del=DB.Exec("DELETE FROM ActivityLog WHERE LoggedAt<datetime('now','-30 days')");ReloadLog();MessageBox.Show($"ລ້າງ Log ສຳເລັດ — ລຶບ {del} ລາຍການ","ສຳເລັດ",MessageBoxButton.OK,MessageBoxImage.Information);};flow.Children.Add(bR);flow.Children.Add(bC);Grid.SetRow(flow,0);root.Children.Add(flow);_logDg=new DataGrid{AutoGenerateColumns=true,IsReadOnly=true,CanUserAddRows=false,BorderThickness=new Thickness(0),Background=System.Windows.Media.Brushes.White};Grid.SetRow(_logDg,1);root.Children.Add(_logDg);ReloadLog();return root;}
        private void ReloadLog(){var dt=DB.Query("SELECT LoggedAt AS ເວລາ,Username AS ຜູ້ໃຊ້,Action AS ການກະທຳ,Detail AS ລາຍລະອຽດ FROM ActivityLog ORDER BY LogID DESC LIMIT 500");_logDg.ItemsSource=dt.DefaultView;}

        private TextBox F(StackPanel p,string lbl,string val=""){L(p,lbl);var tb=new TextBox{Text=val,Margin=new Thickness(0,0,0,12)};p.Children.Add(tb);return tb;}
        private void L(StackPanel p,string t)=>p.Children.Add(new TextBlock{Text=t,FontSize=12,Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),Margin=new Thickness(0,0,0,4)});
        private Button B(string t,string st,double w){var b=new Button{Content=t,Height=34,Width=w,Padding=new Thickness(14,0,14,0),Margin=new Thickness(0,0,0,8),HorizontalAlignment=HorizontalAlignment.Left,Cursor=System.Windows.Input.Cursors.Hand};b.SetResourceReference(Button.StyleProperty,st);return b;}
        private static TabItem MkTab(string h,UIElement c)=>new TabItem{Header=h,Content=c};
    }

    // ════════════════════════════════════════════════════════════
    //  STUDENT SCORE HISTORY  (class-roster picker)
    //
    //  Top: filter bar — ປີ / ຊັ້ນ / ຫ້ອງ + 🔍 search box.
    //  Below: status-filter radio bar — 🔘 ກຳລັງຮຽນ (default) /
    //  🎓 ຈົບ / 📋 ທັງໝົດ — mutex via GroupName, drives the cohort's
    //  Students.Status filter. Alongside: ONE global 📚 ປະຫວັດທັງຫ້ອງ
    //  button (disabled until all three filters set) opens ClassHistoryWindow
    //  for the current (year, grade, room) — there's exactly one class report
    //  for any given selection, so it's deduplicated to the page level.
    //  Then: DataGrid roster of every student who was HISTORICALLY
    //  in that (year, grade, room) — sourced from
    //  DB.GetHistoricalClassRoster (which combines GradeHistory rows
    //  with Students rows for the latest year). NEVER reads only
    //  Students.GradeLevel/ClassRoom — those are current-only.
    //
    //  Per row, ONE action button:
    //      📖 ປະຫວັດສ່ວນຕົວ  → opens StudentHistoryWindow(sid)
    //
    //  No data loads at startup beyond the year-catalogue. The roster
    //  query fires only when Year/Grade/Room all have a value. Both
    //  history windows fire their own queries when their button is
    //  clicked — nothing is preloaded.
    // ════════════════════════════════════════════════════════════
    public class ScoreHistoryPage : UserControl
    {
        private ComboBox _cmbYear=null!, _cmbGrade=null!, _cmbRoom=null!;
        private TextBox  _txtSearch=null!;
        private DataGrid _dg=null!;
        private TextBlock _info=null!;
        private Button   _btnClassHistory=null!;
        private RadioButton _rdoActive=null!, _rdoGraduated=null!, _rdoAll=null!;
        private DataTable _roster=new();   // cached for client-side search

        public ScoreHistoryPage()
        {
            // Three-row grid:  filter card  ·  class-history button bar  ·  roster grid
            var root = H.MkGrid(GridLength.Auto, GridLength.Auto, new GridLength(1, GridUnitType.Star));

            // ── Filter bar ───────────────────────────────────────
            var tb   = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,10,14,10));
            var flow = new WrapPanel();
            _cmbYear  = H.MkCmb(DB.AcademicYears().ToArray(), 110);
            SelectComboValueSH(_cmbYear, DB.CurrentYear);
            _cmbGrade = H.MkCmb(new[]{ "ມ.1", "ມ.2", "ມ.3", "ມ.4" }, 80);
            _cmbRoom  = H.MkCmb(new[]{ "1", "2", "3", "4", "5", "6" }, 80);
            _txtSearch = new TextBox { Width = 200, Margin = new Thickness(0,0,8,0),
                VerticalContentAlignment = VerticalAlignment.Center, Height = 28 };
            _txtSearch.TextChanged += (s,e) => ApplySearch();

            _cmbYear.SelectionChanged  += (s,e) => ReloadRoster();
            _cmbGrade.SelectionChanged += (s,e) => ReloadRoster();
            _cmbRoom.SelectionChanged  += (s,e) => ReloadRoster();

            flow.Children.Add(H.Lbl("ປີ:"));    flow.Children.Add(_cmbYear);
            flow.Children.Add(H.Lbl("ຊັ້ນ:"));  flow.Children.Add(_cmbGrade);
            flow.Children.Add(H.Lbl("ຫ້ອງ:"));  flow.Children.Add(_cmbRoom);
            flow.Children.Add(H.Lbl("🔍"));      flow.Children.Add(_txtSearch);
            tb.Child = flow; Grid.SetRow(tb, 0); root.Children.Add(tb);

            // ── Status filter (mutex radios) + global class-history button ──
            // Three mutually-exclusive radios drive Students.Status filtering:
            //   🔘 ກຳລັງຮຽນ  (default — only active students)
            //   🎓 ຈົບ        (graduated only — Status='ຈົບ')
            //   📋 ທັງໝົດ    (all statuses — ກຳລັງຮຽນ + ຈົບ + ອອກ)
            // Changing a radio re-fires ReloadRoster which re-queries
            // DB.GetHistoricalClassRoster with the matching statusFilter param.
            // Year/Grade/Room dropdowns are NOT touched by status switching.
            // Per-window report-type dropdown (StudentHistoryWindow /
            // ClassHistoryWindow) is independent of this page — never reset.
            var btnBar = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,8,14,8));
            _btnClassHistory = H.Btn("📚  ປະຫວັດທັງຫ້ອງ", "NeutralButton");
            _btnClassHistory.Click += (s, e) => OpenClassHistory();
            _rdoActive    = MkRadio("🔘 ກຳລັງຮຽນ", isChecked: true);
            _rdoGraduated = MkRadio("🎓 ຈົບ",        isChecked: false);
            _rdoAll       = MkRadio("📋 ທັງໝົດ",    isChecked: false);
            _rdoActive.Checked    += (s, e) => ReloadRoster();
            _rdoGraduated.Checked += (s, e) => ReloadRoster();
            _rdoAll.Checked       += (s, e) => ReloadRoster();
            var btnFlow = new WrapPanel();
            btnFlow.Children.Add(_rdoActive);
            btnFlow.Children.Add(_rdoGraduated);
            btnFlow.Children.Add(_rdoAll);
            // Visual separator between the radios and the class-history button.
            btnFlow.Children.Add(new Border {
                Width = 1, Height = 22, Margin = new Thickness(8, 0, 12, 0),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(209,213,219))
            });
            btnFlow.Children.Add(_btnClassHistory);
            btnBar.Child = btnFlow;
            Grid.SetRow(btnBar, 1); root.Children.Add(btnBar);

            // ── Roster DataGrid with ONE action button per row ───
            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White,
                SelectionMode = DataGridSelectionMode.Single
            };
            _dg.Columns.Add(H.Col("ລະຫັດ",      "ລະຫັດ", 110, true));
            _dg.Columns.Add(H.ColStar("ຊື່ນັກຮຽນ", "ຊື່ນັກຮຽນ", true));
            _dg.Columns.Add(H.Col("ເພດ",         "ເພດ", 70, true));
            _dg.Columns.Add(H.Col("ສະຖານະ",      "ສະຖານະ", 100, true));
            _dg.Columns.Add(MkActionCol("📖 ປະຫວັດສ່ວນຕົວ"));
            _dg.AddHandler(Button.ClickEvent, new RoutedEventHandler(OnRowAction));

            _info = new TextBlock {
                FontSize = 12, Margin = new Thickness(4, 8, 0, 0),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))
            };
            var dock = new DockPanel();
            DockPanel.SetDock(_info, Dock.Bottom);
            dock.Children.Add(_info);
            dock.Children.Add(_dg);
            var card = H.MkCard(); card.Child = dock;
            Grid.SetRow(card, 2); root.Children.Add(card);
            Content = root;

            // Nav-context handoff from ClassHubPage: pre-select year/grade/room
            // so the class-history roster loads on arrival.
            if (!string.IsNullOrEmpty(DB.NavGrade))
            {
                SelectComboValueSH(_cmbGrade, DB.NavGrade);
                if (!string.IsNullOrEmpty(DB.NavRoom)) SelectComboValueSH(_cmbRoom, DB.NavRoom);
                if (!string.IsNullOrEmpty(DB.NavYear)) SelectComboValueSH(_cmbYear, DB.NavYear);
                DB.ClearNav();
            }

            ReloadRoster();
        }

        private void OpenClassHistory()
        {
            string year  = _cmbYear.SelectedItem?.ToString()  ?? "";
            string grade = _cmbGrade.SelectedItem?.ToString() ?? "";
            string room  = _cmbRoom.SelectedItem?.ToString()  ?? "";
            if (year == "" || grade == "" || room == "")
            {
                MessageBox.Show("ກະລຸນາເລືອກປີ + ຊັ້ນ + ຫ້ອງ ກ່ອນ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            new ClassHistoryWindow(year, grade, room) {
                Owner = Window.GetWindow(this)
            }.Show();
        }

        private static void SelectComboValueSH(ComboBox cb, string val)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == val) { cb.SelectedIndex = i; return; }
        }

        // Single shared GroupName makes the three radios mutually exclusive.
        private static RadioButton MkRadio(string label, bool isChecked) =>
            new RadioButton {
                Content = label,
                IsChecked = isChecked,
                GroupName = "ScoreHistoryStatus",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 14, 0),
                FontSize = 13
            };

        // Map the currently-checked radio to a Students.Status filter:
        //   active radio    → "ກຳລັງຮຽນ"
        //   graduated radio → "ຈົບ"
        //   all radio       → null  (no SQL filter — every status passes)
        private string? CurrentStatusFilter()
        {
            if (_rdoGraduated?.IsChecked == true) return "ຈົບ";
            if (_rdoAll?.IsChecked == true)       return null;
            return "ກຳລັງຮຽນ";   // default = Active
        }

        // Per-row action button column — only one button now (📖 ປະຫວັດສ່ວນຕົວ).
        // The bubbled Click handler reads the row via the button's DataContext.
        private static DataGridTemplateColumn MkActionCol(string label)
        {
            var col = new DataGridTemplateColumn { Header = "", Width = new DataGridLength(170) };
            var btnFactory = new System.Windows.FrameworkElementFactory(typeof(Button));
            btnFactory.SetValue(Button.ContentProperty, label);
            btnFactory.SetValue(Button.HeightProperty, 26.0);
            btnFactory.SetValue(Button.MarginProperty, new Thickness(2));
            btnFactory.SetValue(Button.CursorProperty, System.Windows.Input.Cursors.Hand);
            btnFactory.SetResourceReference(Button.StyleProperty, "PrimaryButton");
            col.CellTemplate = new System.Windows.DataTemplate { VisualTree = btnFactory };
            return col;
        }

        private void OnRowAction(object sender, RoutedEventArgs e)
        {
            if (e.OriginalSource is not Button btn) return;
            if (btn.DataContext is not DataRowView drv) return;
            int sid = Convert.ToInt32(drv["StudentID"]);
            string label = $"{drv["ລະຫັດ"]}  ·  {drv["ຊື່ນັກຮຽນ"]}";
            new StudentHistoryWindow(sid, label) {
                Owner = Window.GetWindow(this)
            }.Show();
        }

        private void ReloadRoster()
        {
            string year  = _cmbYear.SelectedItem?.ToString()  ?? "";
            string grade = _cmbGrade.SelectedItem?.ToString() ?? "";
            string room  = _cmbRoom.SelectedItem?.ToString()  ?? "";
            bool allPicked = year != "" && grade != "" && room != "";
            // Enable the global class-history button only when all three filters are set.
            if (_btnClassHistory != null) _btnClassHistory.IsEnabled = allPicked;
            if (!allPicked)
            {
                _roster = new DataTable();
                _dg.ItemsSource = null;
                _info.Text = "ກະລຸນາເລືອກປີ + ຊັ້ນ + ຫ້ອງ";
                return;
            }
            string? status = CurrentStatusFilter();
            _roster = DB.GetHistoricalClassRoster(year, grade, room, status);
            ApplySearch();
            DB.Log("ScoreHistoryRoster",
                $"y={year} g={grade} r={room} status={status ?? "ALL"} n={_roster.Rows.Count}");
        }

        private void ApplySearch()
        {
            string q = (_txtSearch?.Text ?? "").Trim();
            if (string.IsNullOrEmpty(q))
            {
                _dg.ItemsSource = _roster.DefaultView;
                _info.Text = $"👥 ນັກຮຽນທັງໝົດ: {_roster.Rows.Count} ຄົນ";
                return;
            }
            // Client-side filter on code OR name. DataView.RowFilter uses LIKE syntax;
            // escape single-quotes the user might type.
            string esc = q.Replace("'", "''");
            var v = new DataView(_roster) {
                RowFilter = $"ລະຫັດ LIKE '%{esc}%' OR ຊື່ນັກຮຽນ LIKE '%{esc}%'"
            };
            _dg.ItemsSource = v;
            _info.Text = $"🔍 ພົບ {v.Count} ຄົນ ຈາກ {_roster.Rows.Count}";
        }
    }

    // ════════════════════════════════════════════════════════════
    //  STUDENT HISTORY WINDOW
    //
    //  Per-student multi-year viewer opened by the 📖 button. For each
    //  academic year the student has data for, renders:
    //      🗓 ເດືອນ 9 / 10 / 11 / 12  (per-subject monthly grid)
    //      📘 ສະຫຼຸບພາກຮຽນ 1
    //      🗓 ເດືອນ 2 / 3 / 4 / 5
    //      📗 ສະຫຼຸບພາກຮຽນ 2
    //      📒 ສະຫຼຸບປະຈຳປີ
    //  Excel + PDF export at the bottom — re-uses ReportPage's COM helper
    //  for PDF conversion.
    // ════════════════════════════════════════════════════════════
    // ── Shared report-type catalog used by both history windows ──
    // The 11 options the user picks from a single dropdown. Order matches the
    // academic year flow: 4 sem-1 months → sem-1 summary → 4 sem-2 months →
    // sem-2 summary → annual.
    // Kind:
    //   "M{n}"  → monthly report for calendar month n (n ∈ {2..5, 9..12})
    //   "S{n}"  → semester n summary  (n ∈ {1, 2})
    //   "A"     → annual summary
    internal class HistoryReportItem
    {
        public string Kind { get; }
        public string Label { get; }
        public HistoryReportItem(string kind, string label) { Kind = kind; Label = label; }
        public override string ToString() => Label;
        public bool IsMonth   => Kind.StartsWith("M");
        public bool IsSemester => Kind.StartsWith("S");
        public bool IsAnnual  => Kind == "A";
        public int Month => int.Parse(Kind.Substring(1));    // valid only when IsMonth
        public int Sem   => int.Parse(Kind.Substring(1));    // valid only when IsSemester
    }

    internal static class HistoryReportCatalog
    {
        // The exact ordering the spec calls for. Ordinal labels ("ເດືອນທີ 1..8")
        // follow the school's academic year rhythm — month 1 = Sept (start of
        // sem 1), month 8 = May (end of sem 2). Final-exam months (Jan / Jun)
        // are NOT user-selectable here — their values feed into the Sem 1 / Sem 2
        // summary options that bracket them.
        public static HistoryReportItem[] Items = new[]
        {
            new HistoryReportItem("M9",  "ເດືອນທີ ກັນຍາ"),
            new HistoryReportItem("M10", "ເດືອນທີ ຕຸລາ"),
            new HistoryReportItem("M11", "ເດືອນທີ ພະຈິກ"),
            new HistoryReportItem("M12", "ເດືອນທີ ທັນວາ"),
            new HistoryReportItem("S1",  "📘 ສະຫຼຸບພາກຮຽນ 1"),
            new HistoryReportItem("M2",  "ເດືອນທີ ກຸມພາ"),
            new HistoryReportItem("M3",  "ເດືອນທີ ມີນາ"),
            new HistoryReportItem("M4",  "ເດືອນທີ ເມສາ"),
            new HistoryReportItem("M5",  "ເດືອນທີ ພຶດສະພາ"),
            new HistoryReportItem("S2",  "📗 ສະຫຼຸບພາກຮຽນ 2"),
            new HistoryReportItem("A",   "📕 ສະຫຼຸບປະຈຳປີ"),
        };

        // Settings key used to remember the user's last pick across windows + sessions.
        public const string LastKey = "score_history_report_type";
        public const string Default = "M9";

        public static int FindIndex(string kind)
        {
            for (int i = 0; i < Items.Length; i++)
                if (Items[i].Kind == kind) return i;
            return 0;
        }
    }

    public class StudentHistoryWindow : Window
    {
        private readonly int _sid;
        private readonly string _label;
        private ComboBox _cmbExportYear=null!, _cmbReport=null!;
        private TextBlock _statusTxt=null!;
        // (year → (grade, room)) snapshot so export picks the historical grade/room
        // for the chosen year without re-querying GradeHistory each time.
        private readonly Dictionary<string,(string grade,string room)> _yrToGradeRoom = new();
        // Cached years list so the body can be rebuilt cheaply when the picked
        // report kind changes — no need to re-query GetStudentHistoryYears.
        private DataTable _years = null!;
        // Body panel exposed as a field so the report-picker can clear + rebuild it.
        private StackPanel _body = null!;

        public StudentHistoryWindow(int sid, string label)
        {
            _sid = sid; _label = label;
            Title = $"ປະຫວັດຄະແນນ — {label}";
            Width = 1100; Height = 760;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252));

            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // Header card
            var hdr = H.MkCard(new Thickness(12,12,12,8), new Thickness(14,10,14,10));
            hdr.Child = new TextBlock {
                Text = $"📖  ປະຫວັດຄະແນນສ່ວນຕົວ: {label}",
                FontSize = 15, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            };
            Grid.SetRow(hdr, 0); root.Children.Add(hdr);

            // Scrollable body — per-year sections. Kept in a field so the
            // report-picker in the action bar can rebuild it in-place.
            _body = new StackPanel { Margin = new Thickness(12,4,12,4) };
            _years = DB.GetStudentHistoryYears(sid);
            foreach (DataRow yr in _years.Rows)
            {
                string y = yr["AcademicYear"].ToString() ?? "";
                string g = yr["GradeLevel"].ToString()  ?? "—";
                string r = yr["ClassRoom"].ToString()   ?? "—";
                _yrToGradeRoom[y] = (g, r);
            }
            var sv = new ScrollViewer {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = _body
            };
            Grid.SetRow(sv, 1); root.Children.Add(sv);

            // Action bar — Year + single Report-Type dropdown (11 options) +
            // Excel + PDF + Close. Status text in front of the dropdown updates
            // on selection change. Last-picked report type persists across
            // windows + sessions via DB.Settings(score_history_report_type).
            var actions = H.MkCard(new Thickness(12,4,12,12), new Thickness(12,8,12,8));
            var bar = new WrapPanel();
            _cmbExportYear  = new ComboBox { Width = 110, Margin = new Thickness(0,0,8,0) };
            foreach (var y in _yrToGradeRoom.Keys) _cmbExportYear.Items.Add(y);
            if (_cmbExportYear.Items.Count > 0) _cmbExportYear.SelectedIndex = 0;

            _cmbReport = new ComboBox { Width = 220, Margin = new Thickness(0,0,8,0) };
            foreach (var it in HistoryReportCatalog.Items) _cmbReport.Items.Add(it);
            string lastKind = (DB.Scalar("SELECT Value FROM Settings WHERE Key=@k",
                null, ("@k", HistoryReportCatalog.LastKey))?.ToString())
                ?? HistoryReportCatalog.Default;
            _cmbReport.SelectedIndex = HistoryReportCatalog.FindIndex(lastKind);

            _statusTxt = new TextBlock {
                FontSize = 12, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 8, 0),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            };
            UpdateStatus();
            // The report-type dropdown does DOUBLE DUTY: it picks what gets
            // exported AND filters the on-screen view. Picking "ເດືອນ ກັນຍາ"
            // hides every other month + summaries so only September's grid
            // stays visible; picking "ສະຫຼຸບພາກຮຽນ 1" shows only that summary,
            // etc. RebuildBody keys off the selected kind string.
            _cmbReport.SelectionChanged += (s, e) =>
            {
                if (_cmbReport.SelectedItem is HistoryReportItem it)
                    DB.SaveSetting(HistoryReportCatalog.LastKey, it.Kind);
                UpdateStatus();
                RebuildBody();
            };

            var bXl    = H.Btn("📋 Excel", "SuccessButton"); bXl.Click    += (s,e) => Export(false);
            var bPdf   = H.Btn("📄 PDF",   "DangerButton");  bPdf.Click   += (s,e) => Export(true);
            var bClose = H.Btn("ປິດ",       "NeutralButton"); bClose.Click += (s,e) => Close();

            bar.Children.Add(H.Lbl("ປີ:"));            bar.Children.Add(_cmbExportYear);
            bar.Children.Add(H.Lbl("ປະເພດລາຍງານ:"));   bar.Children.Add(_cmbReport);
            bar.Children.Add(_statusTxt);
            bar.Children.Add(bXl); bar.Children.Add(bPdf); bar.Children.Add(bClose);
            actions.Child = bar;
            Grid.SetRow(actions, 2); root.Children.Add(actions);

            // First paint — reflect the currently-picked report kind.
            RebuildBody();

            Content = root;
        }

        // Rebuild the scrollable body applying the currently-picked report kind.
        // The 11-option `_cmbReport` catalog drives BOTH the view filter and
        // the export target so teachers see exactly what will be exported.
        //   M9-M12, M2-M5  → show only that month's grid across every year
        //   S1 / S2        → show only that semester's summary card
        //   A              → show only the annual summary
        private void RebuildBody()
        {
            _body.Children.Clear();
            if (_years.Rows.Count == 0)
            {
                _body.Children.Add(new TextBlock {
                    Text = "ບໍ່ມີຂໍ້ມູນປະຫວັດຄະແນນ", FontSize = 13, Margin = new Thickness(4),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))
                });
                return;
            }
            string kind = (_cmbReport?.SelectedItem as HistoryReportItem)?.Kind ?? "";
            foreach (DataRow yr in _years.Rows)
            {
                string y = yr["AcademicYear"].ToString() ?? "";
                string g = yr["GradeLevel"].ToString()  ?? "—";
                string r = yr["ClassRoom"].ToString()   ?? "—";
                _body.Children.Add(BuildYearSection(_sid, y, g, r, kind));
            }
        }

        private void UpdateStatus()
        {
            if (_cmbReport?.SelectedItem is HistoryReportItem it)
                _statusTxt.Text = $"✓ {it.Label}";
            else
                _statusTxt.Text = "";
        }

        // ─── Per-year section — content depends on the picked report kind ───
        //
        //  kind == ""        → legacy full layout (all 8 months + summaries)
        //  kind starts "M"   → only that month's grid (M9..M12, M2..M5)
        //  kind == "S1"      → only sem-1 summary card
        //  kind == "S2"      → only sem-2 summary card
        //  kind == "A"       → only annual summary card
        //  anything else     → falls back to full layout
        //
        // Each year still gets its blue header bar so multi-year students'
        // sections stay visually separated.
        private static UIElement BuildYearSection(int sid, string year, string grade, string room, string kind)
        {
            var outer = new StackPanel { Margin = new Thickness(0,0,0,18) };
            var hdr = new Border {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(31, 64, 122)),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12, 8, 12, 8),
                Margin = new Thickness(0, 0, 0, 6),
                Child = new TextBlock {
                    Text = $"📅  ປີ {year}     ·     ຊັ້ນ {grade}     ·     ຫ້ອງ {room}",
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 14, FontWeight = FontWeights.SemiBold
                }
            };
            outer.Children.Add(hdr);

            var monthly = DB.GetHistoryMonthly(sid, year);

            if (kind.StartsWith("M") && int.TryParse(kind.Substring(1), out int m))
            {
                outer.Children.Add(BuildMonthSection(monthly, m));
                return outer;
            }
            if (kind == "S1")
            {
                // Per-subject table (Mid · Final · Total · Level) + summary card below.
                outer.Children.Add(BuildSemSubjectTable("📘  ສະຫຼຸບພາກຮຽນ 1", DB.GetHistorySemester(sid, year, 1)));
                outer.Children.Add(BuildSemSummary("", DB.GetHistorySemesterSummary(sid, year, 1)));
                return outer;
            }
            if (kind == "S2")
            {
                outer.Children.Add(BuildSemSubjectTable("📗  ສະຫຼຸບພາກຮຽນ 2", DB.GetHistorySemester(sid, year, 2)));
                outer.Children.Add(BuildSemSummary("", DB.GetHistorySemesterSummary(sid, year, 2)));
                return outer;
            }
            if (kind == "A")
            {
                // Per-subject annual table (sem1 · sem2 · annual) + summary card below.
                outer.Children.Add(BuildAnnSubjectTable("📕  ສະຫຼຸບປະຈຳປີ", DB.GetHistoryAnnual(sid, year)));
                outer.Children.Add(BuildAnnSummary(DB.GetHistoryAnnualSummary(sid, year)));
                return outer;
            }

            // Fallback — full year breakdown.
            foreach (int mm in DB.MonthsInSemester(1))
                outer.Children.Add(BuildMonthSection(monthly, mm));
            outer.Children.Add(BuildSemSummary("📘  ສະຫຼຸບພາກຮຽນ 1", DB.GetHistorySemesterSummary(sid, year, 1)));

            foreach (int mm in DB.MonthsInSemester(2))
                outer.Children.Add(BuildMonthSection(monthly, mm));
            outer.Children.Add(BuildSemSummary("📗  ສະຫຼຸບພາກຮຽນ 2", DB.GetHistorySemesterSummary(sid, year, 2)));

            outer.Children.Add(BuildAnnSummary(DB.GetHistoryAnnualSummary(sid, year)));
            return outer;
        }

        private static string MonthName(int m) => m switch {
            1  => "ມັງກອນ",   2  => "ກຸມພາ",   3  => "ມີນາ",     4  => "ເມສາ",
            5  => "ພຶດສະພາ", 6  => "ມິຖຸນາ",  9  => "ກັນຍາ",   10 => "ຕຸລາ",
            11 => "ພະຈິກ",   12 => "ທັນວາ",   _  => $"ເດືອນ {m}"
        };

        private static UIElement BuildMonthSection(DataTable monthly, int month)
        {
            var stack = new StackPanel { Margin = new Thickness(0, 4, 0, 4) };
            stack.Children.Add(new TextBlock {
                Text = $"🗓  ເດືອນ {month} ({MonthName(month)})",
                FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(2, 4, 0, 4),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            var rows = monthly.Select($"ເດືອນ = {month}");
            if (rows.Length == 0)
            {
                stack.Children.Add(new TextBlock {
                    Text = "   (ບໍ່ມີຂໍ້ມູນ)", FontSize = 12,
                    Margin = new Thickness(8, 2, 0, 4),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(156,163,175))
                });
                return stack;
            }
            var dt = new DataTable();
            for (int c = 0; c < monthly.Columns.Count; c++)
            {
                string name = monthly.Columns[c].ColumnName;
                if (name == "ພາກ" || name == "ເດືອນ") continue;
                dt.Columns.Add(name, monthly.Columns[c].DataType);
            }
            foreach (DataRow src in rows)
            {
                var nr = dt.NewRow();
                foreach (DataColumn col in dt.Columns) nr[col.ColumnName] = src[col.ColumnName];
                dt.Rows.Add(nr);
            }
            stack.Children.Add(MkHistoryGrid(dt));
            return stack;
        }

        private static UIElement BuildSemSummary(string title,
            (int subjects, double avg, double total, int rank, int classSize, bool failed) s)
        {
            var box = H.MkCard(new Thickness(0, 6, 0, 10), new Thickness(12, 8, 12, 8));
            box.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 246, 255));
            var stack = new StackPanel();
            // Title is optional — when a per-subject table sits above this
            // summary, we skip the header row to keep the sem block compact.
            if (!string.IsNullOrEmpty(title))
            {
                stack.Children.Add(new TextBlock {
                    Text = title, FontSize = 13, FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 0, 0, 4),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
                });
            }
            string rankTxt = s.failed ? "ຕົກ" : (s.rank > 0 ? s.rank.ToString() : "—");
            string level   = s.failed ? "ຕົກ" : DB.CalcMoESLevel(s.avg);
            stack.Children.Add(new TextBlock {
                Text = $"📊  ສະເລ່ຍເດືອນ {s.avg:F2}    ·    ຄະແນນເສັງລວມ {s.total:F2}    ·    ອັນດັບ {rankTxt} / {s.classSize}    ·    ລະດັບ {level}",
                FontSize = 12,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            });
            box.Child = stack;
            return box;
        }

        // Per-subject semester grid — same visual style as BuildMonthSection.
        // dt columns: ລະຫັດວິຊາ · ຊື່ວິຊາ · ສະເລ່ຍເດືອນ · ຄະແນນເສັງ · ລວມພາກ · ລະດັບ
        // Comes from DB.GetHistorySemester which includes CHA1/LAB1 rows
        // (with their manual eval score in the ລວມພາກ column).
        private static UIElement BuildSemSubjectTable(string title, DataTable dt)
        {
            var stack = new StackPanel { Margin = new Thickness(0, 4, 0, 4) };
            stack.Children.Add(new TextBlock {
                Text = title, FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(2, 4, 0, 4),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            if (dt.Rows.Count == 0)
            {
                stack.Children.Add(new TextBlock {
                    Text = "   (ບໍ່ມີຂໍ້ມູນ)", FontSize = 12,
                    Margin = new Thickness(8, 2, 0, 4),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(156,163,175))
                });
                return stack;
            }
            stack.Children.Add(MkHistoryGrid(dt));
            return stack;
        }

        // Per-subject annual grid — dt columns: ລະຫັດວິຊາ · ຊື່ວິຊາ ·
        // ສະເລ່ຍພາກ1 · ສະເລ່ຍພາກ2 · ສະເລ່ຍປະຈຳປີ. Same visual style as monthly.
        private static UIElement BuildAnnSubjectTable(string title, DataTable dt)
        {
            var stack = new StackPanel { Margin = new Thickness(0, 4, 0, 4) };
            stack.Children.Add(new TextBlock {
                Text = title, FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(2, 4, 0, 4),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            if (dt.Rows.Count == 0)
            {
                stack.Children.Add(new TextBlock {
                    Text = "   (ບໍ່ມີຂໍ້ມູນ)", FontSize = 12,
                    Margin = new Thickness(8, 2, 0, 4),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(156,163,175))
                });
                return stack;
            }
            stack.Children.Add(MkHistoryGrid(dt));
            return stack;
        }

        private static UIElement BuildAnnSummary(
            (double sem1Avg, double sem2Avg, double annualAvg, int rank, int classSize, string level, bool failed) a)
        {
            var box = H.MkCard(new Thickness(0, 6, 0, 0), new Thickness(12, 8, 12, 8));
            box.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(254, 252, 232));
            var stack = new StackPanel();
            stack.Children.Add(new TextBlock {
                Text = "📒  ສະຫຼຸບປະຈຳປີ", FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 4),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            string rankTxt  = a.failed ? "ຕົກ" : (a.rank > 0 ? a.rank.ToString() : "—");
            string levelTxt = a.failed ? "ຕົກ" : a.level;
            stack.Children.Add(new TextBlock {
                Text = $"ສະເລ່ຍພາກ 1 {a.sem1Avg:F2}    ·    ສະເລ່ຍພາກ 2 {a.sem2Avg:F2}    ·    ສະເລ່ຍປະຈຳປີ {a.annualAvg:F2}    ·    ອັນດັບປະຈຳປີ {rankTxt} / {a.classSize}    ·    ລະດັບ {levelTxt}",
                FontSize = 12,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            });
            box.Child = stack;
            return box;
        }

        private static DataGrid MkHistoryGrid(DataTable dt)
            => new DataGrid {
                AutoGenerateColumns = true, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                ItemsSource = dt.DefaultView
            };

        // ─── Excel + PDF export — dropdown-driven ──────────────────────────
        // The single Report-Type dropdown carries the user's intent; this
        // method dispatches on its Kind to one of the four Render* entry
        // points on ReportPage. Historical grade/room come from
        // _yrToGradeRoom — NEVER from Students.GradeLevel/ClassRoom.
        private (bool ok, string year, (string grade,string room) gr) ResolveYear()
        {
            if (_cmbExportYear.SelectedItem is not string year || string.IsNullOrEmpty(year))
            {
                MessageBox.Show("ນັກຮຽນຄົນນີ້ບໍ່ມີຂໍ້ມູນປະຫວັດ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return (false, "", ("",""));
            }
            if (!_yrToGradeRoom.TryGetValue(year, out var gr))
            {
                MessageBox.Show("ບໍ່ສາມາດສ້າງຂໍ້ມູນປະຫວັດປີນີ້ໄດ້", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return (false, "", ("",""));
            }
            return (true, year, gr);
        }

        private void Export(bool toPdf)
        {
            if (_cmbReport.SelectedItem is not HistoryReportItem it)
            {
                MessageBox.Show("ກະລຸນາເລືອກປະເພດລາຍງານ", "ຍັງບໍ່ໄດ້ເລືອກ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var (ok, year, gr) = ResolveYear(); if (!ok) return;

            if (it.IsMonth)
            {
                int month = it.Month;
                RunExport(toPdf, $"ໃບຄະແນນປະຈຳເດືອນ_{_label}_ປີ{year}_ເດືອນ{month}",
                    xlsx => ReportPage.RenderIndividualMonthlyXlsx(_sid, year, month, gr.grade, gr.room, xlsx, out _),
                    $"sid={_sid} y={year} m={month} (monthly)");
            }
            else if (it.IsSemester)
            {
                int sem = it.Sem;
                RunExport(toPdf, $"ໃບຄະແນນພາກຮຽນ{sem}_{_label}_ປີ{year}",
                    xlsx => ReportPage.RenderIndividualSemesterXlsx(_sid, year, sem, gr.grade, gr.room, xlsx, out _),
                    $"sid={_sid} y={year} sem={sem}");
            }
            else // IsAnnual
            {
                RunExport(toPdf, $"ໃບຄະແນນປະຈຳປີ_{_label}_ປີ{year}",
                    xlsx => ReportPage.RenderIndividualAnnualXlsx(_sid, year, gr.grade, gr.room, xlsx, out _),
                    $"sid={_sid} y={year} annual");
            }
        }

        // Common save-dialog + render + optional PDF-convert pipeline. The render
        // callback receives the target xlsx path; for PDF we render to a temp xlsx
        // then convert via Excel COM, mirroring ReportPage's PDF flow.
        private void RunExport(bool toPdf, string baseName, Action<string> render, string logDetail)
        {
            string clean = DB.SafeFileName(baseName);
            var dlg = new SaveFileDialog {
                Filter = toPdf ? "PDF (*.pdf)|*.pdf" : "Excel (*.xlsx)|*.xlsx",
                FileName = clean + (toPdf ? ".pdf" : ".xlsx")
            };
            if (dlg.ShowDialog() != true) return;
            string xlsxPath = toPdf
                ? Path.Combine(Path.GetTempPath(), "histi_" + Guid.NewGuid().ToString("N") + ".xlsx")
                : dlg.FileName;
            try
            {
                render(xlsxPath);
                if (toPdf)
                {
                    ReportPage.ConvertXlsxToPdfViaExcel(xlsxPath, dlg.FileName);
                    try { File.Delete(xlsxPath); } catch { }
                }
                DB.Log("StudentHistoryExport", $"{logDetail} {(toPdf ? "pdf" : "xlsx")}");
                MessageBox.Show($"ບັນທຶກສຳເລັດ:\n{dlg.FileName}", "ສຳເລັດ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ບໍ່ສຳເລັດ:\n{ex.Message}", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // ════════════════════════════════════════════════════════════
    //  CLASS HISTORY WINDOW
    //
    //  Class-wide viewer opened by the 📚 button. Sections for a fixed
    //  (year, grade, room):
    //      🗓 per-month grids — rows=students, cols=subjects (monthly /10)
    //      📘 sem-1 summary  — rows=students, cols=avg/total/rank/level
    //      📗 sem-2 summary  — same shape
    //      📒 annual summary — rows=students, cols=sem1/sem2/annual/rank/level
    //  Pulls from DB.GetClassMonthGrid / GetClassSemesterSummary /
    //  GetClassAnnualSummary — all of which call GetHistoricalClassRoster
    //  so the cohort is fixed across all sub-grids.
    // ════════════════════════════════════════════════════════════
    public class ClassHistoryWindow : Window
    {
        // Fields exposed so the report-type dropdown can filter the on-screen
        // body — same double-duty pattern as StudentHistoryWindow: the picked
        // report kind decides both what shows AND what exports.
        private readonly string _year, _grade, _room;
        private StackPanel _body = null!;
        private ComboBox   _cmbReport = null!;
        private TextBlock  _statusTxt = null!;

        public ClassHistoryWindow(string year, string grade, string room)
        {
            _year = year; _grade = grade; _room = room;
            Title = $"ປະຫວັດທັງຫ້ອງ — ປີ {year} · ຊັ້ນ {grade} · ຫ້ອງ {room}";
            Width = 1200; Height = 800;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252));

            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            var hdr = H.MkCard(new Thickness(12,12,12,8), new Thickness(14,10,14,10));
            hdr.Child = new TextBlock {
                Text = $"📚  ປະຫວັດທັງຫ້ອງ: ປີ {year}     ·     ຊັ້ນ {grade}     ·     ຫ້ອງ {room}",
                FontSize = 15, FontWeight = FontWeights.SemiBold,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            };
            Grid.SetRow(hdr, 0); root.Children.Add(hdr);

            _body = new StackPanel { Margin = new Thickness(12,4,12,4) };
            var sv = new ScrollViewer {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = _body
            };
            Grid.SetRow(sv, 1); root.Children.Add(sv);

            // Action bar — the report-type dropdown does DOUBLE DUTY: pick a
            // month / semester / annual → the body rebuilds to show only
            // that section AND the same pick is what gets exported. Removes
            // the "why doesn't the export match what I see" surprise.
            var actions = H.MkCard(new Thickness(12,4,12,12), new Thickness(12,8,12,8));
            var bar = new WrapPanel();
            _cmbReport = new ComboBox { Width = 220, Margin = new Thickness(0,0,8,0) };
            foreach (var it in HistoryReportCatalog.Items) _cmbReport.Items.Add(it);
            string lastKind = (DB.Scalar("SELECT Value FROM Settings WHERE Key=@k",
                null, ("@k", HistoryReportCatalog.LastKey))?.ToString())
                ?? HistoryReportCatalog.Default;
            _cmbReport.SelectedIndex = HistoryReportCatalog.FindIndex(lastKind);

            _statusTxt = new TextBlock {
                FontSize = 12, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8, 0, 8, 0),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(55,65,81))
            };
            UpdateStatus();
            _cmbReport.SelectionChanged += (s, e) =>
            {
                if (_cmbReport.SelectedItem is HistoryReportItem it)
                    DB.SaveSetting(HistoryReportCatalog.LastKey, it.Kind);
                UpdateStatus();
                RebuildBody();
            };

            var bXl    = H.Btn("📋 Excel", "SuccessButton"); bXl.Click    += (s,e) => Export(year, grade, room, _cmbReport, false);
            var bPdf   = H.Btn("📄 PDF",   "DangerButton");  bPdf.Click   += (s,e) => Export(year, grade, room, _cmbReport, true);
            var bClose = H.Btn("ປິດ",       "NeutralButton"); bClose.Click += (s,e) => Close();

            bar.Children.Add(H.Lbl("ປະເພດລາຍງານ:")); bar.Children.Add(_cmbReport);
            bar.Children.Add(_statusTxt);
            bar.Children.Add(bXl); bar.Children.Add(bPdf); bar.Children.Add(bClose);
            actions.Child = bar;
            Grid.SetRow(actions, 2); root.Children.Add(actions);

            // First paint — respect the currently-picked kind.
            RebuildBody();

            Content = root;
        }

        private void UpdateStatus()
        {
            if (_cmbReport?.SelectedItem is HistoryReportItem it)
                _statusTxt.Text = $"✓ {it.Label}";
            else
                _statusTxt.Text = "";
        }

        // Rebuild the on-screen body applying the currently-picked report kind.
        //   M9-M12, M2-M5  → show only that month's class grid
        //   S1 / S2        → show only that semester's summary
        //   A              → show only the annual summary
        private void RebuildBody()
        {
            _body.Children.Clear();

            int rosterCount = DB.ScalarInt(@"
                SELECT COUNT(*) FROM (
                    SELECT 1 FROM GradeHistory gh
                     WHERE gh.AcademicYear > @y AND gh.FromGrade=@g AND IFNULL(gh.ClassRoom,'')=@r
                       AND gh.AcademicYear = (SELECT MIN(gh2.AcademicYear) FROM GradeHistory gh2
                                              WHERE gh2.StudentID = gh.StudentID AND gh2.AcademicYear > @y)
                    UNION
                    SELECT 1 FROM Students s
                     WHERE s.GradeLevel=@g AND IFNULL(s.ClassRoom,'')=@r
                       AND NOT EXISTS (SELECT 1 FROM GradeHistory gh3
                                       WHERE gh3.StudentID=s.StudentID AND gh3.AcademicYear > @y)
                       AND EXISTS (SELECT 1 FROM Enrollments e
                                   WHERE e.StudentID=s.StudentID AND e.AcademicYear=@y))",
                null, ("@y", _year), ("@g", _grade), ("@r", _room));

            if (rosterCount == 0)
            {
                _body.Children.Add(new TextBlock {
                    Text = "ບໍ່ມີນັກຮຽນໃນຫ້ອງນີ້ສຳລັບປີດັ່ງກ່າວ",
                    FontSize = 13, Margin = new Thickness(4),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))
                });
                return;
            }

            var semBg    = new System.Windows.Media.Color { R=239, G=246, B=255, A=255 };
            var annBg    = new System.Windows.Media.Color { R=254, G=252, B=232, A=255 };
            string kind  = (_cmbReport?.SelectedItem as HistoryReportItem)?.Kind ?? "";

            if (kind.StartsWith("M") && int.TryParse(kind.Substring(1), out int mm))
            {
                _body.Children.Add(BuildMonthBlock(_year, _grade, _room, mm));
                return;
            }
            // Semester + annual views now use the SAME rows-per-student ×
            // columns-per-subject grid layout as the monthly view. This gives
            // teachers a subject-by-subject view for a whole class in one glance
            // (matching the monthly grid's shape) instead of the previous per-
            // student summary of aggregates. CHA1/LAB1 columns stay omitted —
            // matches the monthly grid rule + the aggregate exclusion contract.
            if (kind == "S1")
            {
                _body.Children.Add(BuildSubjectGridBlock("📘  ສະຫຼຸບພາກຮຽນ 1",
                    DB.GetClassSemesterGrid(_year, _grade, _room, 1), semBg));
                return;
            }
            if (kind == "S2")
            {
                _body.Children.Add(BuildSubjectGridBlock("📗  ສະຫຼຸບພາກຮຽນ 2",
                    DB.GetClassSemesterGrid(_year, _grade, _room, 2), semBg));
                return;
            }
            if (kind == "A")
            {
                _body.Children.Add(BuildSubjectGridBlock("📒  ສະຫຼຸບປະຈຳປີ",
                    DB.GetClassAnnualGrid(_year, _grade, _room), annBg));
                return;
            }

            // Fallback — full layout.
            foreach (int m in DB.MonthsInSemester(1))
                _body.Children.Add(BuildMonthBlock(_year, _grade, _room, m));
            _body.Children.Add(BuildSubjectGridBlock("📘  ສະຫຼຸບພາກຮຽນ 1",
                DB.GetClassSemesterGrid(_year, _grade, _room, 1), semBg));
            foreach (int m in DB.MonthsInSemester(2))
                _body.Children.Add(BuildMonthBlock(_year, _grade, _room, m));
            _body.Children.Add(BuildSubjectGridBlock("📗  ສະຫຼຸບພາກຮຽນ 2",
                DB.GetClassSemesterGrid(_year, _grade, _room, 2), semBg));
            _body.Children.Add(BuildSubjectGridBlock("📒  ສະຫຼຸບປະຈຳປີ",
                DB.GetClassAnnualGrid(_year, _grade, _room), annBg));
        }

        // ─── Excel + PDF exports (template-based) ─────────────────────────
        // All four use Templates/ໃບຄະແນນ.xlsx exactly. The cohort is the
        // HISTORICAL roster (DB.GetHistoricalClassRoster) so graduated +
        // promoted students still appear in past-year reports.
        private static (bool ok, DataTable roster) ResolveRoster(string year, string grade, string room)
        {
            var raw = DB.GetHistoricalClassRoster(year, grade, room);
            if (raw.Rows.Count == 0)
            {
                MessageBox.Show("ບໍ່ມີນັກຮຽນໃນຫ້ອງນີ້ສຳລັບປີດັ່ງກ່າວ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return (false, raw);
            }
            // Render*Xlsx wants StudentID + StudentCode + FullName columns.
            // GetHistoricalClassRoster's columns are Lao-named. Reshape.
            var ros = new DataTable();
            ros.Columns.Add("StudentID", typeof(int));
            ros.Columns.Add("StudentCode", typeof(string));
            ros.Columns.Add("FullName", typeof(string));
            ros.Columns.Add("Gender", typeof(string));
            foreach (DataRow src in raw.Rows)
            {
                var nr = ros.NewRow();
                nr["StudentID"]   = Convert.ToInt32(src["StudentID"]);
                nr["StudentCode"] = src["ລະຫັດ"];
                nr["FullName"]    = src["ຊື່ນັກຮຽນ"];
                nr["Gender"]      = src["ເພດ"];
                ros.Rows.Add(nr);
            }
            return (true, ros);
        }

        private static void Export(string year, string grade, string room, ComboBox cmbReport, bool toPdf)
        {
            if (cmbReport.SelectedItem is not HistoryReportItem it)
            {
                MessageBox.Show("ກະລຸນາເລືອກປະເພດລາຍງານ", "ຍັງບໍ່ໄດ້ເລືອກ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var (ok, ros) = ResolveRoster(year, grade, room); if (!ok) return;
            string ng = grade.Replace(".","");

            if (it.IsMonth)
            {
                int month = it.Month;
                RunExport(toPdf, $"ສະຫຼຸບຄະແນນປະຈຳເດືອນ{month}_ປີ{year}_{ng}-{room}",
                    xlsx => ReportPage.RenderClassMonthlyXlsx(year, grade, room, month, ros, xlsx, out _),
                    $"y={year} g={grade} r={room} m={month} (monthly)");
            }
            else if (it.IsSemester)
            {
                int sem = it.Sem;
                RunExport(toPdf, $"ສະຫຼຸບສະເລ່ຍພາກຮຽນ{sem}_ປີ{year}_{ng}-{room}",
                    xlsx => ReportPage.RenderClassSemesterXlsx(year, grade, room, sem, ros, xlsx, out _),
                    $"y={year} g={grade} r={room} sem={sem}");
            }
            else // IsAnnual
            {
                RunExport(toPdf, $"ສະຫຼຸບຄະແນນປະຈຳປີ_ປີ{year}_{ng}-{room}",
                    xlsx => ReportPage.RenderClassAnnualXlsx(year, grade, room, ros, xlsx, out _),
                    $"y={year} g={grade} r={room} annual");
            }
        }

        private static void RunExport(bool toPdf, string baseName, Action<string> render, string logDetail)
        {
            string clean = DB.SafeFileName(baseName);
            var dlg = new SaveFileDialog {
                Filter = toPdf ? "PDF (*.pdf)|*.pdf" : "Excel (*.xlsx)|*.xlsx",
                FileName = clean + (toPdf ? ".pdf" : ".xlsx")
            };
            if (dlg.ShowDialog() != true) return;
            string xlsxPath = toPdf
                ? Path.Combine(Path.GetTempPath(), "histc_" + Guid.NewGuid().ToString("N") + ".xlsx")
                : dlg.FileName;
            try
            {
                render(xlsxPath);
                if (toPdf)
                {
                    ReportPage.ConvertXlsxToPdfViaExcel(xlsxPath, dlg.FileName);
                    try { File.Delete(xlsxPath); } catch { }
                }
                DB.Log("ClassHistoryExport", $"{logDetail} {(toPdf ? "pdf" : "xlsx")}");
                MessageBox.Show($"ບັນທຶກສຳເລັດ:\n{dlg.FileName}", "ສຳເລັດ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ບໍ່ສຳເລັດ:\n{ex.Message}", "ຜິດພາດ",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string MonthName(int m) => m switch {
            1 => "ມັງກອນ", 2 => "ກຸມພາ", 3 => "ມີນາ", 4 => "ເມສາ", 5 => "ພຶດສະພາ",
            6 => "ມິຖຸນາ", 9 => "ກັນຍາ", 10 => "ຕຸລາ", 11 => "ພະຈິກ", 12 => "ທັນວາ",
            _ => $"ເດືອນ {m}"
        };

        private static UIElement BuildMonthBlock(string year, string grade, string room, int month)
        {
            var stack = new StackPanel { Margin = new Thickness(0, 4, 0, 4) };
            stack.Children.Add(new TextBlock {
                Text = $"🗓  ເດືອນ {month} ({MonthName(month)})",
                FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(2, 4, 0, 4),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            var dt = DB.GetClassMonthGrid(year, grade, room, month);
            stack.Children.Add(new DataGrid {
                AutoGenerateColumns = true, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                ItemsSource = dt.DefaultView
            });
            return stack;
        }

        private static UIElement BuildSummaryBlock(string title, DataTable dt, System.Windows.Media.Color bg)
        {
            var box = H.MkCard(new Thickness(0, 6, 0, 10), new Thickness(12, 8, 12, 8));
            box.Background = new System.Windows.Media.SolidColorBrush(bg);
            var stack = new StackPanel();
            stack.Children.Add(new TextBlock {
                Text = title, FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            stack.Children.Add(new DataGrid {
                AutoGenerateColumns = true, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White,
                HeadersVisibility = DataGridHeadersVisibility.Column,
                ItemsSource = dt.DefaultView
            });
            box.Child = stack;
            return box;
        }

        // Subject-per-column class grid — same visual shape as BuildMonthBlock
        // (rows=students, columns=subjects) but used for semester + annual views.
        // Takes a pre-loaded DataTable so callers can pick between
        // GetClassSemesterGrid / GetClassAnnualGrid. The soft background colour
        // matches BuildSummaryBlock so semester/annual stay visually distinct
        // from the plain-white monthly grids above.
        private static UIElement BuildSubjectGridBlock(string title, DataTable dt, System.Windows.Media.Color bg)
        {
            var box = H.MkCard(new Thickness(0, 6, 0, 10), new Thickness(12, 8, 12, 8));
            box.Background = new System.Windows.Media.SolidColorBrush(bg);
            var stack = new StackPanel();
            stack.Children.Add(new TextBlock {
                Text = title, FontSize = 13, FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6),
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(17,24,39))
            });
            if (dt.Rows.Count == 0)
            {
                stack.Children.Add(new TextBlock {
                    Text = "(ບໍ່ມີຂໍ້ມູນ)", FontSize = 12,
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(156,163,175))
                });
            }
            else
            {
                stack.Children.Add(new DataGrid {
                    AutoGenerateColumns = true, IsReadOnly = true, CanUserAddRows = false,
                    BorderThickness = new Thickness(0), Background = System.Windows.Media.Brushes.White,
                    HeadersVisibility = DataGridHeadersVisibility.Column,
                    ItemsSource = dt.DefaultView
                });
            }
            box.Child = stack;
            return box;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  SHARED HELPERS
    // ════════════════════════════════════════════════════════════
    internal static class H
    {
    public static Grid MkGrid(params GridLength[] rows)
    {
        var g=new Grid();
        foreach(var r in rows) g.RowDefinitions.Add(new RowDefinition{Height=r});
        return g;
    }

    public static Border MkCard(Thickness? margin=null,Thickness? padding=null) =>
        new Border{
            Background=System.Windows.Media.Brushes.White,
            CornerRadius=new CornerRadius(10),
            Padding=padding??new Thickness(0),
            Margin=margin??new Thickness(0),
            BorderBrush=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
            BorderThickness=new Thickness(1)
        };

    public static TextBlock Lbl(string t) =>
        new TextBlock{Text=t,VerticalAlignment=VerticalAlignment.Center,
            Margin=new Thickness(0,0,5,0),FontSize=13,
            Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))};

    public static Button Btn(string text,string style)
    {
        var b=new Button{Content=text,Margin=new Thickness(0,0,6,0),Height=34,
            Padding=new Thickness(14,0,14,0),Cursor=System.Windows.Input.Cursors.Hand};
        b.SetResourceReference(Button.StyleProperty,style);
        return b;
    }

    public static ComboBox MkCmb(string[] items,double w)
    {
        var c=new ComboBox{Width=w,Margin=new Thickness(0,0,8,0)};
        foreach(var i in items) c.Items.Add(i);
        c.SelectedIndex=0;
        return c;
    }

    public static DataGridTextColumn Col(string hdr,string path,double w,bool ro=false) =>
        new DataGridTextColumn{Header=hdr,Binding=new System.Windows.Data.Binding(path),Width=new DataGridLength(w),IsReadOnly=ro};

    public static DataGridTextColumn ColStar(string hdr,string path,bool ro=false) =>
        new DataGridTextColumn{Header=hdr,Binding=new System.Windows.Data.Binding(path),Width=new DataGridLength(1,DataGridLengthUnitType.Star),IsReadOnly=ro};
    } // end class H
}
