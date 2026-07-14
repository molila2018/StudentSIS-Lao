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

            var dt = DB.GetActiveStudentPicker(year, grade == "ທັງໝົດ" ? null : grade);

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
            var dt = DB.GetStudentEnrolledSubjects(sid, year);
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
            string grade = DB.GetStudentGradeLevel(sid);
            string label = _cmbStu.Text;

            // Subjects valid for the student's grade — either explicitly tagged to this
            // grade or grade-agnostic (NULL / empty GradeLevel = applies to all grades).
            // In Buekthong's setup ມ.1–ມ.4 share one subject set, so this normally
            // matches the full 14-subject list maintained on the ‘ວິຊາ’ page.
            var subjects = DB.GetSubjectsForGrade(grade);

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
                    int inserted = DB.EnrollBothSemesters(sid, subId, year, conn, tx);
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

            DB.UnenrollSubject(sid, subId, year);
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
}
