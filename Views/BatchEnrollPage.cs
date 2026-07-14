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

            var dt = DB.GetActiveClassRoster(year, grade, room == "ທັງໝົດ" ? null : room);
            _dg.ItemsSource = dt.DefaultView;

            int subjects = DB.CountSubjectsForGrade(grade);

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

            var subjects = DB.GetSubjectsForGrade(grade);
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
                        int n = DB.EnrollBothSemesters(sid, subId, year, conn, tx);
                        rowsInserted += n;
                        rowsSkipped  += 2 - n;
                        newForThisStudent += n;
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
}
