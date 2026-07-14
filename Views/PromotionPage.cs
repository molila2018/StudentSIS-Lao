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
            var ids = new List<int>();
            foreach (DataRow rr in cohort.Rows) ids.Add(Convert.ToInt32(rr["StudentID"]));
            var stuDt = DB.GetStudentsByIds(ids);
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
                    DB.GraduateStudent(row.StudentID, row.GradeLevel, row.ClassRoom,
                        DB.NextYearString(row.AcademicYear), conn, tx);
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

                    DB.ApplyPromotion(row.StudentID, row.GradeLevel, toG, newSt,
                        studentRoom, studentYear, finalY, row.ClassRoom,
                        histTo, action, conn, tx);
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

        private void LoadHistory() =>
            _hist.ItemsSource = DB.GetPromotionHistory().DefaultView;
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

}
