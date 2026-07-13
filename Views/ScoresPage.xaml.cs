using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using StudentSIS.Data;

namespace StudentSIS.Views
{
    // ════════════════════════════════════════════════════════════
    //  ບັນທຶກຄະແນນພາກຮຽນ (semester scores) — class-roster × one-subject
    //
    //  Redesigned to match MonthlyScoresPage's workflow: pick Year +
    //  Grade + Room + Subject + Semester, the grid loads the whole
    //  class for that ONE subject, the teacher fills the final-exam
    //  column down the roster, and Ctrl+S saves the whole class in
    //  one transaction. Switch the Subject combo to start the next
    //  subject — 30 students × 14 subjects is now 14 saves, not 420.
    //
    //  The per-student view that used to live here moved to the
    //  Score History page (sidebar idx 13) — pick a student, press
    //  📖, see the multi-year per-subject breakdown.
    //
    //  CHA1/LAB1 behaviour: picking either subject swaps the grid to
    //  a single /10 EvalScore column. Saves route through
    //  DB.SetEvaluationScore(SEM{N}, code). Academic semester saves
    //  write Scores.FinalScore (Mid is owned by the monthly page via
    //  RecomputeMidFromMonthly) and recompute Total + Level.
    // ════════════════════════════════════════════════════════════
    public partial class ScoresPage : UserControl
    {
        private readonly ObservableCollection<SemesterRow> _rows = new();

        // What the currently-loaded roster represents. Save refuses to write
        // unless the user is still looking at this exact (year, grade, room,
        // subject, sem) tuple — guards against the teacher editing then changing
        // a filter without reloading.
        private string _loadedYear    = "";
        private string _loadedGrade   = "";
        private string _loadedRoom    = "";
        private int    _loadedSubId   = 0;
        private string _loadedSubCode = "";
        private int    _loadedSem     = 0;

        // Auto-load on every filter change (was gated behind a first-manual-load
        // trigger; the ໂຫຼດ button was removed to simplify the workflow).
        private bool _autoReload = true;
        private bool _dirty      = false;

        public ScoresPage()
        {
            InitializeComponent();
            TxtFormula.Text =
                $" ສະເລ່ຍປະຈຳເດືອນ ({DB.MidPct:F0}%) + ສອບເສງພາກຮຽນ ({DB.FinalPct:F0}%) = ຄະແນນລວມ /10  ·  ຜ່ານ ≥ {DB.PassScore}\n" +
                " ‘ສະເລ່ຍປະຈຳເດືອນ’ ຄຳນວນອັດຕະໂນມັດຈາກໜ້າ ‘ຄະແນນປະຈຳເດືອນ’ — ບໍ່ສາມາດແກ້ໄຂໂດຍກົງ\n" +
                " CHA1 (ຄຸນສົມບັດ) ແລະ LAB1 (ການອອກແຮງງານ): ປ້ອນຄະແນນຄ່າດຽວ /10 — ຮອງຮັບ ພາກ 1 · ພາກ 2 · ປະຈຳປີ";

            // Populate filter combos.
            foreach (var y in DB.AcademicYears()) CmbYear.Items.Add(y);
            SelectComboString(CmbYear, DB.CurrentYear);
            foreach (var g in new[] { "ມ.1", "ມ.2", "ມ.3", "ມ.4" }) CmbGrade.Items.Add(g);
            CmbGrade.SelectedIndex = 3;
            foreach (var r in new[] { "1", "2", "3", "4", "5", "6" }) CmbRoom.Items.Add(r);
            CmbRoom.SelectedIndex = 0;
            foreach (var s in new[] { "ກຳລັງຮຽນ", "ຈົບ", "ອອກ", "ທັງໝົດ" }) CmbStatus.Items.Add(s);
            CmbStatus.SelectedIndex = 0;
            CmbSem.Items.Add("ພາກ 1"); CmbSem.Items.Add("ພາກ 2"); CmbSem.Items.Add("ປະຈຳປີ");
            CmbSem.SelectedIndex = DB.CurrentSem - 1;

            ReloadSubjects();
            ApplyImportButtonsEnabled();
            ScoreGrid.ItemsSource = _rows;

            // Keyboard: Ctrl+S = save · F5 = reload · Enter on Final cell handled
            // separately in ScoreGrid_PreviewKeyDown to move down the column.
            PreviewKeyDown += OnPagePreviewKeyDown;
            ScoreGrid.PreviewKeyDown += ScoreGrid_PreviewKeyDown;

            // Nav-context handoff from ClassHubPage: pre-select year/grade/room
            // + Sem (NavSemester = 1/2 → ພາກ 1/2, NavSemester = 3 → ປະຈຳປີ)
            // then auto-load the roster so the teacher lands on the class grid.
            if (!string.IsNullOrEmpty(DB.NavGrade))
            {
                ApplyNavContext();
                DB.ClearNav();
            }
            else if (CmbSubject.SelectedItem != null)
            {
                // No nav-context: still auto-load with the default filters
                // (current year · ມ.4 · room 1 · first subject · current sem).
                Dispatcher.InvokeAsync(() => LoadRoster(manual: false));
            }
        }

        private void ApplyNavContext()
        {
            SelectComboString(CmbGrade, DB.NavGrade);
            if (!string.IsNullOrEmpty(DB.NavRoom))  SelectComboString(CmbRoom, DB.NavRoom);
            if (!string.IsNullOrEmpty(DB.NavYear))  SelectComboString(CmbYear, DB.NavYear);
            if (DB.NavSemester >= 1 && DB.NavSemester <= 3)
                CmbSem.SelectedIndex = DB.NavSemester - 1;
            // Grade may have changed the subject list — refresh it, then annual
            // mode may further filter it to CHA1/LAB1. Both are safe to re-run.
            ReloadSubjects();
            ApplyDisplayMode();
            ApplyImportButtonsEnabled();
            if (CmbSubject.SelectedItem != null)
                Dispatcher.InvokeAsync(() => LoadRoster(manual: false));
        }

        private static void SelectComboString(ComboBox cmb, string value)
        {
            for (int i = 0; i < cmb.Items.Count; i++)
                if (cmb.Items[i]?.ToString() == value) { cmb.SelectedIndex = i; return; }
            if (cmb.Items.Count > 0) cmb.SelectedIndex = 0;
        }

        // ─── Subject combo population ─────────────────────────────
        // Includes every subject valid for the picked grade — academic +
        // CHA1 + LAB1. SubComboItem.Code lets the page swap grid columns
        // (academic ⇄ eval) without a DB round-trip.
        private void ReloadSubjects()
        {
            string grade = CmbGrade.SelectedItem?.ToString() ?? "ມ.4";
            string prev  = (CmbSubject.SelectedItem as SubItem)?.Code ?? "";
            // ANNUAL context only applies to CHA1/LAB1 — academic subjects
            // don't have a yearly-manual score (their annual is computed from
            // Sem1 + Sem2 totals). Filter the picker to match.
            string annualFilter = IsAnnualMode ? "AND SubjectCode IN ('CHA1','LAB1')" : "";
            CmbSubject.Items.Clear();
            var dt = DB.Query($@"
                SELECT SubjectID, SubjectCode, SubjectCode||'  '||SubjectName AS D
                FROM Subjects
                WHERE (GradeLevel=@g OR GradeLevel IS NULL OR GradeLevel='') {annualFilter}
                ORDER BY SortOrder",
                null, ("@g", grade));
            foreach (DataRow r in dt.Rows)
                CmbSubject.Items.Add(new SubItem(
                    Convert.ToInt32(r["SubjectID"]),
                    r["SubjectCode"].ToString()!,
                    r["D"].ToString()!));
            // Try to preserve the previous selection if it still exists for the new grade.
            for (int i = 0; i < CmbSubject.Items.Count; i++)
                if (((SubItem)CmbSubject.Items[i]).Code == prev) { CmbSubject.SelectedIndex = i; return; }
            if (CmbSubject.SelectedIndex < 0 && CmbSubject.Items.Count > 0)
                CmbSubject.SelectedIndex = 0;
        }

        // ─── Filter event handlers ────────────────────────────────

        // Grade change → repopulate subjects (the list depends on grade) then
        // fall through to the standard filter-changed handler.
        private void Grade_Changed(object s, SelectionChangedEventArgs e)
        {
            ReloadSubjects();
            ApplyDisplayMode();
            OnFilterChanged();
        }
        private void Subject_Changed(object s, SelectionChangedEventArgs e)
        {
            ApplyDisplayMode();
            OnFilterChanged();
        }
        // Sem change may flip annual mode → repopulate subjects (CHA1/LAB1 only
        // when annual), then re-run the same downstream steps as Grade_Changed.
        private void Sem_Changed(object s, SelectionChangedEventArgs e)
        {
            ReloadSubjects();
            ApplyDisplayMode();
            ApplyImportButtonsEnabled();
            OnFilterChanged();
        }
        private void Filter_Changed(object s, SelectionChangedEventArgs e) => OnFilterChanged();

        // Auto-reload the roster on filter change. If the grid has unsaved
        // edits, ask via the 3-button SaveConfirmDialog first: ບັນທຶກ (save
        // then reload) · ບໍ່ບັນທຶກ (discard then reload) · ຍົກເລີກ (stay put).
        private void OnFilterChanged()
        {
            if (_dirty)
            {
                var decision = SaveConfirmDialog.Ask(Window.GetWindow(this));
                switch (decision)
                {
                    case SaveConfirmResult.Save:
                        BtnSave_Click(this, new RoutedEventArgs());
                        // BtnSave_Click clears _dirty on success and shows an
                        // error dialog on failure — if it failed, keep _dirty
                        // true and abort the filter change so we don't lose edits.
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
            if (_autoReload && CmbSubject.SelectedItem != null)
                Dispatcher.InvokeAsync(() => LoadRoster(manual: false));
            else
                InvalidateRoster();
        }

        // Retained for legacy call sites; auto-reload always runs since the
        // ໂຫຼດ button was removed, so this path only fires when the user
        // cancels a dirty prompt.
        private void InvalidateRoster()
        {
            if (_loadedSubId == 0) return;
            _rows.Clear();
            _loadedYear = ""; _loadedGrade = ""; _loadedRoom = "";
            _loadedSubId = 0; _loadedSubCode = ""; _loadedSem = 0;
            TxtSaveInfo.Text = "🔄 ຕົວກອງປ່ຽນແລ້ວ — ໂຫຼດຂໍ້ມູນໃໝ່ໃນອີກຄາວ";
        }

        // ─── Display mode (academic vs eval) ──────────────────────
        // Mirrors MonthlyScoresPage.ApplyDisplayMode — toggles column visibility
        // depending on whether the picked subject is CHA1/LAB1 or academic.
        private string SelectedSubjectCode =>
            (CmbSubject.SelectedItem as SubItem)?.Code ?? "";
        private bool IsEvalSubject =>
            SelectedSubjectCode == "CHA1" || SelectedSubjectCode == "LAB1";
        // Third CmbSem option = ປະຈຳປີ. Writes/reads EvaluationScores with
        // Context="ANNUAL" instead of "SEM1"/"SEM2". Only CHA1/LAB1 apply.
        private bool IsAnnualMode => CmbSem.SelectedIndex == 2;

        // Import/Download templates still use the SIS_LAO_SEMESTER_v4 schema
        // (sem 1/2 only). Annual bulk import isn't supported yet — disable
        // the buttons to make the limit obvious.
        private void ApplyImportButtonsEnabled()
        {
            bool annual = IsAnnualMode;
            if (BtnDownload != null) BtnDownload.IsEnabled = !annual;
            if (BtnImport   != null) BtnImport.IsEnabled   = !annual;
        }

        private void ApplyDisplayMode()
        {
            bool eval = IsEvalSubject;
            if (ColMid    != null) ColMid.Visibility    = eval ? Visibility.Collapsed : Visibility.Visible;
            if (ColFinal  != null) ColFinal.Visibility  = eval ? Visibility.Collapsed : Visibility.Visible;
            if (ColTotal  != null) ColTotal.Visibility  = eval ? Visibility.Collapsed : Visibility.Visible;
            if (ColLevel  != null) ColLevel.Visibility  = eval ? Visibility.Collapsed : Visibility.Visible;
            if (ColEval   != null) ColEval.Visibility   = eval ? Visibility.Visible   : Visibility.Collapsed;
        }

        // ─── Load roster ─────────────────────────────────────────

        private void LoadRoster(bool manual)
        {
            string year   = CmbYear.SelectedItem?.ToString()   ?? DB.CurrentYear;
            string grade  = CmbGrade.SelectedItem?.ToString()  ?? "ມ.4";
            string room   = CmbRoom.SelectedItem?.ToString()   ?? "1";
            string status = CmbStatus.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            int    sem    = CmbSem.SelectedIndex + 1;

            if (CmbSubject.SelectedItem is not SubItem sub)
            {
                if (manual) MessageBox.Show("ກະລຸນາເລືອກວິຊາ", "ຍັງບໍ່ໄດ້ເລືອກ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string finalHdr = sem == 1
                ? "ສອບເສງພາກຮຽນ 1 ມ.ກ. (/10)"
                : "ສອບເສງພາກຮຽນ 2 ມິ.ຖ. (/10)";
            string midHdr   = sem == 1
                ? "ສະເລ່ຍປະຈຳເດືອນ ກ.ຍ.–ທ.ວ. (/10)"
                : "ສະເລ່ຍປະຈຳເດືອນ ກ.ພ.–ພ.ພ. (/10)";
            ColFinal.Header = finalHdr;
            ColMid.Header   = midHdr;
            ColEval.Header  = IsAnnualMode ? "ຄະແນນປະຈຳປີ (/10)" : "ຄະແນນ (/10)";

            string statusFilter = status == "ທັງໝົດ" ? "" : "AND s.Status=@st";
            var ps = new System.Collections.Generic.List<(string, object)> {
                ("@subId", sub.Id), ("@year", year), ("@sem", sem),
                ("@grade", grade), ("@room", room)
            };
            if (status != "ທັງໝົດ") ps.Add(("@st", status));

            // One query pulls EVERY active student in (year, grade, room)
            // joined to their enrollment + score row for this subject + sem.
            // LEFT JOIN on Enrollments + Scores so students without an
            // enrollment still appear (with EnrollID=0) — Save then skips
            // them with a helpful note rather than silently dropping.
            var dt = DB.Query($@"
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

            _rows.Clear();
            bool isEval = sub.Code == "CHA1" || sub.Code == "LAB1";
            string evalCtx = IsAnnualMode ? "ANNUAL" : $"SEM{sem}";
            int ord = 1;
            foreach (DataRow r in dt.Rows)
            {
                int sid    = Convert.ToInt32(r["StudentID"]);
                double mid = Convert.ToDouble(r["MidScore"]);
                double fin = Convert.ToDouble(r["FinalScore"]);
                double? eval = isEval ? DB.GetEvaluationScore(sid, year, evalCtx, sub.Code) : null;

                var row = new SemesterRow
                {
                    RowNo       = ord++,
                    StudentID   = sid,
                    EnrollID    = Convert.ToInt32(r["EnrollID"]),
                    ScoreID     = Convert.ToInt32(r["ScoreID"]),
                    StudentCode = r["StudentCode"].ToString()!,
                    FullName    = r["FullName"].ToString()!,
                    MidScore    = isEval ? 0 : mid,
                    FinalScore  = isEval ? 0 : fin,
                    EvalScore   = eval,
                    IsEval      = isEval,
                };
                row.Recalc();
                _rows.Add(row);
            }

            _loadedYear   = year;
            _loadedGrade  = grade;
            _loadedRoom   = room;
            _loadedSubId  = sub.Id;
            _loadedSubCode= sub.Code;
            _loadedSem    = sem;
            _dirty        = false;
            _autoReload   = true;

            string statusLabel = status == "ທັງໝົດ" ? "ທຸກສະຖານະ" : status;
            string lockHint    = status == "ກຳລັງຮຽນ" ? "" : "  ·  🔒 ອ່ານຢ່າງດຽວ";
            int enrolled = 0; foreach (var x in _rows) if (x.EnrollID > 0) enrolled++;
            // Not-enrolled warning is meaningless for ANNUAL (CHA1/LAB1 aren't
            // enrolled — they live in EvaluationScores keyed by student).
            string notEnrolled = (!IsAnnualMode && enrolled < _rows.Count)
                ? $"  ·  ⚠ {_rows.Count - enrolled} ບໍ່ໄດ້ລົງທະບຽນວິຊານີ້"
                : "";
            string semLabel = IsAnnualMode ? "ປະຈຳປີ" : $"ພາກ {sem}";
            TxtSaveInfo.Text = $"📚 {sub.Code}  |  {semLabel}  |  {_rows.Count} ຄົນ ({statusLabel}){notEnrolled}{lockHint}";

            ApplyReadOnly();
            if (_rows.Count == 0 && manual)
                MessageBox.Show("ບໍ່ມີນັກຮຽນທີ່ກົງກັບຕົວກອງ", "ບໍ່ມີຂໍ້ມູນ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Edits are allowed only when status filter = ‘ກຳລັງຮຽນ’. Graduates +
        // withdrawn students are frozen for historical reference.
        private void ApplyReadOnly()
        {
            string status = CmbStatus.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ";
            bool readOnly = status != "ກຳລັງຮຽນ";
            ScoreGrid.IsReadOnly = readOnly;
            BtnSave.IsEnabled    = !readOnly;
        }

        private void MarkDirty()
        {
            if (_dirty) return;
            _dirty = true;
            TxtSaveInfo.Text = "✏ ມີຂໍ້ມູນຍັງບໍ່ໄດ້ບັນທຶກ — ກົດ Ctrl+S ເພື່ອບັນທຶກ";
        }

        // ─── Input filter + commit validator ──────────────────────
        // Same strict integer rule as MonthlyScoresPage. Decimal separators
        // (`.` `,`), letters, signs are blocked at keystroke time.
        private void ScoreGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if ((e.Column == ColFinal || e.Column == ColEval) && e.EditingElement is TextBox tb)
            {
                tb.PreviewTextInput -= IntOnly_PreviewTextInput;
                tb.PreviewTextInput += IntOnly_PreviewTextInput;
                tb.SelectAll();
            }
        }

        private static void IntOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
                if (c < '0' || c > '9') { e.Handled = true; return; }
        }

        private void ScoreGrid_CellEditEnding(object s, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction != DataGridEditAction.Commit) return;
            if (e.Column != ColFinal && e.Column != ColEval) return;
            if (e.EditingElement is not TextBox tb) return;
            if (!DB.TryParseIntScore(tb.Text, 0, 10, out int parsed))
            {
                MessageBox.Show(
                    $"‘{tb.Text}’ ບໍ່ແມ່ນຄະແນນທີ່ຖືກຕ້ອງ.\n" +
                    "ກະລຸນາໃສ່ຕົວເລກເຕັມລະຫວ່າງ 0 ຫາ 10 (ບໍ່ມີທົດສະນິຍົມ).",
                    "ຄະແນນຜິດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                e.Cancel = true;
                Dispatcher.InvokeAsync(() => ScoreGrid.CancelEdit(DataGridEditingUnit.Cell));
                return;
            }
            tb.Text = parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (e.Row.Item is SemesterRow row)
                Dispatcher.InvokeAsync(() => { row.Recalc(); MarkDirty(); });
        }

        // ─── Keyboard ─────────────────────────────────────────────

        // Page-level: Ctrl+S = save, F5 = reload.
        private void OnPagePreviewKeyDown(object sender, KeyEventArgs e)
        {
            bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
            if (e.Key == Key.F5)                  { LoadRoster(manual: true); e.Handled = true; }
            else if (ctrl && e.Key == Key.S)      { BtnSave_Click(sender, new RoutedEventArgs()); e.Handled = true; }
        }

        // Grid-level: Enter on the Final / Eval column commits the cell and
        // moves the selection DOWN to the same column on the next row. This
        // is the primary keyboard-grading flow — type score, Enter, repeat.
        private void ScoreGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            var col = ScoreGrid.CurrentColumn;
            if (col != ColFinal && col != ColEval) return;

            // Commit whatever's in the editor before moving.
            ScoreGrid.CommitEdit(DataGridEditingUnit.Cell, true);

            int idx = ScoreGrid.SelectedIndex;
            if (idx >= 0 && idx < _rows.Count - 1)
            {
                ScoreGrid.SelectedIndex = idx + 1;
                var nextRow = ScoreGrid.Items[idx + 1];
                ScoreGrid.ScrollIntoView(nextRow, col);
                ScoreGrid.CurrentCell = new DataGridCellInfo(nextRow, col);
                ScoreGrid.BeginEdit();
            }
            e.Handled = true;
        }

        // ─── Save ─────────────────────────────────────────────────

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (ScoreGrid.IsReadOnly)
            {
                MessageBox.Show("ບໍ່ສາມາດແກ້ໄຂຄະແນນຂອງນັກຮຽນທີ່ບໍ່ໄດ້ກຳລັງຮຽນ — ປ່ຽນຕົວກອງ ‘ສະຖານະ’ ກັບໄປເປັນ ‘ກຳລັງຮຽນ’ ກ່ອນ",
                    "ອ່ານຢ່າງດຽວ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (_rows.Count == 0)
            {
                MessageBox.Show("ບໍ່ມີຂໍ້ມູນທີ່ຈະບັນທຶກ — ກະລຸນາໂຫຼດກ່ອນ",
                    "ບໍ່ມີຂໍ້ມູນ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            ScoreGrid.CommitEdit(DataGridEditingUnit.Row, true);

            // Save always writes to the tuple that was LOADED (year/sem/subject
            // at load time), not to whatever the combos currently show. This
            // keeps the SaveConfirmDialog "💾 ບັນທຶກ" path correct: if edits
            // are for the previous filter, they go to the previous filter's
            // rows even though the combos have already advanced to the next
            // one. The old filter-drift MessageBox is no longer needed.
            string year  = _loadedYear;
            string code  = _loadedSubCode;
            int    sem   = _loadedSem;
            bool   annual = _loadedSem == 3;

            int saved = 0, skipped = 0;
            string ctx = annual ? "ANNUAL" : $"SEM{sem}";
            using var conn = DB.Open();
            using var tx   = conn.BeginTransaction();
            try
            {
                foreach (var row in _rows)
                {
                    if (row.IsEval)
                    {
                        // CHA1/LAB1: write whatever's in EvalScore (null deletes).
                        // Even rows without an enrollment can store eval scores —
                        // EvaluationScores is keyed by student, not by enrollment.
                        DB.SetEvaluationScore(row.StudentID, year, ctx, code,
                            row.EvalScore.HasValue
                                ? Math.Max(0, Math.Min(10, row.EvalScore.Value))
                                : (double?)null,
                            conn, tx);
                        saved++;
                    }
                    else
                    {
                        // Academic: must have an Enrollments row to write Scores.
                        if (row.EnrollID == 0) { skipped++; continue; }
                        row.FinalScore = Math.Max(0, Math.Min(10, row.FinalScore));
                        row.Recalc();
                        if (row.ScoreID == 0)
                        {
                            DB.ExecTx(@"INSERT INTO Scores(EnrollID,MidScore,FinalScore,TotalScore,Level)
                                        VALUES(@e,@m,@f,@t,@l)",
                                conn, tx,
                                ("@e", row.EnrollID), ("@m", row.MidScore),
                                ("@f", row.FinalScore), ("@t", row.TotalScore),
                                ("@l", row.Level));
                        }
                        else
                        {
                            DB.ExecTx(@"UPDATE Scores
                                        SET FinalScore=@f, TotalScore=@t, Level=@l,
                                            UpdatedAt=datetime('now','localtime')
                                        WHERE ScoreID=@id",
                                conn, tx,
                                ("@f", row.FinalScore), ("@t", row.TotalScore),
                                ("@l", row.Level), ("@id", row.ScoreID));
                        }
                        saved++;
                    }
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

            string semLabel = annual ? "ປະຈຳປີ" : $"ພາກ {sem}";
            DB.Log("SaveSemesterScores",
                $"{code} {semLabel} — ບັນທຶກ {saved} ຄົນ" + (skipped > 0 ? $" (ຂ້າມ {skipped})" : ""));
            _dirty = false;
            TxtSaveInfo.Text =
                $"✅ ບັນທຶກ {saved} ຄົນ ({code} · {semLabel}) — {DateTime.Now:HH:mm:ss}"
                + (skipped > 0 ? $"  ·  ⚠ ຂ້າມ {skipped} (ບໍ່ໄດ້ລົງທະບຽນ)" : "");
        }

        // ─── Excel import (per-subject) ───────────────────────────
        // Subject comes from the main CmbSubject combo — no separate subject
        // dropdown for imports anymore. Status must be ‘ກຳລັງຮຽນ’.
        private bool RequireImportFilters(out string year, out string grade, out string room, out int sem, out string subject)
        {
            year    = CmbYear.SelectedItem?.ToString()  ?? "";
            grade   = CmbGrade.SelectedItem?.ToString() ?? "";
            room    = CmbRoom.SelectedItem?.ToString()  ?? "";
            sem     = CmbSem.SelectedIndex + 1;
            subject = (CmbSubject.SelectedItem as SubItem)?.Code ?? "";
            if (string.IsNullOrEmpty(year) || string.IsNullOrEmpty(grade)
                || string.IsNullOrEmpty(room) || string.IsNullOrEmpty(subject))
            {
                MessageBox.Show("ກະລຸນາເລືອກ ປີ · ຊັ້ນ · ຫ້ອງ · ວິຊາ · ພາກ ກ່ອນ",
                    "ຍັງບໍ່ໄດ້ເລືອກ", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            return true;
        }

        private void BtnDownload_Click(object sender, RoutedEventArgs e)
        {
            if (!RequireImportFilters(out string year, out string grade, out string room, out int sem, out string subject)) return;
            var dlg = new SaveFileDialog {
                Filter   = "Excel|*.xlsx",
                FileName = DB.SafeFileName($"ນຳເຂົ້າຄະແນນພາກຮຽນ_{year}_{grade}-{room}_S{sem}_{subject}.xlsx")
            };
            if (dlg.ShowDialog() != true) return;
            try
            {
                ExcelImport.BuildSemesterTemplate(dlg.FileName, year, grade, room, sem, subject);
                TxtSaveInfo.Text = $"✅ ສ້າງແບບຟອມສຳເລັດ ({subject}) — {Path.GetFileName(dlg.FileName)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ສ້າງແບບຟອມບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            if (!RequireImportFilters(out string year, out string grade, out string room, out int sem, out string subject)) return;
            if ((CmbStatus.SelectedItem?.ToString() ?? "ກຳລັງຮຽນ") != "ກຳລັງຮຽນ")
            {
                MessageBox.Show("ການນຳເຂົ້າຮອງຮັບສະເພາະນັກຮຽນທີ່ ‘ກຳລັງຮຽນ’ — ກະລຸນາປ່ຽນຕົວກອງ ‘ສະຖານະ’ ກັບໄປຂ້າງເທິງ",
                    "ອ່ານຢ່າງດຽວ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var dlg = new OpenFileDialog { Filter = "Excel|*.xlsx" };
            if (dlg.ShowDialog() != true) return;

            ImportResult result;
            try { result = ExcelImport.ParseSemester(dlg.FileName, year, grade, room, sem, subject); }
            catch (Exception ex)
            {
                MessageBox.Show($"ອ່ານໄຟລ໌ບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            var win = new ImportPreviewWindow(result) { Owner = Window.GetWindow(this) };
            if (win.ShowDialog() == true)
            {
                // The roster on screen is the same scope as the import, so a
                // single reload picks up every freshly-imported score.
                LoadRoster(manual: false);
                TxtSaveInfo.Text = $"✅ ນຳເຂົ້າ {win.SavedCount} ແຖວ ({subject}) ສຳເລັດ — {DateTime.Now:HH:mm:ss}";
            }
        }
    }

    // Row model — one student in the class roster, scoped to one subject/sem.
    // Computed properties (TotalScore/Level) refresh via Recalc() whenever
    // FinalScore changes, so the grid's read-only summary columns update live.
    public class SemesterRow : INotifyPropertyChanged
    {
        public int    RowNo       { get; set; }
        public int    StudentID   { get; set; }
        public int    EnrollID    { get; set; }    // 0 if not enrolled
        public int    ScoreID     { get; set; }    // 0 if Scores row doesn't exist yet
        public string StudentCode { get; set; } = "";
        public string FullName    { get; set; } = "";
        public bool   IsEval      { get; set; }

        private double  _mid, _fin, _total;
        private string  _level = "";
        private double? _eval;
        public double  MidScore   { get => _mid;   set { if (_mid == value) return; _mid = value; Notify(); Recalc(); } }
        public double  FinalScore { get => _fin;   set { if (_fin == value) return; _fin = value; Notify(); Recalc(); } }
        public double  TotalScore { get => _total; private set { _total = value; Notify(); } }
        public string  Level      { get => _level; private set { _level = value; Notify(); } }
        public double? EvalScore  { get => _eval;  set { if (_eval == value) return; _eval = value; Notify(); Recalc(); } }

        // CHA1/LAB1: Total = EvalScore (manual eval value, no formula).
        // Academic: Total = Mid×% + Final×%; Level = pass/fail band.
        public void Recalc()
        {
            if (IsEval)
            {
                TotalScore = _eval ?? 0;
                Level      = _eval.HasValue ? DB.CalcLevel(_eval.Value) : "";
            }
            else
            {
                TotalScore = DB.CalcTotal(_mid, _fin);
                Level      = DB.CalcLevel(TotalScore);
            }
        }
        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify([CallerMemberName] string? n = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    // Subject combo item — same shape as MonthlyScoresPage's SubComboItem.
    // Display string is `<CODE>  <Name>` for the dropdown; Code is cached
    // so we can decide academic-vs-eval without re-querying the DB.
    public class SubItem
    {
        public int    Id      { get; }
        public string Code    { get; }
        public string Display { get; }
        public SubItem(int id, string code, string display) { Id = id; Code = code; Display = display; }
        public override string ToString() => Display;
    }
}
