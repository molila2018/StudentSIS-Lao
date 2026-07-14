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
            foreach (DataRow r in DB.GetSubjectsForGrade(g).Rows)
                _subject.Items.Add(new SubComboItem(
                    Convert.ToInt32(r["SubjectID"]),
                    r["SubjectCode"].ToString()!,
                    r["Display"].ToString()!));
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
            var dt = DB.GetMonthlyScoreRoster(subId, year, month, grade, room,
                status == "ທັງໝົດ" ? null : status);
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
            var dt = DB.GetEvalScoreRoster(year, ctx, code, grade, room,
                status == "ທັງໝົດ" ? null : status);
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

                    DB.SaveMonthlyAssessment(row.EnrollID, month,
                        row.ActivityScore, row.DisciplineScore, row.HomeworkScore,
                        conn, tx);
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

}
