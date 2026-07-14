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

}
