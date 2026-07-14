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
            // Student counts split by Status; CreatedDate is the date portion
            // of CreatedAt for compactness.
            _dg.ItemsSource = DB.GetAcademicYearsOverview().DefaultView;
        }

        private void RefreshFooterStats()
        {
            var (total, active, graduated, withdrawn, classrooms) = DB.GetSystemWideStudentStats();
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

            // One-transaction cascade inside DB.DeleteAcademicYearCascade —
            // throws (and rolls back) on failure, so data stays intact.
            try
            {
                DB.DeleteAcademicYearCascade(y);
            }
            catch (Exception ex)
            {
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

            var stats = DB.GetYearStats(year);
            int students = stats.Students,  enrolls = stats.Enrollments,
                scores   = stats.Scores,    monthly = stats.Monthly,
                evals    = stats.Evaluations, attend = stats.Attendance,
                history  = stats.History;
            bool hasData = stats.HasData;

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

            var stats = DB.GetYearStats(year);
            int students = stats.Students,   active  = stats.Active,
                graduated = stats.Graduated, withdrawn = stats.Withdrawn,
                rooms    = stats.Classrooms, enrolls = stats.Enrollments,
                scores   = stats.Scores,     monthly = stats.Monthly,
                evals    = stats.Evaluations, attend = stats.Attendance,
                history  = stats.History;

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

}
