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
            string lastKind = DB.GetSetting(HistoryReportCatalog.LastKey)
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

}
