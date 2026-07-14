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
            string lastKind = DB.GetSetting(HistoryReportCatalog.LastKey)
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

            // Same cohort rule as everywhere else — count via the canonical
            // roster helper instead of duplicating its UNION query here.
            int rosterCount = DB.GetHistoricalClassRoster(_year, _grade, _room).Rows.Count;

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
}
