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

}
