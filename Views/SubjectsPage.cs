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
    //  SUBJECTS PAGE  — filter + search + cascade-aware delete
    // ════════════════════════════════════════════════════════════
    public class SubjectsPage : UserControl
    {
        private DataGrid  _dg     = null!;
        private ComboBox  _fGrade = null!, _fSem = null!;
        private TextBox   _fSearch= null!;
        private TextBlock _lblCount = null!;
        private static readonly string[] AllGrades = { "ມ.1","ມ.2","ມ.3","ມ.4" };

        public SubjectsPage() { Build(); Load(); }

        private void Build()
        {
            var root = H.MkGrid(GridLength.Auto, new GridLength(1, GridUnitType.Star), GridLength.Auto);

            // ── Filter / action bar ──────────────────────────────
            var bar  = H.MkCard(new Thickness(0,0,0,10), new Thickness(14,10,14,10));
            var flow = new WrapPanel();

            var bAdd  = H.Btn("➕  ເພີ່ມວິຊາ",   "SuccessButton"); bAdd.Click  += (s,e) => OpenForm(0);
            var bEdit = H.Btn("✏️  ແກ້ໄຂ",       "PrimaryButton"); bEdit.Click += (s,e) => EditSelected();
            var bDel  = H.Btn("🗑  ລຶບ",          "DangerButton");  bDel.Click  += (s,e) => DeleteSelected();
            flow.Children.Add(bAdd); flow.Children.Add(bEdit); flow.Children.Add(bDel);

            // Visual divider
            flow.Children.Add(new Border {
                Width = 1, Height = 24, Margin = new Thickness(8,4,8,0),
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235))
            });

            flow.Children.Add(H.Lbl("ຊັ້ນ:"));
            _fGrade = new ComboBox { Width = 90, Margin = new Thickness(0,0,8,0) };
            _fGrade.Items.Add("ທັງໝົດ");
            foreach (var g in AllGrades) _fGrade.Items.Add(g);
            _fGrade.SelectedIndex = 0;
            _fGrade.SelectionChanged += (s,e) => Load();
            flow.Children.Add(_fGrade);

            flow.Children.Add(H.Lbl("ພາກ:"));
            _fSem = new ComboBox { Width = 80, Margin = new Thickness(0,0,8,0) };
            _fSem.Items.Add("ທັງໝົດ"); _fSem.Items.Add("1"); _fSem.Items.Add("2");
            _fSem.SelectedIndex = 0;
            _fSem.SelectionChanged += (s,e) => Load();
            flow.Children.Add(_fSem);

            flow.Children.Add(H.Lbl("ຄົ້ນຫາ:"));
            _fSearch = new TextBox { Width = 200, Margin = new Thickness(0,0,8,0) };
            _fSearch.TextChanged += (s,e) => Load();
            flow.Children.Add(_fSearch);

            _lblCount = new TextBlock {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8,0,0,0),
                FontSize = 12, FontWeight = FontWeights.Medium,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99))
            };
            flow.Children.Add(_lblCount);

            bar.Child = flow; Grid.SetRow(bar, 0); root.Children.Add(bar);

            // ── Data grid ────────────────────────────────────────
            _dg = new DataGrid {
                AutoGenerateColumns = false, IsReadOnly = true, CanUserAddRows = false,
                BorderThickness = new Thickness(0),
                Background = System.Windows.Media.Brushes.White
            };
            _dg.Columns.Add(H.Col("ລຳດັບ",   "SortOrder",   70));
            _dg.Columns.Add(H.Col("ລະຫັດ",   "SubjectCode", 110));
            _dg.Columns.Add(H.ColStar("ຊື່ວິຊາ", "SubjectName"));
            _dg.Columns.Add(H.Col("ຊັ້ນ",      "GradeLevel",  70));
            _dg.Columns.Add(H.Col("ພາກ",      "Semester",    60));
            _dg.Columns.Add(H.Col("ປະເພດ",   "Category",   140));
            _dg.MouseDoubleClick += (s,e) => EditSelected();
            var card = H.MkCard(); card.Child = _dg; Grid.SetRow(card, 1); root.Children.Add(card);

            // ── Help footer ──────────────────────────────────────
            var hint = new Border {
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248,250,252)),
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(12,7,12,7),
                Margin = new Thickness(0,8,0,0)
            };
            hint.Child = new TextBlock {
                Text = "💡 ຫຼັງເພີ່ມວິຊາໃໝ່ ກະລຸນາໄປໜ້າ ‘ລົງທະບຽນ (Batch)’ ເພື່ອລົງທະບຽນວິຊາໃຫ້ນັກຮຽນ ຈຶ່ງຈະບັນທຶກຄະແນນໄດ້",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(75,85,99)),
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetRow(hint, 2); root.Children.Add(hint);

            Content = root;
        }

        private void Load()
        {
            string g  = _fGrade.SelectedItem?.ToString() ?? "ທັງໝົດ";
            string sm = _fSem.SelectedItem?.ToString()   ?? "ທັງໝົດ";
            string kw = _fSearch.Text.Trim();

            var dt = DB.SearchSubjects(
                g  == "ທັງໝົດ" ? null : g,
                sm == "ທັງໝົດ" ? null : int.Parse(sm),
                kw);
            _dg.ItemsSource = dt.DefaultView;

            _lblCount.Text = $"📚 {dt.Rows.Count} ຈາກ {DB.CountSubjects()} ວິຊາ";
        }

        private int SelId() => _dg.SelectedItem is DataRowView d ? Convert.ToInt32(d["ID"]) : 0;

        private void OpenForm(int id)
        {
            var f = new SubjectFormWin(id) { Owner = Window.GetWindow(this) };
            if (f.ShowDialog() == true) Load();
        }

        private void EditSelected()
        {
            int id = SelId();
            if (id <= 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກວິຊາທີ່ຕ້ອງການແກ້ໄຂກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            OpenForm(id);
        }

        private void DeleteSelected()
        {
            int id = SelId();
            if (id <= 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກວິຊາທີ່ຕ້ອງການລຶບກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Show subject name + count of dependent records that will cascade.
            string subjectLabel = DB.GetSubjectLabel(id);
            var (enrollCount, scoreCount, monthCount) = DB.GetSubjectCascadeCounts(id);

            string warn = $"ຕ້ອງການລຶບວິຊານີ້ບໍ?\n\n  •  {subjectLabel}\n";
            if (enrollCount + scoreCount + monthCount > 0)
            {
                warn += $"\n⚠ ການລຶບຈະຕັດຂໍ້ມູນຕໍ່ໄປນີ້ດ້ວຍ:\n" +
                        $"   • ການລົງທະບຽນ:        {enrollCount}\n" +
                        $"   • ຄະແນນພາກຮຽນ:      {scoreCount}\n" +
                        $"   • ຄະແນນລາຍເດືອນ:    {monthCount}\n\n" +
                        $"ການກະທຳນີ້ບໍ່ສາມາດຍົກເລີກໄດ້.";
            }

            if (MessageBox.Show(warn, "ຢືນຢັນການລຶບ", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                DB.DeleteSubject(id);
                DB.Log("DelSubject", subjectLabel);
                Load();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ລຶບບໍ່ສຳເລັດ: {ex.Message}", "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    // ════════════════════════════════════════════════════════════
    //  SUBJECT FORM  — Add / Edit a subject
    // ════════════════════════════════════════════════════════════
    public class SubjectFormWin : Window
    {
        private readonly int _id;
        private TextBox _code = null!, _name = null!, _sort = null!;
        private ComboBox _grade = null!, _sem = null!, _cat = null!;
        private static readonly string[] AllGrades = { "ມ.1","ມ.2","ມ.3","ມ.4" };
        private static readonly string[] AllCats   = { "ວິຊາສາມັນ", "ວິຊາເລືອກ", "ວິຊາສ້າງສັນ" };

        public SubjectFormWin(int id)
        {
            _id   = id;
            Title = id == 0 ? "➕  ເພີ່ມວິຊາໃໝ່" : "✏️  ແກ້ໄຂວິຊາ";
            Width = 460; Height = 460;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));
            UseLayoutRounding = true; SnapsToDevicePixels = true;

            var root  = H.MkGrid(new GridLength(1, GridUnitType.Star), GridLength.Auto);
            var card  = H.MkCard(new Thickness(16,16,16,8), new Thickness(20));
            var stack = new StackPanel();

            stack.Children.Add(FL("ລະຫັດວິຊາ *  (ຫ້າມຊ້ຳ — ຕົວຢ່າງ: MATH1, LAO2)"));
            _code = FI(stack);
            stack.Children.Add(FL("ຊື່ວິຊາ *"));
            _name = FI(stack);
            _grade = FC(stack, "ຊັ້ນ", AllGrades, defaultValue: "ມ.4");
            _sem   = FC(stack, "ພາກ", new[]{"1","2"});
            _cat   = FC(stack, "ປະເພດ", AllCats);
            stack.Children.Add(FL("ລຳດັບການສະແດງຜົນ  (ນ້ອຍ → ໃຫຍ່)"));
            _sort = FI(stack, "0");

            card.Child = stack;
            Grid.SetRow(card, 0); root.Children.Add(card);

            // Footer with action buttons
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
            if (id > 0) LoadData();
        }

        private void LoadData()
        {
            var dt = DB.GetSubject(_id);
            if (dt.Rows.Count == 0) return;
            var r = dt.Rows[0];
            _code.Text = r["SubjectCode"].ToString() ?? "";
            _name.Text = r["SubjectName"].ToString() ?? "";
            SC(_grade, r["GradeLevel"].ToString() ?? "");
            SC(_sem,   r["Semester"].ToString()   ?? "1");
            SC(_cat,   r["Category"].ToString()   ?? "");
            _sort.Text = r["SortOrder"].ToString() ?? "0";
        }

        private void Save(object s, RoutedEventArgs e)
        {
            string code = _code.Text.Trim();
            string name = _name.Text.Trim();

            if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(name))
            { MessageBox.Show("ກະລຸນາໃສ່ ລະຫັດ ແລະ ຊື່ວິຊາ", "ຂໍ້ມູນບໍ່ຄົບ", MessageBoxButton.OK, MessageBoxImage.Information); return; }
            if (code.Contains(' '))
            { MessageBox.Show("ລະຫັດວິຊາຫ້າມມີຊ່ອງວ່າງ", "ຂໍ້ມູນບໍ່ຖືກຕ້ອງ", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

            int.TryParse(_sort.Text, out int so);

            try
            {
                int sem = int.Parse(GC(_sem) is { Length: > 0 } v ? v : "1");
                DB.SaveSubject(_id, code, name, GC(_grade), sem, GC(_cat), so);
                DB.Log(_id == 0 ? "AddSubject" : "EditSubject", $"{code} — {name}");
                DialogResult = true; Close();
            }
            catch (SQLiteException ex) when (ex.Message.Contains("UNIQUE"))
            { MessageBox.Show($"ລະຫັດວິຊາ ‘{code}’ ມີຢູ່ແລ້ວ — ກະລຸນາໃຊ້ລະຫັດອື່ນ", "ລະຫັດຊ້ຳ", MessageBoxButton.OK, MessageBoxImage.Warning); }
            catch (Exception ex)
            { MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ: {ex.Message}", "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        // ── Layout helpers ─────────────────────────────────────────
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
        private static ComboBox FC(StackPanel p, string lbl, string[] items, string? defaultValue = null)
        {
            p.Children.Add(FL(lbl));
            var cb = new ComboBox { Margin = new Thickness(0,0,0,12) };
            foreach (var i in items) cb.Items.Add(i);
            int defaultIdx = 0;
            if (defaultValue != null)
                for (int i = 0; i < items.Length; i++)
                    if (items[i] == defaultValue) { defaultIdx = i; break; }
            cb.SelectedIndex = defaultIdx;
            p.Children.Add(cb);
            return cb;
        }
        private static void SC(ComboBox cb, string v)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == v) { cb.SelectedIndex = i; return; }
        }
        private static string GC(ComboBox cb) => cb.SelectedItem?.ToString() ?? "";
    }

}
