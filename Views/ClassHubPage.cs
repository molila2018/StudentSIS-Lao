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
    //  CLASS HUB — class-first navigation (idx 1)
    //
    //  Step 1: pick ມ.1 / ມ.2 / ມ.3 / ມ.4 (each card shows total /
    //          active / graduated counts + the rooms in use)
    //  Step 2: pick the ຫ້ອງ + ປີ filters, see a live stats row
    //          (students, active, subjects enrolled, current sem),
    //          then jump to score-entry / management / reports.
    //
    //  Every action button sets DB.NavGrade/Room/Year/Month/Semester
    //  and calls MainWindow.NavigateToIndex(...) so the destination
    //  page picks the context up on load (ClearNav wipes it after).
    //
    //  Styling: uses the app-wide Card / SectionTitle / PageTitle
    //  styles and the palette brushes (PrimaryBrush, SuccessBrush …)
    //  from Styles/Controls.xaml + Styles/Colors.xaml so the page
    //  matches every other page. No hard-coded RGB.
    // ════════════════════════════════════════════════════════════
    public class ClassHubPage : UserControl
    {
        private static readonly string[] AllGrades = { "ມ.1", "ມ.2", "ມ.3", "ມ.4" };
        private static readonly string[] AllRooms  = { "1", "2", "3", "4", "5", "6" };

        private Grid      _classSelectView = null!;
        private Grid      _hubView         = null!;
        private TextBlock _hubBreadcrumb   = null!;
        private ComboBox  _hubRoom         = null!;
        private ComboBox  _hubYear         = null!;
        private TextBlock _statTotal       = null!;
        private TextBlock _statActive      = null!;
        private TextBlock _statSubjects    = null!;
        private TextBlock _statSem         = null!;
        private string    _selectedGrade   = "";

        public ClassHubPage()
        {
            var root = new Grid();
            _classSelectView = BuildClassSelectView();
            _hubView         = BuildHubView();
            root.Children.Add(_classSelectView);
            root.Children.Add(_hubView);
            Content = root;
            ShowClassSelect();
        }

        // ─── Style + brush helpers ───────────────────────────────
        // Pull styles/brushes from the merged App resources so the page
        // matches every other page — no locally-defined colours.
        private static Border StyledCard(Thickness? margin = null, Thickness? padding = null)
        {
            var b = new Border();
            b.SetResourceReference(Border.StyleProperty, "Card");
            if (margin.HasValue)  b.Margin  = margin.Value;
            if (padding.HasValue) b.Padding = padding.Value;
            return b;
        }
        private static TextBlock Styled(string styleKey, string text)
        {
            var t = new TextBlock { Text = text };
            t.SetResourceReference(TextBlock.StyleProperty, styleKey);
            return t;
        }
        private static System.Windows.Media.SolidColorBrush Brush(string key) =>
            (System.Windows.Media.SolidColorBrush)Application.Current.FindResource(key);

        // ── View 1: Class-selection landing ──────────────────────
        private Grid BuildClassSelectView()
        {
            var g = new Grid();
            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var stack  = new StackPanel();

            // Header card: title + subtitle (left) · current-year badge (right).
            var headerCard = StyledCard(new Thickness(0, 0, 0, 16), new Thickness(20, 14, 20, 14));
            var headerDock = new DockPanel();
            var yearBadge  = new Border {
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(10, 4, 10, 4),
                Background   = Brush("PrimaryLightBrush"),
                VerticalAlignment = VerticalAlignment.Center
            };
            yearBadge.Child = new TextBlock {
                Text = $"📅  ປີການສຶກສາ {DB.CurrentYear}",
                FontSize = 12, FontWeight = FontWeights.SemiBold,
                Foreground = Brush("PrimaryDarkBrush")
            };
            DockPanel.SetDock(yearBadge, Dock.Right);
            headerDock.Children.Add(yearBadge);

            var headerText = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            headerText.Children.Add(Styled("PageTitle", "🏫  ໜ້າຫ້ອງຮຽນ"));
            headerText.Children.Add(new TextBlock {
                Text = "ເລືອກຊັ້ນ ເພື່ອເຂົ້າສູ່ໜ້າບໍລິຫານຫ້ອງຮຽນ ບັນທຶກຄະແນນ ແລະ ອອກລາຍງານ",
                FontSize = 12, Margin = new Thickness(0, 4, 0, 0),
                Foreground = Brush("TextSecondaryBrush")
            });
            headerDock.Children.Add(headerText);
            headerCard.Child = headerDock;
            stack.Children.Add(headerCard);

            // Grade cards — one column each, live counts per grade.
            var grid = new UniformGrid { Rows = 1, Columns = AllGrades.Length };
            foreach (var grade in AllGrades)
                grid.Children.Add(BuildGradeCard(grade));
            stack.Children.Add(grid);

            // Footer hint.
            stack.Children.Add(new TextBlock {
                Text = "💡  ໂຮງຮຽນຮອງຮັບສະເພາະຊັ້ນ ມ.1 – ມ.4",
                FontSize = 11, Margin = new Thickness(0, 20, 0, 0),
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = Brush("TextMutedBrush")
            });

            scroll.Content = stack;
            g.Children.Add(scroll);
            return g;
        }

        // Grade card: total students, active + graduated pills, rooms in use.
        // Rendered as a Button so click-anywhere navigates; hover swaps the
        // border to PrimaryBrush.
        private Button BuildGradeCard(string grade)
        {
            var (total, active, graduated) = DB.GetGradeCardStats(grade);
            var rooms = DB.GetActiveRoomsForGrade(grade);
            string roomList = rooms.Count > 0
                ? "ຫ້ອງ  " + string.Join(" · ", rooms)
                : "ຍັງບໍ່ມີຫ້ອງ";

            var card = new Border {
                Background       = Brush("BgCardBrush"),
                BorderBrush      = Brush("BorderBrush_"),
                BorderThickness  = new Thickness(1),
                CornerRadius     = new CornerRadius(4),
                Padding          = new Thickness(18, 18, 18, 16)
            };
            var stack = new StackPanel();

            // Top row: grade label + icon badge on the right.
            var top   = new DockPanel { LastChildFill = true };
            var badge = new Border {
                Width = 42, Height = 42, CornerRadius = new CornerRadius(21),
                Background = Brush("PrimaryLightBrush"),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            badge.Child = new TextBlock {
                Text = "🏫", FontSize = 20,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            DockPanel.SetDock(badge, Dock.Right);
            top.Children.Add(badge);
            top.Children.Add(new TextBlock {
                Text = grade, FontSize = 30, FontWeight = FontWeights.Bold,
                Foreground = Brush("PrimaryBrush"),
                VerticalAlignment = VerticalAlignment.Center
            });
            stack.Children.Add(top);

            // Big total count.
            stack.Children.Add(new TextBlock {
                Text = "ຈຳນວນນັກຮຽນ",
                FontSize = 11, Margin = new Thickness(0, 14, 0, 4),
                Foreground = Brush("TextSecondaryBrush")
            });
            var countRow = new StackPanel { Orientation = Orientation.Horizontal };
            countRow.Children.Add(new TextBlock {
                Text = total.ToString(), FontSize = 26, FontWeight = FontWeights.Bold,
                Foreground = Brush("TextPrimaryBrush"),
                VerticalAlignment = VerticalAlignment.Bottom
            });
            countRow.Children.Add(new TextBlock {
                Text = " ຄົນ", FontSize = 12, Margin = new Thickness(4, 0, 0, 4),
                Foreground = Brush("TextSecondaryBrush"),
                VerticalAlignment = VerticalAlignment.Bottom
            });
            stack.Children.Add(countRow);

            // Status pills.
            var pills = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 0) };
            pills.Children.Add(MkPill($"ກຳລັງຮຽນ  {active}",   "SuccessBrush"));
            if (graduated > 0)
                pills.Children.Add(MkPill($"ຈົບ  {graduated}", "NeutralBrush"));
            stack.Children.Add(pills);

            // Rooms line.
            stack.Children.Add(new TextBlock {
                Text = roomList, FontSize = 11, Margin = new Thickness(0, 12, 0, 0),
                Foreground = Brush("TextMutedBrush"), TextWrapping = TextWrapping.Wrap
            });

            // CTA line.
            stack.Children.Add(new TextBlock {
                Text = "ກົດເພື່ອເປີດ →", FontSize = 12,
                Margin = new Thickness(0, 14, 0, 0), FontWeight = FontWeights.SemiBold,
                Foreground = Brush("PrimaryBrush")
            });

            card.Child = stack;

            // Button wrapper — transparent chrome so the Border above owns the look.
            var btn = new Button {
                Margin          = new Thickness(6),
                Cursor          = System.Windows.Input.Cursors.Hand,
                BorderThickness = new Thickness(0),
                Background      = System.Windows.Media.Brushes.Transparent,
                Padding         = new Thickness(0),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                VerticalContentAlignment   = VerticalAlignment.Stretch,
                Content         = card,
                Template        = MkTransparentButtonTemplate()
            };
            btn.MouseEnter += (s, e) => card.BorderBrush = Brush("PrimaryBrush");
            btn.MouseLeave += (s, e) => card.BorderBrush = Brush("BorderBrush_");
            btn.Click      += (s, e) => ShowHub(grade);
            return btn;
        }

        // Bare button template so the card border isn't fought by the
        // default button chrome (blue focus rectangle, default background).
        private static ControlTemplate MkTransparentButtonTemplate()
        {
            var tpl = new ControlTemplate(typeof(Button));
            var cp  = new FrameworkElementFactory(typeof(ContentPresenter));
            cp.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
            cp.SetValue(ContentPresenter.VerticalAlignmentProperty,   VerticalAlignment.Stretch);
            tpl.VisualTree = cp;
            return tpl;
        }

        private static Border MkPill(string text, string brushKey)
        {
            var brush = Brush(brushKey);
            var soft  = System.Windows.Media.Color.FromArgb(30, brush.Color.R, brush.Color.G, brush.Color.B);
            var pill  = new Border {
                CornerRadius = new CornerRadius(3),
                Padding      = new Thickness(8, 2, 8, 2),
                Background   = new System.Windows.Media.SolidColorBrush(soft),
                Margin       = new Thickness(0, 0, 6, 0)
            };
            pill.Child = new TextBlock {
                Text = text, FontSize = 11, FontWeight = FontWeights.SemiBold, Foreground = brush
            };
            return pill;
        }

        // ── View 2: Per-class hub ────────────────────────────────
        private Grid BuildHubView()
        {
            var g      = new Grid();
            var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            var stack  = new StackPanel();

            // ── Header card: back button · breadcrumb · ຫ້ອງ+ປີ filters ──
            var headerCard = StyledCard(new Thickness(0, 0, 0, 12), new Thickness(14, 10, 14, 10));
            var headerDock = new DockPanel();
            var back = H.Btn("◀  ກັບໄປ", "NeutralButton");
            back.Click += (s, e) => ShowClassSelect();
            DockPanel.SetDock(back, Dock.Left);
            headerDock.Children.Add(back);

            var rightFilters = new StackPanel {
                Orientation         = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment   = VerticalAlignment.Center
            };
            rightFilters.Children.Add(H.Lbl("ຫ້ອງ:"));
            _hubRoom = H.MkCmb(AllRooms, 70);
            _hubRoom.SelectionChanged += (s, e) => RefreshHubStats();
            rightFilters.Children.Add(_hubRoom);
            rightFilters.Children.Add(H.Lbl("ປີ:"));
            _hubYear = H.MkCmb(DB.AcademicYears().ToArray(), 110);
            SelectComboValue(_hubYear, DB.CurrentYear);
            _hubYear.SelectionChanged += (s, e) => RefreshHubStats();
            rightFilters.Children.Add(_hubYear);
            DockPanel.SetDock(rightFilters, Dock.Right);
            headerDock.Children.Add(rightFilters);

            _hubBreadcrumb = new TextBlock {
                FontSize = 15, FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0),
                Foreground = Brush("TextPrimaryBrush")
            };
            headerDock.Children.Add(_hubBreadcrumb);
            headerCard.Child = headerDock;
            stack.Children.Add(headerCard);

            // ── Stats row: live counts for the picked (grade, room, year) ──
            _statTotal    = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            _statActive   = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            _statSubjects = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            _statSem      = new TextBlock { FontSize = 26, FontWeight = FontWeights.Bold };
            var statsRow  = new UniformGrid { Rows = 1, Columns = 4, Margin = new Thickness(0, 0, 0, 12) };
            statsRow.Children.Add(BuildStatCard("ນັກຮຽນທັງໝົດ", "👥", _statTotal,    "PrimaryBrush",   "#D6E4F7", new Thickness(0, 0, 8, 0)));
            statsRow.Children.Add(BuildStatCard("ກຳລັງຮຽນ",     "✅", _statActive,   "SuccessBrush",   "#DCFCE7", new Thickness(4, 0, 4, 0)));
            statsRow.Children.Add(BuildStatCard("ວິຊາລົງທະບຽນ", "📚", _statSubjects, "SecondaryBrush", "#EDE3FB", new Thickness(4, 0, 4, 0)));
            statsRow.Children.Add(BuildStatCard("ພາກປັດຈຸບັນ",  "📅", _statSem,      "InfoBrush",      "#CFFAFE", new Thickness(8, 0, 0, 0)));
            stack.Children.Add(statsRow);

            // ── Two-column action panels ──
            var twoCol = new Grid();
            twoCol.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            twoCol.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Left column — score entry (monthly · semester · annual).
            var leftCard  = StyledCard(new Thickness(0, 0, 6, 0), new Thickness(16));
            var leftStack = new StackPanel();
            leftStack.Children.Add(Styled("SectionTitle", "📅  ບັນທຶກຄະແນນປະຈຳເດືອນ"));
            leftStack.Children.Add(SectionSubheader("ພາກ 1 — ກ.ຍ. ຫາ ທ.ວ."));
            AddMonthButton(leftStack, "ກ.ຍ. (09)",  9);
            AddMonthButton(leftStack, "ຕ.ລ. (10)", 10);
            AddMonthButton(leftStack, "ພ.ຍ. (11)", 11);
            AddMonthButton(leftStack, "ທ.ວ. (12)", 12);

            leftStack.Children.Add(new Border { Height = 8 });
            leftStack.Children.Add(SectionSubheader("ພາກ 2 — ກ.ພ. ຫາ ພ.ພ."));
            AddMonthButton(leftStack, "ກ.ພ. (02)", 2);
            AddMonthButton(leftStack, "ມີ.ນ. (03)", 3);
            AddMonthButton(leftStack, "ມ.ສ. (04)", 4);
            AddMonthButton(leftStack, "ພ.ພ. (05)", 5);

            leftStack.Children.Add(new Border { Height = 16 });
            leftStack.Children.Add(Styled("SectionTitle", "📝  ບັນທຶກຄະແນນພາກຮຽນ / ປະຈຳປີ"));
            AddSemButton(leftStack,    "📘  ຄະແນນສອບເສງ ພາກ 1", 1);
            AddSemButton(leftStack,    "📗  ຄະແນນສອບເສງ ພາກ 2", 2);
            AddAnnualButton(leftStack, "📕  ຄະແນນປະຈຳປີ (CHA1 / LAB1)");

            leftCard.Child = leftStack;
            Grid.SetColumn(leftCard, 0);
            twoCol.Children.Add(leftCard);

            // Right column — class management + reports/history.
            var rightCard  = StyledCard(new Thickness(6, 0, 0, 0), new Thickness(16));
            var rightStack = new StackPanel();
            rightStack.Children.Add(Styled("SectionTitle", "👥  ບໍລິຫານຫ້ອງຮຽນ"));
            AddActionButton(rightStack, "👥  ຂໍ້ມູນນັກຮຽນຂອງຫ້ອງ",  "PrimaryButton", GoToStudents);
            AddActionButton(rightStack, "📋  ລົງທະບຽນວິຊາ (Batch)", "PrimaryButton", GoToBatchEnroll);

            rightStack.Children.Add(new Border { Height = 16 });
            rightStack.Children.Add(Styled("SectionTitle", "📊  ລາຍງານ & ປະຫວັດ"));
            AddActionButton(rightStack, "📚  ປະຫວັດຄະແນນທັງຫ້ອງ",   "SecondaryButton", GoToScoreHistory);
            AddActionButton(rightStack, "📄  ໃບຄະແນນ ນັກຮຽນ",       "SecondaryButton", GoToScores);
            AddActionButton(rightStack, "📑  ໃບສັນຍາ / ປະຫວັດ",    "NeutralButton",   GoToProfileReports);

            rightCard.Child = rightStack;
            Grid.SetColumn(rightCard, 1);
            twoCol.Children.Add(rightCard);

            stack.Children.Add(twoCol);
            scroll.Content = stack;
            g.Children.Add(scroll);
            return g;
        }

        // Dashboard-style stat card: label · big number (accent brush) · icon badge.
        private static Border BuildStatCard(string label, string icon, TextBlock numberTb, string accentBrushKey, string iconBg, Thickness margin)
        {
            var card = StyledCard(margin, new Thickness(16, 14, 16, 14));
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var left = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            left.Children.Add(new TextBlock {
                Text = label, FontSize = 11, FontWeight = FontWeights.Medium,
                Foreground = Brush("TextSecondaryBrush")
            });
            numberTb.Foreground = Brush(accentBrushKey);
            numberTb.Margin     = new Thickness(0, 4, 0, 0);
            left.Children.Add(numberTb);
            Grid.SetColumn(left, 0);
            grid.Children.Add(left);

            var iconBadge = new Border {
                Width = 40, Height = 40, CornerRadius = new CornerRadius(20),
                Background = (System.Windows.Media.SolidColorBrush)new System.Windows.Media.BrushConverter().ConvertFromString(iconBg)!,
                VerticalAlignment = VerticalAlignment.Top
            };
            iconBadge.Child = new TextBlock {
                Text = icon, FontSize = 18,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center
            };
            Grid.SetColumn(iconBadge, 1);
            grid.Children.Add(iconBadge);

            card.Child = grid;
            return card;
        }

        private static TextBlock SectionSubheader(string text) =>
            new TextBlock {
                Text = text, FontSize = 11, FontWeight = FontWeights.SemiBold,
                Foreground = Brush("TextSecondaryBrush"),
                Margin = new Thickness(0, 0, 0, 6)
            };

        private void AddMonthButton(StackPanel host, string label, int month)
        {
            var btn = H.Btn($"📝  ຄະແນນ ເດືອນ {label}", "PrimaryButton");
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => GoToMonthly(month);
            host.Children.Add(btn);
        }

        private void AddSemButton(StackPanel host, string label, int sem)
        {
            var btn = H.Btn(label, "SuccessButton");
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => GoToSemesterEntry(sem);
            host.Children.Add(btn);
        }

        private void AddAnnualButton(StackPanel host, string label)
        {
            var btn = H.Btn(label, "WarningButton");
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => GoToAnnualEntry();
            host.Children.Add(btn);
        }

        private void AddActionButton(StackPanel host, string label, string style, Action onClick)
        {
            var btn = H.Btn(label, style);
            btn.HorizontalAlignment = HorizontalAlignment.Stretch;
            btn.Width               = double.NaN;
            btn.Margin              = new Thickness(0, 0, 0, 6);
            btn.HorizontalContentAlignment = HorizontalAlignment.Left;
            btn.Padding             = new Thickness(16, 0, 16, 0);
            btn.Click += (s, e) => onClick();
            host.Children.Add(btn);
        }

        private static void SelectComboValue(ComboBox cb, string value)
        {
            for (int i = 0; i < cb.Items.Count; i++)
                if (cb.Items[i]?.ToString() == value) { cb.SelectedIndex = i; return; }
        }

        // ── View state ───────────────────────────────────────────
        private void ShowClassSelect()
        {
            _classSelectView.Visibility = Visibility.Visible;
            _hubView.Visibility         = Visibility.Collapsed;
        }
        private void ShowHub(string grade)
        {
            _selectedGrade = grade;
            _classSelectView.Visibility = Visibility.Collapsed;
            _hubView.Visibility         = Visibility.Visible;
            RefreshHubStats();
        }

        // Refresh the 4 stat cards + breadcrumb whenever room/year changes.
        private void RefreshHubStats()
        {
            string grade = _selectedGrade;
            string room  = _hubRoom.SelectedItem?.ToString() ?? "1";
            string year  = _hubYear.SelectedItem?.ToString() ?? DB.CurrentYear;

            var (total, active, subjects) = DB.GetHubStats(grade, room, year);

            _statTotal.Text    = total.ToString();
            _statActive.Text   = active.ToString();
            _statSubjects.Text = subjects.ToString();
            _statSem.Text      = DB.CurrentSem.ToString();

            _hubBreadcrumb.Text = $"ໜ້າຫ້ອງຮຽນ  /  ຊັ້ນ {grade}  ·  ຫ້ອງ {room}  ·  ປີ {year}";
        }

        // ── Nav-context handoff ──────────────────────────────────
        private void SetNavGradeRoomYear()
        {
            DB.NavGrade = _selectedGrade;
            DB.NavRoom  = _hubRoom.SelectedItem?.ToString() ?? "1";
            DB.NavYear  = _hubYear.SelectedItem?.ToString() ?? DB.CurrentYear;
        }

        private MainWindow? Main => Window.GetWindow(this) as MainWindow;

        private void GoToMonthly(int month)
        {
            SetNavGradeRoomYear();
            DB.NavMonth = month;
            Main?.NavigateToIndex(6, $"ຄະແນນ ເດືອນ {month:D2} — {_selectedGrade}/{DB.NavRoom}");
        }

        // Sem 1/2 exam entry → ScoresPage (idx 5). NavSemester picked up on load.
        private void GoToSemesterEntry(int sem)
        {
            SetNavGradeRoomYear();
            DB.NavSemester = sem;
            Main?.NavigateToIndex(5, $"ຄະແນນສອບເສງ ພາກ {sem} — {_selectedGrade}/{DB.NavRoom}");
        }

        // Annual CHA1/LAB1 → ScoresPage (idx 5). Sentinel value 3 tells the
        // page to pick CmbSem.SelectedIndex=2 (ປະຈຳປີ) on load.
        private void GoToAnnualEntry()
        {
            SetNavGradeRoomYear();
            DB.NavSemester = 3;
            Main?.NavigateToIndex(5, $"ຄະແນນປະຈຳປີ (CHA/LAB) — {_selectedGrade}/{DB.NavRoom}");
        }

        private void GoToStudents()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(2, $"ນັກຮຽນ — {_selectedGrade}/{DB.NavRoom}");
        }
        private void GoToBatchEnroll()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(4, $"ລົງທະບຽນ (Batch) — {_selectedGrade}");
        }
        // Score-history page (idx 13) — every class-wide score report + student
        // history browser lives here (all score reports migrated off ReportPage).
        private void GoToScoreHistory()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(13, $"ປະຫວັດຄະແນນ — {_selectedGrade}/{DB.NavRoom}");
        }
        private void GoToScores()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(5, $"ບັນທຶກຄະແນນ — {_selectedGrade}/{DB.NavRoom}");
        }
        // ReportPage (idx 7) now holds only Enrollment Agreement + Student
        // Profile — the score reports moved to Score History. Keep this button
        // so the two remaining reports are one click from the class hub.
        private void GoToProfileReports()
        {
            SetNavGradeRoomYear();
            Main?.NavigateToIndex(7, $"ໃບສັນຍາ / ປະຫວັດ — {_selectedGrade}/{DB.NavRoom}");
        }
    }

}
