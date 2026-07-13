using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using StudentSIS.Data;

namespace StudentSIS.Views
{
    public partial class MainWindow : Window
    {
        private Button? _active;
        private string? _activeTitle;
        private readonly DispatcherTimer _clock = new() { Interval = TimeSpan.FromSeconds(1) };
        private readonly Func<UserControl>[] _pages;

        public MainWindow()
        {
            InitializeComponent();
            _pages = new Func<UserControl>[] {
                ()=>new DashboardPage(),       // 0
                ()=>new ClassHubPage(),        // 1   class-first navigation hub
                ()=>new StudentsPage(),        // 2
                ()=>new EnrollmentPage(),      // 3
                ()=>new BatchEnrollPage(),     // 4
                ()=>new ScoresPage(),          // 5   semester (mid+final) view
                ()=>new MonthlyScoresPage(),   // 6   monthly continuous-assessment roster
                ()=>new ReportPage(),          // 7
                ()=>new PromotionPage(),       // 8
                ()=>new SubjectsPage(),        // 9
                ()=>new UsersPage(),           // 10
                ()=>new AcademicYearPage(),    // 11  academic-year management
                ()=>new SettingsPage(),        // 12
                ()=>new ScoreHistoryPage(),    // 13  read-only history browser
            };
            TxtSchool.Text    = DB.SchoolName;
            TxtUserName.Text  = DB.CurrentUser;
            TxtUserRole.Text  = DB.CurrentRole;
            TxtInit.Text      = DB.CurrentUser.Length > 0 ? DB.CurrentUser[0].ToString() : "?";
            if (DB.CurrentRole != "admin")
            {
                NavPromo.Visibility    = Visibility.Collapsed;
                NavUsers.Visibility    = Visibility.Collapsed;
                NavAcadYear.Visibility = Visibility.Collapsed;
                NavSett.Visibility     = Visibility.Collapsed;
            }
            _clock.Tick += (_,_) => { TxtClock.Text=DateTime.Now.ToString("HH:mm:ss  dd/MM/yyyy"); TxtInfo.Text=$"👤 {DB.CurrentUser}  ·  {DB.CurrentRole}  ·  ປີ {DB.CurrentYear}"; };
            _clock.Start();

            // Keyboard shortcuts: F5 to refresh, Ctrl+L to log out
            PreviewKeyDown += MainWindow_PreviewKeyDown;

            NavigateTo(NavDash, "ໜ້າຫຼັກ");
        }

        private void Nav_Click(object s, RoutedEventArgs e)
        {
            if (s is Button b)
            {
                string lbl = b.Content is TextBlock tb ? tb.Text.Trim() : "";
                NavigateTo(b, lbl);
            }
        }

        // Programmatic navigation — used by ClassHubPage's action buttons.
        // Resolves the sidebar button by tag-index and triggers the normal nav flow,
        // which preserves the active-state highlighting + page-cache invariants.
        public void NavigateToIndex(int idx, string title)
        {
            Button? btn = idx switch
            {
                0  => NavDash,    1 => NavClassHub, 2 => NavStud,    3 => NavEnroll,
                4  => NavBatch,   5 => NavScores,   6 => NavMonthly, 7 => NavReport,
                8  => NavPromo,   9 => NavSubj,    10 => NavUsers,
                11 => NavAcadYear, 12 => NavSett, 13 => NavHistory,
                _  => null
            };
            if (btn != null) NavigateTo(btn, title);
        }

        private void NavigateTo(Button btn, string? title=null)
        {
            if (_active != null) _active.Style = (Style)FindResource("NavButton");
            _active = btn;
            btn.Style = (Style)FindResource("NavButtonActive");
            if (title != null) { TxtTitle.Text = title; _activeTitle = title; }
            if (btn.Tag is not string ts || !int.TryParse(ts, out int idx) || idx < 0 || idx >= _pages.Length) return;
            try
            {
                Frame.Content = _pages[idx]();
            }
            catch (Exception ex)
            {
                Frame.Content = new TextBlock {
                    Text = $"⚠  Error:\n{ex.Message}\n\n{ex.StackTrace}",
                    Foreground = System.Windows.Media.Brushes.Red, FontSize = 12,
                    Margin = new Thickness(20), TextWrapping = TextWrapping.Wrap,
                    FontFamily = new System.Windows.Media.FontFamily("Consolas")
                };
            }
        }

        private void BtnRefresh_Click(object s, RoutedEventArgs e) => RefreshActive();

        private void RefreshActive()
        {
            if (_active == null) return;
            NavigateTo(_active, _activeTitle);
        }

        private void MainWindow_PreviewKeyDown(object s, KeyEventArgs e)
        {
            // F5: refresh active page
            if (e.Key == Key.F5)
            {
                RefreshActive();
                e.Handled = true;
                return;
            }
            // Ctrl+L: log out
            if (e.Key == Key.L && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                BtnLogout_Click(s, e);
                e.Handled = true;
                return;
            }
        }

        private void BtnLogout_Click(object s, RoutedEventArgs e)
        {
            if (MessageBox.Show("ຕ້ອງການອອກຈາກລະບົບ?", "ຢືນຢັນ", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                DB.Log("Logout"); _clock.Stop();
                new LoginWindow().Show(); Close();
            }
        }
    }
}
