using System.Windows; using System.Windows.Input; using StudentSIS.Data;
namespace StudentSIS.Views
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            DB.Initialize();
            InitializeComponent();
            // Login fields start empty by design — defensively clear in case any
            // tooling auto-populates them. Focus the username field so the user
            // can type immediately without clicking.
            Loaded += (_, _) =>
            {
                TxtUsername.Clear();
                TxtPassword.Clear();
                TxtUsername.Focus();
            };
        }
        private void TitleBar_MouseDown(object s,MouseButtonEventArgs e){if(e.ChangedButton==MouseButton.Left)DragMove();}
        private void BtnClose_Click(object s,RoutedEventArgs e)=>Application.Current.Shutdown();
        private void BtnLogin_Click(object s,RoutedEventArgs e)=>DoLogin();
        private void Input_KeyDown(object s,KeyEventArgs e){if(e.Key==Key.Enter)TxtPassword.Focus();}
        private void Password_KeyDown(object s,KeyEventArgs e){if(e.Key==Key.Enter)DoLogin();}
        private void DoLogin()
        {
            BtnLogin.IsEnabled=false; LblError.Visibility=Visibility.Collapsed;
            var(ok,_,_)=DB.Login(TxtUsername.Text.Trim(),TxtPassword.Password);
            if(ok){new MainWindow().Show();Close();}
            else{LblError.Text="⚠  ຊື່ຜູ້ໃຊ້ ຫຼື ລະຫັດຜ່ານບໍ່ຖືກຕ້ອງ";LblError.Visibility=Visibility.Visible;TxtPassword.Clear();TxtPassword.Focus();BtnLogin.IsEnabled=true;}
        }
    }
}
