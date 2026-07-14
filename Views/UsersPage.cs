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
    //  USERS PAGE
    // ════════════════════════════════════════════════════════════
    public class UsersPage : UserControl
    {
        private DataGrid _dg=null!;
        public UsersPage(){Build();Load();}
        private void Build(){var root=H.MkGrid(GridLength.Auto,new GridLength(1,GridUnitType.Star));var tb=H.MkCard(new Thickness(0,0,0,10),new Thickness(14,10,14,10));var flow=new WrapPanel();var bAdd=H.Btn("➕  ເພີ່ມ","SuccessButton");bAdd.Click+=(s,e)=>{var f=new UserFormWin(0){Owner=Window.GetWindow(this)};if(f.ShowDialog()==true)Load();};var bEdit=H.Btn("✏️  ແກ້ໄຂ","PrimaryButton");bEdit.Click+=(s,e)=>{if(SelId()>0){var f=new UserFormWin(SelId()){Owner=Window.GetWindow(this)};if(f.ShowDialog()==true)Load();}};var bToggle=H.Btn("🔒  ເປີດ/ປິດ","WarningButton");bToggle.Click+=(s,e)=>{int uid=SelId();if(uid<=0){MessageBox.Show("ກະລຸນາເລືອກຜູ້ໃຊ້ກ່ອນ","ຍັງບໍ່ໄດ້ເລືອກ",MessageBoxButton.OK,MessageBoxImage.Information);return;}if(uid==DB.CurrentUserId){MessageBox.Show("ບໍ່ສາມາດປິດບັນຊີຕົນເອງ");return;}var info=DB.GetUserStatus(uid);if(info==null)return;var(uname,wasActive)=info.Value;DB.ToggleUserActive(uid);DB.Log(wasActive?"DeactivateUser":"ActivateUser",uname);Load();};flow.Children.Add(bAdd);flow.Children.Add(bEdit);flow.Children.Add(bToggle);tb.Child=flow;Grid.SetRow(tb,0);root.Children.Add(tb);_dg=new DataGrid{AutoGenerateColumns=true,IsReadOnly=true,CanUserAddRows=false,BorderThickness=new Thickness(0),Background=System.Windows.Media.Brushes.White};var card=H.MkCard();card.Child=_dg;Grid.SetRow(card,1);root.Children.Add(card);Content=root;}
        private void Load()=>_dg.ItemsSource=DB.GetUsersOverview().DefaultView;
        private int SelId()=>_dg.SelectedItem is DataRowView d?Convert.ToInt32(d["ID"]):0;
    }

    public class UserFormWin : Window
    {
        private const int MinPwdLength = 4;

        private readonly int _id;
        private TextBox _user = null!, _name = null!;
        private PasswordBox _pwdNew = null!, _pwdConfirm = null!;
        private ComboBox _role = null!;

        public UserFormWin(int id)
        {
            _id = id;
            Title = id == 0 ? "➕  ເພີ່ມຜູ້ໃຊ້" : "✏️  ແກ້ໄຂຜູ້ໃຊ້";
            Width = 420; Height = 480;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize; ShowInTaskbar = false;
            Background = System.Windows.Media.Brushes.White;

            var stack = new StackPanel { Margin = new Thickness(20) };

            // ── Section 1: User identity ──
            stack.Children.Add(SectionHeader("ຂໍ້ມູນຜູ້ໃຊ້"));
            stack.Children.Add(FieldLabel("ຊື່ຜູ້ໃຊ້ *"));
            _user = FieldInput(stack);
            stack.Children.Add(FieldLabel("ຊື່ເຕັມ *"));
            _name = FieldInput(stack);
            stack.Children.Add(FieldLabel("ບົດບາດ"));
            _role = new ComboBox { Margin = new Thickness(0, 0, 0, 14) };
            _role.Items.Add("admin");
            _role.Items.Add("teacher");
            _role.SelectedIndex = 1;
            stack.Children.Add(_role);

            // ── Section 2: Password ──
            // For Add (id==0): both fields required + min length.
            // For Edit (id>0): leave both blank to keep the existing password;
            // type into both to change it. Confirm field prevents typos that
            // would otherwise lock the user out of the system.
            string pwdHeader = id == 0
                ? "ລະຫັດຜ່ານ"
                : $"ປ່ຽນລະຫັດຜ່ານ (ເວັ້ນຫາກບໍ່ປ່ຽນ)";
            stack.Children.Add(SectionHeader(pwdHeader));
            stack.Children.Add(FieldLabel(id == 0
                ? $"ລະຫັດຜ່ານ * (ຢ່າງໜ້ອຍ {MinPwdLength} ຕົວ)"
                : $"ລະຫັດຜ່ານໃໝ່ (ຢ່າງໜ້ອຍ {MinPwdLength} ຕົວ)"));
            _pwdNew = new PasswordBox { Margin = new Thickness(0, 0, 0, 10) };
            stack.Children.Add(_pwdNew);
            stack.Children.Add(FieldLabel("ຢືນຢັນລະຫັດຜ່ານ"));
            _pwdConfirm = new PasswordBox { Margin = new Thickness(0, 0, 0, 16) };
            stack.Children.Add(_pwdConfirm);

            // ── Buttons ──
            var fp = new StackPanel {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            string okLabel = id == 0 ? "💾  ສ້າງຜູ້ໃຊ້" : "💾  ບັນທຶກການປ່ຽນແປງ";
            var bOK = H.Btn(okLabel, "SuccessButton");
            bOK.MinWidth = 180;
            bOK.Click += Save;
            var bCan = new Button {
                Content = "ຍົກເລີກ", Width = 80, Height = 34,
                Margin = new Thickness(8, 0, 0, 0), IsCancel = true
            };
            bCan.SetResourceReference(Button.StyleProperty, "NeutralButton");
            fp.Children.Add(bOK); fp.Children.Add(bCan);
            stack.Children.Add(fp);
            Content = stack;

            if (id > 0) LoadData();
            // Focus the username field so it's clear where to start typing.
            Loaded += (_, _) => _user.Focus();
        }

        private void LoadData()
        {
            var dt = DB.GetUserBasic(_id);
            if (dt.Rows.Count == 0) return;
            var r = dt.Rows[0];
            _user.Text = r["Username"].ToString() ?? "";
            _name.Text = r["FullName"].ToString() ?? "";
            // Pick the matching role item explicitly (don't rely on .Text setter for non-editable combo).
            string role = r["Role"].ToString() ?? "teacher";
            for (int i = 0; i < _role.Items.Count; i++)
                if (string.Equals(_role.Items[i]?.ToString(), role, StringComparison.OrdinalIgnoreCase))
                { _role.SelectedIndex = i; break; }
        }

        private void Save(object s, RoutedEventArgs e)
        {
            // ── Validation ──
            string uname = _user.Text.Trim();
            string fname = _name.Text.Trim();
            string role  = _role.SelectedItem?.ToString() ?? "teacher";
            string pwdNew = _pwdNew.Password;
            string pwdConf = _pwdConfirm.Password;
            bool hasPwd  = !string.IsNullOrEmpty(pwdNew) || !string.IsNullOrEmpty(pwdConf);

            if (string.IsNullOrEmpty(uname))
            {
                MessageBox.Show("ກະລຸນາໃສ່ຊື່ຜູ້ໃຊ້", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _user.Focus(); return;
            }
            if (string.IsNullOrEmpty(fname))
            {
                MessageBox.Show("ກະລຸນາໃສ່ຊື່ເຕັມ", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _name.Focus(); return;
            }
            // Add: password mandatory. Edit: password optional but if any field
            // is filled both must be filled and must match.
            if (_id == 0 && !hasPwd)
            {
                MessageBox.Show("ກະລຸນາໃສ່ລະຫັດຜ່ານ", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                _pwdNew.Focus(); return;
            }
            if (hasPwd)
            {
                if (pwdNew.Length < MinPwdLength)
                {
                    MessageBox.Show($"ລະຫັດຜ່ານຕ້ອງມີຢ່າງໜ້ອຍ {MinPwdLength} ຕົວ",
                        "ລະຫັດຜ່ານສັ້ນເກີນໄປ", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _pwdNew.Focus(); return;
                }
                if (pwdNew != pwdConf)
                {
                    MessageBox.Show("ລະຫັດຜ່ານໃໝ່ ແລະ ການຢືນຢັນບໍ່ກົງກັນ",
                        "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _pwdConfirm.Clear();
                    _pwdConfirm.Focus(); return;
                }
            }

            // ── Persist ──
            try
            {
                if (_id == 0)
                    DB.InsertUser(uname, pwdNew, fname, role);
                else if (hasPwd)
                    DB.UpdateUserWithPassword(_id, uname, pwdNew, fname, role);
                else
                    DB.UpdateUserProfile(_id, uname, fname, role);
            }
            catch (SQLiteException ex) when (ex.Message.Contains("UNIQUE"))
            {
                MessageBox.Show($"ຊື່ຜູ້ໃຊ້ ‘{uname}’ ມີຢູ່ແລ້ວ — ກະລຸນາໃຊ້ຊື່ອື່ນ",
                    "ຊື່ຊ້ຳ", MessageBoxButton.OK, MessageBoxImage.Warning);
                _user.SelectAll(); _user.Focus();
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ບັນທຶກບໍ່ສຳເລັດ:\n{ex.Message}",
                    "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // ── Verify the write actually landed (round-trip read-back) ──
            // This catches silent failures (e.g. param binding bugs, file locks) and
            // gives the admin certainty that the credentials they just set will work.
            int verifyId = _id;
            if (verifyId == 0)
                verifyId = DB.GetUserIdByUsername(uname);

            var check = DB.GetUserBasic(verifyId);
            bool ok = check.Rows.Count == 1
                   && check.Rows[0]["Username"].ToString() == uname
                   && check.Rows[0]["FullName"].ToString() == fname
                   && string.Equals(check.Rows[0]["Role"].ToString(), role,
                                    StringComparison.OrdinalIgnoreCase);

            // Password round-trip: try to authenticate-style match.
            bool pwdOk = !hasPwd || DB.VerifyUserPassword(verifyId, pwdNew);

            if (!ok || !pwdOk)
            {
                MessageBox.Show(
                    "ບັນທຶກສຳເລັດ ແຕ່ກວດສອບແລ້ວຂໍ້ມູນບໍ່ຕົງ — ກະລຸນາລອງໃໝ່ ຫຼື ກວດກາສິດການຂຽນຖານຂໍ້ມູນ.",
                    "ຄຳເຕືອນ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ── Log + success message ──
            string action;
            string detail;
            if (_id == 0)
            {
                action = "AddUser"; detail = uname;
            }
            else if (hasPwd)
            {
                action = "EditUserWithPwd";
                detail = $"{uname} (password updated)";
            }
            else
            {
                action = "EditUser"; detail = uname;
            }
            DB.Log(action, detail);

            string successMsg = _id == 0
                ? $"ສ້າງຜູ້ໃຊ້ ‘{uname}’ ສຳເລັດ"
                : (hasPwd
                    ? $"ບັນທຶກສຳເລັດ:\n   • ຂໍ້ມູນຜູ້ໃຊ້ປ່ຽນແປງແລ້ວ\n   • ລະຫັດຜ່ານປ່ຽນແປງແລ້ວ\n\n" +
                      "ຄຳແນະນຳ: ໃຫ້ຜູ້ໃຊ້ logout ແລ້ວ login ໃໝ່ດ້ວຍລະຫັດຜ່ານໃໝ່."
                    : $"ບັນທຶກຂໍ້ມູນ ‘{uname}’ ສຳເລັດ");
            MessageBox.Show(successMsg, "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);

            DialogResult = true;
            Close();
        }

        // ── Layout helpers ──
        private static TextBlock SectionHeader(string text) => new TextBlock {
            Text = text,
            FontSize = 13, FontWeight = FontWeights.SemiBold,
            Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(27, 79, 138)),
            Margin = new Thickness(0, 6, 0, 8),
        };
        private static TextBlock FieldLabel(string t) => new TextBlock {
            Text = t, FontSize = 12,
            Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107, 114, 128)),
            Margin = new Thickness(0, 0, 0, 4)
        };
        private static TextBox FieldInput(StackPanel p) {
            var tb = new TextBox { Margin = new Thickness(0, 0, 0, 10) };
            p.Children.Add(tb);
            return tb;
        }
    }

}
