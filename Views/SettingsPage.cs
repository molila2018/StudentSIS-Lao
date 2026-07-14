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
    //  SETTINGS PAGE
    // ════════════════════════════════════════════════════════════
    public class SettingsPage : UserControl
    {
        public SettingsPage()
        {
            var tabs=new TabControl{BorderThickness=new Thickness(0)};
            tabs.Items.Add(MkTab("⚙️  ທົ່ວໄປ",    BuildGeneral()));
            tabs.Items.Add(MkTab("📥  ນຳເຂົ້າ CSV",BuildImport()));
            tabs.Items.Add(MkTab("💾  Backup",      BuildBackup()));
            tabs.Items.Add(MkTab("📋  Log",         BuildLog()));
            Content=tabs;
        }

        private TextBox _school=null!;
        private DataGrid _logDg=null!;

        private UIElement BuildGeneral()
        {
            var st = new StackPanel { Margin = new Thickness(20) };
            _school = F(st, "ຊື່ໂຮງຮຽນ", DB.SchoolName);

            // Academic year + semester are managed on the dedicated ‘ປີການສຶກສາ’
            // admin page (which keeps Settings + the AcademicYears registry in sync),
            // so they are intentionally NOT editable here.
            // Pass-score threshold is kept at its database default and is not editable
            // from the UI — it's a school-wide grading rule that shouldn't change
            // mid-year, and exposing it caused confusion.
            st.Children.Add(new TextBlock {
                Text = "ℹ ປີການສຶກສາ ແລະ ພາກຮຽນ ຈັດການຢູ່ໜ້າ ‘ປີການສຶກສາ’",
                FontSize = 11,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),
                Margin = new Thickness(0,0,0,12)
            });

            var btn = B("💾  ບັນທຶກ", "SuccessButton", 180);
            btn.Click += (s,e) => SaveGeneral();
            st.Children.Add(btn);
            return new ScrollViewer { Content = st };
        }

        private void SaveGeneral()
        {
            string school = _school.Text.Trim();
            if (school.Length == 0)
            {
                MessageBox.Show("ກະລຸນາໃສ່ຊື່ໂຮງຮຽນ", "ຂໍ້ມູນບໍ່ຄົບ",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DB.SaveSetting("school_name", school);
            DB.Log("Settings", "ທົ່ວໄປ");
            MessageBox.Show("ບັນທຶກສຳເລັດ", "ສຳເລັດ",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private UIElement BuildImport()
        {
            // Full column list for the new schema (30 fields).
            // 28 columns — NationalID and student Phone removed.
            const string CsvHeader =
                "StudentCode,FirstName,LastName,Gender,BirthDate," +
                "BirthVillage,BirthDistrict,BirthProvince," +
                "Village,District,Province," +
                "FatherName,FatherAge,FatherJob,FatherVillage,FatherDistrict,FatherProvince,FatherPhone," +
                "MotherName,MotherAge,MotherJob,MotherVillage,MotherDistrict,MotherProvince,MotherPhone," +
                "GradeLevel,ClassRoom,AcademicYear";
            const string CsvSample =
                "S001,ສົມຊາຍ,ໃຈດີ,ຊາຍ,01/01/2009," +
                "ນາໂພ,ຫາດຊາຍຟອງ,ວຽງຈັນ," +
                "ດົງໂດກ,ໄຊທານີ,ວຽງຈັນ," +
                "ທ້າວດີ ໃຈດີ,45,ກະສິກຳ,ດົງໂດກ,ໄຊທານີ,ວຽງຈັນ,0201234567," +
                "ນາງດີ ໃຈດີ,42,ຄ້າຂາຍ,ດົງໂດກ,ໄຊທານີ,ວຽງຈັນ,0207654322," +
                "ມ.4,1,2025-2026";

            var st = new StackPanel { Margin = new Thickness(20) };
            st.Children.Add(new TextBlock {
                Text = "ນຳເຂົ້ານັກຮຽນຈາກໄຟລ໌ CSV\n\n" +
                       "ຄໍລຳ (28 ຂໍ້): " + CsvHeader + "\n\n" +
                       "• ຖ້າລະຫັດຊ້ຳ ຈະຂ້າມ\n" +
                       "• ໄຟລ໌ UTF-8\n" +
                       "• ຄໍລຳທີ່ຂາດ ຈະຖືກບັນທຶກເປັນຄ່າຫວ່າງ — ສົ່ງສະເພາະຄໍລຳທີ່ຕ້ອງການ",
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0,0,0,14)
            });
            var bTpl = B("📄  Template", "NeutralButton", 180);
            bTpl.Click += (s,e) => {
                var dlg = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "students_template.csv" };
                if (dlg.ShowDialog() != true) return;
                File.WriteAllText(dlg.FileName, CsvHeader + "\n" + CsvSample + "\n", Encoding.UTF8);
                MessageBox.Show("ບັນທຶກ Template ສຳເລັດ", "ສຳເລັດ",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            };
            st.Children.Add(bTpl);
            st.Children.Add(new TextBlock { Height = 10 });

            var bImp = B("📂  ເລືອກໄຟລ໌ ແລະ ນຳເຂົ້າ", "PrimaryButton", 220);
            bImp.Click += (s,e) => {
                var dlg = new OpenFileDialog { Filter = "CSV|*.csv" };
                if (dlg.ShowDialog() != true) return;
                var lines = File.ReadAllLines(dlg.FileName, Encoding.UTF8);
                if (lines.Length < 2) { MessageBox.Show("ໄຟລ໌ບໍ່ມີຂໍ້ມູນ", "ບໍ່ມີຂໍ້ມູນ", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                var hdrs = lines[0].Split(',');
                int ins = 0, skip = 0, invalid = 0;
                using var conn = DB.Open();
                for (int i = 1; i < lines.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i])) continue;
                    var vals = lines[i].Split(',');
                    string G(string h)
                    {
                        int ix = Array.IndexOf(hdrs, h);
                        return ix >= 0 && ix < vals.Length ? vals[ix].Trim().Trim('"') : "";
                    }
                    int? Age(string h) =>
                        int.TryParse(G(h), out int n) ? n : (int?)null;
                    // Required, NOT NULL columns — a row missing any of these would
                    // either fail the INSERT or create an unusable record that never
                    // shows in the grade/year-filtered views. Skip + count separately.
                    if (string.IsNullOrWhiteSpace(G("StudentCode"))
                        || string.IsNullOrWhiteSpace(G("FirstName"))
                        || string.IsNullOrWhiteSpace(G("LastName"))
                        || string.IsNullOrWhiteSpace(G("GradeLevel"))
                        || string.IsNullOrWhiteSpace(G("AcademicYear")))
                    { invalid++; continue; }
                    try
                    {
                        // Map CSV columns → DTO; DB.ImportStudentRow owns the SQL
                        // (INSERT OR IGNORE — duplicate codes return 0 = skipped).
                        var dto = new StudentDto {
                            Code = G("StudentCode"), FirstName = G("FirstName"), LastName = G("LastName"),
                            Gender = G("Gender"), BirthDate = G("BirthDate"),
                            BirthVillage = G("BirthVillage"), BirthDistrict = G("BirthDistrict"), BirthProvince = G("BirthProvince"),
                            Village = G("Village"), District = G("District"), Province = G("Province"),
                            FatherName = G("FatherName"), FatherAge = Age("FatherAge"), FatherJob = G("FatherJob"),
                            FatherVillage = G("FatherVillage"), FatherDistrict = G("FatherDistrict"),
                            FatherProvince = G("FatherProvince"), FatherPhone = G("FatherPhone"),
                            MotherName = G("MotherName"), MotherAge = Age("MotherAge"), MotherJob = G("MotherJob"),
                            MotherVillage = G("MotherVillage"), MotherDistrict = G("MotherDistrict"),
                            MotherProvince = G("MotherProvince"), MotherPhone = G("MotherPhone"),
                            GradeLevel = G("GradeLevel"), ClassRoom = G("ClassRoom"), AcademicYear = G("AcademicYear"),
                        };
                        if (DB.ImportStudentRow(dto, conn) > 0) ins++; else skip++;
                    }
                    catch { skip++; }
                }
                DB.Log("ImportCSV", $"{ins} ຄົນ ຂ້າມ {skip} ບໍ່ຄົບ {invalid}");
                MessageBox.Show(
                    $"ນຳເຂົ້າ {ins} ຄົນ\n" +
                    $"ຂ້າມ (ລະຫັດຊ້ຳ) {skip} ຄົນ\n" +
                    $"ບໍ່ຄົບຂໍ້ມູນຈຳເປັນ {invalid} ແຖວ",
                    "ສຳເລັດ", MessageBoxButton.OK, MessageBoxImage.Information);
            };
            st.Children.Add(bImp);
            return new ScrollViewer { Content = st };
        }

        private UIElement BuildBackup(){var st=new StackPanel{Margin=new Thickness(20)};st.Children.Add(new TextBlock{Text=$"ໄຟລ໌: sis_lao.db\nທີ່ຢູ່: {AppDomain.CurrentDomain.BaseDirectory}",Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),TextWrapping=TextWrapping.Wrap,Margin=new Thickness(0,0,0,16)});var bB=B("📦  Backup","PrimaryButton",200);bB.Click+=(s,e)=>{var dlg=new SaveFileDialog{Filter="DB|*.db",FileName=$"backup_{DateTime.Now:yyyyMMdd_HHmm}.db"};if(dlg.ShowDialog()!=true)return;DB.Backup(dlg.FileName);MessageBox.Show("Backup ສຳເລັດ","ສຳເລັດ",MessageBoxButton.OK,MessageBoxImage.Information);};st.Children.Add(bB);st.Children.Add(new TextBlock{Height=10});var bR=B("🔄  Restore","WarningButton",200);bR.Click+=(s,e)=>{if(MessageBox.Show("⚠ Restore ຈະແທນທີ່ຂໍ້ມູນທັງໝົດ?","ຄຳເຕືອນ",MessageBoxButton.YesNo,MessageBoxImage.Warning)!=MessageBoxResult.Yes)return;var dlg=new OpenFileDialog{Filter="DB|*.db"};if(dlg.ShowDialog()!=true)return;DB.Restore(dlg.FileName);DB.LoadSettings();MessageBox.Show("Restore ສຳເລັດ — ກະລຸນາເປີດໂປຣແກຣ ໃໝ່","ສຳເລັດ",MessageBoxButton.OK,MessageBoxImage.Information);Application.Current.Shutdown();System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(System.Diagnostics.Process.GetCurrentProcess().MainModule!.FileName!){UseShellExecute=true});};st.Children.Add(bR);return new ScrollViewer{Content=st};}

        private UIElement BuildLog(){var root=H.MkGrid(GridLength.Auto,new GridLength(1,GridUnitType.Star));var flow=new WrapPanel{Margin=new Thickness(0,0,0,10)};var bR=H.Btn("🔄  ໂຫຼດ","PrimaryButton"); bR.Click+=(s,e)=>ReloadLog();var bC=H.Btn("🗑  ລ້າງ Log (>30 ວັນ)","DangerButton"); bC.Click+=(s,e)=>{if(MessageBox.Show("ລຶບ Log ທີ່ເກົ່າກວ່າ 30 ວັນ? (ກູ້ຄືນບໍ່ໄດ້)","ຢືນຢັນ",MessageBoxButton.YesNo,MessageBoxImage.Warning)!=MessageBoxResult.Yes)return;int del=DB.PruneActivityLog();ReloadLog();MessageBox.Show($"ລ້າງ Log ສຳເລັດ — ລຶບ {del} ລາຍການ","ສຳເລັດ",MessageBoxButton.OK,MessageBoxImage.Information);};flow.Children.Add(bR);flow.Children.Add(bC);Grid.SetRow(flow,0);root.Children.Add(flow);_logDg=new DataGrid{AutoGenerateColumns=true,IsReadOnly=true,CanUserAddRows=false,BorderThickness=new Thickness(0),Background=System.Windows.Media.Brushes.White};Grid.SetRow(_logDg,1);root.Children.Add(_logDg);ReloadLog();return root;}
        private void ReloadLog()=>_logDg.ItemsSource=DB.GetActivityLog().DefaultView;

        private TextBox F(StackPanel p,string lbl,string val=""){L(p,lbl);var tb=new TextBox{Text=val,Margin=new Thickness(0,0,0,12)};p.Children.Add(tb);return tb;}
        private void L(StackPanel p,string t)=>p.Children.Add(new TextBlock{Text=t,FontSize=12,Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128)),Margin=new Thickness(0,0,0,4)});
        private Button B(string t,string st,double w){var b=new Button{Content=t,Height=34,Width=w,Padding=new Thickness(14,0,14,0),Margin=new Thickness(0,0,0,8),HorizontalAlignment=HorizontalAlignment.Left,Cursor=System.Windows.Input.Cursors.Hand};b.SetResourceReference(Button.StyleProperty,st);return b;}
        private static TabItem MkTab(string h,UIElement c)=>new TabItem{Header=h,Content=c};
    }

}
