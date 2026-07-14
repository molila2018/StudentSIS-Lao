using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using StudentSIS.Data;

namespace StudentSIS.Views
{
    public partial class StudentsPage : UserControl
    {
        private List<StudentRow> _rows = new();
        // Cached on first load + after add/delete; avoids re-running COUNT(*) on every keystroke.
        private int _totalCache = -1;

        public StudentsPage()
        {
            InitializeComponent();
            InitFilters();
            // NavContext handoff from ClassHubPage: pre-select grade + room filters
            // so the StudentsPage opens already filtered to the chosen class.
            if (!string.IsNullOrEmpty(DB.NavGrade))
            {
                for (int i = 0; i < CmbGrade.Items.Count; i++)
                    if (CmbGrade.Items[i]?.ToString() == DB.NavGrade) { CmbGrade.SelectedIndex = i; break; }
                if (!string.IsNullOrEmpty(DB.NavRoom))
                    for (int i = 0; i < CmbRoom.Items.Count; i++)
                        if (CmbRoom.Items[i]?.ToString() == DB.NavRoom) { CmbRoom.SelectedIndex = i; break; }
                DB.ClearNav();
            }
            Load();
        }

        private void InitFilters()
        {
            CmbGrade.Items.Add("ທັງໝົດ");
            foreach (var g in new[]{"ມ.1","ມ.2","ມ.3","ມ.4"}) CmbGrade.Items.Add(g);
            CmbGrade.SelectedIndex = 0;

            CmbRoom.Items.Add("ທັງໝົດ");
            foreach (var r in new[]{"1","2","3","4","5"}) CmbRoom.Items.Add(r);
            CmbRoom.SelectedIndex = 0;

            CmbYear.Items.Add("ທັງໝົດ");
            foreach (var y in DB.AcademicYears()) CmbYear.Items.Add(y);
            CmbYear.SelectedIndex = 0;

            CmbStatus.Items.Add("ທັງໝົດ");
            foreach (var s in new[]{"ກຳລັງຮຽນ","ຈົບ","ອອກ"}) CmbStatus.Items.Add(s);
            CmbStatus.SelectedIndex = 0;
        }

        private void Load()
        {
            string grade  = CmbGrade?.SelectedItem?.ToString()  ?? "ທັງໝົດ";
            string room   = CmbRoom?.SelectedItem?.ToString()   ?? "ທັງໝົດ";
            string year   = CmbYear?.SelectedItem?.ToString()   ?? "ທັງໝົດ";
            string status = CmbStatus?.SelectedItem?.ToString() ?? "ທັງໝົດ";
            string kw     = TxtSearch?.Text.Trim() ?? "";

            // "ທັງໝົດ" in a combo means no filter → pass null to the query.
            var dt = DB.SearchStudents(kw,
                grade  == "ທັງໝົດ" ? null : grade,
                room   == "ທັງໝົດ" ? null : room,
                year   == "ທັງໝົດ" ? null : year,
                status == "ທັງໝົດ" ? null : status);
            _rows = new List<StudentRow>();
            foreach (DataRow r in dt.Rows)
                _rows.Add(new StudentRow {
                    StudentID   = Convert.ToInt32(r["StudentID"]),
                    StudentCode = r["StudentCode"].ToString()!,
                    FullName    = r["FullName"].ToString()!,
                    Gender      = r["Gender"].ToString()!,
                    GradeLevel  = r["GradeLevel"].ToString()!,
                    ClassRoom   = r["ClassRoom"].ToString()!,
                    AcademicYear= r["AcademicYear"].ToString()!,
                    ParentName  = r["ParentName"].ToString()!,
                    Status      = r["Status"].ToString()!
                });

            StudentGrid.ItemsSource = _rows;
            int cnt = _rows.Count;
            if (_totalCache < 0) _totalCache = DB.CountAllStudents();
            if (LblCount  != null) LblCount.Text  = $"ພົບ {cnt} ຄົນ";
            if (LblStatus != null) LblStatus.Text = $"ນັກຮຽນທັງໝົດ {_totalCache} ຄົນ  |  ສະແດງ {cnt} ຄົນ";
        }

        private int SelId() => StudentGrid.SelectedItem is StudentRow r ? r.StudentID : 0;

        private void BtnAdd_Click(object s, RoutedEventArgs e)
        { var f = new StudentFormWindow(0) { Owner = Window.GetWindow(this) }; if (f.ShowDialog()==true) { _totalCache = -1; Load(); } }

        private void BtnEdit_Click(object s, RoutedEventArgs e)
        { int id=SelId(); if(id<=0){MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນກ່ອນ");return;} var f=new StudentFormWindow(id){Owner=Window.GetWindow(this)}; if(f.ShowDialog()==true)Load(); }

        private void BtnDelete_Click(object s, RoutedEventArgs e)
        {
            int id = SelId();
            if (id <= 0)
            {
                MessageBox.Show("ກະລຸນາເລືອກນັກຮຽນທີ່ຕ້ອງການລຶບກ່ອນ", "ຍັງບໍ່ໄດ້ເລືອກ",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Look up the student label + count dependent records that will be cascade-deleted.
            string label = DB.GetStudentLabel(id);
            var (enrollCount, scoreCount, monthCount, historyCount) = DB.GetStudentCascadeCounts(id);

            string warn = $"ຕ້ອງການລຶບນັກຮຽນຄົນນີ້ບໍ?\n\n  •  {label}\n";
            if (enrollCount + scoreCount + monthCount + historyCount > 0)
            {
                warn += $"\n⚠ ການລຶບຈະຕັດຂໍ້ມູນຕໍ່ໄປນີ້ດ້ວຍ:\n" +
                        $"   • ການລົງທະບຽນ:          {enrollCount}\n" +
                        $"   • ຄະແນນພາກຮຽນ:        {scoreCount}\n" +
                        $"   • ຄະແນນລາຍເດືອນ:      {monthCount}\n" +
                        $"   • ປະຫວັດການຂຶ້ນຊັ້ນ:  {historyCount}\n\n" +
                        $"ການກະທຳນີ້ບໍ່ສາມາດຍົກເລີກໄດ້.";
            }

            if (MessageBox.Show(warn, "ຢືນຢັນການລຶບ", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                DB.DeleteStudent(id);
                DB.Log("DelStudent", label);
                _totalCache = -1;
                Load();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ລຶບບໍ່ສຳເລັດ: {ex.Message}", "ຜິດພາດ", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Filter_Changed(object s, EventArgs e) => Load();
        private void Grid_DoubleClick(object s, MouseButtonEventArgs e)
        { int id=SelId(); if(id<=0)return; var f=new StudentFormWindow(id){Owner=Window.GetWindow(this)}; if(f.ShowDialog()==true)Load(); }
    }

    public class StudentRow
    {
        public int    StudentID   {get;set;}
        public string StudentCode {get;set;}="";
        public string FullName    {get;set;}="";
        public string Gender      {get;set;}="";
        public string GradeLevel  {get;set;}="";
        public string ClassRoom   {get;set;}="";
        public string AcademicYear{get;set;}="";
        public string ParentName  {get;set;}="";
        public string Status      {get;set;}="";
    }
}
