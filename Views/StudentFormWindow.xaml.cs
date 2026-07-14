using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Windows;
using System.Windows.Controls;
using StudentSIS.Data;

namespace StudentSIS.Views
{
    public partial class StudentFormWindow : Window
    {
        private readonly int _id;
        public StudentFormWindow(int id)
        {
            _id = id;
            InitializeComponent();
            Title = id == 0 ? "➕  ເພີ່ມນັກຮຽນໃໝ່" : "✏️  ແກ້ໄຂຂໍ້ມູນນັກຮຽນ";

            // Academic years come from the system registry (Settings/AcademicYears),
            // so this list always stays in sync with the configured years instead of
            // a hardcoded snapshot. New students default to the current year.
            CmbYear.Items.Clear();
            foreach (var y in DB.AcademicYears()) CmbYear.Items.Add(y);
            SetCbi(CmbYear, DB.CurrentYear);

            if (id > 0) Load();
        }

        private void Load()
        {
            var dt = DB.GetStudent(_id);
            if (dt.Rows.Count == 0) return;
            var r = dt.Rows[0];

            // Personal
            TxtCode.Text        = S(r, "StudentCode");
            TxtFirst.Text       = S(r, "FirstName");
            TxtLast.Text        = S(r, "LastName");
            SetCbi(CmbGender,     S(r, "Gender"));
            TxtBirth.Text       = S(r, "BirthDate");
            // NationalID + student Phone fields removed from the form.

            // Birth place
            TxtBirthVill.Text   = S(r, "BirthVillage");
            TxtBirthCity.Text   = S(r, "BirthDistrict");
            TxtBirthProv.Text   = S(r, "BirthProvince");

            // Current address
            TxtVillage.Text     = S(r, "Village");
            TxtDistrict.Text    = S(r, "District");
            TxtProvince.Text    = S(r, "Province");

            // Father — fallback to legacy ParentName / ParentPhone if FatherName/FatherPhone empty
            string faName = S(r, "FatherName");  if (string.IsNullOrWhiteSpace(faName)) faName = S(r, "ParentName");
            string faTel  = S(r, "FatherPhone"); if (string.IsNullOrWhiteSpace(faTel))  faTel  = S(r, "ParentPhone");
            TxtFaName.Text      = faName;
            TxtFaAge.Text       = S(r, "FatherAge");
            TxtFaJob.Text       = S(r, "FatherJob");
            TxtFaVill.Text      = S(r, "FatherVillage");
            TxtFaCity.Text      = S(r, "FatherDistrict");
            TxtFaProv.Text      = S(r, "FatherProvince");
            TxtFaTel.Text       = faTel;

            // Mother
            TxtMaName.Text      = S(r, "MotherName");
            TxtMaAge.Text       = S(r, "MotherAge");
            TxtMaJob.Text       = S(r, "MotherJob");
            TxtMaVill.Text      = S(r, "MotherVillage");
            TxtMaCity.Text      = S(r, "MotherDistrict");
            TxtMaProv.Text      = S(r, "MotherProvince");
            TxtMaTel.Text       = S(r, "MotherPhone");

            // Academic
            SetCbi(CmbGrade,      S(r, "GradeLevel"));
            SetCbi(CmbRoom,       S(r, "ClassRoom"));
            SetCbi(CmbYear,       S(r, "AcademicYear"));
            SetCbi(CmbStatus,     S(r, "Status"));
            TxtNote.Text        = S(r, "Note");
        }

        private void BtnSave_Click(object s, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtCode.Text)
             || string.IsNullOrWhiteSpace(TxtFirst.Text)
             || string.IsNullOrWhiteSpace(TxtLast.Text))
            { MessageBox.Show("ກະລຸນາໃສ່ ລະຫັດ, ຊື່, ນາມສະກຸນ", "ຂໍ້ມູນບໍ່ຄົບ"); return; }

            try
            {
                // Map form controls → DTO; DB.SaveStudent owns the SQL.
                // NationalID + Phone columns were dropped from the form — the
                // columns still exist in the schema but are no longer written.
                var dto = new StudentDto
                {
                    Code       = TxtCode.Text.Trim(),
                    FirstName  = TxtFirst.Text.Trim(),
                    LastName   = TxtLast.Text.Trim(),
                    Gender     = Gcb(CmbGender),
                    BirthDate  = TxtBirth.Text.Trim(),

                    BirthVillage  = TxtBirthVill.Text.Trim(),
                    BirthDistrict = TxtBirthCity.Text.Trim(),
                    BirthProvince = TxtBirthProv.Text.Trim(),

                    Village  = TxtVillage.Text.Trim(),
                    District = TxtDistrict.Text.Trim(),
                    Province = TxtProvince.Text.Trim(),

                    FatherName     = TxtFaName.Text.Trim(),
                    FatherAge      = ParseAge(TxtFaAge.Text),
                    FatherJob      = TxtFaJob.Text.Trim(),
                    FatherVillage  = TxtFaVill.Text.Trim(),
                    FatherDistrict = TxtFaCity.Text.Trim(),
                    FatherProvince = TxtFaProv.Text.Trim(),
                    FatherPhone    = TxtFaTel.Text.Trim(),

                    MotherName     = TxtMaName.Text.Trim(),
                    MotherAge      = ParseAge(TxtMaAge.Text),
                    MotherJob      = TxtMaJob.Text.Trim(),
                    MotherVillage  = TxtMaVill.Text.Trim(),
                    MotherDistrict = TxtMaCity.Text.Trim(),
                    MotherProvince = TxtMaProv.Text.Trim(),
                    MotherPhone    = TxtMaTel.Text.Trim(),

                    GradeLevel   = Gcb(CmbGrade),
                    ClassRoom    = Gcb(CmbRoom),
                    AcademicYear = Gcb(CmbYear),
                    Status       = Gcb(CmbStatus),
                    Note         = TxtNote.Text.Trim(),
                };
                DB.SaveStudent(_id, dto);
                DB.Log(_id == 0 ? "AddStudent" : "EditStudent", TxtCode.Text);
                DialogResult = true; Close();
            }
            catch (SQLiteException ex) when (ex.Message.Contains("UNIQUE"))
            { MessageBox.Show("ລະຫັດນັກຮຽນນີ້ມີຢູ່ແລ້ວ"); }
        }

        // ── Helpers ──────────────────────────────────────────────────
        private static string S(System.Data.DataRow r, string col)
        {
            if (!r.Table.Columns.Contains(col)) return "";
            var v = r[col];
            return v == DBNull.Value ? "" : v.ToString() ?? "";
        }

        // Age stored as INTEGER; empty / non-numeric → DBNull
        private static int? ParseAge(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            return int.TryParse(s.Trim(), out int n) ? n : (int?)null;
        }

        private static void SetCbi(ComboBox cb, string val)
        {
            foreach (var item in cb.Items)
            {
                string? c = item is ComboBoxItem cbi ? cbi.Content?.ToString() : item?.ToString();
                if (c == val) { cb.SelectedItem = item; return; }
            }
            cb.Text = val;
        }
        private static string Gcb(ComboBox cb)
            => cb.SelectedItem is ComboBoxItem cbi ? cbi.Content?.ToString() ?? "" : cb.SelectedItem?.ToString() ?? cb.Text ?? "";
    }
}
