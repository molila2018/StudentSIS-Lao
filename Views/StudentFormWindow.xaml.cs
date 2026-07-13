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
            var dt = DB.Query("SELECT * FROM Students WHERE StudentID=@id", null, ("@id", _id));
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
                // NationalID + Phone columns dropped from the form. The columns still
                // exist in the schema for backward compatibility but no longer written here.
                string sql = _id == 0
                    ? @"INSERT INTO Students(
                          StudentCode, FirstName, LastName, Gender, BirthDate,
                          BirthVillage, BirthDistrict, BirthProvince,
                          Village, District, Province,
                          FatherName, FatherAge, FatherJob, FatherVillage, FatherDistrict, FatherProvince, FatherPhone,
                          MotherName, MotherAge, MotherJob, MotherVillage, MotherDistrict, MotherProvince, MotherPhone,
                          ParentName, ParentPhone,
                          GradeLevel, ClassRoom, AcademicYear, Status, Note)
                       VALUES(
                          @code, @fn, @ln, @gd, @bd,
                          @bvi, @bdi, @bpv,
                          @vi, @di, @pv,
                          @faN, @faA, @faJ, @faV, @faC, @faP, @faT,
                          @maN, @maA, @maJ, @maV, @maC, @maP, @maT,
                          @pn, @pp,
                          @gl, @rm, @yr, @st, @nt)"
                    : @"UPDATE Students SET
                          StudentCode=@code, FirstName=@fn, LastName=@ln, Gender=@gd,
                          BirthDate=@bd,
                          BirthVillage=@bvi, BirthDistrict=@bdi, BirthProvince=@bpv,
                          Village=@vi, District=@di, Province=@pv,
                          FatherName=@faN, FatherAge=@faA, FatherJob=@faJ,
                          FatherVillage=@faV, FatherDistrict=@faC, FatherProvince=@faP, FatherPhone=@faT,
                          MotherName=@maN, MotherAge=@maA, MotherJob=@maJ,
                          MotherVillage=@maV, MotherDistrict=@maC, MotherProvince=@maP, MotherPhone=@maT,
                          ParentName=@pn, ParentPhone=@pp,
                          GradeLevel=@gl, ClassRoom=@rm, AcademicYear=@yr, Status=@st, Note=@nt
                       WHERE StudentID=@id";

                var ps = new List<(string, object)>
                {
                    ("@code", TxtCode.Text.Trim()),
                    ("@fn",   TxtFirst.Text.Trim()),
                    ("@ln",   TxtLast.Text.Trim()),
                    ("@gd",   Gcb(CmbGender)),
                    ("@bd",   TxtBirth.Text.Trim()),
                    // Birth place
                    ("@bvi",  TxtBirthVill.Text.Trim()),
                    ("@bdi",  TxtBirthCity.Text.Trim()),
                    ("@bpv",  TxtBirthProv.Text.Trim()),
                    // Current address
                    ("@vi",   TxtVillage.Text.Trim()),
                    ("@di",   TxtDistrict.Text.Trim()),
                    ("@pv",   TxtProvince.Text.Trim()),
                    // Father
                    ("@faN",  TxtFaName.Text.Trim()),
                    ("@faA",  ParseAge(TxtFaAge.Text)),
                    ("@faJ",  TxtFaJob.Text.Trim()),
                    ("@faV",  TxtFaVill.Text.Trim()),
                    ("@faC",  TxtFaCity.Text.Trim()),
                    ("@faP",  TxtFaProv.Text.Trim()),
                    ("@faT",  TxtFaTel.Text.Trim()),
                    // Mother
                    ("@maN",  TxtMaName.Text.Trim()),
                    ("@maA",  ParseAge(TxtMaAge.Text)),
                    ("@maJ",  TxtMaJob.Text.Trim()),
                    ("@maV",  TxtMaVill.Text.Trim()),
                    ("@maC",  TxtMaCity.Text.Trim()),
                    ("@maP",  TxtMaProv.Text.Trim()),
                    ("@maT",  TxtMaTel.Text.Trim()),
                    // Legacy mirrors (so old queries / CSV exports keep showing a parent)
                    ("@pn",   TxtFaName.Text.Trim()),
                    ("@pp",   TxtFaTel.Text.Trim()),
                    // Academic
                    ("@gl",   Gcb(CmbGrade)),
                    ("@rm",   Gcb(CmbRoom)),
                    ("@yr",   Gcb(CmbYear)),
                    ("@st",   Gcb(CmbStatus)),
                    ("@nt",   TxtNote.Text.Trim()),
                };
                if (_id > 0) ps.Add(("@id", _id));
                DB.Exec(sql, null, ps.ToArray());
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
        private static object ParseAge(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return DBNull.Value;
            return int.TryParse(s.Trim(), out int n) ? (object)n : DBNull.Value;
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
