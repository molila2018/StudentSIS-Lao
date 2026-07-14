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

    public class ComboItem{public string Tag{get;}public string Display{get;}public ComboItem(string t,string d){Tag=t;Display=d;}}
    public class SubComboItem{
        public int Id{get;}
        public string Code{get;}                  // SubjectCode — used to detect CHA1/LAB1 without a DB query
        private readonly string _d;
        public SubComboItem(int id,string d){ Id=id; Code=""; _d=d; }
        public SubComboItem(int id,string code,string d){ Id=id; Code=code; _d=d; }
        public override string ToString()=>_d;
    }

}
