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
    //  SHARED HELPERS
    // ════════════════════════════════════════════════════════════
    internal static class H
    {
    public static Grid MkGrid(params GridLength[] rows)
    {
        var g=new Grid();
        foreach(var r in rows) g.RowDefinitions.Add(new RowDefinition{Height=r});
        return g;
    }

    public static Border MkCard(Thickness? margin=null,Thickness? padding=null) =>
        new Border{
            Background=System.Windows.Media.Brushes.White,
            CornerRadius=new CornerRadius(10),
            Padding=padding??new Thickness(0),
            Margin=margin??new Thickness(0),
            BorderBrush=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
            BorderThickness=new Thickness(1)
        };

    public static TextBlock Lbl(string t) =>
        new TextBlock{Text=t,VerticalAlignment=VerticalAlignment.Center,
            Margin=new Thickness(0,0,5,0),FontSize=13,
            Foreground=new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(107,114,128))};

    public static Button Btn(string text,string style)
    {
        var b=new Button{Content=text,Margin=new Thickness(0,0,6,0),Height=34,
            Padding=new Thickness(14,0,14,0),Cursor=System.Windows.Input.Cursors.Hand};
        b.SetResourceReference(Button.StyleProperty,style);
        return b;
    }

    public static ComboBox MkCmb(string[] items,double w)
    {
        var c=new ComboBox{Width=w,Margin=new Thickness(0,0,8,0)};
        foreach(var i in items) c.Items.Add(i);
        c.SelectedIndex=0;
        return c;
    }

    public static DataGridTextColumn Col(string hdr,string path,double w,bool ro=false) =>
        new DataGridTextColumn{Header=hdr,Binding=new System.Windows.Data.Binding(path),Width=new DataGridLength(w),IsReadOnly=ro};

    public static DataGridTextColumn ColStar(string hdr,string path,bool ro=false) =>
        new DataGridTextColumn{Header=hdr,Binding=new System.Windows.Data.Binding(path),Width=new DataGridLength(1,DataGridLengthUnitType.Star),IsReadOnly=ro};
    } // end class H
}
