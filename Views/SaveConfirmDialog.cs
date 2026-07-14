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
    //  SAVE-CONFIRM DIALOG
    //
    //  3-button dirty-state prompt for both score-entry pages (ScoresPage
    //  + MonthlyScoresPage). Replaces the plain YesNo `ຈະຍົກເລີກບໍ?`
    //  MessageBox — the old buttons were English-locale ("Yes"/"No") and
    //  the wording (`ຈະຍົກເລີກບໍ?`) conflated "cancel edits" with "cancel
    //  the filter change".
    //
    //  Three explicit outcomes:
    //    💾 ບັນທຶກ      → save the dirty edits first, then let the caller
    //                     proceed with whatever prompted the dialog
    //    ❌ ບໍ່ບັນທຶກ  → discard edits, proceed
    //    ↩ ຍົກເລີກ     → keep edits, block the caller's action (stay put)
    //
    //  Usage:
    //      var d = SaveConfirmDialog.Ask(Window.GetWindow(this));
    //      switch (d) { case SaveConfirmResult.Save: …; case … }
    // ════════════════════════════════════════════════════════════
    public enum SaveConfirmResult { Save, DontSave, Cancel }

    public class SaveConfirmDialog : Window
    {
        public SaveConfirmResult Decision { get; private set; } = SaveConfirmResult.Cancel;

        public SaveConfirmDialog()
        {
            Title = "ຢືນຢັນ";
            Width = 440; Height = 210;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(244,246,250));
            UseLayoutRounding = true; SnapsToDevicePixels = true;

            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var body = new Border {
                Background = System.Windows.Media.Brushes.White,
                Padding = new Thickness(22, 20, 22, 20)
            };
            var flow = new StackPanel { Orientation = Orientation.Horizontal };
            flow.Children.Add(new TextBlock {
                Text = "❓", FontSize = 34,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 18, 0)
            });
            flow.Children.Add(new TextBlock {
                Text = "ມີຂໍ້ມູນທີ່ຍັງບໍ່ໄດ້ບັນທຶກ\nຕ້ອງການບັນທຶກກ່ອນປ່ຽນຕົວກອງບໍ?",
                FontSize = 14, TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(31,41,55))
            });
            body.Child = flow;
            Grid.SetRow(body, 0); root.Children.Add(body);

            var foot = new Border {
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(229,231,235)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(14, 10, 14, 10)
            };
            var fp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var bSave     = H.Btn("💾  ບັນທຶກ",   "SuccessButton"); bSave.Width     = 110;
            var bDontSave = H.Btn("ບໍ່ບັນທຶກ",    "DangerButton");  bDontSave.Width = 110; bDontSave.Margin = new Thickness(8, 0, 0, 0);
            var bCancel   = H.Btn("ຍົກເລີກ",      "NeutralButton"); bCancel.Width   = 90;  bCancel.Margin   = new Thickness(8, 0, 0, 0); bCancel.IsCancel = true;
            bSave.Click     += (s, e) => { Decision = SaveConfirmResult.Save;     DialogResult = true;  Close(); };
            bDontSave.Click += (s, e) => { Decision = SaveConfirmResult.DontSave; DialogResult = true;  Close(); };
            bCancel.Click   += (s, e) => { Decision = SaveConfirmResult.Cancel;   DialogResult = false; Close(); };
            fp.Children.Add(bSave); fp.Children.Add(bDontSave); fp.Children.Add(bCancel);
            foot.Child = fp;
            Grid.SetRow(foot, 1); root.Children.Add(foot);

            Content = root;
        }

        // Convenience: build + show + return the decision. Owner is optional
        // (falls back to Application.Current.MainWindow) so the dialog centres
        // properly and stays modal to the parent.
        public static SaveConfirmResult Ask(Window? owner)
        {
            var dlg = new SaveConfirmDialog();
            if (owner != null) dlg.Owner = owner;
            else if (Application.Current.MainWindow != null && Application.Current.MainWindow.IsLoaded)
                dlg.Owner = Application.Current.MainWindow;
            dlg.ShowDialog();
            return dlg.Decision;
        }
    }

}
