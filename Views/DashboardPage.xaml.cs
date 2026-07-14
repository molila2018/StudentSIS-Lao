using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using StudentSIS.Data;

namespace StudentSIS.Views
{
    public partial class DashboardPage : UserControl
    {
        private List<(string label, int val, Color color)> _chart = new();

        public DashboardPage() { InitializeComponent(); Loaded += (_, _) => Load(); }

        private void Load()
        {
            TxtTotal.Text  = DB.CountAllStudents().ToString();
            TxtActive.Text = DB.CountActiveStudents().ToString();
            TxtSubj.Text   = DB.CountSubjects().ToString();
            TxtFail.Text   = DB.CountFailingStudents().ToString();

            var gdt = DB.GetStudentCountsByGrade();
            var rows = new List<GradeRow>();
            _chart.Clear();
            Color[] pal = { Color.FromRgb(27,79,138), Color.FromRgb(21,128,61), Color.FromRgb(107,62,160),
                             Color.FromRgb(14,116,144), Color.FromRgb(180,90,0), Color.FromRgb(185,28,28) };
            int ci = 0;
            foreach (DataRow r in gdt.Rows)
            {
                rows.Add(new GradeRow { GradeLevel = r["GradeLevel"].ToString()!, Count = Convert.ToInt32(r["Count"]), Active = Convert.ToInt32(r["Active"]) });
                _chart.Add((r["GradeLevel"].ToString()!, Convert.ToInt32(r["Count"]), pal[ci++ % pal.Length]));
            }
            GradeTable.ItemsSource = rows;

            var adt = DB.GetRecentAnnouncements();
            var anns = new List<AnnRow>();
            foreach (DataRow r in adt.Rows) anns.Add(new AnnRow { Title = r["Title"].ToString()!, CreatedAt = r["CreatedAt"].ToString()! });
            AnnList.ItemsSource = anns;
            DrawChart();
        }

        private void ChartCanvas_SizeChanged(object s, SizeChangedEventArgs e) => DrawChart();

        private void DrawChart()
        {
            ChartCanvas.Children.Clear();
            if (_chart.Count == 0) return;
            double cw = ChartCanvas.ActualWidth, ch = ChartCanvas.ActualHeight;
            if (cw < 10 || ch < 10) return;
            double pL = 32, pB = 28, pT = 14, pR = 10;
            double chartW = cw - pL - pR, chartH = ch - pT - pB;
            int maxV = 1; foreach (var (_, v, _) in _chart) if (v > maxV) maxV = v;
            int n = _chart.Count;
            double barW = Math.Max(10, (chartW - (n + 1) * 8) / n);
            double gap  = (chartW - barW * n) / (n + 1);

            for (int i = 1; i <= 4; i++)
            {
                double y = pT + chartH - chartH * i / 4.0;
                ChartCanvas.Children.Add(new Line { X1=pL, Y1=y, X2=cw-pR, Y2=y,
                    Stroke = new SolidColorBrush(Color.FromArgb(30,0,0,0)), StrokeThickness=1,
                    StrokeDashArray = new DoubleCollection { 4, 4 } });
                var yl = new TextBlock { Text = (maxV*i/4).ToString(), FontSize=10,
                    Foreground = new SolidColorBrush(Color.FromRgb(156,163,175)) };
                Canvas.SetLeft(yl, 2); Canvas.SetTop(yl, y-8); ChartCanvas.Children.Add(yl);
            }
            for (int i = 0; i < n; i++)
            {
                var (label, val, col) = _chart[i];
                double bh = val == 0 ? 2 : chartH * val / maxV;
                double bx = pL + gap + i * (barW + gap), by = pT + chartH - bh;
                var bar = new Rectangle { Width=barW, Height=bh, RadiusX=4, RadiusY=4, Fill=new SolidColorBrush(col) };
                Canvas.SetLeft(bar, bx); Canvas.SetTop(bar, by); ChartCanvas.Children.Add(bar);
                if (val > 0) {
                    var vl = new TextBlock { Text=val.ToString(), FontSize=10, FontWeight=FontWeights.SemiBold,
                        Foreground=new SolidColorBrush(Color.FromRgb(30,41,59)) };
                    Canvas.SetLeft(vl, bx+barW/2-6); Canvas.SetTop(vl, by-17); ChartCanvas.Children.Add(vl);
                }
                var xl = new TextBlock { Text=label, FontSize=10, Foreground=new SolidColorBrush(Color.FromRgb(107,114,128)) };
                Canvas.SetLeft(xl, bx+barW/2-label.Length*3.5); Canvas.SetTop(xl, pT+chartH+5);
                ChartCanvas.Children.Add(xl);
            }
        }
    }

    public class GradeRow { public string GradeLevel{get;set;}=""; public int Count{get;set;} public int Active{get;set;} }
    public class AnnRow   { public string Title{get;set;}="";      public string CreatedAt{get;set;}=""; }
}
