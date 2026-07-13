using System; using System.Globalization; using System.Windows; using System.Windows.Data; using System.Windows.Media;
namespace StudentSIS.Helpers
{
    public class LevelToFgConverter : IValueConverter {
        public object Convert(object v,Type t,object p,CultureInfo c)=>(v?.ToString()??"") switch{
            "ດີຫຼາຍ"=>new SolidColorBrush(Color.FromRgb(21,128,61)),"ດີ"=>new SolidColorBrush(Color.FromRgb(27,79,138)),
            "ຜ່ານ"=>new SolidColorBrush(Color.FromRgb(146,64,14)),"ບໍ່ຜ່ານ"=>new SolidColorBrush(Color.FromRgb(185,28,28)),
            _=>new SolidColorBrush(Color.FromRgb(107,114,128))};
        public object ConvertBack(object v,Type t,object p,CultureInfo c)=>Binding.DoNothing;}

    public class LevelToBgConverter : IValueConverter {
        public object Convert(object v,Type t,object p,CultureInfo c)=>(v?.ToString()??"") switch{
            "ດີຫຼາຍ"=>new SolidColorBrush(Color.FromRgb(220,252,231)),"ດີ"=>new SolidColorBrush(Color.FromRgb(214,228,247)),
            "ຜ່ານ"=>new SolidColorBrush(Color.FromRgb(254,243,199)),"ບໍ່ຜ່ານ"=>new SolidColorBrush(Color.FromRgb(254,226,226)),
            _=>new SolidColorBrush(Color.FromRgb(243,244,246))};
        public object ConvertBack(object v,Type t,object p,CultureInfo c)=>Binding.DoNothing;}

    public class StatusToBgConverter : IValueConverter {
        public object Convert(object v,Type t,object p,CultureInfo c)=>(v?.ToString()??"") switch{
            "ກຳລັງຮຽນ"=>new SolidColorBrush(Color.FromRgb(220,252,231)),"ຈົບ"=>new SolidColorBrush(Color.FromRgb(214,228,247)),
            "ອອກ"=>new SolidColorBrush(Color.FromRgb(254,226,226)),_=>new SolidColorBrush(Color.FromRgb(243,244,246))};
        public object ConvertBack(object v,Type t,object p,CultureInfo c)=>Binding.DoNothing;}

    public class StatusToFgConverter : IValueConverter {
        public object Convert(object v,Type t,object p,CultureInfo c)=>(v?.ToString()??"") switch{
            "ກຳລັງຮຽນ"=>new SolidColorBrush(Color.FromRgb(21,128,61)),"ຈົບ"=>new SolidColorBrush(Color.FromRgb(27,79,138)),
            "ອອກ"=>new SolidColorBrush(Color.FromRgb(185,28,28)),_=>new SolidColorBrush(Color.FromRgb(107,114,128))};
        public object ConvertBack(object v,Type t,object p,CultureInfo c)=>Binding.DoNothing;}
}
