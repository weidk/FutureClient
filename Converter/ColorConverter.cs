using System;
using System.Windows.Data;
using System.Globalization;

namespace FutureClient.Converter
{
    public class ColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                double pos = double.Parse(value.ToString());
                if (pos > 0)
                {
                    return "Red";
                }
                else
                {
                    return "Green";
                }
            }
            else
            {
                return value;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
