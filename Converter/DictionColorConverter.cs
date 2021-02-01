using System;
using System.Windows.Data;
using System.Globalization;

namespace FutureClient.Converter
{
    public class DictionColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                string diction = value.ToString();
                if (diction == "1"|| diction == "多")
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
