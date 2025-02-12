using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace ScoreCard.Converts
{
    public class InverseBoolToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && parameter is string stringParameter)
            {
                var options = stringParameter.Split('|');
                if (options.Length == 2)
                {
                    return boolValue ? options[1] : options[0];
                }
            }
            return value?.ToString() ?? string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string stringValue && parameter is string stringParameter)
            {
                var options = stringParameter.Split('|');
                if (options.Length == 2)
                {
                    return stringValue == options[1];
                }
            }
            return false;
        }
    }
}