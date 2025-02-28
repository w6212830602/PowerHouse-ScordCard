using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Converts
{
    public class StringEqualsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string stringValue && parameter is string paramValue)
            {
                var parts = paramValue.Split(',');
                if (parts.Length >= 3 && stringValue == parts[0])
                {
                    return parts[1];
                }
                else if (parts.Length >= 3)
                {
                    return parts[2];
                }
            }
            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}