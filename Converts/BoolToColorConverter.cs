using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Converts
{
    public class BoolToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && parameter is string colorParameter)
            {
                var colors = colorParameter.Split(',');
                if (colors.Length >= 2)
                {
                    return boolValue ? Color.FromArgb(colors[0]) : Color.FromArgb(colors[1]);
                }
            }
            return Colors.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class InverseBoolToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && parameter is string colorParameter)
            {
                var colors = colorParameter.Split(',');
                if (colors.Length >= 2)
                {
                    return !boolValue ? Color.FromArgb(colors[0]) : Color.FromArgb(colors[1]);
                }
            }
            return Colors.Transparent;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}