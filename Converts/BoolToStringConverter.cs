
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Maui.Controls;

namespace ScoreCard.Converts
{
    public class BoolToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (targetType == typeof(FontAttributes))
            {
                // 直接返回FontAttributes枚举值
                if (value is bool boolValue1) // 变量名改为boolValue1
                {
                    return boolValue1 ? FontAttributes.Bold : FontAttributes.None;
                }
                return FontAttributes.None;
            }

            if (value is bool boolValue2 && parameter is string stringParameter) // 变量名改为boolValue2
            {
                var options = stringParameter.Split('|');
                if (options.Length == 2)
                {
                    return boolValue2 ? options[0] : options[1];
                }
            }

            return value?.ToString() ?? string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is FontAttributes fontAttributes)
            {
                return fontAttributes == FontAttributes.Bold;
            }

            if (value is string stringValue && parameter is string stringParameter)
            {
                var options = stringParameter.Split('|');
                if (options.Length == 2)
                {
                    return stringValue == options[0];
                }
            }

            return false;
        }
    }
}