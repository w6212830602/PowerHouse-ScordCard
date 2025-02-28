
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
            if (value is bool boolValue && parameter is string stringParameter)
            {
                var options = stringParameter.Split('|');
                if (options.Length == 2)
                {
                    // 根據目標類型返回適當的值
                    if (targetType == typeof(FontAttributes))
                    {
                        // 特殊處理 FontAttributes 轉換
                        if (boolValue)
                        {
                            return FontAttributes.Bold;
                        }
                        else
                        {
                            return FontAttributes.None;
                        }
                    }

                    return boolValue ? options[0] : options[1];
                }
            }

            // 如果目標類型是 FontAttributes，確保返回有效值
            if (targetType == typeof(FontAttributes))
            {
                return FontAttributes.None;
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
                    return stringValue == options[0];
                }
            }

            // 如果值是 FontAttributes.Bold，表示為 true
            if (value is FontAttributes fontAttributes)
            {
                return fontAttributes == FontAttributes.Bold;
            }

            return false;
        }
    }
}