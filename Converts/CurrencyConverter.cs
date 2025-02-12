using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Converts
{
    public class CurrencyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is decimal amount)
            {
                if (Math.Abs(amount) >= 1000000)
                {
                    return $"${amount / 1000000:N1}M";
                }
                if (Math.Abs(amount) >= 1000)
                {
                    return $"${amount / 1000:N1}K";
                }
                return $"${amount:N0}";
            }
            return "$0";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
