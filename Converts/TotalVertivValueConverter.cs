// 创建一个新文件 Converts/TotalVertivValueConverter.cs
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ScoreCard.Models;

namespace ScoreCard.Converts
{
    public class TotalVertivValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.VertivValue);
                return $"${total:N2}";
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}