// 创建/修改文件 Converts/TotalPOValueConverter.cs
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ScoreCard.Models;

namespace ScoreCard.Converts
{
    public class TotalPOValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.VertivValue); // 使用VertivValue替代POValue
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