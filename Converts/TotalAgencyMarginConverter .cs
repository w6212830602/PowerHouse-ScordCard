using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ScoreCard.Models;

namespace ScoreCard.Converts
{
    // 更新轉換器名稱及邏輯，從Commission改為Margin
    public class TotalAgencyMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.AgencyMargin);
                return $"${total:N2}";
            }
            else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
            {
                decimal total = salesData.Sum(x => x.AgencyMargin);
                return $"${total:N2}";
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TotalBuyResellMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.BuyResellMargin);
                return $"${total:N2}";
            }
            else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
            {
                decimal total = salesData.Sum(x => x.BuyResellMargin);
                return $"${total:N2}";
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TotalMarginConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.TotalMargin);
                return $"${total:N2}";
            }
            else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
            {
                decimal total = salesData.Sum(x => x.TotalMargin);
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