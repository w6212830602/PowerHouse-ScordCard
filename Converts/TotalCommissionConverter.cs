using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using ScoreCard.Models;

namespace ScoreCard.Converts
{
    public class TotalAgencyCommissionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.AgencyCommission);
                return $"${total:N2}";
            }
            else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
            {
                decimal total = salesData.Sum(x => x.AgencyCommission);
                return $"${total:N2}";
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TotalBuyResellCommissionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.BuyResellCommission);
                return $"${total:N2}";
            }
            else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
            {
                decimal total = salesData.Sum(x => x.BuyResellCommission);
                return $"${total:N2}";
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TotalCommissionConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.TotalCommission);
                return $"${total:N2}";
            }
            else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
            {
                decimal total = salesData.Sum(x => x.TotalCommission);
                return $"${total:N2}";
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class TotalPOValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ObservableCollection<ProductSalesData> productData)
            {
                decimal total = productData.Sum(x => x.POValue);
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