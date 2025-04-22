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
            try
            {
                if (value is ObservableCollection<ProductSalesData> productData)
                {
                    // For product data, use VertivValue/POValue property
                    decimal total = 0;
                    foreach (var item in productData)
                    {
                        total += item.VertivValue; // Use VertivValue directly
                    }
                    return $"${total:N2}";
                }
                else if (value is ObservableCollection<SalesLeaderboardItem> salesData)
                {
                    // For sales rep data, use VertivValue property
                    decimal total = 0;
                    foreach (var item in salesData)
                    {
                        total += item.VertivValue;
                    }
                    return $"${total:N2}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in TotalVertivValueConverter: {ex.Message}");
            }
            return "$0.00";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}