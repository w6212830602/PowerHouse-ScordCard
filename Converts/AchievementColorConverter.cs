using System;
using System.Globalization;
using Microsoft.Maui.Controls;
using Microsoft.Maui.Graphics;

namespace ScoreCard.Converts
{
    public class AchievementColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double achievementPercent && parameter is string threshold)
            {
                if (double.TryParse(threshold, out double thresholdValue))
                {
                    // 如果達成率大於或等於閾值（默認為100%），則顯示綠色，否則顯示紅色
                    return achievementPercent >= thresholdValue ? Color.FromArgb("#10B981") : Color.FromArgb("#EF4444");
                }
            }

            // 默認返回紅色
            return Color.FromArgb("#EF4444");
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}