using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ScoreCard.Models
{
    public class SalesAnalysisSummary : INotifyPropertyChanged
    {
        private decimal _totalTarget;
        private decimal _totalAchievement;
        private decimal _totalMargin;
        private decimal _remainingTarget;
        private decimal _achievementPercentage;
        private decimal _marginPercentage;
        private decimal _remainingTargetPercentage;
        private List<SalesRepPerformance> _topPerformers = new();
        private List<MonthlyPerformance> _monthlyData = new();

        // 基本屬性 - 添加完整的屬性實現
        public decimal TotalTarget
        {
            get => _totalTarget;
            set { _totalTarget = value; OnPropertyChanged(); }
        }

        public decimal TotalAchievement
        {
            get => _totalAchievement;
            set { _totalAchievement = value; OnPropertyChanged(); }
        }

        public decimal TotalMargin
        {
            get => _totalMargin;
            set { _totalMargin = value; OnPropertyChanged(); }
        }

        public decimal RemainingTarget
        {
            get => _remainingTarget;
            set { _remainingTarget = value; OnPropertyChanged(); }
        }

        private decimal _remainingToTarget;
        public decimal RemainingToTarget
        {
            get => _remainingToTarget;
            set { _remainingToTarget = value; OnPropertyChanged(); }
        }

        // 百分比屬性
        public decimal AchievementPercentage
        {
            get => _achievementPercentage;
            set { _achievementPercentage = value; OnPropertyChanged(); }
        }

        public decimal MarginPercentage
        {
            get => _marginPercentage;
            set { _marginPercentage = value; OnPropertyChanged(); }
        }

        public decimal RemainingTargetPercentage
        {
            get => _remainingTargetPercentage;
            set { _remainingTargetPercentage = value; OnPropertyChanged(); }
        }

        private decimal _actualRemaining;
        public decimal ActualRemaining
        {
            get => _actualRemaining;
            set { _actualRemaining = value; OnPropertyChanged(); }
        }


        // 格式化的顯示字串
        public string AchievementDisplay => $"{AchievementPercentage:N1}%";
        public string MarginDisplay => $"{MarginPercentage:N1}%";
        public string RemainingTargetDisplay => $"{RemainingTargetPercentage:N1}% to achieve";

        // 排行榜和圖表數據
        public List<SalesRepPerformance> TopPerformers
        {
            get => _topPerformers;
            set { _topPerformers = value; OnPropertyChanged(); }
        }

        public List<MonthlyPerformance> MonthlyData
        {
            get => _monthlyData;
            set { _monthlyData = value; OnPropertyChanged(); }
        }

        // INotifyPropertyChanged 實現
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}