using System;
using System.Collections.Generic;

namespace ScoreCard.Models
{
    public class SalesAnalysisData
    {
        public DateTime Date { get; set; }
        public decimal Target { get; set; }
        public decimal Achievement { get; set; }
        public decimal Margin { get; set; }
        public string SalesRep { get; set; }
        public string ProductType { get; set; }
        public string Status { get; set; }
        public string Department { get; set; }
        public string LOB { get; set; }
        public decimal Commission { get; set; }
    }
    public class MonthlyPerformance
    {
        public DateTime Month { get; set; }
        public decimal Target { get; set; }
        public decimal Achievement { get; set; }
        public decimal Margin { get; set; }
        public decimal AchievementPercentage => Target > 0 ? (Achievement / Target) * 100 : 0;
    }


    public class SalesRepPerformance
    {
        public string SalesRep { get; set; }
        public decimal Achievement { get; set; }
        public decimal Commission { get; set; }
        public decimal Target { get; set; }
        public decimal AchievementPercentage => Target > 0 ? (Achievement / Target) * 100 : 0;
        public string AchievementDisplay => $"{AchievementPercentage:N1}%";
    }

    public enum TimeRange
    {
        YTD,
        Q1,
        Q2,
        Q3,
        Q4,
        Custom
    }

    public enum ViewType
    {
        ByProduct,
        ByRep,
        ByDeptLOB
    }
}