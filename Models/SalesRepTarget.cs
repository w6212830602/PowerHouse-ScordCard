// 在 Models 目錄下創建一個新檔案 SalesRepTarget.cs
using CommunityToolkit.Mvvm.ComponentModel;

namespace ScoreCard.Models
{
    public partial class SalesRepTarget : ObservableObject
    {
        [ObservableProperty]
        private string _salesRep;

        [ObservableProperty]
        private int _fiscalYear;

        [ObservableProperty]
        private decimal _annualTarget;

        [ObservableProperty]
        private decimal _q1Target;

        [ObservableProperty]
        private decimal _q2Target;

        [ObservableProperty]
        private decimal _q3Target;

        [ObservableProperty]
        private decimal _q4Target;
    }
}