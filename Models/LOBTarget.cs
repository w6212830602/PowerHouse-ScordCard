using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace ScoreCard.Models
{
    public partial class LOBTarget : ObservableObject
    {
        private string _lob;
        private int _fiscalYear;
        private decimal _annualTarget;
        private decimal _q1Target;
        private decimal _q2Target;
        private decimal _q3Target;
        private decimal _q4Target;

        // 手動添加屬性，不使用ObservableProperty特性
        public string LOB
        {
            get => _lob;
            set => SetProperty(ref _lob, value);
        }

        public int FiscalYear
        {
            get => _fiscalYear;
            set => SetProperty(ref _fiscalYear, value);
        }

        public decimal AnnualTarget
        {
            get => _annualTarget;
            set => SetProperty(ref _annualTarget, value);
        }

        public decimal Q1Target
        {
            get => _q1Target;
            set => SetProperty(ref _q1Target, value);
        }

        public decimal Q2Target
        {
            get => _q2Target;
            set => SetProperty(ref _q2Target, value);
        }

        public decimal Q3Target
        {
            get => _q3Target;
            set => SetProperty(ref _q3Target, value);
        }

        public decimal Q4Target
        {
            get => _q4Target;
            set => SetProperty(ref _q4Target, value);
        }
    }
}