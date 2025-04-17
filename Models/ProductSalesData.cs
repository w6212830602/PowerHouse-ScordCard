using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class ProductSalesData
    {
        public string ProductType { get; set; } = string.Empty;

        // 基本屬性
        public decimal AgencyMargin { get; set; }
        public decimal BuyResellMargin { get; set; }
        public decimal TotalMargin { get; set; }

        // 向後兼容的屬性
        public decimal AgencyCommission { get => AgencyMargin; set => AgencyMargin = value; }
        public decimal BuyResellCommission { get => BuyResellMargin; set => BuyResellMargin = value; }
        public decimal TotalCommission { get => TotalMargin; set => TotalMargin = value; }

        public decimal POValue { get; set; }
        public decimal PercentageOfTotal { get; set; }

        // 添加此屬性以修復錯誤
        public bool IsInProgress { get; set; }
    }

}