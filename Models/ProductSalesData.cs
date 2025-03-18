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

        // 如果您已經更新了類別中的屬性名稱
        public decimal AgencyMargin { get; set; }
        public decimal BuyResellMargin { get; set; }
        public decimal TotalMargin { get; set; }

        // 或者在繼續使用舊屬性的情況下保持兼容性:
        public decimal AgencyCommission { get; set; } // 與 AgencyMargin 相同
        public decimal BuyResellCommission { get; set; } // 與 BuyResellMargin 相同
        public decimal TotalCommission { get; set; }  // 與 TotalMargin 相同

        public decimal POValue { get; set; }
        public decimal PercentageOfTotal { get; set; }
    }
}