using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class ProductSalesData
    {
        public ProductSalesData()
        {
            // 初始化非空屬性
            ProductType = string.Empty;
        }

        public string ProductType { get; set; }

        // 屬性名稱已更改為Margin
        public decimal AgencyMargin { get; set; }
        public decimal BuyResellMargin { get; set; }
        public decimal TotalMargin { get; set; }

        public decimal POValue { get; set; }
        public decimal PercentageOfTotal { get; set; }
    }
}