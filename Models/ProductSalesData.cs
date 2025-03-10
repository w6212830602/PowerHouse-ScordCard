﻿using System;
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
        public decimal AgencyCommission { get; set; }
        public decimal BuyResellCommission { get; set; }
        public decimal TotalCommission { get; set; }
        public decimal POValue { get; set; }
        public decimal PercentageOfTotal { get; set; }
    }
}
