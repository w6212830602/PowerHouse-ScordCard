using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class ProductSalesData
    {
        public string ProductType { get; set; }
        public decimal AgencyCommission { get; set; }
        public decimal BuyResellCommission { get; set; }
        public decimal TotalCommission => AgencyCommission + BuyResellCommission;
        public decimal POValue { get; set; }
        public decimal PercentageOfTotal { get; set; }
    }
}
