using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class SalesLeaderboardItem
    {
        public int Rank { get; set; }
        public string SalesRep { get; set; }
        public decimal AgencyCommission { get; set; }
        public decimal BuyResellCommission { get; set; }

        // 將 TotalCommission 從計算屬性改為一般可讀寫屬性
        public decimal TotalCommission { get; set; }
    }
}
