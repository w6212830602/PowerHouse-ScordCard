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

        // 保持原始屬性名稱以避免錯誤
        public decimal AgencyCommission { get; set; }
        public decimal BuyResellCommission { get; set; }
        public decimal TotalCommission { get; set; }
    }

}