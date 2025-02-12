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
        public decimal TotalCommission => AgencyCommission + BuyResellCommission;
    }
}
