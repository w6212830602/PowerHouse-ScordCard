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

        // 屬性名稱已更改為Margin
        public decimal AgencyMargin { get; set; }
        public decimal BuyResellMargin { get; set; }
        public decimal TotalMargin { get; set; }

    }
}