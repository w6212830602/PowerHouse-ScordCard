using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class DepartmentLobData
    {
        public int Rank { get; set; }
        public string LOB { get; set; }
        public decimal MarginTarget { get; set; }
        public decimal MarginYTD { get; set; }
        public decimal MarginPercentage => MarginTarget > 0 ? MarginYTD / MarginTarget : 0;

    }
}
