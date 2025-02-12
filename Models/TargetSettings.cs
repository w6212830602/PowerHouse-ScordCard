using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class FiscalYearTarget
    {
        public int FiscalYear { get; set; }
        public decimal AnnualTarget { get; set; }
        public decimal Q1Target { get; set; }
        public decimal Q2Target { get; set; }
        public decimal Q3Target { get; set; }
        public decimal Q4Target { get; set; }
    }

    public class TargetSettings
    {
        public List<FiscalYearTarget> CompanyTargets { get; set; }
    }

}
