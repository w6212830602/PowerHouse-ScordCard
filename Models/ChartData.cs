using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class ChartData
    {
        public string Label { get; set; }
        public decimal Target { get; set; }
        public decimal Achievement { get; set; }
    }
}
