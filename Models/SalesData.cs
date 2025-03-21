﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScoreCard.Models
{
    public class SalesData
    {
        public DateTime ReceivedDate { get; set; }
        public string SalesRep { get; set; }
        public string Status { get; set; }
        public string ProductType { get; set; }
        public decimal POValue { get; set; }
        public decimal VertivValue { get; set; }
        public decimal BuyResellValue { get; set; } // 新增: J列
        public decimal AgencyMargin { get; set; }   // 新增: M列
        public decimal TotalCommission { get; set; }
        public decimal CommissionPercentage { get; set; }
        public string Department { get; set; }
        public DateTime? CompletionDate { get; set; } // 新增: Y列的完成日期

        // 財年計算 (8月以後為新的一年)
        public int FiscalYear
        {
            get => ReceivedDate.Month >= 8 ? ReceivedDate.Year + 1 : ReceivedDate.Year;
        }

        // 財年季度計算 
        public int Quarter
        {
            get => ReceivedDate.Month switch
            {
                8 or 9 or 10 => 1,    // Q1
                11 or 12 or 1 => 2,   // Q2
                2 or 3 or 4 => 3,     // Q3
                5 or 6 or 7 => 4,     // Q4
                _ => 0
            };
        }

        // 訂單是否已完成
        public bool IsCompleted => CompletionDate.HasValue;
    }
}