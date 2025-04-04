// 修改 Models/SalesData.cs

public class SalesData
{
    public DateTime ReceivedDate { get; set; }
    public DateTime? CompletionDate { get; set; }
    public string SalesRep { get; set; }
    public string Status { get; set; }
    public string ProductType { get; set; }
    public decimal POValue { get; set; }
    public decimal VertivValue { get; set; }
    public decimal BuyResellValue { get; set; } // J列
    public decimal AgencyMargin { get; set; }   // M列
    public decimal TotalCommission { get; set; } // N列
    public decimal CommissionPercentage { get; set; }
    public string Department { get; set; }

    // 添加新屬性，表示這是一個未完成項目（Y列為空）
    public bool IsRemaining { get; set; }

    // 添加新屬性，表示這個記錄是否應該計入季度業績
    public bool HasQuarterAssigned { get; set; }

    // 季度計算日期（僅對於已完成的訂單）
    public DateTime QuarterDate { get; set; }

    // 財年計算 (8月以後為新的一年)
    public int FiscalYear
    {
        get => HasQuarterAssigned ? (QuarterDate.Month >= 8 ? QuarterDate.Year + 1 : QuarterDate.Year) : 0;
    }

    // 財年季度計算 
    public int Quarter
    {
        get => HasQuarterAssigned ? QuarterDate.Month switch
        {
            8 or 9 or 10 => 1,    // Q1
            11 or 12 or 1 => 2,   // Q2
            2 or 3 or 4 => 3,     // Q3
            5 or 6 or 7 => 4,     // Q4
            _ => 0
        } : 0; // 未分配季度的訂單返回0
    }

    // 訂單是否已完成 - 使用 CompletionDate 來確定
    public bool IsCompleted => CompletionDate.HasValue;
}
