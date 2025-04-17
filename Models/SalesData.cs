// 修改 Models/SalesData.cs

public class SalesData
{
    // Initialize all required string properties to empty string to avoid null
    public SalesData()
    {
        // Initialize string properties to prevent null references
        SalesRep = string.Empty;
        Status = string.Empty;
        ProductType = string.Empty;
        Department = string.Empty;
    }

    public DateTime ReceivedDate { get; set; }
    public DateTime? CompletionDate { get; set; }
    public string SalesRep { get; set; }
    public string Status { get; set; }
    public string ProductType { get; set; }
    public decimal POValue { get; set; }
    public decimal VertivValue { get; set; } // 已經存在的 VertivValue 屬性
    public decimal BuyResellValue { get; set; }
    public decimal AgencyMargin { get; set; }
    public decimal TotalCommission { get; set; }
    public decimal CommissionPercentage { get; set; }
    public string Department { get; set; }

    // Flag for in-progress status
    public bool IsInProgress { get; set; }

    // Flag for remaining items
    public bool IsRemaining { get; set; }

    // Has quarter assigned
    public bool HasQuarterAssigned { get; set; }

    // Quarter date for calculation
    public DateTime QuarterDate { get; set; }

    // Fiscal year calculation
    public int FiscalYear
    {
        get => HasQuarterAssigned ? (QuarterDate.Month >= 8 ? QuarterDate.Year + 1 : QuarterDate.Year) : 0;
    }

    // Quarter calculation
    public int Quarter
    {
        get => HasQuarterAssigned ? QuarterDate.Month switch
        {
            8 or 9 or 10 => 1,    // Q1
            11 or 12 or 1 => 2,   // Q2
            2 or 3 or 4 => 3,     // Q3
            5 or 6 or 7 => 4,     // Q4
            _ => 0
        } : 0;
    }

    // Helper property to determine overall status
    public string StatusType
    {
        get
        {
            if (CompletionDate.HasValue)
                return "Completed";
            else if (IsInProgress)
                return "InProgress";
            else
                return "Booked";
        }
    }
}