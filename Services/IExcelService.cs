using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ScoreCard.Models;

namespace ScoreCard.Services
{
    public interface IExcelService
    {
        Task<(List<SalesData> data, DateTime lastUpdated)> LoadDataAsync(string filePath = Constants.EXCEL_FILE_NAME);
        Task<bool> UpdateDataAsync(string filePath = Constants.EXCEL_FILE_NAME, List<SalesData> data = null);
        Task<bool> MonitorFileChangesAsync(CancellationToken token);
        void ClearCache();

        // 獲取待完成項目的剩餘金額（Y欄為空的記錄的N欄總和）
        decimal GetRemainingAmount();

        // 獲取各季度已完成金額的方法（僅計算Y欄有值的記錄）
        decimal GetQ1Achieved();
        decimal GetQ2Achieved();
        decimal GetQ3Achieved();
        decimal GetQ4Achieved();
        decimal GetTotalAchieved();

        // 新增獲取摘要數據的方法
        List<ProductSalesData> GetProductSalesData();
        List<SalesLeaderboardItem> GetSalesLeaderboardData();
        List<DepartmentLobData> GetDepartmentLobData();

        // 新增獲取銷售代表和LOB列表的方法
        List<string> GetAllSalesReps();
        List<string> GetAllLOBs();

        event EventHandler<DateTime> DataUpdated;
    }
}