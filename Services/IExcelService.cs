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

        // 新增獲取摘要數據的方法
        List<ProductSalesData> GetProductSalesData();
        List<SalesLeaderboardItem> GetSalesLeaderboardData();
        List<DepartmentLobData> GetDepartmentLobData();

        event EventHandler<DateTime> DataUpdated;
    }
}