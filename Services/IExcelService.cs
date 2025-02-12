using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading.Tasks;
using ScoreCard.Services;
using ScoreCard.Models;
using ScoreCard.Views;



namespace ScoreCard.Services
{
    public interface IExcelService
    {
        Task<(List<SalesData> data, DateTime lastUpdated)> LoadDataAsync(string filePath = Constants.EXCEL_FILE_NAME);
        Task<bool> UpdateDataAsync(string filePath = Constants.EXCEL_FILE_NAME, List<SalesData> data = null);
        Task<bool> MonitorFileChangesAsync(CancellationToken token);  // 簡化參數
        event EventHandler<DateTime> DataUpdated;
    }
}

