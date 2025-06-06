﻿using System;
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

        // 獲取正在進行中的金額（Y欄和N欄為空的H欄總和*0.12）
        decimal GetInProgressAmount();

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