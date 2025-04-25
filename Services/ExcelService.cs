using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ScoreCard.Models;
using OfficeOpenXml;
using System.Diagnostics;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace ScoreCard.Services
{
    public class ExcelService : IExcelService
    {
        private readonly string _worksheetName = Constants.WORKSHEET_NAME;
        private readonly string _leaderboardSheetName = "Sales Leaderboard";
        private readonly string _byDeptLobSheetName = "By Dept-LOB";
        private readonly string _byRepSheetName = "By Rep";
        private FileSystemWatcher _watcher;
        public event EventHandler<DateTime> DataUpdated;
        private List<SalesData> _allSalesData = new List<SalesData>();

        private List<ProductSalesData> _productSalesCache = new List<ProductSalesData>();
        private List<SalesLeaderboardItem> _salesLeaderboardCache = new List<SalesLeaderboardItem>();
        private List<DepartmentLobData> _departmentLobCache = new List<DepartmentLobData>();
        private List<SalesData> _recentDataCache = new List<SalesData>();

        // 保存剩餘金額的靜態變量
        private static decimal _remainingAmount = 0;

        // 保存In Progress金額（記錄Y欄和N欄均為空的H欄總和的12%）
        private static decimal _inProgressAmount = 0;

        // 添加一個方法來獲取剩餘金額
        public decimal GetRemainingAmount()
        {
            return _remainingAmount;
        }

        // 添加一個方法來獲取正在進行中的金額（Y欄和N欄為空的H欄總和*0.12）
        public decimal GetInProgressAmount()
        {
            return _inProgressAmount;
        }

        // In ExcelService.cs, modify the LoadDataAsync method to include the completion date
        public async Task<(List<SalesData> data, DateTime lastUpdated)> LoadDataAsync(string filePath = Constants.EXCEL_FILE_NAME)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string fullPath = Path.Combine(Constants.BASE_PATH, filePath);
                    Debug.WriteLine($"尝试加载 Excel 文件: {fullPath}");

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var data = new List<SalesData>();
                    DateTime lastUpdated = File.Exists(fullPath) ? File.GetLastWriteTime(fullPath) : DateTime.Now;

                    using (var package = new ExcelPackage(new FileInfo(fullPath)))
                    {
                        // 1. 读取原始数据
                        var worksheet = package.Workbook.Worksheets[_worksheetName];
                        if (worksheet == null)
                        {
                            Debug.WriteLine($"找不到工作表: '{_worksheetName}'");

                            // 尝试使用索引查找工作表
                            if (package.Workbook.Worksheets.Count > 0)
                            {
                                worksheet = package.Workbook.Worksheets[0];
                                Debug.WriteLine($"改用第一个工作表: {worksheet.Name}");
                            }
                            else
                            {
                                throw new Exception("Excel 文件中没有工作表");
                            }
                        }

                        Debug.WriteLine($"成功找到工作表: {worksheet.Name}");
                        var rowCount = worksheet.Dimension?.Rows ?? 0;
                        Debug.WriteLine($"工作表有 {rowCount} 行数据");

                        if (rowCount > 0)
                        {
                            // 绘制表头，确认列名称
                            var headers = new List<string>();
                            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                            {
                                var headerValue = worksheet.Cells[1, col].Text;
                                headers.Add(headerValue);
                            }
                            Debug.WriteLine($"表头: {string.Join(", ", headers)}");

                            // 保存剩余金额的计算，计算具有未完成日期的订单的 TotalCommission 总和
                            decimal remainingAmount = 0;

                            // 保存正在进行中的金额，计算Y列和N列均为空的记录的H列总和的12%
                            decimal inProgressAmount = 0;

                            // 添加计数器和记录已处理行的集合
                            int inProgressCount = 0;
                            decimal totalHValue = 0;
                            decimal calculatedInProgressAmount = 0; // 本地变量，不使用静态变量进行中间计算
                            HashSet<int> processedRows = new HashSet<int>();

                            Debug.WriteLine("===== 开始计算进行中金额 =====");

                            for (int row = 2; row <= rowCount; row++)
                            {
                                // 檢查是否已處理過該行
                                if (!processedRows.Contains(row))
                                {
                                    try
                                    {
                                        // 讀取接收日期（A列）
                                        var receivedDateCell = worksheet.Cells[row, 1];
                                        DateTime receivedDate;
                                        if (receivedDateCell.Value is DateTime date)
                                        {
                                            receivedDate = date;
                                        }
                                        else
                                        {
                                            // 嘗試解析日期
                                            if (!DateTime.TryParse(receivedDateCell.Text, out receivedDate))
                                            {
                                                Debug.WriteLine($"無法解析第 {row} 行的接收日期: {receivedDateCell.Text}");
                                                continue;
                                            }
                                        }

                                        // 讀取完成日期（Y列）
                                        var completionDateCell = worksheet.Cells[row, 25]; // Y列
                                        DateTime? completionDate = null;
                                        if (completionDateCell.Value != null)
                                        {
                                            if (completionDateCell.Value is DateTime completionDateTime)
                                            {
                                                completionDate = completionDateTime;
                                            }
                                            else
                                            {
                                                // 嘗試解析完成日期
                                                DateTime parsedCompletionDate;
                                                if (DateTime.TryParse(completionDateCell.Text, out parsedCompletionDate))
                                                {
                                                    completionDate = parsedCompletionDate;
                                                }
                                            }
                                        }

                                        // 讀取產品類型 (AD列) 
                                        string productType = worksheet.Cells[row, 30].GetValue<string>() ?? "Unknown";

                                        // 讀取銷售代表 (Z列)
                                        string salesRep = worksheet.Cells[row, 26].GetValue<string>() ?? "Unknown";

                                        // 讀取總佣金 (N列)
                                        decimal totalCommission = GetDecimalValue(worksheet.Cells[row, 14]);

                                        // 讀取 PO 值 (G列)
                                        decimal poValue = GetDecimalValue(worksheet.Cells[row, 7]);

                                        // 讀取 Vertiv 值 (H列)
                                        decimal vertivValue = GetDecimalValue(worksheet.Cells[row, 8]);

                                        // 讀取 Agency Margin (M列)
                                        decimal agencyMargin = GetDecimalValue(worksheet.Cells[row, 13]);

                                        // 讀取 Buy Resell Value (J列)
                                        decimal buyResellValue = GetDecimalValue(worksheet.Cells[row, 10]);

                                        // 輸出前幾行的詳細信息用於調試
                                        if (row <= 10)
                                        {
                                            Debug.WriteLine($"行 {row}: " +
                                                           $"接收日期={receivedDate:yyyy-MM-dd}, " +
                                                           $"完成日期={completionDate?.ToString("yyyy-MM-dd") ?? "未完成"}, " +
                                                           $"產品={productType}, " +
                                                           $"代表={salesRep}, " +
                                                           $"總佣金=${totalCommission:N2}, " +
                                                           $"PO值=${poValue:N2}, " +
                                                           $"Vertiv值=${vertivValue:N2}");
                                        }

                                        // 根据完成日期设置订单状态
                                        string status = completionDate.HasValue ? "Completed" : "Booked";

                                        // 檢查是否是 In Progress 數據
                                        bool isInProgress = !completionDate.HasValue && totalCommission == 0;
                                        if (isInProgress)
                                        {
                                            status = "InProgress";
                                            Debug.WriteLine($"檢測到 In Progress 記錄：行 {row}, VertivValue=${vertivValue:N2}");
                                        }

                                        // 如果Y列（完成日期）為空，將總佣金添加到剩餘金額中，但不計入季度業績
                                        if (!completionDate.HasValue)
                                        {
                                            // 這是已預訂但未完成的訂單，將其添加到剩餘金額
                                            remainingAmount += totalCommission;

                                            // 如果是 In Progress 狀態（N列也為空），則計算預期佣金
                                            if (isInProgress)
                                            {
                                                // 獲取H列（Vertiv Value）的值，並添加其12%到inProgressAmount
                                                decimal hColumnValue = vertivValue; // H列是第8列
                                                decimal commission = hColumnValue * 0.12m;
                                                totalHValue += hColumnValue;
                                                calculatedInProgressAmount += commission;
                                                inProgressCount++;
                                                Debug.WriteLine($"[主表] 行 {row}: 添加進行中金額 ${hColumnValue} * 12% = ${commission}");
                                            }
                                        }

                                        var salesData = new SalesData
                                        {
                                            ReceivedDate = receivedDate,
                                            POValue = poValue,        // G列 - PO Value
                                            VertivValue = vertivValue, // H列 - Vertiv Value
                                            BuyResellValue = buyResellValue, // J列 - Buy Resell
                                            AgencyMargin = agencyMargin,   // M列 - Agency Margin
                                            TotalCommission = totalCommission, // N列 - Total Commission
                                            CommissionPercentage = GetDecimalValue(worksheet.Cells[row, 16]), // P列
                                            Status = status,      // 根據Y列和N列確定狀態
                                            CompletionDate = completionDate, // Y列 - 完成日期
                                            SalesRep = salesRep,    // Z列
                                            ProductType = productType, // AD列 - Product Type
                                            Department = worksheet.Cells[row, 29].GetValue<string>() ?? string.Empty,   // AC列 - Department/LOB

                                            // 添加一個標志，表示這是一個"剩餘"項目（Y列為空）
                                            IsRemaining = !completionDate.HasValue,

                                            // 添加一個標志，表示這是一個"進行中"項目（Y列和N列都為空）
                                            IsInProgress = isInProgress,

                                            // 重要：只有已完成的訂單才設置 QuarterDate，未完成的設置為 null
                                            HasQuarterAssigned = completionDate.HasValue,
                                            QuarterDate = completionDate ?? DateTime.MinValue // 使用完成日期作為季度計算日期
                                        };

                                        if (row <= 5)
                                        {
                                            Debug.WriteLine($"第 {row} 行: 接收日期={salesData.ReceivedDate:yyyy-MM-dd}, " +
                                                           $"完成日期={salesData.CompletionDate?.ToString("yyyy-MM-dd") ?? "未完成"}, " +
                                                           $"計入季度={salesData.HasQuarterAssigned}, " +
                                                           (salesData.HasQuarterAssigned ? $"季度計算日期={salesData.QuarterDate:yyyy-MM-dd}, 季度={salesData.Quarter}, " : "") +
                                                           $"POValue=${salesData.POValue}, " +
                                                           $"VertivValue=${salesData.VertivValue}, " +
                                                           $"TotalCommission=${salesData.TotalCommission}, " +
                                                           $"Status={salesData.Status}, " +
                                                           $"IsInProgress={salesData.IsInProgress}");
                                        }

                                        if (IsValidSalesData(salesData))
                                        {
                                            // 跳过 cancelled 状态的数据
                                            if (!salesData.Status?.ToLower().Contains("cancelled") ?? true)
                                            {
                                                data.Add(salesData);
                                            }
                                        }

                                        // 標記該行已處理
                                        processedRows.Add(row);
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"加載第 {row} 行數據時發生錯誤: {ex.Message}");
                                        // 繼續處理下一行
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine($"行 {row} 已處理過，跳過");
                                }
                            }

                            // 在所有行处理完毕后，一次性设置静态变量
                            _remainingAmount = remainingAmount;
                            _inProgressAmount = calculatedInProgressAmount;

                            Debug.WriteLine($"===== 计算完成 =====");
                            Debug.WriteLine($"进行中订单总数: {inProgressCount}");
                            Debug.WriteLine($"进行中订单H列总额: ${totalHValue:N2}");
                            Debug.WriteLine($"计算得到的12%佣金总额: ${calculatedInProgressAmount:N2}");
                            Debug.WriteLine($"计算得到的剩余金额: ${_remainingAmount:N2}");
                            Debug.WriteLine($"设置的正在进行中金额: ${_inProgressAmount:N2}");
                            Debug.WriteLine($"===== 计算结束 =====");
                        }

                        // 尝试读取摘要工作表
                        try
                        {
                            LoadSummarySheets(package);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"读取摘要工作表时发生错误: {ex.Message}");
                        }
                    }

                    Debug.WriteLine($"成功加载 {data.Count} 条有效记录");
                    return (data, lastUpdated);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"加载 Excel 数据时发生错误: {ex.Message}\n{ex.StackTrace}");
                    return (new List<SalesData>(), DateTime.Now);
                }
            });
        }

        // 載入所有摘要工作表
        private void LoadSummarySheets(ExcelPackage package)
        {
            try
            {
                Debug.WriteLine("===== 開始加載摘要工作表 =====");

                // 先載入 Sales Leaderboard
                var leaderboardSheet = FindWorksheet(package, _leaderboardSheetName);
                if (leaderboardSheet != null)
                {
                    LoadSalesLeaderboardData(leaderboardSheet);
                }
                else
                {
                    Debug.WriteLine($"找不到 {_leaderboardSheetName} 工作表");
                }

                // 載入 By Dept-LOB
                var deptLobSheet = FindWorksheet(package, _byDeptLobSheetName);
                if (deptLobSheet != null)
                {
                    LoadDeptLobData(deptLobSheet);
                }
                else
                {
                    Debug.WriteLine($"找不到 {_byDeptLobSheetName} 工作表");
                }

                // 載入 By Rep
                var byRepSheet = FindWorksheet(package, _byRepSheetName);
                if (byRepSheet != null)
                {
                    LoadByRepData(byRepSheet);
                }
                else
                {
                    Debug.WriteLine($"找不到 {_byRepSheetName} 工作表");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"讀取摘要工作表時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
            }
            finally
            {
                Debug.WriteLine("===== 摘要工作表加載完成 =====");
            }
        }

        // 找到工作表，支援名稱或索引查找
        private ExcelWorksheet FindWorksheet(ExcelPackage package, string sheetName)
        {
            // 先按名稱嘗試
            var worksheet = package.Workbook.Worksheets[sheetName];
            if (worksheet != null)
            {
                Debug.WriteLine($"成功找到工作表: {sheetName}");
                return worksheet;
            }

            // 查找包含指定名稱的工作表
            foreach (var sheet in package.Workbook.Worksheets)
            {
                if (sheet.Name.Contains(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine($"找到類似名稱的工作表: {sheet.Name}");
                    return sheet;
                }
            }

            Debug.WriteLine($"找不到工作表: {sheetName}");
            return null;
        }

        // 載入 Sales Leaderboard 工作表數據
        private void LoadSalesLeaderboardData(ExcelWorksheet sheet)
        {
            try
            {
                int totalRows = sheet.Dimension?.Rows ?? 0;
                if (totalRows <= 1) // 如果只有標題行或根本沒有行
                {
                    Debug.WriteLine("Sales Leaderboard 工作表沒有數據或只有標題行");
                    return;
                }

                Debug.WriteLine($"Sales Leaderboard 工作表有 {totalRows} 行數據");

                // 尋找列索引（注意：Excel 中第一行通常是標題）
                int salesRepColIndex = -1;
                int agencyMarginColIndex = -1;
                int buyResellMarginColIndex = -1;
                int totalMarginColIndex = -1;
                int vertivValueColIndex = -1; // 新增的Vertiv Value列索引

                // 掃描標題行找出欄位位置
                for (int col = 1; col <= sheet.Dimension.Columns; col++)
                {
                    string headerText = (sheet.Cells[1, col].Text ?? "").Trim().ToLower();

                    if (headerText.Contains("sales rep") || headerText.Contains("rep"))
                        salesRepColIndex = col;
                    else if (headerText.Contains("agency") && headerText.Contains("margin"))
                        agencyMarginColIndex = col;
                    else if (headerText.Contains("buy"))
                        buyResellMarginColIndex = col;
                    else if (headerText.Contains("total") && headerText.Contains("margin"))
                        totalMarginColIndex = col;
                    else if (headerText.Contains("vertiv") || headerText.Contains("value"))
                        vertivValueColIndex = col;
                }

                // 檢查是否找到所需欄位
                if (salesRepColIndex < 0 || totalMarginColIndex < 0)
                {
                    Debug.WriteLine("Sales Leaderboard 工作表中找不到必要的欄位");
                    return;
                }

                // 處理資料行
                var data = new List<SalesLeaderboardItem>();
                for (int row = 2; row <= totalRows; row++) // 從第二行開始，跳過標題
                {
                    string salesRep = sheet.Cells[row, salesRepColIndex].Text;

                    // 跳過空行或總計行
                    if (string.IsNullOrWhiteSpace(salesRep) || salesRep.ToLower().Contains("total"))
                        continue;

                    decimal agencyMargin = 0;
                    decimal buyResellMargin = 0;
                    decimal totalMargin = 0;
                    decimal vertivValue = 0;

                    // 讀取 Agency Margin
                    if (agencyMarginColIndex > 0)
                    {
                        decimal.TryParse(
                            sheet.Cells[row, agencyMarginColIndex].Text.Replace("$", "").Replace(",", ""),
                            out agencyMargin);
                    }

                    // 讀取 Buy Resell Margin
                    if (buyResellMarginColIndex > 0)
                    {
                        decimal.TryParse(
                            sheet.Cells[row, buyResellMarginColIndex].Text.Replace("$", "").Replace(",", ""),
                            out buyResellMargin);
                    }

                    // 讀取 Total Margin
                    if (totalMarginColIndex > 0)
                    {
                        decimal.TryParse(
                            sheet.Cells[row, totalMarginColIndex].Text.Replace("$", "").Replace(",", ""),
                            out totalMargin);
                    }

                    // 讀取 Vertiv Value
                    if (vertivValueColIndex > 0)
                    {
                        decimal.TryParse(
                            sheet.Cells[row, vertivValueColIndex].Text.Replace("$", "").Replace(",", ""),
                            out vertivValue);
                    }

                    // 如果 Total Margin 為 0 且有 Agency 或 BuyResell，計算總和
                    if (totalMargin == 0 && (agencyMargin > 0 || buyResellMargin > 0))
                    {
                        totalMargin = agencyMargin + buyResellMargin;
                    }

                    data.Add(new SalesLeaderboardItem
                    {
                        SalesRep = salesRep,
                        AgencyMargin = agencyMargin,
                        BuyResellMargin = buyResellMargin,
                        TotalMargin = totalMargin,
                        VertivValue = vertivValue, // 新增Vertiv Value
                        Rank = 0 // 先設為 0，稍後再計算
                    });
                }

                // 按照 Total Margin 排序並設定排名
                data = data.OrderByDescending(x => x.TotalMargin).ToList();
                for (int i = 0; i < data.Count; i++)
                {
                    data[i].Rank = i + 1;
                }

                // 更新緩存
                _salesLeaderboardCache = data;
                Debug.WriteLine($"成功從工作表載入 {data.Count} 條銷售代表數據");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入 Sales Leaderboard 數據時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                // 不使用硬編碼資料，而是保持緩存為當前值或空列表
                if (_salesLeaderboardCache == null)
                    _salesLeaderboardCache = new List<SalesLeaderboardItem>();
            }
        }

        // 載入 By Dept-LOB 工作表數據
        private void LoadDeptLobData(ExcelWorksheet sheet)
        {
            try
            {
                int totalRows = sheet.Dimension?.Rows ?? 0;
                if (totalRows <= 1) // If only header or no rows
                {
                    Debug.WriteLine("By Dept-LOB worksheet has no data");
                    return;
                }

                var deptLobData = new List<DepartmentLobData>();

                // Find proper column indices for the By Dept-LOB sheet
                int lobColumnIndex = -1;
                int marginTargetColumnIndex = -1;
                int marginYtdColumnIndex = -1;

                // Scan first row for headers to identify columns
                for (int col = 1; col <= sheet.Dimension.Columns; col++)
                {
                    string headerText = (sheet.Cells[1, col].Text ?? "").Trim().ToLower();
                    if (headerText.Contains("lob"))
                        lobColumnIndex = col;
                    else if (headerText.Contains("target") || headerText.Contains("f25 margin target"))
                        marginTargetColumnIndex = col;
                    else if (headerText.Contains("ytd") || headerText.Contains("margin ytd"))
                        marginYtdColumnIndex = col;
                }

                if (lobColumnIndex < 0 || marginTargetColumnIndex < 0 || marginYtdColumnIndex < 0)
                {
                    Debug.WriteLine("Couldn't find required columns in By Dept-LOB sheet");
                    return;
                }

                // Process data rows
                for (int row = 2; row <= totalRows; row++)
                {
                    string lob = sheet.Cells[row, lobColumnIndex].Text;

                    if (string.IsNullOrWhiteSpace(lob))
                        continue;

                    decimal marginTarget = 0;
                    decimal marginYtd = 0;

                    // Try parse margin target
                    decimal.TryParse(
                        sheet.Cells[row, marginTargetColumnIndex].Text.Replace("$", "").Replace(",", ""),
                        out marginTarget);

                    // Try parse margin YTD
                    decimal.TryParse(
                        sheet.Cells[row, marginYtdColumnIndex].Text.Replace("$", "").Replace(",", ""),
                        out marginYtd);

                    deptLobData.Add(new DepartmentLobData
                    {
                        Rank = row - 1, // Assign rank based on row position
                        LOB = lob,
                        MarginTarget = marginTarget,
                        MarginYTD = marginYtd
                    });
                }

                // Check if we parsed any data
                if (deptLobData.Any())
                {
                    // Update the "Total" row's rank to 0 to ensure it appears at the bottom
                    var totalRow = deptLobData.FirstOrDefault(x => x.LOB.ToLower() == "total");
                    if (totalRow != null)
                        totalRow.Rank = 0;

                    _departmentLobCache = deptLobData;
                    Debug.WriteLine($"Successfully loaded {deptLobData.Count} Dept-LOB records from Excel");
                    return;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading By Dept-LOB data: {ex.Message}");
            }

            // If we failed to load data, we'll add sample data in the calling method
        }

        // 載入 By Rep 工作表數據
        private void LoadByRepData(ExcelWorksheet sheet)
        {
            try
            {
                int totalRows = sheet.Dimension?.Rows ?? 0;
                if (totalRows == 0)
                {
                    Debug.WriteLine("By Rep 工作表沒有數據");
                    return;
                }

                // 直接使用硬編碼的測試數據
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入 By Rep 數據時發生錯誤: {ex.Message}");
            }
        }

        public void ClearCache()
        {
            _productSalesCache = new List<ProductSalesData>();
            _salesLeaderboardCache = new List<SalesLeaderboardItem>();
            _departmentLobCache = new List<DepartmentLobData>();

            // 記錄緩存清除
            Debug.WriteLine("緩存已完全清除");
        }


        private DateTime _lastCacheUpdate = DateTime.MinValue;

        // 獲取產品銷售數據（從緩存）
        public List<ProductSalesData> GetProductSalesData()
        {
            try
            {
                Debug.WriteLine("GetProductSalesData: 開始獲取產品銷售數據");

                // 如果需要，先嘗試加載數據
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    Debug.WriteLine("GetProductSalesData: 沒有已加載的數據，嘗試從文件加載");
                    try
                    {
                        var (data, _) = LoadDataAsync().GetAwaiter().GetResult();
                        _allSalesData = data;
                        Debug.WriteLine($"GetProductSalesData: 從文件加載了 {_allSalesData.Count} 條記錄");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"GetProductSalesData: 從文件加載數據失敗: {ex.Message}");
                        return new List<ProductSalesData>(); // 返回空列表，而不是使用硬編碼數據
                    }
                }

                // 使用所有數據（不過濾）來計算產品銷售數據
                var products = _allSalesData
                    .GroupBy(x => NormalizeProductType(x.ProductType))
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                        TotalMargin = Math.Round(g.Sum(x => x.TotalCommission), 2),
                        // 修改为使用 VertivValue 代替 POValue
                        POValue = Math.Round(g.Sum(x => x.VertivValue), 2)
                    })
                    .OrderByDescending(x => x.POValue)
                    .ToList();

                // 計算百分比
                decimal totalPOValue = products.Sum(p => p.POValue);
                foreach (var product in products)
                {
                    product.PercentageOfTotal = totalPOValue > 0
                        ? Math.Round((product.POValue / totalPOValue) * 100, 1)
                        : 0;
                }

                Debug.WriteLine($"GetProductSalesData: 計算了 {products.Count} 個產品類型的數據");

                return products;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetProductSalesData: 處理數據時發生錯誤: {ex.Message}");
                return new List<ProductSalesData>(); // 返回空列表
            }
        }

        // 標準化產品類型名稱
        private string NormalizeProductType(string productType)
        {
            if (string.IsNullOrEmpty(productType))
                return "Other";

            // 转为小写以便比较
            string lowercaseType = productType.ToLowerInvariant();

            if (lowercaseType.Contains("thermal"))
                return "Thermal";
            if (lowercaseType.Contains("power") || lowercaseType.Contains("saskpower"))
                return "Power";
            if (lowercaseType.Contains("channel"))
                return "Channel";
            if (lowercaseType.Contains("service"))
                return "Service";
            if (lowercaseType.Contains("batts") || lowercaseType.Contains("caps") || lowercaseType.Contains("batt"))
                return "Batts & Caps";

            // 如果没有匹配，返回原始名称
            return productType;
        }


        // 獲取銷售代表排行榜數據
        public List<SalesLeaderboardItem> GetSalesLeaderboardData()
        {
            try
            {
                Debug.WriteLine("GetSalesLeaderboardData: 開始獲取銷售代表排行榜數據");

                // 如果需要，先嘗試加載數據
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    Debug.WriteLine("GetSalesLeaderboardData: 沒有已加載的數據，嘗試從文件加載");
                    try
                    {
                        var (data, _) = LoadDataAsync().GetAwaiter().GetResult();
                        _allSalesData = data;
                        Debug.WriteLine($"GetSalesLeaderboardData: 從文件加載了 {_allSalesData.Count} 條記錄");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"GetSalesLeaderboardData: 從文件加載數據失敗: {ex.Message}");
                        return new List<SalesLeaderboardItem>(); // 返回空列表
                    }
                }

                // 使用所有數據（不過濾）來計算銷售代表排行榜數據
                var reps = _allSalesData
                    .GroupBy(x => x.SalesRep)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new SalesLeaderboardItem
                    {
                        SalesRep = g.Key,
                        AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                        TotalMargin = Math.Round(g.Sum(x => x.TotalCommission), 2),
                        // 修改为使用 VertivValue 代替 POValue
                        VertivValue = Math.Round(g.Sum(x => x.VertivValue), 2)
                    })
                    .OrderByDescending(x => x.TotalMargin)
                    .ToList();

                // 設置排名
                for (int i = 0; i < reps.Count; i++)
                {
                    reps[i].Rank = i + 1;
                }

                Debug.WriteLine($"GetSalesLeaderboardData: 計算了 {reps.Count} 個銷售代表的數據");

                return reps;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetSalesLeaderboardData: 處理數據時發生錯誤: {ex.Message}");
                return new List<SalesLeaderboardItem>(); // 返回空列表
            }
        }

        // 獲取部門/LOB數據
        public List<DepartmentLobData> GetDepartmentLobData()
        {
            try
            {
                Debug.WriteLine("GetDepartmentLobData: 開始獲取部門/LOB數據");

                // 如果需要，先嘗試加載數據
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    Debug.WriteLine("GetDepartmentLobData: 沒有已加載的數據，嘗試從文件加載");
                    try
                    {
                        var (data, _) = LoadDataAsync().GetAwaiter().GetResult();
                        _allSalesData = data;
                        Debug.WriteLine($"GetDepartmentLobData: 從文件加載了 {_allSalesData.Count} 條記錄");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"GetDepartmentLobData: 從文件加載數據失敗: {ex.Message}");
                        return new List<DepartmentLobData>(); // 返回空列表
                    }
                }

                // 設置預定義的LOB和目標值
                var lobData = new Dictionary<string, (decimal target, decimal ytd)>
        {
            { "Power", (850000, 0) },
            { "Thermal", (720000, 0) },
            { "Channel", (650000, 0) },
            { "Service", (580000, 0) },
            { "Batts & Caps", (450000, 0) }
        };

                decimal totalTarget = lobData.Sum(kv => kv.Value.target);
                decimal totalYtd = 0;

                // 從所有數據計算YTD值
                if (_allSalesData != null && _allSalesData.Any())
                {
                    Debug.WriteLine($"GetDepartmentLobData: 處理 {_allSalesData.Count} 條銷售數據進行LOB分類");

                    // 記錄找到的所有Department值，幫助調試
                    var allDepts = _allSalesData
                        .Select(x => x.Department)
                        .Where(d => !string.IsNullOrWhiteSpace(d))
                        .Distinct()
                        .ToList();

                    Debug.WriteLine($"GetDepartmentLobData: 找到的所有Department值: {string.Join(", ", allDepts)}");

                    foreach (var item in _allSalesData)
                    {
                        // 跳過沒有Department值的記錄
                        if (string.IsNullOrWhiteSpace(item.Department))
                            continue;

                        // 標準化Department值
                        string lob = NormalizeDepartment(item.Department);

                        // 如果不是已定義的LOB，跳過或添加到「其他」類別
                        if (!lobData.ContainsKey(lob))
                        {
                            if (!lobData.ContainsKey("Other"))
                            {
                                lobData["Other"] = (200000, 0);
                                totalTarget += 200000;
                            }

                            var current = lobData["Other"];
                            lobData["Other"] = (current.target, current.ytd + item.TotalCommission);
                            totalYtd += item.TotalCommission;
                            continue;
                        }

                        // 更新此LOB的YTD值
                        var lobCurrent = lobData[lob];
                        lobData[lob] = (lobCurrent.target, lobCurrent.ytd + item.TotalCommission);

                        // 更新總YTD
                        totalYtd += item.TotalCommission;
                    }
                }
                else
                {
                    Debug.WriteLine("GetDepartmentLobData: 沒有銷售數據可用於LOB分類");
                }

                // 轉換為DepartmentLobData列表
                var result = new List<DepartmentLobData>();
                int rank = 1;

                foreach (var entry in lobData)
                {
                    result.Add(new DepartmentLobData
                    {
                        Rank = rank++,
                        LOB = entry.Key,
                        MarginTarget = entry.Value.target,
                        MarginYTD = entry.Value.ytd
                    });
                }

                // 添加總計行
                result.Add(new DepartmentLobData
                {
                    Rank = 0,  // 設為0，使總計行排在最後
                    LOB = "Total",
                    MarginTarget = totalTarget,
                    MarginYTD = totalYtd
                });

                // 按YTD降序排序（但保持Total在最後）
                result = result
                    .OrderBy(x => x.LOB == "Total" ? 1 : 0)  // Total排在最後
                    .ThenByDescending(x => x.MarginYTD)      // 其他按YTD降序
                    .ToList();

                // 重新分配排名
                for (int i = 0; i < result.Count; i++)
                {
                    if (result[i].LOB != "Total")
                    {
                        result[i].Rank = i + 1;
                    }
                }

                Debug.WriteLine($"GetDepartmentLobData: 生成了 {result.Count} 條LOB數據");
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetDepartmentLobData: 獲取部門數據時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                return new List<DepartmentLobData>(); // 返回空列表
            }
        }



        // 標準化Department名稱
        private string NormalizeDepartment(string department)
        {
            if (string.IsNullOrWhiteSpace(department))
                return "Other";

            if (department.Contains("Power", StringComparison.OrdinalIgnoreCase))
                return "Power";
            if (department.Contains("Thermal", StringComparison.OrdinalIgnoreCase))
                return "Thermal";
            if (department.Contains("Channel", StringComparison.OrdinalIgnoreCase))
                return "Channel";
            if (department.Contains("Service", StringComparison.OrdinalIgnoreCase))
                return "Service";
            if (department.Contains("Batts", StringComparison.OrdinalIgnoreCase) ||
                department.Contains("Caps", StringComparison.OrdinalIgnoreCase))
                return "Batts & Caps";

            return department; // Keep original if not matching predefined categories
        }

        // 輔助函數：將產品類型映射到LOB
        private string GetLOBFromProductType(string productType)
        {
            if (string.IsNullOrWhiteSpace(productType))
                return "Other";

            if (productType.Contains("Power", StringComparison.OrdinalIgnoreCase))
                return "Power";
            if (productType.Contains("Thermal", StringComparison.OrdinalIgnoreCase))
                return "Thermal";
            if (productType.Contains("Channel", StringComparison.OrdinalIgnoreCase))
                return "Channel";
            if (productType.Contains("Service", StringComparison.OrdinalIgnoreCase))
                return "Service";

            // 其他產品類型歸為"Other"，這裡可以根據需要調整
            return "Other";
        }

        // 設置過濾後的數據
        public void SetFilteredData(List<SalesData> filteredData)
        {
            if (filteredData != null)
            {
                _recentDataCache = filteredData;
                Debug.WriteLine($"已設置 {filteredData.Count} 條過濾後的數據到ExcelService");

                // 清除現有緩存，確保數據重新計算
                _departmentLobCache.Clear();
            }
        }

        public List<SalesData> GetInProgressItems(DateTime startDate, DateTime endDate)
        {
            try
            {
                // Ensure data is loaded
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    var (data, _) = LoadDataAsync().GetAwaiter().GetResult();
                    _allSalesData = data;
                }

                // Filter by date range and find records with empty Y column (CompletionDate) 
                // and empty N column (TotalCommission)
                var inProgressItems = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate.Date &&
                           x.ReceivedDate.Date <= endDate.Date &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission == 0)
                    .ToList();

                Debug.WriteLine($"Found {inProgressItems.Count} in-progress items");
                return inProgressItems;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting in-progress items: {ex.Message}");
                return new List<SalesData>();
            }
        }

        public async Task<bool> UpdateDataAsync(string filePath = Constants.EXCEL_FILE_NAME, List<SalesData> data = null)
        {
            try
            {
                string fullPath = Path.Combine(Constants.BASE_PATH, filePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(fullPath)))
                {
                    var worksheet = package.Workbook.Worksheets[_worksheetName];
                    if (worksheet == null)
                        throw new Exception($"找不到工作表: '{_worksheetName}'");

                    if (data != null)
                    {
                        // TODO: 實作更新Excel的邏輯
                    }

                    await package.SaveAsync();
                }

                DataUpdated?.Invoke(this, DateTime.Now);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新 Excel 數據時發生錯誤: {ex.Message}");
                return false;
            }
        }

        public async Task<bool> MonitorFileChangesAsync(CancellationToken token)
        {
            try
            {
                string fullPath = Path.Combine(Constants.BASE_PATH, Constants.EXCEL_FILE_NAME);
                var fileInfo = new FileInfo(fullPath);

                if (!fileInfo.Exists)
                {
                    Debug.WriteLine($"監控的文件不存在: {fullPath}");
                    return false;
                }

                _watcher = new FileSystemWatcher
                {
                    Path = fileInfo.DirectoryName,
                    Filter = fileInfo.Name,
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size
                };
                _watcher.Changed += async (s, e) =>
                {
                    try
                    {
                        await Task.Delay(1000, token);
                        DataUpdated?.Invoke(this, File.GetLastWriteTime(fullPath));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"檔案變更處理錯誤: {ex.Message}");
                    }
                };

                _watcher.EnableRaisingEvents = true;

                try
                {
                    await Task.Delay(-1, token);
                }
                catch (OperationCanceledException)
                {
                    _watcher.EnableRaisingEvents = false;
                    _watcher.Dispose();
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"監控檔案時發生錯誤: {ex.Message}");
                return false;
            }
        }

        private decimal GetDecimalValue(ExcelRange cell)
        {
            if (cell == null || cell.Value == null || string.IsNullOrWhiteSpace(cell.Text))
                return 0;

            string value = cell.Text.Replace("$", "").Replace(",", "").Trim();

            if (decimal.TryParse(value, out decimal result))
                return result;

            return 0;
        }

        private bool IsValidSalesData(SalesData data)
        {
            if (data.ReceivedDate == DateTime.MinValue)
            {
                return false;
            }
            if (string.IsNullOrWhiteSpace(data.SalesRep))
            {
                return false;
            }
            return true;
        }


        /// <summary>
        /// 獲取所有銷售代表列表
        /// </summary>
        /// <returns>銷售代表列表</returns>
        public List<string> GetAllSalesReps()
        {
            try
            {
                // 如果 _recentDataCache 有數據，從中獲取
                if (_recentDataCache != null && _recentDataCache.Any())
                {
                    return _recentDataCache
                        .Where(x => !string.IsNullOrWhiteSpace(x.SalesRep))
                        .Select(x => x.SalesRep)
                        .Distinct()
                        .OrderBy(x => x)
                        .ToList();
                }

                // 否則，嘗試從主資料源讀取
                var (data, _) = LoadDataAsync().GetAwaiter().GetResult();
                return data
                    .Where(x => !string.IsNullOrWhiteSpace(x.SalesRep))
                    .Select(x => x.SalesRep)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"獲取銷售代表列表時出錯: {ex.Message}");
                // 返回一些默認數據
                return new List<string> { "Brandon", "Chris", "Isaac", "Mark", "Nathan", "Tania" };
            }
        }

        /// <summary>
        /// 獲取所有產品線(LOB)列表
        /// </summary>
        /// <returns>產品線列表</returns>
        public List<string> GetAllLOBs()
        {
            try
            {
                Debug.WriteLine("Getting all LOBs from Excel");

                // Try to load data if needed
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    var (data, _) = LoadDataAsync().GetAwaiter().GetResult();
                    _allSalesData = data;
                }

                // Extract LOBs from Department field
                var lobs = _allSalesData
                    .Where(x => !string.IsNullOrWhiteSpace(x.Department))
                    .Select(x => NormalizeDepartment(x.Department))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                Debug.WriteLine($"Found {lobs.Count} unique LOBs from Excel");

                // If no LOBs found, return default values
                if (!lobs.Any())
                {
                    Debug.WriteLine("No LOBs found, returning default values");
                    return new List<string> { "Power", "Thermal", "Channel", "Service", "Batts & Caps" };
                }

                return lobs;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting LOB list: {ex.Message}");
                // Return default values in case of error
                return new List<string> { "Power", "Thermal", "Channel", "Service", "Batts & Caps" };
            }
        }
    }
}