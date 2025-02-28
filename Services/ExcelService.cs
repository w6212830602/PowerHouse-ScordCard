using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ScoreCard.Models;
using OfficeOpenXml;
using System.Diagnostics;

namespace ScoreCard.Services
{
    public class ExcelService : IExcelService
    {
        private readonly string _worksheetName = Constants.WORKSHEET_NAME;
        private FileSystemWatcher _watcher;
        public event EventHandler<DateTime> DataUpdated;

        public async Task<(List<SalesData> data, DateTime lastUpdated)> LoadDataAsync(string filePath = Constants.EXCEL_FILE_NAME)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string fullPath = Path.Combine(Constants.BASE_PATH, filePath);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var data = new List<SalesData>();
                    DateTime lastUpdated = File.GetLastWriteTime(fullPath);

                    using (var package = new ExcelPackage(new FileInfo(fullPath)))
                    {
                        var worksheet = package.Workbook.Worksheets[_worksheetName];
                        if (worksheet == null)
                            throw new Exception($"Worksheet '{_worksheetName}' not found");

                        var rowCount = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                var salesData = new SalesData
                                {
                                    ReceivedDate = worksheet.Cells[row, 1].GetValue<DateTime>(),  // A欄
                                    POValue = GetDecimalValue(worksheet.Cells[row, 7]),           // G欄
                                    VertivValue = GetDecimalValue(worksheet.Cells[row, 8]),       // H欄
                                    TotalCommission = GetDecimalValue(worksheet.Cells[row, 14]),  // N欄
                                    CommissionPercentage = GetDecimalValue(worksheet.Cells[row, 16]), // P欄
                                    Status = worksheet.Cells[row, 18].GetValue<string>(),         // R欄
                                    SalesRep = worksheet.Cells[row, 26].GetValue<string>(),       // Z欄
                                    ProductType = worksheet.Cells[row, 30].GetValue<string>()     // AD欄
                                };
                                Debug.WriteLine($"Reading row {row}: Date={salesData.ReceivedDate}, Status={salesData.Status}");

                                if (IsValidSalesData(salesData))
                                {
                                    // 跳過 cancelled 狀態的數據
                                    if (!salesData.Status?.ToLower().Contains("cancelled") ?? true)
                                    {
                                        data.Add(salesData);
                                    }
                                    else
                                    {
                                        Debug.WriteLine($"Skipped cancelled record at row {row}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error reading row {row}: {ex.Message}");
                                continue;
                            }
                        }
                    }

                    // 按季度分組統計
                    var quarterlyStats = data.GroupBy(x => x.Quarter)
                                           .ToDictionary(g => g.Key, g => new
                                           {
                                               TotalPOValue = g.Sum(x => x.POValue),
                                               TotalVertivValue = g.Sum(x => x.VertivValue),
                                               TotalCommission = g.Sum(x => x.TotalCommission)
                                           });

                    System.Diagnostics.Debug.WriteLine($"Loaded {data.Count} valid records");
                    foreach (var stat in quarterlyStats)
                    {
                        System.Diagnostics.Debug.WriteLine($"Q{stat.Key}: Target=${stat.Value.TotalPOValue:N0}, Achieved=${stat.Value.TotalVertivValue:N0}");
                    }

                    return (data, lastUpdated);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error loading Excel data: {ex.Message}");
                    throw;
                }
            });
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
                        throw new Exception($"Worksheet '{_worksheetName}' not found");

                    if (data != null)
                    {
                        // TODO: 實作更新Excel的邏輯
                        // 這裡可以根據需求添加具體的更新邏輯
                    }

                    await package.SaveAsync();
                }

                DataUpdated?.Invoke(this, DateTime.Now);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating Excel data: {ex.Message}");
                return false;
            }
        }

        public async Task<bool> MonitorFileChangesAsync(CancellationToken token)
        {
            try
            {
                string fullPath = Path.Combine(Constants.BASE_PATH, Constants.EXCEL_FILE_NAME);
                var fileInfo = new FileInfo(fullPath);

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
                        await Task.Delay(1000, token);  // 使用傳入的 token
                        DataUpdated?.Invoke(this, File.GetLastWriteTime(fullPath));
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error in file change handler: {ex.Message}");
                    }
                };

                _watcher.EnableRaisingEvents = true;

                try
                {
                    await Task.Delay(-1, token);  // 使用傳入的 token
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
                Debug.WriteLine($"Error monitoring file: {ex.Message}");
                return false;
            }
        }

        private decimal GetDecimalValue(ExcelRange cell)
        {
            if (cell == null || string.IsNullOrWhiteSpace(cell.Text))
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
                Debug.WriteLine($"記錄無效: 日期為空值");
                return false;
            }
            if (string.IsNullOrWhiteSpace(data.SalesRep))
            {
                Debug.WriteLine($"記錄無效: SalesRep為空值，日期: {data.ReceivedDate}");
                return false;
            }
            if (string.IsNullOrWhiteSpace(data.Status))
            {
                Debug.WriteLine($"記錄無效: Status為空值，日期: {data.ReceivedDate}");
                return false;
            }
            if (data.POValue < 0)
            {
                Debug.WriteLine($"記錄無效: POValue為負值 ({data.POValue})，日期: {data.ReceivedDate}");
                return false;
            }
            return true;
        }
    }
}