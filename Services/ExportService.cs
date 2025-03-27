using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ScoreCard.Models;


namespace ScoreCard.Services
{
    public class ExportService : IExportService
    {

        private readonly IExcelService _excelService;
        public ExportService(IExcelService excelService)
        {
            _excelService = excelService;
        }

        // 獲取產品數據的方法
        private List<ProductSalesData> GetProductDataForSexyReport()
        {
            return _excelService.GetProductSalesData();
        }

        /// <summary>
        /// 將數據匯出為Excel檔案
        /// </summary>
        public async Task<bool> ExportToExcelAsync<T>(IEnumerable<T> data, string fileName, string title)
        {
            if (data == null || !data.Any())
            {
                Debug.WriteLine("無數據可匯出");
                await ShowErrorMessage("匯出錯誤", "沒有資料可以匯出。請確保視圖中有顯示數據。");
                return false;
            }

            try
            {
                Debug.WriteLine($"開始匯出 Excel，數據項數: {data.Count()}");

                string exportPath = GetExportDirectory();
                string fullPath = Path.Combine(exportPath, $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Report");

                    // 標題
                    worksheet.Cells[1, 1].Value = title;
                    worksheet.Cells[1, 1, 1, 10].Merge = true;
                    worksheet.Cells[1, 1].Style.Font.Bold = true;
                    worksheet.Cells[1, 1].Style.Font.Size = 16;
                    worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    // 生成日期
                    worksheet.Cells[2, 1].Value = $"Generate date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                    worksheet.Cells[2, 1, 2, 10].Merge = true;
                    worksheet.Cells[2, 1].Style.Font.Size = 12;
                    worksheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    int currentRow = 4;

                    // 第一個表格 - F25 Sales Rep Commission
                    currentRow = AddTableToWorksheet(worksheet, data, currentRow, "Sales Rep Commission by Product Type (Margin Achieved)");

                    // 添加間隔
                    currentRow += 2;

                    // 第二個表格 - Sexy Report
                    // 不管當前是什麼視圖，都添加 Sexy Report
                    // 需要從 DetailedSalesViewModel 獲取產品數據
                    if (data.FirstOrDefault() is ProductSalesData productData)
                    {
                        // 直接使用產品數據
                        var productDataList = data.Cast<ProductSalesData>().ToList();
                        currentRow = AddSexyReportToWorksheet(worksheet, productDataList, currentRow);
                    }
                    else
                    {
                        // 如果是 SalesLeaderboardItem 視圖，需要獲取產品數據
                        // 這裡需要通過服務或屬性獲取產品數據
                        var productDataList = GetProductDataForSexyReport();
                        if (productDataList != null && productDataList.Any())
                        {
                            currentRow = AddSexyReportToWorksheet(worksheet, productDataList, currentRow);
                        }
                    }

                    // 設置列寬自適應
                    for (int i = 1; i <= 10; i++) // 假設最多10列
                    {
                        worksheet.Column(i).AutoFit();
                    }

                    // 保存文件
                    await package.SaveAsAsync(new FileInfo(fullPath));
                    Debug.WriteLine($"Excel 檔案已保存至: {fullPath}");
                }

                // 顯示成功消息
                await ShowSuccessMessage("Excel Report Generated", $"File saved to: {fullPath}");

                // 嘗試打開文件夾
#if WINDOWS
        OpenFolder(exportPath);
#endif

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"匯出Excel時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await ShowErrorMessage("匯出錯誤", $"匯出Excel檔案時發生錯誤: {ex.Message}");
                return false;
            }
        }

        // 添加表格到工作表
        private int AddTableToWorksheet<T>(ExcelWorksheet worksheet, IEnumerable<T> data, int startRow, string tableTitle)
        {
            // 表格標題
            worksheet.Cells[startRow, 1].Value = tableTitle;
            worksheet.Cells[startRow, 1, startRow, 6].Merge = true;
            worksheet.Cells[startRow, 1].Style.Font.Bold = true;
            worksheet.Cells[startRow, 1].Style.Font.Size = 12;
            startRow++;

            // 手動定義欄位名稱和對應屬性
            List<(string Header, string Property)> columns = new List<(string, string)>();

            // 根據類型添加適當欄位
            var firstItem = data.FirstOrDefault();
            if (firstItem is ProductSalesData)
            {
                columns.Add(("Product Type", "ProductType"));
                columns.Add(("Agency Commission", "AgencyCommission"));
                columns.Add(("Buy Resell Commission", "BuyResellCommission"));
                columns.Add(("Total Commission", "TotalCommission"));
                columns.Add(("PO Value", "POValue"));
                columns.Add(("% of Total", "PercentageOfTotal"));
            }
            else if (firstItem is SalesLeaderboardItem)
            {
                columns.Add(("Rank", "Rank"));
                columns.Add(("Sales Rep", "SalesRep"));
                columns.Add(("Agency Commission", "AgencyCommission"));
                columns.Add(("Buy Resell Commission", "BuyResellCommission"));
                columns.Add(("Total Commission", "TotalCommission"));
            }

            // 寫入表頭
            int colIndex = 1;
            foreach (var col in columns)
            {
                worksheet.Cells[startRow, colIndex].Value = col.Header;
                worksheet.Cells[startRow, colIndex].Style.Font.Bold = true;
                worksheet.Cells[startRow, colIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, colIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                worksheet.Cells[startRow, colIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[startRow, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                colIndex++;
            }
            startRow++;

            // 寫入數據
            Type itemType = firstItem.GetType();
            decimal totalAgencyComm = 0;
            decimal totalBuyResellComm = 0;
            decimal totalComm = 0;
            decimal totalPOValue = 0;

            foreach (var item in data)
            {
                colIndex = 1;
                foreach (var col in columns)
                {
                    var prop = itemType.GetProperty(col.Property);
                    if (prop != null)
                    {
                        try
                        {
                            var value = prop.GetValue(item);

                            if (value != null)
                            {
                                // 直接寫入值
                                worksheet.Cells[startRow, colIndex].Value = value;

                                // 累計總計
                                if (col.Property == "AgencyCommission" && value is decimal agencyComm)
                                    totalAgencyComm += agencyComm;
                                else if (col.Property == "BuyResellCommission" && value is decimal buyResellComm)
                                    totalBuyResellComm += buyResellComm;
                                else if (col.Property == "TotalCommission" && value is decimal comm)
                                    totalComm += comm;
                                else if (col.Property == "POValue" && value is decimal poValue)
                                    totalPOValue += poValue;

                                // 設定格式
                                if (value is decimal decimalValue)
                                {
                                    if (col.Property == "PercentageOfTotal")
                                    {
                                        worksheet.Cells[startRow, colIndex].Style.Numberformat.Format = "0.0%";
                                        worksheet.Cells[startRow, colIndex].Value = decimalValue / 100;
                                    }
                                    else
                                    {
                                        worksheet.Cells[startRow, colIndex].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    worksheet.Cells[startRow, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                }
                                else if (value is int || value is double || value is float)
                                {
                                    worksheet.Cells[startRow, colIndex].Style.Numberformat.Format = "#,##0";
                                    worksheet.Cells[startRow, colIndex].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"寫入單元格時出錯: 行={startRow}, 列={colIndex}, 屬性={col.Property}, 錯誤={ex.Message}");
                        }
                    }

                    // 添加邊框
                    worksheet.Cells[startRow, colIndex].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    colIndex++;
                }
                startRow++;
            }

            // 添加總計行
            if (firstItem is ProductSalesData)
            {
                worksheet.Cells[startRow, 1].Value = "Grand Total";
                worksheet.Cells[startRow, 1].Style.Font.Bold = true;
                worksheet.Cells[startRow, 2].Value = totalAgencyComm;
                worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[startRow, 3].Value = totalBuyResellComm;
                worksheet.Cells[startRow, 3].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[startRow, 4].Value = totalComm;
                worksheet.Cells[startRow, 4].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[startRow, 5].Value = totalPOValue;
                worksheet.Cells[startRow, 5].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[startRow, 6].Value = 1.0; // 100%
                worksheet.Cells[startRow, 6].Style.Numberformat.Format = "0.00%";

                // 設置背景顏色
                for (int i = 1; i <= 6; i++)
                {
                    worksheet.Cells[startRow, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[startRow, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                    worksheet.Cells[startRow, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }
                startRow++;
            }

            return startRow;
        }

        // 添加Sexy Report到工作表
        private int AddSexyReportToWorksheet(ExcelWorksheet worksheet, List<ProductSalesData> productDataList, int startRow)
        {
            // 表格標題
            worksheet.Cells[startRow, 1].Value = "Sexy Report - Who's POs are bigger";
            worksheet.Cells[startRow, 1, startRow, 6].Merge = true;
            worksheet.Cells[startRow, 1].Style.Font.Bold = true;
            worksheet.Cells[startRow, 1].Style.Font.Size = 12;
            startRow++;

            // 表格頭
            worksheet.Cells[startRow, 1].Value = "Product Type";
            worksheet.Cells[startRow, 1].Style.Font.Bold = true;
            worksheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            worksheet.Cells[startRow, 2].Value = "PO Value";
            worksheet.Cells[startRow, 2].Style.Font.Bold = true;
            worksheet.Cells[startRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[startRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            worksheet.Cells[startRow, 3].Value = "% of Grand Total";
            worksheet.Cells[startRow, 3].Style.Font.Bold = true;
            worksheet.Cells[startRow, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[startRow, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            // 設置邊框
            for (int i = 1; i <= 3; i++)
            {
                worksheet.Cells[startRow, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
            startRow++;

            // 按PO值排序的產品數據
            var sortedData = productDataList.OrderByDescending(p => p.POValue).ToList();
            decimal totalPOValue = sortedData.Sum(p => p.POValue);

            // 寫入數據
            foreach (var product in sortedData)
            {
                worksheet.Cells[startRow, 1].Value = product.ProductType;
                worksheet.Cells[startRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells[startRow, 2].Value = product.POValue;
                worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[startRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells[startRow, 3].Value = product.PercentageOfTotal / 100;
                worksheet.Cells[startRow, 3].Style.Numberformat.Format = "0.0%";
                worksheet.Cells[startRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                worksheet.Cells[startRow, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                startRow++;
            }

            // 添加總計行
            worksheet.Cells[startRow, 1].Value = "Grand Total";
            worksheet.Cells[startRow, 1].Style.Font.Bold = true;
            worksheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[startRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells[startRow, 2].Value = totalPOValue;
            worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells[startRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[startRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[startRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[startRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            worksheet.Cells[startRow, 3].Value = 1.0; // 100%
            worksheet.Cells[startRow, 3].Style.Numberformat.Format = "0.00%";
            worksheet.Cells[startRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            worksheet.Cells[startRow, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[startRow, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            worksheet.Cells[startRow, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);

            startRow++;
            return startRow;
        }

        /// <summary>
        /// 將數據匯出為PDF檔案
        /// </summary>
        /// <summary>
        /// 將數據匯出為PDF檔案
        /// </summary>
        /// <summary>
        /// 將數據匯出為PDF檔案
        /// </summary>
        public async Task<bool> ExportToPdfAsync<T>(IEnumerable<T> data, string fileName, string title)
        {
            if (data == null || !data.Any())
            {
                Debug.WriteLine("無數據可匯出");
                await ShowErrorMessage("匯出錯誤", "沒有資料可以匯出。請確保視圖中有顯示數據。");
                return false;
            }

            try
            {
                Debug.WriteLine($"開始匯出 PDF，數據項數: {data.Count()}");

                string exportPath = GetExportDirectory();
                string fullPath = Path.Combine(exportPath, $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");

                // 創建文件夾（如果不存在）
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath));

                // 這裡先不使用 iTextSharp，改用簡單的文本文件
                using (var writer = new StreamWriter(fullPath))
                {
                    await writer.WriteLineAsync($"*** {title} ***");
                    await writer.WriteLineAsync($"Generate date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    await writer.WriteLineAsync("");

                    var firstItem = data.FirstOrDefault();

                    if (firstItem is ProductSalesData)
                    {
                        await writer.WriteLineAsync("F25 - Sales Rep Commission by Product Type (Margin Achieved)");
                        await writer.WriteLineAsync("----------------------------------------------------");
                        await writer.WriteLineAsync("Product Type\tAgency Commission\tBuy Resell Commission\tTotal Commission\tPO Value\t% of Total");

                        decimal totalAgencyComm = 0;
                        decimal totalBuyResellComm = 0;
                        decimal totalComm = 0;
                        decimal totalPOValue = 0;

                        foreach (var item in data)
                        {
                            if (item is ProductSalesData product)
                            {
                                await writer.WriteLineAsync($"{product.ProductType}\t${product.AgencyMargin:N2}\t${product.BuyResellMargin:N2}\t${product.TotalMargin:N2}\t${product.POValue:N2}\t{product.PercentageOfTotal:N1}%");

                                totalAgencyComm += product.AgencyMargin;
                                totalBuyResellComm += product.BuyResellMargin;
                                totalComm += product.TotalMargin;
                                totalPOValue += product.POValue;
                            }
                        }

                        await writer.WriteLineAsync($"Grand Total\t${totalAgencyComm:N2}\t${totalBuyResellComm:N2}\t${totalComm:N2}\t${totalPOValue:N2}\t100.0%");

                        await writer.WriteLineAsync("");
                        await writer.WriteLineAsync("Sexy Report - Who's POs are bigger");
                        await writer.WriteLineAsync("-----------------------------");
                        await writer.WriteLineAsync("Product Type\tPO Value\t% of Grand Total");

                        var productDataList = data.Cast<ProductSalesData>().ToList();
                        var sortedData = productDataList.OrderByDescending(p => p.POValue).ToList();

                        foreach (var product in sortedData)
                        {
                            await writer.WriteLineAsync($"{product.ProductType}\t${product.POValue:N2}\t{product.PercentageOfTotal:N1}%");
                        }

                        await writer.WriteLineAsync($"Grand Total\t${totalPOValue:N2}\t100.0%");
                    }
                    else if (firstItem is SalesLeaderboardItem)
                    {
                        await writer.WriteLineAsync("Sales Representatives Performance");
                        await writer.WriteLineAsync("-----------------------------");
                        await writer.WriteLineAsync("Rank\tSales Rep\tAgency Commission\tBuy Resell Commission\tTotal Commission");

                        foreach (var item in data)
                        {
                            if (item is SalesLeaderboardItem rep)
                            {
                                await writer.WriteLineAsync($"{rep.Rank}\t{rep.SalesRep}\t${rep.AgencyMargin:N2}\t${rep.BuyResellMargin:N2}\t${rep.TotalMargin:N2}");
                            }
                        }
                    }
                }

                // 顯示成功消息，但說明這只是臨時 PDF 格式
                await ShowSuccessMessage("PDF報表已生成",
                    $"檔案已保存到: {fullPath}\n\n" +
                    "注意：目前PDF輸出採用簡化格式。未來版本將提供完整格式化的PDF。");

                // 嘗試打開文件夾
#if WINDOWS
        OpenFolder(exportPath);
#endif

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"匯出PDF時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await ShowErrorMessage("匯出錯誤", $"匯出PDF檔案時發生錯誤: {ex.Message}");
                return false;
            }
        }


        /// <summary>
        /// 將數據匯出為CSV檔案
        /// </summary>
        public async Task<bool> ExportToCsvAsync<T>(IEnumerable<T> data, string fileName)
        {
            if (data == null || !data.Any())
            {
                Debug.WriteLine("無數據可匯出");
                return false;
            }

            try
            {
                string exportPath = GetExportDirectory();
                string fullPath = Path.Combine(exportPath, $"{fileName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

                // 手動定義欄位名稱和對應屬性
                List<(string Header, string Property)> columns = new List<(string, string)>();

                // 根據類型添加適當欄位
                var firstItem = data.FirstOrDefault();
                if (firstItem is ProductSalesData)
                {
                    columns.Add(("Product Type", "ProductType"));
                    columns.Add(("Agency Commission", "AgencyCommission"));
                    columns.Add(("Buy Resell Commission", "BuyResellCommission"));
                    columns.Add(("Total Commission", "TotalCommission"));
                    columns.Add(("PO Value", "POValue"));
                    columns.Add(("% of Total", "PercentageOfTotal"));
                }
                else if (firstItem is SalesLeaderboardItem)
                {
                    columns.Add(("Rank", "Rank"));
                    columns.Add(("Sales Rep", "SalesRep"));
                    columns.Add(("Agency Commission", "AgencyCommission"));
                    columns.Add(("Buy Resell Commission", "BuyResellCommission"));
                    columns.Add(("Total Commission", "TotalCommission"));
                }

                using (var writer = new StreamWriter(fullPath, false, Encoding.UTF8))
                {
                    // 寫入表頭
                    var header = string.Join(",", columns.Select(c => $"\"{c.Header}\""));
                    await writer.WriteLineAsync(header);

                    // 寫入數據行
                    Type itemType = firstItem.GetType();
                    foreach (var item in data)
                    {
                        var values = new List<string>();
                        foreach (var col in columns)
                        {
                            var prop = itemType.GetProperty(col.Property);
                            if (prop != null)
                            {
                                var value = prop.GetValue(item);
                                string formattedValue = FormatCsvValue(value, prop.PropertyType);
                                values.Add(formattedValue);
                            }
                            else
                            {
                                values.Add("\"\"");
                            }
                        }
                        await writer.WriteLineAsync(string.Join(",", values));
                    }
                }

                // 顯示成功消息
                await ShowSuccessMessage("CSV報表已生成", $"檔案已保存到: {fullPath}");

                // 嘗試打開文件夾
#if WINDOWS
                OpenFolder(exportPath);
#endif

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"匯出CSV時發生錯誤: {ex.Message}");
                await ShowErrorMessage("匯出錯誤", $"匯出CSV檔案時發生錯誤: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 打印預覽報表
        /// </summary>
        public async Task<bool> PrintReportAsync<T>(IEnumerable<T> data, string title)
        {
            if (data == null || !data.Any())
            {
                Debug.WriteLine("無數據可打印");
                return false;
            }

            try
            {
                // 未實現打印功能
                await ShowSuccessMessage("打印功能未實現", "在該版本中，打印功能尚未完全實現。請考慮使用Excel匯出代替。");

                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"打印報表時發生錯誤: {ex.Message}");
                await ShowErrorMessage("打印錯誤", $"打印報表時發生錯誤: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 獲取匯出文件的目錄，如果不存在則創建
        /// </summary>
        public string GetExportDirectory()
        {
            string baseDir = string.Empty;

            // 依照平台獲取適當的文件夾
#if ANDROID
            baseDir = Path.Combine(Android.OS.Environment.ExternalStorageDirectory.AbsolutePath, Android.OS.Environment.DirectoryDownloads);
#elif IOS
            baseDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
#else
            // Windows、macOS等平台
            baseDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
#endif

            // 創建導出目錄
            string exportDir = Path.Combine(baseDir, "ScoreCard_Exports");
            if (!Directory.Exists(exportDir))
            {
                Directory.CreateDirectory(exportDir);
            }

            return exportDir;
        }

        #region 輔助方法

        /// <summary>
        /// 判斷屬性是否需要跳過（不需要匯出）
        /// </summary>
        private bool ShouldSkipProperty(string propertyName)
        {
            // 跳過不需要匯出的屬性，例如IsSelected標記或導航屬性
            string[] skipProperties = { "IsSelected", "Parent", "Children", "Id" };
            return skipProperties.Contains(propertyName);
        }

        /// <summary>
        /// 獲取屬性的顯示名稱
        /// </summary>
        private string GetDisplayName(string propertyName)
        {
            // 將屬性名稱轉換為更友好的顯示名稱
            // 例如: ProductType -> Product Type
            var mapping = new Dictionary<string, string>
            {
                { "ProductType", "Product Type" },
                { "SalesRep", "Sales Rep" },
                { "AgencyCommission", "Agency Commission" },
                { "BuyResellCommission", "Buy Resell Commission" },
                { "TotalCommission", "Total Commission" },
                { "POValue", "PO Value" },
                { "PercentageOfTotal", "% of Total" },
                { "Rank", "Rank" }
            };

            if (mapping.ContainsKey(propertyName))
            {
                return mapping[propertyName];
            }

            // 自動生成顯示名稱
            StringBuilder result = new StringBuilder();
            foreach (char c in propertyName)
            {
                if (char.IsUpper(c) && result.Length > 0)
                {
                    result.Append(' ');
                }
                result.Append(c);
            }
            return result.ToString();
        }

        /// <summary>
        /// 判斷是否為數字類型
        /// </summary>
        private bool IsNumericType(Type type)
        {
            return Type.GetTypeCode(Nullable.GetUnderlyingType(type) ?? type) switch
            {
                TypeCode.Decimal or
                TypeCode.Double or
                TypeCode.Single or
                TypeCode.Int32 or
                TypeCode.Int64 or
                TypeCode.Int16 => true,
                _ => false
            };
        }

        /// <summary>
        /// 格式化CSV值
        /// </summary>
        private string FormatCsvValue(object value, Type propertyType)
        {
            if (value == null)
                return "\"\"";

            // 處理特殊類型
            if (propertyType == typeof(DateTime))
            {
                return $"\"{((DateTime)value).ToString("yyyy-MM-dd")}\"";
            }
            else if (IsNumericType(propertyType))
            {
                // 數字類型不需要引號
                return value.ToString();
            }
            else
            {
                // 字符串類型，需要轉義引號
                return $"\"{value.ToString().Replace("\"", "\"\"")}\"";
            }
        }

        /// <summary>
        /// 嘗試打開文件夾
        /// </summary>
        private void OpenFolder(string folderPath)
        {
            try
            {
                // 僅在Windows平台上嘗試打開文件夾
                if (DeviceInfo.Platform == DevicePlatform.WinUI)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = folderPath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"嘗試打開文件夾時發生錯誤: {ex.Message}");
            }
        }

        /// <summary>
        /// 顯示成功消息
        /// </summary>
        private async Task ShowSuccessMessage(string title, string message)
        {
            await MainThread.InvokeOnMainThreadAsync(async () =>
            {
                await Application.Current.MainPage.DisplayAlert(title, message, "確定");
            });
        }

        /// <summary>
        /// 顯示錯誤消息
        /// </summary>
        private async Task ShowErrorMessage(string title, string message)
        {
            await MainThread.InvokeOnMainThreadAsync(async () =>
            {
                await Application.Current.MainPage.DisplayAlert(title, message, "確定");
            });
        }

        #endregion
    }
}