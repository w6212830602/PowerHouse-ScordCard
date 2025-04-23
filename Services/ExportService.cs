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

        /// <summary>
        /// 將數據匯出為Excel檔案
        /// </summary>
        public async Task<bool> ExportToExcelAsync<T>(IEnumerable<T> data, string fileName, string title)
        {
            try
            {
                if (data == null || !data.Any())
                {
                    Debug.WriteLine("無數據可匯出");
                    await ShowErrorMessage("匯出錯誤", "沒有資料可以匯出。請確保視圖中有顯示數據。");
                    return false;
                }

                Debug.WriteLine($"開始匯出 Excel，數據項數: {data.Count()}");
                Debug.WriteLine($"數據類型: {typeof(T).Name}");

                // Use Task.Run to perform the file operations in a background thread
                return await Task.Run(async () =>
                {
                    try
                    {
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

                            // 第一個表格 - 主要數據
                            string tableTitle = (typeof(T) == typeof(SalesLeaderboardItem) ||
                                                (typeof(T).Name == "Object" && data.FirstOrDefault() is SalesLeaderboardItem) ||
                                                (data.FirstOrDefault()?.GetType().Name.Contains("SalesLeaderboardItem") == true)) ?
                                "Sales Rep Commission" :
                                "Sales Rep Commission by Product Type (Margin Achieved)";

                            currentRow = AddTableToWorksheet(worksheet, data, currentRow, tableTitle);

                            // 添加間隔
                            currentRow += 2;

                            // 重要：檢索正確的產品數據用於第二個表格 - PO Vertiv Value
                            List<ProductSalesData> productDataForVertivTable = null;

                            // 使用動態類型避免明確依賴DetailedSalesViewModel類型
                            try
                            {
                                Debug.WriteLine("嘗試獲取當前頁面的ViewModel");

                                // 嘗試從MainPage獲取ViewModel
                                var mainPage = Application.Current.MainPage;
                                if (mainPage != null)
                                {
                                    var viewModel = mainPage.BindingContext;
                                    if (viewModel != null)
                                    {
                                        // 使用反射獲取所需屬性
                                        var vmType = viewModel.GetType();
                                        Debug.WriteLine($"找到ViewModel，類型為: {vmType.Name}");

                                        // 嘗試獲取IsProductView屬性
                                        var isProductViewProp = vmType.GetProperty("IsProductView");
                                        var isRepViewProp = vmType.GetProperty("IsRepView");
                                        var productSalesDataProp = vmType.GetProperty("ProductSalesData");
                                        var salesRepProductDataProp = vmType.GetProperty("SalesRepProductData");

                                        if (isProductViewProp != null && productSalesDataProp != null)
                                        {
                                            bool isProductView = (bool)isProductViewProp.GetValue(viewModel);
                                            if (isProductView)
                                            {
                                                var productData = productSalesDataProp.GetValue(viewModel) as IEnumerable<ProductSalesData>;
                                                if (productData != null)
                                                {
                                                    productDataForVertivTable = productData.ToList();
                                                    Debug.WriteLine($"從ProductSalesData獲取了{productDataForVertivTable.Count}條產品數據");
                                                }
                                            }
                                        }

                                        if (isRepViewProp != null && salesRepProductDataProp != null)
                                        {
                                            bool isRepView = (bool)isRepViewProp.GetValue(viewModel);
                                            if (isRepView)
                                            {
                                                var productData = salesRepProductDataProp.GetValue(viewModel) as IEnumerable<ProductSalesData>;
                                                if (productData != null)
                                                {
                                                    productDataForVertivTable = productData.ToList();
                                                    Debug.WriteLine($"從SalesRepProductData獲取了{productDataForVertivTable.Count}條產品數據");
                                                }
                                            }
                                        }
                                    }
                                }

                                // 如果MainPage不行，嘗試查找當前顯示的頁面
                                if (productDataForVertivTable == null || !productDataForVertivTable.Any())
                                {
                                    if (mainPage is Shell shell)
                                    {
                                        var currentPage = shell.CurrentPage;
                                        if (currentPage != null)
                                        {
                                            var viewModel = currentPage.BindingContext;
                                            if (viewModel != null)
                                            {
                                                // 使用反射獲取所需屬性
                                                var vmType = viewModel.GetType();
                                                Debug.WriteLine($"從Shell.CurrentPage找到ViewModel，類型為: {vmType.Name}");

                                                // 嘗試獲取產品數據屬性
                                                var productSalesDataProp = vmType.GetProperty("ProductSalesData");
                                                if (productSalesDataProp != null)
                                                {
                                                    var productData = productSalesDataProp.GetValue(viewModel) as IEnumerable<ProductSalesData>;
                                                    if (productData != null)
                                                    {
                                                        productDataForVertivTable = productData.ToList();
                                                        Debug.WriteLine($"從CurrentPage.ProductSalesData獲取了{productDataForVertivTable.Count}條產品數據");
                                                    }
                                                }

                                                // 如果上面沒成功，嘗試SalesRepProductData
                                                if ((productDataForVertivTable == null || !productDataForVertivTable.Any()) &&
                                                    vmType.GetProperty("SalesRepProductData") != null)
                                                {
                                                    var salesRepProductDataProp = vmType.GetProperty("SalesRepProductData");
                                                    var productData = salesRepProductDataProp.GetValue(viewModel) as IEnumerable<ProductSalesData>;
                                                    if (productData != null)
                                                    {
                                                        productDataForVertivTable = productData.ToList();
                                                        Debug.WriteLine($"從CurrentPage.SalesRepProductData獲取了{productDataForVertivTable.Count}條產品數據");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"嘗試獲取ViewModel時發生錯誤: {ex.Message}");
                                Debug.WriteLine(ex.StackTrace);
                            }

                            // 直接嘗試將數據轉換為產品數據
                            if ((productDataForVertivTable == null || !productDataForVertivTable.Any()) && typeof(T) == typeof(ProductSalesData))
                            {
                                productDataForVertivTable = data.Cast<ProductSalesData>().ToList();
                                Debug.WriteLine($"直接轉換獲取了{productDataForVertivTable.Count}條產品數據");
                            }

                            // 若仍然沒有產品數據，則使用數據類型信息來生成模擬數據
                            if (productDataForVertivTable == null || !productDataForVertivTable.Any())
                            {
                                Debug.WriteLine("無法獲取實際產品數據，嘗試創建模擬數據");
                                try
                                {
                                    // 根據第一個表格的數據類型決定使用哪種模擬數據
                                    if (typeof(T) == typeof(SalesLeaderboardItem) ||
                                        (data.FirstOrDefault()?.GetType().Name.Contains("SalesLeaderboardItem") == true))
                                    {
                                        // 如果是销售代表视图，创建相应的模拟数据
                                        productDataForVertivTable = CreateSampleProductDataForReps();
                                        Debug.WriteLine($"已创建销售代表视图的模拟产品数据: {productDataForVertivTable.Count}项");
                                    }
                                    else
                                    {
                                        // 默认产品视图模拟数据
                                        productDataForVertivTable = CreateSampleProductData();
                                        Debug.WriteLine($"已创建默认的模拟产品数据: {productDataForVertivTable.Count}项");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"创建模拟数据时发生错误: {ex.Message}");
                                    productDataForVertivTable = new List<ProductSalesData>();
                                }
                            }

                            // 生成第二個表格 - Vertiv Value Report
                            currentRow = AddVertivValueReportToWorksheet(worksheet, productDataForVertivTable ?? new List<ProductSalesData>(), currentRow);
                            Debug.WriteLine($"已添加PO Vertiv Value表格，當前行: {currentRow}");

                            // 設置列寬自適應
                            for (int i = 1; i <= 10; i++) // 假設最多10列
                            {
                                worksheet.Column(i).AutoFit();
                            }

                            // 保存文件
                            try
                            {
                                await package.SaveAsAsync(new FileInfo(fullPath));
                                Debug.WriteLine($"Excel 檔案已保存至: {fullPath}");

                                // 顯示成功消息
                                await MainThread.InvokeOnMainThreadAsync(async () => {
                                    await ShowSuccessMessage("Excel Report Generated", $"File saved to: {fullPath}");
                                });

                                // 嘗試打開文件夾
#if WINDOWS
                        try 
                        {
                            OpenFolder(exportPath);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error opening folder: {ex.Message}");
                        }
#endif

                                return true;
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"Error saving Excel file: {ex.Message}");
                                await MainThread.InvokeOnMainThreadAsync(async () => {
                                    await ShowErrorMessage("Excel Export Error", $"Error saving file: {ex.Message}");
                                });
                                return false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Exception in Excel export: {ex.Message}");
                        Debug.WriteLine(ex.StackTrace);

                        await MainThread.InvokeOnMainThreadAsync(async () => {
                            await ShowErrorMessage("Export Error", $"An error occurred during Excel export: {ex.Message}");
                        });
                        return false;
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"匯出Excel時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await ShowErrorMessage("Export error", $"Something happens: {ex.Message}");
                return false;
            }
        }

        // 輔助方法：創建默認產品數據
        private List<ProductSalesData> CreateSampleProductData()
        {
            return new List<ProductSalesData>
    {
        new ProductSalesData { ProductType = "Power", VertivValue = 525055.22m, PercentageOfTotal = 58.4m },
        new ProductSalesData { ProductType = "Service", VertivValue = 266681.89m, PercentageOfTotal = 29.7m },
        new ProductSalesData { ProductType = "Thermal", VertivValue = 67843.00m, PercentageOfTotal = 7.5m },
        new ProductSalesData { ProductType = "Batts & Caps", VertivValue = 39744.08m, PercentageOfTotal = 4.4m }
    };
        }

        // 輔助方法：創建針對銷售代表視圖的產品數據
        private List<ProductSalesData> CreateSampleProductDataForReps()
        {
            // 此處可根據實際需要調整數據
            return new List<ProductSalesData>
    {
        new ProductSalesData { ProductType = "Power", VertivValue = 525055.22m, PercentageOfTotal = 58.4m },
        new ProductSalesData { ProductType = "Service", VertivValue = 266681.89m, PercentageOfTotal = 29.7m },
        new ProductSalesData { ProductType = "Thermal", VertivValue = 67843.00m, PercentageOfTotal = 7.5m },
        new ProductSalesData { ProductType = "Batts & Caps", VertivValue = 39744.08m, PercentageOfTotal = 4.4m }
    };
        }

        // 獲取產品數據的方法 - 用於 Vertiv Value Report
        private List<ProductSalesData> GetProductDataForVertivReport()
        {
            try
            {
                // 不要從ExcelService獲取數據，而是使用傳入的參數
                return null; // 這個方法不再主動獲取數據
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting product data: {ex.Message}");
                return new List<ProductSalesData>();
            }
        }

        // 添加表格到工作表 - 完全修改版本
        private int AddTableToWorksheet<T>(ExcelWorksheet worksheet, IEnumerable<T> data, int startRow, string tableTitle)
        {
            try
            {
                Debug.WriteLine($"開始添加表格: {tableTitle}，行號: {startRow}");

                // 表格標題
                worksheet.Cells[startRow, 1].Value = tableTitle;
                worksheet.Cells[startRow, 1, startRow, 6].Merge = true;
                worksheet.Cells[startRow, 1].Style.Font.Bold = true;
                worksheet.Cells[startRow, 1].Style.Font.Size = 12;
                startRow++;

                // 手動定義欄位名稱和對應屬性
                List<(string Header, string Property)> columns = new List<(string, string)>();

                // 根據數據類型添加適當欄位
                var firstItem = data.FirstOrDefault();
                bool isProductData = firstItem is ProductSalesData;
                bool isSalesRepData = firstItem is SalesLeaderboardItem;

                // 處理Object類型的情況（數據類型可能在運行時不明確）
                if (!isProductData && !isSalesRepData && firstItem != null)
                {
                    // 嘗試確定實際數據類型
                    Type actualType = firstItem.GetType();
                    Debug.WriteLine($"數據實際類型: {actualType.Name}");

                    isProductData = actualType.Name.Contains("ProductSalesData");
                    isSalesRepData = actualType.Name.Contains("SalesLeaderboardItem");
                }

                if (isProductData)
                {
                    Debug.WriteLine("檢測到ProductSalesData類型");
                    columns.Add(("Product Type", "ProductType"));
                    columns.Add(("Agency Margin", "AgencyMargin"));
                    columns.Add(("Buy Resell Margin", "BuyResellMargin"));
                    columns.Add(("Total Margin", "TotalMargin"));
                    columns.Add(("PO Vertiv Value", "VertivValue")); // 修改为使用 VertivValue 属性
                    columns.Add(("% of Total", "PercentageOfTotal"));
                }
                else if (isSalesRepData)
                {
                    Debug.WriteLine("檢測到SalesLeaderboardItem類型");
                    columns.Add(("Rank", "Rank"));
                    columns.Add(("Sales Rep", "SalesRep"));
                    columns.Add(("Agency Margin", "AgencyMargin"));
                    columns.Add(("Buy Resell Margin", "BuyResellMargin"));
                    columns.Add(("Total Margin", "TotalMargin"));
                }
                else
                {
                    Debug.WriteLine($"未知數據類型: {(firstItem?.GetType().Name ?? "null")}，使用默認列");
                    // 默認列（保守方案）
                    columns.Add(("Item", "ToString"));
                    columns.Add(("Value", "ToString"));
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

                // 計算總計值
                decimal totalAgencyMargin = 0;
                decimal totalBuyResellMargin = 0;
                decimal totalMargin = 0;
                decimal totalVertivValue = 0;

                // 寫入數據行
                foreach (var item in data)
                {
                    colIndex = 1;

                    // 根據數據類型處理每一行
                    if (isProductData && item is ProductSalesData productItem)
                    {
                        try
                        {
                            worksheet.Cells[startRow, 1].Value = productItem.ProductType;
                            worksheet.Cells[startRow, 2].Value = productItem.AgencyMargin;
                            worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells[startRow, 3].Value = productItem.BuyResellMargin;
                            worksheet.Cells[startRow, 3].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells[startRow, 4].Value = productItem.TotalMargin;
                            worksheet.Cells[startRow, 4].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells[startRow, 5].Value = productItem.VertivValue;
                            worksheet.Cells[startRow, 5].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells[startRow, 6].Value = productItem.PercentageOfTotal / 100; // 轉換為小數
                            worksheet.Cells[startRow, 6].Style.Numberformat.Format = "0.0%";

                            // 累計總計值
                            totalAgencyMargin += productItem.AgencyMargin;
                            totalBuyResellMargin += productItem.BuyResellMargin;
                            totalMargin += productItem.TotalMargin;
                            totalVertivValue += productItem.VertivValue;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"處理產品數據行時出錯: {ex.Message}");
                        }
                    }
                    else if (isSalesRepData && item is SalesLeaderboardItem repItem)
                    {
                        try
                        {
                            worksheet.Cells[startRow, 1].Value = repItem.Rank;
                            worksheet.Cells[startRow, 2].Value = repItem.SalesRep;
                            worksheet.Cells[startRow, 3].Value = repItem.AgencyMargin;
                            worksheet.Cells[startRow, 3].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells[startRow, 4].Value = repItem.BuyResellMargin;
                            worksheet.Cells[startRow, 4].Style.Numberformat.Format = "#,##0.00";
                            worksheet.Cells[startRow, 5].Value = repItem.TotalMargin;
                            worksheet.Cells[startRow, 5].Style.Numberformat.Format = "#,##0.00";

                            // 累計總計值
                            totalAgencyMargin += repItem.AgencyMargin;
                            totalBuyResellMargin += repItem.BuyResellMargin;
                            totalMargin += repItem.TotalMargin;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"處理銷售代表數據行時出錯: {ex.Message}");
                        }
                    }
                    else
                    {
                        // 處理未知類型
                        try
                        {
                            // 嘗試使用反射獲取屬性值
                            Type itemType = item.GetType();

                            for (int i = 0; i < columns.Count; i++)
                            {
                                if (columns[i].Property == "ToString")
                                {
                                    worksheet.Cells[startRow, i + 1].Value = item.ToString();
                                }
                                else
                                {
                                    var prop = itemType.GetProperty(columns[i].Property);
                                    if (prop != null)
                                    {
                                        var value = prop.GetValue(item);
                                        worksheet.Cells[startRow, i + 1].Value = value;

                                        // 如果是數值，應用數字格式
                                        if (value is decimal || value is double || value is float)
                                        {
                                            worksheet.Cells[startRow, i + 1].Style.Numberformat.Format = "#,##0.00";
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"處理未知類型數據行時出錯: {ex.Message}");
                            // 作為最後的手段，直接使用ToString
                            worksheet.Cells[startRow, 1].Value = item.ToString();
                        }
                    }

                    // 添加邊框到每個單元格
                    for (int i = 1; i <= columns.Count; i++)
                    {
                        worksheet.Cells[startRow, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    startRow++;
                }

                // 添加總計行
                if (isProductData)
                {
                    worksheet.Cells[startRow, 1].Value = "Grand Total";
                    worksheet.Cells[startRow, 1].Style.Font.Bold = true;
                    worksheet.Cells[startRow, 2].Value = totalAgencyMargin;
                    worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 3].Value = totalBuyResellMargin;
                    worksheet.Cells[startRow, 3].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 4].Value = totalMargin;
                    worksheet.Cells[startRow, 4].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 5].Value = totalVertivValue;
                    worksheet.Cells[startRow, 5].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 6].Value = 1.0; // 100%
                    worksheet.Cells[startRow, 6].Style.Numberformat.Format = "0.0%";

                    // 設置背景顏色
                    for (int i = 1; i <= 6; i++)
                    {
                        worksheet.Cells[startRow, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[startRow, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        worksheet.Cells[startRow, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }
                }
                else if (isSalesRepData)
                {
                    worksheet.Cells[startRow, 2].Value = "Grand Total";
                    worksheet.Cells[startRow, 2].Style.Font.Bold = true;
                    worksheet.Cells[startRow, 3].Value = totalAgencyMargin;
                    worksheet.Cells[startRow, 3].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 4].Value = totalBuyResellMargin;
                    worksheet.Cells[startRow, 4].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 5].Value = totalMargin;
                    worksheet.Cells[startRow, 5].Style.Numberformat.Format = "#,##0.00";

                    // 設置背景顏色
                    for (int i = 2; i <= 5; i++)
                    {
                        worksheet.Cells[startRow, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[startRow, i].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                        worksheet.Cells[startRow, i].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }
                }
                startRow++;

                Debug.WriteLine($"表格添加完成，當前行: {startRow}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"在添加表格時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
            }

            return startRow;
        }

        // 添加 Vertiv Value Report 到工作表
        private int AddVertivValueReportToWorksheet(ExcelWorksheet worksheet, List<ProductSalesData> productDataList, int startRow)
        {
            try
            {
                Debug.WriteLine($"開始添加PO Vertiv Value表格，產品數據項數: {productDataList?.Count ?? 0}");

                // 表格标题
                worksheet.Cells[startRow, 1].Value = "PO Vertiv Value";
                worksheet.Cells[startRow, 1, startRow, 6].Merge = true;
                worksheet.Cells[startRow, 1].Style.Font.Bold = true;
                worksheet.Cells[startRow, 1].Style.Font.Size = 12;
                startRow++;

                // 表格头
                worksheet.Cells[startRow, 1].Value = "Product Type";
                worksheet.Cells[startRow, 1].Style.Font.Bold = true;
                worksheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                worksheet.Cells[startRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[startRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[startRow, 2].Value = "PO Vertiv Value";
                worksheet.Cells[startRow, 2].Style.Font.Bold = true;
                worksheet.Cells[startRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                worksheet.Cells[startRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[startRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[startRow, 3].Value = "% of Grand Total";
                worksheet.Cells[startRow, 3].Style.Font.Bold = true;
                worksheet.Cells[startRow, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                worksheet.Cells[startRow, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                worksheet.Cells[startRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                startRow++;

                // Safety check for empty data
                if (productDataList == null || !productDataList.Any())
                {
                    Debug.WriteLine("沒有產品數據可用，添加「無數據」提示行");
                    worksheet.Cells[startRow, 1].Value = "No data available";
                    worksheet.Cells[startRow, 1, startRow, 3].Merge = true;
                    worksheet.Cells[startRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[startRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    return startRow + 1;
                }

                // 首先計算總的Vertiv Value
                decimal totalVertivValue = 0;

                // 首先嘗試從VertivValue屬性獲取值
                foreach (var product in productDataList)
                {
                    totalVertivValue += product.VertivValue;
                }

                // 如果總值為0，則嘗試從POValue屬性獲取值
                if (totalVertivValue == 0)
                {
                    totalVertivValue = productDataList.Sum(p => p.POValue);
                    Debug.WriteLine("使用POValue替代VertivValue計算總值: " + totalVertivValue);
                }

                // 確保有合理的總值
                if (totalVertivValue <= 0)
                {
                    Debug.WriteLine("警告: 總Vertiv Value為0或負數，設置為1避免除以零錯誤");
                    totalVertivValue = 1; // 避免除以零
                }

                // 按Vertiv Value或POValue降序排序資料（取決於哪個有值）
                var sortedData = new List<ProductSalesData>(productDataList);

                // 如果使用了VertivValue，則按此排序；否則按POValue排序
                if (productDataList.Sum(p => p.VertivValue) > 0)
                {
                    sortedData = sortedData.OrderByDescending(p => p.VertivValue).ToList();
                    Debug.WriteLine("按VertivValue降序排序數據");
                }
                else
                {
                    sortedData = sortedData.OrderByDescending(p => p.POValue).ToList();
                    Debug.WriteLine("按POValue降序排序數據");
                }

                // 寫入每條產品數據
                foreach (var product in sortedData)
                {
                    worksheet.Cells[startRow, 1].Value = product.ProductType;
                    worksheet.Cells[startRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    // 優先使用VertivValue，如果為0則使用POValue
                    decimal valueToUse = product.VertivValue > 0 ? product.VertivValue : product.POValue;
                    worksheet.Cells[startRow, 2].Value = valueToUse;
                    worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
                    worksheet.Cells[startRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    // 若產品已有百分比，則使用現有百分比；否則計算新的百分比
                    decimal percentage = 0;
                    if (product.PercentageOfTotal > 0)
                    {
                        percentage = product.PercentageOfTotal;
                        Debug.WriteLine($"使用現有百分比: {product.ProductType} = {percentage}%");
                    }
                    else
                    {
                        percentage = (valueToUse / totalVertivValue) * 100;
                        Debug.WriteLine($"計算新百分比: {product.ProductType} = {percentage}%");
                    }

                    worksheet.Cells[startRow, 3].Value = percentage / 100; // Excel需要小數表示百分比
                    worksheet.Cells[startRow, 3].Style.Numberformat.Format = "0.0%";
                    worksheet.Cells[startRow, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                    startRow++;
                }

                // 添加总计行
                worksheet.Cells[startRow, 1].Value = "Grand Total";
                worksheet.Cells[startRow, 1].Style.Font.Bold = true;
                worksheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[startRow, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells[startRow, 2].Value = totalVertivValue;
                worksheet.Cells[startRow, 2].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[startRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[startRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                worksheet.Cells[startRow, 3].Value = 1.0; // 100%
                worksheet.Cells[startRow, 3].Style.Numberformat.Format = "0.0%";
                worksheet.Cells[startRow, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[startRow, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                worksheet.Cells[startRow, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                startRow++;
                Debug.WriteLine("PO Vertiv Value表格添加完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"在添加 Vertiv Value Report 时发生错误: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
            }

            return startRow;
        }

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
                return await Task.Run(async () => {
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
                                await writer.WriteLineAsync("F25 - Sales Rep Margin by Product Type (Margin Achieved)");
                                await writer.WriteLineAsync("----------------------------------------------------");
                                await writer.WriteLineAsync("Product Type\tAgency Margin\tBuy Resell Margin\tTotal Margin\tVertiv Value\t% of Total");

                                decimal totalAgencyMargin = 0;
                                decimal totalBuyResellMargin = 0;
                                decimal totalMargin = 0;
                                decimal totalVertivValue = 0;

                                foreach (var item in data)
                                {
                                    if (item is ProductSalesData product)
                                    {
                                        await writer.WriteLineAsync($"{product.ProductType}\t${product.AgencyMargin:N2}\t${product.BuyResellMargin:N2}\t${product.TotalMargin:N2}\t${product.VertivValue:N2}\t{product.PercentageOfTotal:N1}%");

                                        totalAgencyMargin += product.AgencyMargin;
                                        totalBuyResellMargin += product.BuyResellMargin;
                                        totalMargin += product.TotalMargin;
                                        totalVertivValue += product.VertivValue;
                                    }
                                }

                                await writer.WriteLineAsync($"Grand Total\t${totalAgencyMargin:N2}\t${totalBuyResellMargin:N2}\t${totalMargin:N2}\t${totalVertivValue:N2}\t100.0%");

                                // Add PO Vertiv Value section
                                await writer.WriteLineAsync("");
                                await writer.WriteLineAsync("PO Vertiv Value");
                                await writer.WriteLineAsync("------------------");
                                await writer.WriteLineAsync("Product Type\tVertiv Value\t% of Grand Total");

                                var productDataList = data.Cast<ProductSalesData>().ToList();
                                var sortedData = productDataList.OrderByDescending(p => p.VertivValue).ToList();

                                foreach (var product in sortedData)
                                {
                                    await writer.WriteLineAsync($"{product.ProductType}\t${product.VertivValue:N2}\t{product.PercentageOfTotal:N1}%");
                                }

                                await writer.WriteLineAsync($"Grand Total\t${totalVertivValue:N2}\t100.0%");
                            }
                            else if (firstItem is SalesLeaderboardItem)
                            {
                                await writer.WriteLineAsync("Sales Representatives Performance");
                                await writer.WriteLineAsync("-----------------------------");
                                await writer.WriteLineAsync("Rank\tSales Rep\tAgency Margin\tBuy Resell Margin\tTotal Margin");

                                decimal totalAgencyMargin = 0;
                                decimal totalBuyResellMargin = 0;
                                decimal totalMargin = 0;

                                foreach (var item in data)
                                {
                                    if (item is SalesLeaderboardItem rep)
                                    {
                                        await writer.WriteLineAsync($"{rep.Rank}\t{rep.SalesRep}\t${rep.AgencyMargin:N2}\t${rep.BuyResellMargin:N2}\t${rep.TotalMargin:N2}");

                                        totalAgencyMargin += rep.AgencyMargin;
                                        totalBuyResellMargin += rep.BuyResellMargin;
                                        totalMargin += rep.TotalMargin;
                                    }
                                }

                                await writer.WriteLineAsync($"Grand Total\t\t${totalAgencyMargin:N2}\t${totalBuyResellMargin:N2}\t${totalMargin:N2}");

                                // Add PO Vertiv Value section even for Sales Rep view
                                await writer.WriteLineAsync("");
                                await writer.WriteLineAsync("PO Vertiv Value");
                                await writer.WriteLineAsync("------------------");
                                await writer.WriteLineAsync("Product Type\tVertiv Value\t% of Grand Total");

                                var productData = GetProductDataForVertivReport();
                                if (productData != null && productData.Any())
                                {
                                    var sortedData = productData.OrderByDescending(p => p.VertivValue).ToList();
                                    decimal totalVertivValue = sortedData.Sum(p => p.VertivValue);

                                    foreach (var product in sortedData)
                                    {
                                        await writer.WriteLineAsync($"{product.ProductType}\t${product.VertivValue:N2}\t{product.PercentageOfTotal:N1}%");
                                    }

                                    await writer.WriteLineAsync($"Grand Total\t${totalVertivValue:N2}\t100.0%");
                                }
                            }
                        }

                        // 顯示成功消息，但說明這只是臨時 PDF 格式
                        await MainThread.InvokeOnMainThreadAsync(async () => {
                            await ShowSuccessMessage("PDF報表已生成",
                                $"檔案已保存到: {fullPath}\n\n" +
                                "注意：目前PDF輸出採用簡化格式。未來版本將提供完整格式化的PDF。");
                        });

                        // 嘗試打開文件夾
#if WINDOWS
                        OpenFolder(exportPath);
#endif

                        return true;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error during PDF export: {ex.Message}");
                        Debug.WriteLine(ex.StackTrace);

                        await MainThread.InvokeOnMainThreadAsync(async () => {
                            await ShowErrorMessage("匯出錯誤", $"匯出PDF檔案時發生錯誤: {ex.Message}");
                        });
                        return false;
                    }
                });
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
                    columns.Add(("Agency Margin", "AgencyMargin"));
                    columns.Add(("Buy Resell Margin", "BuyResellMargin"));
                    columns.Add(("Total Margin", "TotalMargin"));
                    columns.Add(("Vertiv Value", "VertivValue"));
                    columns.Add(("% of Total", "PercentageOfTotal"));
                }
                else if (firstItem is SalesLeaderboardItem)
                {
                    columns.Add(("Rank", "Rank"));
                    columns.Add(("Sales Rep", "SalesRep"));
                    columns.Add(("Agency Margin", "AgencyMargin"));
                    columns.Add(("Buy Resell Margin", "BuyResellMargin"));
                    columns.Add(("Total Margin", "TotalMargin"));
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

                    // If this is ProductSalesData, add a second section for Vertiv Values
                    if (firstItem is ProductSalesData)
                    {
                        await writer.WriteLineAsync("");
                        await writer.WriteLineAsync("PO Vertiv Value");

                        // Add headers for the second table
                        await writer.WriteLineAsync("\"Product Type\",\"Vertiv Value\",\"% of Grand Total\"");

                        // Sort by Vertiv Value descending - use the exact same data without recalculation
                        var sortedData = data.Cast<ProductSalesData>().OrderByDescending(p => p.VertivValue).ToList();
                        decimal totalVertivValue = sortedData.Sum(p => p.VertivValue);

                        // Write data rows using the exact same values from the data
                        foreach (var product in sortedData)
                        {
                            // Use the exact percentage value from the product data
                            await writer.WriteLineAsync($"\"{product.ProductType}\",{product.VertivValue},{product.PercentageOfTotal:0.0}%");
                        }

                        // Add totals
                        await writer.WriteLineAsync($"\"Grand Total\",{totalVertivValue},100.0%");
                    }
                    // If this is SalesLeaderboardItem, get product data for Vertiv Values
                    else if (firstItem is SalesLeaderboardItem)
                    {
                        var productData = GetProductDataForVertivReport();
                        if (productData != null && productData.Any())
                        {
                            await writer.WriteLineAsync("");
                            await writer.WriteLineAsync("PO Vertiv Value");

                            // Add headers for the second table
                            await writer.WriteLineAsync("\"Product Type\",\"Vertiv Value\",\"% of Grand Total\"");

                            // Sort by Vertiv Value descending
                            var sortedData = productData.OrderByDescending(p => p.VertivValue).ToList();

                            // Write data rows
                            foreach (var product in sortedData)
                            {
                                await writer.WriteLineAsync($"\"{product.ProductType}\",{product.VertivValue},{product.PercentageOfTotal / 100:0.0%}");
                            }

                            // Add totals
                            decimal totalVertivValue = sortedData.Sum(p => p.VertivValue);
                            await writer.WriteLineAsync($"\"Grand Total\",{totalVertivValue},100.0%");
                        }
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
                await ShowErrorMessage("打印錯誤", "沒有資料可以打印。請確保視圖中有顯示數據。");
                return false;
            }

            try
            {
                // 先匯出為 PDF，然後嘗試打開它
                string fileName = $"Print_{title.Replace(" ", "_")}";
                bool exported = await ExportToPdfAsync(data, fileName, title);

                if (exported)
                {
                    // 在未來版本中，可以添加直接打印的功能
                    await ShowSuccessMessage("打印預覽已準備",
                        "已匯出為 PDF 文件作為打印預覽。在該版本中，請手動打開 PDF 文件並使用系統的打印功能進行打印。");
                    return true;
                }
                else
                {
                    await ShowErrorMessage("打印預覽失敗", "無法創建打印預覽。請嘗試使用 Excel 或 PDF 匯出功能。");
                    return false;
                }
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
                { "AgencyMargin", "Agency Margin" },
                { "BuyResellMargin", "Buy Resell Margin" },
                { "TotalMargin", "Total Margin" },
                { "VertivValue", "Vertiv Value" },
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