using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ScoreCard.Models;
using ScoreCard.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace ScoreCard.ViewModels
{
    public partial class DetailedSalesViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;
        private readonly IExportService _exportService;
        private List<SalesData> _allSalesData;
        private List<SalesData> _filteredSalesData;

        #region Properties

        [ObservableProperty]
        private DateTime _startDate = DateTime.Now.AddDays(-30);  // 預設為最近30天

        [ObservableProperty]
        private DateTime _endDate = DateTime.Now;

        [ObservableProperty]
        private bool _isLoading;

        [ObservableProperty]
        private string _viewType = "ByProduct";

        [ObservableProperty]
        private ObservableCollection<ProductSalesData> _productSalesData = new();

        [ObservableProperty]
        private ObservableCollection<SalesLeaderboardItem> _salesRepData = new();

        [ObservableProperty]
        private ObservableCollection<ProductSalesData> _salesRepProductData = new();

        [ObservableProperty]
        private ObservableCollection<string> _selectedSalesReps = new();

        [ObservableProperty]
        private ObservableCollection<string> _availableSalesReps = new();

        [ObservableProperty]
        private bool _isExportOptionsVisible;

        // 控制銷售代表選擇彈出視窗的顯示
        [ObservableProperty]
        private bool _isRepSelectionPopupVisible = false;

        // 新增：跟踪銷售代表選擇狀態
        [ObservableProperty]
        private ObservableCollection<RepSelectionItem> _repSelectionItems = new();

        // 臨時存儲選擇狀態
        private ObservableCollection<RepSelectionItem> _tempRepSelectionItems = new();

        // 計算屬性 - 用於繫結到 UI
        public bool IsProductView => ViewType == "ByProduct";
        public bool IsRepView => ViewType == "ByRep";

        // 顯示選中銷售代表的文本
        public string SelectedRepsText => SelectedSalesReps.Count == 0 ||
                                         (SelectedSalesReps.Count == 1 && SelectedSalesReps[0] == "All Reps")
                                         ? "All Reps"
                                         : $"{SelectedSalesReps.Count} selected";

        #endregion

        #region Property Changed Methods

        // 每當 ViewType 變更時更新 UI 屬性
        partial void OnViewTypeChanged(string value)
        {
            Debug.WriteLine($"視圖類型變更為: {value}");
            OnPropertyChanged(nameof(IsProductView));
            OnPropertyChanged(nameof(IsRepView));

            // 清除緩存並重新載入數據
            _excelService.ClearCache();
            FilterDataByDateRange();
            LoadFilteredData();
        }

        // 當 StartDate 變更時觸發的方法
        partial void OnStartDateChanged(DateTime value)
        {
            Debug.WriteLine($"開始日期變更為: {value:yyyy-MM-dd}");
            // 確保日期範圍有效
            if (value <= EndDate)
            {
                // 使用 Task.Run 異步執行，避免阻塞 UI 線程
                Task.Run(() => FilterDataAndReload());
            }
            else
            {
                // 如果開始日期比結束日期晚，自動調整結束日期
                EndDate = value;
            }
        }

        // 當 EndDate 變更時觸發的方法
        partial void OnEndDateChanged(DateTime value)
        {
            Debug.WriteLine($"結束日期變更為: {value:yyyy-MM-dd}");
            // 確保日期範圍有效
            if (value >= StartDate)
            {
                // 使用 Task.Run 異步執行，避免阻塞 UI 線程
                Task.Run(() => FilterDataAndReload());
            }
            else
            {
                // 如果結束日期比開始日期早，自動調整開始日期
                StartDate = value;
            }
        }

        // 當 SelectedSalesReps 變更時觸發的方法
        partial void OnSelectedSalesRepsChanged(ObservableCollection<string> value)
        {
            Debug.WriteLine($"選擇的銷售代表變更為: {string.Join(", ", value)}");
            OnPropertyChanged(nameof(SelectedRepsText));
        }

        #endregion

        #region Commands

        [RelayCommand]
        private async Task FilterData()
        {
            Debug.WriteLine("執行過濾...");
            try
            {
                IsLoading = true;
                await Task.Delay(300);  // 添加延遲以顯示載入效果
                FilterDataByDateRange();
                LoadFilteredData();
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        private void ChangeView(string viewType)
        {
            Debug.WriteLine($"切換視圖至: {viewType}，當前視圖: {ViewType}");
            if (ViewType != viewType)
            {
                ViewType = viewType;
                Debug.WriteLine($"視圖已切換到: {ViewType}");
            }
            else
            {
                Debug.WriteLine("視圖沒有變化，手動觸發數據重載");
                // 即使視圖沒有變化，也強制重載數據
                _excelService.ClearCache();
                FilterDataByDateRange();
                LoadFilteredData();
            }
        }

        [RelayCommand]
        private void ToggleExportOptions()
        {
            IsExportOptionsVisible = !IsExportOptionsVisible;
            Debug.WriteLine($"匯出選項顯示狀態: {IsExportOptionsVisible}");
        }

        public List<ProductSalesData> GetProductDataForExport()
        {
            // 不管當前視圖是什麼，總是返回產品數據
            return ProductSalesData.ToList();
        }


        [RelayCommand]
        private async Task Export(string format)
        {
            IsExportOptionsVisible = false;
            Debug.WriteLine($"匯出格式: {format}");

            try
            {
                IsLoading = true;
                await Task.Delay(300); // 添加短暫延遲以顯示載入效果

                // 準備要匯出的數據
                var exportData = GetDataToExport();
                if (exportData == null || !exportData.Any())
                {
                    await Application.Current.MainPage.DisplayAlert("無數據", "沒有可匯出的數據", "確定");
                    return;
                }

                // 生成適當的文件名
                string exportFileName = GenerateFileName();
                string exportTitle = GenerateReportTitle();

                bool exportSuccess = false;
                switch (format.ToLower())
                {
                    case "excel":
                        exportSuccess = await _exportService.ExportToExcelAsync(exportData, exportFileName, exportTitle);
                        break;
                    case "pdf":
                        exportSuccess = await _exportService.ExportToPdfAsync(exportData, exportFileName, exportTitle);
                        break;
                    case "csv":
                        exportSuccess = await _exportService.ExportToCsvAsync(exportData, exportFileName);
                        break;
                    case "print":
                        exportSuccess = await _exportService.PrintReportAsync(exportData, exportTitle);
                        break;
                    default:
                        await Application.Current.MainPage.DisplayAlert("不支持的格式", $"不支持的匯出格式: {format}", "確定");
                        break;
                }

                if (exportSuccess)
                {
                    Debug.WriteLine($"成功匯出為 {format} 格式");
                }
                else
                {
                    Debug.WriteLine($"匯出 {format} 格式失敗");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"匯出時發生錯誤: {ex.Message}");
                await Application.Current.MainPage.DisplayAlert("匯出錯誤", $"匯出時發生錯誤: {ex.Message}", "確定");
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// 獲取要匯出的數據
        /// </summary>
        private IEnumerable<object> GetDataToExport()
        {
            Debug.WriteLine($"準備匯出數據，當前視圖類型: {ViewType}");

            // 根據當前視圖類型返回相應的數據
            if (IsProductView && ProductSalesData.Any())
            {
                var data = ProductSalesData.ToList();
                Debug.WriteLine($"匯出產品視圖數據，共 {data.Count} 項");

                // 打印出前幾條數據的內容進行檢查
                for (int i = 0; i < Math.Min(data.Count, 3); i++)
                {
                    var item = data[i];
                    Debug.WriteLine($"樣本數據 {i + 1}: " +
                        $"ProductType={item.ProductType}, " +
                        $"AgencyMargin={item.AgencyMargin}, " +
                        $"BuyResellMargin={item.BuyResellMargin}, " +
                        $"TotalMargin={item.TotalMargin}, " +
                        $"POValue={item.POValue}, " +
                        $"PercentageOfTotal={item.PercentageOfTotal}");
                }
                return data;
            }
            else if (IsRepView && SalesRepData.Any())
            {
                var data = SalesRepData.ToList();
                Debug.WriteLine($"匯出銷售代表視圖數據，共 {data.Count} 項");
                return data;
            }

            Debug.WriteLine("沒有找到可匯出的數據");
            return new List<object>();
        }

        /// <summary>
        /// 生成匯出的文件名
        /// </summary>
        private string GenerateFileName()
        {
            string viewTypeText = IsProductView ? "ByProduct" : "ByRep";
            string dateRangeText = $"{StartDate:yyyyMMdd}_to_{EndDate:yyyyMMdd}";
            return $"SalesAnalysis_{viewTypeText}_{dateRangeText}";
        }

        /// <summary>
        /// 生成報表標題
        /// </summary>
        private string GenerateReportTitle()
        {
            string viewTypeText = IsProductView ? "By Product" : "By Sales Rep";
            string dateRangeText = $"{StartDate:yyyy-MM-dd} to {EndDate:yyyy-MM-dd}";
            return $"Sales Analysis Report - {viewTypeText} ({dateRangeText})";
        }

        [RelayCommand]
        private async Task NavigateToSummary()
        {
            try
            {
                await Shell.Current.GoToAsync("//SalesAnalysis");
                Debug.WriteLine("導航到摘要視圖");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"導航錯誤: {ex.Message}");
                await Application.Current.MainPage.DisplayAlert(
                    "導航錯誤",
                    $"無法導航到摘要頁面: {ex.Message}",
                    "確定");
            }
        }

        // 新增命令：處理銷售代表選擇變更
        [RelayCommand]
        private void ToggleRepSelection(RepSelectionItem item)
        {
            if (item == null) return;

            Debug.WriteLine($"切換銷售代表選擇: {item.Name}, 之前狀態: {item.IsSelected}");

            // 處理"All Reps"的特殊情況
            if (item.Name == "All Reps")
            {
                // 如果選中"All Reps"，取消選中其他所有項
                if (!item.IsSelected)
                {
                    item.IsSelected = true;

                    // 更新其他項的選中狀態
                    foreach (var otherItem in RepSelectionItems)
                    {
                        if (otherItem.Name != "All Reps")
                        {
                            otherItem.IsSelected = false;
                        }
                    }
                }
            }
            else
            {
                // 切換當前項的選中狀態
                item.IsSelected = !item.IsSelected;

                // 如果選中了非"All Reps"項，取消選中"All Reps"
                if (item.IsSelected)
                {
                    var allRepsItem = RepSelectionItems.FirstOrDefault(x => x.Name == "All Reps");
                    if (allRepsItem != null && allRepsItem.IsSelected)
                    {
                        allRepsItem.IsSelected = false;
                    }
                }

                // 如果沒有選中項，則自動選中"All Reps"
                if (!RepSelectionItems.Any(x => x.Name != "All Reps" && x.IsSelected))
                {
                    var allRepsItem = RepSelectionItems.FirstOrDefault(x => x.Name == "All Reps");
                    if (allRepsItem != null)
                    {
                        allRepsItem.IsSelected = true;
                    }
                }
            }

            // 注意：在彈出視窗中，我們不立即過濾數據，而是等待用戶點擊Apply按鈕
            // 這樣可以避免每次點擊都重新載入數據，提高用戶體驗
        }

        // 顯示銷售代表選擇彈出視窗
        [RelayCommand]
        private void ToggleRepSelectionPopup()
        {
            if (!IsRepSelectionPopupVisible)
            {
                // 顯示彈出視窗前，備份當前選擇狀態
                _tempRepSelectionItems = new ObservableCollection<RepSelectionItem>();
                foreach (var item in RepSelectionItems)
                {
                    _tempRepSelectionItems.Add(new RepSelectionItem
                    {
                        Name = item.Name,
                        IsSelected = item.IsSelected
                    });
                }
            }

            IsRepSelectionPopupVisible = !IsRepSelectionPopupVisible;
            Debug.WriteLine($"銷售代表選擇彈出視窗 {(IsRepSelectionPopupVisible ? "已顯示" : "已隱藏")}");
        }

        // 關閉彈出視窗並撤銷更改
        [RelayCommand]
        private void CloseRepSelectionPopup()
        {
            // 還原選擇狀態
            if (_tempRepSelectionItems.Any())
            {
                RepSelectionItems.Clear();
                foreach (var item in _tempRepSelectionItems)
                {
                    RepSelectionItems.Add(item);
                }
            }

            IsRepSelectionPopupVisible = false;
            Debug.WriteLine("取消更改，關閉彈出視窗");
        }

        // 應用選擇並關閉彈出視窗
        [RelayCommand]
        private void ApplyRepSelection()
        {
            // 更新 SelectedSalesReps 集合
            SelectedSalesReps.Clear();
            foreach (var item in RepSelectionItems)
            {
                if (item.IsSelected)
                {
                    SelectedSalesReps.Add(item.Name);
                }
            }

            // 如果沒有選中項，默認選中"All Reps"
            if (!SelectedSalesReps.Any())
            {
                var allRepsItem = RepSelectionItems.FirstOrDefault(x => x.Name == "All Reps");
                if (allRepsItem != null)
                {
                    allRepsItem.IsSelected = true;
                    SelectedSalesReps.Add("All Reps");
                }
            }

            IsRepSelectionPopupVisible = false;
            Debug.WriteLine($"應用選擇，已選: {string.Join(", ", SelectedSalesReps)}");
            OnPropertyChanged(nameof(SelectedRepsText));

            // 清除緩存並重新過濾數據
            _excelService.ClearCache();

            // 重要：清除現有產品數據，避免在重新加載前顯示舊數據
            MainThread.BeginInvokeOnMainThread(() =>
            {
                // 清空當前顯示的數據
                SalesRepProductData = new ObservableCollection<ProductSalesData>();
            });

            MainThread.BeginInvokeOnMainThread(async () =>
            {
                try
                {
                    IsLoading = true;
                    await FilterDataCommand.ExecuteAsync(null);
                    Debug.WriteLine($"過濾完成，已選: {string.Join(", ", SelectedSalesReps)}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"過濾發生錯誤: {ex.Message}");
                }
                finally
                {
                    IsLoading = false;
                }
            });
        }

        #endregion

        public DetailedSalesViewModel(IExcelService excelService, IExportService exportService)
        {
            _excelService = excelService;
            _exportService = exportService;
            Debug.WriteLine("DetailedSalesViewModel 已初始化");

            // 初始化已選擇的銷售代表列表
            SelectedSalesReps = new ObservableCollection<string> { "All Reps" };

            // 避免在構造函數中使用異步方法
            // 而是在事件循環的下一個循環執行初始化
            MainThread.BeginInvokeOnMainThread(() => InitializeAsync());
        }

        private async void InitializeAsync()
        {
            try
            {
                Debug.WriteLine("開始初始化數據");
                IsLoading = true;

                // 1. 載入原始數據
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                _allSalesData = data;
                Debug.WriteLine($"成功載入 {_allSalesData?.Count ?? 0} 條原始數據");

                // 2. 提取銷售代表列表
                var reps = _allSalesData
                    ?.Select(x => x.SalesRep)
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList() ?? new List<string>();

                if (!reps.Any())
                {
                    // 如果沒有從數據中提取到銷售代表，添加一些示例
                    reps = new List<string> { "Isaac", "Brandon", "Chris", "Mark", "Nathan" };
                }

                // 添加"全部"選項作為第一個選項
                reps.Insert(0, "All Reps");

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    AvailableSalesReps = new ObservableCollection<string>(reps);

                    // 初始化銷售代表選擇項
                    RepSelectionItems = new ObservableCollection<RepSelectionItem>();
                    foreach (var rep in reps)
                    {
                        RepSelectionItems.Add(new RepSelectionItem
                        {
                            Name = rep,
                            IsSelected = rep == "All Reps" // 默認只選中"All Reps"
                        });
                    }

                    // 初始化選中項
                    SelectedSalesReps = new ObservableCollection<string> { "All Reps" };
                    OnPropertyChanged(nameof(SelectedRepsText));
                });

                Debug.WriteLine($"載入了 {AvailableSalesReps.Count} 個銷售代表選項");

                // 3. 按日期範圍過濾數據
                FilterDataByDateRange();

                // 4. 加載適合當前視圖的數據
                LoadFilteredData();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"初始化時發生錯誤: {ex.Message}");

                // 發生錯誤時加載一些示例數據
                LoadSampleData();
            }
            finally
            {
                await MainThread.InvokeOnMainThreadAsync(() => IsLoading = false);
            }
        }

        private async void FilterDataAndReload()
        {
            try
            {
                MainThread.BeginInvokeOnMainThread(() => IsLoading = true);

                // 清除緩存，確保數據重新載入
                _excelService.ClearCache();

                // 過濾數據
                FilterDataByDateRange();

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    try
                    {
                        // 在主線程上載入過濾後的數據到 UI
                        LoadFilteredData();
                        Debug.WriteLine("已根據條件更新表格數據");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"加載過濾數據到UI時發生錯誤: {ex.Message}");
                    }
                    finally
                    {
                        IsLoading = false;
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"過濾和重新載入數據時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    try
                    {
                        // 發生錯誤時，嘗試載入樣本數據
                        LoadSampleData();
                    }
                    catch
                    {
                        // 最後的保險措施
                    }
                    finally
                    {
                        IsLoading = false;
                    }
                });
            }
        }

        private void FilterDataByDateRange()
        {
            if (_allSalesData == null || !_allSalesData.Any())
            {
                Debug.WriteLine("沒有原始數據可以過濾");
                _filteredSalesData = new List<SalesData>();

                // 顯示警告訊息給用戶
                MainThread.BeginInvokeOnMainThread(async () =>
                {
                    await Application.Current.MainPage.DisplayAlert(
                        "無可用數據",
                        "無法從Excel檔案讀取銷售數據。請確認檔案存在且可訪問。",
                        "確定");
                });
                return;
            }

            try
            {
                // 獲取最新的日期值
                var currentStartDate = StartDate.Date;
                var currentEndDate = EndDate.Date.AddDays(1).AddSeconds(-1); // 包含結束日期的整天

                Debug.WriteLine($"過濾日期範圍: {currentStartDate:yyyy-MM-dd HH:mm:ss} 到 {currentEndDate:yyyy-MM-dd HH:mm:ss}");

                // 輸出原始數據中的日期範圍，幫助確認有效數據
                var allDates = _allSalesData.Select(x => x.ReceivedDate.Date).Distinct().OrderBy(d => d).ToList();
                Debug.WriteLine($"原始數據中的所有不同日期 ({allDates.Count}個): {string.Join(", ", allDates.Take(10))}...");

                // 根據日期範圍過濾
                _filteredSalesData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= currentStartDate.Date &&
                               x.ReceivedDate.Date <= currentEndDate.Date)
                    .ToList();

                Debug.WriteLine($"日期過濾後剩餘 {_filteredSalesData.Count} 條數據");

                // 輸出過濾後的所有日期，確認過濾是否正確
                var filteredDates = _filteredSalesData.Select(x => x.ReceivedDate.Date).Distinct().OrderBy(d => d).ToList();
                Debug.WriteLine($"過濾後數據中的所有不同日期 ({filteredDates.Count}個): {string.Join(", ", filteredDates)}");

                // 檢查日期範圍內是否有任何數據
                if (_filteredSalesData.Count == 0)
                {
                    Debug.WriteLine("選擇的日期範圍內沒有數據");

                    // 顯示警告訊息給用戶
                    MainThread.BeginInvokeOnMainThread(async () =>
                    {
                        await Application.Current.MainPage.DisplayAlert(
                            "No Data",
    $"No sales data found in the period from {currentStartDate:yyyy-MM-dd} to {currentEndDate:yyyy-MM-dd}.",
    "OK");
                    });

                    // 保持_filteredSalesData為空列表，這樣UI將顯示無數據狀態
                    return;
                }

                // 記錄過濾前的數據，用於檢查特定銷售代表是否有數據
                var preFilterSalesReps = _filteredSalesData
                    .Select(x => x.SalesRep)
                    .Where(x => !string.IsNullOrEmpty(x))
                    .Distinct()
                    .ToList();

                Debug.WriteLine($"日期範圍內的銷售代表: {string.Join(", ", preFilterSalesReps)}");

                // 根據銷售代表過濾
                if (SelectedSalesReps.Any() && !SelectedSalesReps.Contains("All Reps"))
                {
                    var beforeCount = _filteredSalesData.Count;

                    // 找出選擇的哪些代表在當前日期範圍內沒有數據
                    var missingReps = SelectedSalesReps
                        .Where(rep => !preFilterSalesReps.Any(x =>
                            string.Equals(x, rep, StringComparison.OrdinalIgnoreCase) ||
                            x.Contains(rep, StringComparison.OrdinalIgnoreCase) ||
                            rep.Contains(x, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    if (missingReps.Any())
                    {
                        Debug.WriteLine($"警告: 以下選中的銷售代表在當前日期範圍內沒有數據: {string.Join(", ", missingReps)}");

                        // 顯示警告訊息給用戶
                        MainThread.BeginInvokeOnMainThread(async () =>
                        {
                            await Application.Current.MainPage.DisplayAlert(
                                "Some Representatives Have No Data",
                                $"The following selected sales representatives have no data records in the selected date range: {string.Join(", ", missingReps)}",
                                "OK");
                        });
                    }

                    // 只過濾存在的代表的數據
                    _filteredSalesData = _filteredSalesData
                        .Where(x => !string.IsNullOrEmpty(x.SalesRep) &&
                                SelectedSalesReps.Any(r =>
                                    string.Equals(x.SalesRep, r, StringComparison.OrdinalIgnoreCase) ||
                                    x.SalesRep.Contains(r, StringComparison.OrdinalIgnoreCase) ||
                                    r.Contains(x.SalesRep, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    Debug.WriteLine($"銷售代表過濾: 從 {beforeCount} 條記錄中篩選，" +
                                   $"剩餘 {_filteredSalesData.Count} 條數據，已選: {string.Join(", ", SelectedSalesReps)}");

                    // 打印一些匹配到的記錄，幫助調試
                    if (_filteredSalesData.Any())
                    {
                        Debug.WriteLine("匹配的記錄示例:");
                        foreach (var record in _filteredSalesData.Take(Math.Min(3, _filteredSalesData.Count)))
                        {
                            Debug.WriteLine($"  SalesRep: {record.SalesRep}, ProductType: {record.ProductType}");
                        }
                    }

                    // 如果過濾後沒有任何數據
                    if (_filteredSalesData.Count == 0)
                    {
                        Debug.WriteLine("過濾後沒有任何匹配的數據");

                        // 顯示警告訊息給用戶
                        MainThread.BeginInvokeOnMainThread(async () =>
                        {
                            await Application.Current.MainPage.DisplayAlert(
                                "No Matching Data",
                                $"The selected sales representatives have no matching data records in the current date range.",
                                "OK");
                        });
                    }
                }
                else
                {
                    Debug.WriteLine("使用All Reps設置，顯示所有符合日期範圍的數據");
                }

                // 輸出最終過濾結果的詳細信息
                Debug.WriteLine("最終過濾後的數據結果:");
                foreach (var item in _filteredSalesData.Take(Math.Min(10, _filteredSalesData.Count)))
                {
                    Debug.WriteLine($"日期: {item.ReceivedDate:yyyy-MM-dd}, 產品: {item.ProductType}, PO值: ${item.POValue:N2}, 銷售代表: {item.SalesRep}");
                }

                // 按產品類型分組統計
                var productStats = _filteredSalesData
                    .GroupBy(x => x.ProductType)
                    .Select(g => new
                    {
                        ProductType = g.Key,
                        Count = g.Count(),
                        TotalPOValue = g.Sum(x => x.POValue)
                    }).OrderByDescending(x => x.TotalPOValue)
                    .ToList();

                Debug.WriteLine("產品類型統計:");
                foreach (var stat in productStats)
                {
                    Debug.WriteLine($"產品類型: {stat.ProductType}, 數量: {stat.Count}, 總PO值: ${stat.TotalPOValue:N2}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"過濾數據時發生錯誤: {ex.Message}");
                _filteredSalesData = new List<SalesData>();

                // 顯示錯誤訊息給用戶
                MainThread.BeginInvokeOnMainThread(async () =>
                {
                    await Application.Current.MainPage.DisplayAlert(
                        "過濾錯誤",
                        $"處理數據時發生錯誤: {ex.Message}",
                        "確定");
                });
            }
        }

        // 新增方法 - 生成適合當前選擇的示例數據
        private List<SalesData> GenerateMatchingTestData()
        {
            var testData = new List<SalesData>();

            // 確定要生成數據的銷售代表列表
            List<string> repsToGenerate;
            if (SelectedSalesReps.Contains("All Reps") || !SelectedSalesReps.Any())
            {
                // 如果選擇了All Reps或沒有選擇任何代表，為所有可能的銷售代表生成數據
                repsToGenerate = new List<string> { "Isaac", "Brandon", "Chris", "Mark", "Nathan" };
            }
            else
            {
                // 僅為選中的銷售代表生成數據
                repsToGenerate = SelectedSalesReps.ToList();
            }

            Debug.WriteLine($"為以下銷售代表生成示例數據: {string.Join(", ", repsToGenerate)}");

            // 針對每個銷售代表生成數據
            foreach (var rep in repsToGenerate)
            {
                // 生成不同產品類型的數據
                foreach (var product in new[] { "Power", "Thermal", "Channel", "Service", "Batts & Caps" })
                {
                    // 根據不同代表和產品類型設置不同的基準金額
                    decimal baseAmount = rep switch
                    {
                        "Isaac" => 30000m,
                        "Brandon" => 25000m,
                        "Chris" => 20000m,
                        "Mark" => 18000m,
                        "Nathan" => 15000m,
                        _ => 10000m
                    };

                    // 根據產品類型調整基準金額
                    decimal productMultiplier = product switch
                    {
                        "Thermal" => 2.5m,
                        "Power" => 2.0m,
                        "Channel" => 1.5m,
                        "Batts & Caps" => 1.2m,
                        "Service" => 1.0m,
                        _ => 1.0m
                    };

                    decimal finalAmount = baseAmount * productMultiplier;

                    // 為每個月份生成一條數據
                    for (int month = 1; month <= 12; month++)
                    {
                        // 生成一個在日期範圍內的日期
                        var year = (month >= 8) ? 2023 : 2024;
                        var date = new DateTime(year, month, 15);

                        // 如果日期在過濾範圍內才添加
                        if (date >= StartDate && date <= EndDate)
                        {
                            testData.Add(new SalesData
                            {
                                ReceivedDate = date,
                                SalesRep = rep,
                                Status = month % 2 == 0 ? "Booked" : "Completed",
                                ProductType = product,
                                POValue = finalAmount,
                                VertivValue = finalAmount * 0.85m,
                                TotalCommission = finalAmount * 0.1m,
                                CommissionPercentage = 0.1m
                            });
                        }
                    }
                }
            }

            Debug.WriteLine($"已生成 {testData.Count} 條匹配當前選擇的示例數據");
            return testData;
        }

        private void LoadFilteredData()
        {
            try
            {
                Debug.WriteLine($"載入已過濾數據，當前視圖類型: {ViewType}");
                Debug.WriteLine($"過濾條件 - 開始日期: {StartDate:yyyy-MM-dd}, 結束日期: {EndDate:yyyy-MM-dd}");
                Debug.WriteLine($"已選銷售代表: {string.Join(", ", SelectedSalesReps)}");

                // 根據當前視圖類型載入相應數據
                if (ViewType == "ByProduct")
                {
                    LoadProductData();
                }
                else if (ViewType == "ByRep")
                {
                    LoadSalesRepData();
                }

                Debug.WriteLine("數據載入完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入已過濾數據時發生錯誤: {ex.Message}");
            }
        }

        private void LoadProductData()
        {
            try
            {
                Debug.WriteLine("Loading product view data");

                if (_filteredSalesData == null || !_filteredSalesData.Any())
                {
                    Debug.WriteLine("No filtered data available");
                    ProductSalesData = new ObservableCollection<ProductSalesData>();
                    return;
                }

                // Group by product type and calculate totals
                var productData = _filteredSalesData
                    .GroupBy(x => x.ProductType)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        // Direct from Excel's corresponding columns
                        AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                        // Use TotalCommission directly from the N column instead of summing
                        TotalMargin = Math.Round(g.Sum(x => x.TotalCommission), 2),
                        POValue = Math.Round(g.Sum(x => x.POValue), 2)
                    })
                    .OrderByDescending(x => x.POValue)
                    .ToList();

                // Calculate PO Value percentages
                decimal totalPOValue = productData.Sum(p => p.POValue);
                foreach (var product in productData)
                {
                    product.PercentageOfTotal = totalPOValue > 0
                        ? Math.Round((product.POValue / totalPOValue) * 100, 2)
                        : 0;
                }

                ProductSalesData = new ObservableCollection<ProductSalesData>(productData);
                Debug.WriteLine($"Loaded {productData.Count} product data items");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading product data: {ex.Message}");

                // In case of error, display empty data
                var emptyProductData = new List<ProductSalesData>();
                foreach (var productType in new[] { "Thermal", "Power", "Channel", "Service", "Batts & Caps" })
                {
                    emptyProductData.Add(new ProductSalesData
                    {
                        ProductType = productType,
                        AgencyMargin = 0,
                        BuyResellMargin = 0,
                        TotalMargin = 0,
                        POValue = 0,
                        PercentageOfTotal = 0
                    });
                }

                MainThread.BeginInvokeOnMainThread(() =>
                {
                    ProductSalesData = new ObservableCollection<ProductSalesData>(emptyProductData);
                });
            }
        }


        // 標準化產品類型名稱，處理可能的大小寫或拼寫差異
        private string NormalizeProductType(string productType)
        {
            if (string.IsNullOrEmpty(productType))
                return "Other";

            // 轉為小寫以便比較
            string lowercaseType = productType.ToLowerInvariant();

            if (lowercaseType.Contains("thermal"))
                return "Thermal";
            if (lowercaseType.Contains("power"))
                return "Power";
            if (lowercaseType.Contains("channel"))
                return "Channel";
            if (lowercaseType.Contains("service"))
                return "Service";
            if (lowercaseType.Contains("batts") || lowercaseType.Contains("caps") || lowercaseType.Contains("batt"))
                return "Batts & Caps";

            // 如果沒有匹配，返回原始名稱
            return productType;
        }

        // 標準化銷售代表名稱
        private string NormalizeSalesRep(string salesRep)
        {
            if (string.IsNullOrWhiteSpace(salesRep))
                return "Unknown";

            return salesRep.Trim();
        }

        private void LoadSalesRepData()
        {
            try
            {
                var salesRepData = new List<SalesLeaderboardItem>();

                // 處理實際數據
                if (_filteredSalesData?.Any() == true)
                {
                    salesRepData = _filteredSalesData
                        .GroupBy(x => x.SalesRep)
                        .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                        .Select(g => new SalesLeaderboardItem
                        {
                            SalesRep = g.Key,
                            // 直接從Excel的M列獲取Agency Margin
                            AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                            // 直接從Excel的J列獲取Buy Resell
                            BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                            // Total Margin是兩者之和
                            TotalMargin = Math.Round(g.Sum(x => x.AgencyMargin) + g.Sum(x => x.BuyResellValue), 2)
                        })
                        .OrderByDescending(x => x.TotalMargin)
                        .ToList();
                }


                // 使用過濾後的數據
                var repData = _filteredSalesData
                    .GroupBy(x => NormalizeSalesRep(x.SalesRep))
                    .Where(g => !string.IsNullOrEmpty(g.Key))
                    .Select(g =>
                    {
                        var rep = new SalesLeaderboardItem
                        {
                            SalesRep = g.Key,
                            // 更新屬性名稱
                            AgencyMargin = g.Sum(x => x.TotalCommission * 0.7m),
                            BuyResellMargin = g.Sum(x => x.TotalCommission * 0.3m)
                        };

                        // 明確設置 TotalMargin
                        rep.TotalMargin = rep.AgencyMargin + rep.BuyResellMargin;
                        return rep;
                    })
                    .OrderByDescending(x => x.TotalMargin)
                    .ToList();

                if (repData.Any())
                {
                    // 更新排名
                    for (int i = 0; i < repData.Count; i++)
                    {
                        repData[i].Rank = i + 1;
                    }

                    // 處理選中但在數據中不存在的銷售代表（在已選擇銷售代表模式下）
                    if (SelectedSalesReps.Any() && !SelectedSalesReps.Contains("All Reps"))
                    {
                        var existingReps = repData.Select(r => r.SalesRep).ToList();

                        // 為不存在於數據中但已選擇的代表創建空記錄
                        foreach (var selectedRep in SelectedSalesReps)
                        {
                            // 檢查該代表是否已存在於結果中（考慮標準化名稱）
                            bool exists = existingReps.Any(r =>
                                string.Equals(r, selectedRep, StringComparison.OrdinalIgnoreCase) ||
                                r.Contains(selectedRep, StringComparison.OrdinalIgnoreCase) ||
                                selectedRep.Contains(r, StringComparison.OrdinalIgnoreCase));

                            if (!exists)
                            {
                                // 添加一條空記錄
                                repData.Add(new SalesLeaderboardItem
                                {
                                    Rank = repData.Count + 1,
                                    SalesRep = selectedRep,
                                    AgencyMargin = 0,
                                    BuyResellMargin = 0,
                                    TotalMargin = 0
                                });

                                Debug.WriteLine($"為已選擇但無數據的銷售代表添加空記錄: {selectedRep}");
                            }
                        }

                        // 重新排序
                        repData = repData.OrderByDescending(x => x.TotalMargin).ToList();

                        // 重新設置排名
                        for (int i = 0; i < repData.Count; i++)
                        {
                            repData[i].Rank = i + 1;
                        }
                    }

                    MainThread.BeginInvokeOnMainThread(() =>
                    {
                        SalesRepData = new ObservableCollection<SalesLeaderboardItem>(repData);
                    });

                    Debug.WriteLine($"已載入 {repData.Count} 條銷售代表數據");

                    // 載入銷售代表產品數據（新增）
                    LoadSalesRepProductData();

                    return;
                }
                else
                {
                    // 如果沒有銷售代表數據
                    Debug.WriteLine("沒有銷售代表數據可顯示");

                    // 創建空的銷售代表數據
                    var emptyRepData = CreateEmptySalesRepData();

                    MainThread.BeginInvokeOnMainThread(() =>
                    {
                        SalesRepData = new ObservableCollection<SalesLeaderboardItem>(emptyRepData);
                    });

                    Debug.WriteLine("已載入空銷售代表數據（所有值為0）");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入銷售代表數據時發生錯誤: {ex.Message}");

                // 發生錯誤時顯示空數據
                var emptyRepData = CreateEmptySalesRepData();

                MainThread.BeginInvokeOnMainThread(() =>
                {
                    SalesRepData = new ObservableCollection<SalesLeaderboardItem>(emptyRepData);
                });
            }
        }

        // 新增：載入銷售代表產品數據
        private void LoadSalesRepProductData()
        {
            try
            {
                Debug.WriteLine("載入銷售代表產品數據");

                if (_filteredSalesData == null || !_filteredSalesData.Any())
                {
                    Debug.WriteLine("沒有過濾後的數據可用於銷售代表產品分析");
                    // 創建空白數據
                    var emptyProductData = CreateEmptySalesRepProductData();
                    MainThread.BeginInvokeOnMainThread(() =>
                    {
                        SalesRepProductData = new ObservableCollection<ProductSalesData>(emptyProductData);
                    });
                    return;
                }

                // 確定要處理的數據
                var dataToProcess = _filteredSalesData;

                // 如果選擇了特定銷售代表（不是All Reps），則按所選代表過濾
                if (SelectedSalesReps.Any() && !SelectedSalesReps.Contains("All Reps"))
                {
                    dataToProcess = _filteredSalesData
                        .Where(x => !string.IsNullOrEmpty(x.SalesRep) &&
                                SelectedSalesReps.Any(r =>
                                    string.Equals(x.SalesRep, r, StringComparison.OrdinalIgnoreCase) ||
                                    x.SalesRep.Contains(r, StringComparison.OrdinalIgnoreCase) ||
                                    r.Contains(x.SalesRep, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    Debug.WriteLine($"按選定的銷售代表過濾後剩餘 {dataToProcess.Count} 條記錄");

                    // 如果過濾後沒有數據，顯示空白數據
                    if (!dataToProcess.Any())
                    {
                        Debug.WriteLine("選定的銷售代表在當前日期範圍內沒有數據");
                        var emptyProductData = CreateEmptySalesRepProductData();
                        MainThread.BeginInvokeOnMainThread(() =>
                        {
                            SalesRepProductData = new ObservableCollection<ProductSalesData>(emptyProductData);
                        });
                        return;
                    }
                }

                // 按產品類型分組計算
                var productData = dataToProcess
                    .GroupBy(x => NormalizeProductType(x.ProductType))
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        POValue = Math.Round(g.Sum(x => x.POValue), 2),
                        PercentageOfTotal = 0  // 先設為0，下面再計算
                    })
                    .OrderByDescending(x => x.POValue)
                    .ToList();

                // 計算百分比
                decimal totalPOValue = productData.Sum(p => p.POValue);
                foreach (var product in productData)
                {
                    product.PercentageOfTotal = totalPOValue > 0
                        ? Math.Round((product.POValue / totalPOValue) * 100, 2)
                        : 0;
                }

                MainThread.BeginInvokeOnMainThread(() =>
                {
                    SalesRepProductData = new ObservableCollection<ProductSalesData>(productData);
                });

                Debug.WriteLine($"已載入 {productData.Count} 條銷售代表產品數據");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入銷售代表產品數據時發生錯誤: {ex.Message}");

                // 發生錯誤時創建空白數據
                var emptyProductData = CreateEmptySalesRepProductData();
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    SalesRepProductData = new ObservableCollection<ProductSalesData>(emptyProductData);
                });
            }
        }

        // 新增：創建空白的銷售代表產品數據
        private List<ProductSalesData> CreateEmptySalesRepProductData()
        {
            var emptyData = new List<ProductSalesData>();

            // 為常見產品類型創建空白記錄
            foreach (var productType in new[] { "Power", "Thermal", "Channel", "Service", "Batts & Caps" })
            {
                emptyData.Add(new ProductSalesData
                {
                    ProductType = productType,
                    POValue = 0,
                    PercentageOfTotal = 0
                });
            }

            return emptyData;
        }

        // 修改：載入樣本銷售代表產品數據（現在不再需要固定示例數據）
        private void LoadSampleSalesRepProductData()
        {
            // Create empty data instead of using hardcoded sample data
            var emptyData = CreateEmptySalesRepProductData();

            MainThread.BeginInvokeOnMainThread(() =>
            {
                SalesRepProductData = new ObservableCollection<ProductSalesData>(emptyData);
            });

            Debug.WriteLine("Loaded empty sales rep product data");
        }


        // 創建空的銷售代表數據
        private List<SalesLeaderboardItem> CreateEmptySalesRepData()
        {
            var emptyData = new List<SalesLeaderboardItem>();

            // 如果有選擇特定的銷售代表
            if (SelectedSalesReps.Any() && !SelectedSalesReps.Contains("All Reps"))
            {
                // 為每個選中的銷售代表創建空記錄
                for (int i = 0; i < SelectedSalesReps.Count; i++)
                {
                    emptyData.Add(new SalesLeaderboardItem
                    {
                        Rank = i + 1,
                        SalesRep = SelectedSalesReps[i],
                        AgencyMargin = 0,
                        BuyResellMargin = 0,
                        TotalMargin = 0
                    });
                }
            }
            else
            {
                // 如果選擇了"All Reps"或沒有選擇，則創建常見銷售代表的空記錄
                var commonReps = new[] { "Isaac", "Brandon", "Chris", "Mark", "Nathan" };

                for (int i = 0; i < commonReps.Length; i++)
                {
                    emptyData.Add(new SalesLeaderboardItem
                    {
                        Rank = i + 1,
                        SalesRep = commonReps[i],
                        AgencyMargin = 0,
                        BuyResellMargin = 0,
                        TotalMargin = 0
                    });
                }
            }

            return emptyData;
        }

        // 當所有其他方法都失敗時，載入固定的示例數據
        private void LoadSampleData()
        {
            LoadSampleProductData();
            LoadSampleSalesRepData();
            LoadSampleSalesRepProductData();
        }

        private void LoadSampleProductData()
        {
            var tempData = new List<ProductSalesData>
            {
                new ProductSalesData
                {
                    ProductType = "Thermal",
                    AgencyMargin = 744855.43m,
                    BuyResellMargin = 116206.36m,
                    TotalMargin = 861061.79m, // 確保明確設置TotalCommission
                    POValue = 7358201.65m,
                    PercentageOfTotal = 41.0m
                },
                new ProductSalesData
                {
                    ProductType = "Power",
                    AgencyMargin = 296743.08m,
                    BuyResellMargin = 8737.33m,
                    TotalMargin = 305481.01m,
                    POValue = 5466144.65m,
                    PercentageOfTotal = 31.0m
                },
                new ProductSalesData
                {
                    ProductType = "Batts & Caps",
                    AgencyMargin = 250130.95m,
                    BuyResellMargin = 0.00m,
                    TotalMargin = 250130.95m,
                    POValue = 2061423.30m,
                    PercentageOfTotal = 12.0m
                },
                new ProductSalesData
                {
                    ProductType = "Channel",
                    AgencyMargin = 167353.03m,
                    BuyResellMargin = 8323.03m,
                    TotalMargin = 175676.06m,
                    POValue = 1416574.65m,
                    PercentageOfTotal = 8.0m
                },
                new ProductSalesData
                {
                    ProductType = "Service",
                    AgencyMargin = 101556.42m,
                    BuyResellMargin = 0.00m,
                    TotalMargin = 101556.42m,
                    POValue = 1272318.58m,
                    PercentageOfTotal = 7.0m
                }
            };

            // 確保數據不重複 - 將樣本數據直接設置為有序版本
            MainThread.BeginInvokeOnMainThread(() =>
            {
                ProductSalesData = new ObservableCollection<ProductSalesData>(
                    tempData.OrderByDescending(p => p.POValue)
                );
            });

            Debug.WriteLine($"已載入 {tempData.Count} 條示例產品數據");
        }

        private void LoadSampleSalesRepData()
        {
            var sampleData = new List<SalesLeaderboardItem>
            {
                new SalesLeaderboardItem
                {
                    Rank = 1,
                    SalesRep = "Isaac",
                    AgencyMargin = 350186.00m,
                    BuyResellMargin = 0.00m,
                    TotalMargin = 350186.00m
                },
                new SalesLeaderboardItem
                {
                    Rank = 2,
                    SalesRep = "Brandon",
                    AgencyMargin = 301802.40m,
                    BuyResellMargin = 38165.70m,
                    TotalMargin = 339968.10m
                },
                new SalesLeaderboardItem
                {
                    Rank = 3,
                    SalesRep = "Chris",
                    AgencyMargin = 186411.10m,
                    BuyResellMargin = 0.00m,
                    TotalMargin = 186411.10m
                },
                new SalesLeaderboardItem
                {
                    Rank = 4,
                    SalesRep = "Mark",
                    AgencyMargin = 124680.50m,
                    BuyResellMargin = 18920.30m,
                    TotalMargin = 143600.80m
                },
                new SalesLeaderboardItem
                {
                    Rank = 5,
                    SalesRep = "Nathan",
                    AgencyMargin = 104582.20m,
                    BuyResellMargin = 21060.80m,
                    TotalMargin = 125643.00m
                }
            };

            MainThread.BeginInvokeOnMainThread(() =>
            {
                SalesRepData = new ObservableCollection<SalesLeaderboardItem>(sampleData);
            });

            Debug.WriteLine($"已載入 {sampleData.Count} 條示例銷售代表數據");
        }

    }
}