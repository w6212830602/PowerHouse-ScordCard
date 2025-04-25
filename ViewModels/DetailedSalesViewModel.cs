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
        private bool _isAllStatus = true; // Default to All status

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

        [ObservableProperty]
        private bool _isBookedStatus = false; // 默認為 Booked 狀態

        [ObservableProperty]
        private bool _isInProgressStatus = false;

        [ObservableProperty]
        private bool _isInvoicedStatus = false;


        // 臨時存儲選擇狀態
        private ObservableCollection<RepSelectionItem> _tempRepSelectionItems = new();

        // 計算屬性 - 用於繫結到 UI
        public bool IsProductView => ViewType == "ByProduct";
        public bool IsRepView => ViewType == "ByRep";

        public bool IsBookedStatusActive => IsBookedStatus;
        public bool IsInProgressStatusActive => IsInProgressStatus;
        public bool IsInvoicedStatusActive => IsInvoicedStatus;


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
            OnPropertyChanged(nameof(SelectedRepsDetailText)); // 添加这一行
        }

        #endregion

        #region Commands


        [RelayCommand]
        private async Task ChangeStatus(string status)
        {
            Debug.WriteLine($"Attempting to switch to status: {status}");
            bool changed = false;

            switch (status.ToLower())
            {
                case "all":
                    if (!IsAllStatus)
                    {
                        IsAllStatus = true;
                        IsBookedStatus = false;
                        IsInProgressStatus = false;
                        IsInvoicedStatus = false;
                        changed = true;
                    }
                    break;

                case "booked":
                    if (!IsBookedStatus)
                    {
                        IsAllStatus = false;
                        IsBookedStatus = true;
                        IsInProgressStatus = false;
                        IsInvoicedStatus = false;
                        changed = true;
                    }
                    break;

                case "inprogress":
                    if (!IsInProgressStatus)
                    {
                        IsAllStatus = false;
                        IsBookedStatus = false;
                        IsInProgressStatus = true;
                        IsInvoicedStatus = false;
                        changed = true;
                    }
                    break;

                case "invoiced": // Changed from "completed" to "invoiced" for consistency
                    if (!IsInvoicedStatus)
                    {
                        IsAllStatus = false;
                        IsBookedStatus = false;
                        IsInProgressStatus = false;
                        IsInvoicedStatus = true;
                        changed = true;
                    }
                    break;
            }

            if (changed)
            {
                Debug.WriteLine($"Status changed to: {status}, reloading data");
                await FilterDataCommand.ExecuteAsync(null);
            }
            else
            {
                Debug.WriteLine($"Status not changed, still: {status}");
            }
        }


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
            var productData = GetCurrentViewProductData();

            // 確保數據已經排序且百分比計算正確
            if (productData.Any())
            {
                // 已按VertivValue/POValue排序
                decimal totalValue = productData.Sum(p => p.VertivValue);

                // 更新百分比
                foreach (var product in productData)
                {
                    product.PercentageOfTotal = Math.Round((product.VertivValue / totalValue) * 100, 1);
                }
            }

            return productData;
        }


        [RelayCommand]
        private async Task Export(string format)
        {
            IsExportOptionsVisible = false;
            Debug.WriteLine($"匯出格式: {format}");

            try
            {
                IsLoading = true;
                await Task.Delay(300);

                // 使用 GetDataToExport 方法獲取數據！
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
            Debug.WriteLine($"Preparing data for export, current view type: {ViewType}");

            // 根據當前視圖類型，返回適當的數據
            if (IsProductView && ProductSalesData.Any())
            {
                // 返回產品視圖數據 - 確保每個產品的VertivValue和百分比都已正確設置
                var data = ProductSalesData.ToList();

                // 確保所有產品都有正確的VertivValue和PercentageOfTotal屬性
                decimal totalVertivValue = data.Sum(p => p.VertivValue);
                if (totalVertivValue > 0)
                {
                    foreach (var product in data)
                    {
                        // 如果百分比未設置，則計算它
                        if (product.PercentageOfTotal <= 0)
                        {
                            product.PercentageOfTotal = Math.Round((product.VertivValue / totalVertivValue) * 100, 1);
                        }
                    }
                }

                Debug.WriteLine($"正在匯出產品視圖數據，{data.Count}項，" +
                               $"第一項：{(data.Any() ? data[0].ProductType : "無")}，" +
                               $"總計：${data.Sum(p => p.VertivValue):N2}");
                return data;
            }
            else if (IsRepView && SalesRepData.Any())
            {
                var data = SalesRepData.ToList();
                Debug.WriteLine($"正在匯出銷售代表視圖數據，{data.Count}項");
                return data;
            }
            else if (IsRepView && SalesRepProductData.Any())
            {
                // 匯出銷售代表產品數據 - 包含PO Vertiv Value表所需的信息
                var data = SalesRepProductData.ToList();

                // 確保百分比計算正確
                decimal totalPOValue = data.Sum(p => p.POValue);
                if (totalPOValue > 0)
                {
                    foreach (var product in data)
                    {
                        product.PercentageOfTotal = Math.Round((product.POValue / totalPOValue) * 100, 1);
                    }
                }

                Debug.WriteLine($"作為後備方案匯出銷售代表產品數據，{data.Count}項");
                return data;
            }

            Debug.WriteLine("沒有找到可匯出的數據");
            return new List<object>();
        }

        public List<ProductSalesData> GetCurrentViewProductData()
        {
            // 根據當前視圖返回適當的產品數據
            if (IsProductView)
            {
                // 返回當前產品視圖中顯示的數據
                return ProductSalesData.ToList();
            }
            else if (IsRepView)
            {
                // 返回銷售代表產品視圖數據
                return SalesRepProductData.ToList();
            }

            // 默認情況下返回空列表
            return new List<ProductSalesData>();
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
            OnPropertyChanged(nameof(SelectedRepsDetailText));

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
                Debug.WriteLine("No original data available for filtering");
                _filteredSalesData = new List<SalesData>();
                return;
            }

            try
            {
                // Get the current date values
                var currentStartDate = StartDate.Date;
                var currentEndDate = EndDate.Date.AddDays(1).AddSeconds(-1); // Include the entire end date

                Debug.WriteLine($"Filtering date range: {currentStartDate:yyyy-MM-dd} to {currentEndDate:yyyy-MM-dd}");

                // New handling for "All" status
                if (IsAllStatus)
                {
                    // All status: Include all records within the date range
                    // For Booked/InProgress: filter by ReceivedDate
                    // For Invoiced: filter by CompletionDate
                    var bookedData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= currentStartDate.Date &&
                               x.ReceivedDate.Date <= currentEndDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission > 0)
                        .ToList();

                    var inProgressData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= currentStartDate.Date &&
                               x.ReceivedDate.Date <= currentEndDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission == 0)
                        .ToList();

                    var invoicedData = _allSalesData
                        .Where(x => x.CompletionDate.HasValue &&
                               x.CompletionDate.Value.Date >= currentStartDate.Date &&
                               x.CompletionDate.Value.Date <= currentEndDate.Date)
                        .ToList();

                    // Combine all data
                    _filteredSalesData = bookedData.Concat(inProgressData).Concat(invoicedData).ToList();
                    Debug.WriteLine($"[All] Combined filtered records: {_filteredSalesData.Count} (Booked: {bookedData.Count}, In Progress: {inProgressData.Count}, Invoiced: {invoicedData.Count})");
                }
                // Apply individual status filters (existing logic)
                else if (IsBookedStatus)
                {
                    // Booked: A列(ReceivedDate)在日期範圍內，Y列(CompletionDate)為空，N列(TotalCommission)有值
                    _filteredSalesData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= currentStartDate.Date &&
                               x.ReceivedDate.Date <= currentEndDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission > 0)
                        .ToList();

                    Debug.WriteLine($"[Booked] Filtered records: {_filteredSalesData.Count}");
                }
                else if (IsInProgressStatus)
                {
                    // In Progress: A列(ReceivedDate)在日期範圍內，Y列(CompletionDate)為空，N列(TotalCommission)為零或空
                    _filteredSalesData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= currentStartDate.Date &&
                               x.ReceivedDate.Date <= currentEndDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission == 0)
                        .ToList();

                    Debug.WriteLine($"[In Progress] Filtered records: {_filteredSalesData.Count}");
                }
                else if (IsInvoicedStatus)
                {
                    // Invoiced/Completed: Y列(CompletionDate)在日期範圍內且不為空
                    _filteredSalesData = _allSalesData
                        .Where(x => x.CompletionDate.HasValue &&
                               x.CompletionDate.Value.Date >= currentStartDate.Date &&
                               x.CompletionDate.Value.Date <= currentEndDate.Date)
                        .ToList();

                    Debug.WriteLine($"[Invoiced] Filtered records: {_filteredSalesData.Count}");
                }
                else
                {
                    // Default: empty data
                    _filteredSalesData = new List<SalesData>();
                    Debug.WriteLine("No status selected, returning empty data");
                }

                // Apply sales rep filtering if needed
                if (SelectedSalesReps.Any() && !SelectedSalesReps.Contains("All Reps"))
                {
                    var beforeCount = _filteredSalesData.Count;

                    _filteredSalesData = _filteredSalesData
                        .Where(x => !string.IsNullOrEmpty(x.SalesRep) &&
                                SelectedSalesReps.Any(r =>
                                    string.Equals(x.SalesRep, r, StringComparison.OrdinalIgnoreCase) ||
                                    x.SalesRep.Contains(r, StringComparison.OrdinalIgnoreCase) ||
                                    r.Contains(x.SalesRep, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    Debug.WriteLine($"Sales Rep filtering: from {beforeCount} records, " +
                                   $"remaining {_filteredSalesData.Count} records");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error filtering data: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                _filteredSalesData = new List<SalesData>();
            }
        }


        private void LoadFilteredData()
        {
            try
            {
                Debug.WriteLine($"載入已過濾數據，當前視圖類型: {ViewType}");

                // 根據當前視圖類型載入相應數據
                if (ViewType == "ByProduct")
                {
                    LoadProductData(_filteredSalesData);
                }
                else if (ViewType == "ByRep")
                {
                    LoadSalesRepData(_filteredSalesData);

                    // 重要：同時載入 By Rep 視圖下的產品數據，用於 PO Value 表格
                    LoadSalesRepProductData();
                }

                Debug.WriteLine("數據載入完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入已過濾數據時發生錯誤: {ex.Message}");
            }
        }

        // 創建空白的銷售代表產品數據的輔助方法
        private List<ProductSalesData> CreateEmptySalesRepProductDataForView()
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



        private async Task LoadProductData(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine($"Loading product data, record count: {data.Count}");

                // 判斷是否處於 In Progress 狀態
                bool isInProgressMode = IsInProgressStatus;

                var products = data
                    .GroupBy(x => x.ProductType)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        // 如果是 In Progress 模式，则使用 POValue * 0.12 作为 Total Margin
                        // 进一步拆分为 75% Agency 和 25% Buy Resell
                        AgencyMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.12m) : // 75% of expected commission (0.12)
                            g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.00m) : // 25% of expected commission (0.12)
                            g.Sum(x => x.BuyResellValue), 2),
                        // Total Margin 直接计算，不使用 TotalCommission
                        TotalMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.12m) : // Expected margin for in progress items
                            g.Sum(x => x.TotalCommission), 2),
                        // 修改为使用 VertivValue 代替 POValue
                        POValue = Math.Round(g.Sum(x => x.VertivValue), 2),
                        // 标记是否为 In Progress 模式（用于 UI 显示）
                        IsInProgress = isInProgressMode
                    })
                    .OrderByDescending(x => x.POValue)
                    .ToList();

                Debug.WriteLine($"Number of products after grouping: {products.Count}");

                if (products.Any())
                {
                    // 計算百分比
                    decimal totalPOValue = products.Sum(p => p.POValue);
                    foreach (var product in products)
                    {
                        product.PercentageOfTotal = totalPOValue > 0
                            ? Math.Round((product.POValue / totalPOValue) * 100, 2)
                            : 0;
                    }

                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        ProductSalesData = new ObservableCollection<ProductSalesData>(products);
                    });

                    Debug.WriteLine($"Successfully loaded {products.Count} product data items");
                }
                else
                {
                    Debug.WriteLine("No product data to display");
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        ProductSalesData = new ObservableCollection<ProductSalesData>();
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading product data: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    ProductSalesData = new ObservableCollection<ProductSalesData>();
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

        private async Task LoadSalesRepData(List<SalesData> data)
        {
            try
            {
                // 判斷是否處於 In Progress 狀態
                bool isInProgressMode = IsInProgressStatus;

                var reps = data
                    .GroupBy(x => x.SalesRep)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new SalesLeaderboardItem
                    {
                        SalesRep = g.Key,
                        // 如果是 In Progress 模式，則使用 POValue * 0.12 作為 Total Margin
                        // 進一步拆分為 75% Agency 和 25% Buy Resell
                        AgencyMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.12m) : // 75% of expected commission (0.12)
                            g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.00m) : // 25% of expected commission (0.12)
                            g.Sum(x => x.BuyResellValue), 2),
                        // Total Margin 直接計算，不使用 TotalCommission
                        TotalMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.12m) : // Expected margin for in progress items
                            g.Sum(x => x.TotalCommission), 2)
                    })
                    .OrderByDescending(x => x.TotalMargin)
                    .ToList();

                // 設置排名
                for (int i = 0; i < reps.Count; i++)
                {
                    reps[i].Rank = i + 1;
                }

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    SalesRepData = new ObservableCollection<SalesLeaderboardItem>(reps);
                });

                Debug.WriteLine($"Sales rep data loaded, {reps.Count} items");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading sales rep data: {ex.Message}");
            }
        }

        // 显示选中销售代表的详细名单（最多5个）
        public string SelectedRepsDetailText
        {
            get
            {
                if (SelectedSalesReps.Count == 0 || (SelectedSalesReps.Count == 1 && SelectedSalesReps[0] == "All Reps"))
                    return "All Reps";

                // 取前5个代表名称
                var topReps = SelectedSalesReps.Where(r => r != "All Reps").Take(5).ToList();

                // 如果没有选择具体代表，返回"All Reps"
                if (!topReps.Any())
                    return "All Reps";

                // 如果选择的代表少于等于5个，全部显示
                if (topReps.Count <= 5)
                    return string.Join(", ", topReps);

                // 如果超过5个，显示前5个加...
                return string.Join(", ", topReps) + "...";
            }
        }


        // 新增：載入銷售代表產品數據
        private async Task LoadSalesRepProductData()
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

                // 按產品類型分組計算
                var productData = dataToProcess
                    .GroupBy(x => NormalizeProductType(x.ProductType))
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        // 使用 VertivValue 代替 POValue
                        POValue = Math.Round(g.Sum(x => x.VertivValue), 2),
                        PercentageOfTotal = 0  // 先设为0，下面再计算
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