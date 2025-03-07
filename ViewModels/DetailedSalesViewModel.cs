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
        private string _selectedSalesRep = "All Reps";

        [ObservableProperty]
        private ObservableCollection<string> _salesReps = new();

        [ObservableProperty]
        private bool _isExportOptionsVisible;

        // 新增：已選擇的銷售代表列表（支持複選）
        [ObservableProperty]
        private ObservableCollection<string> _selectedSalesReps = new();

        // 新增：是否啟用多選模式
        [ObservableProperty]
        private bool _isMultiSelectMode = false;

        // 新增：跟踪銷售代表選擇狀態
        [ObservableProperty]
        private ObservableCollection<RepSelectionItem> _repSelectionItems = new();

        // 控制銷售代表選擇彈出視窗的顯示
        [ObservableProperty]
        private bool _isRepSelectionPopupVisible = false;

        // 臨時存儲選擇狀態
        private ObservableCollection<RepSelectionItem> _tempRepSelectionItems = new();

        // 計算屬性 - 用於繫結到 UI
        public bool IsProductView => ViewType == "ByProduct";
        public bool IsRepView => ViewType == "ByRep";

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

        // 當 SelectedSalesRep 變更時觸發的方法
        partial void OnSelectedSalesRepChanged(string value)
        {
            Debug.WriteLine($"選擇的銷售代表變更為: {value}");
            // 使用 Task.Run 異步執行，避免阻塞 UI 線程
            Task.Run(() => FilterDataAndReload());
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

        [RelayCommand]
        private async Task Export(string format)
        {
            IsExportOptionsVisible = false;
            Debug.WriteLine($"匯出格式: {format}");

            // 簡單訊息提示
            await Application.Current.MainPage.DisplayAlert(
                "匯出",
                $"正在將數據匯出為 {format} 格式",
                "確定");
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

        // 新增命令：切換多選模式
        [RelayCommand]
        private void ToggleMultiSelectMode()
        {
            IsMultiSelectMode = !IsMultiSelectMode;
            Debug.WriteLine($"多選模式已{(IsMultiSelectMode ? "啟用" : "禁用")}");

            // 重置選擇
            if (IsMultiSelectMode)
            {
                // 進入多選模式，清空已選項並更新選擇項狀態
                SelectedSalesReps.Clear();

                foreach (var item in RepSelectionItems)
                {
                    item.IsSelected = false;
                }

                // 默認選中"All Reps"
                var allRepsItem = RepSelectionItems.FirstOrDefault(x => x.Name == "All Reps");
                if (allRepsItem != null)
                {
                    allRepsItem.IsSelected = true;
                    SelectedSalesReps.Add("All Reps");
                }
            }
            else
            {
                // 退出多選模式，設置單選為第一個選中項或All Reps
                SelectedSalesRep = SelectedSalesReps.FirstOrDefault() ?? "All Reps";
            }

            // 重新過濾數據
            FilterDataAndReload();
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
                    SelectedSalesReps.Clear();
                    SelectedSalesReps.Add("All Reps");

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

                // 更新SelectedSalesReps集合
                if (item.IsSelected)
                {
                    // 選中當前項，同時取消選中"All Reps"
                    SelectedSalesReps.Add(item.Name);

                    var allRepsItem = RepSelectionItems.FirstOrDefault(x => x.Name == "All Reps");
                    if (allRepsItem != null && allRepsItem.IsSelected)
                    {
                        allRepsItem.IsSelected = false;
                        SelectedSalesReps.Remove("All Reps");
                    }
                }
                else
                {
                    // 取消選中當前項
                    SelectedSalesReps.Remove(item.Name);

                    // 如果沒有選中項，則自動選中"All Reps"
                    if (!RepSelectionItems.Any(x => x.Name != "All Reps" && x.IsSelected))
                    {
                        var allRepsItem = RepSelectionItems.FirstOrDefault(x => x.Name == "All Reps");
                        if (allRepsItem != null)
                        {
                            allRepsItem.IsSelected = true;
                            SelectedSalesReps.Add("All Reps");
                        }
                    }
                }
            }

            Debug.WriteLine($"已選銷售代表: {string.Join(", ", SelectedSalesReps)}");
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

            // 重新過濾數據
            FilterDataAndReload();
        }

        #endregion

        public DetailedSalesViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            Debug.WriteLine("DetailedSalesViewModel 已初始化");

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

                await MainThread.InvokeOnMainThreadAsync(() => {
                    SalesReps = new ObservableCollection<string>(reps);
                    SelectedSalesRep = "All Reps"; // 設置默認選項為"全部"

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
                });

                Debug.WriteLine($"載入了 {SalesReps.Count} 個銷售代表選項");

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

        private void FilterDataAndReload()
        {
            try
            {
                MainThread.BeginInvokeOnMainThread(() => IsLoading = true);

                // 清除緩存
                _excelService.ClearCache();

                FilterDataByDateRange();

                MainThread.BeginInvokeOnMainThread(() => {
                    // 在主線程上載入過濾後的數據到 UI
                    LoadFilteredData();
                    IsLoading = false;
                    Debug.WriteLine("根據條件更新了表格數據");
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"過濾和重新載入數據時發生錯誤: {ex.Message}");
                MainThread.BeginInvokeOnMainThread(() => IsLoading = false);
            }
        }

        private void FilterDataByDateRange()
        {
            if (_allSalesData == null || !_allSalesData.Any())
            {
                Debug.WriteLine("沒有原始數據可以過濾");
                _filteredSalesData = new List<SalesData>();
                return;
            }

            try
            {
                // 獲取最新的日期值
                var currentStartDate = StartDate.Date;
                var currentEndDate = EndDate.Date.AddDays(1).AddSeconds(-1); // 包含結束日期的整天

                Debug.WriteLine($"過濾日期範圍: {currentStartDate:yyyy-MM-dd HH:mm:ss} 到 {currentEndDate:yyyy-MM-dd HH:mm:ss}");

                // 輸出幾筆原始數據的日期，用於檢查
                for (int i = 0; i < Math.Min(_allSalesData.Count, 5); i++)
                {
                    Debug.WriteLine($"原始數據 {i}: ReceivedDate={_allSalesData[i].ReceivedDate:yyyy-MM-dd HH:mm:ss}");
                }

                // 根據日期範圍過濾
                _filteredSalesData = _allSalesData
                    .Where(x => x.ReceivedDate >= currentStartDate &&
                               x.ReceivedDate <= currentEndDate)
                    .ToList();

                Debug.WriteLine($"日期過濾後剩餘 {_filteredSalesData.Count} 條數據");

                // 根據銷售代表過濾
                if (IsMultiSelectMode && SelectedSalesReps.Any() && !SelectedSalesReps.Contains("All Reps"))
                {
                    // 多選模式，過濾符合任一選定代表的記錄
                    _filteredSalesData = _filteredSalesData
                        .Where(x => SelectedSalesReps.Contains(x.SalesRep))
                        .ToList();

                    Debug.WriteLine($"多選銷售代表過濾後剩餘 {_filteredSalesData.Count} 條數據，已選: {string.Join(", ", SelectedSalesReps)}");
                }
                else if (!IsMultiSelectMode && !string.IsNullOrEmpty(SelectedSalesRep) && SelectedSalesRep != "All Reps")
                {
                    // 單選模式，過濾符合所選代表的記錄
                    _filteredSalesData = _filteredSalesData
                        .Where(x => x.SalesRep == SelectedSalesRep)
                        .ToList();

                    Debug.WriteLine($"單選銷售代表過濾後剩餘 {_filteredSalesData.Count} 條數據，已選: {SelectedSalesRep}");
                }

                // 如果過濾後沒有數據
                if (_filteredSalesData.Count == 0)
                {
                    Debug.WriteLine($"過濾後沒有數據，回退到示例數據");

                    // 使用日期範圍生成合適的示例數據
                    LoadSampleData();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"過濾數據時發生錯誤: {ex.Message}");
                _filteredSalesData = new List<SalesData>();
            }
        }

        private void LoadFilteredData()
        {
            // 根據當前視圖類型載入相應數據
            if (ViewType == "ByProduct")
            {
                LoadProductData();
            }
            else if (ViewType == "ByRep")
            {
                LoadSalesRepData();
            }
        }

        private void LoadProductData()
        {
            try
            {
                Debug.WriteLine("載入產品視圖數據");

                // 優先使用過濾後的數據
                if (_filteredSalesData != null && _filteredSalesData.Any())
                {
                    var productData = _filteredSalesData
                        .GroupBy(x => x.ProductType)
                        .Where(g => !string.IsNullOrEmpty(g.Key))
                        .Select(g =>
                        {
                            var product = new ProductSalesData
                            {
                                ProductType = g.Key,
                                AgencyCommission = g.Sum(x => x.TotalCommission * 0.7m),  // 假設 70% 為代理傭金
                                BuyResellCommission = g.Sum(x => x.TotalCommission * 0.3m),  // 假設 30% 為買賣傭金
                                POValue = g.Sum(x => x.POValue)
                            };

                            // 明確設置 TotalCommission
                            product.TotalCommission = product.AgencyCommission + product.BuyResellCommission;
                            return product;
                        })
                        .OrderByDescending(x => x.POValue)
                        .ToList();

                    // 計算百分比
                    if (productData.Any())
                    {
                        decimal totalPO = productData.Sum(p => p.POValue);

                        foreach (var product in productData)
                        {
                            product.PercentageOfTotal = totalPO > 0 ? Math.Round((product.POValue / totalPO) * 100, 1) : 0;
                        }

                        // 去重處理
                        var uniqueProducts = productData
                            .GroupBy(p => p.ProductType)
                            .Select(g => g.First())
                            .OrderByDescending(p => p.POValue)
                            .ToList();

                        MainThread.BeginInvokeOnMainThread(() => {
                            ProductSalesData = new ObservableCollection<ProductSalesData>(uniqueProducts);
                        });

                        Debug.WriteLine($"從過濾數據計算出 {uniqueProducts.Count} 條產品數據（已去重）");
                        return;
                    }
                }

                // 如果過濾數據為空，嘗試從緩存獲取
                var cachedProductData = _excelService.GetProductSalesData();

                if (cachedProductData.Any())
                {
                    // 進行去重處理，確保每個產品類型只出現一次
                    var uniqueProducts = cachedProductData
                        .GroupBy(p => p.ProductType)
                        .Select(g => g.First())
                        .OrderByDescending(p => p.POValue)
                        .ToList();

                    MainThread.BeginInvokeOnMainThread(() => {
                        ProductSalesData = new ObservableCollection<ProductSalesData>(uniqueProducts);
                    });

                    Debug.WriteLine($"從緩存載入了 {uniqueProducts.Count} 條產品數據（已去重）");
                    return;
                }

                // 如果緩存也沒有數據，載入示例數據
                Debug.WriteLine("過濾後的數據和緩存都為空，載入示例產品數據");
                LoadSampleProductData();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入產品數據時發生錯誤: {ex.Message}");
                LoadSampleProductData();
            }
        }

        private void LoadSalesRepData()
        {
            try
            {
                Debug.WriteLine("載入銷售代表視圖數據");

                // 優先使用過濾後的數據
                if (_filteredSalesData != null && _filteredSalesData.Any())
                {
                    var repData = _filteredSalesData
                        .GroupBy(x => x.SalesRep)
                        .Where(g => !string.IsNullOrEmpty(g.Key))
                        .Select((g, index) =>
                        {
                            var rep = new SalesLeaderboardItem
                            {
                                Rank = index + 1,
                                SalesRep = g.Key,
                                AgencyCommission = g.Sum(x => x.TotalCommission * 0.7m),
                                BuyResellCommission = g.Sum(x => x.TotalCommission * 0.3m)
                            };

                            // 明確設置 TotalCommission
                            rep.TotalCommission = rep.AgencyCommission + rep.BuyResellCommission;
                            return rep;
                        })
                        .OrderByDescending(x => x.TotalCommission)
                        .ToList();

                    // 去重處理
                    var filteredReps = repData
                        .GroupBy(r => r.SalesRep)
                        .Select(g => g.First())
                        .OrderByDescending(r => r.TotalCommission)
                        .ToList();

                    // 更新排名
                    for (int i = 0; i < filteredReps.Count; i++)
                    {
                        filteredReps[i].Rank = i + 1;
                    }

                    if (filteredReps.Any())
                    {
                        MainThread.BeginInvokeOnMainThread(() => {
                            SalesRepData = new ObservableCollection<SalesLeaderboardItem>(filteredReps);
                        });

                        Debug.WriteLine($"從過濾數據計算出 {filteredReps.Count} 條銷售代表數據（已去重）");
                        return;
                    }
                }

                // 如果過濾數據為空，嘗試從緩存獲取
                var cachedSalesRepData = _excelService.GetSalesLeaderboardData();

                if (cachedSalesRepData.Any())
                {
                    // 確保沒有重複的銷售代表
                    var repsFromCache = cachedSalesRepData
                        .GroupBy(r => r.SalesRep)
                        .Select(g => g.First())
                        .OrderByDescending(r => r.TotalCommission)
                        .ToList();

                    // 更新排名
                    for (int i = 0; i < repsFromCache.Count; i++)
                    {
                        repsFromCache[i].Rank = i + 1;
                    }

                    MainThread.BeginInvokeOnMainThread(() => {
                        SalesRepData = new ObservableCollection<SalesLeaderboardItem>(repsFromCache);
                    });

                    Debug.WriteLine($"從緩存載入了 {repsFromCache.Count} 條銷售代表數據（已去重）");
                    return;
                }

                // 如果緩存也沒有數據，載入示例數據
                Debug.WriteLine("過濾後的數據和緩存都為空，載入示例銷售代表數據");
                LoadSampleSalesRepData();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入銷售代表數據時發生錯誤: {ex.Message}");
                LoadSampleSalesRepData();
            }
        }

        // 當所有其他方法都失敗時，載入固定的示例數據
        private void LoadSampleData()
        {
            LoadSampleProductData();
            LoadSampleSalesRepData();
        }

        private void LoadSampleProductData()
        {
            var tempData = new List<ProductSalesData>
            {
                new ProductSalesData
                {
                    ProductType = "Thermal",
                    AgencyCommission = 744855.43m,
                    BuyResellCommission = 116206.36m,
                    TotalCommission = 861061.79m,
                    POValue = 7358201.65m,
                    PercentageOfTotal = 41.0m
                },
                new ProductSalesData
                {
                    ProductType = "Power",
                    AgencyCommission = 296743.08m,
                    BuyResellCommission = 8737.33m,
                    TotalCommission = 305481.01m,
                    POValue = 5466144.65m,
                    PercentageOfTotal = 31.0m
                },
                new ProductSalesData
                {
                    ProductType = "Batts & Caps",
                    AgencyCommission = 250130.95m,
                    BuyResellCommission = 0.00m,
                    TotalCommission = 250130.95m,
                    POValue = 2061423.30m,
                    PercentageOfTotal = 12.0m
                },
                new ProductSalesData
                {
                    ProductType = "Channel",
                    AgencyCommission = 167353.03m,
                    BuyResellCommission = 8323.03m,
                    TotalCommission = 175676.06m,
                    POValue = 1416574.65m,
                    PercentageOfTotal = 8.0m
                },
                new ProductSalesData
                {
                    ProductType = "Service",
                    AgencyCommission = 101556.42m,
                    BuyResellCommission = 0.00m,
                    TotalCommission = 101556.42m,
                    POValue = 1272318.58m,
                    PercentageOfTotal = 7.0m
                }
            };

            // 確保數據不重複 - 將樣本數據直接設置為有序版本
            MainThread.BeginInvokeOnMainThread(() => {
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
                    AgencyCommission = 350186.00m,
                    BuyResellCommission = 0.00m,
                    TotalCommission = 350186.00m
                },
                new SalesLeaderboardItem
                {
                    Rank = 2,
                    SalesRep = "Brandon",
                    AgencyCommission = 301802.40m,
                    BuyResellCommission = 38165.70m,
                    TotalCommission = 339968.10m
                },
                new SalesLeaderboardItem
                {
                    Rank = 3,
                    SalesRep = "Chris",
                    AgencyCommission = 186411.10m,
                    BuyResellCommission = 0.00m,
                    TotalCommission = 186411.10m
                },
                new SalesLeaderboardItem
                {
                    Rank = 4,
                    SalesRep = "Mark",
                    AgencyCommission = 124680.50m,
                    BuyResellCommission = 18920.30m,
                    TotalCommission = 143600.80m
                },
                new SalesLeaderboardItem
                {
                    Rank = 5,
                    SalesRep = "Nathan",
                    AgencyCommission = 104582.20m,
                    BuyResellCommission = 21060.80m,
                    TotalCommission = 125643.00m
                }
            };

            MainThread.BeginInvokeOnMainThread(() => {
                SalesRepData = new ObservableCollection<SalesLeaderboardItem>(sampleData);
            });

            Debug.WriteLine($"已載入 {sampleData.Count} 條示例銷售代表數據");
        }
    }

    // 用於跟踪銷售代表選擇狀態的類
    public partial class RepSelectionItem : ObservableObject
    {
        public string Name { get; set; }

        [ObservableProperty]
        private bool _isSelected;
    }
}