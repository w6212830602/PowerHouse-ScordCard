using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ScoreCard.Models;
using ScoreCard.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace ScoreCard.ViewModels
{
    public partial class SalesAnalysisViewModel : ObservableObject
    {

        // 服務依賴注入
        private readonly IExcelService _excelService;
        private readonly ITargetService _targetService;
        private List<SalesData> _allSalesData;
        private List<SalesData> _filteredData;

        #region 屬性

        // 日期範圍控制
        [ObservableProperty]
        private DateTime _startDate = DateTime.Now.AddMonths(-3);

        [ObservableProperty]
        private DateTime _endDate = DateTime.Now;

        [ObservableProperty]
        private bool _isAllStatus = true; // Default to All status selected

        // 載入狀態
        [ObservableProperty]
        private bool _isLoading;

        // 視圖類型控制
        [ObservableProperty]
        private bool _isSummaryView = true;

        [ObservableProperty]
        private string _viewType = "ByProduct"; // ByProduct, ByRep - removed ByDeptLOB

        [ObservableProperty]
        private bool _isBookedStatus = false;

        // Add new properties instead of computed ones
        [ObservableProperty]
        private bool _isInProgressStatus = false;

        [ObservableProperty]
        private bool _isCompletedStatus = false;

        [ObservableProperty]
        private string _status = "All"; // Default status

        // Computed properties for status checking
        public bool IsBookedStatusActive => Status == "Booked";
        public bool IsInProgressStatusActive => Status == "InProgress";
        public bool IsCompletedStatusActive => Status == "Completed";


        // 摘要數據
        [ObservableProperty]
        private SalesAnalysisSummary _summary;

        // 圖表數據
        [ObservableProperty]
        private ObservableCollection<ChartData> _targetVsAchievementData = new();

        [ObservableProperty]
        private ObservableCollection<ChartData> _achievementTrendData = new();

        [ObservableProperty]
        private double _yAxisMaximum = 20;

        // 產品數據
        [ObservableProperty]
        private ObservableCollection<ProductSalesData> _productSalesData = new();

        // 銷售代表排行榜
        [ObservableProperty]
        private ObservableCollection<SalesLeaderboardItem> _salesLeaderboard = new();

        public bool IsAllStatusActive => Status == "All";

        // 視圖計算屬性
        public bool IsProductView => ViewType == "ByProduct";
        public bool IsRepView => ViewType == "ByRep";
        // Removed IsDeptLobView property

        #endregion

        #region 構造函數

        public SalesAnalysisViewModel(IExcelService excelService, ITargetService targetService)
        {
            Debug.WriteLine("初始化 SalesAnalysisViewModel");
            _excelService = excelService;
            _targetService = targetService;

            // 初始化摘要數據
            Summary = new SalesAnalysisSummary();

            // 訂閱服務的事件通知
            _excelService.DataUpdated += OnDataUpdated;
            _targetService.TargetsUpdated += OnTargetsUpdated;

            // 初始化資料
            InitializeAsync();
        }

        #endregion

        #region 命令

        // 切換視圖（摘要/詳細）
        [RelayCommand]
        private async Task SwitchView(string viewType)
        {
            if (viewType.ToLower() == "summary")
            {
                IsSummaryView = true;
                await ReloadDataAsync();
            }
            else if (viewType.ToLower() == "detailed")
            {
                IsSummaryView = false;
                await Shell.Current.GoToAsync("DetailedAnalysis");
            }
        }

        // 變更視圖類型（產品/代表）
        [RelayCommand]
        private async Task ChangeViewType(string newViewType)
        {
            if (!string.IsNullOrEmpty(newViewType) && ViewType != newViewType)
            {
                ViewType = newViewType;
                OnPropertyChanged(nameof(IsProductView));
                OnPropertyChanged(nameof(IsRepView));
                await LoadLeaderboardDataAsync();
            }
            else if (ViewType == newViewType)
            {
                // 即使視圖類型相同，也重新載入數據
                await LoadLeaderboardDataAsync();
            }
        }

        // 變更狀態過濾器（Booked/Completed）
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
                        IsCompletedStatus = false;
                        Status = "All";
                        changed = true;
                    }
                    break;

                case "booked":
                    if (!IsBookedStatus)
                    {
                        IsAllStatus = false;
                        IsBookedStatus = true;
                        IsInProgressStatus = false;
                        IsCompletedStatus = false;
                        Status = "Booked";
                        changed = true;
                    }
                    break;

                case "inprogress":
                    if (!IsInProgressStatus)
                    {
                        IsAllStatus = false;
                        IsBookedStatus = false;
                        IsInProgressStatus = true;
                        IsCompletedStatus = false;
                        Status = "InProgress";
                        changed = true;
                    }
                    break;

                case "completed":
                    if (!IsCompletedStatus)
                    {
                        IsAllStatus = false;
                        IsBookedStatus = false;
                        IsInProgressStatus = false;
                        IsCompletedStatus = true;
                        Status = "Completed";
                        changed = true;
                    }
                    break;
            }

            if (changed)
            {
                Debug.WriteLine($"Status changed to: {status}, reloading data");
                await FilterData(); // Direct method call instead of using Command
            }
            else
            {
                Debug.WriteLine($"Status not changed, still: {status}");
            }
        }

        [RelayCommand]
        private async Task FilterData()
        {
            try
            {
                IsLoading = true;
                Debug.WriteLine("执行数据过滤...");

                // 过滤数据
                FilterDataByDateRange();

                // 加载排行榜数据
                await LoadLeaderboardDataAsync();

                Debug.WriteLine("数据过滤完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"过滤数据时出错: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }




        // 導航到詳細分析頁面
        [RelayCommand]
        private async Task NavigateToDetailed()
        {
            try
            {
                await Shell.Current.GoToAsync("//DetailedAnalysis");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"導航錯誤: {ex.Message}");
                await Application.Current.MainPage.DisplayAlert("導航錯誤", $"無法導航至詳細頁面: {ex.Message}", "確定");
            }
        }

        #endregion

        #region 數據載入方法

        private async void InitializeAsync()
        {
            try
            {
                Debug.WriteLine("開始初始化 Dashboard");

                // Initialize target service
                await _targetService.InitializeAsync();

                // 載入 Excel 數據
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                _allSalesData = data ?? new List<SalesData>();

                // 設定日期範圍（如果未設置）
                if (StartDate == default || EndDate == default)
                {
                    var dates = _allSalesData.Select(x => x.ReceivedDate).OrderBy(x => x).ToList();
                    if (dates.Any())
                    {
                        EndDate = dates.Last().Date;
                        StartDate = EndDate.AddMonths(-3).Date;
                    }
                    else
                    {
                        EndDate = DateTime.Now.Date;
                        StartDate = EndDate.AddMonths(-3).Date;
                    }
                }

                // 初始載入數據
                await ReloadDataAsync();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"初始化錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
            }
            finally
            {
                IsLoading = false;
            }
        }

        // 重新載入所有數據
        // Modified ReloadDataAsync method in SalesAnalysisViewModel.cs
        public async Task ReloadDataAsync()
        {
            try
            {
                IsLoading = true;

                // Reload raw data if needed
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    try
                    {
                        var (data, lastUpdated) = await _excelService.LoadDataAsync();
                        _allSalesData = data ?? new List<SalesData>();
                    }
                    catch (Exception ex)
                    {
                        await MainThread.InvokeOnMainThreadAsync(async () =>
                        {
                            await Application.Current.MainPage.DisplayAlert(
                                "Read Error",
                                $"Unable to read Excel file: {ex.Message}",
                                "OK");
                        });
                        IsLoading = false;

                        // Clear all data displays instead of showing sample data
                        ClearAllDisplays();
                        return;
                    }
                }

                // Filter data
                FilterDataByDateRange();

                // Load summary data
                await LoadSummaryDataAsync();

                // Load chart data
                await LoadChartDataAsync();

                // Load leaderboard data
                await LoadLeaderboardDataAsync();

                // Update chart axes
                UpdateChartAxes();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reloading data: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await MainThread.InvokeOnMainThreadAsync(async () =>
                {
                    await Application.Current.MainPage.DisplayAlert(
                        "Error",
                        $"Error loading data: {ex.Message}",
                        "OK");
                });

                // Clear all data displays instead of showing sample data
                ClearAllDisplays();
            }
            finally
            {
                IsLoading = false;
            }
        }

        // New method to clear all data displays
        private void ClearAllDisplays()
        {
            MainThread.BeginInvokeOnMainThread(() =>
            {
                // Clear chart data
                TargetVsAchievementData = new ObservableCollection<ChartData>();
                AchievementTrendData = new ObservableCollection<ChartData>();

                // Clear product data
                ProductSalesData = new ObservableCollection<ProductSalesData>();

                // Clear sales rep data
                SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>();

                // Set default axes values
                YAxisMaximum = 5;
            });
        }

        // 過濾數據按日期範圍
        private void FilterDataByDateRange()
        {
            try
            {
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    Debug.WriteLine("No data available for filtering");
                    _filteredData = new List<SalesData>();
                    return;
                }

                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1); // Include the entire end date

                Debug.WriteLine($"Filtering date range: {startDate:yyyy-MM-dd} to {endDate:yyyy-MM-dd}");

                // 準備三個不同的數據集
                List<SalesData> bookedData = new List<SalesData>();
                List<SalesData> inProgressData = new List<SalesData>();
                List<SalesData> completedData = new List<SalesData>();

                // 無論選擇什麼狀態，總是提取所有三種狀態的數據
                // Booked數據: A列在日期範圍內，Y列為空，N列有值
                bookedData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate.Date &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission > 0)
                    .ToList();

                // In Progress數據: A列在日期範圍內，Y列和N列均為空
                inProgressData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate.Date &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission == 0)
                    .ToList();

                // Completed數據: Y列在日期範圍內且不為空
                completedData = _allSalesData
                    .Where(x => x.CompletionDate.HasValue &&
                           x.CompletionDate.Value.Date >= startDate &&
                           x.CompletionDate.Value.Date <= endDate.Date)
                    .ToList();

                Debug.WriteLine($"找到 Booked數據: {bookedData.Count} 條, In Progress數據: {inProgressData.Count} 條, Completed數據: {completedData.Count} 條");

                // 根據選擇的狀態過濾數據
                if (IsAllStatus)
                {
                    // All狀態: 合併所有數據
                    _filteredData = new List<SalesData>();
                    _filteredData.AddRange(bookedData);
                    _filteredData.AddRange(inProgressData);
                    _filteredData.AddRange(completedData);
                    Debug.WriteLine($"[All] 合併過濾後的記錄: {_filteredData.Count}");
                }
                else if (IsBookedStatus)
                {
                    _filteredData = bookedData;
                    Debug.WriteLine($"[Booked] 過濾後的記錄: {_filteredData.Count}");
                }
                else if (IsInProgressStatus)
                {
                    _filteredData = inProgressData;
                    Debug.WriteLine($"[In Progress] 過濾後的記錄: {_filteredData.Count}");

                    // 輸出樣本記錄用於調試
                    foreach (var item in _filteredData.Take(Math.Min(5, _filteredData.Count)))
                    {
                        Debug.WriteLine($"In Progress樣本: " +
                            $"ReceivedDate={item.ReceivedDate:yyyy-MM-dd}, " +
                            $"POValue=${item.POValue:N2}, " +
                            $"預期佣金(POValue*0.12)=${item.POValue * 0.12m:N2}");
                    }
                }
                else if (IsCompletedStatus)
                {
                    _filteredData = completedData;
                    Debug.WriteLine($"[Completed] 過濾後的記錄: {_filteredData.Count}");
                }
                else
                {
                    // 默認: 空數據
                    _filteredData = new List<SalesData>();
                    Debug.WriteLine("未選擇狀態，返回空數據");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"過濾數據時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                _filteredData = new List<SalesData>();
            }
        }


        // 載入摘要數據
        // 修改LoadSummaryDataAsync方法中的數值轉換部分
        private async Task LoadSummaryDataAsync()
        {
            try
            {
                var currentDate = DateTime.Now;
                var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;

                // 從目標服務獲取目標
                var companyTarget = _targetService.GetCompanyTarget(currentFiscalYear);
                decimal targetValue = companyTarget?.AnnualTarget ?? 4000000m;

                // 使用所有數據，處理三種狀態 (Booked、In Progress、Invoiced)
                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1);

                // 已完成的數據（使用完成日期）
                decimal completedAchievement = _allSalesData
                    .Where(x => x.CompletionDate.HasValue &&
                            x.CompletionDate.Value.Date >= startDate &&
                            x.CompletionDate.Value.Date <= endDate)
                    .Sum(x => x.TotalCommission);

                // 已預訂但未完成的數據（N欄有值）
                decimal bookedMarginNotInvoiced = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission > 0) // 有佣金的記錄
                    .Sum(x => x.TotalCommission);

                // In Progress的預期佣金（接收日期在範圍內，Y列為空，N列為0）
                decimal inProgressExpectedMargin = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission == 0) // 沒有佣金的記錄
                    .Sum(x => x.VertivValue * 0.12m); // 計算預期佣金（Vertiv值的12%）

                // 所有達成的總和（已完成 + 已預訂 + 預期）
                decimal totalAchievement = completedAchievement + bookedMarginNotInvoiced + inProgressExpectedMargin;

                // 計算實際剩餘金額
                decimal actualRemaining = Math.Max(0, targetValue - totalAchievement);

                // 計算「Remaining to Target」（減去已預訂但未完成的訂單）
                decimal remainingToTarget = Math.Max(0, actualRemaining - bookedMarginNotInvoiced - inProgressExpectedMargin);

                // 計算百分比
                decimal achievementPercentage = 0;
                decimal marginPercentage = 0;
                decimal remainingTargetPercentage = 0;

                if (targetValue > 0)
                {
                    achievementPercentage = Math.Round((totalAchievement / targetValue) * 100, 1);
                    remainingTargetPercentage = Math.Round((actualRemaining / targetValue) * 100, 1);
                    marginPercentage = Math.Round(((bookedMarginNotInvoiced + inProgressExpectedMargin) / targetValue) * 100, 1);
                }

                // 轉換為百萬單位 - 使用更高精度的四捨五入，保留到千位數
                decimal targetValueM = Math.Round(targetValue / 1000000m, 6);
                decimal totalAchievementM = Math.Round(totalAchievement / 1000000m, 6);
                decimal remainingToTargetM = Math.Round(remainingToTarget / 1000000m, 6);
                decimal actualRemainingM = Math.Round(actualRemaining / 1000000m, 6);

                // 轉換已Booked但未Invoiced的Margin為百萬單位
                decimal bookedMarginNotInvoicedM = Math.Round((bookedMarginNotInvoiced + inProgressExpectedMargin) / 1000000m, 6);

                // 更新UI
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    if (Summary == null)
                    {
                        Summary = new SalesAnalysisSummary();
                    }
                    Summary.TotalTarget = targetValueM;
                    Summary.TotalAchievement = totalAchievementM;
                    Summary.TotalMargin = bookedMarginNotInvoicedM; // Booked Margin not Invoiced + In Progress預期
                    Summary.RemainingTarget = remainingToTargetM; // 原來屬性用於Remaining to target
                    Summary.ActualRemaining = actualRemainingM; // 新屬性用於Actual Remaining
                    Summary.AchievementPercentage = achievementPercentage;
                    Summary.MarginPercentage = marginPercentage; // 使用Booked Margin的比率
                    Summary.RemainingTargetPercentage = remainingTargetPercentage;
                });

                Debug.WriteLine($"摘要數據: Target=${targetValueM}M, Achievement=${totalAchievementM}M ({achievementPercentage}%), " +
                               $"Remaining to target=${remainingToTargetM}M, Actual Remaining=${actualRemainingM}M, " +
                               $"Booked Margin not Invoiced=${bookedMarginNotInvoicedM}M ({marginPercentage}%)");

                // 輸出詳細數據用於調試
                Debug.WriteLine($"已完成 (Invoiced): ${completedAchievement:N2}");
                Debug.WriteLine($"已預訂 (Booked): ${bookedMarginNotInvoiced:N2}");
                Debug.WriteLine($"進行中 (In Progress): ${inProgressExpectedMargin:N2}");
                Debug.WriteLine($"總達成: ${totalAchievement:N2}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入摘要數據時發生錯誤: {ex.Message}");
                // 確保摘要不為空
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    if (Summary == null)
                    {
                        Summary = new SalesAnalysisSummary
                        {
                            TotalTarget = 10,
                            TotalAchievement = 5.123m,
                            TotalMargin = 3.456m,
                            RemainingTarget = 5.789m,
                            ActualRemaining = 5.321m,
                            AchievementPercentage = 50,
                            MarginPercentage = 30,
                            RemainingTargetPercentage = 50
                        };
                    }
                });
            }
        }

        // 載入圖表數據
        private async Task LoadChartDataAsync()
        {
            try
            {
                if (_filteredData == null || !_filteredData.Any())
                {
                    // 確保集合不為 null，而是一個空集合
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        TargetVsAchievementData = new ObservableCollection<ChartData>();
                        AchievementTrendData = new ObservableCollection<ChartData>();
                        // 設置一個合理的 Y 軸最大值
                        YAxisMaximum = 5;
                    });

                    Debug.WriteLine("沒有數據可用於圖表，顯示空圖表");
                    return;
                }

                // 建立月度數據集
                var monthlyData = new List<dynamic>();

                // 根據選中的狀態，創建不同的月度數據
                // 準備三個不同的數據集
                var bookedData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= StartDate.Date &&
                           x.ReceivedDate.Date <= EndDate.Date.AddDays(1).AddSeconds(-1) &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission > 0)
                    .ToList();

                var inProgressData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= StartDate.Date &&
                           x.ReceivedDate.Date <= EndDate.Date.AddDays(1).AddSeconds(-1) &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission == 0)
                    .ToList();

                var invoicedData = _allSalesData
                    .Where(x => x.CompletionDate.HasValue &&
                           x.CompletionDate.Value.Date >= StartDate.Date &&
                           x.CompletionDate.Value.Date <= EndDate.Date.AddDays(1).AddSeconds(-1))
                    .ToList();

                // 建立月份分組，依據狀態處理不同的數據
                var allMonths = new HashSet<(int Year, int Month)>();

                // 添加所有獨特的月份
                foreach (var item in bookedData)
                {
                    allMonths.Add((item.ReceivedDate.Year, item.ReceivedDate.Month));
                }

                foreach (var item in inProgressData)
                {
                    allMonths.Add((item.ReceivedDate.Year, item.ReceivedDate.Month));
                }

                foreach (var item in invoicedData)
                {
                    allMonths.Add((item.CompletionDate.Value.Year, item.CompletionDate.Value.Month));
                }

                // 對每個月份進行處理
                foreach (var (year, month) in allMonths.OrderBy(m => m.Year).ThenBy(m => m.Month))
                {
                    // 計算不同狀態的數據
                    var bookedForMonth = bookedData
                        .Where(x => x.ReceivedDate.Year == year && x.ReceivedDate.Month == month)
                        .ToList();

                    var inProgressForMonth = inProgressData
                        .Where(x => x.ReceivedDate.Year == year && x.ReceivedDate.Month == month)
                        .ToList();

                    var invoicedForMonth = invoicedData
                        .Where(x => x.CompletionDate.Value.Year == year && x.CompletionDate.Value.Month == month)
                        .ToList();

                    // 計算各種數據
                    decimal bookedCommission = bookedForMonth.Sum(x => x.TotalCommission);
                    decimal inProgressEstimate = inProgressForMonth.Sum(x => x.VertivValue * 0.12m);
                    decimal invoicedCommission = invoicedForMonth.Sum(x => x.TotalCommission);

                    // 計算總和
                    decimal totalCommission = 0;
                    decimal totalPOValue = 0;

                    // 根據狀態選擇要包含的數據
                    if (IsAllStatus)
                    {
                        totalCommission = bookedCommission + inProgressEstimate + invoicedCommission;
                        totalPOValue = bookedForMonth.Sum(x => x.POValue) +
                                      inProgressForMonth.Sum(x => x.POValue) +
                                      invoicedForMonth.Sum(x => x.POValue);
                    }
                    else if (IsBookedStatus)
                    {
                        totalCommission = bookedCommission;
                        totalPOValue = bookedForMonth.Sum(x => x.POValue);
                    }
                    else if (IsInProgressStatus)
                    {
                        totalCommission = inProgressEstimate;
                        totalPOValue = inProgressForMonth.Sum(x => x.POValue);
                    }
                    else if (IsCompletedStatus)
                    {
                        totalCommission = invoicedCommission;
                        totalPOValue = invoicedForMonth.Sum(x => x.POValue);
                    }

                    // 添加到月度數據
                    monthlyData.Add(new
                    {
                        YearMonth = new { Year = year, Month = month },
                        POValue = totalPOValue,
                        MarginValue = totalCommission,
                        CommissionValue = totalCommission
                    });
                }

                // 檢查處理後的資料
                Debug.WriteLine($"處理後的月度資料筆數: {monthlyData.Count}");

                // 準備圖表資料
                var newTargetVsAchievementData = new ObservableCollection<ChartData>();
                var newAchievementTrendData = new ObservableCollection<ChartData>();

                foreach (var month in monthlyData)
                {
                    var label = $"{month.YearMonth.Year}/{month.YearMonth.Month:D2}";

                    // 所有模式下都添加預期佣金到總佣金中
                    decimal commissionValueInMillions = Math.Round(month.CommissionValue / 1000000m, 2);
                    decimal marginValueInMillions = Math.Round(month.MarginValue / 1000000m, 2);

                    // 從目標服務獲取該月份的目標值
                    int fiscalYear = month.YearMonth.Month >= 8 ? month.YearMonth.Year + 1 : month.YearMonth.Year;
                    int quarter = GetQuarterFromMonth(month.YearMonth.Month);
                    decimal targetValue = _targetService.GetCompanyQuarterlyTarget(fiscalYear, quarter) / 3; // 將季度目標平均分配到月
                    decimal targetValueInMillions = Math.Round(targetValue / 1000000m, 2);

                    // 添加到目標與達成對比圖表
                    newTargetVsAchievementData.Add(new ChartData
                    {
                        Label = label,
                        Target = targetValueInMillions,
                        Achievement = commissionValueInMillions
                    });

                    // 添加到達成趨勢圖表
                    newAchievementTrendData.Add(new ChartData
                    {
                        Label = label,
                        Target = targetValueInMillions,
                        Achievement = commissionValueInMillions
                    });
                }

                // 記錄生成的圖表資料
                Debug.WriteLine($"生成 Target vs Achievement 資料: {newTargetVsAchievementData.Count} 筆");
                Debug.WriteLine($"生成 Achievement Trend 資料: {newAchievementTrendData.Count} 筆");

                // 更新 UI
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    TargetVsAchievementData = newTargetVsAchievementData;
                    AchievementTrendData = newAchievementTrendData;

                    // 設置一個合理的 Y 軸最大值
                    double maxTarget = 0;
                    double maxAchievement = 0;

                    if (newTargetVsAchievementData.Any())
                    {
                        maxTarget = (double)newTargetVsAchievementData.Max(d => d.Target);
                        maxAchievement = (double)newTargetVsAchievementData.Max(d => d.Achievement);
                    }

                    double maxValue = Math.Max(maxTarget, maxAchievement);
                    maxValue = Math.Max(maxValue * 1.2, 5); // 增加 20% 空間，但確保至少為 5
                    YAxisMaximum = Math.Ceiling(maxValue);

                    Debug.WriteLine($"設置 Y 軸最大值: {YAxisMaximum}");
                });

                Debug.WriteLine("圖表資料加載完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入圖表資料時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 確保發生錯誤時圖表資料不為 null
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    TargetVsAchievementData = new ObservableCollection<ChartData>();
                    AchievementTrendData = new ObservableCollection<ChartData>();
                    YAxisMaximum = 5;
                });
            }
        }

        // 載入排行榜數據
        private async Task LoadLeaderboardDataAsync()
        {
            try
            {
                IsLoading = true;

                // 使用之前已經過濾好的數據
                if (_filteredData == null || !_filteredData.Any())
                {
                    Debug.WriteLine("沒有過濾後的數據可用於加載排行榜");

                    // 根據視圖類型清空數據
                    if (IsProductView)
                    {
                        ProductSalesData = new ObservableCollection<ProductSalesData>();
                    }
                    else if (IsRepView)
                    {
                        SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>();
                    }

                    return;
                }

                Debug.WriteLine($"用於計算的記錄數: {_filteredData.Count} 條");

                // 檢查當前視圖和狀態
                bool isAllStatus = IsAllStatus;
                bool isInProgressStatus = IsInProgressStatus;
                bool isBookedStatus = IsBookedStatus;
                bool isCompletedStatus = IsCompletedStatus;
                bool isProductView = IsProductView;
                bool isRepView = IsRepView;

                Debug.WriteLine($"當前視圖和狀態: IsProductView={isProductView}, IsRepView={isRepView}, " +
                              $"IsAllStatus={isAllStatus}, IsInProgressStatus={isInProgressStatus}, " +
                              $"IsBookedStatus={isBookedStatus}, IsCompletedStatus={isCompletedStatus}");

                // 如果我們處於All視圖，則一開始就提取所有三種狀態的數據
                List<SalesData> allData = _allSalesData;
                List<SalesData> bookedData = new List<SalesData>();
                List<SalesData> inProgressData = new List<SalesData>();
                List<SalesData> completedData = new List<SalesData>();

                // 使用共同的日期範圍過濾
                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1);

                if (allData != null && allData.Any())
                {
                    // 提取所有三種狀態的數據
                    bookedData = allData
                        .Where(x => x.ReceivedDate.Date >= startDate &&
                               x.ReceivedDate.Date <= endDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission > 0)
                        .ToList();

                    inProgressData = allData
                        .Where(x => x.ReceivedDate.Date >= startDate &&
                               x.ReceivedDate.Date <= endDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission == 0)
                        .ToList();

                    completedData = allData
                        .Where(x => x.CompletionDate.HasValue &&
                               x.CompletionDate.Value.Date >= startDate &&
                               x.CompletionDate.Value.Date <= endDate.Date)
                        .ToList();

                    Debug.WriteLine($"分離數據: Booked={bookedData.Count}, InProgress={inProgressData.Count}, Completed={completedData.Count}");
                }

                // 根據視圖類型和狀態加載相應數據
                if (isProductView)
                {
                    if (isAllStatus)
                    {
                        // 創建一個合併的數據集
                        List<SalesData> combinedData = new List<SalesData>();
                        combinedData.AddRange(bookedData);
                        combinedData.AddRange(inProgressData);
                        combinedData.AddRange(completedData);

                        await LoadProductData(combinedData);
                        Debug.WriteLine("加載了All狀態的產品數據");
                    }
                    else if (isBookedStatus)
                    {
                        await LoadProductData(bookedData);
                        Debug.WriteLine("加載了Booked狀態的產品數據");
                    }
                    else if (isInProgressStatus)
                    {
                        await LoadProductData(inProgressData);
                        Debug.WriteLine("加載了In Progress狀態的產品數據");
                    }
                    else if (isCompletedStatus)
                    {
                        await LoadProductData(completedData);
                        Debug.WriteLine("加載了Completed狀態的產品數據");
                    }
                    else
                    {
                        // 默認情況
                        await LoadProductData(_filteredData);
                        Debug.WriteLine("加載了過濾後的產品數據");
                    }
                }
                else if (isRepView)
                {
                    if (isAllStatus)
                    {
                        // 創建一個合併的數據集
                        List<SalesData> combinedData = new List<SalesData>();
                        combinedData.AddRange(bookedData);
                        combinedData.AddRange(inProgressData);
                        combinedData.AddRange(completedData);

                        await LoadSalesRepData(combinedData);
                        Debug.WriteLine("加載了All狀態的銷售代表數據");
                    }
                    else if (isBookedStatus)
                    {
                        await LoadSalesRepData(bookedData);
                        Debug.WriteLine("加載了Booked狀態的銷售代表數據");
                    }
                    else if (isInProgressStatus)
                    {
                        await LoadSalesRepData(inProgressData);
                        Debug.WriteLine("加載了In Progress狀態的銷售代表數據");
                    }
                    else if (isCompletedStatus)
                    {
                        await LoadSalesRepData(completedData);
                        Debug.WriteLine("加載了Completed狀態的銷售代表數據");
                    }
                    else
                    {
                        // 默認情況
                        await LoadSalesRepData(_filteredData);
                        Debug.WriteLine("加載了過濾後的銷售代表數據");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加載排行榜數據時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 出錯時清空數據
                if (IsProductView)
                {
                    ProductSalesData = new ObservableCollection<ProductSalesData>();
                }
                else if (IsRepView)
                {
                    SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>();
                }
            }
            finally
            {
                IsLoading = false;
            }
        }

        // 在 SalesAnalysisViewModel.cs 中添加以下方法，或替換現有的同名方法
        private async Task LoadProductData(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine($"載入產品數據，記錄數: {data.Count}");

                // 檢查當前視圖狀態
                bool isAllStatus = IsAllStatus;
                bool isInProgressStatus = IsInProgressStatus;
                bool isBookedStatus = IsBookedStatus;
                bool isCompletedStatus = IsCompletedStatus;

                Debug.WriteLine($"當前視圖狀態: All={isAllStatus}, InProgress={isInProgressStatus}, Booked={isBookedStatus}, Completed={isCompletedStatus}");

                // 分離不同狀態的數據，以便進行後續處理
                List<SalesData> bookedData = new List<SalesData>();
                List<SalesData> inProgressData = new List<SalesData>();
                List<SalesData> completedData = new List<SalesData>();

                // 我們永遠需要按狀態分離數據
                if (data != null && data.Any())
                {
                    bookedData = data.Where(x => !x.CompletionDate.HasValue && x.TotalCommission > 0).ToList();
                    inProgressData = data.Where(x => !x.CompletionDate.HasValue && x.TotalCommission == 0).ToList();
                    completedData = data.Where(x => x.CompletionDate.HasValue).ToList();

                    Debug.WriteLine($"數據分佈: Booked={bookedData.Count}, InProgress={inProgressData.Count}, Invoiced={completedData.Count}");
                }

                List<ProductSalesData> products = new List<ProductSalesData>();

                if (isAllStatus)
                {
                    // 對於All視圖，我們需要合併所有狀態的數據
                    // 首先按產品類型分組處理所有數據
                    var allProductTypes = new HashSet<string>();

                    // 收集所有不同的產品類型
                    foreach (var item in bookedData)
                    {
                        allProductTypes.Add(NormalizeProductType(item.ProductType));
                    }
                    foreach (var item in inProgressData)
                    {
                        allProductTypes.Add(NormalizeProductType(item.ProductType));
                    }
                    foreach (var item in completedData)
                    {
                        allProductTypes.Add(NormalizeProductType(item.ProductType));
                    }

                    // 移除空的產品類型
                    allProductTypes.Remove(string.Empty);
                    allProductTypes.Remove(null);

                    // 對每種產品類型分別處理三種狀態的數據
                    foreach (var productType in allProductTypes)
                    {
                        // Booked數據 (CompletionDate為空，TotalCommission有值)
                        var bookedForProduct = bookedData.Where(x => NormalizeProductType(x.ProductType) == productType).ToList();
                        decimal bookedAgencyMargin = bookedForProduct.Sum(x => x.AgencyMargin);
                        decimal bookedBuyResellMargin = bookedForProduct.Sum(x => x.BuyResellValue);
                        decimal bookedTotalMargin = bookedForProduct.Sum(x => x.TotalCommission);
                        decimal bookedVertivValue = bookedForProduct.Sum(x => x.VertivValue);

                        // In Progress數據 (CompletionDate為空，TotalCommission為0)
                        var inProgressForProduct = inProgressData.Where(x => NormalizeProductType(x.ProductType) == productType).ToList();
                        decimal expectedCommission = inProgressForProduct.Sum(x => x.VertivValue * 0.12m);
                        decimal inProgressVertivValue = inProgressForProduct.Sum(x => x.VertivValue);

                        // Completed數據 (CompletionDate有值)
                        var completedForProduct = completedData.Where(x => NormalizeProductType(x.ProductType) == productType).ToList();
                        decimal completedAgencyMargin = completedForProduct.Sum(x => x.AgencyMargin);
                        decimal completedBuyResellMargin = completedForProduct.Sum(x => x.BuyResellValue);
                        decimal completedTotalMargin = completedForProduct.Sum(x => x.TotalCommission);
                        decimal completedVertivValue = completedForProduct.Sum(x => x.VertivValue);

                        // 合併所有狀態的數據
                        decimal totalAgencyMargin = bookedAgencyMargin + expectedCommission + completedAgencyMargin;
                        decimal totalBuyResellMargin = bookedBuyResellMargin + completedBuyResellMargin; // In Progress的Buy Resell Margin為0
                        decimal totalMargin = bookedTotalMargin + expectedCommission + completedTotalMargin;
                        decimal totalVertivValue = bookedVertivValue + inProgressVertivValue + completedVertivValue;

                        products.Add(new ProductSalesData
                        {
                            ProductType = productType,
                            AgencyMargin = Math.Round(totalAgencyMargin, 2),
                            BuyResellMargin = Math.Round(totalBuyResellMargin, 2),
                            TotalMargin = Math.Round(totalMargin, 2),
                            VertivValue = Math.Round(totalVertivValue, 2),
                            POValue = Math.Round(totalVertivValue, 2) // POValue與VertivValue相同
                        });

                        Debug.WriteLine($"All視圖 - 產品: {productType}, " +
                                       $"總Agency: ${totalAgencyMargin:N2}, " +
                                       $"總BuyResell: ${totalBuyResellMargin:N2}, " +
                                       $"總Margin: ${totalMargin:N2}, " +
                                       $"總VertivValue: ${totalVertivValue:N2}");
                    }
                }
                else
                {
                    // 對於特定狀態的視圖，使用對應狀態的數據
                    List<SalesData> dataToProcess = new List<SalesData>();

                    if (isBookedStatus)
                    {
                        dataToProcess = bookedData;
                    }
                    else if (isInProgressStatus)
                    {
                        dataToProcess = inProgressData;
                    }
                    else if (isCompletedStatus)
                    {
                        dataToProcess = completedData;
                    }
                    else
                    {
                        // 默認情況下使用所有數據
                        dataToProcess = data;
                    }

                    var productGroups = dataToProcess
                        .GroupBy(x => NormalizeProductType(x.ProductType))
                        .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                        .ToList();

                    foreach (var group in productGroups)
                    {
                        decimal agencyMargin, buyResellMargin, totalMargin, vertivValue;

                        if (isInProgressStatus)
                        {
                            // In Progress模式下使用預期佣金
                            agencyMargin = Math.Round(group.Sum(x => x.VertivValue * 0.12m), 2);
                            buyResellMargin = 0; // In Progress模式下Buy Resell Margin為0
                            totalMargin = agencyMargin; // 總佣金等於預期佣金
                            vertivValue = Math.Round(group.Sum(x => x.VertivValue), 2);
                        }
                        else
                        {
                            // Booked或Completed模式下使用實際數據
                            agencyMargin = Math.Round(group.Sum(x => x.AgencyMargin), 2);
                            buyResellMargin = Math.Round(group.Sum(x => x.BuyResellValue), 2);
                            totalMargin = Math.Round(group.Sum(x => x.TotalCommission), 2);
                            vertivValue = Math.Round(group.Sum(x => x.VertivValue), 2);
                        }

                        products.Add(new ProductSalesData
                        {
                            ProductType = group.Key,
                            AgencyMargin = agencyMargin,
                            BuyResellMargin = buyResellMargin,
                            TotalMargin = totalMargin,
                            VertivValue = vertivValue,
                            POValue = vertivValue // POValue與VertivValue相同
                        });

                        Debug.WriteLine($"狀態視圖 - 產品: {group.Key}, " +
                                       $"Agency: ${agencyMargin:N2}, " +
                                       $"BuyResell: ${buyResellMargin:N2}, " +
                                       $"Total: ${totalMargin:N2}, " +
                                       $"VertivValue: ${vertivValue:N2}");
                    }
                }

                // 排序並計算百分比
                products = products.OrderByDescending(x => x.VertivValue).ToList();

                if (products.Any())
                {
                    decimal totalVertivValue = products.Sum(p => p.VertivValue);
                    foreach (var product in products)
                    {
                        product.PercentageOfTotal = totalVertivValue > 0
                            ? Math.Round((product.VertivValue / totalVertivValue) * 100, 1)
                            : 0;
                    }

                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        ProductSalesData = new ObservableCollection<ProductSalesData>(products);
                    });

                    Debug.WriteLine($"成功載入 {products.Count} 個產品數據項目");
                }
                else
                {
                    Debug.WriteLine("沒有產品數據可顯示");
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        ProductSalesData = new ObservableCollection<ProductSalesData>();
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入產品數據時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    ProductSalesData = new ObservableCollection<ProductSalesData>();
                });
            }
        }


        // 載入產品數據
        private async Task LoadProductDataAsync(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine($"Loading product data, record count: {data.Count}");

                // 跟蹤每種狀態的記錄數
                int bookedCount = data.Count(x => !x.CompletionDate.HasValue && x.TotalCommission > 0);
                int inProgressCount = data.Count(x => !x.CompletionDate.HasValue && x.TotalCommission == 0);
                int invoicedCount = data.Count(x => x.CompletionDate.HasValue);

                Debug.WriteLine($"數據分佈: Booked={bookedCount}, InProgress={inProgressCount}, Invoiced={invoicedCount}");

                // 檢查是否處於 In Progress 模式
                bool isInProgressMode = IsInProgressStatus;

                // 計算 In Progress 項目的總預期佣金，用於調試
                decimal totalInProgressCommission = data
                    .Where(x => !x.CompletionDate.HasValue && x.TotalCommission == 0)
                    .Sum(x => x.VertivValue * 0.12m);

                Debug.WriteLine($"In Progress 項目的預期總佣金: ${totalInProgressCommission:N2}");

                var products = data
                    .GroupBy(x => NormalizeProductType(x.ProductType))
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g =>
                    {
                        // 針對 In Progress 模式，計算預期佣金
                        decimal expectedCommission = 0;
                        if (isInProgressMode)
                        {
                            expectedCommission = g.Sum(x => x.VertivValue * 0.12m);
                        }

                        return new ProductSalesData
                        {
                            ProductType = g.Key,
                            // 在 In Progress 模式下，將所有預期佣金都放在 Agency Margin
                            AgencyMargin = Math.Round(isInProgressMode ?
                                expectedCommission : // In Progress 模式 - 使用預期佣金
                                g.Sum(x => x.AgencyMargin), 2), // 其他模式 - 使用實際 Agency Margin
                                                                // Buy Resell Margin 在 In Progress 模式下為 0
                            BuyResellMargin = Math.Round(isInProgressMode ?
                                0 : // In Progress 模式下為 0
                                g.Sum(x => x.BuyResellValue), 2), // 其他模式 - 使用實際 Buy Resell Margin
                                                                  // Total Margin 等於 Agency + Buy Resell
                            TotalMargin = Math.Round(isInProgressMode ?
                                expectedCommission : // In Progress 模式 - 使用預期佣金
                                g.Sum(x => x.TotalCommission), 2), // 其他模式 - 使用實際 Total Commission
                                                                   // 記錄 Vertiv Value
                            VertivValue = Math.Round(g.Sum(x => x.VertivValue), 2),
                            POValue = Math.Round(g.Sum(x => x.POValue), 2),
                            // 標記項目來源
                            IsInProgress = isInProgressMode
                        };
                    })
                    .OrderByDescending(x => x.VertivValue)
                    .ToList();

                Debug.WriteLine($"Number of products after grouping: {products.Count}");

                if (products.Any())
                {
                    // 計算百分比
                    decimal totalPOValue = products.Sum(p => p.VertivValue);
                    foreach (var product in products)
                    {
                        product.PercentageOfTotal = totalPOValue > 0
                            ? Math.Round((product.VertivValue / totalPOValue) * 100, 2)
                            : 0;

                        Debug.WriteLine($"Product: {product.ProductType}, " +
                                       $"Agency: ${product.AgencyMargin}, " +
                                       $"BuyResell: ${product.BuyResellMargin}, " +
                                       $"Total: ${product.TotalMargin}, " +
                                       $"Vertiv Value: ${product.VertivValue}, " +
                                       $"Percentage: {product.PercentageOfTotal}");
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

        // 在 SalesAnalysisViewModel.cs 中添加或替換此方法
        private async Task LoadSalesRepData(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine("載入銷售代表數據...");

                // 檢查當前視圖狀態
                bool isAllStatus = IsAllStatus;
                bool isInProgressStatus = IsInProgressStatus;
                bool isBookedStatus = IsBookedStatus;
                bool isCompletedStatus = IsCompletedStatus;

                Debug.WriteLine($"當前視圖狀態: All={isAllStatus}, InProgress={isInProgressStatus}, Booked={isBookedStatus}, Completed={isCompletedStatus}");

                // 分離不同狀態的數據
                List<SalesData> bookedData = new List<SalesData>();
                List<SalesData> inProgressData = new List<SalesData>();
                List<SalesData> completedData = new List<SalesData>();

                if (data != null && data.Any())
                {
                    bookedData = data.Where(x => !x.CompletionDate.HasValue && x.TotalCommission > 0).ToList();
                    inProgressData = data.Where(x => !x.CompletionDate.HasValue && x.TotalCommission == 0).ToList();
                    completedData = data.Where(x => x.CompletionDate.HasValue).ToList();

                    Debug.WriteLine($"數據分佈: Booked={bookedData.Count}, InProgress={inProgressData.Count}, Completed={completedData.Count}");
                }

                List<SalesLeaderboardItem> reps = new List<SalesLeaderboardItem>();

                if (isAllStatus)
                {
                    // 對於All視圖，我們需要合併所有狀態的數據
                    // 收集所有不同的銷售代表
                    var allReps = new HashSet<string>();

                    foreach (var item in bookedData)
                    {
                        if (!string.IsNullOrWhiteSpace(item.SalesRep))
                            allReps.Add(item.SalesRep);
                    }
                    foreach (var item in inProgressData)
                    {
                        if (!string.IsNullOrWhiteSpace(item.SalesRep))
                            allReps.Add(item.SalesRep);
                    }
                    foreach (var item in completedData)
                    {
                        if (!string.IsNullOrWhiteSpace(item.SalesRep))
                            allReps.Add(item.SalesRep);
                    }

                    // 對每個銷售代表分別處理三種狀態的數據
                    foreach (var salesRep in allReps)
                    {
                        // Booked數據
                        var bookedForRep = bookedData.Where(x => x.SalesRep == salesRep).ToList();
                        decimal bookedAgencyMargin = bookedForRep.Sum(x => x.AgencyMargin);
                        decimal bookedBuyResellMargin = bookedForRep.Sum(x => x.BuyResellValue);
                        decimal bookedTotalMargin = bookedForRep.Sum(x => x.TotalCommission);
                        decimal bookedVertivValue = bookedForRep.Sum(x => x.VertivValue);

                        // In Progress數據 - 計算預期佣金
                        var inProgressForRep = inProgressData.Where(x => x.SalesRep == salesRep).ToList();
                        decimal expectedCommission = inProgressForRep.Sum(x => x.VertivValue * 0.12m);
                        decimal inProgressVertivValue = inProgressForRep.Sum(x => x.VertivValue);

                        // Completed數據
                        var completedForRep = completedData.Where(x => x.SalesRep == salesRep).ToList();
                        decimal completedAgencyMargin = completedForRep.Sum(x => x.AgencyMargin);
                        decimal completedBuyResellMargin = completedForRep.Sum(x => x.BuyResellValue);
                        decimal completedTotalMargin = completedForRep.Sum(x => x.TotalCommission);
                        decimal completedVertivValue = completedForRep.Sum(x => x.VertivValue);

                        // 合併所有狀態的數據
                        decimal totalAgencyMargin = bookedAgencyMargin + expectedCommission + completedAgencyMargin;
                        decimal totalBuyResellMargin = bookedBuyResellMargin + completedBuyResellMargin; // In Progress的Buy Resell Margin為0
                        decimal totalMargin = bookedTotalMargin + expectedCommission + completedTotalMargin;
                        decimal totalVertivValue = bookedVertivValue + inProgressVertivValue + completedVertivValue;

                        reps.Add(new SalesLeaderboardItem
                        {
                            SalesRep = salesRep,
                            AgencyMargin = Math.Round(totalAgencyMargin, 2),
                            BuyResellMargin = Math.Round(totalBuyResellMargin, 2),
                            TotalMargin = Math.Round(totalMargin, 2),
                            VertivValue = Math.Round(totalVertivValue, 2)
                        });

                        Debug.WriteLine($"All視圖 - 銷售代表: {salesRep}, " +
                                       $"總Agency: ${totalAgencyMargin:N2}, " +
                                       $"總BuyResell: ${totalBuyResellMargin:N2}, " +
                                       $"總Margin: ${totalMargin:N2}, " +
                                       $"總VertivValue: ${totalVertivValue:N2}");
                    }
                }
                else
                {
                    // 對於特定狀態的視圖，使用對應狀態的數據
                    List<SalesData> dataToProcess = new List<SalesData>();

                    if (isBookedStatus)
                    {
                        dataToProcess = bookedData;
                    }
                    else if (isInProgressStatus)
                    {
                        dataToProcess = inProgressData;
                    }
                    else if (isCompletedStatus)
                    {
                        dataToProcess = completedData;
                    }
                    else
                    {
                        // 默認情況下使用所有數據
                        dataToProcess = data;
                    }

                    var repGroups = dataToProcess
                        .GroupBy(x => x.SalesRep)
                        .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                        .ToList();

                    foreach (var group in repGroups)
                    {
                        decimal agencyMargin, buyResellMargin, totalMargin, vertivValue;

                        if (isInProgressStatus)
                        {
                            // In Progress模式下使用預期佣金
                            agencyMargin = Math.Round(group.Sum(x => x.VertivValue * 0.12m), 2);
                            buyResellMargin = 0; // In Progress模式下Buy Resell Margin為0
                            totalMargin = agencyMargin; // 總佣金等於預期佣金
                            vertivValue = Math.Round(group.Sum(x => x.VertivValue), 2);
                        }
                        else
                        {
                            // Booked或Completed模式下使用實際數據
                            agencyMargin = Math.Round(group.Sum(x => x.AgencyMargin), 2);
                            buyResellMargin = Math.Round(group.Sum(x => x.BuyResellValue), 2);
                            totalMargin = Math.Round(group.Sum(x => x.TotalCommission), 2);
                            vertivValue = Math.Round(group.Sum(x => x.VertivValue), 2);
                        }

                        reps.Add(new SalesLeaderboardItem
                        {
                            SalesRep = group.Key,
                            AgencyMargin = agencyMargin,
                            BuyResellMargin = buyResellMargin,
                            TotalMargin = totalMargin,
                            VertivValue = vertivValue
                        });

                        Debug.WriteLine($"狀態視圖 - 銷售代表: {group.Key}, " +
                                       $"Agency: ${agencyMargin:N2}, " +
                                       $"BuyResell: ${buyResellMargin:N2}, " +
                                       $"Total: ${totalMargin:N2}, " +
                                       $"VertivValue: ${vertivValue:N2}");
                    }
                }

                // 排序並設置排名
                reps = reps.OrderByDescending(x => x.TotalMargin).ToList();
                for (int i = 0; i < reps.Count; i++)
                {
                    reps[i].Rank = i + 1;
                }

                if (reps.Any())
                {
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>(reps);
                    });

                    Debug.WriteLine($"成功載入 {reps.Count} 個銷售代表數據項目");
                }
                else
                {
                    Debug.WriteLine("沒有銷售代表數據可顯示");
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>();
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入銷售代表數據時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>();
                });
            }
        }

        // 載入銷售代表數據
        private async Task LoadSalesRepDataAsync(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine("Loading sales rep data...");

                // 追蹤數據分布
                int bookedCount = data.Count(x => !x.CompletionDate.HasValue && x.TotalCommission > 0);
                int inProgressCount = data.Count(x => !x.CompletionDate.HasValue && x.TotalCommission == 0);
                int invoicedCount = data.Count(x => x.CompletionDate.HasValue);

                Debug.WriteLine($"Data distribution: Booked={bookedCount}, InProgress={inProgressCount}, Invoiced={invoicedCount}");

                // 檢查是否處於 In Progress 模式
                bool isInProgressMode = IsInProgressStatus;

                // 記錄 In Progress 項目的總預期佣金
                decimal totalInProgressCommission = data
                    .Where(x => !x.CompletionDate.HasValue && x.TotalCommission == 0)
                    .Sum(x => x.VertivValue * 0.12m);

                Debug.WriteLine($"Total expected commission for In Progress items: ${totalInProgressCommission:N2}");

                var reps = data
                    .GroupBy(x => x.SalesRep)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g =>
                    {
                        // 針對 In Progress 模式，計算預期佣金
                        decimal expectedCommission = 0;
                        if (isInProgressMode)
                        {
                            expectedCommission = g.Sum(x => x.VertivValue * 0.12m);
                        }

                        return new SalesLeaderboardItem
                        {
                            SalesRep = g.Key,
                            // 在 In Progress 模式下，將所有預期佣金都放在 Agency Margin
                            AgencyMargin = Math.Round(isInProgressMode ?
                                expectedCommission : // In Progress 模式 - 使用預期佣金
                                g.Sum(x => x.AgencyMargin), 2), // 其他模式 - 使用實際 Agency Margin
                                                                // Buy Resell Margin 在 In Progress 模式下為 0
                            BuyResellMargin = Math.Round(isInProgressMode ?
                                0 : // In Progress 模式下為 0
                                g.Sum(x => x.BuyResellValue), 2), // 其他模式 - 使用實際 Buy Resell Margin
                                                                  // Total Margin 等於 Agency + Buy Resell
                            TotalMargin = Math.Round(isInProgressMode ?
                                expectedCommission : // In Progress 模式 - 使用預期佣金
                                g.Sum(x => x.TotalCommission), 2), // 其他模式 - 使用實際 Total Commission
                                                                   // 記錄 Vertiv 值
                            VertivValue = Math.Round(g.Sum(x => x.VertivValue), 2)
                        };
                    })
                    .OrderByDescending(x => x.TotalMargin)
                    .ToList();

                // 設置排名
                for (int i = 0; i < reps.Count; i++)
                {
                    reps[i].Rank = i + 1;
                }

                // 輸出每個銷售代表的佣金明細，用於調試
                foreach (var rep in reps.Take(Math.Min(5, reps.Count)))
                {
                    Debug.WriteLine($"Rep: {rep.SalesRep}, " +
                                   $"Agency: ${rep.AgencyMargin}, " +
                                   $"BuyResell: ${rep.BuyResellMargin}, " +
                                   $"Total: ${rep.TotalMargin}");
                }

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>(reps);
                });

                Debug.WriteLine($"Sales rep data loaded, {reps.Count} items");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading sales rep data: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
            }
        }


        // 標準化產品類型名稱
        private string NormalizeProductType(string productType)
        {
            if (string.IsNullOrEmpty(productType))
                return "Other";

            // 轉為小寫以便比較
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

            // 如果沒有匹配，返回原始名稱
            return productType;
        }
        // 更新圖表軸
        private void UpdateChartAxes()
        {
            try
            {
                // 找出圖表中的最大值，用於設置 Y 軸的最大值
                double maxTarget = 0;
                double maxAchievement = 0;

                if (TargetVsAchievementData.Any())
                {
                    maxTarget = (double)TargetVsAchievementData.Max(d => d.Target);
                    maxAchievement = (double)TargetVsAchievementData.Max(d => d.Achievement);
                }

                double maxValue = Math.Max(maxTarget, maxAchievement);
                // 增加 20% 空間，使圖表不會太緊湊
                maxValue = maxValue * 1.2;
                // 向上取整到下一個整數
                maxValue = Math.Ceiling(maxValue);

                // 確保最小有 5 的空間
                YAxisMaximum = Math.Max(5, maxValue);

                Debug.WriteLine($"圖表 Y 軸最大值設為: {YAxisMaximum}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新圖表軸時發生錯誤: {ex.Message}");
                YAxisMaximum = 10; // 默認值
            }
        }

        #region 輔助方法

        // 從月份獲取季度
        private int GetQuarterFromMonth(int month)
        {
            // 財年 Q1: 8-10月, Q2: 11-1月, Q3: 2-4月, Q4: 5-7月
            return month switch
            {
                8 or 9 or 10 => 1,
                11 or 12 or 1 => 2,
                2 or 3 or 4 => 3,
                5 or 6 or 7 => 4,
                _ => 1
            };
        }

        // 處理數據更新通知
        private void OnDataUpdated(object sender, DateTime lastUpdated)
        {
            Debug.WriteLine($"收到 Excel 數據更新通知: {lastUpdated}");
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await ReloadDataAsync();
            });
        }

        // 處理目標更新通知
        private void OnTargetsUpdated(object sender, EventArgs e)
        {
            Debug.WriteLine("收到目標更新通知");
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await ReloadDataAsync();
            });
        }

        #endregion

        #region 樣本數據

        // 載入樣本圖表數據
        private void LoadSampleChartData()
        {
            // Instead of loading sample data, clear the charts
            TargetVsAchievementData = new ObservableCollection<ChartData>();
            AchievementTrendData = new ObservableCollection<ChartData>();

            Debug.WriteLine("No data available for charts, displaying empty charts instead of sample data");
        }

        #endregion // 结束 #region 樣本數据

        #endregion // 结束 #region 辅助方法

    }
}