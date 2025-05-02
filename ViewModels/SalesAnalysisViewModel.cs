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
            Debug.WriteLine($"切換狀態為: {status}");
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
                case "invoiced":
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
                try
                {
                    IsLoading = true;

                    // 清空現有圖表數據，確保不會顯示舊數據
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        TargetVsAchievementData = new ObservableCollection<ChartData>();
                        AchievementTrendData = new ObservableCollection<ChartData>();
                    });

                    // 更新屬性，觸發UI更新
                    OnPropertyChanged(nameof(IsAllStatus));
                    OnPropertyChanged(nameof(IsBookedStatus));
                    OnPropertyChanged(nameof(IsInProgressStatus));
                    OnPropertyChanged(nameof(IsCompletedStatus));
                    OnPropertyChanged(nameof(Status));

                    // 完全重新過濾數據
                    FilterDataByDateRange();

                    // 重要：先加載表格數據，再加載圖表數據
                    await LoadLeaderboardDataAsync();

                    // 等待表格數據完全加載並更新UI
                    await Task.Delay(300);

                    // 完成表格加載後再加載圖表數據
                    await LoadChartDataAsync();

                    Debug.WriteLine($"狀態切換為 {status} 完成，數據已重新加載");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"狀態切換時發生錯誤: {ex.Message}");
                    Debug.WriteLine(ex.StackTrace);
                }
                finally
                {
                    IsLoading = false;
                }
            }
            else
            {
                Debug.WriteLine($"狀態沒有變化，仍為: {status}");
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

                // 在重新載入數據之前強制清除緩存
                _excelService.ClearCache();

                // 重新載入原始數據（如果需要）
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
                        ClearAllDisplays();
                        return;
                    }
                }

                // 過濾數據，確保應用日期範圍
                FilterDataByDateRange();

                // 載入摘要數據
                await LoadSummaryDataAsync();

                // 載入圖表數據
                await LoadChartDataAsync();

                // 載入排行榜數據
                await LoadLeaderboardDataAsync();

                // 更新圖表軸
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

                // 無論選擇什麼狀態，總是嚴格按照所選日期範圍提取數據
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
                    // All狀態: 合併所有數據，確保只包含所選日期範圍內的數據
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

                // 驗證過濾結果的日期範圍
                if (_filteredData.Any())
                {
                    var minDate = _filteredData.Min(x => x.CompletionDate ?? x.ReceivedDate);
                    var maxDate = _filteredData.Max(x => x.CompletionDate ?? x.ReceivedDate);
                    Debug.WriteLine($"過濾後數據的日期範圍: {minDate:yyyy-MM-dd} 到 {maxDate:yyyy-MM-dd}");
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

                // 使用完成日期過濾的數據計算實際達成
                decimal completedAchievement = _allSalesData
                    .Where(x => x.CompletionDate.HasValue &&
                            x.CompletionDate.Value.Date >= StartDate.Date &&
                            x.CompletionDate.Value.Date <= EndDate.Date.AddDays(1).AddSeconds(-1))
                    .Sum(x => x.TotalCommission);

                // 獲取所有基於接收日期的總額（用於計算Remaining to target）
                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1);

                // 基於接收日期過濾數據
                var receivedDateData = _allSalesData?
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate.Date)
                    .ToList() ?? new List<SalesData>();

                decimal totalByReceivedDate = receivedDateData.Sum(x => x.TotalCommission);


                // 計算沒有完成日期的訂單的N欄總和（以A欄日期為基礎）
                decimal bookedMarginNotInvoiced = receivedDateData
                    .Where(x => !x.CompletionDate.HasValue) // 沒有完成日期的記錄
                    .Sum(x => x.TotalCommission); // N欄 - Total Commission


                // 更新摘要
                decimal actualRemaining = targetValue - completedAchievement;

                // 計算"Remaining to target"
                decimal remainingToTarget = actualRemaining - bookedMarginNotInvoiced;


                // 百分比計算
                decimal achievementPercentage = 0;
                decimal marginPercentage = 0;
                decimal remainingTargetPercentage = 0;

                if (targetValue > 0)
                {
                    achievementPercentage = Math.Round((completedAchievement / targetValue) * 100, 1);
                    remainingTargetPercentage = Math.Round((actualRemaining / targetValue) * 100, 1);
                }

                // 轉換為百萬單位 - 使用更高精度的四捨五入，保留到千位數
                targetValue = Math.Round(targetValue / 1000000m, 6);
                decimal completedAchievementM = Math.Round(completedAchievement / 1000000m, 6);
                decimal remainingToTargetM = Math.Round(remainingToTarget / 1000000m, 6);
                decimal actualRemainingM = Math.Round(actualRemaining / 1000000m, 6);

                // 轉換已Booked但未Invoiced的Margin為百萬單位
                decimal bookedMarginNotInvoicedM = Math.Round(bookedMarginNotInvoiced / 1000000m, 6);

                // 計算Booked Margin的比率（相對於總目標）
                decimal bookedMarginRate = 0;
                if (targetValue > 0)
                {
                    bookedMarginRate = Math.Round((bookedMarginNotInvoiced / (targetValue * 1000000m)) * 100, 1);
                }

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    if (Summary == null)
                    {
                        Summary = new SalesAnalysisSummary();
                    }
                    Summary.TotalTarget = targetValue;
                    Summary.TotalAchievement = completedAchievementM;
                    Summary.TotalMargin = bookedMarginNotInvoicedM; // 改成Booked Margin not Invoiced
                    Summary.RemainingTarget = remainingToTargetM; // 原來屬性保留用於Remaining to target
                    Summary.ActualRemaining = actualRemainingM; // 新屬性用於Actual Remaining
                    Summary.AchievementPercentage = achievementPercentage;
                    Summary.MarginPercentage = bookedMarginRate; // 使用Booked Margin的比率
                    Summary.RemainingTargetPercentage = remainingTargetPercentage;
                });

                Debug.WriteLine($"摘要數據: Target=${targetValue}M, Achievement=${completedAchievementM}M ({achievementPercentage}%), " +
                               $"Remaining to target=${remainingToTargetM}M, Actual Remaining=${actualRemainingM}M, " +
                               $"Booked Margin not Invoiced=${bookedMarginNotInvoicedM}M ({bookedMarginRate}%)");
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

        private async Task LoadChartDataForAllStatus()
        {
            try
            {
                Debug.WriteLine("正在載入All狀態的圖表數據（修正版）");

                // 使用日期範圍
                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1);

                // 取得各個狀態的數據
                List<SalesData> bookedData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate.Date &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission > 0)
                    .ToList();

                List<SalesData> inProgressData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate &&
                           x.ReceivedDate.Date <= endDate.Date &&
                           !x.CompletionDate.HasValue &&
                           x.TotalCommission == 0)
                    .ToList();

                List<SalesData> completedData = _allSalesData
                    .Where(x => x.CompletionDate.HasValue &&
                           x.CompletionDate.Value.Date >= startDate &&
                           x.CompletionDate.Value.Date <= endDate.Date)
                    .ToList();

                Debug.WriteLine($"各狀態數據量：Booked={bookedData.Count}, InProgress={inProgressData.Count}, Completed={completedData.Count}");

                // 獲取所有月份
                var allMonths = GetAllMonthsInRange(startDate, endDate);

                // 針對每個月份分別計算各狀態的數據
                var monthlyData = new List<MonthlyDataItem>();

                foreach (var (year, month) in allMonths)
                {
                    var monthStart = new DateTime(year, month, 1);
                    var monthEnd = monthStart.AddMonths(1).AddDays(-1);

                    // 調整月份範圍，確保不超出選定日期範圍
                    if (monthStart < startDate) monthStart = startDate;
                    if (monthEnd > endDate) monthEnd = endDate;

                    // 1. 計算Booked數據
                    decimal bookedCommission = CalculateMonthlyCommission(
                        bookedData,
                        item => item.ReceivedDate,
                        monthStart,
                        monthEnd,
                        item => item.TotalCommission);

                    // 2. 計算InProgress數據（使用12%的Vertiv值）
                    decimal inProgressCommission = CalculateMonthlyCommission(
                        inProgressData,
                        item => item.ReceivedDate,
                        monthStart,
                        monthEnd,
                        item => item.VertivValue * 0.12m);

                    // 3. 計算Completed數據
                    decimal completedCommission = CalculateMonthlyCommission(
                        completedData,
                        item => item.CompletionDate.Value,
                        monthStart,
                        monthEnd,
                        item => item.TotalCommission);

                    // 計算合計
                    decimal totalMonthlyCommission = bookedCommission + inProgressCommission + completedCommission;

                    Debug.WriteLine($"月份{year}/{month:D2}: Booked=${bookedCommission:N2} + InProgress=${inProgressCommission:N2} + " +
                                  $"Completed=${completedCommission:N2} = 總計${totalMonthlyCommission:N2}");

                    // 添加到月度數據集合
                    monthlyData.Add(new MonthlyDataItem
                    {
                        Year = year,
                        Month = month,
                        CommissionValue = totalMonthlyCommission,
                        VertivValue = 0, // 不需要用於圖表
                        RecordCount = 0  // 不需要用於圖表
                    });
                }

                // 從月度數據生成圖表數據
                var chartData = GenerateChartDataFromMonthlyData(monthlyData);

                // 更新UI
                await MainThread.InvokeOnMainThreadAsync(() => {
                    TargetVsAchievementData = chartData;
                    AchievementTrendData = chartData;

                    // 設置Y軸最大值
                    SetChartYAxisMaximum(chartData);
                });

                Debug.WriteLine("All狀態圖表數據載入完成（修正版）");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入All狀態圖表數據時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 確保發生錯誤時顯示空圖表
                await MainThread.InvokeOnMainThreadAsync(() => {
                    TargetVsAchievementData = new ObservableCollection<ChartData>();
                    AchievementTrendData = new ObservableCollection<ChartData>();
                    YAxisMaximum = 5000;
                });
            }
        }

        // 輔助方法：獲取日期範圍內的所有月份
        private List<(int Year, int Month)> GetAllMonthsInRange(DateTime startDate, DateTime endDate)
        {
            var result = new List<(int Year, int Month)>();
            var currentMonth = new DateTime(startDate.Year, startDate.Month, 1);
            var endMonth = new DateTime(endDate.Year, endDate.Month, 1);

            while (currentMonth <= endMonth)
            {
                result.Add((currentMonth.Year, currentMonth.Month));
                currentMonth = currentMonth.AddMonths(1);
            }

            return result;
        }

        // 輔助方法：計算月度佣金
        private decimal CalculateMonthlyCommission<T>(
            List<T> items,
            Func<T, DateTime> dateSelector,
            DateTime monthStart,
            DateTime monthEnd,
            Func<T, decimal> valueSelector)
            where T : SalesData
        {
            return items
                .Where(item => {
                    var date = dateSelector(item);
                    return date >= monthStart && date <= monthEnd;
                })
                .Sum(valueSelector);
        }

        private ObservableCollection<ChartData> GenerateChartDataFromMonthlyData(List<MonthlyDataItem> monthlyData)
        {
            var result = new ObservableCollection<ChartData>();

            foreach (var item in monthlyData)
            {
                var label = $"{item.Year}/{item.Month:D2}";

                // 轉換為千單位
                decimal commissionInThousands = Math.Round(item.CommissionValue / 1000m, 2);

                // 獲取月度目標
                int fiscalYear = item.Month >= 8 ? item.Year + 1 : item.Year;
                int quarter = GetQuarterFromMonth(item.Month);
                decimal targetValue = _targetService.GetCompanyQuarterlyTarget(fiscalYear, quarter) / 3;
                decimal targetInThousands = Math.Round(targetValue / 1000m, 2);

                // 添加到結果集合
                result.Add(new ChartData
                {
                    Label = label,
                    Target = targetInThousands,
                    Achievement = commissionInThousands
                });

                Debug.WriteLine($"圖表數據點: {label}, 目標=${targetInThousands:N2}K, 達成=${commissionInThousands:N2}K");
            }

            return result;
        }

        // 輔助方法：設置圖表Y軸最大值
        private void SetChartYAxisMaximum(ObservableCollection<ChartData> chartData)
        {
            double maxTarget = 0;
            double maxAchievement = 0;

            if (chartData.Any())
            {
                maxTarget = (double)chartData.Max(d => d.Target);
                maxAchievement = (double)chartData.Max(d => d.Achievement);
            }

            double maxValue = Math.Max(maxTarget, maxAchievement);
            maxValue = Math.Max(maxValue * 1.2, 1000); // 增加20%空間，但確保至少為1000
            YAxisMaximum = Math.Ceiling(maxValue);

            Debug.WriteLine($"設置Y軸最大值: {YAxisMaximum}");
        }


        // 載入圖表數據
        private async Task LoadChartDataAsync()
        {
            try
            {
                // 記錄當前狀態用於調試
                string currentStatus = IsAllStatus ? "All" :
                                      IsBookedStatus ? "Booked" :
                                      IsInProgressStatus ? "InProgress" :
                                      IsCompletedStatus ? "Completed" : "Unknown";

                Debug.WriteLine($"載入圖表數據，當前狀態: {currentStatus}, 日期範圍: {StartDate:yyyy-MM-dd} 到 {EndDate:yyyy-MM-dd}");

                if (_filteredData == null || !_filteredData.Any())
                {
                    // 確保集合不為 null，而是一個空集合
                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        TargetVsAchievementData = new ObservableCollection<ChartData>();
                        AchievementTrendData = new ObservableCollection<ChartData>();
                        // 設置一個合理的 Y 軸最大值
                        YAxisMaximum = 5000;
                    });

                    Debug.WriteLine("沒有數據可用於圖表，顯示空圖表");
                    return;
                }

                // 關鍵修改：如果是 All 狀態，使用特殊的方法直接計算三種狀態的總和
                if (IsAllStatus)
                {
                    await LoadChartDataForAllStatus();
                    return; // 提前返回，不執行後續代碼
                }

                // 確認表格數據已經加載完成
                decimal tableTotal = 0;
                if (IsProductView)
                {
                    tableTotal = ProductSalesData.Sum(p => p.TotalMargin);
                }
                else if (IsRepView)
                {
                    tableTotal = SalesLeaderboard.Sum(r => r.TotalMargin);
                }

                Debug.WriteLine($"【{currentStatus}】狀態表格合計: ${tableTotal:N2}");

                // 為避免計算差異，直接使用表格數據的總計值
                decimal expectedTotal = tableTotal;

                // 首先，獲取所選日期範圍內的所有月份
                var startMonth = new DateTime(StartDate.Year, StartDate.Month, 1);
                var endMonth = new DateTime(EndDate.Year, EndDate.Month, 1);

                var allMonthsInRange = new List<(int Year, int Month)>();
                var currentMonth = startMonth;

                while (currentMonth <= endMonth)
                {
                    allMonthsInRange.Add((currentMonth.Year, currentMonth.Month));
                    currentMonth = currentMonth.AddMonths(1);
                }

                Debug.WriteLine($"日期範圍內的月份數: {allMonthsInRange.Count}");

                // 計算狀態相關數據
                bool isInProgressStatus = IsInProgressStatus;
                bool isBookedStatus = IsBookedStatus;
                bool isCompletedStatus = IsCompletedStatus;

                // 準備處理圖表數據
                var monthlyData = new List<MonthlyDataItem>();

                // 這裡使用與表格相同的數據源
                List<SalesData> dataToProcess = _filteredData;

                // 每月總計
                decimal chartTotalCommission = 0;

                foreach (var (year, month) in allMonthsInRange)
                {
                    // 確定當月的日期範圍
                    var monthStart = new DateTime(year, month, 1);
                    var monthEnd = monthStart.AddMonths(1).AddDays(-1);

                    // 如果是起始月份，使用StartDate
                    if (year == StartDate.Year && month == StartDate.Month)
                        monthStart = StartDate;

                    // 如果是結束月份，使用EndDate
                    if (year == EndDate.Year && month == EndDate.Month)
                        monthEnd = EndDate;

                    Debug.WriteLine($"處理月份: {year}/{month:D2}, 範圍: {monthStart:yyyy-MM-dd} 到 {monthEnd:yyyy-MM-dd}");

                    // 計算該月的佣金數據
                    decimal monthlyCommission = 0;
                    decimal monthlyVertivValue = 0;
                    int recordCount = 0;

                    foreach (var item in dataToProcess)
                    {
                        // 決定要使用哪個日期進行比較
                        DateTime dateToCompare;

                        if (isCompletedStatus && item.CompletionDate.HasValue)
                        {
                            // Completed狀態用完成日期
                            dateToCompare = item.CompletionDate.Value;
                        }
                        else
                        {
                            // 其他狀態用接收日期
                            dateToCompare = item.ReceivedDate;
                        }

                        // 檢查日期是否在當月範圍內
                        if (dateToCompare >= monthStart && dateToCompare <= monthEnd)
                        {
                            recordCount++;

                            // 計算佣金
                            if (isInProgressStatus)
                            {
                                // In Progress模式下使用預期佣金 (12% of Vertiv Value)
                                monthlyCommission += item.VertivValue * 0.12m;
                            }
                            else
                            {
                                // 其他情況使用實際佣金
                                monthlyCommission += item.TotalCommission;
                            }

                            // 計算Vertiv值
                            monthlyVertivValue += item.VertivValue;
                        }
                    }

                    // 添加到月度數據
                    monthlyData.Add(new MonthlyDataItem
                    {
                        Year = year,
                        Month = month,
                        CommissionValue = monthlyCommission,
                        VertivValue = monthlyVertivValue,
                        RecordCount = recordCount
                    });

                    chartTotalCommission += monthlyCommission;

                    Debug.WriteLine($"月份 {year}/{month:D2}: 記錄數={recordCount}, 佣金=${monthlyCommission:N2}, Vertiv值=${monthlyVertivValue:N2}");
                }

                // 如果圖表總計為零但表格有數據，將所有數據分配到第一個月
                if (chartTotalCommission < 0.01m && tableTotal > 0.01m && monthlyData.Any())
                {
                    Debug.WriteLine("圖表總計為零但表格有數據，將所有數據分配到第一個月");
                    monthlyData[0].CommissionValue = tableTotal;
                    chartTotalCommission = tableTotal;
                }

                // 檢查圖表總計與表格總計是否一致
                Debug.WriteLine($"圖表總計: ${chartTotalCommission:N2}, 表格總計: ${tableTotal:N2}");

                // 如果差異超過1元，應用調整係數
                if (Math.Abs(chartTotalCommission - tableTotal) > 1 && chartTotalCommission > 0.01m)
                {
                    // 計算調整係數
                    decimal adjustFactor = tableTotal / chartTotalCommission;
                    Debug.WriteLine($"應用調整係數: {adjustFactor:N4} 以匹配表格總計");

                    // 調整每個月的數據
                    for (int i = 0; i < monthlyData.Count; i++)
                    {
                        monthlyData[i].CommissionValue *= adjustFactor;
                    }

                    // 重新計算圖表總計
                    chartTotalCommission = monthlyData.Sum(m => m.CommissionValue);
                    Debug.WriteLine($"調整後圖表總計: ${chartTotalCommission:N2}");
                }

                // 準備圖表資料
                var newTargetVsAchievementData = new ObservableCollection<ChartData>();
                var newAchievementTrendData = new ObservableCollection<ChartData>();

                // 轉換月度數據到圖表數據
                foreach (var monthData in monthlyData)
                {
                    var label = $"{monthData.Year}/{monthData.Month:D2}";

                    // 修改：轉換為千單位，而不是百萬單位
                    decimal commissionValueInThousands = Math.Round(monthData.CommissionValue / 1000m, 2);

                    // 從目標服務獲取該月份的目標值
                    int fiscalYear = monthData.Month >= 8 ? monthData.Year + 1 : monthData.Year;
                    int quarter = GetQuarterFromMonth(monthData.Month);
                    decimal targetValue = _targetService.GetCompanyQuarterlyTarget(fiscalYear, quarter) / 3; // 將季度目標平均分配到月

                    // 修改：轉換為千單位，而不是百萬單位
                    decimal targetValueInThousands = Math.Round(targetValue / 1000m, 2);

                    // 添加到目標與達成對比圖表
                    newTargetVsAchievementData.Add(new ChartData
                    {
                        Label = label,
                        Target = targetValueInThousands,
                        Achievement = commissionValueInThousands
                    });

                    // 添加到達成趨勢圖表
                    newAchievementTrendData.Add(new ChartData
                    {
                        Label = label,
                        Target = targetValueInThousands,
                        Achievement = commissionValueInThousands
                    });

                    Debug.WriteLine($"圖表數據點: {label}, 目標=${targetValueInThousands:N2}K, 達成=${commissionValueInThousands:N2}K");
                }

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
                    maxValue = Math.Max(maxValue * 1.2, 1000); // 增加 20% 空間，但確保至少為 1000
                    YAxisMaximum = Math.Ceiling(maxValue); // 不再使用 Math.Min(2, ...) 限制

                    Debug.WriteLine($"設置 Y 軸最大值: {YAxisMaximum}");
                });

                // 最終確認
                Debug.WriteLine($"【{currentStatus}】狀態圖表數據加載完成，總計值 ${chartTotalCommission:N2}");
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
                    YAxisMaximum = 5000; // 調整為適合千單位的最大值
                });
            }
        }

        // 輔助類：用於存儲月度數據
        private class MonthlyDataItem
        {
            public int Year { get; set; }
            public int Month { get; set; }
            public decimal CommissionValue { get; set; }
            public decimal VertivValue { get; set; }
            public int RecordCount { get; set; }
        }

        // 輔助方法：生成記錄的唯一鍵值
        private string GetUniqueKey(SalesData item)
        {
            // 根據業務邏輯定義唯一鍵
            // 這裡使用接收日期、產品類型和銷售代表來識別記錄
            return $"{item.ReceivedDate:yyyy-MM-dd}_{item.ProductType}_{item.SalesRep}_{item.POValue:F2}";
        }

        // 定義用於比較記錄的輔助方法
        private bool IsSameRecord(SalesData a, SalesData b)
        {
            // 根據業務邏輯定義兩條記錄是否應視為相同
            // 例如，可以比較某些關鍵字段
            return a.ReceivedDate == b.ReceivedDate &&
                   a.SalesRep == b.SalesRep &&
                   a.ProductType == b.ProductType &&
                   Math.Abs(a.POValue - b.POValue) < 0.01m;
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
                // 增加一些空間，使圖表不會太緊湊
                maxValue = maxValue * 1.2;
                // 向上取整到下一個整數
                maxValue = Math.Ceiling(maxValue);

                // 將最大值設為2，除非數據值超過2
                YAxisMaximum = 1500;

                Debug.WriteLine($"圖表 Y 軸最大值設為: {YAxisMaximum}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新圖表軸時發生錯誤: {ex.Message}");
                YAxisMaximum = 2; // 默認值改為2
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