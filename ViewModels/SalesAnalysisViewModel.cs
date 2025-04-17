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

        // 載入狀態
        [ObservableProperty]
        private bool _isLoading;

        // 視圖類型控制
        [ObservableProperty]
        private bool _isSummaryView = true;

        [ObservableProperty]
        private string _viewType = "ByProduct"; // ByProduct, ByRep - removed ByDeptLOB

        [ObservableProperty]
        private bool _isBookedStatus = true; // Default to Booked status

        // Add new properties instead of computed ones
        [ObservableProperty]
        private bool _isInProgressStatus = false;

        [ObservableProperty]
        private bool _isCompletedStatus = false;

        [ObservableProperty]
        private string _status = "Booked"; // Default status

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
            Debug.WriteLine($"尝试切换到状态: {status}");
            bool changed = false;

            switch (status.ToLower())
            {
                case "booked":
                    if (!IsBookedStatus)
                    {
                        IsBookedStatus = true;
                        IsInProgressStatus = false;
                        IsCompletedStatus = false;
                        changed = true;
                    }
                    break;

                case "inprogress":
                    if (!IsInProgressStatus)
                    {
                        IsBookedStatus = false;
                        IsInProgressStatus = true;
                        IsCompletedStatus = false;
                        changed = true;
                    }
                    break;

                case "completed":
                    if (!IsCompletedStatus)
                    {
                        IsBookedStatus = false;
                        IsInProgressStatus = false;
                        IsCompletedStatus = true;
                        changed = true;
                    }
                    break;
            }

            if (changed)
            {
                Debug.WriteLine($"状态已更改为: {status}，重新加载数据");
                await FilterData(); // 直接调用方法，而不是通过Command
            }
            else
            {
                Debug.WriteLine($"状态未更改，仍然是: {status}");
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
                    Debug.WriteLine("没有可用数据进行过滤");
                    _filteredData = new List<SalesData>();
                    return;
                }

                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1); // 包含整个结束日期

                Debug.WriteLine($"过滤日期范围: {startDate:yyyy-MM-dd} 到 {endDate:yyyy-MM-dd}");

                // 根据所选标签应用不同的过滤条件
                if (IsBookedStatus)
                {
                    // Booked: A列(ReceivedDate)在日期范围内，Y列(CompletionDate)为空，N列(TotalCommission)有值
                    _filteredData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= startDate &&
                               x.ReceivedDate.Date <= endDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission != 0)
                        .ToList();

                    Debug.WriteLine($"[Booked] 过滤后记录数: {_filteredData.Count}");
                }
                else if (IsInProgressStatus)
                {
                    // In Progress: A列(ReceivedDate)在日期范围内，Y列(CompletionDate)为空，N列(TotalCommission)为零或空
                    _filteredData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= startDate &&
                               x.ReceivedDate.Date <= endDate.Date &&
                               !x.CompletionDate.HasValue &&
                               x.TotalCommission == 0)
                        .ToList();

                    Debug.WriteLine($"[In Progress] 过滤后记录数: {_filteredData.Count}");

                    // 打印前几条记录的详细信息用于调试
                    foreach (var item in _filteredData.Take(Math.Min(5, _filteredData.Count)))
                    {
                        Debug.WriteLine($"In Progress 示例: " +
                            $"接收日期={item.ReceivedDate:yyyy-MM-dd}, " +
                            $"POValue=${item.POValue:N2}, " +
                            $"预期佣金(POValue*0.12)=${item.POValue * 0.12m:N2}");
                    }
                }
                else if (IsCompletedStatus)
                {
                    // Completed: Y列(CompletionDate)在日期范围内且不为空
                    _filteredData = _allSalesData
                        .Where(x => x.CompletionDate.HasValue &&
                               x.CompletionDate.Value.Date >= startDate &&
                               x.CompletionDate.Value.Date <= endDate.Date)
                        .ToList();

                    Debug.WriteLine($"[Completed] 过滤后记录数: {_filteredData.Count}");
                }
                else
                {
                    // 默认：显示空数据
                    _filteredData = new List<SalesData>();
                    Debug.WriteLine("没有选择任何状态，返回空数据");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"过滤数据时出错: {ex.Message}");
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

                // 計算"Remaining to target"
                decimal remainingToTarget = targetValue - totalByReceivedDate;

                // 計算沒有完成日期的訂單的N欄總和（以A欄日期為基礎）
                decimal bookedMarginNotInvoiced = receivedDateData
                    .Where(x => !x.CompletionDate.HasValue) // 沒有完成日期的記錄
                    .Sum(x => x.TotalCommission); // N欄 - Total Commission

                // 更新摘要
                decimal actualRemaining = targetValue - completedAchievement;

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

                // 根據狀態選擇合適的日期欄位進行分組
                var monthlyData = new List<dynamic>();

                if (IsBookedStatus || IsInProgressStatus) // 如果是 Booked 或 In Progress 狀態
                {
                    // 用接收日期 (ReceivedDate) 進行分組
                    monthlyData = _filteredData
                        .GroupBy(x => new
                        {
                            Year = x.ReceivedDate.Year,
                            Month = x.ReceivedDate.Month
                        })
                        .Select(g => new
                        {
                            YearMonth = g.Key,
                            POValue = g.Sum(x => x.POValue),
                            MarginValue = g.Sum(x => x.TotalCommission),
                            CommissionValue = g.Sum(x => x.TotalCommission)
                        })
                        .OrderBy(x => x.YearMonth.Year)
                        .ThenBy(x => x.YearMonth.Month)
                        .ToList<dynamic>();
                }
                else // 如果是 Completed 狀態
                {
                    // 用完成日期 (CompletionDate) 進行分組
                    monthlyData = _filteredData
                        .Where(x => x.CompletionDate.HasValue)
                        .GroupBy(x => new
                        {
                            Year = x.CompletionDate.Value.Year,
                            Month = x.CompletionDate.Value.Month
                        })
                        .Select(g => new
                        {
                            YearMonth = g.Key,
                            POValue = g.Sum(x => x.POValue),
                            MarginValue = g.Sum(x => x.TotalCommission),
                            CommissionValue = g.Sum(x => x.TotalCommission)
                        })
                        .OrderBy(x => x.YearMonth.Year)
                        .ThenBy(x => x.YearMonth.Month)
                        .ToList<dynamic>();
                }

                // 檢查處理後的資料
                Debug.WriteLine($"處理後的月度資料筆數: {monthlyData.Count}");

                // 準備圖表資料
                var newTargetVsAchievementData = new ObservableCollection<ChartData>();
                var newAchievementTrendData = new ObservableCollection<ChartData>();

                foreach (var month in monthlyData)
                {
                    var label = $"{month.YearMonth.Year}/{month.YearMonth.Month:D2}";
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

                // 使用之前已经过滤好的数据
                if (_filteredData == null || !_filteredData.Any())
                {
                    Debug.WriteLine("没有过滤后的数据可用于加载排行榜");

                    // 根据视图类型清空数据
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

                Debug.WriteLine($"用于计算的记录数: {_filteredData.Count} 条");

                // 根据视图类型加载相应数据
                switch (ViewType)
                {
                    case "ByProduct":
                        await LoadProductDataAsync(_filteredData);
                        break;
                    case "ByRep":
                        await LoadSalesRepDataAsync(_filteredData);
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加载排行榜数据时出错: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 出错时清空数据
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


        // 載入產品數據
        private async Task LoadProductDataAsync(List<SalesData> data)
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
                        // 如果是 In Progress 模式，則使用 POValue * 0.12 作為 Total Margin
                        // 進一步拆分為 75% Agency 和 25% Buy Resell
                        AgencyMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.12m) : // 75% of expected commission (0.12)
                            g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0m) : // 25% of expected commission (0.12)
                            g.Sum(x => x.BuyResellValue), 2),
                        // Total Margin 直接計算，不使用 TotalCommission
                        TotalMargin = Math.Round(isInProgressMode ?
                            g.Sum(x => x.POValue * 0.12m) : // Expected margin for in progress items
                            g.Sum(x => x.TotalCommission), 2),
                        // 修改为使用 VertivValue 代替 POValue
                        POValue = Math.Round(g.Sum(x => x.VertivValue), 2),
                        // 標記是否為 In Progress 模式（用於 UI 顯示）
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
                            ? Math.Round((product.POValue / totalPOValue), 2) // 將百分比存為小數
                            : 0;

                        Debug.WriteLine($"Product: {product.ProductType}, " +
                                       $"Agency: ${product.AgencyMargin}, " +
                                       $"BuyResell: ${product.BuyResellMargin}, " +
                                       $"Total: ${product.TotalMargin}, " +
                                       $"Vertiv Value: ${product.POValue}, " +
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


        // 載入銷售代表數據
        private async Task LoadSalesRepDataAsync(List<SalesData> data)
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
                            g.Sum(x => x.POValue * 0m) : // 25% of expected commission (0.12)
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
                    SalesLeaderboard = new ObservableCollection<SalesLeaderboardItem>(reps);
                });

                Debug.WriteLine($"Sales rep data loaded, {reps.Count} items");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading sales rep data: {ex.Message}");
            }
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