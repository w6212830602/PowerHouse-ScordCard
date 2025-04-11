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
        private bool _isBookedStatus = true; // true=Booked, false=Completed

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
            bool wasBooked = IsBookedStatus;
            IsBookedStatus = status.ToLower() == "booked";

            if (wasBooked != IsBookedStatus)
            {
                await LoadLeaderboardDataAsync();
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
        public async Task ReloadDataAsync()
        {
            try
            {
                IsLoading = true;

                // 重新載入原始數據（如果需要）
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    var (data, lastUpdated) = await _excelService.LoadDataAsync();
                    _allSalesData = data ?? new List<SalesData>();
                }

                // 過濾數據
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
                Debug.WriteLine($"重新載入數據時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 確保有一些默認數據顯示
                LoadSampleData();
            }
            finally
            {
                IsLoading = false;
            }
        }

        // 過濾數據按日期範圍
        private void FilterDataByDateRange()
        {
            try
            {
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    _filteredData = new List<SalesData>();
                    return;
                }

                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1); // 包含結束日期整天

                Debug.WriteLine($"過濾日期範圍: {startDate:yyyy-MM-dd} 到 {endDate:yyyy-MM-dd}");

                // 修改：使用完成日期(Y列)進行過濾，而不是接收日期(A列)
                _filteredData = _allSalesData
                    .Where(x => x.CompletionDate.HasValue &&
                           x.CompletionDate.Value.Date >= startDate &&
                           x.CompletionDate.Value.Date <= endDate.Date)
                    .ToList();

                Debug.WriteLine($"日期過濾後數據: {_filteredData.Count} 條記錄，" +
                              $"全部為已完成記錄");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"過濾數據時發生錯誤: {ex.Message}");
                _filteredData = new List<SalesData>();
            }
        }

        // 載入摘要數據
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
                decimal completedAchievement = _filteredData?.Sum(x => x.TotalCommission) ?? 0;

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

                // 這裡totalAchievement和totalMargin現在基本相同
                if (completedAchievement > 0)
                {
                    marginPercentage = Math.Round((completedAchievement / completedAchievement) * 100, 1);
                }

                // 轉換為百萬單位
                targetValue = Math.Round(targetValue / 1000000m, 2);
                decimal completedAchievementM = Math.Round(completedAchievement / 1000000m, 2);
                decimal remainingToTargetM = Math.Round(remainingToTarget / 1000000m, 2);
                decimal actualRemainingM = Math.Round(actualRemaining / 1000000m, 2);

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    if (Summary == null)
                    {
                        Summary = new SalesAnalysisSummary();
                    }
                    Summary.TotalTarget = targetValue;
                    Summary.TotalAchievement = completedAchievementM;
                    Summary.TotalMargin = completedAchievementM; // 同上
                    Summary.RemainingTarget = remainingToTargetM; // 原來屬性保留用於Remaining to target
                    Summary.ActualRemaining = actualRemainingM; // 新屬性用於Actual Remaining
                    Summary.AchievementPercentage = achievementPercentage;
                    Summary.MarginPercentage = marginPercentage;
                    Summary.RemainingTargetPercentage = remainingTargetPercentage;
                });

                Debug.WriteLine($"摘要數據: Target=${targetValue}M, Achievement=${completedAchievementM}M ({achievementPercentage}%), " +
                               $"Remaining to target=${remainingToTargetM}M, Actual Remaining=${actualRemainingM}M");
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
                            TotalAchievement = 5,
                            TotalMargin = 5,
                            RemainingTarget = 5,
                            ActualRemaining = 5,
                            AchievementPercentage = 50,
                            MarginPercentage = 100,
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
                    LoadSampleChartData();
                    return;
                }

                // 按月份分組 - 使用已過濾的數據，現在全部都是有完成日期的記錄
                var monthlyData = _filteredData
                    .GroupBy(x => new
                    {
                        Year = x.CompletionDate.Value.Year,  // 這裡不再需要做空值檢查，因為已過濾
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
                    .ToList();

                // 載入圖表數據
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

                // 更新 UI
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    TargetVsAchievementData = newTargetVsAchievementData;
                    AchievementTrendData = newAchievementTrendData;
                });

                Debug.WriteLine($"圖表數據已更新，共 {newTargetVsAchievementData.Count} 個數據點");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入圖表數據時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                LoadSampleChartData();
            }
        }

        // 載入排行榜數據
        private async Task LoadLeaderboardDataAsync()
        {
            try
            {
                if (_filteredData == null)
                {
                    FilterDataByDateRange();
                }

                // 再次確認數據可用性
                if (_filteredData == null || !_filteredData.Any())
                {
                    LoadSampleLeaderboardData();
                    return;
                }

                // 因為現在 _filteredData 已經只包含完成日期在選定範圍內的記錄
                // 所以不需要再根據完成日期進行過濾
                var statusFiltered = _filteredData;

                Debug.WriteLine($"用於計算的記錄數: {statusFiltered.Count} 條");

                // 根據視圖類型載入相應數據
                switch (ViewType)
                {
                    case "ByProduct":
                        await LoadProductDataAsync(statusFiltered);
                        break;
                    case "ByRep":
                        await LoadSalesRepDataAsync(statusFiltered);
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入排行榜數據時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                LoadSampleLeaderboardData();
            }
        }


        // 載入產品數據
        private async Task LoadProductDataAsync(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine($"載入產品數據，資料筆數: {data.Count}");

                // 輸出前幾條記錄的詳細信息，用於調試
                foreach (var item in data.Take(Math.Min(5, data.Count)))
                {
                    Debug.WriteLine($"樣本數據: 接收日期={item.ReceivedDate:yyyy-MM-dd}, " +
                                   $"完成日期={item.CompletionDate?.ToString("yyyy-MM-dd") ?? "未完成"}, " +
                                   $"產品類型={item.ProductType}, " +
                                   $"總佣金=${item.TotalCommission}, " +
                                   $"PO值=${item.POValue}");
                }

                var products = data
                    .GroupBy(x => x.ProductType)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                        TotalMargin = Math.Round(g.Sum(x => x.TotalCommission), 2),
                        POValue = Math.Round(g.Sum(x => x.POValue), 2)
                    })
                    .OrderByDescending(x => x.POValue)
                    .ToList();

                Debug.WriteLine($"分組後產品數: {products.Count}");

                if (products.Any())
                {
                    // 計算百分比
                    decimal totalPOValue = products.Sum(p => p.POValue);
                    foreach (var product in products)
                    {
                        product.PercentageOfTotal = totalPOValue > 0
                            ? Math.Round((product.POValue / totalPOValue), 1)
                            : 0;

                        Debug.WriteLine($"產品: {product.ProductType}, " +
                                       $"Agency: ${product.AgencyMargin}, " +
                                       $"BuyResell: ${product.BuyResellMargin}, " +
                                       $"Total: ${product.TotalMargin}, " +
                                       $"PO: ${product.POValue}, " +
                                       $"百分比: {product.PercentageOfTotal}%");
                    }

                    await MainThread.InvokeOnMainThreadAsync(() =>
                    {
                        ProductSalesData = new ObservableCollection<ProductSalesData>(products);
                    });

                    Debug.WriteLine($"成功載入 {products.Count} 項產品數據");
                }
                else
                {
                    Debug.WriteLine("沒有產品數據可顯示");
                    LoadSampleProductData();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入產品數據時發生錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                LoadSampleProductData();
            }
        }


        // 載入銷售代表數據
        private async Task LoadSalesRepDataAsync(List<SalesData> data)
        {
            try
            {
                var reps = data
                    .GroupBy(x => x.SalesRep)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new SalesLeaderboardItem
                    {
                        SalesRep = g.Key,
                        AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                        // Use TotalCommission directly instead of summing
                        TotalMargin = Math.Round(g.Sum(x => x.TotalCommission), 2)
                    })
                    .OrderByDescending(x => x.TotalMargin)
                    .ToList();

                // Set ranking
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
                LoadSampleSalesRepData();
            }
        }

        // 標準化部門名稱
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

            return department;
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

        // 載入樣本數據（當有錯誤或無數據時使用）
        private void LoadSampleData()
        {
            LoadSampleChartData();
            LoadSampleLeaderboardData();
        }

        // 載入樣本圖表數據
        private void LoadSampleChartData()
        {
            // 目標 vs 達成圖表
            var targetVsAchievementSample = new ObservableCollection<ChartData>
            {
                new ChartData { Label = "1月", Target = 2.5m, Achievement = 2.0m },
                new ChartData { Label = "2月", Target = 2.5m, Achievement = 2.3m },
                new ChartData { Label = "3月", Target = 2.5m, Achievement = 2.7m },
                new ChartData { Label = "4月", Target = 2.5m, Achievement = 2.1m },
                new ChartData { Label = "5月", Target = 2.5m, Achievement = 3.0m },
                new ChartData { Label = "6月", Target = 2.5m, Achievement = 3.2m }
            };
            TargetVsAchievementData = targetVsAchievementSample;

            // 達成趨勢圖表
            var achievementTrendSample = new ObservableCollection<ChartData>
            {
                new ChartData { Label = "1月", Achievement = 2.0m },
                new ChartData { Label = "2月", Achievement = 2.3m },
                new ChartData { Label = "3月", Achievement = 2.7m },
                new ChartData { Label = "4月", Achievement = 2.1m },
                new ChartData { Label = "5月", Achievement = 3.0m },
                new ChartData { Label = "6月", Achievement = 3.2m }
            };
            AchievementTrendData = achievementTrendSample;
        }

        // 載入樣本排行榜數據
        private void LoadSampleLeaderboardData()
        {
            LoadSampleProductData();
            LoadSampleSalesRepData();
        }

        // 載入樣本產品數據
        private void LoadSampleProductData()
        {
            var productSample = new ObservableCollection<ProductSalesData>
    {
        new ProductSalesData
        {
            ProductType = "Thermal",
            AgencyMargin = 744855.43m,
            BuyResellMargin = 116206.36m,
            TotalMargin = 861061.79m,
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

            ProductSalesData = productSample;
        }

        // 載入樣本銷售代表數據
        private void LoadSampleSalesRepData()
        {
            var salesRepSample = new ObservableCollection<SalesLeaderboardItem>
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

            SalesLeaderboard = salesRepSample;
        }
        #endregion // 结束 #region 樣本數据

        #endregion // 结束 #region 辅助方法

    }
}