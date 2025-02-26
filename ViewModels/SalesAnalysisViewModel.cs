using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ScoreCard.Models;
using ScoreCard.Services;
using System.Linq;
using System.Diagnostics;

namespace ScoreCard.ViewModels
{
    public partial class SalesAnalysisViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;
        private List<SalesData> _allSalesData;

        [ObservableProperty]
        private DateTime _startDate;

        [ObservableProperty]
        private DateTime _endDate;

        [ObservableProperty]
        private bool _isSummaryView = true;

        [ObservableProperty]
        private bool _isLoading;

        [ObservableProperty]
        private SalesAnalysisSummary _summary;

        // 改用ObservableProperty來定義圖表相關屬性
        [ObservableProperty]
        private ObservableCollection<Models.ChartData> _targetVsAchievementData = new();

        [ObservableProperty]
        private ObservableCollection<Models.ChartData> _achievementTrendData = new();

        [ObservableProperty]
        private ObservableCollection<SalesRepPerformance> _leaderboard = new();

        [ObservableProperty]
        private double _yAxisMaximum = 20;

        private ICommand _switchViewCommand;
        public ICommand SwitchViewCommand => _switchViewCommand;

        // 使用公共屬性暴露 LoadDataCommand
        public IAsyncRelayCommand LoadDataCommand { get; }

        // 添加日期變更的特殊處理
        public SalesAnalysisViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            LoadDataCommand = new AsyncRelayCommand(LoadDataAsync);
            _switchViewCommand = new RelayCommand<string>(ExecuteSwitchView);

            // 初始化集合
            _targetVsAchievementData = new ObservableCollection<Models.ChartData>();
            _achievementTrendData = new ObservableCollection<Models.ChartData>();
            _leaderboard = new ObservableCollection<SalesRepPerformance>();

            // 初始化 Summary
            _summary = new SalesAnalysisSummary
            {
                TotalTarget = 0,
                TotalAchievement = 0,
                TotalMargin = 0
            };

            // 顯式監聽 StartDate 和 EndDate 屬性變更
            PropertyChanged += async (s, e) =>
            {
                if (e.PropertyName == nameof(StartDate))
                {
                    Debug.WriteLine($"ViewModel - StartDate changed to: {StartDate:yyyy/MM/dd}");
                    await OnDateRangeChangedAsync();
                }
                else if (e.PropertyName == nameof(EndDate))
                {
                    Debug.WriteLine($"ViewModel - EndDate changed to: {EndDate:yyyy/MM/dd}");
                    await OnDateRangeChangedAsync();
                }
            };

            // 初始載入
            InitializeAsync();
        }

        // 初始化
        private async void InitializeAsync()
        {
            try
            {
                IsLoading = true;

                // 初始載入數據
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                _allSalesData = data;

                // 設定初始日期範圍
                if (StartDate == default || EndDate == default)
                {
                    var dates = _allSalesData.Select(x => x.ReceivedDate).OrderBy(x => x).ToList();
                    if (dates.Any())
                    {
                        // 避免在這裡觸發重複載入
                        var startDate = dates.First().Date;
                        var endDate = dates.Last().Date;

                        // 先設置結束日期再設置開始日期，避免兩個日期變更都觸發數據載入
                        _endDate = endDate;  // 直接設置欄位避免觸發事件
                        _startDate = startDate;
                        OnPropertyChanged(nameof(EndDate)); // 然後手動觸發UI更新
                        OnPropertyChanged(nameof(StartDate));

                        Debug.WriteLine($"Initial date range set - Start: {StartDate:yyyy/MM/dd}, End: {EndDate:yyyy/MM/dd}");
                    }
                    else
                    {
                        // 若沒有數據，設置合理的預設值
                        _endDate = DateTime.Now.Date;
                        _startDate = DateTime.Now.AddMonths(-3).Date;
                        OnPropertyChanged(nameof(EndDate));
                        OnPropertyChanged(nameof(StartDate));
                    }
                }

                // 顯式調用數據載入
                await LoadDataAsync();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in InitializeAsync: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                IsLoading = false;
            }
        }
        public async Task ReloadDataAsync()
        {
            Debug.WriteLine("手動觸發數據重載");
            await LoadDataAsync();
            await ForceRefreshCharts(); // 強制圖表刷新
        }


        private void ExecuteSwitchView(string viewType)
        {
            if (!string.IsNullOrEmpty(viewType))
            {
                IsSummaryView = viewType.ToLower() == "summary";
            }
        }

        // 專用的日期範圍變更處理方法
        private async Task OnDateRangeChangedAsync()
        {
            // 確保日期範圍有效並且兩個都已設置
            if (StartDate != default && EndDate != default && StartDate <= EndDate)
            {
                Debug.WriteLine($"Valid date range: {StartDate:yyyy/MM/dd} to {EndDate:yyyy/MM/dd} - Reloading data...");
                await LoadDataAsync();
            }
            else
            {
                Debug.WriteLine($"Invalid date range: {StartDate:yyyy/MM/dd} to {EndDate:yyyy/MM/dd} - Skipping reload");
            }
        }

        // 載入數據
        private async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;
                Debug.WriteLine($"LoadDataAsync - Using date range: {StartDate:yyyy/MM/dd} to {EndDate:yyyy/MM/dd}");

                // 確保已有數據載入
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    var (data, lastUpdated) = await _excelService.LoadDataAsync();
                    _allSalesData = data;
                    Debug.WriteLine($"Loaded {_allSalesData.Count} records from Excel");
                }

                // 強制使用當前的 StartDate 和 EndDate
                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1); // 包含結束日期的整天

                Debug.WriteLine($"實際過濾使用的日期範圍: {startDate:yyyy-MM-dd} 到 {endDate:yyyy-MM-dd}");

                // 過濾數據
                var filteredData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate && x.ReceivedDate.Date <= endDate.Date)
                    .ToList();

                Debug.WriteLine($"過濾後數據: {filteredData.Count} 條記錄在 {startDate:yyyy-MM-dd} 和 {endDate:yyyy-MM-dd} 之間");

                // 資料檢查點 - 輸出找到的記錄 
                foreach (var record in filteredData.Take(5))
                {
                    Debug.WriteLine($"記錄: {record.ReceivedDate:yyyy-MM-dd} - POValue: {record.POValue}, VertivValue: {record.VertivValue}");
                }

                // 按月份分組並排序
                var monthlyData = filteredData
                    .GroupBy(x => new
                    {
                        Year = x.ReceivedDate.Year,
                        Month = x.ReceivedDate.Month
                    })
                    .Select(g => new Models.ChartData
                    {
                        Label = $"{g.Key.Year}/{g.Key.Month:D2}",
                        Target = Math.Round(g.Sum(x => x.POValue) / 1000000m, 2),
                        Achievement = Math.Round(g.Sum(x => x.VertivValue) / 1000000m, 2)
                    })
                    .OrderBy(x => x.Label)
                    .ToList();

                Debug.WriteLine($"Generated {monthlyData.Count} monthly data points");
                foreach (var point in monthlyData)
                {
                    Debug.WriteLine($"  {point.Label}: Target=${point.Target}M, Achievement=${point.Achievement}M");
                }

                // 在主線程上進行 UI 更新
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    // 強制 UI 更新 - 創建新的集合實例
                    var newTargetVsAchievementData = new ObservableCollection<Models.ChartData>(
                        monthlyData.Select(item => new Models.ChartData
                        {
                            Label = item.Label,
                            Target = item.Target,
                            Achievement = item.Achievement
                        })
                    );

                    TargetVsAchievementData = newTargetVsAchievementData;

                    var newAchievementTrendData = new ObservableCollection<Models.ChartData>(
                        monthlyData.Select(item => new Models.ChartData
                        {
                            Label = item.Label,
                            Target = item.Target,
                            Achievement = item.Achievement
                        })
                    );

                    AchievementTrendData = newAchievementTrendData;

                    // 手動通知
                    OnPropertyChanged(nameof(TargetVsAchievementData));
                    OnPropertyChanged(nameof(AchievementTrendData));

                    // 更新匯總數據
                    Summary = new SalesAnalysisSummary
                    {
                        TotalTarget = Math.Round(filteredData.Sum(x => x.POValue) / 1000000m, 2),
                        TotalAchievement = Math.Round(filteredData.Sum(x => x.VertivValue) / 1000000m, 2),
                        TotalMargin = Math.Round(filteredData.Sum(x => x.TotalCommission) / 1000000m, 2)
                    };

                    Debug.WriteLine($"Updated summary: Target=${Summary.TotalTarget}M, Achievement=${Summary.TotalAchievement}M, Margin=${Summary.TotalMargin}M");

                    // 更新排行榜
                    var newLeaderboard = new ObservableCollection<SalesRepPerformance>();

                    var leaderboardData = filteredData
                        .GroupBy(x => x.SalesRep)
                        .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                        .Select(g => new SalesRepPerformance
                        {
                            SalesRep = g.Key,
                            Achievement = Math.Round(g.Sum(x => x.VertivValue), 2),
                            Commission = Math.Round(g.Sum(x => x.TotalCommission), 2),
                            Target = Math.Round(g.Sum(x => x.POValue), 2)
                        })
                        .OrderByDescending(x => x.Achievement)
                        .Take(10)
                        .ToList();

                    Leaderboard = new ObservableCollection<SalesRepPerformance>(leaderboardData);
                    OnPropertyChanged(nameof(Leaderboard));

                    UpdateChartAxes();
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in LoadDataAsync: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                await Application.Current.MainPage.DisplayAlert("資料載入錯誤", $"無法載入銷售數據: {ex.Message}", "確定");
            }
            finally
            {
                IsLoading = false;
            }
        }

        public async Task ForceRefreshCharts()
        {
            await MainThread.InvokeOnMainThreadAsync(() => {
                // 通知所有繫結屬性已更改
                OnPropertyChanged(nameof(TargetVsAchievementData));
                OnPropertyChanged(nameof(AchievementTrendData));
                OnPropertyChanged(nameof(YAxisMaximum));
                OnPropertyChanged(nameof(Leaderboard));

                // 觸發整個 ViewModel 刷新
                OnPropertyChanged(string.Empty);
            });
        }


        private void UpdateChartAxes()
        {
            try
            {
                if (TargetVsAchievementData?.Any() == true)
                {
                    var maxTarget = TargetVsAchievementData.Max(x => Convert.ToDouble(x.Target));
                    var maxAchievement = TargetVsAchievementData.Max(x => Convert.ToDouble(x.Achievement));
                    var maxValue = Math.Max(maxTarget, maxAchievement);

                    // 設置一個稍大的最大值以便於查看
                    YAxisMaximum = Math.Ceiling(maxValue * 1.2);
                    Debug.WriteLine($"Updated Y-axis maximum to {YAxisMaximum}");
                }
                else
                {
                    // 若無數據，設置預設值
                    YAxisMaximum = 5;
                    Debug.WriteLine("No data points, set default Y-axis maximum to 5");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in UpdateChartAxes: {ex.Message}");
                YAxisMaximum = 5;
            }
        }
    }
}