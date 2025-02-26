using System;
using System.Collections.ObjectModel;
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

        public SalesAnalysisViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            LoadDataCommand = new AsyncRelayCommand(LoadDataAsync);
            _switchViewCommand = new RelayCommand<string>(ExecuteSwitchView);

            // 監聽日期變化
            PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(StartDate) || e.PropertyName == nameof(EndDate))
                {
                    Debug.WriteLine($"Date changed - Start: {StartDate:yyyy/MM/dd}, End: {EndDate:yyyy/MM/dd}");
                    LoadDataCommand.ExecuteAsync(null);
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
                // 初始載入數據
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                _allSalesData = data;

                // 設定初始日期範圍（如果尚未設定）
                if (StartDate == default || EndDate == default)
                {
                    var dates = _allSalesData.Select(x => x.ReceivedDate).OrderBy(x => x).ToList();
                    if (dates.Any())
                    {
                        StartDate = dates.First().Date;
                        EndDate = dates.Last().Date;
                        Debug.WriteLine($"Initial date range set - Start: {StartDate:yyyy/MM/dd}, End: {EndDate:yyyy/MM/dd}");
                    }
                }
               
                await LoadDataCommand.ExecuteAsync(null);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in InitializeAsync: {ex.Message}");
            }
        }
        public IAsyncRelayCommand LoadDataCommand { get; }

        private void ExecuteSwitchView(string viewType)
        {
            if (!string.IsNullOrEmpty(viewType))
            {
                IsSummaryView = viewType.ToLower() == "summary";
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
                if (_allSalesData == null)
                {
                    var (data, lastUpdated) = await _excelService.LoadDataAsync();
                    _allSalesData = data;
                }

                // 正確處理日期過濾
                var startDate = StartDate.Date;
                var endDate = EndDate.Date;

                // 過濾數據：只比較日期部分
                var filteredData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= startDate && x.ReceivedDate.Date <= endDate)
                    .ToList();

                Debug.WriteLine($"Filtered data count: {filteredData.Count}");

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

                Debug.WriteLine($"Monthly data points in range: {monthlyData.Count}");
                foreach (var point in monthlyData)
                {
                    Debug.WriteLine($"Data point: {point.Label}, Target: {point.Target}, Achievement: {point.Achievement}");
                }

                // 更新圖表數據
                TargetVsAchievementData.Clear();
                AchievementTrendData.Clear();
                foreach (var item in monthlyData)
                {
                    TargetVsAchievementData.Add(item);
                    AchievementTrendData.Add(item);
                }

                // 更新匯總數據
                Summary = new SalesAnalysisSummary
                {
                    TotalTarget = Math.Round(filteredData.Sum(x => x.POValue) / 1000000m, 2),
                    TotalAchievement = Math.Round(filteredData.Sum(x => x.VertivValue) / 1000000m, 2),
                    TotalMargin = Math.Round(filteredData.Sum(x => x.TotalCommission) / 1000000m, 2)
                };

                // 更新排行榜
                var leaderboardData = filteredData
                    .GroupBy(x => x.SalesRep)
                    .Select(g => new SalesRepPerformance
                    {
                        SalesRep = g.Key,
                        Achievement = Math.Round(g.Sum(x => x.VertivValue), 2),
                        Commission = Math.Round(g.Sum(x => x.TotalCommission), 2),
                        Target = Math.Round(g.Sum(x => x.POValue), 2)
                    })
                    .OrderByDescending(x => x.Achievement)
                    .Take(5)
                    .ToList();

                Leaderboard.Clear();
                foreach (var item in leaderboardData)
                {
                    Leaderboard.Add(item);
                }

                UpdateChartAxes();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in LoadDataAsync: {ex.Message}");
                await Application.Current.MainPage.DisplayAlert("Error", ex.Message, "OK");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async void OnDateRangeChanged()
        {
            Debug.WriteLine($"Date range changed: {StartDate:yyyy/MM/dd} - {EndDate:yyyy/MM/dd}");
            await LoadDataCommand.ExecuteAsync(null);
        }

        private void UpdateChartAxes()
        {
            if (TargetVsAchievementData?.Any() == true)
            {
                var maxValue = Math.Max(
                    TargetVsAchievementData.Max(x => Convert.ToDouble(x.Target)),
                    TargetVsAchievementData.Max(x => Convert.ToDouble(x.Achievement))
                );

                YAxisMaximum = Math.Ceiling(maxValue * 1.2);
                Debug.WriteLine($"Updated Y-axis maximum to {YAxisMaximum}");
            }
        }
    }
}