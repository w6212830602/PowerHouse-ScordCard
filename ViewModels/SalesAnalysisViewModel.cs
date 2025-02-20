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

            // 設定初始日期範圍為當前會計年度
            var currentDate = DateTime.Now;
            var fiscalYearStart = currentDate.Month >= 8 ?
                new DateTime(currentDate.Year, 8, 1) :
                new DateTime(currentDate.Year - 1, 8, 1);

            StartDate = fiscalYearStart;
            EndDate = fiscalYearStart.AddYears(1).AddDays(-1);

            PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(StartDate) || e.PropertyName == nameof(EndDate))
                {
                    LoadDataCommand.ExecuteAsync(null);
                }
            };

            LoadDataCommand.ExecuteAsync(null);
        }

        public IAsyncRelayCommand LoadDataCommand { get; }

        private void ExecuteSwitchView(string viewType)
        {
            if (!string.IsNullOrEmpty(viewType))
            {
                IsSummaryView = viewType.ToLower() == "summary";
            }
        }

        private async Task LoadDataAsync()
        {
            try
            {
                IsLoading = true;

                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                Debug.WriteLine($"Loaded data count: {data?.Count ?? 0}");

                _allSalesData = data;

                // 過濾數據：忽略時間部分，只比較日期
                var filteredData = _allSalesData
                    .Where(x => x.ReceivedDate.Date >= StartDate.Date && x.ReceivedDate.Date <= EndDate.Date)
                    .ToList();

                Debug.WriteLine($"Filtered data count: {filteredData.Count}, Date range: {StartDate:yyyy-MM-dd} to {EndDate:yyyy-MM-dd}");

                // 按月份分組計算數據
                var monthlyData = filteredData
                    .GroupBy(x => new { Year = x.ReceivedDate.Year, Month = x.ReceivedDate.Month })
                    .Select(g => new Models.ChartData
                    {
                        Label = $"{g.Key.Year}/{g.Key.Month:D2}",
                        Target = Math.Round(g.Sum(x => x.POValue) / 1000000m, 2),
                        Achievement = Math.Round(g.Sum(x => x.VertivValue) / 1000000m, 2)
                    })
                    .OrderBy(x => x.Label)
                    .ToList();

                Debug.WriteLine($"Monthly data points: {monthlyData.Count}");
                foreach (var point in monthlyData)
                {
                    Debug.WriteLine($"Month: {point.Label}, Target: {point.Target}, Achievement: {point.Achievement}");
                }

                // 清除並更新圖表數據
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

                // 更新Y軸最大值
                UpdateChartAxes();

                Debug.WriteLine($"Data load completed. Summary - Target: {Summary.TotalTarget}M, Achievement: {Summary.TotalAchievement}M");
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