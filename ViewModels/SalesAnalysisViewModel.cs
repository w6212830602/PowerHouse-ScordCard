using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ScoreCard.Models;
using ScoreCard.Services;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.Maui;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ScoreCard.ViewModels
{
    public partial class SalesAnalysisViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;

        [ObservableProperty]
        private DateTime startDate = DateTime.Now.AddDays(-30);

        [ObservableProperty]
        private DateTime endDate = DateTime.Now;

        [ObservableProperty]
        private bool isLoading;

        [ObservableProperty]
        private string selectedTimeRange = "YTD";

        [ObservableProperty]
        private string selectedViewType = "By Rep";

        [ObservableProperty]
        private SalesAnalysisSummary summary;

        [ObservableProperty]
        private ObservableCollection<SalesRepPerformance> leaderboard;

        [ObservableProperty]
        private ObservableCollection<MonthlyPerformance> monthlyPerformance;

        [ObservableProperty]
        private bool isSummaryView = true;

        [ObservableProperty]
        private ObservableCollection<ChartData> targetVsAchievementData;

        [ObservableProperty]
        private ObservableCollection<ChartData> achievementTrendData;

        public SalesAnalysisViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            _excelService.DataUpdated += OnDataUpdated;

            summary = new SalesAnalysisSummary();
            leaderboard = new ObservableCollection<SalesRepPerformance>();
            monthlyPerformance = new ObservableCollection<MonthlyPerformance>();
            targetVsAchievementData = new ObservableCollection<ChartData>();
            achievementTrendData = new ObservableCollection<ChartData>();

            Task.Run(async () => await LoadDataAsync());
        }

        [RelayCommand]
        private async Task Refresh()
        {
            await LoadDataAsync();
        }

        private void OnDataUpdated(object sender, DateTime e)
        {
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await LoadDataAsync();
            });
        }

        private async Task LoadDataAsync()
        {
            if (IsLoading) return;

            try
            {
                IsLoading = true;
                var (data, lastUpdated) = await _excelService.LoadDataAsync();

                // 先過濾時間範圍
                var filteredData = FilterDataByTimeRange(data);

                // 更新摘要信息
                Summary = new SalesAnalysisSummary
                {
                    TotalTarget = filteredData.Sum(d => d.POValue),
                    TotalAchievement = filteredData.Sum(d => d.VertivValue),
                    TotalMargin = filteredData.Sum(d => d.TotalCommission),
                    TopPerformers = GetTopPerformers(filteredData),
                    MonthlyData = GetMonthlyData(filteredData)
                };

                // 更新圖表數據
                UpdateChartData(filteredData);

                // 更新排行榜
                Leaderboard = new ObservableCollection<SalesRepPerformance>(Summary.TopPerformers);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading data: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
            }
        }

        private List<SalesData> FilterDataByTimeRange(List<SalesData> data)
        {
            if (SelectedTimeRange == "Custom")
            {
                return data.Where(d => d.ReceivedDate >= StartDate && d.ReceivedDate <= EndDate).ToList();
            }

            return data.Where(d =>
            {
                return SelectedTimeRange switch
                {
                    "YTD" => d.ReceivedDate.Year == DateTime.Now.Year,
                    "Q1" => d.Quarter == 1,
                    "Q2" => d.Quarter == 2,
                    "Q3" => d.Quarter == 3,
                    "Q4" => d.Quarter == 4,
                    _ => true
                };
            }).ToList();
        }

        private List<SalesRepPerformance> GetTopPerformers(List<SalesData> data)
        {
            return data.GroupBy(d => d.SalesRep)
                      .Select(g => new SalesRepPerformance
                      {
                          SalesRep = g.Key,
                          Achievement = g.Sum(d => d.VertivValue),
                          Commission = g.Sum(d => d.TotalCommission),
                          Target = g.Sum(d => d.POValue)
                      })
                      .OrderByDescending(p => p.Achievement)
                      .Take(10)
                      .ToList();
        }

        private List<MonthlyPerformance> GetMonthlyData(List<SalesData> data)
        {
            return data.GroupBy(x => new DateTime(x.ReceivedDate.Year, x.ReceivedDate.Month, 1))
                       .Select(g => new MonthlyPerformance
                       {
                           Month = g.Key,
                           Target = g.Sum(x => x.POValue),
                           Achievement = g.Sum(x => x.VertivValue),
                           Margin = g.Sum(x => x.TotalCommission)
                       })
                       .OrderBy(x => x.Month)
                       .ToList();
        }


        private void UpdateChartData(List<SalesData> filteredData)
        {
            var monthlyData = GetMonthlyData(filteredData);

            TargetVsAchievementData = new ObservableCollection<ChartData>(
                monthlyData.Select(m => new ChartData
                {
                    Label = m.Month.ToString("MMM"),
                    Target = m.Target / 1000,
                    Achievement = m.Achievement / 1000
                })
            );

            AchievementTrendData = new ObservableCollection<ChartData>(
                monthlyData.Select(m => new ChartData
                {
                    Label = m.Month.ToString("MMM"),
                    Achievement = m.Achievement / 1000
                })
            );
        }

        partial void OnSelectedTimeRangeChanged(string value)
        {
            LoadDataAsync().ConfigureAwait(false);
        }

        partial void OnStartDateChanged(DateTime value)
        {
            LoadDataAsync().ConfigureAwait(false);
        }

        partial void OnEndDateChanged(DateTime value)
        {
            LoadDataAsync().ConfigureAwait(false);
        }

        partial void OnSelectedViewTypeChanged(string value)
        {
            LoadDataAsync().ConfigureAwait(false);
        }

        [RelayCommand]
        private void SwitchView()
        {
            IsSummaryView = !IsSummaryView;
        }
    }
}