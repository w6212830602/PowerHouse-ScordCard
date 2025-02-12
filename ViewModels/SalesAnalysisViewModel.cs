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

        public SalesAnalysisViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            _excelService.DataUpdated += OnDataUpdated;

            // 初始化摘要和排行榜
            summary = new SalesAnalysisSummary();
            leaderboard = new ObservableCollection<SalesRepPerformance>();

            Task.Run(async () => await LoadDataAsync());
        }

        [RelayCommand]
        private async Task Refresh()
        {
            await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            if (IsLoading) return;

            try
            {
                IsLoading = true;
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                await ProcessDataAsync(data);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading data: {ex.Message}");
                // TODO: 顯示錯誤訊息
            }
            finally
            {
                IsLoading = false;
            }
        }

        private Task ProcessDataAsync(List<SalesData> data)
        {
            return MainThread.InvokeOnMainThreadAsync(() =>
            {
                // 根據選擇的時間範圍過濾數據
                var filteredData = FilterDataByTimeRange(data);

                // 計算摘要
                Summary = new SalesAnalysisSummary
                {
                    TotalTarget = filteredData.Sum(d => d.POValue),
                    TotalAchievement = filteredData.Sum(d => d.VertivValue),
                    TotalMargin = filteredData.Sum(d => d.TotalCommission),
                    TopPerformers = GetTopPerformers(filteredData)
                };

                // 更新排行榜
                Leaderboard = new ObservableCollection<SalesRepPerformance>(
                    GetTopPerformers(filteredData)
                );
            });
        }

        private List<SalesData> FilterDataByTimeRange(List<SalesData> data)
        {
            return data.Where(d =>
            {
                if (string.IsNullOrEmpty(SelectedTimeRange) || SelectedTimeRange == "YTD")
                    return d.ReceivedDate.Year == DateTime.Now.Year;

                return SelectedTimeRange switch
                {
                    "Q1" => d.ReceivedDate.Month >= 1 && d.ReceivedDate.Month <= 3,
                    "Q2" => d.ReceivedDate.Month >= 4 && d.ReceivedDate.Month <= 6,
                    "Q3" => d.ReceivedDate.Month >= 7 && d.ReceivedDate.Month <= 9,
                    "Q4" => d.ReceivedDate.Month >= 10 && d.ReceivedDate.Month <= 12,
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

        private void OnDataUpdated(object sender, DateTime e)
        {
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await LoadDataAsync();
            });
        }
    }
}