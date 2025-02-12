using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;
using ScoreCard.Services;
using ScoreCard.Models;
using System.Diagnostics;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;


namespace ScoreCard.ViewModels
{
    public partial class DashboardViewModel : ObservableObject
    {
        // 基本屬性
        [ObservableProperty]
        private string selectedOption;

        [ObservableProperty]
        private List<string> options;

        [ObservableProperty]
        private string lastUpdated = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

        // 年度總體數據
        [ObservableProperty]
        private decimal annualTarget;

        [ObservableProperty]
        private decimal achievement;

        [ObservableProperty]
        private decimal remaining;

        [ObservableProperty]
        private double achievementProgress;

        // 季度基本目標
        [ObservableProperty]
        private decimal q1Target;

        [ObservableProperty]
        private decimal q2Target;

        [ObservableProperty]
        private decimal q3Target;

        [ObservableProperty]
        private decimal q4Target;

        // 季度達成金額
        [ObservableProperty]
        private decimal q1Achieved;

        [ObservableProperty]
        private decimal q2Achieved;

        [ObservableProperty]
        private decimal q3Achieved;

        [ObservableProperty]
        private decimal q4Achieved;

        // 季度轉移金額
        [ObservableProperty]
        private decimal q2Added;

        [ObservableProperty]
        private decimal q3Added;

        [ObservableProperty]
        private decimal q4Added;

        // 通知
        [ObservableProperty]
        private ObservableCollection<NotificationItem> notifications = new();

        [ObservableProperty]
        private bool isLoading;

        // Excel服務和設定
        private readonly IExcelService _excelService;
        private readonly TargetSettings _targetSettings;
        private CancellationTokenSource _cancellationTokenSource;

        public DashboardViewModel(IExcelService excelService)
        {
            Debug.WriteLine("DashboardViewModel 建構函數開始初始化");
            _excelService = excelService;
            _excelService.DataUpdated += OnDataUpdated;
            _cancellationTokenSource = new CancellationTokenSource();

            // 讀取設定檔
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            _targetSettings = config.GetSection("TargetSettings").Get<TargetSettings>();

            // 取得目前財年
            var currentDate = DateTime.Now;
            var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;

            Debug.WriteLine($"當前日期: {currentDate}, 當前財年: {currentFiscalYear}");

            // 設定財年選項（降序排列）
            Options = new List<string> {
                $"FY{currentFiscalYear + 1}",
                $"FY{currentFiscalYear}",
                $"FY{currentFiscalYear - 1}"
            };

            // 設定預設選項為當前財年
            SelectedOption = $"FY{currentFiscalYear}";

            InitializeNotifications();
            InitializeAsync();
            Debug.WriteLine($"DashboardViewModel 初始化完成，選擇的財年: {SelectedOption}");
        }

        private void UpdateTargets(int fiscalYear)
        {
            var yearTarget = _targetSettings.CompanyTargets
                .FirstOrDefault(t => t.FiscalYear == fiscalYear);

            if (yearTarget != null)
            {
                AnnualTarget = yearTarget.AnnualTarget;
                Q1Target = yearTarget.Q1Target;
                Q2Target = yearTarget.Q2Target;
                Q3Target = yearTarget.Q3Target;
                Q4Target = yearTarget.Q4Target;

                Debug.WriteLine($"已載入 FY{fiscalYear} 目標值:");
                Debug.WriteLine($"Annual: ${AnnualTarget:N0}");
                Debug.WriteLine($"Q1: ${Q1Target:N0}");
                Debug.WriteLine($"Q2: ${Q2Target:N0}");
                Debug.WriteLine($"Q3: ${Q3Target:N0}");
                Debug.WriteLine($"Q4: ${Q4Target:N0}");
            }
            else
            {
                Debug.WriteLine($"警告: 找不到 FY{fiscalYear} 的目標值設定");
            }
        }

        partial void OnSelectedOptionChanged(string value)
        {
            Debug.WriteLine($"選擇的財年改變為: {value}");
            if (!string.IsNullOrEmpty(value))
            {
                LoadDataAsync();
            }
        }

        private int GetSelectedFiscalYear()
        {
            if (string.IsNullOrEmpty(SelectedOption))
            {
                var currentDate = DateTime.Now;
                return currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;
            }

            return int.Parse(SelectedOption.Replace("FY", ""));
        }

        private async void InitializeAsync()
        {
            try
            {
                Debug.WriteLine("開始初始化 Dashboard");
                await LoadDataAsync();
                Debug.WriteLine("開始監控 Excel 檔案變更");
                await _excelService.MonitorFileChangesAsync(_cancellationTokenSource.Token);
                Debug.WriteLine("Excel 檔案監控已啟動");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"初始化過程發生錯誤: {ex.Message}");
                Debug.WriteLine($"錯誤詳細資訊: {ex.StackTrace}");
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    Notifications.Add(new NotificationItem { Message = "Error initializing dashboard" });
                });
            }
        }

        private async Task LoadDataAsync()
        {
            try
            {
                Debug.WriteLine("開始載入 Excel 數據");
                IsLoading = true;
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                Debug.WriteLine($"成功載入數據，共 {data.Count} 筆記錄");

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    UpdateDashboard(data);
                    LastUpdated = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                    Debug.WriteLine("數據已更新到 UI");
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入數據時發生錯誤: {ex.Message}");
                Debug.WriteLine($"錯誤詳細資訊: {ex.StackTrace}");
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    Notifications.Add(new NotificationItem { Message = "Error loading data from Excel" });
                });
            }
            finally
            {
                IsLoading = false;
                Debug.WriteLine("LoadDataAsync 完成");
            }
        }

        private void UpdateDashboard(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine("開始更新儀表板數據");

                // 獲取選擇的財年並更新目標值
                var selectedFiscalYear = GetSelectedFiscalYear();
                UpdateTargets(selectedFiscalYear);

                // 篩選選擇的財年數據
                var yearData = data.Where(x => x.FiscalYear == selectedFiscalYear).ToList();
                Debug.WriteLine($"目前財年: FY{selectedFiscalYear}, 數據筆數: {yearData.Count}");

                // 按季度分組計算實際達成值
                var quarterlyData = yearData.GroupBy(x => x.Quarter)
                                       .ToDictionary(g => g.Key, g => new
                                       {
                                           Achieved = g.Sum(x => x.TotalCommission),
                                           MonthlyBreakdown = g.GroupBy(x => x.ReceivedDate.Month)
                                                             .Select(m => new
                                                             {
                                                                 Month = m.Key,
                                                                 Commission = m.Sum(x => x.TotalCommission)
                                                             }).ToList()
                                       });

                // 更新各季度達成值
                Q1Achieved = quarterlyData.GetValueOrDefault(1)?.Achieved ?? 0;
                Q2Achieved = quarterlyData.GetValueOrDefault(2)?.Achieved ?? 0;
                Q3Achieved = quarterlyData.GetValueOrDefault(3)?.Achieved ?? 0;
                Q4Achieved = quarterlyData.GetValueOrDefault(4)?.Achieved ?? 0;

                // 計算總體達成值和進度
                var totalAchieved = Q1Achieved + Q2Achieved + Q3Achieved + Q4Achieved;
                Achievement = AnnualTarget > 0 ? Math.Round((totalAchieved / AnnualTarget) * 100, 1) : 0;
                Remaining = AnnualTarget - totalAchieved;
                AchievementProgress = AnnualTarget > 0 ? (double)(totalAchieved / AnnualTarget) : 0;

                // 更新通知
                UpdateNotifications();
                Debug.WriteLine("儀表板數據更新完成");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新儀表板時發生錯誤: {ex.Message}");
                Debug.WriteLine($"錯誤詳細資訊: {ex.StackTrace}");
                Notifications.Add(new NotificationItem { Message = "Error updating dashboard data" });
            }
        }

        private void UpdateNotifications()
        {
            Debug.WriteLine("開始更新通知");

            var newNotifications = new List<NotificationItem>();

            // Q1 未達成通知
            if (Q1Carried > 0)
            {
                Debug.WriteLine($"Q1 未達成通知: ${Q1Carried:N0}");
                newNotifications.Add(new NotificationItem
                {
                    Message = $"Q1 Target not achieved! ${Q1Carried:N0} carried to Q2"
                });
            }

            // Q2 未達成通知
            if (Q2Carried > 0)
            {
                Debug.WriteLine($"Q2 未達成通知: ${Q2Carried:N0}");
                newNotifications.Add(new NotificationItem
                {
                    Message = $"Q2 Target not achieved! ${Q2Carried:N0} carried to Q3"
                });
            }

            // Q3 未達成通知
            if (Q3Carried > 0)
            {
                Debug.WriteLine($"Q3 未達成通知: ${Q3Carried:N0}");
                newNotifications.Add(new NotificationItem
                {
                    Message = $"Q3 Target not achieved! ${Q3Carried:N0} carried to Q4"
                });
            }

            MainThread.BeginInvokeOnMainThread(() =>
            {
                Notifications.Clear();
                foreach (var notification in newNotifications)
                {
                    Notifications.Add(notification);
                }
            });

            Debug.WriteLine($"通知更新完成，共 {Notifications.Count} 條通知");
        }

        private void OnDataUpdated(object sender, DateTime lastUpdated)
        {
            Debug.WriteLine($"收到檔案更新通知，時間: {lastUpdated}");
            MainThread.InvokeOnMainThreadAsync(async () =>
            {
                LastUpdated = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                await LoadDataAsync();
                Debug.WriteLine("檔案更新處理完成");
            });
        }

        private void InitializeNotifications()
        {
            Debug.WriteLine("初始化通知列表");
            MainThread.BeginInvokeOnMainThread(() =>
            {
                Notifications = new ObservableCollection<NotificationItem>();
            });
            Debug.WriteLine("通知列表初始化完成");
        }
        // 清理資源
        public void Cleanup()
        {
            Debug.WriteLine("開始清理資源");
            if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                try
                {
                    _cancellationTokenSource.Cancel();
                    _cancellationTokenSource.Dispose();
                    _cancellationTokenSource = null;
                }
                catch (ObjectDisposedException)
                {
                    Debug.WriteLine("CancellationTokenSource 已經被釋放");
                }
            }
            Debug.WriteLine("資源清理完成");
        }

        // 計算屬性
        public decimal Q1FinalTarget => Q1Target;
        public decimal Q2FinalTarget => Q2Target + Q1Carried;  // 改用 Q1Carried 而不是 Q2Added
        public decimal Q3FinalTarget => Q3Target + Q2Carried;  // 改用 Q2Carried 而不是 Q3Added
        public decimal Q4FinalTarget => Q4Target + Q3Carried;  // 改用 Q3Carried 而不是 Q4Added

        // Achievement 計算保持不變
        public decimal Q1Achievement => Q1FinalTarget > 0 ? Math.Round((Q1Achieved / Q1FinalTarget) * 100, 1) : 0;
        public decimal Q2Achievement => Q2FinalTarget > 0 ? Math.Round((Q2Achieved / Q2FinalTarget) * 100, 1) : 0;
        public decimal Q3Achievement => Q3FinalTarget > 0 ? Math.Round((Q3Achieved / Q3FinalTarget) * 100, 1) : 0;
        public decimal Q4Achievement => Q4FinalTarget > 0 ? Math.Round((Q4Achieved / Q4FinalTarget) * 100, 1) : 0;

        // Carried 計算保持不變
        public decimal Q1Carried => Math.Max(0, Q1FinalTarget - Q1Achieved);
        public decimal Q2Carried => Math.Max(0, Q2FinalTarget - Q2Achieved);
        public decimal Q3Carried => Math.Max(0, Q3FinalTarget - Q3Achieved);

        // Exceeded 計算保持不變
        public decimal Q2Exceeded => Math.Max(0, Q2Achieved - Q2FinalTarget);
        public decimal Q3Exceeded => Math.Max(0, Q3Achieved - Q3FinalTarget);
        public decimal Q4Exceeded => Math.Max(0, Q4Achieved - Q4FinalTarget);
    }

    public class NotificationItem
    {
        public string Message { get; set; }
    }


}