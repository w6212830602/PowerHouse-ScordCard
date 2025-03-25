using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using System.Runtime.CompilerServices;
using CommunityToolkit.Mvvm.Input;
using ScoreCard.Models;
using ScoreCard.Services;
using System.Linq;
using System.Diagnostics;
using System.Text.Json;

namespace ScoreCard.ViewModels
{
    public partial class SalesAnalysisViewModel : INotifyPropertyChanged
    {
        private readonly IExcelService _excelService;
        private readonly ITargetService _targetService; // 新增目標服務
        private List<SalesData> _allSalesData;

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        #endregion

        #region 屬性

        // 基本屬性
        private DateTime _startDate;
        public DateTime StartDate
        {
            get => _startDate;
            set
            {
                if (SetProperty(ref _startDate, value))
                {
                    Debug.WriteLine($"ViewModel - StartDate changed to: {_startDate:yyyy/MM/dd}");
                    OnDateRangeChangedAsync().ConfigureAwait(false);
                }
            }
        }

        private DateTime _endDate;
        public DateTime EndDate
        {
            get => _endDate;
            set
            {
                if (SetProperty(ref _endDate, value))
                {
                    Debug.WriteLine($"ViewModel - EndDate changed to: {_endDate:yyyy/MM/dd}");
                    OnDateRangeChangedAsync().ConfigureAwait(false);
                }
            }
        }

        private bool _isSummaryView = true;
        public bool IsSummaryView
        {
            get => _isSummaryView;
            set => SetProperty(ref _isSummaryView, value);
        }

        private bool _isLoading;
        public bool IsLoading
        {
            get => _isLoading;
            set => SetProperty(ref _isLoading, value);
        }

        private SalesAnalysisSummary _summary;
        public SalesAnalysisSummary Summary
        {
            get => _summary;
            set => SetProperty(ref _summary, value);
        }

        // 新增使用目標值的屬性
        private decimal _targetValue; // 當前選擇的目標值
        public decimal TargetValue
        {
            get => _targetValue;
            set => SetProperty(ref _targetValue, value);
        }

        // 圖表相關屬性
        private ObservableCollection<Models.ChartData> _targetVsAchievementData = new();
        public ObservableCollection<Models.ChartData> TargetVsAchievementData
        {
            get => _targetVsAchievementData;
            set => SetProperty(ref _targetVsAchievementData, value);
        }

        private ObservableCollection<Models.ChartData> _achievementTrendData = new();
        public ObservableCollection<Models.ChartData> AchievementTrendData
        {
            get => _achievementTrendData;
            set => SetProperty(ref _achievementTrendData, value);
        }

        private ObservableCollection<SalesRepPerformance> _leaderboard = new();
        public ObservableCollection<SalesRepPerformance> Leaderboard
        {
            get => _leaderboard;
            set => SetProperty(ref _leaderboard, value);
        }

        private double _yAxisMaximum = 20;
        public double YAxisMaximum
        {
            get => _yAxisMaximum;
            set => SetProperty(ref _yAxisMaximum, value);
        }

        // Leaderboard 頁籤切換相關屬性
        private string _viewType = "ByProduct";
        public string ViewType
        {
            get => _viewType;
            set
            {
                if (SetProperty(ref _viewType, value))
                {
                    OnPropertyChanged(nameof(IsProductView));
                    OnPropertyChanged(nameof(IsRepView));
                    OnPropertyChanged(nameof(IsDeptLobView));
                    LoadLeaderboardDataAsync().ConfigureAwait(false);
                }
            }
        }

        private bool _isBookedStatus = true;
        public bool IsBookedStatus
        {
            get => _isBookedStatus;
            set
            {
                if (SetProperty(ref _isBookedStatus, value))
                {
                    LoadLeaderboardDataAsync().ConfigureAwait(false);
                }
            }
        }

        // 各視圖數據集合
        private ObservableCollection<SalesLeaderboardItem> _salesLeaderboard = new();
        public ObservableCollection<SalesLeaderboardItem> SalesLeaderboard
        {
            get => _salesLeaderboard;
            set => SetProperty(ref _salesLeaderboard, value);
        }

        private ObservableCollection<ProductSalesData> _productSalesData = new();
        public ObservableCollection<ProductSalesData> ProductSalesData
        {
            get => _productSalesData;
            set => SetProperty(ref _productSalesData, value);
        }

        private ObservableCollection<DepartmentLobData> _departmentLobData = new();
        public ObservableCollection<DepartmentLobData> DepartmentLobData
        {
            get => _departmentLobData;
            set => SetProperty(ref _departmentLobData, value);
        }

        // 視圖可見性計算屬性
        public bool IsProductView
        {
            get => ViewType == "ByProduct";
            set
            {
                if (value && ViewType != "ByProduct")
                {
                    ViewType = "ByProduct";
                }
            }
        }

        public bool IsRepView
        {
            get => ViewType == "ByRep";
            set
            {
                if (value && ViewType != "ByRep")
                {
                    ViewType = "ByRep";
                }
            }
        }

        public bool IsDeptLobView
        {
            get => ViewType == "ByDeptLOB";
            set
            {
                if (value && ViewType != "ByDeptLOB")
                {
                    ViewType = "ByDeptLOB";
                }
            }
        }
        #endregion

        #region 命令

        public ICommand SwitchViewCommand { get; }
        public ICommand ChangeViewTypeCommand { get; }
        public ICommand ChangeStatusCommand { get; }
        public IAsyncRelayCommand LoadDataCommand { get; }

        #endregion

        public SalesAnalysisViewModel(IExcelService excelService, ITargetService targetService)
        {
            _excelService = excelService;
            _targetService = targetService; // 初始化目標服務
            _summary = new SalesAnalysisSummary
            {
                TotalTarget = 0,
                TotalAchievement = 0,
                TotalMargin = 0
            };

            // 初始化命令
            LoadDataCommand = new AsyncRelayCommand(LoadDataAsync);
            SwitchViewCommand = new RelayCommand<string>(ExecuteSwitchView);
            ChangeViewTypeCommand = new RelayCommand<string>(ExecuteChangeViewType);
            ChangeStatusCommand = new RelayCommand<string>(ExecuteChangeStatus);

            // 訂閱目標更新事件
            _targetService.TargetsUpdated += OnTargetsUpdated;

            // 初始化
            InitializeAsync();
        }

        // 處理目標更新事件
        private void OnTargetsUpdated(object sender, EventArgs e)
        {
            Debug.WriteLine("收到目標更新通知");
            MainThread.InvokeOnMainThreadAsync(async () =>
            {
                await ReloadDataAsync();
                Debug.WriteLine("目標更新後，數據已重新載入");
            });
        }

        // 切換 Summary/Detailed 視圖
        private void ExecuteSwitchView(string viewType)
        {
            if (!string.IsNullOrEmpty(viewType))
            {
                if (viewType.ToLower() == "summary")
                {
                    IsSummaryView = true;
                }
                else if (viewType.ToLower() == "detailed")
                {
                    // 如果切換到詳細視圖，則導航到詳細頁面
                    NavigateToDetailedViewCommand.Execute(null);
                    return;
                }
            }
        }

        // 切換 Leaderboard 視圖類型 (Product/Rep/Dept-LOB)
        private void ExecuteChangeViewType(string viewTypeStr)
        {
            try
            {
                if (!string.IsNullOrEmpty(viewTypeStr))
                {
                    // 記錄原始視圖類型
                    string originalViewType = ViewType;
                    Debug.WriteLine($"視圖切換請求: 從 {originalViewType} 到 {viewTypeStr}");

                    // 如果視圖類型沒變，重新加載當前視圖數據而不改變視圖類型
                    if (originalViewType == viewTypeStr)
                    {
                        Debug.WriteLine("重新加載當前視圖數據，不改變視圖類型");

                        // 由於我們沒有改變視圖類型，不會自動觸發數據重載
                        // 因此需要手動觸發數據重載
                        MainThread.BeginInvokeOnMainThread(async () => {
                            try
                            {
                                // 確保在加載前，視圖類型已經正確設置
                                await LoadLeaderboardDataAsync();
                                // 明確觸發 UI 更新
                                OnPropertyChanged(nameof(ProductSalesData));
                                OnPropertyChanged(nameof(SalesLeaderboard));
                                OnPropertyChanged(nameof(DepartmentLobData));
                                Debug.WriteLine($"已重新加載視圖 {viewTypeStr} 的數據");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"重新加載數據錯誤: {ex.Message}");
                            }
                        });
                        return;
                    }

                    // 設置新的視圖類型
                    ViewType = viewTypeStr;

                    // 明確觸發所有相關的屬性更新
                    OnPropertyChanged(nameof(IsProductView));
                    OnPropertyChanged(nameof(IsRepView));
                    OnPropertyChanged(nameof(IsDeptLobView));

                    // 確保立即加載對應的數據
                    MainThread.BeginInvokeOnMainThread(async () => {
                        try
                        {
                            await LoadLeaderboardDataAsync();
                            Debug.WriteLine($"視圖已成功切換為: {viewTypeStr}，IsProductView={IsProductView}, IsRepView={IsRepView}, IsDeptLobView={IsDeptLobView}");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"視圖切換後加載數據錯誤: {ex.Message}");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"視圖切換錯誤: {ex.Message}");
            }
        }

        // 切換狀態過濾器 (Booked/Completed)
        private void ExecuteChangeStatus(string status)
        {
            IsBookedStatus = status.ToLower() == "booked";
            Debug.WriteLine($"狀態過濾器切換為: {status}");
        }

        // 初始化
        private async void InitializeAsync()
        {
            try
            {
                IsLoading = true;

                // 初始化目標服務
                await _targetService.InitializeAsync();

                // 初始載入數據
                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                _allSalesData = data;

                // 設定初始日期範圍
                if (StartDate == default || EndDate == default)
                {
                    var dates = _allSalesData.Select(x => x.ReceivedDate).OrderBy(x => x).ToList();
                    if (dates.Any())
                    {
                        // 直接設置欄位避免觸發事件
                        _endDate = dates.Last().Date;
                        _startDate = dates.First().Date;
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

                // 載入現在財政年度的目標
                var currentDate = DateTime.Now;
                var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;
                var companyTarget = _targetService.GetCompanyTarget(currentFiscalYear);
                if (companyTarget != null)
                {
                    TargetValue = companyTarget.AnnualTarget;
                    Debug.WriteLine($"已載入 FY{currentFiscalYear} 目標值: ${TargetValue:N0}");
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

                // 確保圖表數據集合不為空
                if (TargetVsAchievementData == null)
                    TargetVsAchievementData = new ObservableCollection<Models.ChartData>();

                if (AchievementTrendData == null)
                    AchievementTrendData = new ObservableCollection<Models.ChartData>();

                // 確保已有數據載入
                if (_allSalesData == null || !_allSalesData.Any())
                {
                    try
                    {
                        var (data, lastUpdated) = await _excelService.LoadDataAsync();
                        _allSalesData = data ?? new List<SalesData>();
                        Debug.WriteLine($"Loaded {_allSalesData.Count} records from Excel");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error loading Excel data: {ex.Message}");
                        _allSalesData = CreateSampleData(); // 確保始終有數據可用
                    }
                }

                // 從目標服務獲取當前財年的目標
                var currentDate = DateTime.Now;
                var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;
                var companyTarget = _targetService.GetCompanyTarget(currentFiscalYear);
                if (companyTarget != null)
                {
                    TargetValue = companyTarget.AnnualTarget;
                    Debug.WriteLine($"載入 FY{currentFiscalYear} 目標值: ${TargetValue:N0}");
                }

                // 強制使用當前的 StartDate 和 EndDate
                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1); // 包含結束日期的整天

                Debug.WriteLine($"實際過濾使用的日期範圍: {startDate:yyyy-MM-dd} 到 {endDate:yyyy-MM-dd}");

                // 過濾數據
                var filteredData = new List<SalesData>();
                try
                {
                    filteredData = _allSalesData
                        .Where(x => x.ReceivedDate.Date >= startDate && x.ReceivedDate.Date <= endDate.Date)
                        .ToList();

                    Debug.WriteLine($"過濾後數據: {filteredData.Count} 條記錄在 {startDate:yyyy-MM-dd} 和 {endDate:yyyy-MM-dd} 之間");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error filtering data: {ex.Message}");
                    filteredData = CreateSampleData(); // 確保始終有過濾後的數據
                }

                // 如果過濾後沒有數據，添加默認樣本數據
                if (filteredData == null || !filteredData.Any())
                {
                    Debug.WriteLine("No data after filtering, using sample data");
                    filteredData = CreateSampleData();
                }

                // 按月份分組並排序
                var monthlyData = new List<Models.ChartData>();
                try
                {
                    monthlyData = filteredData
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
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error generating monthly data: {ex.Message}");
                    // 創建默認月度數據
                    monthlyData = new List<Models.ChartData>
                    {
                        new Models.ChartData { Label = DateTime.Now.ToString("yyyy/MM"), Target = 1, Achievement = 0.5m }
                    };
                }

                // 如果沒有月度數據，添加默認數據點
                if (monthlyData == null || !monthlyData.Any())
                {
                    monthlyData = new List<Models.ChartData>
                    {
                        new Models.ChartData { Label = "No Data", Target = 0, Achievement = 0 }
                    };
                }

                // 在主線程上進行 UI 更新
                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    try
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

                        // 確保圖表有數據
                        if (TargetVsAchievementData.Count == 0)
                        {
                            TargetVsAchievementData.Add(new Models.ChartData { Label = "No Data", Target = 0, Achievement = 0 });
                        }

                        if (AchievementTrendData.Count == 0)
                        {
                            AchievementTrendData.Add(new Models.ChartData { Label = "No Data", Target = 0, Achievement = 0 });
                        }

                        // 更新匯總數據
                        var totalTarget = Math.Round(filteredData.Sum(x => x.POValue) / 1000000m, 2);
                        var totalAchievement = Math.Round(filteredData.Sum(x => x.VertivValue) / 1000000m, 2);
                        var totalMargin = Math.Round(filteredData.Sum(x => x.TotalCommission) / 1000000m, 2);

                        // 防止零值或負值
                        totalTarget = Math.Max(0.01m, totalTarget);
                        totalAchievement = Math.Max(0.01m, totalAchievement);
                        totalMargin = Math.Max(0.01m, totalMargin);

                        Summary = new SalesAnalysisSummary
                        {
                            TotalTarget = totalTarget,
                            TotalAchievement = totalAchievement,
                            TotalMargin = totalMargin
                        };

                        Debug.WriteLine($"Updated summary: Target=${Summary.TotalTarget}M, Achievement=${Summary.TotalAchievement}M, Margin=${Summary.TotalMargin}M");

                        // 更新排行榜數據
                        LoadLeaderboardData(filteredData);

                        // 確保更新圖表軸
                        UpdateChartAxes();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error updating UI: {ex.Message}");

                        // 確保即使出錯也有默認數據
                        if (TargetVsAchievementData.Count == 0)
                        {
                            TargetVsAchievementData.Add(new Models.ChartData { Label = "Error", Target = 1, Achievement = 0 });
                        }

                        if (AchievementTrendData.Count == 0)
                        {
                            AchievementTrendData.Add(new Models.ChartData { Label = "Error", Target = 1, Achievement = 0 });
                        }

                        Summary = new SalesAnalysisSummary
                        {
                            TotalTarget = 1,
                            TotalAchievement = 0.5m,
                            TotalMargin = 0.1m
                        };
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in LoadDataAsync: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");

                try
                {
                    await MainThread.InvokeOnMainThreadAsync(() => {
                        // 確保圖表數據不為空
                        if (TargetVsAchievementData == null || TargetVsAchievementData.Count == 0)
                        {
                            TargetVsAchievementData = new ObservableCollection<Models.ChartData> {
                                new Models.ChartData { Label = "Error", Target = 1, Achievement = 0 }
                            };
                        }

                        if (AchievementTrendData == null || AchievementTrendData.Count == 0)
                        {
                            AchievementTrendData = new ObservableCollection<Models.ChartData> {
                                new Models.ChartData { Label = "Error", Target = 1, Achievement = 0 }
                            };
                        }

                        // 確保有有效的 Y 軸最大值
                        YAxisMaximum = 5;

                        // 確保 Summary 不為空
                        if (Summary == null)
                        {
                            Summary = new SalesAnalysisSummary
                            {
                                TotalTarget = 1,
                                TotalAchievement = 0.5m,
                                TotalMargin = 0.1m
                            };
                        }
                    });
                }
                catch (Exception innerEx)
                {
                    Debug.WriteLine($"Fatal error in error handling: {innerEx.Message}");
                }

                // 顯示錯誤，但不要讓應用程式崩潰
                try
                {
                    await Application.Current.MainPage.DisplayAlert("資料載入錯誤", "無法載入銷售數據，請稍後再試", "確定");
                }
                catch
                {
                    // 最後的防線 - 即使顯示錯誤對話框失敗也不崩潰
                }
            }
            finally
            {
                IsLoading = false;
            }
        }

        // 載入 Leaderboard 數據 (不更新其他頁面數據)
        private async Task LoadLeaderboardDataAsync()
        {
            try
            {
                Debug.WriteLine($"LoadLeaderboardDataAsync - 當前視圖: {ViewType}");
                IsLoading = true;

                if (_allSalesData == null || !_allSalesData.Any())
                {
                    var (data, lastUpdated) = await _excelService.LoadDataAsync();
                    _allSalesData = data;
                    Debug.WriteLine($"已載入 {_allSalesData?.Count ?? 0} 條記錄");
                }

                var startDate = StartDate.Date;
                var endDate = EndDate.Date.AddDays(1).AddSeconds(-1);

                // 僅按指定的日期範圍過濾數據，不再擴大範圍
                var filteredData = _allSalesData?
                    .Where(x => x.ReceivedDate.Date >= startDate && x.ReceivedDate.Date <= endDate.Date)
                    .ToList() ?? new List<SalesData>();

                Debug.WriteLine($"日期過濾後有 {filteredData.Count} 條記錄");

                // 清除現有緩存，確保數據重新生成
                _excelService.ClearCache();

                await MainThread.InvokeOnMainThreadAsync(() => {
                    LoadLeaderboardData(filteredData);
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadLeaderboardDataAsync 錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 發生錯誤時，不使用示例數據，而是顯示空白數據
                await MainThread.InvokeOnMainThreadAsync(() => {
                    LoadLeaderboardData(new List<SalesData>());
                });
            }
            finally
            {
                IsLoading = false;
            }
        }

        private List<SalesData> CreateSampleData()
        {
            // 創建一些樣本數據，確保在沒有實際數據時界面也能顯示內容
            return new List<SalesData>
            {
                new SalesData
                {
                    ReceivedDate = DateTime.Now,
                    ProductType = "Thermal",
                    POValue = 8058197.44m,
                    VertivValue = 8058197.44m,
                    TotalCommission = 899545.64m,
                    SalesRep = "Sample Rep 1",
                    Status = "Booked"
                },
                new SalesData
                {
                    ReceivedDate = DateTime.Now,
                    ProductType = "Power",
                    POValue = 5857870.43m,
                    VertivValue = 5857870.43m,
                    TotalCommission = 305481.01m,
                    SalesRep = "Sample Rep 2",
                    Status = "Booked"
                },
                new SalesData
                {
                    ReceivedDate = DateTime.Now,
                    ProductType = "Batts & Caps",
                    POValue = 2169156.88m,
                    VertivValue = 2169156.88m,
                    TotalCommission = 344608.00m,
                    SalesRep = "Sample Rep 3",
                    Status = "Booked"
                },
                new SalesData
                {
                    ReceivedDate = DateTime.Now,
                    ProductType = "Service",
                    POValue = 2140938.32m,
                    VertivValue = 2140938.32m,
                    TotalCommission = 160595.74m,
                    SalesRep = "Sample Rep 4",
                    Status = "Booked"
                },
                new SalesData
                {
                    ReceivedDate = DateTime.Now,
                    ProductType = "Channel",
                    POValue = 1863238.05m,
                    VertivValue = 1863238.05m,
                    TotalCommission = 214071.40m,
                    SalesRep = "Sample Rep 5",
                    Status = "Booked"
                }
            };
        }

        // 載入 Leaderboard 數據的核心邏輯
        private void LoadLeaderboardData(List<SalesData> filteredData)
        {
            try
            {
                Debug.WriteLine($"LoadLeaderboardData - 當前視圖: {ViewType}, 數據量: {filteredData?.Count ?? 0}");

                // 先根據 IsBookedStatus 過濾數據
                var statusFilteredData = filteredData.Where(x =>
                    (IsBookedStatus && x.Status == "Booked") ||
                    (!IsBookedStatus && x.Status == "Completed")
                ).ToList();

                Debug.WriteLine($"狀態過濾後剩餘數據量: {statusFilteredData.Count}");

                // 根據不同的視圖類型加載對應的數據
                switch (ViewType)
                {
                    case "ByProduct":
                        LoadProductData(statusFilteredData);
                        OnPropertyChanged(nameof(ProductSalesData));
                        break;

                    case "ByRep":
                        LoadSalesRepData(statusFilteredData);
                        OnPropertyChanged(nameof(SalesLeaderboard));
                        break;

                    case "ByDeptLOB":
                        // 重要：設置過濾後的數據到ExcelService
                        if (_excelService is ExcelService excelSvc)
                        {
                            excelSvc.SetFilteredData(statusFilteredData);
                        }

                        // 加載部門/LOB數據
                        DepartmentLobData = new ObservableCollection<DepartmentLobData>(_excelService.GetDepartmentLobData());
                        Debug.WriteLine($"已加載 {DepartmentLobData.Count} 條部門/LOB數據");
                        OnPropertyChanged(nameof(DepartmentLobData));
                        break;

                    default:
                        Debug.WriteLine($"未知的視圖類型: {ViewType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadLeaderboardData 錯誤: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
            }
        }

        // 載入產品數據
        private void LoadProductData(List<SalesData> filteredData)
        {
            try
            {
                Debug.WriteLine("載入產品視圖數據");

                if (filteredData == null || !filteredData.Any())
                {
                    Debug.WriteLine("沒有過濾後的數據可用");
                    ProductSalesData = new ObservableCollection<ProductSalesData>();
                    return;
                }

                // 按產品類型（AD列）分組，並計算每個類型的PO Value（G列）總和
                var productData = filteredData
                    .GroupBy(x => x.ProductType)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key))
                    .Select(g => new ProductSalesData
                    {
                        ProductType = g.Key,
                        // 直接從Excel的相應列獲取數值
                        AgencyMargin = Math.Round(g.Sum(x => x.AgencyMargin), 2),
                        BuyResellMargin = Math.Round(g.Sum(x => x.BuyResellValue), 2),
                        POValue = Math.Round(g.Sum(x => x.POValue), 2),
                        // TotalMargin是兩個Margin的總和
                        TotalMargin = Math.Round(g.Sum(x => x.AgencyMargin) + g.Sum(x => x.BuyResellValue), 2)
                    })
                    .OrderByDescending(x => x.POValue)
                    .ToList();

                // 計算PO Value百分比
                decimal totalPOValue = productData.Sum(p => p.POValue);
                foreach (var product in productData)
                {
                    product.PercentageOfTotal = totalPOValue > 0
                        ? Math.Round((product.POValue / totalPOValue) * 100, 2)
                        : 0;
                }

                ProductSalesData = new ObservableCollection<ProductSalesData>(productData);
                Debug.WriteLine($"已載入 {productData.Count} 條產品數據");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入產品數據時發生錯誤: {ex.Message}");

                // 發生錯誤
                Debug.WriteLine($"")