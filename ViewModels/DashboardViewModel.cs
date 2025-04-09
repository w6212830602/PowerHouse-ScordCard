using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ScoreCard.Models;
using ScoreCard.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace ScoreCard.ViewModels
{
    public partial class DashboardViewModel : ObservableObject
    {
        // Dependencies
        private readonly IExcelService _excelService;
        private readonly ITargetService _targetService;
        private CancellationTokenSource _cancellationTokenSource;

        // Basic properties
        [ObservableProperty]
        private string _selectedOption;

        [ObservableProperty]
        private List<string> _options;

        [ObservableProperty]
        private string _lastUpdated = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

        // Edit mode
        [ObservableProperty]
        private bool _isEditMode;

        // Target editing tracking
        private decimal _originalAnnualTarget;
        private decimal _originalQ1Target;
        private decimal _originalQ2Target;
        private decimal _originalQ3Target;
        private decimal _originalQ4Target;

        private bool _isQ1Edited;
        private bool _isQ2Edited;
        private bool _isQ3Edited;
        private bool _isQ4Edited;
        private bool _isAnnualEdited;

        // Main target properties with full property implementations
        private decimal _annualTarget;
        public decimal AnnualTarget
        {
            get => _annualTarget;
            set
            {
                if (SetProperty(ref _annualTarget, value))
                {
                    // Only handle distribution if in edit mode and value actually changed
                    if (IsEditMode && _annualTarget != _originalAnnualTarget)
                    {
                        _isAnnualEdited = true;
                        DistributeAnnualTarget();
                    }

                    // Update display properties
                    OnPropertyChanged(nameof(AnnualTargetDisplay));
                    UpdateRemainingValues();
                }
            }
        }

        private decimal _q1Target;
        public decimal Q1Target
        {
            get => _q1Target;
            set
            {
                if (SetProperty(ref _q1Target, value))
                {
                    if (IsEditMode && _q1Target != _originalQ1Target)
                    {
                        _isQ1Edited = true;
                        RecalculateAnnualTarget();
                        RedistributeRemainingAmount();
                    }

                    OnPropertyChanged(nameof(Q1TargetDisplay));
                    OnPropertyChanged(nameof(Q1FinalTarget));
                    OnPropertyChanged(nameof(Q1Achievement));
                    OnPropertyChanged(nameof(Q1AchievementDisplay));
                    UpdateRemainingValues();
                }
            }
        }

        private decimal _q2Target;
        public decimal Q2Target
        {
            get => _q2Target;
            set
            {
                if (SetProperty(ref _q2Target, value))
                {
                    if (IsEditMode && _q2Target != _originalQ2Target)
                    {
                        _isQ2Edited = true;
                        RecalculateAnnualTarget();
                        RedistributeRemainingAmount();
                    }

                    OnPropertyChanged(nameof(Q2TargetDisplay));
                    OnPropertyChanged(nameof(Q2FinalTarget));
                    OnPropertyChanged(nameof(Q2Achievement));
                    OnPropertyChanged(nameof(Q2AchievementDisplay));
                    UpdateRemainingValues();
                }
            }
        }

        private decimal _q3Target;
        public decimal Q3Target
        {
            get => _q3Target;
            set
            {
                if (SetProperty(ref _q3Target, value))
                {
                    if (IsEditMode && _q3Target != _originalQ3Target)
                    {
                        _isQ3Edited = true;
                        RecalculateAnnualTarget();
                        RedistributeRemainingAmount();
                    }

                    OnPropertyChanged(nameof(Q3TargetDisplay));
                    OnPropertyChanged(nameof(Q3FinalTarget));
                    OnPropertyChanged(nameof(Q3Achievement));
                    OnPropertyChanged(nameof(Q3AchievementDisplay));
                    UpdateRemainingValues();
                }
            }
        }

        private decimal _q4Target;
        public decimal Q4Target
        {
            get => _q4Target;
            set
            {
                if (SetProperty(ref _q4Target, value))
                {
                    if (IsEditMode && _q4Target != _originalQ4Target)
                    {
                        _isQ4Edited = true;
                        RecalculateAnnualTarget();
                        RedistributeRemainingAmount();
                    }

                    OnPropertyChanged(nameof(Q4TargetDisplay));
                    OnPropertyChanged(nameof(Q4FinalTarget));
                    OnPropertyChanged(nameof(Q4Achievement));
                    OnPropertyChanged(nameof(Q4AchievementDisplay));
                    UpdateRemainingValues();
                }
            }
        }

        public string TestQ1Display => "Test Q1: +$295,671";

        // Achievement values
        [ObservableProperty]
        private decimal _q1Achieved;

        [ObservableProperty]
        private decimal _q2Achieved;

        [ObservableProperty]
        private decimal _q3Achieved;

        [ObservableProperty]
        private decimal _q4Achieved;

        [ObservableProperty]
        private decimal _achievement;

        [ObservableProperty]
        private decimal _remaining;

        [ObservableProperty]
        private double _achievementProgress;

        [ObservableProperty]
        private decimal _notInvoiced;

        [ObservableProperty]
        private decimal _inProgress;

        // Notifications
        [ObservableProperty]
        private ObservableCollection<NotificationItem> _notifications = new();

        [ObservableProperty]
        private bool _isLoading;

        // Status message
        [ObservableProperty]
        private string _statusMessage;

        [ObservableProperty]
        private bool _isStatusMessageVisible;

        [ObservableProperty]
        private bool _isStatusSuccess;

        // Format display properties
        public string AnnualTargetDisplay => $"${AnnualTarget:N0}";
        public string AchievementDisplay => $"Achievement:{Achievement:0.0}%";
        public string RemainingDisplay => $"${Remaining:N0}";
        public string TotalAchievedDisplay => $"${Q1Achieved + Q2Achieved + Q3Achieved + Q4Achieved:N0}";
        public string NotInvoicedDisplay => $"${NotInvoiced:N0} Booked";
        public string InProgressDisplay => $"${InProgress:N0} In progress";

        // Quarterly achievement percentages
        public double Q1Achievement => Q1FinalTarget > 0 ? (double)Math.Round((Q1Achieved / Q1FinalTarget) * 100, 1) : 0;
        public double Q2Achievement => Q2FinalTarget > 0 ? (double)Math.Round((Q2Achieved / Q2FinalTarget) * 100, 1) : 0;
        public double Q3Achievement => Q3FinalTarget > 0 ? (double)Math.Round((Q3Achieved / Q3FinalTarget) * 100, 1) : 0;
        public double Q4Achievement => Q4FinalTarget > 0 ? (double)Math.Round((Q4Achieved / Q4FinalTarget) * 100, 1) : 0;

        // Quarterly display properties
        public string Q1AchievementDisplay => $"{Q1Achievement:0.0}%";
        public string Q2AchievementDisplay => $"{Q2Achievement:0.0}%";
        public string Q3AchievementDisplay => $"{Q3Achievement:0.0}%";
        public string Q4AchievementDisplay => $"{Q4Achievement:0.0}%";

        public string Q1TargetDisplay => $"${Q1FinalTarget:N0}";
        public string Q2TargetDisplay => $"${Q2FinalTarget:N0}";
        public string Q3TargetDisplay => $"${Q3FinalTarget:N0}";
        public string Q4TargetDisplay => $"${Q4FinalTarget:N0}";

        public string Q1BaseDisplay => $"Base: ${Q1Target:N0}";
        public string Q2BaseDisplay => $"Base: ${Q2Target:N0}";
        public string Q3BaseDisplay => $"Base: ${Q3Target:N0}";
        public string Q4BaseDisplay => $"Base: ${Q4Target:N0}";

        public string Q1AchievedDisplay => $"${Q1Achieved:N0}";
        public string Q2AchievedDisplay => $"${Q2Achieved:N0}";
        public string Q3AchievedDisplay => $"${Q3Achieved:N0}";
        public string Q4AchievedDisplay => $"${Q4Achieved:N0}";

        // Carried/exceeded calculations and display
        public decimal Q1Carried => Math.Max(0, Q1FinalTarget - Q1Achieved);
        public decimal Q2Carried => Math.Max(0, Q2FinalTarget - Q2Achieved);
        public decimal Q3Carried => Math.Max(0, Q3FinalTarget - Q3Achieved);

        public decimal Q1Exceeded => Math.Max(0, Q1Achieved - Q1FinalTarget);
        public decimal Q2Exceeded => Math.Max(0, Q2Achieved - Q2FinalTarget);
        public decimal Q3Exceeded => Math.Max(0, Q3Achieved - Q3FinalTarget);
        public decimal Q4Exceeded => Math.Max(0, Q4Achieved - Q4FinalTarget);

        public string Q1CarriedDisplay => $"Carried: ${Q1Carried:N0} →";
        public string Q2CarriedDisplay => $"Carried: ${Q2Carried:N0} →";
        public string Q3CarriedDisplay => $"Carried: ${Q3Carried:N0} →";

        public string Q1CarriedAddedDisplay => $"+${Q1Carried:N0}";
        public string Q2CarriedAddedDisplay => $"+${Q2Carried:N0}";
        public string Q3CarriedAddedDisplay => $"+${Q3Carried:N0}";

        public string Q1ExceededDisplay => $"Exceeded: +${Q1Exceeded:N0}";
        public string Q2ExceededDisplay => $"Exceeded: +${Q2Exceeded:N0}";
        public string Q3ExceededDisplay => $"Exceeded: +${Q3Exceeded:N0}";
        public string Q4ExceededDisplay => $"Exceeded: +${Q4Exceeded:N0}";

        // Final target calculations including carried amounts
        public decimal Q1FinalTarget => Q1Target;
        public decimal Q2FinalTarget => Q2Target + Q1Carried;
        public decimal Q3FinalTarget => Q3Target + Q2Carried;
        public decimal Q4FinalTarget => Q4Target + Q3Carried;

        public DashboardViewModel(IExcelService excelService, ITargetService targetService)
        {
            Debug.WriteLine("DashboardViewModel 建構函數開始初始化");
            _excelService = excelService;
            _targetService = targetService;
            _excelService.DataUpdated += OnDataUpdated;
            _targetService.TargetsUpdated += OnTargetsUpdated;
            _cancellationTokenSource = new CancellationTokenSource();

            // Get current fiscal year
            var currentDate = DateTime.Now;
            var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;

            Debug.WriteLine($"當前日期: {currentDate}, 當前財年: {currentFiscalYear}");

            // Set fiscal year options (descending order)
            Options = new List<string> {
                $"FY{currentFiscalYear + 1}",
                $"FY{currentFiscalYear}",
                $"FY{currentFiscalYear - 1}"
            };

            // Set default option to current fiscal year
            SelectedOption = $"FY{currentFiscalYear}";
            
            InitializeNotifications();
            InitializeAsync();
            Debug.WriteLine($"DashboardViewModel 初始化完成，選擇的財年: {SelectedOption}");

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

            if (int.TryParse(SelectedOption.Replace("FY", ""), out int result))
            {
                return result;
            }

            // Default to current fiscal year
            var date = DateTime.Now;
            return date.Month >= 8 ? date.Year + 1 : date.Year;
        }

        private async void InitializeAsync()
        {
            try
            {
                Debug.WriteLine("開始初始化 Dashboard");

                // Initialize target service
                await _targetService.InitializeAsync();

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
                Debug.WriteLine("開始載入數據");
                IsLoading = true;

                // Get selected fiscal year and update targets
                var selectedFiscalYear = GetSelectedFiscalYear();
                await UpdateTargets(selectedFiscalYear);

                var (data, lastUpdated) = await _excelService.LoadDataAsync();
                Debug.WriteLine($"成功載入數據，共 {data.Count} 筆記錄");

                await MainThread.InvokeOnMainThreadAsync(() =>
                {
                    // Update dashboard based on new data
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

                Debug.WriteLine($"Q1Achieved (type: {Q1Achieved.GetType()}): {Q1Achieved}");
                Debug.WriteLine($"Q1FinalTarget (type: {Q1FinalTarget.GetType()}): {Q1FinalTarget}");
                Debug.WriteLine($"Q1Achieved - Q1FinalTarget = {Q1Achieved - Q1FinalTarget}");
                Debug.WriteLine($"Q1Exceeded: {Q1Exceeded}");
                Debug.WriteLine($"{Q1ExceededDisplay}");



            }
        }

        private async Task UpdateTargets(int fiscalYear)
        {
            try
            {
                // Exit edit mode if active
                if (IsEditMode)
                {
                    await SaveChanges();
                }

                // Reset editing flags
                ResetEditFlags();

                // Get company target for the selected fiscal year
                var companyTarget = _targetService.GetCompanyTarget(fiscalYear);

                if (companyTarget != null)
                {
                    // Store original values for tracking changes
                    _originalAnnualTarget = companyTarget.AnnualTarget;
                    _originalQ1Target = companyTarget.Q1Target;
                    _originalQ2Target = companyTarget.Q2Target;
                    _originalQ3Target = companyTarget.Q3Target;
                    _originalQ4Target = companyTarget.Q4Target;

                    // Set current values
                    AnnualTarget = companyTarget.AnnualTarget;
                    Q1Target = companyTarget.Q1Target;
                    Q2Target = companyTarget.Q2Target;
                    Q3Target = companyTarget.Q3Target;
                    Q4Target = companyTarget.Q4Target;

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
                    // Set default values to avoid calculation issues
                    AnnualTarget = 4335000;
                    Q1Target = 1083750;
                    Q2Target = 1083750;
                    Q3Target = 1083750;
                    Q4Target = 1083750;

                    // Store as originals
                    _originalAnnualTarget = AnnualTarget;
                    _originalQ1Target = Q1Target;
                    _originalQ2Target = Q2Target;
                    _originalQ3Target = Q3Target;
                    _originalQ4Target = Q4Target;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新目標時發生錯誤: {ex.Message}");
                // Set default values to avoid calculation issues
                AnnualTarget = 4335000;
                Q1Target = 1083750;
                Q2Target = 1083750;
                Q3Target = 1083750;
                Q4Target = 1083750;

                // Store as originals
                _originalAnnualTarget = AnnualTarget;
                _originalQ1Target = Q1Target;
                _originalQ2Target = Q2Target;
                _originalQ3Target = Q3Target;
                _originalQ4Target = Q4Target;
            }
        }

        private void UpdateDashboard(List<SalesData> data)
        {
            try
            {
                Debug.WriteLine("開始更新儀表板數據");

                // Get selected fiscal year 
                var selectedFiscalYear = GetSelectedFiscalYear();

                // Filter data for selected fiscal year
                var yearData = data.Where(x => x.FiscalYear == selectedFiscalYear).ToList();
                Debug.WriteLine($"目前財年: FY{selectedFiscalYear}, 數據筆數: {yearData.Count}");

                // Group by quarter and calculate actual achievement values
                var quarterlyData = yearData.GroupBy(x => x.Quarter)
                                       .ToDictionary(g => g.Key, g => new
                                       {
                                           // 将POValue改为TotalCommission
                                           Achieved = g.Sum(x => x.TotalCommission),
                                           MonthlyBreakdown = g.GroupBy(x => x.ReceivedDate.Month)
                                                             .Select(m => new
                                                             {
                                                                 Month = m.Key,
                                                                 // 这里也从POValue改为TotalCommission
                                                                 Value = m.Sum(x => x.TotalCommission)
                                                             }).ToList()
                                       });


                // Update quarterly achievement values
                Q1Achieved = quarterlyData.GetValueOrDefault(1)?.Achieved ?? 0;
                Q2Achieved = quarterlyData.GetValueOrDefault(2)?.Achieved ?? 0;
                Q3Achieved = quarterlyData.GetValueOrDefault(3)?.Achieved ?? 0;
                Q4Achieved = quarterlyData.GetValueOrDefault(4)?.Achieved ?? 0;

                // Calculate overall achievement percentage and remaining amount
                UpdateRemainingValues();

                // Output calculated values for debugging
                Debug.WriteLine($"季度目標值：Q1=${Q1Target}, Q2=${Q2Target}, Q3=${Q3Target}, Q4=${Q4Target}");
                Debug.WriteLine($"季度達成值：Q1=${Q1Achieved}, Q2=${Q2Achieved}, Q3=${Q3Achieved}, Q4=${Q4Achieved}");
                Debug.WriteLine($"最終目標值：Q1=${Q1FinalTarget}, Q2=${Q2FinalTarget}, Q3=${Q3FinalTarget}, Q4=${Q4FinalTarget}");
                Debug.WriteLine($"達成百分比：Q1={Q1Achievement}%, Q2={Q2Achievement}%, Q3={Q3Achievement}%, Q4={Q4Achievement}%");

                // Update all display properties
                RefreshAllDisplayProperties();

                // Update notifications
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

        private void RefreshAllDisplayProperties()
        {
            OnPropertyChanged(nameof(AnnualTargetDisplay));
            OnPropertyChanged(nameof(AchievementDisplay));
            OnPropertyChanged(nameof(RemainingDisplay));
            OnPropertyChanged(nameof(TotalAchievedDisplay));
            OnPropertyChanged(nameof(NotInvoicedDisplay));
            OnPropertyChanged(nameof(InProgressDisplay));

            OnPropertyChanged(nameof(Q1AchievementDisplay));
            OnPropertyChanged(nameof(Q2AchievementDisplay));
            OnPropertyChanged(nameof(Q3AchievementDisplay));
            OnPropertyChanged(nameof(Q4AchievementDisplay));

            OnPropertyChanged(nameof(Q1TargetDisplay));
            OnPropertyChanged(nameof(Q2TargetDisplay));
            OnPropertyChanged(nameof(Q3TargetDisplay));
            OnPropertyChanged(nameof(Q4TargetDisplay));

            OnPropertyChanged(nameof(Q1BaseDisplay));
            OnPropertyChanged(nameof(Q2BaseDisplay));
            OnPropertyChanged(nameof(Q3BaseDisplay));
            OnPropertyChanged(nameof(Q4BaseDisplay));

            OnPropertyChanged(nameof(Q1AchievedDisplay));
            OnPropertyChanged(nameof(Q2AchievedDisplay));
            OnPropertyChanged(nameof(Q3AchievedDisplay));
            OnPropertyChanged(nameof(Q4AchievedDisplay));

            OnPropertyChanged(nameof(Q1FinalTarget));
            OnPropertyChanged(nameof(Q2FinalTarget));
            OnPropertyChanged(nameof(Q3FinalTarget));
            OnPropertyChanged(nameof(Q4FinalTarget));

            OnPropertyChanged(nameof(Q1Carried));
            OnPropertyChanged(nameof(Q2Carried));
            OnPropertyChanged(nameof(Q3Carried));

            OnPropertyChanged(nameof(Q1Exceeded));
            OnPropertyChanged(nameof(Q2Exceeded));
            OnPropertyChanged(nameof(Q3Exceeded));
            OnPropertyChanged(nameof(Q4Exceeded));

            OnPropertyChanged(nameof(Q1CarriedDisplay));
            OnPropertyChanged(nameof(Q2CarriedDisplay));
            OnPropertyChanged(nameof(Q3CarriedDisplay));

            OnPropertyChanged(nameof(Q1CarriedAddedDisplay));
            OnPropertyChanged(nameof(Q2CarriedAddedDisplay));
            OnPropertyChanged(nameof(Q3CarriedAddedDisplay));

            OnPropertyChanged(nameof(Q1ExceededDisplay));
            OnPropertyChanged(nameof(Q2ExceededDisplay));
            OnPropertyChanged(nameof(Q3ExceededDisplay));
            OnPropertyChanged(nameof(Q4ExceededDisplay));
        }

        private void UpdateRemainingValues()
        {
            try
            {
                // 計算總達成值 - 這是已經完成的部分
                var totalAchieved = Q1Achieved + Q2Achieved + Q3Achieved + Q4Achieved;

                // 從 ExcelService 獲取未完成訂單的剩餘金額 (即 Y 欄為空的記錄的 N 欄總和)
                decimal remainingAmount = 0;
                decimal inProgressAmount = 0;
                if (_excelService != null)
                {
                    remainingAmount = _excelService.GetRemainingAmount();
                    Debug.WriteLine($"從 ExcelService 獲取的剩餘金額: ${remainingAmount:N2}");

                    // 獲取正在進行中的金額 (即 Y 欄和 N 欄均為空的記錄的 H 欄總和*0.12)
                    inProgressAmount = _excelService.GetInProgressAmount();
                    Debug.WriteLine($"從 ExcelService 獲取的正在進行中金額: ${inProgressAmount:N2}");
                }

                // 計算達成百分比
                Achievement = AnnualTarget > 0 ? Math.Round((totalAchieved / AnnualTarget) * 100, 1) : 0;

                // 使用新的計算方式設置剩餘金額 - 使用 Y 欄為空的訂單的 N 欄總和
                NotInvoiced = remainingAmount;

                // 設置正在進行中金額
                InProgress = inProgressAmount;

                Remaining = AnnualTarget - totalAchieved;

                // 確保 Remaining 不會小於 0
                if (Remaining < 0) Remaining = 0;

                // 計算進度條的進度 - 使用達成百分比計算
                AchievementProgress = AnnualTarget > 0 ? (double)(totalAchieved / AnnualTarget) : 0;

                // 確保進度不超過 100%
                AchievementProgress = Math.Min(AchievementProgress, 1.0);

                Debug.WriteLine($"更新值: Achievement=${totalAchieved:N2}, {Achievement}%, Remaining=${Remaining:N2}, Progress={AchievementProgress:P2}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"更新 Remaining 值時發生錯誤: {ex.Message}");
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
                    Message = $"Q1 target not achieved! ${Q1Carried:N0} carried over to Q2"
                });
            }

            // Q2 未達成通知
            if (Q2Carried > 0)
            {
                Debug.WriteLine($"Q2 未達成通知: ${Q2Carried:N0}");
                newNotifications.Add(new NotificationItem
                {
                    Message = $"Q2 target not achieved! ${Q2Carried:N0} carried over to Q3"
                });
            }

            // Q3 未達成通知
            if (Q3Carried > 0)
            {
                Debug.WriteLine($"Q3 未達成通知: ${Q3Carried:N0}");
                newNotifications.Add(new NotificationItem
                {
                    Message = $"Q3 target not achieved! ${Q3Carried:N0} carried over to Q4"
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

        // 處理目標更新通知
        private void OnTargetsUpdated(object sender, EventArgs e)
        {
            Debug.WriteLine("收到目標更新通知");
            MainThread.InvokeOnMainThreadAsync(async () =>
            {
                // Get selected fiscal year and update targets
                var selectedFiscalYear = GetSelectedFiscalYear();
                await UpdateTargets(selectedFiscalYear);
                await LoadDataAsync();
                Debug.WriteLine("目標更新處理完成");
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

        // Target distribution and calculation methods

        private void DistributeAnnualTarget()
        {
            Debug.WriteLine($"分配年度目標: ${AnnualTarget:N0}");

            // Calculate the number of quarters that haven't been manually edited
            int uneditedQuarters = 0;
            if (!_isQ1Edited) uneditedQuarters++;
            if (!_isQ2Edited) uneditedQuarters++;
            if (!_isQ3Edited) uneditedQuarters++;
            if (!_isQ4Edited) uneditedQuarters++;

            // If all quarters have been edited or no quarters have been edited
            if (uneditedQuarters == 0 || uneditedQuarters == 4)
            {
                // Distribute evenly to all quarters
                decimal quarterlyAmount = Math.Round(AnnualTarget / 4);
                Q1Target = quarterlyAmount;
                Q2Target = quarterlyAmount;
                Q3Target = quarterlyAmount;
                Q4Target = quarterlyAmount;

                // Adjust any rounding discrepancy to Q4
                decimal totalQuarterly = Q1Target + Q2Target + Q3Target + Q4Target;
                if (totalQuarterly != AnnualTarget)
                {
                    Q4Target += (AnnualTarget - totalQuarterly);
                }

                // Reset edit flags since we've evenly distributed
                _isQ1Edited = false;
                _isQ2Edited = false;
                _isQ3Edited = false;
                _isQ4Edited = false;
            }
            else
            {
                // Distribute remaining amount among unedited quarters
                RedistributeRemainingAmount();
            }

            Debug.WriteLine($"目標分配完成: Q1=${Q1Target:N0}, Q2=${Q2Target:N0}, Q3=${Q3Target:N0}, Q4=${Q4Target:N0}");
        }

        private void RedistributeRemainingAmount()
        {
            if (!IsEditMode) return;

            Debug.WriteLine("重新分配剩餘目標金額");

            // Count unedited quarters
            int uneditedQuarters = 0;
            if (!_isQ1Edited) uneditedQuarters++;
            if (!_isQ2Edited) uneditedQuarters++;
            if (!_isQ3Edited) uneditedQuarters++;
            if (!_isQ4Edited) uneditedQuarters++;

            // If all quarters have been edited, do nothing
            if (uneditedQuarters == 0) return;

            // Calculate total of edited quarters
            decimal totalEdited = 0;
            if (_isQ1Edited) totalEdited += Q1Target;
            if (_isQ2Edited) totalEdited += Q2Target;
            if (_isQ3Edited) totalEdited += Q3Target;
            if (_isQ4Edited) totalEdited += Q4Target;

            // Calculate remaining amount to distribute
            decimal remainingAmount = AnnualTarget - totalEdited;

            // If remaining amount is negative, handle special case
            if (remainingAmount < 0)
            {
                Debug.WriteLine("警告: 编辑的季度目标超过了年度目标，将重置所有季度目标");
                DistributeAnnualTarget();
                return;
            }

            // Distribute evenly among unedited quarters
            decimal amountPerQuarter = Math.Round(remainingAmount / uneditedQuarters);

            if (!_isQ1Edited) Q1Target = amountPerQuarter;
            if (!_isQ2Edited) Q2Target = amountPerQuarter;
            if (!_isQ3Edited) Q3Target = amountPerQuarter;
            if (!_isQ4Edited) Q4Target = amountPerQuarter;

            // Adjust any rounding discrepancy to last unedited quarter
            decimal totalQuarterly = Q1Target + Q2Target + Q3Target + Q4Target;
            decimal adjustment = AnnualTarget - totalQuarterly;

            if (adjustment != 0)
            {
                if (!_isQ4Edited) Q4Target += adjustment;
                else if (!_isQ3Edited) Q3Target += adjustment;
                else if (!_isQ2Edited) Q2Target += adjustment;
                else if (!_isQ1Edited) Q1Target += adjustment;
            }

            Debug.WriteLine($"重新分配完成: Q1=${Q1Target:N0}, Q2=${Q2Target:N0}, Q3=${Q3Target:N0}, Q4=${Q4Target:N0}");
        }

        private void RecalculateAnnualTarget()
        {
            if (!IsEditMode) return;

            // Only recalculate annual if it wasn't manually edited
            if (!_isAnnualEdited)
            {
                decimal newAnnualTarget = Q1Target + Q2Target + Q3Target + Q4Target;
                _annualTarget = newAnnualTarget; // Direct assignment to avoid recursion
                OnPropertyChanged(nameof(AnnualTarget));
                OnPropertyChanged(nameof(AnnualTargetDisplay));
                Debug.WriteLine($"重新計算年度目標: ${AnnualTarget:N0}");
            }
        }

        private void ResetEditFlags()
        {
            _isQ1Edited = false;
            _isQ2Edited = false;
            _isQ3Edited = false;
            _isQ4Edited = false;
            _isAnnualEdited = false;

            Debug.WriteLine("已重置編輯標記");
        }

        // Edit mode commands

        [RelayCommand]
        private void ToggleEditMode()
        {
            IsEditMode = !IsEditMode;

            if (IsEditMode)
            {
                // Entering edit mode
                _originalAnnualTarget = AnnualTarget;
                _originalQ1Target = Q1Target;
                _originalQ2Target = Q2Target;
                _originalQ3Target = Q3Target;
                _originalQ4Target = Q4Target;

                ResetEditFlags();

                ShowStatusMessage("Edit mode on, you can change targets now", true);
                Debug.WriteLine("進入編輯模式");
            }
            else
            {
                // Exiting edit mode without saving
                AnnualTarget = _originalAnnualTarget;
                Q1Target = _originalQ1Target;
                Q2Target = _originalQ2Target;
                Q3Target = _originalQ3Target;
                Q4Target = _originalQ4Target;

                ResetEditFlags();
                ShowStatusMessage("Edit mode off，Targets resets", false);
                Debug.WriteLine("退出編輯模式而不儲存");
            }
        }

        [RelayCommand]
        private async Task SaveChanges()
        {
            try
            {
                IsLoading = true;

                // Create FiscalYearTarget object to save
                var fiscalYear = GetSelectedFiscalYear();
                var companyTarget = new FiscalYearTarget
                {
                    FiscalYear = fiscalYear,
                    AnnualTarget = AnnualTarget,
                    Q1Target = Q1Target,
                    Q2Target = Q2Target,
                    Q3Target = Q3Target,
                    Q4Target = Q4Target
                };

                // Save to target service
                bool success = await _targetService.UpdateCompanyTargetAsync(companyTarget);
                if (success)
                {
                    // Update originals
                    _originalAnnualTarget = AnnualTarget;
                    _originalQ1Target = Q1Target;
                    _originalQ2Target = Q2Target;
                    _originalQ3Target = Q3Target;
                    _originalQ4Target = Q4Target;

                    // Exit edit mode
                    IsEditMode = false;
                    ResetEditFlags();

                    ShowStatusMessage("Target set success", true);
                    Debug.WriteLine("目標設定已儲存");
                }
                else
                {
                    ShowStatusMessage("Target set failed, Please try again", false);
                    Debug.WriteLine("儲存目標設定失敗");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"儲存目標設定時發生錯誤: {ex.Message}");
                ShowStatusMessage($"Fail to save: {ex.Message}", false);
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        private void CancelEditing()
        {
            // Restore original values
            AnnualTarget = _originalAnnualTarget;
            Q1Target = _originalQ1Target;
            Q2Target = _originalQ2Target;
            Q3Target = _originalQ3Target;
            Q4Target = _originalQ4Target;

            // Exit edit mode
            IsEditMode = false;
            ResetEditFlags();

            ShowStatusMessage("Edit canceled", false);
            Debug.WriteLine("已取消編輯");
        }

        [RelayCommand]
        private void ResetQuarterlyTargets()
        {
            if (!IsEditMode) return;

            // Reset all quarters to be evenly distributed
            DistributeAnnualTarget();

            ShowStatusMessage("Resetted Target", true);
            Debug.WriteLine("已重置季度目標");
        }

        // Status message display
        private void ShowStatusMessage(string message, bool isSuccess)
        {
            StatusMessage = message;
            IsStatusSuccess = isSuccess;
            IsStatusMessageVisible = true;

            // Hide message after delay
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await Task.Delay(5000);
                IsStatusMessageVisible = false;
            });
        }

        // Clean up resources
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
    }
}