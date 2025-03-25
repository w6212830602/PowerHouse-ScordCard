using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Configuration;
using ScoreCard.Models;
using ScoreCard.Services;

namespace ScoreCard.ViewModels
{
    public partial class SettingsViewModel : ObservableObject
    {
        private readonly IExcelService _excelService;
        private readonly ITargetService _targetService;
        private IConfiguration _configuration;

        // Selected tab
        [ObservableProperty]
        private string _selectedTab = "CompanyTarget";

        // Company Targets
        [ObservableProperty]
        private ObservableCollection<FiscalYearTarget> _companyTargets = new();

        // Individual Targets (Sales Reps)
        [ObservableProperty]
        private ObservableCollection<SalesRepTarget> _salesRepTargets = new();

        // LOB Targets
        [ObservableProperty]
        private ObservableCollection<LOBTarget> _lobTargets = new();

        // Edit mode flags
        [ObservableProperty]
        private bool _isEditingTargets = false;

        // Loading indicator
        [ObservableProperty]
        private bool _isLoading;

        // Selected fiscal year for individual and LOB targets
        [ObservableProperty]
        private string _selectedFiscalYear;

        // Available fiscal years 
        [ObservableProperty]
        private ObservableCollection<string> _fiscalYears = new();

        // Message to display after save operations
        [ObservableProperty]
        private string _statusMessage;

        // Status message visibility
        [ObservableProperty]
        private bool _isStatusMessageVisible;

        // Status message is success or error
        [ObservableProperty]
        private bool _isStatusSuccess;

        [RelayCommand]
        private void EditTargets()
        {
            IsEditingTargets = true;
        }

        // 為每行添加編輯按鈕的命令
        [RelayCommand]
        private void EditTargetRow(object parameter)
        {
            // 根據當前標籤確定編輯的是哪種目標
            if (SelectedTab == "CompanyTarget" && parameter is FiscalYearTarget target)
            {
                // 儲存原始值，以便在取消時還原
                _originalCompanyTarget = new FiscalYearTarget
                {
                    FiscalYear = target.FiscalYear,
                    AnnualTarget = target.AnnualTarget,
                    Q1Target = target.Q1Target,
                    Q2Target = target.Q2Target,
                    Q3Target = target.Q3Target,
                    Q4Target = target.Q4Target
                };

                // 設置為編輯模式
                _editingRowId = target.FiscalYear;
                IsEditingRow = true;
            }
            else if (SelectedTab == "IndividualTarget" && parameter is SalesRepTarget repTarget)
            {
                // 類似的處理銷售代表目標的編輯
                _editingRowId = repTarget.SalesRep;
                IsEditingRow = true;
            }
            else if (SelectedTab == "LOBTargets" && parameter is LOBTarget lobTarget)
            {
                // 處理 LOB 目標的編輯
                _editingRowId = lobTarget.LOB;
                IsEditingRow = true;
            }
        }

        public SettingsViewModel(IExcelService excelService, ITargetService targetService)
        {
            _excelService = excelService;
            _targetService = targetService;
            InitializeAsync();
        }

        private async void InitializeAsync()
        {
            try
            {
                IsLoading = true;

                // Initialize target service
                await _targetService.InitializeAsync();

                // Load configuration
                _configuration = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // Load company targets from settings
                LoadCompanyTargetsFromSettings();

                // Setup fiscal years
                SetupFiscalYears();

                // Load sales reps and LOB data
                await LoadSalesRepTargets();
                await LoadLOBTargets();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing SettingsViewModel: {ex.Message}");
                ShowStatusMessage($"Error initializing settings: {ex.Message}", false);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void SetupFiscalYears()
        {
            // Create list of available fiscal years based on company targets
            FiscalYears.Clear();

            foreach (var target in CompanyTargets)
            {
                FiscalYears.Add($"FY{target.FiscalYear}");
            }

            // Select the most recent fiscal year by default
            if (FiscalYears.Any())
            {
                SelectedFiscalYear = FiscalYears.First();
            }
        }

        private void LoadCompanyTargetsFromSettings()
        {
            try
            {
                // Get company targets from target service
                var companyTargets = new List<FiscalYearTarget>();
                
                // Get current fiscal year and several previous/next years
                var currentDate = DateTime.Now;
                var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;
                
                for (int year = currentFiscalYear - 2; year <= currentFiscalYear + 2; year++)
                {
                    var target = _targetService.GetCompanyTarget(year);
                    if (target != null)
                    {
                        companyTargets.Add(target);
                    }
                }

                if (companyTargets.Any())
                {
                    CompanyTargets = new ObservableCollection<FiscalYearTarget>(
                        companyTargets.OrderByDescending(t => t.FiscalYear));
                }
                else
                {
                    // Default targets if none found in settings
                    CompanyTargets = new ObservableCollection<FiscalYearTarget>
                    {
                        new FiscalYearTarget
                        {
                            FiscalYear = currentFiscalYear + 1,
                            AnnualTarget = 5000000,
                            Q1Target = 1250000,
                            Q2Target = 1250000,
                            Q3Target = 1250000,
                            Q4Target = 1250000
                        },
                        new FiscalYearTarget
                        {
                            FiscalYear = currentFiscalYear,
                            AnnualTarget = 4500000,
                            Q1Target = 1125000,
                            Q2Target = 1125000,
                            Q3Target = 1125000,
                            Q4Target = 1125000
                        },
                        new FiscalYearTarget
                        {
                            FiscalYear = currentFiscalYear - 1,
                            AnnualTarget = 4000000,
                            Q1Target = 1000000,
                            Q2Target = 1000000,
                            Q3Target = 1000000,
                            Q4Target = 1000000
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading company targets: {ex.Message}");
                ShowStatusMessage($"Error loading company targets: {ex.Message}", false);
            }
        }

        private async Task LoadSalesRepTargets()
        {
            try
            {
                // 獲取當前選擇的財年值
                int currentFiscalYear = GetSelectedFiscalYearValue();
                Debug.WriteLine($"正在加載 {currentFiscalYear} 財年的銷售代表目標");

                // 嘗試從目標服務中獲取已保存的目標
                var savedTargets = _targetService.GetSalesRepTargets(currentFiscalYear);

                if (savedTargets?.Any() == true)
                {
                    // 如果已有保存的目標，直接使用
                    SalesRepTargets = new ObservableCollection<SalesRepTarget>(savedTargets);
                    Debug.WriteLine($"已加載 {savedTargets.Count} 個銷售代表目標");
                    return;
                }

                // 從 Excel 獲取所有銷售代表
                var allReps = _excelService.GetAllSalesReps();
                Debug.WriteLine($"從 Excel 讀取到 {allReps.Count} 個銷售代表");

                // 獲取公司年度總目標用於分配
                var companyTarget = CompanyTargets.FirstOrDefault(t => t.FiscalYear == currentFiscalYear);
                decimal annualTarget = companyTarget?.AnnualTarget ?? 4000000m;

                // 為每個銷售代表創建目標
                SalesRepTargets = new ObservableCollection<SalesRepTarget>();

                if (allReps.Any())
                {
                    // 計算每位代表的平均目標值
                    decimal avgTarget = annualTarget / allReps.Count;
                    // 四捨五入到最接近的 1000
                    avgTarget = Math.Round(avgTarget / 1000) * 1000;
                    decimal quarterlyTarget = avgTarget / 4;

                    foreach (var rep in allReps)
                    {
                        SalesRepTargets.Add(new SalesRepTarget
                        {
                            SalesRep = rep,
                            FiscalYear = currentFiscalYear,
                            AnnualTarget = avgTarget,
                            Q1Target = quarterlyTarget,
                            Q2Target = quarterlyTarget,
                            Q3Target = quarterlyTarget,
                            Q4Target = quarterlyTarget
                        });
                    }

                    Debug.WriteLine($"已為 {allReps.Count} 個銷售代表創建目標");
                }
                else
                {
                    // 如果沒有找到銷售代表，添加默認數據
                    AddDefaultSalesRepTargets(currentFiscalYear);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加載銷售代表目標時出錯: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // 確保視圖始終有數據顯示
                if (SalesRepTargets == null || !SalesRepTargets.Any())
                {
                    SalesRepTargets = new ObservableCollection<SalesRepTarget>();
                    AddDefaultSalesRepTargets(GetSelectedFiscalYearValue());
                }

                // 向用戶顯示錯誤消息
                ShowStatusMessage($"加載銷售代表目標失敗: {ex.Message}", false);
            }
        }

        // 辅助方法：添加默认销售代表目标
        private void AddDefaultSalesRepTargets(int fiscalYear)
        {
            var defaultReps = new[]
            {
        ("Brandon", 1000000m),
        ("Chris", 900000m),
        ("Isaac", 850000m),
        ("Mark", 800000m),
        ("Nathan", 750000m)
    };

            foreach (var (name, target) in defaultReps)
            {
                decimal quarterlyTarget = target / 4;
                SalesRepTargets.Add(new SalesRepTarget
                {
                    SalesRep = name,
                    FiscalYear = fiscalYear,
                    AnnualTarget = target,
                    Q1Target = quarterlyTarget,
                    Q2Target = quarterlyTarget,
                    Q3Target = quarterlyTarget,
                    Q4Target = quarterlyTarget
                });
            }

            Debug.WriteLine($"已添加 {defaultReps.Length} 个默认销售代表目标");
        }

        private async Task LoadLOBTargets()
        {
            try
            {
                // 獲取當前財年
                int currentFiscalYear = GetSelectedFiscalYearValue();

                // 獲取 LOB 目標從目標服務
                var savedTargets = _targetService.GetLOBTargets(currentFiscalYear);

                if (savedTargets?.Any() == true)
                {
                    LobTargets = new ObservableCollection<LOBTarget>(savedTargets);
                    Debug.WriteLine($"已加載 {savedTargets.Count} 個 LOB 目標");
                    return;
                }

                // 從 Excel 獲取所有 LOB
                var allLOBs = _excelService.GetAllLOBs();
                Debug.WriteLine($"從 Excel 讀取到 {allLOBs.Count} 個 LOB");

                // 獲取公司年度總目標用於分配
                var companyTarget = CompanyTargets.FirstOrDefault(t => t.FiscalYear == currentFiscalYear);
                decimal annualTarget = companyTarget?.AnnualTarget ?? 4000000m;

                LobTargets = new ObservableCollection<LOBTarget>();

                if (allLOBs.Any())
                {
                    // 計算每個 LOB 的平均目標值
                    decimal avgTarget = annualTarget / allLOBs.Count;
                    // 四捨五入到最接近的 1000
                    avgTarget = Math.Round(avgTarget / 1000) * 1000;
                    decimal quarterlyTarget = avgTarget / 4;

                    foreach (var lob in allLOBs)
                    {
                        LobTargets.Add(new LOBTarget
                        {
                            LOB = lob,
                            FiscalYear = currentFiscalYear,
                            AnnualTarget = avgTarget,
                            Q1Target = quarterlyTarget,
                            Q2Target = quarterlyTarget,
                            Q3Target = quarterlyTarget,
                            Q4Target = quarterlyTarget
                        });
                    }

                    Debug.WriteLine($"已為 {allLOBs.Count} 個 LOB 創建目標");
                }
                else
                {
                    // 如果沒有找到 LOB，添加默認數據
                    AddDefaultLOBTargets(currentFiscalYear);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"加載 LOB 目標時出錯: {ex.Message}");

                // 確保視圖始終有數據顯示
                if (LobTargets == null || !LobTargets.Any())
                {
                    LobTargets = new ObservableCollection<LOBTarget>();
                    AddDefaultLOBTargets(GetSelectedFiscalYearValue());
                }

                ShowStatusMessage($"加載 LOB 目標失敗: {ex.Message}", false);
            }
        }

        // 添加默認 LOB 目標的輔助方法
        private void AddDefaultLOBTargets(int fiscalYear)
        {
            var defaultLOBs = new[]
            {
        ("Power", 1000000m),
        ("Thermal", 900000m),
        ("Channel", 750000m),
        ("Service", 500000m),
        ("Batts & Caps", 400000m)
    };

            foreach (var (name, target) in defaultLOBs)
            {
                decimal quarterlyTarget = target / 4;
                LobTargets.Add(new LOBTarget
                {
                    LOB = name,
                    FiscalYear = fiscalYear,
                    AnnualTarget = target,
                    Q1Target = quarterlyTarget,
                    Q2Target = quarterlyTarget,
                    Q3Target = quarterlyTarget,
                    Q4Target = quarterlyTarget
                });
            }

            Debug.WriteLine($"已添加 {defaultLOBs.Length} 個默認 LOB 目標");
        }

        private int GetSelectedFiscalYearValue()
        {
            if (string.IsNullOrEmpty(SelectedFiscalYear))
            {
                // Default to current year if nothing is selected
                var currentDate = DateTime.Now;
                return currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;
            }

            if (int.TryParse(SelectedFiscalYear.Replace("FY", ""), out int result))
            {
                return result;
            }

            // Default to current year if parsing fails
            var date = DateTime.Now;
            return date.Month >= 8 ? date.Year + 1 : date.Year;
        }

        // Tab switching command
        [RelayCommand]
        private void SwitchTab(string tabName)
        {
            if (!string.IsNullOrEmpty(tabName) && SelectedTab != tabName)
            {
                SelectedTab = tabName;

                // Load appropriate data for selected tab
                if (tabName == "IndividualTarget" || tabName == "LOBTargets")
                {
                    MainThread.BeginInvokeOnMainThread(async () =>
                    {
                        IsLoading = true;
                        try
                        {
                            if (tabName == "IndividualTarget")
                            {
                                await LoadSalesRepTargets();
                            }
                            else if (tabName == "LOBTargets")
                            {
                                await LoadLOBTargets();
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error loading data for tab {tabName}: {ex.Message}");
                        }
                        finally
                        {
                            IsLoading = false;
                        }
                    });
                }
            }
        }

        // Toggle edit mode
        [RelayCommand]
        private void ToggleEditMode()
        {
            IsEditingTargets = !IsEditingTargets;
        }

        // Update annual target based on quarterly targets
        [RelayCommand]
        private void UpdateAnnualTarget(object parameter)
        {
            if (parameter is FiscalYearTarget target)
            {
                target.AnnualTarget = target.Q1Target + target.Q2Target + target.Q3Target + target.Q4Target;
                OnPropertyChanged(nameof(CompanyTargets));
            }
            else if (parameter is SalesRepTarget repTarget)
            {
                repTarget.AnnualTarget = repTarget.Q1Target + repTarget.Q2Target + repTarget.Q3Target + repTarget.Q4Target;
                OnPropertyChanged(nameof(SalesRepTargets));
            }
            else if (parameter is LOBTarget lobTarget)
            {
                lobTarget.AnnualTarget = lobTarget.Q1Target + lobTarget.Q2Target + lobTarget.Q3Target + lobTarget.Q4Target;
                OnPropertyChanged(nameof(LobTargets));
            }
        }

        // Distribute annual target to quarters
        [RelayCommand]
        private void DistributeTarget(object parameter)
        {
            if (parameter is FiscalYearTarget target)
            {
                decimal quarterlyTarget = target.AnnualTarget / 4;
                target.Q1Target = quarterlyTarget;
                target.Q2Target = quarterlyTarget;
                target.Q3Target = quarterlyTarget;
                target.Q4Target = quarterlyTarget;
                OnPropertyChanged(nameof(CompanyTargets));
            }
            else if (parameter is SalesRepTarget repTarget)
            {
                decimal quarterlyTarget = repTarget.AnnualTarget / 4;
                repTarget.Q1Target = quarterlyTarget;
                repTarget.Q2Target = quarterlyTarget;
                repTarget.Q3Target = quarterlyTarget;
                repTarget.Q4Target = quarterlyTarget;
                OnPropertyChanged(nameof(SalesRepTargets));
            }
            else if (parameter is LOBTarget lobTarget)
            {
                decimal quarterlyTarget = lobTarget.AnnualTarget / 4;
                lobTarget.Q1Target = quarterlyTarget;
                lobTarget.Q2Target = quarterlyTarget;
                lobTarget.Q3Target = quarterlyTarget;
                lobTarget.Q4Target = quarterlyTarget;
                OnPropertyChanged(nameof(LobTargets));
            }
        }

        // Save changes command
        [RelayCommand]
        private async Task SaveChanges()
        {
            try
            {
                IsLoading = true;

                // Set edit mode to false
                IsEditingTargets = false;

                // Save company targets to appsettings.json
                await SaveCompanyTargetsAsync();

                // Save sales rep targets to file
                await SaveSalesRepTargetsAsync();

                // Save LOB targets to file
                await SaveLOBTargetsAsync();

                // Notify target service that targets have been updated
                _targetService.NotifyTargetsUpdated();

                // Show success message
                ShowStatusMessage("All target settings have been saved successfully.", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving changes: {ex.Message}");
                ShowStatusMessage($"Failed to save target settings: {ex.Message}", false);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async Task SaveCompanyTargetsAsync()
        {
            try
            {
                // Update TargetSettings object
                var targetSettings = new TargetSettings
                {
                    CompanyTargets = CompanyTargets.ToList()
                };

                // Get the path to the appsettings.json file
                string settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");

                // Read the existing file content
                string json = await File.ReadAllTextAsync(settingsPath);

                // Deserialize to a dictionary
                var jsonDoc = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);

                // Serialize the updated TargetSettings
                var targetSettingsJson = JsonSerializer.Serialize(targetSettings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Update the TargetSettings node in the jsonDoc
                var targetSettingsElement = JsonSerializer.Deserialize<JsonElement>(targetSettingsJson);
                jsonDoc["TargetSettings"] = targetSettingsElement;

                // Serialize back to JSON
                var updatedJson = JsonSerializer.Serialize(jsonDoc, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Write back to the file
                await File.WriteAllTextAsync(settingsPath, updatedJson);

                Debug.WriteLine("Company targets saved to appsettings.json");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving company targets: {ex.Message}");
                throw; // Re-throw to be caught by the caller
            }
        }

        private async Task SaveSalesRepTargetsAsync()
        {
            if (SalesRepTargets == null || !SalesRepTargets.Any())
                return;

            try
            {
                // Set the current fiscal year for all targets
                int currentFY = GetSelectedFiscalYearValue();
                foreach (var target in SalesRepTargets)
                {
                    target.FiscalYear = currentFY;
                }

                // Create directory if it doesn't exist
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");
                Directory.CreateDirectory(targetDir);

                // Save to file
                string filePath = Path.Combine(targetDir, $"SalesRepTargets_{currentFY}.json");
                string json = JsonSerializer.Serialize(SalesRepTargets, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                await File.WriteAllTextAsync(filePath, json);
                Debug.WriteLine($"Sales rep targets saved to {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving sales rep targets: {ex.Message}");
                throw; // Re-throw to be caught by the caller
            }
        }

        private async Task SaveLOBTargetsAsync()
        {
            if (LobTargets == null || !LobTargets.Any())
                return;

            try
            {
                // Set the current fiscal year for all targets
                int currentFY = GetSelectedFiscalYearValue();
                foreach (var target in LobTargets)
                {
                    target.FiscalYear = currentFY;
                }

                // Create directory if it doesn't exist
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");
                Directory.CreateDirectory(targetDir);

                // Save to file
                string filePath = Path.Combine(targetDir, $"LOBTargets_{currentFY}.json");
                string json = JsonSerializer.Serialize(LobTargets, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                await File.WriteAllTextAsync(filePath, json);
                Debug.WriteLine($"LOB targets saved to {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving LOB targets: {ex.Message}");
                throw; // Re-throw to be caught by the caller
            }
        }

        // Add new fiscal year
        [RelayCommand]
        private void AddFiscalYear()
        {
            try
            {
                // Find the highest fiscal year
                int maxYear = CompanyTargets.Max(t => t.FiscalYear);

                // Create new fiscal year with values copied from the previous year
                var previousYear = CompanyTargets.FirstOrDefault(t => t.FiscalYear == maxYear);
                var newTarget = new FiscalYearTarget
                {
                    FiscalYear = maxYear + 1,
                    AnnualTarget = previousYear?.AnnualTarget * 1.05m ?? 5000000m, // 5% increase or default
                    Q1Target = previousYear?.Q1Target * 1.05m ?? 1250000m,
                    Q2Target = previousYear?.Q2Target * 1.05m ?? 1250000m,
                    Q3Target = previousYear?.Q3Target * 1.05m ?? 1250000m,
                    Q4Target = previousYear?.Q4Target * 1.05m ?? 1250000m
                };

                // Round values to nearest 10,000
                newTarget.AnnualTarget = Math.Round(newTarget.AnnualTarget / 10000) * 10000;
                newTarget.Q1Target = Math.Round(newTarget.Q1Target / 10000) * 10000;
                newTarget.Q2Target = Math.Round(newTarget.Q2Target / 10000) * 10000;
                newTarget.Q3Target = Math.Round(newTarget.Q3Target / 10000) * 10000;
                newTarget.Q4Target = Math.Round(newTarget.Q4Target / 10000) * 10000;

                // Add new fiscal year to the collection (at the beginning)
                CompanyTargets.Insert(0, newTarget);

                // Update fiscal years list
                SetupFiscalYears();

                ShowStatusMessage($"Added new fiscal year FY{maxYear + 1}", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding fiscal year: {ex.Message}");
                ShowStatusMessage($"Error adding fiscal year: {ex.Message}", false);
            }
        }

        // Add sales rep command
        [RelayCommand]
        private void AddSalesRep()
        {
            try
            {
                // Get average target from existing reps
                decimal avgTarget = 500000m; // Default fallback
                if (SalesRepTargets.Any())
                {
                    avgTarget = SalesRepTargets.Average(r => r.AnnualTarget);
                }

                // Round to nearest 50,000
                avgTarget = Math.Round(avgTarget / 50000) * 50000;
                decimal quarterlyTarget = avgTarget / 4;

                SalesRepTargets.Add(new SalesRepTarget
                {
                    SalesRep = "New Rep",
                    FiscalYear = GetSelectedFiscalYearValue(),
                    AnnualTarget = avgTarget,
                    Q1Target = quarterlyTarget,
                    Q2Target = quarterlyTarget,
                    Q3Target = quarterlyTarget,
                    Q4Target = quarterlyTarget
                });

                ShowStatusMessage("Added new sales rep", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding sales rep: {ex.Message}");
                ShowStatusMessage($"Error adding sales rep: {ex.Message}", false);
            }
        }

        // Remove sales rep command
        [RelayCommand]
        private void RemoveSalesRep(SalesRepTarget repTarget)
        {
            if (repTarget != null)
            {
                SalesRepTargets.Remove(repTarget);
                ShowStatusMessage($"Removed sales rep {repTarget.SalesRep}", true);
            }
        }

        // Add LOB command
        [RelayCommand]
        private void AddLOB()
        {
            try
            {
                // Get average target from existing LOBs
                decimal avgTarget = 500000m; // Default fallback
                if (LobTargets.Any())
                {
                    avgTarget = LobTargets.Average(l => l.AnnualTarget);
                }

                // Round to nearest 50,000
                avgTarget = Math.Round(avgTarget / 50000) * 50000;
                decimal quarterlyTarget = avgTarget / 4;

                LobTargets.Add(new LOBTarget
                {
                    LOB = "New Line of Business",
                    FiscalYear = GetSelectedFiscalYearValue(),
                    AnnualTarget = avgTarget,
                    Q1Target = quarterlyTarget,
                    Q2Target = quarterlyTarget,
                    Q3Target = quarterlyTarget,
                    Q4Target = quarterlyTarget
                });

                ShowStatusMessage("Added new LOB", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding LOB: {ex.Message}");
                ShowStatusMessage($"Error adding LOB: {ex.Message}", false);
            }
        }

        // Remove LOB command
        [RelayCommand]
        private void RemoveLOB(LOBTarget lobTarget)
        {
            if (lobTarget != null)
            {
                LobTargets.Remove(lobTarget);
                ShowStatusMessage($"Removed LOB {lobTarget.LOB}", true);
            }
        }

        // Change fiscal year for individual and LOB targets
        partial void OnSelectedFiscalYearChanged(string value)
        {
            if (string.IsNullOrEmpty(value))
                return;

            Debug.WriteLine($"Selected fiscal year changed to: {value}");

            // Load targets for the selected fiscal year
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                IsLoading = true;
                try
                {
                    await LoadSalesRepTargets();
                    await LoadLOBTargets();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error loading targets for fiscal year {value}: {ex.Message}");
                    ShowStatusMessage($"Error loading targets: {ex.Message}", false);
                }
                finally
                {
                    IsLoading = false;
                }
            });
        }

        private void ShowStatusMessage(string message, bool isSuccess)
        {
            StatusMessage = message;
            IsStatusSuccess = isSuccess;
            IsStatusMessageVisible = true;

            // Hide the message after a delay
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await Task.Delay(5000);
                IsStatusMessageVisible = false;
            });
        }

        [RelayCommand]
        private void AddRep()
        {
            if (!IsEditingTargets)
                return;

            try
            {
                // 獲取平均目標值
                decimal avgTarget = 500000m;
                if (SalesRepTargets.Any())
                {
                    avgTarget = SalesRepTargets.Average(r => r.AnnualTarget);
                }

                // 四捨五入到最接近的 50,000
                avgTarget = Math.Round(avgTarget / 50000) * 50000;
                decimal quarterlyTarget = avgTarget / 4;

                // 創建新的銷售代表目標
                var newTarget = new SalesRepTarget
                {
                    SalesRep = "新代表",
                    FiscalYear = GetSelectedFiscalYearValue(),
                    AnnualTarget = avgTarget,
                    Q1Target = quarterlyTarget,
                    Q2Target = quarterlyTarget,
                    Q3Target = quarterlyTarget,
                    Q4Target = quarterlyTarget
                };

                // 添加到集合
                SalesRepTargets.Add(newTarget);

                // 顯示成功消息
                ShowStatusMessage("已添加新銷售代表，請更改代表名稱並設定目標值", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"添加代表時出錯: {ex.Message}");
                ShowStatusMessage($"添加代表失敗: {ex.Message}", false);
            }
        }

        [RelayCommand]
        private void AddLOB()
        {
            if (!IsEditingTargets)
                return;

            try
            {
                // 獲取平均目標值
                decimal avgTarget = 500000m;
                if (LobTargets.Any())
                {
                    avgTarget = LobTargets.Average(l => l.AnnualTarget);
                }

                // 四捨五入到最接近的 50,000
                avgTarget = Math.Round(avgTarget / 50000) * 50000;
                decimal quarterlyTarget = avgTarget / 4;

                // 創建新的 LOB 目標
                var newTarget = new LOBTarget
                {
                    LOB = "新產品線",
                    FiscalYear = GetSelectedFiscalYearValue(),
                    AnnualTarget = avgTarget,
                    Q1Target = quarterlyTarget,
                    Q2Target = quarterlyTarget,
                    Q3Target = quarterlyTarget,
                    Q4Target = quarterlyTarget
                };

                // 添加到集合
                LobTargets.Add(newTarget);

                // 顯示成功消息
                ShowStatusMessage("已添加新產品線，請更改名稱並設定目標值", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"添加 LOB 時出錯: {ex.Message}");
                ShowStatusMessage($"添加產品線失敗: {ex.Message}", false);
            }
        }
    }
}