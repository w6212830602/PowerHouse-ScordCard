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
        private bool _isEditingRow;
        private string _editingRowId;
        private FiscalYearTarget _originalCompanyTarget;
        private bool _isInitialized = false;

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
        private void ToggleEditMode()
        {
            IsEditingTargets = !IsEditingTargets;
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
                _editingRowId = target.FiscalYear.ToString();
                _isEditingRow = true;
            }
            else if (SelectedTab == "IndividualTarget" && parameter is SalesRepTarget repTarget)
            {
                // 類似的處理銷售代表目標的編輯
                _editingRowId = repTarget.SalesRep;
                _isEditingRow = true;
            }
            else if (SelectedTab == "LOBTargets" && parameter is LOBTarget lobTarget)
            {
                // 處理 LOB 目標的編輯
                _editingRowId = lobTarget.LOB;
                _isEditingRow = true;
            }
        }

        public SettingsViewModel(IExcelService excelService, ITargetService targetService)
        {
            _excelService = excelService;
            _targetService = targetService;

            // 注意：不在構造函數中調用 InitializeAsync，而是在頁面出現時調用
            SetupDefaultValues();
        }

        // 設置一些默認值，以便在初始化完成前頁面顯示不會為空
        private void SetupDefaultValues()
        {
            var currentDate = DateTime.Now;
            var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;

            // 設置默認財年列表
            FiscalYears.Clear();
            for (int year = currentFiscalYear - 2; year <= currentFiscalYear + 2; year++)
            {
                FiscalYears.Add($"FY{year}");
            }

            // 默認選擇當前財年
            SelectedFiscalYear = $"FY{currentFiscalYear}";
        }

        // 在頁面出現時調用此方法
        public async Task InitializeAsync()
        {
            // 如果已初始化，則不重複執行
            if (_isInitialized) return;

            try
            {
                IsLoading = true;

                // 初始化目標服務（異步）
                await _targetService.InitializeAsync();

                // 載入設定
                _configuration = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                // 載入公司目標
                LoadCompanyTargetsFromSettings();

                // 設置財年選項
                SetupFiscalYears();

                // 載入銷售代表目標
                await LoadSalesRepTargets();

                // 載入LOB目標
                await LoadLOBTargets();

                _isInitialized = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"初始化SettingsViewModel時出錯: {ex.Message}");
                ShowStatusMessage($"載入設定時出錯: {ex.Message}", false);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void SetupFiscalYears()
        {
            FiscalYears.Clear();

            foreach (var target in CompanyTargets)
            {
                FiscalYears.Add($"FY{target.FiscalYear}");
            }

            // 如果沒有找到財年，添加默認值
            if (!FiscalYears.Any())
            {
                var currentDate = DateTime.Now;
                var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;

                for (int year = currentFiscalYear - 2; year <= currentFiscalYear + 2; year++)
                {
                    FiscalYears.Add($"FY{year}");
                }
            }

            // 選擇第一個財年（通常是最新的）
            if (FiscalYears.Any())
            {
                SelectedFiscalYear = FiscalYears.First();
            }
        }

        private void LoadCompanyTargetsFromSettings()
        {
            try
            {
                var companyTargets = new List<FiscalYearTarget>();

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
                    // 如果沒有找到目標，創建默認目標
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
                Debug.WriteLine($"載入公司目標時出錯: {ex.Message}");
                ShowStatusMessage($"載入公司目標時出錯: {ex.Message}", false);
            }
        }

        private async Task LoadSalesRepTargets()
        {
            try
            {
                int currentFiscalYear = GetSelectedFiscalYearValue();

                // 從目標服務獲取已保存的目標
                var savedTargets = _targetService.GetSalesRepTargets(currentFiscalYear);

                if (savedTargets?.Any() == true)
                {
                    // 直接使用已保存的目標
                    SalesRepTargets = new ObservableCollection<SalesRepTarget>(savedTargets);
                    return;
                }

                // 從Excel獲取銷售代表列表（注意這裡使用異步方法）
                List<string> allReps = await Task.Run(() => GetSalesRepsAsync());

                // 獲取公司目標用於計算平均目標
                var companyTarget = CompanyTargets.FirstOrDefault(t => t.FiscalYear == currentFiscalYear);
                decimal annualTarget = companyTarget?.AnnualTarget ?? 4000000m;

                // 創建新的銷售代表目標列表
                var salesRepTargets = new List<SalesRepTarget>();

                if (allReps.Any())
                {
                    // 計算平均目標
                    decimal avgTarget = annualTarget / allReps.Count;
                    avgTarget = Math.Round(avgTarget / 1000) * 1000;
                    decimal quarterlyTarget = avgTarget / 4;

                    foreach (var rep in allReps)
                    {
                        salesRepTargets.Add(new SalesRepTarget
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
                }
                else
                {
                    // 如果沒有找到銷售代表，添加默認數據
                    AddDefaultSalesRepTargets(salesRepTargets, currentFiscalYear);
                }

                // 更新UI
                SalesRepTargets = new ObservableCollection<SalesRepTarget>(salesRepTargets);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入銷售代表目標時出錯: {ex.Message}");

                // 創建默認數據，確保UI不為空
                var defaultTargets = new List<SalesRepTarget>();
                AddDefaultSalesRepTargets(defaultTargets, GetSelectedFiscalYearValue());
                SalesRepTargets = new ObservableCollection<SalesRepTarget>(defaultTargets);

                ShowStatusMessage($"載入銷售代表目標時出錯: {ex.Message}", false);
            }
        }

        // 從Excel異步獲取銷售代表列表
        private async Task<List<string>> GetSalesRepsAsync()
        {
            try
            {
                return await Task.Run(() => _excelService.GetAllSalesReps());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"獲取銷售代表列表時出錯: {ex.Message}");
                return new List<string> { "Brandon", "Chris", "Isaac", "Mark", "Nathan", "Tania" };
            }
        }

        // 添加默認銷售代表目標
        private void AddDefaultSalesRepTargets(List<SalesRepTarget> targets, int fiscalYear)
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
                targets.Add(new SalesRepTarget
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
        }

        private async Task LoadLOBTargets()
        {
            try
            {
                int currentFiscalYear = GetSelectedFiscalYearValue();

                // 從目標服務獲取已保存的目標
                var savedTargets = _targetService.GetLOBTargets(currentFiscalYear);

                if (savedTargets?.Any() == true)
                {
                    // 直接使用已保存的目標
                    LobTargets = new ObservableCollection<LOBTarget>(savedTargets);
                    return;
                }

                // 從Excel獲取LOB列表（使用異步方法）
                List<string> allLOBs = await Task.Run(() => GetLOBsAsync());

                // 獲取公司目標用於計算平均目標
                var companyTarget = CompanyTargets.FirstOrDefault(t => t.FiscalYear == currentFiscalYear);
                decimal annualTarget = companyTarget?.AnnualTarget ?? 4000000m;

                // 創建新的LOB目標列表
                var lobTargets = new List<LOBTarget>();

                if (allLOBs.Any())
                {
                    // 計算平均目標
                    decimal avgTarget = annualTarget / allLOBs.Count;
                    avgTarget = Math.Round(avgTarget / 1000) * 1000;
                    decimal quarterlyTarget = avgTarget / 4;

                    foreach (var lob in allLOBs)
                    {
                        lobTargets.Add(new LOBTarget
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
                }
                else
                {
                    // 如果沒有找到LOB，添加默認數據
                    AddDefaultLOBTargets(lobTargets, currentFiscalYear);
                }

                // 更新UI
                LobTargets = new ObservableCollection<LOBTarget>(lobTargets);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"載入LOB目標時出錯: {ex.Message}");

                // 創建默認數據，確保UI不為空
                var defaultTargets = new List<LOBTarget>();
                AddDefaultLOBTargets(defaultTargets, GetSelectedFiscalYearValue());
                LobTargets = new ObservableCollection<LOBTarget>(defaultTargets);

                ShowStatusMessage($"載入LOB目標時出錯: {ex.Message}", false);
            }
        }

        // 從Excel異步獲取LOB列表
        private async Task<List<string>> GetLOBsAsync()
        {
            try
            {
                return await Task.Run(() => _excelService.GetAllLOBs());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"獲取LOB列表時出錯: {ex.Message}");
                return new List<string> { "Power", "Thermal", "Channel", "Service", "Batts & Caps" };
            }
        }

        // 添加默認LOB目標
        private void AddDefaultLOBTargets(List<LOBTarget> targets, int fiscalYear)
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
                targets.Add(new LOBTarget
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
        }

        private int GetSelectedFiscalYearValue()
        {
            if (string.IsNullOrEmpty(SelectedFiscalYear))
            {
                // 如果未選擇財年，使用當前財年
                var currentDate = DateTime.Now;
                return currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;
            }

            if (int.TryParse(SelectedFiscalYear.Replace("FY", ""), out int result))
            {
                return result;
            }

            // 如果解析失敗，使用當前財年
            var date = DateTime.Now;
            return date.Month >= 8 ? date.Year + 1 : date.Year;
        }

        // 切換標籤
        [RelayCommand]
        private async Task SwitchTab(string tabName)
        {
            if (!string.IsNullOrEmpty(tabName) && SelectedTab != tabName)
            {
                SelectedTab = tabName;

                // 載入相應標籤的數據
                if (!_isInitialized) return; // 如果尚未初始化完成，不進行數據載入

                if (tabName == "IndividualTarget" || tabName == "LOBTargets")
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
                        Debug.WriteLine($"切換到標籤 {tabName} 時載入數據出錯: {ex.Message}");
                    }
                    finally
                    {
                        IsLoading = false;
                    }
                }
            }
        }

        // 根據季度目標更新年度目標
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

        // 將年度目標平均分配到季度
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

        // 保存更改
        [RelayCommand]
        private async Task SaveChanges()
        {
            try
            {
                IsLoading = true;

                // 關閉編輯模式
                IsEditingTargets = false;

                // 保存公司目標
                await SaveCompanyTargetsAsync();

                // 保存銷售代表目標
                await SaveSalesRepTargetsAsync();

                // 保存LOB目標
                await SaveLOBTargetsAsync();

                // 通知目標服務目標已更新
                _targetService.NotifyTargetsUpdated();

                // 顯示成功消息
                ShowStatusMessage("所有目標設定已成功儲存。", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存更改時出錯: {ex.Message}");
                ShowStatusMessage($"儲存目標設定失敗: {ex.Message}", false);
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
                // 建立 TargetSettings 物件
                var targetSettings = new TargetSettings
                {
                    CompanyTargets = CompanyTargets.ToList()
                };

                // 獲取 appsettings.json 文件路徑
                string settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");

                // 讀取現有文件內容
                string json = await File.ReadAllTextAsync(settingsPath);

                // 反序列化為字典
                var jsonDoc = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);

                // 序列化更新後的 TargetSettings
                var targetSettingsJson = JsonSerializer.Serialize(targetSettings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // 更新 jsonDoc 中的 TargetSettings 節點
                var targetSettingsElement = JsonSerializer.Deserialize<JsonElement>(targetSettingsJson);
                jsonDoc["TargetSettings"] = targetSettingsElement;

                // 重新序列化為 JSON
                var updatedJson = JsonSerializer.Serialize(jsonDoc, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // 寫回文件
                await File.WriteAllTextAsync(settingsPath, updatedJson);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存公司目標時出錯: {ex.Message}");
                throw; // 重新拋出異常，由調用者捕獲
            }
        }

        private async Task SaveSalesRepTargetsAsync()
        {
            if (SalesRepTargets == null || !SalesRepTargets.Any())
                return;

            try
            {
                // 設置當前財年
                int currentFY = GetSelectedFiscalYearValue();
                foreach (var target in SalesRepTargets)
                {
                    target.FiscalYear = currentFY;
                }

                // 創建目錄（如果不存在）
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");
                Directory.CreateDirectory(targetDir);

                // 保存到文件
                string filePath = Path.Combine(targetDir, $"SalesRepTargets_{currentFY}.json");
                string json = JsonSerializer.Serialize(SalesRepTargets, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                await File.WriteAllTextAsync(filePath, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存銷售代表目標時出錯: {ex.Message}");
                throw; // 重新拋出異常，由調用者捕獲
            }
        }

        private async Task SaveLOBTargetsAsync()
        {
            if (LobTargets == null || !LobTargets.Any())
                return;

            try
            {
                // 設置當前財年
                int currentFY = GetSelectedFiscalYearValue();
                foreach (var target in LobTargets)
                {
                    target.FiscalYear = currentFY;
                }

                // 創建目錄（如果不存在）
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");
                Directory.CreateDirectory(targetDir);

                // 保存到文件
                string filePath = Path.Combine(targetDir, $"LOBTargets_{currentFY}.json");
                string json = JsonSerializer.Serialize(LobTargets, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                await File.WriteAllTextAsync(filePath, json);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存LOB目標時出錯: {ex.Message}");
                throw; // 重新拋出異常，由調用者捕獲
            }
        }

        // 添加新財年
        [RelayCommand]
        private void AddFiscalYear()
        {
            try
            {
                // 找出最高財年
                int maxYear = CompanyTargets.Max(t => t.FiscalYear);

                // 創建新財年，複製上一年的值並增加5%
                var previousYear = CompanyTargets.FirstOrDefault(t => t.FiscalYear == maxYear);
                var newTarget = new FiscalYearTarget
                {
                    FiscalYear = maxYear + 1,
                    AnnualTarget = previousYear?.AnnualTarget * 1.05m ?? 5000000m,
                    Q1Target = previousYear?.Q1Target * 1.05m ?? 1250000m,
                    Q2Target = previousYear?.Q2Target * 1.05m ?? 1250000m,
                    Q3Target = previousYear?.Q3Target * 1.05m ?? 1250000m,
                    Q4Target = previousYear?.Q4Target * 1.05m ?? 1250000m
                };

                // 四捨五入到最接近的10,000
                newTarget.AnnualTarget = Math.Round(newTarget.AnnualTarget / 10000) * 10000;
                newTarget.Q1Target = Math.Round(newTarget.Q1Target / 10000) * 10000;
                newTarget.Q2Target = Math.Round(newTarget.Q2Target / 10000) * 10000;
                newTarget.Q3Target = Math.Round(newTarget.Q3Target / 10000) * 10000;
                newTarget.Q4Target = Math.Round(newTarget.Q4Target / 10000) * 10000;

                // 將新財年添加到集合開頭
                CompanyTargets.Insert(0, newTarget);

                // 更新財年列表
                SetupFiscalYears();

                ShowStatusMessage($"已新增財年 FY{maxYear + 1}", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"新增財年時出錯: {ex.Message}");
                ShowStatusMessage($"新增財年時出錯: {ex.Message}", false);
            }
        }

        // 添加銷售代表
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

                // 四捨五入到最接近的50,000
                avgTarget = Math.Round(avgTarget / 50000) * 50000;
                decimal quarterlyTarget = avgTarget / 4;

                // 創建新銷售代表目標
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

                ShowStatusMessage("已添加新銷售代表，請更改代表名稱並設定目標值", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"添加代表時出錯: {ex.Message}");
                ShowStatusMessage($"添加代表失敗: {ex.Message}", false);
            }
        }

        // 移除銷售代表
        [RelayCommand]
        private void RemoveSalesRep(SalesRepTarget repTarget)
        {
            if (repTarget != null)
            {
                SalesRepTargets.Remove(repTarget);
                ShowStatusMessage($"已移除銷售代表 {repTarget.SalesRep}", true);
            }
        }

        // 添加LOB
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

                // 四捨五入到最接近的50,000
                avgTarget = Math.Round(avgTarget / 50000) * 50000;
                decimal quarterlyTarget = avgTarget / 4;

                // 創建新LOB目標
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

                ShowStatusMessage("已添加新產品線，請更改名稱並設定目標值", true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"添加LOB時出錯: {ex.Message}");
                ShowStatusMessage($"添加產品線失敗: {ex.Message}", false);
            }
        }

        // 移除LOB
        [RelayCommand]
        private void RemoveLOB(LOBTarget lobTarget)
        {
            if (lobTarget != null)
            {
                LobTargets.Remove(lobTarget);
                ShowStatusMessage($"已移除產品線 {lobTarget.LOB}", true);
            }
        }

        // 當選擇的財年變更時
        partial void OnSelectedFiscalYearChanged(string value)
        {
            if (string.IsNullOrEmpty(value) || !_isInitialized)
                return;

            Debug.WriteLine($"所選財年變更為: {value}");

            // 為選擇的財年載入目標
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
                    Debug.WriteLine($"載入財年 {value} 的目標時出錯: {ex.Message}");
                    ShowStatusMessage($"載入目標時出錯: {ex.Message}", false);
                }
                finally
                {
                    IsLoading = false;
                }
            });
        }

        // 顯示狀態消息
        private void ShowStatusMessage(string message, bool isSuccess)
        {
            StatusMessage = message;
            IsStatusSuccess = isSuccess;
            IsStatusMessageVisible = true;

            // 延遲後隱藏消息
            MainThread.BeginInvokeOnMainThread(async () =>
            {
                await Task.Delay(5000);
                IsStatusMessageVisible = false;
            });
        }
    }
}