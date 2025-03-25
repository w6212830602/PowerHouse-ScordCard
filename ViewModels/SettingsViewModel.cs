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
                // Get the current fiscal year (parsed from the selected fiscal year string)
                int currentFiscalYear = GetSelectedFiscalYearValue();

                // Get sales rep targets from target service
                var savedTargets = _targetService.GetSalesRepTargets(currentFiscalYear);
                
                if (savedTargets?.Any() == true)
                {
                    SalesRepTargets = new ObservableCollection<SalesRepTarget>(savedTargets);
                    Debug.WriteLine($"Loaded {savedTargets.Count} sales rep targets");
                    return;
                }

                // If no saved targets, generate from leaderboard data
                var leaderboardData = _excelService.GetSalesLeaderboardData();

                // Create sales rep targets
                SalesRepTargets = new ObservableCollection<SalesRepTarget>();

                if (leaderboardData.Any())
                {
                    // Calculate an appropriate target based on current performance
                    // This is just an example logic - adjust as needed for your business rules
                    foreach (var item in leaderboardData)
                    {
                        // Use the annual target from the company targets and distribute based on rep performance
                        var companyTarget = CompanyTargets.FirstOrDefault(t => t.FiscalYear == currentFiscalYear);
                        decimal annualTarget = companyTarget?.AnnualTarget ?? 4000000;

                        // Calculate a weighted target based on the rep's ranking
                        decimal weight = 1.0m / Math.Max(1, item.Rank + 1);  // 確保除數至少為1
                        decimal repTarget = annualTarget * weight;

                        // Ensure a minimum target
                        repTarget = Math.Max(repTarget, annualTarget * 0.05m);

                        var roundedAnnualTarget = Math.Round(repTarget, -3);  // Round to nearest thousand
                        var quarterlyTarget = Math.Round(roundedAnnualTarget / 4, 0);  // Round to nearest whole number

                        SalesRepTargets.Add(new SalesRepTarget
                        {
                            SalesRep = item.SalesRep,
                            FiscalYear = currentFiscalYear,
                            AnnualTarget = roundedAnnualTarget,
                            Q1Target = quarterlyTarget,
                            Q2Target = quarterlyTarget,
                            Q3Target = quarterlyTarget,
                            Q4Target = quarterlyTarget
                        });
                    }
                }
                else
                {
                    // Add default sample data if no leaderboard data
                    SalesRepTargets.Add(new SalesRepTarget
                    {
                        SalesRep = "Brandon",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 1000000,
                        Q1Target = 250000,
                        Q2Target = 250000,
                        Q3Target = 250000,
                        Q4Target = 250000
                    });

                    SalesRepTargets.Add(new SalesRepTarget
                    {
                        SalesRep = "Chris",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 900000,
                        Q1Target = 225000,
                        Q2Target = 225000,
                        Q3Target = 225000,
                        Q4Target = 225000
                    });

                    SalesRepTargets.Add(new SalesRepTarget
                    {
                        SalesRep = "Isaac",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 850000,
                        Q1Target = 212500,
                        Q2Target = 212500,
                        Q3Target = 212500,
                        Q4Target = 212500
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading sales rep targets: {ex.Message}");
                ShowStatusMessage($"Error loading sales rep targets: {ex.Message}", false);
            }
        }

        private async Task LoadLOBTargets()
        {
            try
            {
                // Get the current fiscal year
                int currentFiscalYear = GetSelectedFiscalYearValue();

                // Get LOB targets from target service
                var savedTargets = _targetService.GetLOBTargets(currentFiscalYear);
                
                if (savedTargets?.Any() == true)
                {
                    LobTargets = new ObservableCollection<LOBTarget>(savedTargets);
                    Debug.WriteLine($"Loaded {savedTargets.Count} LOB targets");
                    return;
                }

                var deptLobData = _excelService.GetDepartmentLobData();

                LobTargets = new ObservableCollection<LOBTarget>();

                if (deptLobData.Any())
                {
                    // Use existing data, skipping "Total" entry
                    foreach (var item in deptLobData.Where(d => d.LOB != "Total"))
                    {
                        LobTargets.Add(new LOBTarget
                        {
                            LOB = item.LOB,
                            FiscalYear = currentFiscalYear,
                            AnnualTarget = item.MarginTarget,
                            Q1Target = item.MarginTarget / 4,
                            Q2Target = item.MarginTarget / 4,
                            Q3Target = item.MarginTarget / 4,
                            Q4Target = item.MarginTarget / 4
                        });
                    }
                }
                else
                {
                    // Add default sample data
                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Power",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 1000000,
                        Q1Target = 250000,
                        Q2Target = 250000,
                        Q3Target = 250000,
                        Q4Target = 250000
                    });

                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Thermal",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 900000,
                        Q1Target = 225000,
                        Q2Target = 225000,
                        Q3Target = 225000,
                        Q4Target = 225000
                    });

                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Channel",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 750000,
                        Q1Target = 187500,
                        Q2Target = 187500,
                        Q3Target = 187500,
                        Q4Target = 187500
                    });

                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Service",
                        FiscalYear = currentFiscalYear,
                        AnnualTarget = 500000,
                        Q1Target = 125000,
                        Q2Target = 125000,
                        Q3Target = 125000,
                        Q4Target = 125000
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading LOB targets: {ex.Message}`");
                ShowStatusMessage($"Error loading LOB targets: {ex.Message}", false);
            }
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
    }
}