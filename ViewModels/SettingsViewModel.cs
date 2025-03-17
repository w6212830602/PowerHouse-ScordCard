using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
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

        public SettingsViewModel(IExcelService excelService)
        {
            _excelService = excelService;
            InitializeAsync();
        }

        private async void InitializeAsync()
        {
            try
            {
                IsLoading = true;

                // Load company targets from settings
                LoadCompanyTargetsFromSettings();

                // Setup fiscal years
                SetupFiscalYears();

                // Load sales reps and LOB data from Excel service
                await LoadSalesRepTargets();
                await LoadLOBTargets();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing SettingsViewModel: {ex.Message}");
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
                // Read from appsettings.json
                var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

                var targetSettings = config.GetSection("TargetSettings").Get<TargetSettings>();

                if (targetSettings?.CompanyTargets != null)
                {
                    CompanyTargets = new ObservableCollection<FiscalYearTarget>(
                        targetSettings.CompanyTargets.OrderByDescending(t => t.FiscalYear));
                }
                else
                {
                    // Default targets if none found in settings
                    CompanyTargets = new ObservableCollection<FiscalYearTarget>
                    {
                        new FiscalYearTarget
                        {
                            FiscalYear = 2025,
                            AnnualTarget = 10000000,
                            Q1Target = 850000,
                            Q2Target = 650000,
                            Q3Target = 1500000,
                            Q4Target = 1500000
                        },
                        new FiscalYearTarget
                        {
                            FiscalYear = 2024,
                            AnnualTarget = 10000000,
                            Q1Target = 720000,
                            Q2Target = 720000,
                            Q3Target = 720000,
                            Q4Target = 720000
                        },
                        new FiscalYearTarget
                        {
                            FiscalYear = 2023,
                            AnnualTarget = 10000000,
                            Q1Target = 650000,
                            Q2Target = 580000,
                            Q3Target = 1230000,
                            Q4Target = 1230000
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading company targets: {ex.Message}");
            }
        }

        private async Task LoadSalesRepTargets()
        {
            try
            {
                var leaderboardData = _excelService.GetSalesLeaderboardData();

                // Create sales rep targets
                SalesRepTargets = new ObservableCollection<SalesRepTarget>();

                if (leaderboardData.Any())
                {
                    // Use existing data
                    foreach (var item in leaderboardData)
                    {
                        SalesRepTargets.Add(new SalesRepTarget
                        {
                            SalesRep = item.SalesRep,
                            AnnualTarget = 10000000, // Default values 
                            Q1Target = 850000,
                            Q2Target = 650000,
                            Q3Target = 1500000,
                            Q4Target = 1500000
                        });
                    }
                }
                else
                {
                    // Add default sample data
                    SalesRepTargets.Add(new SalesRepTarget
                    {
                        SalesRep = "Brandon",
                        AnnualTarget = 10000000,
                        Q1Target = 850000,
                        Q2Target = 650000,
                        Q3Target = 1500000,
                        Q4Target = 1500000
                    });

                    SalesRepTargets.Add(new SalesRepTarget
                    {
                        SalesRep = "Chris",
                        AnnualTarget = 10000000,
                        Q1Target = 720000,
                        Q2Target = 720000,
                        Q3Target = 720000,
                        Q4Target = 720000
                    });

                    SalesRepTargets.Add(new SalesRepTarget
                    {
                        SalesRep = "Tania",
                        AnnualTarget = 10000000,
                        Q1Target = 650000,
                        Q2Target = 580000,
                        Q3Target = 1230000,
                        Q4Target = 1230000
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading sales rep targets: {ex.Message}");
            }
        }

        private async Task LoadLOBTargets()
        {
            try
            {
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
                            AnnualTarget = item.MarginTarget * 10, // Just an example multiplier
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
                        AnnualTarget = 10000000,
                        Q1Target = 850000,
                        Q2Target = 650000,
                        Q3Target = 1500000,
                        Q4Target = 1500000
                    });

                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Thermal",
                        AnnualTarget = 10000000,
                        Q1Target = 720000,
                        Q2Target = 720000,
                        Q3Target = 720000,
                        Q4Target = 720000
                    });

                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Channel",
                        AnnualTarget = 10000000,
                        Q1Target = 650000,
                        Q2Target = 580000,
                        Q3Target = 1230000,
                        Q4Target = 1230000
                    });

                    LobTargets.Add(new LOBTarget
                    {
                        LOB = "Service",
                        AnnualTarget = 40000000,
                        Q1Target = 10000000,
                        Q2Target = 10000000,
                        Q3Target = 10000000,
                        Q4Target = 10000000
                    });
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading LOB targets: {ex.Message}");
            }
        }

        // Tab switching command
        [RelayCommand]
        private void SwitchTab(string tabName)
        {
            if (!string.IsNullOrEmpty(tabName) && SelectedTab != tabName)
            {
                SelectedTab = tabName;
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

                // In a real app, you would implement saving to settings/database here
                // This example will update the appsettings.json file
                await Task.Delay(500); // Simulate save operation

                // Update TargetSettings object
                var targetSettings = new TargetSettings
                {
                    CompanyTargets = CompanyTargets.ToList()
                };

                // Save TargetSettings to appsettings.json
                // This would require file I/O which is limited in this example

                // Show confirmation
                await Application.Current.MainPage.DisplayAlert(
                    "Success",
                    "Target settings have been saved successfully.",
                    "OK");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving changes: {ex.Message}");
                await Application.Current.MainPage.DisplayAlert(
                    "Error",
                    $"Failed to save target settings. Error: {ex.Message}",
                    "OK");
            }
            finally
            {
                IsLoading = false;
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

                // Add a new fiscal year with the next year
                CompanyTargets.Insert(0, new FiscalYearTarget
                {
                    FiscalYear = maxYear + 1,
                    AnnualTarget = 10000000,
                    Q1Target = 2500000,
                    Q2Target = 2500000,
                    Q3Target = 2500000,
                    Q4Target = 2500000
                });

                // Update fiscal years list
                SetupFiscalYears();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error adding fiscal year: {ex.Message}");
            }
        }

        // Add sales rep command
        [RelayCommand]
        private void AddSalesRep()
        {
            SalesRepTargets.Add(new SalesRepTarget
            {
                SalesRep = "New Rep",
                AnnualTarget = 10000000,
                Q1Target = 2500000,
                Q2Target = 2500000,
                Q3Target = 2500000,
                Q4Target = 2500000
            });
        }

        // Remove sales rep command
        [RelayCommand]
        private void RemoveSalesRep(SalesRepTarget repTarget)
        {
            if (repTarget != null)
            {
                SalesRepTargets.Remove(repTarget);
            }
        }

        // Add LOB command
        [RelayCommand]
        private void AddLOB()
        {
            LobTargets.Add(new LOBTarget
            {
                LOB = "New Line of Business",
                AnnualTarget = 10000000,
                Q1Target = 2500000,
                Q2Target = 2500000,
                Q3Target = 2500000,
                Q4Target = 2500000
            });
        }

        // Remove LOB command
        [RelayCommand]
        private void RemoveLOB(LOBTarget lobTarget)
        {
            if (lobTarget != null)
            {
                LobTargets.Remove(lobTarget);
            }
        }

        // Change fiscal year for individual and LOB targets
        partial void OnSelectedFiscalYearChanged(string value)
        {
            if (string.IsNullOrEmpty(value))
                return;

            // You could implement logic here to load targets for the selected fiscal year
            Debug.WriteLine($"Selected fiscal year changed to: {value}");

            // In a real app, you would load targets from a database or settings for this fiscal year
            // For this example, we'll just keep using the existing data
        }
    }

    // Models for the settings page
    public class SalesRepTarget
    {
        public string SalesRep { get; set; }
        public decimal AnnualTarget { get; set; }
        public decimal Q1Target { get; set; }
        public decimal Q2Target { get; set; }
        public decimal Q3Target { get; set; }
        public decimal Q4Target { get; set; }
    }

    public class LOBTarget
    {
        public string LOB { get; set; }
        public decimal AnnualTarget { get; set; }
        public decimal Q1Target { get; set; }
        public decimal Q2Target { get; set; }
        public decimal Q3Target { get; set; }
        public decimal Q4Target { get; set; }
    }
}