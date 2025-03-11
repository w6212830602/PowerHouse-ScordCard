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
                await Task.Delay(500); // Simulate save operation

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
                    "Failed to save target settings. Please try again.",
                    "OK");
            }
            finally
            {
                IsLoading = false;
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