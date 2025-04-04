﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using ScoreCard.Models;
using ScoreCard.ViewModels;

namespace ScoreCard.Services
{
    public interface ITargetService
    {
        Task InitializeAsync();

        // Company target methods
        FiscalYearTarget GetCompanyTarget(int fiscalYear);
        decimal GetCompanyQuarterlyTarget(int fiscalYear, int quarter);

        // Sales rep target methods
        List<SalesRepTarget> GetSalesRepTargets(int fiscalYear);
        decimal GetSalesRepTarget(int fiscalYear, string salesRep);
        decimal GetSalesRepQuarterlyTarget(int fiscalYear, string salesRep, int quarter);

        // LOB target methods
        List<LOBTarget> GetLOBTargets(int fiscalYear);
        decimal GetLOBTarget(int fiscalYear, string lob);
        decimal GetLOBQuarterlyTarget(int fiscalYear, string lob, int quarter);

        // File monitoring and update notifications
        void MonitorTargetFiles();
        void NotifyTargetsUpdated();

        // Events
        event EventHandler TargetsUpdated;

        Task<bool> UpdateCompanyTargetAsync(FiscalYearTarget target);
    }

    public class TargetService : ITargetService
    {
        private readonly IConfiguration _configuration;
        private readonly Dictionary<int, List<SalesRepTarget>> _salesRepTargetsByYear = new();
        private readonly Dictionary<int, List<LOBTarget>> _lobTargetsByYear = new();
        private List<FiscalYearTarget> _companyTargets = new();
        private bool _isInitialized = false;
        private FileSystemWatcher _watcher;

        public event EventHandler TargetsUpdated;

        public TargetService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public async Task<bool> UpdateCompanyTargetAsync(FiscalYearTarget target)
        {
            try
            {
                Debug.WriteLine($"Updating company target for fiscal year {target.FiscalYear}");

                // Find and update the target in the _companyTargets list
                var existingTarget = _companyTargets.FirstOrDefault(t => t.FiscalYear == target.FiscalYear);
                if (existingTarget != null)
                {
                    // Update existing target
                    existingTarget.AnnualTarget = target.AnnualTarget;
                    existingTarget.Q1Target = target.Q1Target;
                    existingTarget.Q2Target = target.Q2Target;
                    existingTarget.Q3Target = target.Q3Target;
                    existingTarget.Q4Target = target.Q4Target;
                }
                else
                {
                    // Add new target
                    _companyTargets.Add(target);
                }

                // Save to appsettings.json
                string settingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");

                // Read existing file content
                string json = await File.ReadAllTextAsync(settingsPath);

                // Deserialize to dictionary
                var jsonDoc = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);

                // Update TargetSettings
                var targetSettings = new TargetSettings
                {
                    CompanyTargets = _companyTargets.ToList()
                };

                // Serialize updated TargetSettings
                var targetSettingsJson = JsonSerializer.Serialize(targetSettings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Update jsonDoc with new TargetSettings
                var targetSettingsElement = JsonSerializer.Deserialize<JsonElement>(targetSettingsJson);
                jsonDoc["TargetSettings"] = targetSettingsElement;

                // Reserialize to JSON
                var updatedJson = JsonSerializer.Serialize(jsonDoc, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Write back to file
                await File.WriteAllTextAsync(settingsPath, updatedJson);

                // Notify that targets have been updated
                NotifyTargetsUpdated();

                Debug.WriteLine($"Company target for fiscal year {target.FiscalYear} updated successfully");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating company target: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);
                return false;
            }
        }

        public async Task InitializeAsync()
        {
            if (_isInitialized)
                return;

            try
            {
                Debug.WriteLine("TargetService: Starting initialization");

                // Load company targets with error handling
                try
                {
                    LoadCompanyTargetsFromSettings();
                    Debug.WriteLine("TargetService: Company targets loaded successfully");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"TargetService: Error loading company targets: {ex.Message}");
                    // Create default company targets
                    _companyTargets = CreateDefaultCompanyTargets();
                }

                _isInitialized = true;
                Debug.WriteLine("TargetService: Initialization completed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TargetService: Critical error in initialization: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // Ensure we mark as initialized even on error
                _isInitialized = true;
                throw; // Re-throw to be handled by caller
            }
        }

        private List<FiscalYearTarget> CreateDefaultCompanyTargets()
        {
            var currentDate = DateTime.Now;
            var currentFiscalYear = currentDate.Month >= 8 ? currentDate.Year + 1 : currentDate.Year;

            return new List<FiscalYearTarget>
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



        private void LoadCompanyTargetsFromSettings()
        {
            try
            {
                var targetSettings = _configuration.GetSection("TargetSettings").Get<TargetSettings>();

                if (targetSettings?.CompanyTargets != null && targetSettings.CompanyTargets.Any())
                {
                    _companyTargets = targetSettings.CompanyTargets.ToList();
                }
                else
                {
                    // Default targets if none found in settings
                    var currentYear = DateTime.Now.Year;
                    var currentFiscalYear = DateTime.Now.Month >= 8 ? currentYear + 1 : currentYear;

                    _companyTargets = new List<FiscalYearTarget>
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

                Debug.WriteLine($"Loaded {_companyTargets.Count} company targets from settings");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading company targets: {ex.Message}");
                throw;
            }
        }

        private void LoadSalesRepTargetsFromFile(int fiscalYear)
        {
            try
            {
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");
                string filePath = Path.Combine(targetDir, $"SalesRepTargets_{fiscalYear}.json");

                if (!File.Exists(filePath))
                {
                    Debug.WriteLine($"No sales rep targets file found for fiscal year {fiscalYear}");
                    _salesRepTargetsByYear[fiscalYear] = new List<SalesRepTarget>();
                    return;
                }

                string json = File.ReadAllText(filePath);
                var targets = JsonSerializer.Deserialize<List<SalesRepTarget>>(json);

                if (targets != null && targets.Any())
                {
                    _salesRepTargetsByYear[fiscalYear] = targets;
                    Debug.WriteLine($"Loaded {targets.Count} sales rep targets for fiscal year {fiscalYear}");
                }
                else
                {
                    _salesRepTargetsByYear[fiscalYear] = new List<SalesRepTarget>();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading sales rep targets for fiscal year {fiscalYear}: {ex.Message}");
                _salesRepTargetsByYear[fiscalYear] = new List<SalesRepTarget>();
            }
        }

        private void LoadLOBTargetsFromFile(int fiscalYear)
        {
            try
            {
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");
                string filePath = Path.Combine(targetDir, $"LOBTargets_{fiscalYear}.json");

                if (!File.Exists(filePath))
                {
                    Debug.WriteLine($"No LOB targets file found for fiscal year {fiscalYear}");
                    _lobTargetsByYear[fiscalYear] = new List<LOBTarget>();
                    return;
                }

                string json = File.ReadAllText(filePath);
                var targets = JsonSerializer.Deserialize<List<LOBTarget>>(json);

                if (targets != null && targets.Any())
                {
                    _lobTargetsByYear[fiscalYear] = targets;
                    Debug.WriteLine($"Loaded {targets.Count} LOB targets for fiscal year {fiscalYear}");
                }
                else
                {
                    _lobTargetsByYear[fiscalYear] = new List<LOBTarget>();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading LOB targets for fiscal year {fiscalYear}: {ex.Message}");
                _lobTargetsByYear[fiscalYear] = new List<LOBTarget>();
            }
        }

        public FiscalYearTarget GetCompanyTarget(int fiscalYear)
        {
            var target = _companyTargets.FirstOrDefault(t => t.FiscalYear == fiscalYear);

            if (target == null)
            {
                // If no target found, use the most recent available target
                if (_companyTargets.Any())
                {
                    var mostRecentYear = _companyTargets.Max(t => t.FiscalYear);
                    var nearestYear = _companyTargets
                        .OrderBy(t => Math.Abs(t.FiscalYear - fiscalYear))
                        .First().FiscalYear;

                    target = _companyTargets.First(t => t.FiscalYear == nearestYear);
                    Debug.WriteLine($"No company target found for fiscal year {fiscalYear}, using {nearestYear} instead");
                }
                else
                {
                    // Create a default target if we don't have any targets at all
                    target = new FiscalYearTarget
                    {
                        FiscalYear = fiscalYear,
                        AnnualTarget = 4000000,
                        Q1Target = 1000000,
                        Q2Target = 1000000,
                        Q3Target = 1000000,
                        Q4Target = 1000000
                    };
                    Debug.WriteLine($"No company targets found, created default target for {fiscalYear}");
                }
            }

            return target;
        }

        public decimal GetCompanyQuarterlyTarget(int fiscalYear, int quarter)
        {
            if (quarter < 1 || quarter > 4)
                throw new ArgumentOutOfRangeException(nameof(quarter), "Quarter must be between 1 and 4");

            var target = GetCompanyTarget(fiscalYear);

            return quarter switch
            {
                1 => target.Q1Target,
                2 => target.Q2Target,
                3 => target.Q3Target,
                4 => target.Q4Target,
                _ => 0 // Never reached due to the check above
            };
        }

        public List<SalesRepTarget> GetSalesRepTargets(int fiscalYear)
        {
            if (!_salesRepTargetsByYear.ContainsKey(fiscalYear))
            {
                LoadSalesRepTargetsFromFile(fiscalYear);
            }

            return _salesRepTargetsByYear[fiscalYear];
        }

        public decimal GetSalesRepTarget(int fiscalYear, string salesRep)
        {
            var targets = GetSalesRepTargets(fiscalYear);
            var target = targets.FirstOrDefault(t => t.SalesRep.Equals(salesRep, StringComparison.OrdinalIgnoreCase));

            if (target != null)
            {
                return target.AnnualTarget;
            }
            else
            {
                // If no specific target is set for this rep, calculate a reasonable default
                // based on the company target and the number of sales reps
                var companyTarget = GetCompanyTarget(fiscalYear);
                int repCount = targets.Count > 0 ? targets.Count : 5; // Default to 5 reps if none defined

                return companyTarget.AnnualTarget / repCount;
            }
        }

        public decimal GetSalesRepQuarterlyTarget(int fiscalYear, string salesRep, int quarter)
        {
            if (quarter < 1 || quarter > 4)
                throw new ArgumentOutOfRangeException(nameof(quarter), "Quarter must be between 1 and 4");

            var targets = GetSalesRepTargets(fiscalYear);
            var target = targets.FirstOrDefault(t => t.SalesRep.Equals(salesRep, StringComparison.OrdinalIgnoreCase));

            if (target != null)
            {
                return quarter switch
                {
                    1 => target.Q1Target,
                    2 => target.Q2Target,
                    3 => target.Q3Target,
                    4 => target.Q4Target,
                    _ => 0 // Never reached due to the check above
                };
            }
            else
            {
                // If no specific target is set for this rep, calculate a reasonable default
                var annualTarget = GetSalesRepTarget(fiscalYear, salesRep);
                return annualTarget / 4; // Divide evenly across quarters
            }
        }

        public List<LOBTarget> GetLOBTargets(int fiscalYear)
        {
            if (!_lobTargetsByYear.ContainsKey(fiscalYear))
            {
                LoadLOBTargetsFromFile(fiscalYear);
            }

            return _lobTargetsByYear[fiscalYear];
        }

        public decimal GetLOBTarget(int fiscalYear, string lob)
        {
            var targets = GetLOBTargets(fiscalYear);
            var target = targets.FirstOrDefault(t => t.LOB.Equals(lob, StringComparison.OrdinalIgnoreCase));

            if (target != null)
            {
                return target.AnnualTarget;
            }
            else
            {
                // If no specific target is set for this LOB, calculate a reasonable default
                // based on the company target
                var companyTarget = GetCompanyTarget(fiscalYear);
                int lobCount = targets.Count > 0 ? targets.Count : 4; // Default to 4 LOBs if none defined

                return companyTarget.AnnualTarget / lobCount;
            }
        }

        public decimal GetLOBQuarterlyTarget(int fiscalYear, string lob, int quarter)
        {
            if (quarter < 1 || quarter > 4)
                throw new ArgumentOutOfRangeException(nameof(quarter), "Quarter must be between 1 and 4");

            var targets = GetLOBTargets(fiscalYear);
            var target = targets.FirstOrDefault(t => t.LOB.Equals(lob, StringComparison.OrdinalIgnoreCase));

            if (target != null)
            {
                return quarter switch
                {
                    1 => target.Q1Target,
                    2 => target.Q2Target,
                    3 => target.Q3Target,
                    4 => target.Q4Target,
                    _ => 0 // Never reached due to the check above
                };
            }
            else
            {
                // If no specific target is set for this LOB, calculate a reasonable default
                var annualTarget = GetLOBTarget(fiscalYear, lob);
                return annualTarget / 4; // Divide evenly across quarters
            }
        }

        // Monitor target files for changes
        public void MonitorTargetFiles()
        {
            try
            {
                string targetDir = Path.Combine(AppContext.BaseDirectory, "Targets");

                // Create directory if it doesn't exist
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                // Setup file watcher
                _watcher = new FileSystemWatcher(targetDir)
                {
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName,
                    Filter = "*.json",
                    EnableRaisingEvents = true
                };

                _watcher.Changed += OnTargetFileChanged;
                _watcher.Created += OnTargetFileChanged;
                _watcher.Renamed += OnTargetFileChanged;

                Debug.WriteLine($"Started monitoring target files in {targetDir}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error setting up file monitoring: {ex.Message}");
            }
        }

        private void OnTargetFileChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                Debug.WriteLine($"Target file changed: {e.Name}, {e.ChangeType}");

                // Allow the file to be completely written
                Task.Delay(500).Wait();

                // Reload the affected file
                if (e.Name.StartsWith("SalesRepTargets_"))
                {
                    string yearStr = e.Name.Replace("SalesRepTargets_", "").Replace(".json", "");
                    if (int.TryParse(yearStr, out int year))
                    {
                        LoadSalesRepTargetsFromFile(year);
                    }
                }
                else if (e.Name.StartsWith("LOBTargets_"))
                {
                    string yearStr = e.Name.Replace("LOBTargets_", "").Replace(".json", "");
                    if (int.TryParse(yearStr, out int year))
                    {
                        LoadLOBTargetsFromFile(year);
                    }
                }

                // Notify that targets have been updated
                NotifyTargetsUpdated();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing target file change: {ex.Message}");
            }
        }

        // Method to notify other components when targets are updated
        public void NotifyTargetsUpdated()
        {
            try
            {
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    TargetsUpdated?.Invoke(this, EventArgs.Empty);
                    Debug.WriteLine("Target update notification sent");
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error notifying of target updates: {ex.Message}");
            }
        }
    }
}