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

        // Events
        event EventHandler TargetsUpdated;
    }

    public class TargetService : ITargetService
    {
        private readonly IConfiguration _configuration;
        private readonly Dictionary<int, List<SalesRepTarget>> _salesRepTargetsByYear = new();
        private readonly Dictionary<int, List<LOBTarget>> _lobTargetsByYear = new();
        private List<FiscalYearTarget> _companyTargets = new();
        private bool _isInitialized = false;

        public event EventHandler TargetsUpdated;

        public TargetService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public async Task InitializeAsync()
        {
            if (_isInitialized)
                return;

            try
            {
                await Task.Run(() =>
                {
                    // Load company targets from appsettings.json
                    LoadCompanyTargetsFromSettings();

                    // Load sales rep and LOB targets for all fiscal years
                    foreach (var target in _companyTargets)
                    {
                        int fiscalYear = target.FiscalYear;
                        LoadSalesRepTargetsFromFile(fiscalYear);
                        LoadLOBTargetsFromFile(fiscalYear);
                    }
                });

                _isInitialized = true;
                Debug.WriteLine("Target service initialized successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error initializing target service: {ex.Message}");
                throw;
            }
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
                    _companyTargets = new List<FiscalYearTarget>
                    {
                        new FiscalYearTarget
                        {
                            FiscalYear = DateTime.Now.Year + 1,
                            AnnualTarget = 5000000,
                            Q1Target = 1250000,
                            Q2Target = 1250000,
                            Q3Target = 1250000,
                            Q4Target = 1250000
                        },
                        new FiscalYearTarget
                        {
                            FiscalYear = DateTime.Now.Year,
                            AnnualTarget = 4500000,
                            Q1Target = 1125000,
                            Q2Target = 1125000,
                            Q3Target = 1125000,
                            Q4Target = 1125000
                        },
                        new FiscalYearTarget
                        {
                            FiscalYear = DateTime.Now.Year - 1,
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
                var mostRecentYear = _companyTargets.Max(t => t.FiscalYear);
                target = _companyTargets.First(t => t.FiscalYear == mostRecentYear);
                Debug.WriteLine($"No company target found for fiscal year {fiscalYear}, using {mostRecentYear} instead");
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

        // Method to notify other components when targets are updated
        public void NotifyTargetsUpdated()
        {
            TargetsUpdated?.Invoke(this, EventArgs.Empty);
        }
    }
}