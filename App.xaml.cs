using Microsoft.Extensions.DependencyInjection;
using ScoreCard.Services;
using ScoreCard.ViewModels;
using ScoreCard.Views;
using Microsoft.Extensions.Logging;
using System.Diagnostics;


namespace ScoreCard
{
    public partial class App : Application
    {
        private readonly ILogger<App> _logger;
        private readonly IServiceProvider _serviceProvider;
        private CancellationTokenSource _cts;

        public App()
        {
            try
            {
                // 設置全局異常處理
                AppDomain.CurrentDomain.UnhandledException += (s, e) =>
                {
                    Debug.WriteLine($"未處理的異常: {e.ExceptionObject}");
                };

                InitializeComponent();
                MainPage = new AppShell();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"應用程式初始化錯誤: {ex.Message}");
            }
        }

        private void InitializeFileMonitoring()
        {
            try
            {
                var excelService = _serviceProvider.GetRequiredService<IExcelService>();
                MainThread.BeginInvokeOnMainThread(async () =>
                {
                    try
                    {
                        // 直接使用 CancellationToken
                        await excelService.MonitorFileChangesAsync(_cts.Token);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "監控檔案時發生錯誤");
                    }
                });
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "檔案監控初始化失敗");
            }
        }

        protected override void OnStart()
        {
            try
            {
                Debug.WriteLine("應用程式啟動: 正在初始化 Syncfusion 許可證...");
                Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1NMaF1cXmhNYVppR2Nbek5xdF9HZlZSTGYuP1ZhSXxWdkZiWX5ecXJRRGZaWEQ=");
                Debug.WriteLine("Syncfusion 許可證註冊成功");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Syncfusion 許可證錯誤: {ex.Message}");
                Debug.WriteLine($"堆疊跟踪: {ex.StackTrace}");
            }

            base.OnStart();
        }
        protected override Window CreateWindow(IActivationState activationState)
        {
            var window = base.CreateWindow(activationState);

            window.Created += (s, e) =>
            {
                _logger?.LogInformation("視窗已創建");
            };


            window.Destroying += (s, e) =>
            {
                try
                {
                    _logger?.LogInformation("正在清理資源");
                    _cts?.Cancel();
                    _cts?.Dispose();
                    _cts = null;
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "清理資源時發生錯誤");
                }
            };

            return window;
        }
    }
}