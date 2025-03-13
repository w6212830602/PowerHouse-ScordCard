using Microsoft.Extensions.Logging;
using ScoreCard.Views;
using ScoreCard.Services;
using ScoreCard.ViewModels;
using Microsoft.Maui.LifecycleEvents;
using Syncfusion.Maui.Core.Hosting;

namespace ScoreCard
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder
                .UseMauiApp<App>()
                .ConfigureSyncfusionCore()    // 添加 Syncfusion 配置
                .ConfigureFonts(fonts =>
                {
                    fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                    fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
                });

            // 添加視窗設定
            builder.ConfigureLifecycleEvents(events =>
            {
#if WINDOWS
                events.AddWindows(windows => windows
                    .OnWindowCreated(window =>
                    {
                        window.ExtendsContentIntoTitleBar = false;
                    }));
#endif
            });

            // Register Services
            builder.Services.AddSingleton<IExcelService, ExcelService>();
            builder.Services.AddSingleton<IExportService, ExportService>(); // 新增匯出服務

            // Register ViewModels
            builder.Services.AddTransient<DashboardViewModel>();
            builder.Services.AddTransient<SalesAnalysisViewModel>();
            builder.Services.AddTransient<SettingsViewModel>();
            builder.Services.AddTransient<DetailedSalesViewModel>();

            // Register Pages
            builder.Services.AddTransient<DashboardPage>();
            builder.Services.AddTransient<SalesAnalysisPage>();
            builder.Services.AddTransient<SettingsPage>();
            builder.Services.AddTransient<DetailedSalesPage>();


#if DEBUG
            builder.Logging.AddDebug();
#endif

            return builder.Build();
        }
    }
}