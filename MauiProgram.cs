using Microsoft.Extensions.Logging;
using ScoreCard.Views;
using ScoreCard.Services;
using ScoreCard.ViewModels;
using ScoreCard.Views;
using Microsoft.Maui.LifecycleEvents;


namespace ScoreCard
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            var builder = MauiApp.CreateBuilder();
            builder
                .UseMauiApp<App>()
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

            // Register ViewModels
            builder.Services.AddTransient<ViewModels.DashboardViewModel>();
            builder.Services.AddTransient<ViewModels.SalesAnalysisViewModel>();
            builder.Services.AddTransient<ViewModels.SettingsViewModel>();

            // Register Pages
            builder.Services.AddTransient<Views.DashboardPage>();
            builder.Services.AddTransient<Views.SalesAnalysisPage>();
            builder.Services.AddTransient<Views.SettingsPage>();

            return builder.Build();
        }
    }
}