using System.Diagnostics;
using Microsoft.Maui.Controls;

namespace ScoreCard.Views
{
    public partial class SalesAnalysisPage : ContentPage
    {
        public SalesAnalysisPage(ViewModels.SalesAnalysisViewModel viewModel)
        {
            InitializeComponent();
            BindingContext = viewModel;

            // 新增頁面載入完成事件
            this.Loaded += SalesAnalysisPage_Loaded;

            // 直接修改 ViewModel 屬性，然後調用重載
            DateRangePicker.StartDateChanged += async (s, e) => {
                Debug.WriteLine($"Page捕獲到StartDate變更: {e:yyyy-MM-dd}");
                // 直接設置 ViewModel 的日期 - 這樣可以確保欄位值更新
                viewModel.StartDate = e;
                await Task.Delay(100); // 給 UI 和繫結一點時間更新
                await viewModel.ReloadDataAsync();
            };

            DateRangePicker.EndDateChanged += async (s, e) => {
                Debug.WriteLine($"Page捕獲到EndDate變更: {e:yyyy-MM-dd}");
                // 直接設置 ViewModel 的日期
                viewModel.EndDate = e;
                await Task.Delay(100); // 給 UI 和繫結一點時間更新
                await viewModel.ReloadDataAsync();
            };
        }
        protected override void OnAppearing()
        {
            base.OnAppearing();
            Debug.WriteLine($"SalesAnalysisPage OnAppearing - BindingContext 是否為 null: {BindingContext == null}");
            if (BindingContext is ViewModels.SalesAnalysisViewModel vm)
            {
                Debug.WriteLine($"ViewModel 的圖表數據: {vm.TargetVsAchievementData?.Count ?? 0} 項");
            }
        }


        private async void SalesAnalysisPage_Loaded(object sender, EventArgs e)
        {
            // 延遲初始化圖表數據，給 UI 和控件足夠的時間初始化
            await Task.Delay(100);
            var vm = BindingContext as ViewModels.SalesAnalysisViewModel;
            await vm?.ReloadDataAsync();
        }

    }
}