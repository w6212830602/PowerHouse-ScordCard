using System.Diagnostics;
using ScoreCard.ViewModels;

namespace ScoreCard.Views
{
    public partial class DetailedSalesPage : ContentPage
    {
        public DetailedSalesPage(ViewModels.DetailedSalesViewModel viewModel)
        {
            InitializeComponent();
            BindingContext = viewModel;

            // 處理日期範圍變更事件
            this.Loaded += (s, e) => {
                Debug.WriteLine("DetailedSalesPage 已加載");
            };
        }
    }
}