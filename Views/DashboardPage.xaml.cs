using ScoreCard.Views;
using ScoreCard.ViewModels;

namespace ScoreCard.Views
{
    public partial class DashboardPage : ContentPage
    {
        public DashboardPage(DashboardViewModel viewModel)
        {
            InitializeComponent();
            BindingContext = viewModel;
        }

        protected override void OnDisappearing()
        {
            base.OnDisappearing();
            (BindingContext as DashboardViewModel)?.Cleanup();
        }

    }
}
