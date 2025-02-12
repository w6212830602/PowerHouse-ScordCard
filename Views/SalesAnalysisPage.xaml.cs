using Microsoft.Maui.Controls;

namespace ScoreCard.Views
{
    public partial class SalesAnalysisPage : ContentPage
    {
        public SalesAnalysisPage(ViewModels.SalesAnalysisViewModel viewModel)
        {
            InitializeComponent();
            BindingContext = viewModel;
        }
    }
}