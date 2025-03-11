using System.Diagnostics;
using ScoreCard.ViewModels;

namespace ScoreCard.Views
{
    public partial class SettingsPage : ContentPage
    {
        private readonly SettingsViewModel _viewModel;

        public SettingsPage(SettingsViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            BindingContext = viewModel;

            // 訂閱頁面載入事件
            this.Loaded += OnPageLoaded;
        }

        private void OnPageLoaded(object sender, EventArgs e)
        {
            Debug.WriteLine("SettingsPage loaded");
        }

        protected override void OnAppearing()
        {
            base.OnAppearing();
            Debug.WriteLine("SettingsPage appeared");
        }

        protected override void OnDisappearing()
        {
            base.OnDisappearing();
            Debug.WriteLine("SettingsPage disappeared");
        }
    }
}