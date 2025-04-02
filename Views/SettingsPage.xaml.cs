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

        protected override async void OnAppearing()
        {
            base.OnAppearing();

            try
            {
                Debug.WriteLine("SettingsPage OnAppearing");
                await _viewModel.InitializeAsync();
                Debug.WriteLine("SettingsPage initialization completed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in SettingsPage.OnAppearing: {ex.Message}");
                Debug.WriteLine(ex.StackTrace);

                // Show error to user
                await DisplayAlert("Error", "Failed to initialize settings page. Please try again later.", "OK");
            }
        }
        protected override void OnDisappearing()
        {
            base.OnDisappearing();
            Debug.WriteLine("SettingsPage disappeared");
        }
    }
}