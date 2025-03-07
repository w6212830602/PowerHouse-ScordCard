using System.Diagnostics;
using ScoreCard.ViewModels;
using Microsoft.Maui.Graphics;


namespace ScoreCard.Views
{
    public partial class DetailedSalesPage : ContentPage
    {
        private readonly DetailedSalesViewModel _viewModel;
        private Border _repSelectionPopup;

        public DetailedSalesPage(DetailedSalesViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            BindingContext = viewModel;

            // 處理日期範圍變更事件
            // 處理日期範圍變更事件
            this.Loaded += (s, e) => {
                Debug.WriteLine("DetailedSalesPage 已加載");

                // 獲取彈出視窗引用
                _repSelectionPopup = this.FindByName<Border>("RepSelectionPopup");

                // 獲取頁面的主要內容容器
                var mainContent = this.Content as View;
                if (mainContent != null)
                {
                    var backgroundTapGesture = new TapGestureRecognizer();
                    backgroundTapGesture.Tapped += OnBackgroundTapped;
                    mainContent.GestureRecognizers.Add(backgroundTapGesture);
                }
            };
            // 手動連接CustomDatePicker事件到我們的自定義處理方法
            var datePickers = this.FindByName<ScoreCard.Controls.CustomDatePicker>("DateRangePicker");
            if (datePickers != null)
            {
                datePickers.StartDateChanged += OnStartDateChanged;
                datePickers.EndDateChanged += OnEndDateChanged;
            }
            else
            {
                Debug.WriteLine("警告：無法找到DateRangePicker控件");
            }
        }

        // 處理開始日期變更的方法
        private void OnStartDateChanged(object sender, DateTime e)
        {
            Debug.WriteLine($"DetailedSalesPage捕獲到開始日期變更：{e:yyyy-MM-dd}");

            // 直接設置視圖模型的屬性值
            if (_viewModel != null)
            {
                _viewModel.StartDate = e;

                // 刷新數據命令
                Task.Run(async () => {
                    try
                    {
                        await Task.Delay(100); // 給數據綁定一點時間更新
                        await _viewModel.FilterDataCommand.ExecuteAsync(null);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"更新開始日期時發生錯誤: {ex.Message}");
                    }
                });
            }
        }

        // 處理結束日期變更的方法
        private void OnEndDateChanged(object sender, DateTime e)
        {
            Debug.WriteLine($"DetailedSalesPage捕獲到結束日期變更：{e:yyyy-MM-dd}");

            // 直接設置視圖模型的屬性值
            if (_viewModel != null)
            {
                _viewModel.EndDate = e;

                // 刷新數據命令
                Task.Run(async () => {
                    try
                    {
                        await Task.Delay(100); // 給數據綁定一點時間更新
                        await _viewModel.FilterDataCommand.ExecuteAsync(null);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"更新結束日期時發生錯誤: {ex.Message}");
                    }
                });
            }
        }

        // 處理背景點擊，用於關閉彈出視窗
        // 修改 OnBackgroundTapped 方法
        private void OnBackgroundTapped(object sender, EventArgs e)
        {
            if (_viewModel != null && _viewModel.IsRepSelectionPopupVisible)
            {
                // 確保點擊的 sender 是一個 View
                if (e is TappedEventArgs tappedEvent && sender is View)
                {
                    var touchPoint = (Microsoft.Maui.Graphics.Point)tappedEvent.GetPosition((View)sender);

                    // 如果彈出視窗存在且可見
                    if (_repSelectionPopup != null && _repSelectionPopup.IsVisible)
                    {
                        Rect popupBounds = new Rect(
                            _repSelectionPopup.X,
                            _repSelectionPopup.Y,
                            _repSelectionPopup.Width,
                            _repSelectionPopup.Height);

                        // 如果點擊在彈出視窗外，關閉彈出視窗
                        if (!popupBounds.Contains(touchPoint))
                        {
                            Debug.WriteLine("背景點擊，關閉彈出視窗");
                            _viewModel.CloseRepSelectionPopupCommand.Execute(null);
                        }
                    }
                }
                else
                {
                    // 如果無法獲取點擊坐標，默認關閉彈出視窗
                    _viewModel.CloseRepSelectionPopupCommand.Execute(null);
                }
            }
        }

        protected override void OnAppearing()
        {
            base.OnAppearing();
            Debug.WriteLine("DetailedSalesPage 出現");

            // 頁面顯示時也重新加載一次數據
            Task.Run(async () => {
                try
                {
                    await Task.Delay(300); // 等待UI完成加載
                    await _viewModel?.FilterDataCommand.ExecuteAsync(null);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"頁面加載時重新載入數據發生錯誤: {ex.Message}");
                }
            });
        }

        protected override bool OnBackButtonPressed()
        {
            // 如果彈出視窗正在顯示，按返回鍵關閉彈出視窗
            if (_viewModel != null && _viewModel.IsRepSelectionPopupVisible)
            {
                _viewModel.CloseRepSelectionPopupCommand.Execute(null);
                return true; // 表示已處理返回事件
            }

            return base.OnBackButtonPressed();
        }
    }
}