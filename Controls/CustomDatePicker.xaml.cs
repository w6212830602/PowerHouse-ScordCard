using Microsoft.Maui.Controls;
using System;

namespace ScoreCard.Controls
{
    public partial class CustomDatePicker : ContentView
    {
        public static readonly BindableProperty StartDateProperty =
            BindableProperty.Create(nameof(StartDate), typeof(DateTime), typeof(CustomDatePicker),
                DateTime.Now, propertyChanged: OnDateChanged);

        public static readonly BindableProperty EndDateProperty =
            BindableProperty.Create(nameof(EndDate), typeof(DateTime), typeof(CustomDatePicker),
                DateTime.Now, propertyChanged: OnDateChanged);

        private DateTime _displayMonth;
        private bool _isStartDateActive;

        // 添加一個關閉按鈕或點擊外部區域來關閉日曆
        public CustomDatePicker()
        {
            InitializeComponent();
            BindingContext = this;
            _displayMonth = DateTime.Now;
            _isStartDateActive = true; // 預設是選擇開始日期

            var dateTapGesture = new TapGestureRecognizer();
            dateTapGesture.Tapped += (s, e) =>
            {
                if (!Calendar.IsVisible)
                {
                    _isStartDateActive = true; // 每次打開日曆時，都重置為選擇開始日期
                    _displayMonth = StartDate;
                }
                Calendar.IsVisible = !Calendar.IsVisible;
                UpdateCalendarDisplay();
            };
            DateDisplay.GestureRecognizers.Add(dateTapGesture);

            // 添加背景點擊事件來關閉日曆
            var backgroundTapGesture = new TapGestureRecognizer();
            backgroundTapGesture.Tapped += (s, e) =>
            {
                if (Calendar.IsVisible)
                {
                    Calendar.IsVisible = false;
                }
            };
            this.GestureRecognizers.Add(backgroundTapGesture);

            // 防止日曆區域的點擊事件影響背景
            var calendarTapGesture = new TapGestureRecognizer();
            calendarTapGesture.Tapped += (s, e) =>
            {
                // 不做任何事，只是攔截事件
            };
            Calendar.GestureRecognizers.Add(calendarTapGesture);

            UpdateCalendarDisplay();
        }

        public DateTime StartDate
        {
            get => (DateTime)GetValue(StartDateProperty);
            set => SetValue(StartDateProperty, value);
        }

        public DateTime EndDate
        {
            get => (DateTime)GetValue(EndDateProperty);
            set => SetValue(EndDateProperty, value);
        }

        private static void OnDateChanged(BindableObject bindable, object oldValue, object newValue)
        {
            var picker = (CustomDatePicker)bindable;
            picker.UpdateCalendarDisplay();
        }

        private void OnStartDateClicked(object sender, EventArgs e)
        {
            _isStartDateActive = true;
            _displayMonth = StartDate;
            Calendar.IsVisible = true;
            UpdateCalendarDisplay();
        }

        private void OnEndDateClicked(object sender, EventArgs e)
        {
            _isStartDateActive = false;
            _displayMonth = EndDate;
            Calendar.IsVisible = true;
            UpdateCalendarDisplay();
        }

        private void OnPreviousMonthClicked(object sender, EventArgs e)
        {
            _displayMonth = _displayMonth.AddMonths(-1);
            UpdateCalendarDisplay();
        }

        private void OnNextMonthClicked(object sender, EventArgs e)
        {
            _displayMonth = _displayMonth.AddMonths(1);
            UpdateCalendarDisplay();
        }

        private void UpdateCalendarDisplay()
        {
            UpdateMonthGrid(CurrentMonthGrid, _displayMonth);
            UpdateMonthGrid(NextMonthGrid, _displayMonth.AddMonths(1));

            CurrentMonthLabel.Text = _displayMonth.ToString("MMMM yyyy");
            NextMonthLabel.Text = _displayMonth.AddMonths(1).ToString("MMMM yyyy");
        }

        private void UpdateMonthGrid(Grid grid, DateTime month)
        {
            grid.Children.Clear();

            var firstDay = new DateTime(month.Year, month.Month, 1);
            int offset = ((int)firstDay.DayOfWeek + 6) % 7;
            int daysInMonth = DateTime.DaysInMonth(month.Year, month.Month);

            for (int i = 0; i < 42; i++)
            {
                int dayNumber = i - offset + 1;
                if (dayNumber > 0 && dayNumber <= daysInMonth)
                {
                    var dateButton = new Button
                    {
                        Text = dayNumber.ToString(),
                        BackgroundColor = Colors.Transparent,
                        TextColor = Colors.Black,
                        HeightRequest = 36,
                        WidthRequest = 36,
                        CornerRadius = 18,
                        Margin = new Thickness(2),
                        Padding = new Thickness(0)
                    };

                    var currentDate = new DateTime(month.Year, month.Month, dayNumber);

                    // 設置選中狀態的視覺效果
                    if (currentDate >= StartDate && currentDate <= EndDate)
                    {
                        dateButton.BackgroundColor = Color.FromArgb("#FF3B30");
                        dateButton.TextColor = Colors.White;
                    }

                    var dateForHandler = currentDate;
                    dateButton.Clicked += (s, e) =>
                    {
                        if (_isStartDateActive)
                        {
                            StartDate = dateForHandler;
                            _isStartDateActive = false; // 選完開始日期後，切換到選擇結束日期
                            EndDate = dateForHandler; // 初始化結束日期為開始日期
                        }
                        else
                        {
                            if (dateForHandler < StartDate)
                            {
                                // 如果選擇的結束日期比開始日期早，則重新開始選擇
                                StartDate = dateForHandler;
                                _isStartDateActive = false;
                                EndDate = dateForHandler;
                            }
                            else
                            {
                                EndDate = dateForHandler;
                                Calendar.IsVisible = false; // 選擇完結束日期後關閉日曆
                            }
                        }

                        MainThread.BeginInvokeOnMainThread(() =>
                        {
                            UpdateCalendarDisplay();
                        });
                    };

                    grid.Add(dateButton, i % 7, i / 7);
                }
            }
        }


    }
}