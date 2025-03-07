using Microsoft.Maui.Controls;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace ScoreCard.Controls
{
    public partial class CustomDatePicker : ContentView, INotifyPropertyChanged
    {

        public static readonly BindableProperty StartDateProperty =
            BindableProperty.Create(nameof(StartDate), typeof(DateTime), typeof(CustomDatePicker),
                DateTime.Now, BindingMode.TwoWay, propertyChanged: OnDateChanged);

        public static readonly BindableProperty EndDateProperty =
            BindableProperty.Create(nameof(EndDate), typeof(DateTime), typeof(CustomDatePicker),
                DateTime.Now, BindingMode.TwoWay, propertyChanged: OnDateChanged);

        private DateTime _displayMonth;
        private bool _isStartDateActive;
        private bool _isInternalChange; // 標記內部變更，避免循環觸發

        // 日期變更事件
        public event EventHandler<DateTime> StartDateChanged;
        public event EventHandler<DateTime> EndDateChanged;
        public new event PropertyChangedEventHandler PropertyChanged;




        protected override void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            base.OnPropertyChanged(propertyName);
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public CustomDatePicker()
        {
            InitializeComponent();
            this.BindingContext = this; // 確保 BindingContext 是自己，避免繼承父元素的 BindingContext
            _displayMonth = DateTime.Now;
            _isStartDateActive = true;

            // 設置日期顯示的點擊事件
            var dateTapGesture = new TapGestureRecognizer();
            dateTapGesture.Tapped += (s, e) =>
            {
                Debug.WriteLine("Date display tapped");
                if (!Calendar.IsVisible)
                {
                    _isStartDateActive = true;
                    _displayMonth = StartDate;
                }
                Calendar.IsVisible = !Calendar.IsVisible;
                UpdateCalendarDisplay();
            };
            DateDisplay.GestureRecognizers.Add(dateTapGesture);

            // 背景點擊事件（隱藏日曆）
            var backgroundTapGesture = new TapGestureRecognizer();
            backgroundTapGesture.Tapped += (s, e) =>
            {
                Debug.WriteLine("Background tapped");
                if (Calendar.IsVisible)
                {
                    Calendar.IsVisible = false;
                }
            };
            this.GestureRecognizers.Add(backgroundTapGesture);

            // 日曆點擊事件（防止冒泡）
            var calendarTapGesture = new TapGestureRecognizer();
            calendarTapGesture.Tapped += (s, e) =>
            {
                Debug.WriteLine("Calendar tapped - stopping propagation");
                // 阻止事件傳播到背景
            };
            Calendar.GestureRecognizers.Add(calendarTapGesture);

            UpdateCalendarDisplay();
        }

        public DateTime StartDate
        {
            get => (DateTime)GetValue(StartDateProperty);
            set
            {
                if ((DateTime)GetValue(StartDateProperty) != value)
                {
                    _isInternalChange = true;
                    SetValue(StartDateProperty, value);
                    StartDateChanged?.Invoke(this, value);  // 確保事件被觸發
                    OnPropertyChanged(nameof(StartDate));
                    Debug.WriteLine($"[CustomDatePicker] StartDate set to: {value:yyyy-MM-dd}");
                    _isInternalChange = false;
                }
            }
        }

        public DateTime EndDate
        {
            get => (DateTime)GetValue(EndDateProperty);
            set
            {
                if ((DateTime)GetValue(EndDateProperty) != value)
                {
                    _isInternalChange = true;
                    SetValue(EndDateProperty, value);
                    EndDateChanged?.Invoke(this, value);  // 確保事件被觸發
                    OnPropertyChanged(nameof(EndDate));
                    Debug.WriteLine($"[CustomDatePicker] EndDate set to: {value:yyyy-MM-dd}");
                    _isInternalChange = false;
                }
            }
        }

        private void OnDateDayClicked(object sender, EventArgs e)
        {
            var dateButton = sender as Button;
            if (dateButton?.BindingContext is DateTime dateForHandler)
            {
                Debug.WriteLine($"[CustomDatePicker] Date button clicked: {dateForHandler:yyyy-MM-dd}");

                DateTime oldStart = StartDate;
                DateTime oldEnd = EndDate;

                if (_isStartDateActive)
                {
                    // 設置起始日期
                    StartDate = dateForHandler;
                    _isStartDateActive = false;

                    // 若選擇的起始日期比當前結束日期晚，則同時設置結束日期
                    if (dateForHandler > EndDate)
                    {
                        EndDate = dateForHandler;
                    }
                }
                else
                {
                    // 若選擇的結束日期比起始日期早，則交換順序
                    if (dateForHandler < StartDate)
                    {
                        EndDate = StartDate;
                        StartDate = dateForHandler;
                    }
                    else
                    {
                        EndDate = dateForHandler;
                        Calendar.IsVisible = false;  // 選擇完結束日期後自動關閉日曆
                    }
                }

                // 手動觸發事件，確保 ViewModel 接收到變更
                if (StartDate != oldStart)
                {
                    Debug.WriteLine($"[CustomDatePicker] Explicitly triggering StartDateChanged: {oldStart:yyyy-MM-dd} -> {StartDate:yyyy-MM-dd}");
                    StartDateChanged?.Invoke(this, StartDate);
                }

                if (EndDate != oldEnd)
                {
                    Debug.WriteLine($"[CustomDatePicker] Explicitly triggering EndDateChanged: {oldEnd:yyyy-MM-dd} -> {EndDate:yyyy-MM-dd}");
                    EndDateChanged?.Invoke(this, EndDate);
                }

                UpdateCalendarDisplay();
            }
        }


        private static void OnDateChanged(BindableObject bindable, object oldValue, object newValue)
        {
            var picker = (CustomDatePicker)bindable;

            // 避免處理內部變更導致的循環觸發
            if (picker._isInternalChange)
                return;

            // 更新日曆顯示
            picker.UpdateCalendarDisplay();

            // 確保組件在繫結的屬性變更時，主動通知 ViewModel
            if (oldValue is DateTime oldDate && newValue is DateTime newDate && oldDate != newDate)
            {
                string propertyName = null;

                // 判斷是哪個屬性變更了
                if (bindable.GetType().GetProperty(nameof(StartDate)).GetValue(bindable).Equals(newValue))
                {
                    propertyName = nameof(StartDate);
                    Debug.WriteLine($"[CustomDatePicker] OnDateChanged: StartDate from {oldDate:yyyy-MM-dd} to {newDate:yyyy-MM-dd}");
                    picker.StartDateChanged?.Invoke(picker, newDate);
                }
                else if (bindable.GetType().GetProperty(nameof(EndDate)).GetValue(bindable).Equals(newValue))
                {
                    propertyName = nameof(EndDate);
                    Debug.WriteLine($"[CustomDatePicker] OnDateChanged: EndDate from {oldDate:yyyy-MM-dd} to {newDate:yyyy-MM-dd}");
                    picker.EndDateChanged?.Invoke(picker, newDate);
                }
            }
        }

        private void OnStartDateClicked(object sender, EventArgs e)
        {
            Debug.WriteLine("[CustomDatePicker] Start date clicked");
            _isStartDateActive = true;
            _displayMonth = StartDate;
            Calendar.IsVisible = true;
            UpdateCalendarDisplay();
        }

        private void OnEndDateClicked(object sender, EventArgs e)
        {
            Debug.WriteLine("[CustomDatePicker] End date clicked");
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
            int offset = ((int)firstDay.DayOfWeek + 6) % 7; // 調整為星期一開始
            int daysInMonth = DateTime.DaysInMonth(month.Year, month.Month);

            for (int i = 0; i < 42; i++) // 6週 x 7天
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

                    // 設置所選日期範圍的視覺效果
                    if (currentDate >= StartDate && currentDate <= EndDate)
                    {
                        dateButton.BackgroundColor = Color.FromArgb("#FF3B30");
                        dateButton.TextColor = Colors.White;
                    }

                    var dateForHandler = currentDate;
                    dateButton.Clicked += (s, e) =>
                    {
                        Debug.WriteLine($"[CustomDatePicker] Date button clicked: {dateForHandler:yyyy-MM-dd}");

                        DateTime oldStart = StartDate;
                        DateTime oldEnd = EndDate;

                        if (_isStartDateActive)
                        {
                            // 設置起始日期
                            StartDate = dateForHandler;
                            _isStartDateActive = false;

                            // 若選擇的起始日期比當前結束日期晚，則同時設置結束日期
                            if (dateForHandler > EndDate)
                            {
                                EndDate = dateForHandler;
                            }
                        }
                        else
                        {
                            // 若選擇的結束日期比起始日期早，則交換順序
                            if (dateForHandler < StartDate)
                            {
                                EndDate = StartDate;
                                StartDate = dateForHandler;
                            }
                            else
                            {
                                EndDate = dateForHandler;
                                Calendar.IsVisible = false;
                            }
                        }

                        // 手動觸發事件，確保 ViewModel 能接收到變更
                        MainThread.BeginInvokeOnMainThread(() =>
                        {
                            if (StartDate != oldStart)
                            {
                                Debug.WriteLine($"[CustomDatePicker] Explicitly triggering StartDateChanged: {oldStart:yyyy-MM-dd} -> {StartDate:yyyy-MM-dd}");
                                StartDateChanged?.Invoke(this, StartDate);
                            }

                            if (EndDate != oldEnd)
                            {
                                Debug.WriteLine($"[CustomDatePicker] Explicitly triggering EndDateChanged: {oldEnd:yyyy-MM-dd} -> {EndDate:yyyy-MM-dd}");
                                EndDateChanged?.Invoke(this, EndDate);
                            }

                            UpdateCalendarDisplay();
                        });
                    };

                    grid.Add(dateButton, i % 7, i / 7);
                }
            }
        }
    }
}