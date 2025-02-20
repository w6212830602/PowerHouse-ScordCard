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

        // 添加日期变更事件
        public event EventHandler<DateTime> StartDateChanged;
        public event EventHandler<DateTime> EndDateChanged;
        public event PropertyChangedEventHandler PropertyChanged;


        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }


        public CustomDatePicker()
        {
            InitializeComponent();
            BindingContext = this;
            _displayMonth = DateTime.Now;
            _isStartDateActive = true;

            var dateTapGesture = new TapGestureRecognizer();
            dateTapGesture.Tapped += (s, e) =>
            {
                if (!Calendar.IsVisible)
                {
                    _isStartDateActive = true;
                    _displayMonth = StartDate;
                }
                Calendar.IsVisible = !Calendar.IsVisible;
                UpdateCalendarDisplay();
            };
            DateDisplay.GestureRecognizers.Add(dateTapGesture);

            var backgroundTapGesture = new TapGestureRecognizer();
            backgroundTapGesture.Tapped += (s, e) =>
            {
                if (Calendar.IsVisible)
                {
                    Calendar.IsVisible = false;
                }
            };
            this.GestureRecognizers.Add(backgroundTapGesture);

            var calendarTapGesture = new TapGestureRecognizer();
            calendarTapGesture.Tapped += (s, e) => { };
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
                    SetValue(StartDateProperty, value);
                    StartDateChanged?.Invoke(this, value);
                    OnPropertyChanged(nameof(StartDate));
                    Debug.WriteLine($"Start date changed to: {value:yyyy-MM-dd}");
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
                    SetValue(EndDateProperty, value);
                    EndDateChanged?.Invoke(this, value);
                    OnPropertyChanged(nameof(EndDate));
                    Debug.WriteLine($"End date changed to: {value:yyyy-MM-dd}");
                }
            }
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
                            _isStartDateActive = false;
                            EndDate = dateForHandler;
                        }
                        else
                        {
                            if (dateForHandler < StartDate)
                            {
                                StartDate = dateForHandler;
                                _isStartDateActive = false;
                                EndDate = dateForHandler;
                            }
                            else
                            {
                                EndDate = dateForHandler;
                                Calendar.IsVisible = false;

                                // 触发日期更改事件
                                StartDateChanged?.Invoke(this, StartDate);
                                EndDateChanged?.Invoke(this, EndDate);
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