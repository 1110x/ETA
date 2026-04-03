using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using System;

namespace ETA.Views;

/// <summary>
/// 인라인 캘린더 공통 생성 헬퍼.
/// 📅 버튼 클릭 → 캘린더 토글, 날짜 선택 시 콜백 + 자동 닫힘, ESC로 닫기.
/// </summary>
public static class InlineCalendarHelper
{
    private static readonly FontFamily Font =
        new("avares://ETA/Assets/Fonts#Pretendard");

    /// <summary>
    /// 인라인 캘린더 Border를 생성합니다.
    /// </summary>
    /// <param name="onDateSelected">날짜 선택 시 콜백</param>
    /// <param name="toggleButton">캘린더를 토글할 버튼 (null이면 직접 제어)</param>
    /// <returns>캘린더를 감싼 Border (IsVisible=false 상태로 시작)</returns>
    public static Border Create(Action<DateTime> onDateSelected, Button? toggleButton = null)
    {
        var calendar = new Calendar
        {
            SelectionMode = CalendarSelectionMode.SingleDate,
            FontFamily = Font,
            FontSize = AppTheme.FontMD,
            DisplayDate = DateTime.Today,
            Focusable = true,
        };

        var border = new Border
        {
            IsVisible = false,
            Background = AppTheme.BgInput,
            BorderBrush = AppTheme.BorderSubtle,
            BorderThickness = new Avalonia.Thickness(0, 0, 0, 1),
            Padding = new Avalonia.Thickness(4),
            Margin = new Avalonia.Thickness(0, 4, 0, 0),
            Focusable = true,
            Child = calendar,
        };

        // 날짜 선택 → 콜백 + 닫기
        calendar.SelectedDatesChanged += (_, _) =>
        {
            if (calendar.SelectedDate.HasValue)
            {
                var date = calendar.SelectedDate.Value;
                border.IsVisible = false;
                onDateSelected(date);
                calendar.SelectedDates.Clear();
            }
        };

        // ESC → 닫기
        border.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape && border.IsVisible)
            {
                border.IsVisible = false;
                e.Handled = true;
            }
        };
        calendar.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Escape && border.IsVisible)
            {
                border.IsVisible = false;
                e.Handled = true;
            }
        };

        // 토글 버튼 연결
        if (toggleButton != null)
        {
            toggleButton.Click += (_, _) =>
            {
                border.IsVisible = !border.IsVisible;
                if (border.IsVisible)
                {
                    calendar.SelectedDates.Clear();
                    calendar.DisplayDate = DateTime.Today;
                    calendar.Focus();
                }
            };
        }

        return border;
    }

    /// <summary>캘린더 토글 (외부에서 직접 제어할 때)</summary>
    public static void Toggle(Border calendarBorder)
    {
        calendarBorder.IsVisible = !calendarBorder.IsVisible;
        if (calendarBorder.IsVisible && calendarBorder.Child is Calendar cal)
        {
            cal.SelectedDates.Clear();
            cal.DisplayDate = DateTime.Today;
            cal.Focus();
        }
    }

    /// <summary>캘린더 닫기</summary>
    public static void Hide(Border calendarBorder)
    {
        calendarBorder.IsVisible = false;
    }
}
