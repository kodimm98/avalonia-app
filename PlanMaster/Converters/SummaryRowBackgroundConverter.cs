using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace PlanMaster.Converters;

public sealed class SummaryRowBackgroundConverter : IValueConverter
{
    private static readonly IBrush Summary = new SolidColorBrush(Color.Parse("#F1F3F6"));
    private static readonly IBrush Normal = Brushes.Transparent;

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b && b ? Summary : Normal;

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}