using System;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace PlanMaster.Converters;

public sealed class SummaryRowFontWeightConverter : IValueConverter
{
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value is bool b && b ? FontWeight.SemiBold : FontWeight.Normal;

    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => throw new NotSupportedException();
}