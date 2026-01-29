using System;
using System.Globalization;
using Avalonia;
using Avalonia.Data.Converters;
using Avalonia.Controls;

namespace PlanMaster.Converters;

public sealed class BoolToColumnWidthConverter : IValueConverter
{
    public static readonly BoolToColumnWidthConverter Instance = new();

    // true  -> 280 (показано)
    // false -> 0   (скрыто)
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var shown = value is bool b && b;
        var width = shown ? 280 : 0;
        return new GridLength(width, GridUnitType.Pixel);
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => AvaloniaProperty.UnsetValue!;
}