using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace PlanMaster.Converters;

public sealed class SheetNameRuConverter : IValueConverter
{
    public static readonly SheetNameRuConverter Instance = new();

    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        var s = value?.ToString() ?? "";
        if (s.StartsWith("Table", StringComparison.OrdinalIgnoreCase))
        {
            // "Table 2" -> "Таблица 2"
            return "Таблица" + s.Substring(5);
        }
        return s;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        => value?.ToString() ?? "";
}