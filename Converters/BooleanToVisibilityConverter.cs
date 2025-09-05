using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ExcelDynamicViewer.Converters;

/// <summary>
/// Converts a boolean value to Visibility (Visible/Collapsed).
/// </summary>
public class BooleanToVisibilityConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        return value is true ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}