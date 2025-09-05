using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelDynamicViewer.Converters;

public class RowIndexConverter : IValueConverter
{
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is DataGridRow row)
        {
            return row.GetIndex() + 1;
        }
        return string.Empty;
    }

    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}