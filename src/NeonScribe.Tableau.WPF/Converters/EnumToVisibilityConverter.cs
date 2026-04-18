using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace NeonScribe.Tableau.WPF.Converters;

/// <summary>
/// Returns Visible when the bound enum value matches the ConverterParameter, Collapsed otherwise.
/// </summary>
public class EnumToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value == null || parameter == null)
            return Visibility.Collapsed;

        return value.ToString() == parameter.ToString() ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}
