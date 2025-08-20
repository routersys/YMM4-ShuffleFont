using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Threading.Tasks;

namespace FontShuffle
{
    public class BoolToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                return (value is bool v && v) ?
                    (SolidColorBrush)new BrushConverter().ConvertFrom("#FFD700")! :
                    (SolidColorBrush)new BrushConverter().ConvertFrom("#FF808080")!;
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "BoolToColor変換"));
                return (SolidColorBrush)new BrushConverter().ConvertFrom("#FF808080")!;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class HasCustomSettingsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                if (value is FontItem font)
                {
                    return font.HasCustomSettings ? Visibility.Visible : Visibility.Collapsed;
                }
                if (value is string fontName)
                {
                    return Visibility.Collapsed;
                }
                return Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "HasCustomSettings変換"));
                return Visibility.Collapsed;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                if (value is System.Windows.Media.Color color)
                {
                    return new SolidColorBrush(color);
                }
                return new SolidColorBrush(Colors.White);
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "ColorToBrush変換"));
                return new SolidColorBrush(Colors.White);
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                if (value is SolidColorBrush brush)
                {
                    return brush.Color;
                }
                return Colors.White;
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "BrushToColor変換"));
                return Colors.White;
            }
        }
    }

    public class OrderNumberConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                if (value is ListBoxItem item)
                {
                    var listBox = ItemsControl.ItemsControlFromItemContainer(item) as ListBox;
                    if (listBox != null)
                    {
                        int index = -1;
                        for (int i = 0; i < listBox.Items.Count; i++)
                        {
                            if (listBox.Items[i] == item.DataContext)
                            {
                                index = i;
                                break;
                            }
                        }
                        return index >= 0 ? $"{index + 1}." : "";
                    }
                }
                return "";
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "OrderNumber変換"));
                return "";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}