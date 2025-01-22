using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;

namespace OMNI.QMS.Converters
{
    public class QIRIDNumbertoVisibility : IValueConverter, IMultiValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (parameter == null)
            {
                return value == null ? Visibility.Collapsed : Visibility.Visible;
            }
            else if (parameter != null && value == null)
            {
                return Visibility.Visible;
            }
            else if (parameter != null && (int)value > 0)
            {
                return Visibility.Visible;
            }
            else
            {
                return Visibility.Collapsed;
            }
        }
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value is Visibility && (Visibility)value == Visibility.Visible;
        }

        #region IMultiValueConverter Implementation

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Count() == 2)
            {
                var _id = values[0] != null && int.TryParse(values[0].ToString(), out int i) ? i : 0;
                var _legacy = values[1] != null && bool.TryParse(values[1].ToString(), out bool b) ? b : false;
                return _id > 0 && _legacy ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
