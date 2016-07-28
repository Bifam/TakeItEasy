using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace TakeItEasy
{
    public class CustomTemplateSelector : DataTemplateSelector
    {
        public DataTemplate SuggestTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
                return SuggestTemplate;
        }
    }
}
