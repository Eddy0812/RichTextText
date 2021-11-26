using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WpfRichTextBoxEdit.Resources
{
    public class FontColorSource : List<string>
    {
        public FontColorSource()
        {
            List<string> colorinfo = new List<string>();
            Type type = typeof(System.Windows.Media.Brushes);
            System.Reflection.PropertyInfo[] info = type.GetProperties();
            foreach (System.Reflection.PropertyInfo pi in info)
            {
                colorinfo.Add(pi.Name);
            }
            this.AddRange(colorinfo);
        }
    }
}
