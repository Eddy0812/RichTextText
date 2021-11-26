using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WpfRichTextBoxEdit.Resources
{
    public class FontSizeSource : List<double>
    {
        public FontSizeSource()
        {
            this.AddRange(Enumerable.Range(10, 62).Where(p => p % 4 == 0).Select(p => (double)p));
        }
    }
}
