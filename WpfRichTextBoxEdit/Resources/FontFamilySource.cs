using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WpfRichTextBoxEdit.Resources
{
    public class FontFamilySource : List<System.Drawing.FontFamily> 
    {
        public FontFamilySource()
        {
            this.AddRange(new System.Drawing.Text.InstalledFontCollection().Families.Select(p => p).ToList());
        }
    }
}
