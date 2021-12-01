using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace WpfRichTextBoxEdit.Common
{
    public class WordHelp
    {
        ///
        /// 读取 word文档 返回内容，此方法需在项目应用中引用com组件：Microsoft.Office.Interop.Word，如果与PC上安装的Office版本对不上，则解析报错，不能使用
        ///
        //////
        public static string GetWordContent(string path)
        {
            try
            {
                Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Type wordType = app.GetType();
                Word.Document doc = null;
                object unknow = Type.Missing;
                app.Visible = false;

                object file = path;
                doc = app.Documents.Open(ref file,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow, ref unknow,
                ref unknow, ref unknow, ref unknow);
                app.Selection.Copy();
                int count = doc.Paragraphs.Count;
                StringBuilder sb = new StringBuilder();
                for (int i = 1; i <= count; i++)
                {

                    sb.Append(doc.Paragraphs[i].Range.Text.Trim());
                }
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    var fText = p.Range.FormattedText;
                    var tables = p.Range.Tables;
                    var inlines = p.Range.InlineShapes;
                }
                doc.Close(ref unknow, ref unknow, ref unknow);
                wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null);
                doc = null;
                app = null;
                //垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();

                string temp = sb.ToString();
                //if (temp.Length > 200)
                // return temp.Substring(0, 200);
                //else
                return temp;
            }
            catch(Exception ex)
            {
                return ex.Message;
            }
        }

        public static void OpenDoc(string path, ref Word.Document doc, ref Word.Application app)
        {
            
            Type wordType = app.GetType();
            doc = null;
            object unknow = Type.Missing;
            object readOnly = (object)true;
            app.Visible = false;
            
            object file = path;
            doc = app.Documents.Open(ref file,
            ref unknow, ref unknow, ref unknow, ref unknow,
            ref unknow, ref unknow, ref unknow, ref unknow,
            ref unknow, ref unknow, ref unknow, ref unknow,
            ref unknow, ref unknow, ref unknow);
            //doc.SelectAllEditableRanges();
            
        }
    }
}
