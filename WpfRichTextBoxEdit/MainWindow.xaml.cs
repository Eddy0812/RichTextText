using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace WpfRichTextBoxEdit
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        RichTxtHelp richTxtHelp;
        public MainWindow()
        {
            InitializeComponent();
            //richTxtHelp = new RichTxtHelp(rtbMain, this);
        }

        private void comboBoxFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!rtbMain.Selection.IsEmpty)
            {
                TextRange range = new TextRange(rtbMain.Selection.Start, rtbMain.Selection.End);
                var v = ((System.Drawing.FontFamily)comboBoxFontFamily.SelectedValue).Name.ToString();
                range.ApplyPropertyValue(RichTextBox.FontFamilyProperty, v);
            }
        }

        private void comboBoxFontColor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!rtbMain.Selection.IsEmpty)
            {
                TextRange range = new TextRange(rtbMain.Selection.Start, rtbMain.Selection.End);
                Brush b = (Brush)new BrushConverter().ConvertFromString(comboBoxFontColor.SelectedValue.ToString());

                range.ApplyPropertyValue(RichTextBox.ForegroundProperty, b);
            }
        }

        private void comboBoxFontSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!rtbMain.Selection.IsEmpty)
            {
                rtbMain.Selection.ApplyPropertyValue(RichTextBox.FontSizeProperty, comboBoxFontSize.SelectedValue);
            }
            rtbMain.Focus();
        }

        /// <summary>
        /// 确认按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            //tableHelp.BindAllTabMouse();
        }

        /// <summary>
        /// 取消按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        bool isopen;
        /// <summary>
        /// 打开文件
        /// </summary>
        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.FileName = "";
            open.DefaultExt = ".xaml";
            open.Filter = "xaml文件(.xaml)|*.xaml|txt文件(.txt)|*.txt";
            Stream checkStream = null;

            if ((bool)open.ShowDialog())
            {
                try
                {
                    if ((checkStream = open.OpenFile()) != null)
                    {
                        isopen = true;
                        FileStream fs;
                        if (File.Exists(open.FileName))
                        {
                            fs = new FileStream(open.FileName, FileMode.Open, FileAccess.Read);
                            StreamReader streamReader = new StreamReader(fs, System.Text.Encoding.UTF8);

                            using (fs)
                            {
                                TextRange text = new TextRange(rtbMain.Document.ContentStart, rtbMain.Document.ContentEnd);
                                if (!open.FileName.Substring(open.FileName.Length - 4, 4).Contains(".txt"))
                                {
                                    text.Load(fs, DataFormats.XamlPackage);
                                }
                                else
                                {
                                    text.Load(fs, DataFormats.Text);
                                }
                                foreach (var block in rtbMain.Document.Blocks)
                                {
                                    if (block is BlockUIContainer bc)
                                    {
                                        if (bc.Child is Image image)
                                        {
                                            //image.Stretch = Stretch.Fill;
                                            image.Width = rtbMain.Width- rtbMain.Document.PagePadding.Left-rtbMain.Document.PagePadding.Right;
                                        }
                                    }
                                    else if(block is Paragraph p)
                                    {
                                        foreach (var item in p.Inlines)
                                        {
                                            if(item is InlineUIContainer ic)
                                            {
                                                if(ic.Child is Image image)
                                                {

                                                }
                                            }
                                        }
                                    }
                                }
                                isopen = false;
                            }

                            streamReader.Dispose();
                            streamReader.Close();
                            fs.Dispose();
                            fs.Close();

                        }
                        checkStream.Dispose();
                        checkStream.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("错误: 无法从磁盘读取文件。原始错误: " + ex.Message);
                }
            }
            else
            {
            }
        }

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "xaml文件(.xaml)|*.xaml|txt文件(.txt)|*.txt";
            try
            {
                if ((bool)save.ShowDialog())
                {
                    SaveFile(save.FileName);
                    MessageBox.Show("保存成功");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void SaveFile(string _fileName)
        {
            TextRange range;
            FileStream fStream;
            range = new TextRange(this.rtbMain.Document.ContentStart, this.rtbMain.Document.ContentEnd);
            fStream = new FileStream(_fileName, FileMode.Create);
            if (!_fileName.Substring(_fileName.Length - 4, 4).Contains(".txt"))
            {
                range.Save(fStream, DataFormats.XamlPackage);
            }
            else
            {
                range.Save(fStream, DataFormats.Text);
            }
            fStream.Close();
        }

        /// <summary>
        /// 打印(未测试)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPrint_Click(object sender, RoutedEventArgs e)
        {
            PrintCommand();
        }
        private void PrintCommand()
        {
            PrintDialog pd = new PrintDialog();
            if ((pd.ShowDialog() == true))
            {
                //use either one of the below      
                pd.PrintVisual(this.rtbMain as Visual, "printing as visual");
                pd.PrintDocument((((IDocumentPaginatorSource)this.rtbMain.Document).DocumentPaginator), "printing as paginator");
            }
        }

        private void btnAddImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "图像文件(*.jpg,*.jpe,*.gif,*.bmp,*.png)|*.jpg;*.jpe;*.gif;*.bmp;*.png|所有文件(*.*)|*.*";
            if ((bool)ofd.ShowDialog())
            {
                BitmapImage image= new BitmapImage(new Uri(ofd.FileName, UriKind.RelativeOrAbsolute));
                richTxtHelp.InsertImage(image);
            }
        }

        //合并单元格
        private void btnMerge_Click(object sender, RoutedEventArgs e)
        {
            richTxtHelp.MergeCell(rtbMain.Selection);
        }

        //撤销
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            rtbMain.Undo();
            richTxtHelp.UpdateColumn();
            richTxtHelp.BindAllTabMouse();
        }

        //重做
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            rtbMain.Redo();
            richTxtHelp.UpdateColumn();
            richTxtHelp.BindAllTabMouse();
        }

        //拆分单元格
        private void btnMerge_Click_1(object sender, RoutedEventArgs e)
        {
            richTxtHelp.SplitCell(null);
        }

        //插入行
        private void btnInsertRow_Click(object sender, RoutedEventArgs e)
        {
            richTxtHelp.InsertRow(null);
        }

        //插入列
        private void btnInsertColumn_Click(object sender, RoutedEventArgs e)
        {
            richTxtHelp.InsertCol(null);
        }

        //删除行
        private void btnRemoveRow_Click(object sender, RoutedEventArgs e)
        {
            richTxtHelp.RemoveRow(null);
        }

        //删除列
        private void btnInsertColumn_Click_1(object sender, RoutedEventArgs e)
        {
            richTxtHelp.RemoveCol(null);
        }

        //插入表格
        private void btnInsertTable_Click(object sender, RoutedEventArgs e)
        {
            richTxtHelp.InsertTable(5, 4);
        }

        private void pdfResolve_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.FileName = "";
            open.DefaultExt = ".pdf";
            open.Filter = "pdf文件(.pdf)|*.pdf";

            if ((bool)open.ShowDialog())
            {
                //Common.PdfHelp.OpenFile(open.FileName);
                //var pdfHelp = new Common.PdfHelp1();
                //TextRange textRange = new TextRange(rtbMain.Document.ContentStart, rtbMain.Document.ContentEnd);
                //textRange.Text = pdfHelp.Text;
                var pdfHelp = new Common.PdfHelp2();
                pdfHelp.ExtractImage(open.FileName);
                foreach (var item in pdfHelp.Images)
                {
                    richTxtHelp.InsertImage(item);
                }
                
            }
        }

        private void pdfResoveText_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.FileName = "";
            open.DefaultExt = ".pdf";
            open.Filter = "pdf文件(.pdf)|*.pdf";

            if ((bool)open.ShowDialog())
            {
                var pdfHelp = new Common.PdfHelp2();
                pdfHelp.ExtractText(open.FileName);
                TextRange textRange = new TextRange(rtbMain.Document.ContentStart, rtbMain.Document.ContentEnd);
                textRange.Text = pdfHelp.Text;
            }
                
        }

        
        private void wordResove_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open= new OpenFileDialog();
            open.FileName = "";
            open.DefaultExt = ".doc";
            open.Filter = "word文件|*.doc;*.docx|所有文件|*.*";
            if ((bool)open.ShowDialog())
            {
                Word.Application app = new Microsoft.Office.Interop.Word.Application();
                Word.Document doc=null;
                
                new Thread(() => {

                    Common.WordHelp.OpenDoc(open.FileName,ref doc,ref app);
                    doc.Select();
                    app.Selection.Copy();
                    Dispatcher.Invoke(new Action(() => {
                        IDataObject iData = Clipboard.GetDataObject();
                        if (iData.GetDataPresent(DataFormats.Bitmap))
                        {
                            var img = Clipboard.GetImage();
                        }
                        if (iData.GetDataPresent(DataFormats.Rtf))
                        {
                            var rtf = iData.GetData(DataFormats.Rtf);
                        }
                        rtbMain.Paste();//粘贴需要时间
                        
                        doc.Close();
                        doc = null;//此时粘贴完成了
                        app.Quit();//推出app会提示剪贴板中还有内容，不确定默认是清除了还是保存了剪贴板中的内容
                    })
                        );
                    
                }).Start();
                
                //Clipboard.Clear();
                //var text=Common.WordHelp.GetWordContent(open.FileName);
                //TextRange textRange = new TextRange(rtbMain.Document.ContentStart, rtbMain.Document.ContentEnd);
                //textRange.Text=text;
            }
        }
    }
}
