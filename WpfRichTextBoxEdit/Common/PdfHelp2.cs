using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Media.Imaging;

namespace WpfRichTextBoxEdit.Common
{
    public class PdfHelp2
    {
        public string Text { get; set; }
        public void OpenFile()
        {
            
        }

        public void ExtractText(string fileName)
        {
            string fileContent = string.Empty;
            StringBuilder sbFileContent = new StringBuilder();
            PdfReader reader = null;
            try
            {
                reader = new PdfReader(fileName);
            }
            catch (Exception)
            {
                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }
            }

            try
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    sbFileContent.AppendLine(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                    reader = null;
                }
            }
            fileContent = sbFileContent.ToString();
            Text = fileContent;
        }

        public List<BitmapImage> Images = new List<BitmapImage>();
        public void ExtractImage(string pdfFile)
        {
            PdfReader pdfReader = new PdfReader(pdfFile);
            for (int pageNumber = 1; pageNumber <= pdfReader.NumberOfPages; pageNumber++)
            {

                PdfReader pdf = new PdfReader(pdfFile);

                PdfDictionary pg = pdf.GetPageN(pageNumber);

                PdfDictionary res = (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));

                PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));

                try
                {
                    foreach (PdfName name in xobj.Keys)
                    {

                        PdfObject obj = xobj.Get(name);

                        if (obj.IsIndirect())
                        {

                            PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);

                            string width = tg.Get(PdfName.WIDTH).ToString();

                            string height = tg.Get(PdfName.HEIGHT).ToString();

                            //ImageRenderInfo imgRI = ImageRenderInfo.CreateForXObject((GraphicsState)new Matrix(float.Parse(width), float.Parse(height)), (PRIndirectReference)obj, tg);

                            ImageRenderInfo imgRI = ImageRenderInfo.CreateForXObject(new GraphicsState(), (PRIndirectReference)obj, tg);
                            RenderImage(imgRI);
                        }

                    }

                }
                catch
                {
                    continue;
                }

            }

        }

        private void RenderImage(ImageRenderInfo renderInfo)
        {
            PdfImageObject image = renderInfo.GetImage();

            using (var dotnetImg = image.GetDrawingImage())
            {
                if (dotnetImg != null)

                {

                    using (MemoryStream ms = new MemoryStream())
                    {

                        dotnetImg.Save(ms, ImageFormat.Tiff);
                        BitmapImage ix = new BitmapImage();
                        ix.BeginInit();
                        ix.CacheOption= BitmapCacheOption.OnLoad;
                        ix.StreamSource= ms;
                        ix.EndInit();
                        Images.Add(ix);
                        //Bitmap d = new Bitmap(dotnetImg);

                        //d.Save(@"");

                    }

                }

            }

        }
    }
}
