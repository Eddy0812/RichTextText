using java.util;
using org.pdfbox.pdmodel;
using org.pdfbox.pdmodel.graphics.xobject;
using org.pdfbox.util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WpfRichTextBoxEdit.Common
{
    public class PdfHelp1
    {
        public string Text { get; set; }
        public  void OpenFile(string pdfPath)
        {
            PDDocument doc = PDDocument.load(pdfPath);
            PDDocumentInformation info=doc.getDocumentInformation();
            Dictionary<string, string> parameters = new Dictionary<string, string>()
            {
                {"标题",info.getTitle() },
                {"主题",info.getSubject()},
                {"作者",info.getAuthor()},
                {"关键字",info.getKeywords()},
                {"应用程序",info.getCreator()},
                {"创建时间",info.getCreationDate().ToString() },
                {"修改时间",info.getModificationDate().toString()},
            };
            PDDocumentCatalog cata = doc.getDocumentCatalog();
            List pages = cata.getAllPages();
            int count = 1;
            for (int i = 0; i < pages.size(); i++)
            {
                PDPage page = (PDPage)pages.get(i);
                if (null != page)
                {
                    PDResources res = page.findResources();
                    var contents = page.getContents();

                    //获取页面图片信息  
                    Map imgs = res.getImages();
                    if (null != imgs)
                    {
                        Set keySet = imgs.keySet();
                        Iterator it = keySet.iterator();
                        while (it.hasNext())
                        {
                            Object obj = it.next();
                            PDXObjectImage img = (PDXObjectImage)imgs.get(obj);
                            count++;
                        }
                    }
                }
            } 
            PDFTextStripper stripper = new PDFTextStripper();
            string txt = stripper.getText(doc);
            Text= txt;
            var stripper1 = new org.pdfbox.util.PDFTextStripperByArea();
        }

    }
}
