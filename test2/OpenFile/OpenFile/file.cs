using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace File
{
    class Program
    {
        static void Main(string[] args)
        {
            string Str = @"G:\科研细则.docx";//找到文件所在路径
            using (WordprocessingDocument doc = WordprocessingDocument.Open(Str, true))//打开文件操作
            {
                Body body = doc.MainDocumentPart.Document.Body;//创建对象
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    Console.WriteLine(paragraph.InnerText);
                }
            }
            Console.ReadLine();
        }
    }
}
