﻿using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test3
{
    class RWord
    {
        //源文档
        public string[] readList()
        {
            string fileName = @"D:\Program Files\VS2017_workplace\test3\国考_原题.docx";
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(fileName, false))
            {
                // Create a Body object.
                DocumentFormat.OpenXml.Wordprocessing.Body body =
           wordprocessingDocument.MainDocumentPart.Document.Body;

                List<string> list = new List<string>();
                foreach (var g in body.Elements())
                {
                    list.Add(g.InnerText);
                }
                //创建一个字符型的一维数组
                char[] chArr = new char[list.Count];
                //进行数组的初始化
                string[] str = list.ToArray();
                //将其转化为字符数组
                modify CH = new modify();            
                chArr = CH.changeCh(str,list.Count);
                //再转换为string[]
                string[] s = new string[chArr.Length];
                //调用函数
                s = CH.changeStr(chArr);
                return s;
                
            }
        }
        //第一个文档
        public string[] readList1()
        {
            string fileName = @"D:\Program Files\VS2017_workplace\test3\国考_标准答案1.docx";
            
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(fileName, false))
            {
                // Create a Body object.
                DocumentFormat.OpenXml.Wordprocessing.Body body =
           wordprocessingDocument.MainDocumentPart.Document.Body;

                List<string> list1 = new List<string>();
                foreach (var g in body.Elements())
                {
                    list1.Add(g.InnerText);
                }
                //初始化一个字符数组
                char[] chArr = new char[list1.Count];
                //将list转化为字符串数组
                string[] str = list1.ToArray();
                //调用函数转化为字符数组
                modify CH = new modify();
                chArr = CH.changeCh(str, list1.Count);
                //再转换为string[]
                string[] s = new string[chArr.Length];
                //调用函数
                s = CH.changeStr(chArr);
                return s;
            }
        }
        //第二个文档
        public string[] readList2()
        {
            string fileName = @"D:\Program Files\VS2017_workplace\test3\国考_标准答案2.docx";
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(fileName, false))
            {
                // Create a Body object.
                DocumentFormat.OpenXml.Wordprocessing.Body body =
           wordprocessingDocument.MainDocumentPart.Document.Body;

                List<string> list2 = new List<string>();
                foreach (var g in body.Elements())
                {
                    list2.Add(g.InnerText);
                }
                //初始化一个字符数组
                char[] chArr = new char[list2.Count];
                //将list转化为字符串数组
                string[] str = list2.ToArray();
                //调用函数转化为字符数组
                modify CH = new modify();
                chArr = CH.changeCh(str, list2.Count);
                //再转换为string[]
                string[] s = new string[chArr.Length];
                //调用函数
                s = CH.changeStr(chArr);
                return s;
            }
        }
        //第三个文档
        public string[] readList3()
        {
            string fileName = @"D:\Program Files\VS2017_workplace\test3\国考_标准答案3.docx";
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(fileName, false))
            {
                // Create a Body object.
                DocumentFormat.OpenXml.Wordprocessing.Body body =
           wordprocessingDocument.MainDocumentPart.Document.Body;

                List<string> list3 = new List<string>();
                foreach (var g in body.Elements())
                {
                    list3.Add(g.InnerText);
                }
                //初始化一个字符数组
                char[] chArr = new char[list3.Count];
                //将list转化为字符串数组
                string[] str = list3.ToArray();
                //调用函数转化为字符数组
                modify CH = new modify();
                chArr = CH.changeCh(str, list3.Count);
                //再转换为string[]
                string[] s = new string[chArr.Length];
                //调用函数
                s = CH.changeStr(chArr);
                return s;
            }
        }
        //第四个文档
        public string[] readList4()
        {
            string fileName = @"D:\Program Files\VS2017_workplace\test3\国考_原题.docx";
            using (WordprocessingDocument wordprocessingDocument =
            WordprocessingDocument.Open(fileName, false))
            {
                // Create a Body object.
                DocumentFormat.OpenXml.Wordprocessing.Body body =
           wordprocessingDocument.MainDocumentPart.Document.Body;

                List<string> list4 = new List<string>();
                foreach (var g in body.Elements())
                {
                    list4.Add(g.InnerText);
                }
                //初始化一个字符数组
                char[] chArr = new char[list4.Count];
                //将list转化为字符串数组
                string[] str = list4.ToArray();
                //调用函数转化为字符数组
                modify CH = new modify();
                chArr = CH.changeCh(str, list4.Count);
                //再转换为string[]
                string[] s = new string[chArr.Length];
                //调用函数
                s = CH.changeStr(chArr);
                return s;

            }
        }
    }
}
