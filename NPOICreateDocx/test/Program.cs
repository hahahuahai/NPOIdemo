using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream stream = File.OpenRead(@"F:\汇编文档.docx"))
            {
                XWPFDocument doc = new XWPFDocument(stream);
                IList<XWPFParagraph> lstPr = doc.Paragraphs;
                foreach (XWPFParagraph pr in lstPr)
                {
                    foreach (XWPFRun gr in pr.GetRuns())
                    {
                        //gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)50;
                        //gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)50;
                        //gr.GetCTR().AddNewRPr().AddNewB().val = true; //加粗
                        //gr.GetCTR().AddNewRPr().AddNewColor().val = "red";//字体颜色
                        Console.WriteLine(gr.GetFontFamily());
                    }
                    Console.WriteLine("{0}",pr.GetRuns().Count);    //1 1 4 3
                }
                FileStream sw = File.OpenWrite(@"F:\output.docx");
                doc.Write(sw);
                sw.Close();
                Console.WriteLine("完成。");
                Console.ReadKey();
            }
        }

    }
}
