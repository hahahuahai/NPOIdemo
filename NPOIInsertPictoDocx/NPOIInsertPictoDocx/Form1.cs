using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
/* 
 * 实现了页面高、宽设置、插图。
 * 插图有两种方式：inline,anchor
 * 内联(inline)对应嵌入型，
 * 锚(anchor)对应四周型、紧密型、穿越型、上下型、文字上方型、文字下方型
 * 四周型、紧密型、穿越型有四种自动换行方式，四周型不需要多边形（wrapPolygon）
 * wrapText值为：bothSides（两边），left（只在左侧），right（只在右侧），largest（只在最宽一侧）
 * 自动换行方式：
 * 两边：需要四个lineTo
 * posHOffset，posVOffset为图的左上角坐标，1cm=360000EMUS，width,height图宽和高，1cm=360000EMUS
 * 本例子提供的NPOI已经过修改。 
 * 重新修改NPOI关于anchor插图AddPicture函数
 * vs2010
 * netframework4
 * 创建的docx在word2007可以打开
 * 重新修改日期 2014-10-8
 */
namespace NPOIInsertPictoDocx
{
    public partial class Form1 : Form
    {
        const String m_savefilepath = "d:\\doc\\NPOI\\Picture";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //inline插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("inline","","");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\inline.docx");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //anchor-topandbottom(上下型)
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapTopAndBottom", "");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapTopAndBottom.docx");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //anchor-none(文字之上型)
            //anchor.BehindDoc=0
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapNoneBehindDoc", "");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapNoneBehindDoc.docx");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //anchor-none(文字之上型)
            //anchor.BehindDoc=1
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapBehindDoc", "");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapBehindDoc.docx");
        }

        private void button2_Besides_Click(object sender, EventArgs e)
        {
            //anchor-wrapSquare-besides插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapSquare", "Besides");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapSquareBesides.docx");
        }

        private void button2_Left_Click(object sender, EventArgs e)
        {
            //anchor-wrapSquare-left插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapSquare", "Left");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapSquareLeft.docx");

        }

        private void button2_Right_Click(object sender, EventArgs e)
        {
            //anchor-wrapSquare-right插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapSquare", "Right");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapSquareRight.docx");
        }

        private void button2_Largest_Click(object sender, EventArgs e)
        {
            //anchor-wrapSquare-largest插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapSquare", "Largest");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapSquareLargest.docx");

        }

        private void button3_Besides_Click(object sender, EventArgs e)
        {
            //anchor-wrapTight-besides插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapTight", "Besides");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapTightBesides.docx");

        }

        private void button3_Left_Click(object sender, EventArgs e)
        {
            //anchor-wrapTight-left插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapTight", "Left");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapTightLeft.docx");

        }

        private void button3_Right_Click(object sender, EventArgs e)
        {
            //anchor-wrapTight-right插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapTight", "Right");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapTightRight.docx");

        }

        private void button3_Largest_Click(object sender, EventArgs e)
        {
            //anchor-wrapTight-largest插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapTight", "Largest");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapTightLargest.docx");

        }

        private void button4_Besides_Click(object sender, EventArgs e)
        {
            //anchor-wrapThrough-Besides插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapThrough", "Besides");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapThroughBesides.docx");

        }

        private void button4_Left_Click(object sender, EventArgs e)
        {
            //anchor-wrapThrough-Left插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapThrough", "Left");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapThroughLeft.docx");

        }

        private void button4_Right_Click(object sender, EventArgs e)
        {
            //anchor-wrapThrough-Right插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapThrough", "Right");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapThroughRight.docx");

        }

        private void button4_Largest_Click(object sender, EventArgs e)
        {
            //anchor-wrapThrough-Largest插图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = InsertPictoDocx("anchor", "wrapThrough", "Largest");
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\wrapThroughLargest.docx");

        }

        protected XWPFDocument InsertPictoDocx(String parType,String wrapType,String posType)
        {
            //parType:inline，anchor
            //wrapType:当parType="anchor"时的值为：wrapSquare(四周)，wrapTight(紧密)，wrapThrough(穿越)
            XWPFDocument m_Docx = new XWPFDocument();
            //页面设置
            //A4:W=11906,h=16838
            //CT_SectPr m_SectPr = m_Docx.Document.body.AddNewSectPr();
            m_Docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_SectPr = m_Docx.Document.body.sectPr;    
            //页面设置A4横向
            m_SectPr.pgSz.w = (ulong)16838;
            m_SectPr.pgSz.h = (ulong)11906;
            FileStream gfs = null;
            
            //XWPFParagraph gp = m_Docx.CreateParagraph();
            CT_P m_p = m_Docx.Document.body.AddNewP();
            m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
            XWPFParagraph gp = new XWPFParagraph(m_p, m_Docx); 

            XWPFRun gr = gp.CreateRun();
            gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = "黑体";
            gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = "黑体";
            gr.GetCTR().AddNewRPr().AddNewRFonts().hint = ST_Hint.eastAsia;
            gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)44;
            gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)44;
            gr.GetCTR().AddNewRPr().AddNewColor().val = "red";   
            gr.SetText("NPOI插图");

            gfs = new FileStream("f:\\1.jpg", FileMode.Open, FileAccess.Read);
            m_p = m_Docx.Document.body.AddNewP();
            m_p.AddNewPPr().AddNewJc().val = ST_Jc.both;//段落两端对齐
            gp = new XWPFParagraph(m_p, m_Docx);
            gr = gp.CreateRun();
            gr.SetText("NPOI，顾名思义，就是POI的.NET版本。那POI又是什么呢？POI是一套用Java写成的库，能够帮助开 发者在没有安装微软Office的情况下读写Office 97-2003的文件，支持的文件格式包括xls, doc, ppt等 。目前POI的稳定版本中支持Excel文件格式(xls和xlsx)，其他的都属于不稳定版本（放在poi的scrachpad目录 中）。");

            m_p = m_Docx.Document.body.AddNewP();
            m_p.AddNewPPr().AddNewJc().val = ST_Jc.both;//段落两端对齐
            gp = new XWPFParagraph(m_p, m_Docx);
            gr = gp.CreateRun();
            gr.SetText("NPOI，顾名思义，就是POI的.NET版本。那POI又是什么呢？POI是一套用Java写成的库，能够帮助开 发者在没有安装微软Office的情况下读写Office 97-2003的文件，支持的文件格式包括xls, doc, ppt等 。目前POI的稳定版本中支持Excel文件格式(xls和xlsx)，其他的都属于不稳定版本（放在poi的scrachpad目录 中）。"); 
            if(parType=="inline")
            {
                //inline方式插图
                gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1000000, 1000000);
                m_p = m_Docx.Document.body.AddNewP();
                m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                gp = new XWPFParagraph(m_p, m_Docx);
                gr = gp.CreateRun();
                gr.SetText("inline插图");
            }
            else if (parType == "anchor")
            {
                //anchor方式插图
                CT_Anchor anchor = new CT_Anchor();
                //图片距正文上(distT)、下(distB)、左(distL)、右(distR)的距离。114300EMUS=3.1mm
                anchor.distT = 0u;
                anchor.distB = 0u;
                anchor.distL = 114300u;
                anchor.distR = 114300u;
                anchor.simplePos1 = false;
                anchor.relativeHeight = 251658240u;
                anchor.behindDoc = false;
                anchor.locked = false;
                anchor.layoutInCell = true;
                anchor.allowOverlap = true;

                CT_Positive2D simplePos = new CT_Positive2D();
                simplePos.x = 0;
                simplePos.y = 0;

                CT_EffectExtent effectExtent = new CT_EffectExtent();
                effectExtent.l = 0;
                effectExtent.t = 0;
                effectExtent.r = 0;
                effectExtent.b = 0;

                //图片与文字关系
                //四周型：CT_WrapSquare，紧密型：CT_WrapTight，穿越型：CT_WrapThrough
                if (wrapType == "wrapSquare")
                {
                    //四周型
                    gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = "宋体";
                    gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = "宋体";

                    gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)28;//四号
                    gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)28;

                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.both;
        
                    gr.GetCTR().AddNewRPr().AddNewB().val = true; //加粗      
                    gp.IndentationFirstLine = Indentation("宋体", 28, 2,FontStyle.Bold);
                    //图左上角坐标
                    CT_PosH posH = new CT_PosH();
                    posH.relativeFrom = ST_RelFromH.column;
                    posH.posOffset = 4000000;//单位：EMUS,1CM=360000EMUS
                    CT_PosV posV = new CT_PosV();
                    posV.relativeFrom = ST_RelFromV.paragraph;
                    posV.posOffset = 200000;
                    CT_WrapSquare wrapSquare = new CT_WrapSquare();
                    if (posType == "Besides")
                    {
                        //两侧
                         wrapSquare.wrapText = ST_WrapText.bothSides;
                    }
                    else if (posType == "Left")
                    {
                        //左侧
                        wrapSquare.wrapText = ST_WrapText.left;
                    }
                    else if (posType == "Right")
                    {
                        //右侧
                        wrapSquare.wrapText = ST_WrapText.right;
                    }
                    else if (posType == "Largest")
                    {
                        //最大一侧
                        wrapSquare.wrapText = ST_WrapText.largest;
                    }
                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1000000, 1000000, posH, posV, wrapSquare,anchor,simplePos,effectExtent);
                    m_p = m_Docx.Document.body.AddNewP();
                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                    gp = new XWPFParagraph(m_p, m_Docx);
                    gr = gp.CreateRun();
                    gr.SetText("anchor-wrapSquare(四周)-" + posType + "插图");
                }
                else if (wrapType == "wrapTight")
                {
                    //紧密型
                    gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = "宋体";
                    gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = "宋体";

                    gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)28;//四号
                    gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)28;

                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.both;
                    m_p.AddNewPPr().AddNewSpacing().line = "400";//行距固定20磅
                    m_p.AddNewPPr().AddNewSpacing().lineRule = ST_LineSpacingRule.exact; 

                    //gr.GetCTR().AddNewRPr().AddNewB().val = true; //加粗      
                    gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);

                    CT_WrapTight wrapTight = new CT_WrapTight();
                    if (posType == "Besides")
                    {
                        //两侧
                        wrapTight.wrapText = ST_WrapText.bothSides;
                    }
                    else if (posType == "Left")
                    {
                        //左侧
                        wrapTight.wrapText = ST_WrapText.left;
                    }
                    else if (posType == "Right")
                    {
                        //右侧
                        wrapTight.wrapText = ST_WrapText.right;
                    }
                    else if (posType == "Largest")
                    {
                        //最大一侧
                        wrapTight.wrapText = ST_WrapText.largest;
                    }

                    wrapTight.wrapPolygon = new CT_WrapPath();
                    wrapTight.wrapPolygon.edited = false;
                    wrapTight.wrapPolygon.start = new CT_Positive2D();
                    wrapTight.wrapPolygon.start.x = 0;
                    wrapTight.wrapPolygon.start.y = 0;
                    CT_Positive2D lineTo = new CT_Positive2D();
                    wrapTight.wrapPolygon.lineTo = new List<CT_Positive2D>();
                    lineTo = new CT_Positive2D();
                    lineTo.x = 0;
                    lineTo.y = 21394;
                    wrapTight.wrapPolygon.lineTo.Add(lineTo);
                    lineTo = new CT_Positive2D();
                    lineTo.x = 21806;
                    lineTo.y = 21394;
                    wrapTight.wrapPolygon.lineTo.Add(lineTo);
                    lineTo = new CT_Positive2D();
                    lineTo.x = 21806;
                    lineTo.y = 0;
                    wrapTight.wrapPolygon.lineTo.Add(lineTo);
                    lineTo = new CT_Positive2D();
                    lineTo.x = 0;
                    lineTo.y = 0;
                    wrapTight.wrapPolygon.lineTo.Add(lineTo);
                    //图位置
                    CT_PosH posH = new CT_PosH();
                    posH.relativeFrom = ST_RelFromH.column;
                    posH.posOffset = 4000000;
                    CT_PosV posV = new CT_PosV();
                    posV.relativeFrom = ST_RelFromV.paragraph;
                    posV.posOffset = -432000;//-1.2cm*360000

                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 720000, 720000, posH, posV, wrapTight, anchor, simplePos, effectExtent);
                    m_p = m_Docx.Document.body.AddNewP();
                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                    gp = new XWPFParagraph(m_p, m_Docx);
                    gr = gp.CreateRun();
                    gr.SetText("anchor-wrapTight(紧密)插图");
                }
                else if (wrapType == "wrapThrough")
                {
                    //穿越型-两边
                    gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = "宋体";
                    gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = "宋体";

                    gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)28;//四号
                    gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)28;

                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.both;
                    m_p.AddNewPPr().AddNewSpacing().line = "400";//行距固定20磅
                    m_p.AddNewPPr().AddNewSpacing().lineRule = ST_LineSpacingRule.exact; 

                    CT_WrapThrough wrapThrough = new CT_WrapThrough();
                    if (posType == "Besides")
                    {
                        //两侧
                        wrapThrough.wrapText = ST_WrapText.bothSides;
                    }
                    else if (posType == "Left")
                    {
                        //左侧
                        wrapThrough.wrapText = ST_WrapText.left;
                    }
                    else if (posType == "Right")
                    {
                        //右侧
                        wrapThrough.wrapText = ST_WrapText.right;
                    }
                    else if (posType == "Largest")
                    {
                        //最大一侧
                        wrapThrough.wrapText = ST_WrapText.largest;
                    }
                    wrapThrough.wrapPolygon = new CT_WrapPath();
                    wrapThrough.wrapPolygon.edited = false;
                    wrapThrough.wrapPolygon.start = new CT_Positive2D();
                    wrapThrough.wrapPolygon.start.x = 0;
                    wrapThrough.wrapPolygon.start.y = 0;
                    CT_Positive2D lineTo = new CT_Positive2D();
                    wrapThrough.wrapPolygon.lineTo = new List<CT_Positive2D>();
                    lineTo = new CT_Positive2D();
                    lineTo.x = 0;
                    lineTo.y = 21394;
                    wrapThrough.wrapPolygon.lineTo.Add(lineTo);
                    lineTo = new CT_Positive2D();
                    lineTo.x = 21806;
                    lineTo.y = 21394;
                    wrapThrough.wrapPolygon.lineTo.Add(lineTo);
                    lineTo = new CT_Positive2D();
                    lineTo.x = 21806;
                    lineTo.y = 0;
                    wrapThrough.wrapPolygon.lineTo.Add(lineTo);
                    lineTo = new CT_Positive2D();
                    lineTo.x = 0;
                    lineTo.y = 0;
                    wrapThrough.wrapPolygon.lineTo.Add(lineTo);
                    CT_PosH posH = new CT_PosH();
                    posH.relativeFrom = ST_RelFromH.column;
                    posH.posOffset = 4000000;
                    CT_PosV posV = new CT_PosV();
                    posV.relativeFrom = ST_RelFromV.paragraph;
                    posV.posOffset = -432000;//-1.2cm*360000

                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 720000, 720000, posH, posV, wrapThrough, anchor, simplePos, effectExtent);
                    //gp = m_Docx.CreateParagraph();
                    //gp.GetCTPPr().AddNewJc().val = ST_Jc.center; //水平居中
                    m_p = m_Docx.Document.body.AddNewP();
                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                    gp = new XWPFParagraph(m_p, m_Docx);
                    gr = gp.CreateRun();
                    gr.SetText("anchor-wrapThrough(穿越)插图");
                }
                else if (wrapType=="wrapTopAndBottom")
                {
                    //上下型
                    //图左上角坐标
                    CT_PosH posH = new CT_PosH();
                    posH.relativeFrom = ST_RelFromH.column;
                    posH.posOffset = 4000000;//单位：EMUS,1CM=360000EMUS
                    CT_PosV posV = new CT_PosV();
                    posV.relativeFrom = ST_RelFromV.paragraph;
                    posV.posOffset = 200000;
                    CT_WrapTopBottom wrapTopandBottom = new CT_WrapTopBottom();
                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1000000, 1000000, posH, posV, wrapTopandBottom, anchor, simplePos, effectExtent);
                    m_p = m_Docx.Document.body.AddNewP();
                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                    gp = new XWPFParagraph(m_p, m_Docx);
                    gr = gp.CreateRun();
                    gr.SetText("anchor-wrapTopAndBottom(上下)插图");
                }
                else if (wrapType == "wrapNoneBehindDoc")
                {
                    //上方型
                    //图左上角坐标
                    CT_PosH posH = new CT_PosH();
                    posH.relativeFrom = ST_RelFromH.column;
                    posH.posOffset = 4000000;//单位：EMUS,1CM=360000EMUS
                    CT_PosV posV = new CT_PosV();
                    posV.relativeFrom = ST_RelFromV.paragraph;
                    posV.posOffset = 0;
                    CT_WrapNone wrapNone = new CT_WrapNone();
                    anchor.behindDoc = false;
                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1000000, 1000000, posH, posV, wrapNone, anchor, simplePos, effectExtent);
                    m_p = m_Docx.Document.body.AddNewP();
                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                    gp = new XWPFParagraph(m_p, m_Docx);
                    gr = gp.CreateRun();
                    gr.SetText("anchor-wrapNoneBehindDoc插图");
                }
                else if (wrapType == "wrapBehindDoc")
                {
                    //下方型
                    //图左上角坐标
                    CT_PosH posH = new CT_PosH();
                    posH.relativeFrom = ST_RelFromH.column;
                    posH.posOffset = 4000000;//单位：EMUS,1CM=360000EMUS
                    CT_PosV posV = new CT_PosV();
                    posV.relativeFrom = ST_RelFromV.paragraph;
                    posV.posOffset = 0;
                    CT_WrapNone wrapNone = new CT_WrapNone();
                    anchor.behindDoc = true;
                    gr.AddPicture(gfs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "1.jpg", 1000000, 1000000, posH, posV, wrapNone, anchor, simplePos, effectExtent);
                    m_p = m_Docx.Document.body.AddNewP();
                    m_p.AddNewPPr().AddNewJc().val = ST_Jc.center;//段落水平居中
                    gp = new XWPFParagraph(m_p, m_Docx);
                    gr = gp.CreateRun();
                    gr.SetText("anchor-wrapBehindDoc插图");
                }

            }
            gfs.Close();
            return m_Docx;
        }
        static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();

                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }
        }
        protected int Indentation(String fontname, int fontsize, int Indentationfonts,FontStyle fs)
        {
            //字显示宽度，用于段首行缩进
            /*字号与fontsize关系
             * 初号（0号）=84，小初=72，1号=52，2号=44，小2=36，3号=32，小3=30，4号=28，小4=24，5号=21，小5=18，6号=15，小6=13，7号=11，8号=10
             */
            Graphics m_tmpGr = this.CreateGraphics();
            m_tmpGr.PageUnit = GraphicsUnit.Point;
            SizeF size = m_tmpGr.MeasureString("好", new Font(fontname, fontsize * 0.75F, fs));
            return (int)size.Width * Indentationfonts * 10;
        }




  
    }
}
