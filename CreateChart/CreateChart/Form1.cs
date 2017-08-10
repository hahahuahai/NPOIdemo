using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.OpenXmlFormats.Dml.Chart;
using NPOI.OpenXmlFormats.Dml.WordProcessing;


/*
 * 本例子提供的NPOI是tonyqus提供的2.1.1.0源码经过修改编译。
 * 例中包括：
 * 1、页眉页脚设置
 * 2、插图表操作：分inline和anchor两种方式，提供饼图和柱状图实例，其它图表没有提供实例
 * vs2010
 * netframework4
 * 创建的docx在word2007可以打开
 * 2014-9-18
 * 
 */
namespace CreateChart
{
    public partial class Form1 : Form
    {
        const String m_savefilepath = "d:\\doc\\NPOI";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //inline
            /*
             * 创建饼图
             */
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = CreatepieCharttoDocxwithinline();
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\NPOIChart\\CreatepieChartwithinline.docx");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //inline
            /*
             * 创建柱状图
             */

            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = CreatebarCharttoDocxwithinline();
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\NPOIChart\\CreatebarChartwithinline.docx");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //anchor方式的饼图
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            m_Docx = CreateCharttoDocxwithAnchor();
            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, m_savefilepath + "\\NPOIChart\\Chartwithanchor.docx");

        }

        protected XWPFDocument CreatepieCharttoDocxwithinline()
        {
            //inline
            //创建不同设计的饼图
            XWPFDocument m_Docx = new XWPFDocument();
            //页面设置
            //A4:W=11906,h=16838
            //CT_SectPr m_SectPr = m_Docx.Document.body.AddNewSectPr();
            m_Docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_SectPr = m_Docx.Document.body.sectPr;
            //页面设置A4横向
            m_SectPr.pgSz.w = (ulong)16838;
            m_SectPr.pgSz.h = (ulong)11906;

            //创建页脚
            CT_Ftr m_ftr = new CT_Ftr();
            m_ftr.AddNewP().AddNewR().AddNewT().Value = "fff";//页脚内容
            //创建页脚关系（footern.xml）
            XWPFRelation Frelation = XWPFRelation.FOOTER;
            XWPFFooter m_f = (XWPFFooter)m_Docx.CreateRelationship(Frelation, XWPFFactory.GetInstance(), m_Docx.FooterList.Count + 1);
            //设置页脚
            m_f.SetHeaderFooter(m_ftr);
            CT_HdrFtrRef m_HdrFtr = m_SectPr.AddNewFooterReference();
            m_HdrFtr.type = ST_HdrFtr.@default;
            m_HdrFtr.id = m_f.GetPackageRelationship().Id;

            //创建页眉
            CT_Hdr m_Hdr = new CT_Hdr();
            m_Hdr.AddNewP().AddNewR().AddNewT().Value = "hhh";//页眉内容
            //创建页眉关系（headern.xml）
            XWPFRelation Hrelation = XWPFRelation.HEADER;
            XWPFHeader m_h = (XWPFHeader)m_Docx.CreateRelationship(Hrelation, XWPFFactory.GetInstance(), m_Docx.HeaderList.Count + 1);
            //设置页眉
            m_h.SetHeaderFooter(m_Hdr);
            m_HdrFtr = m_SectPr.AddNewHeaderReference();
            m_HdrFtr.type = ST_HdrFtr.@default;
            m_HdrFtr.id = m_h.GetPackageRelationship().Id;

            //插入图表（饼图）
            //插入xlsx
            //创建xlsx
            XSSFWorkbook workbook = new XSSFWorkbook();
            //创建表单1（饼图）
            ISheet sheet = workbook.CreateSheet("Sheet1");
            //表单1饼图数据
            //         销售额
            //第一季度 8.2
            //第二季度 3.2
            //第三季度 1.4
            //第四季度 1.2

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell = row.CreateCell(0);
            cell = row.CreateCell(1);
            cell.SetCellValue("销售额");
            row = sheet.CreateRow(1);
            cell = row.CreateCell(0);
            cell.SetCellValue("第一季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(8.2);
            row = sheet.CreateRow(2);
            cell = row.CreateCell(0);
            cell.SetCellValue("第二季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(3.2);
            row = sheet.CreateRow(3);
            cell = row.CreateCell(0);
            cell.SetCellValue("第三季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(1.4);
            row = sheet.CreateRow(4);
            cell = row.CreateCell(0);
            cell.SetCellValue("第四季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(1.2);

            //将xlsx数据转为btye，因为在插入第一张图表后workbook流会关闭
            MemoryStream msxlsxData = new MemoryStream();
            workbook.Write(msxlsxData);
            msxlsxData.Flush();
            byte[] bxlsxData = msxlsxData.ToArray();
 
            //创建\word\charts\chartn.xml内容（简单饼图）
            CT_ChartSpace ctpiechartspace = new CT_ChartSpace();

            ctpiechartspace.date1904 = new CT_Boolean();
            ctpiechartspace.date1904.val = 1;
            ctpiechartspace.lang = new CT_TextLanguageID();
            ctpiechartspace.lang.val = "zh-CN";

            CT_Chart m_chart = ctpiechartspace.AddNewChart();
            m_chart.plotArea = new CT_PlotArea();
            m_chart.plotArea.pieChart = new List<CT_PieChart>();
            //饼图
            CT_PieChart m_piechart = new CT_PieChart();
            m_piechart.varyColors = new CT_Boolean();
            m_piechart.varyColors.val = 1;
            m_piechart.ser = new List<CT_PieSer>();
            CT_PieSer m_pieser = new CT_PieSer();
            //标题
            m_pieser.tx = new CT_SerTx();
            m_pieser.tx.strRef = new CT_StrRef();
            m_pieser.tx.strRef.f = "Sheet1!$B$1";
            m_pieser.tx.strRef.strCache = new CT_StrData();
            m_pieser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_pieser.tx.strRef.strCache.ptCount.val = 1;
            CT_StrVal m_strval = new CT_StrVal();
            m_strval.idx = 0;
            m_strval.v = "销售额";
            m_pieser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_pieser.tx.strRef.strCache.pt.Add(m_strval); 
            //行标题
            m_pieser.cat = new CT_AxDataSource();
            m_pieser.cat.strRef = new CT_StrRef();
            m_pieser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_pieser.cat.strRef.strCache = new CT_StrData();
            m_pieser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_pieser.cat.strRef.strCache.ptCount.val = 4;
            m_pieser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_strval = new CT_StrVal();
            m_strval.idx = 0;
            m_strval.v = "第一季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 1;
            m_strval.v = "第二季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 2;
            m_strval.v = "第三季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 3;
            m_strval.v = "第四季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            //值
            m_pieser.val = new CT_NumDataSource();
            m_pieser.val.numRef = new CT_NumRef();
            m_pieser.val.numRef.f = "Sheet1!$B$2:$B$5";
            m_pieser.val.numRef.numCache = new CT_NumData();
            m_pieser.val.numRef.numCache.formatCode = "General";
            m_pieser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_pieser.val.numRef.numCache.ptCount.val = 4;
            m_pieser.val.numRef.numCache.pt = new List<CT_NumVal>();
            CT_NumVal m_numval = new CT_NumVal();
            m_numval.idx = 0;
            m_numval.v = "8.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 1;
            m_numval.v = "3.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 2;
            m_numval.v = "1.4";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 3;
            m_numval.v = "1.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval); 
            m_piechart.ser.Add(m_pieser);

            m_chart.plotArea.pieChart.Add(m_piechart);


            m_chart.legend = new CT_Legend();
            m_chart.legend.legendPos = new CT_LegendPos();
            m_chart.legend.legendPos.val = ST_LegendPos.r;
            m_chart.plotVisOnly = new CT_Boolean();
            m_chart.plotVisOnly.val = 1;

            XWPFParagraph gp = m_Docx.CreateParagraph();
            XWPFRun gr = gp.CreateRun();
            gp = m_Docx.CreateParagraph();
            gr = gp.CreateRun();
            gr.AddChartSpace(new XSSFWorkbook(new MemoryStream(bxlsxData)) , ctpiechartspace, 5274310, 3076575);

            gr.AddBreak();//分页 

            //创建\word\charts\chartn.xml内容（格式饼图）
            ctpiechartspace = new CT_ChartSpace();

            ctpiechartspace.date1904 = new CT_Boolean();
            ctpiechartspace.date1904.val = 1;
            ctpiechartspace.lang = new CT_TextLanguageID();
            ctpiechartspace.lang.val = "zh-CN";

            m_chart = ctpiechartspace.AddNewChart();
            //图表标题
            m_chart.title = new CT_Title();
            //省略，则标题采用tx的值
            m_chart.title.tx = new CT_Tx();
            m_chart.title.tx.rich = new CT_TextBody();
            m_chart.title.tx.rich.AddNewBodyPr();
            m_chart.title.tx.rich.AddNewLstStyle();
            m_chart.title.tx.rich.p = new List<NPOI.OpenXmlFormats.Dml.CT_TextParagraph>();
            NPOI.OpenXmlFormats.Dml.CT_TextParagraph charttp = new NPOI.OpenXmlFormats.Dml.CT_TextParagraph();
            charttp.AddNewPPr().defRPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            NPOI.OpenXmlFormats.Dml.CT_RegularTextRun r = charttp.AddNewR();
            r.rPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            r.rPr.lang = "zh-CN";
            r.rPr.altLang = "en-US";
            r.t = "销售额饼图";
            m_chart.title.tx.rich.p.Add(charttp);

            m_chart.title.overlay = new CT_Boolean();
            m_chart.title.overlay.val = 0;
            m_chart.autoTitleDeleted = new CT_Boolean();
            m_chart.autoTitleDeleted.val = 0; 

            m_chart.plotArea = new CT_PlotArea();
            m_chart.plotArea.pieChart = new List<CT_PieChart>();
            //饼图
            m_piechart = new CT_PieChart();
            m_piechart.varyColors = new CT_Boolean();
            m_piechart.varyColors.val = 1;
            m_piechart.ser = new List<CT_PieSer>();
            m_pieser = new CT_PieSer();
            //m_piechart.ser.Add(m_pieser);
            //标题
            m_pieser.tx = new CT_SerTx();
            m_pieser.tx.strRef = new CT_StrRef();
            m_pieser.tx.strRef.f = "Sheet1!$B$1";
            m_pieser.tx.strRef.strCache = new CT_StrData();
            m_pieser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_pieser.tx.strRef.strCache.ptCount.val = 1;
            m_strval = new CT_StrVal();
            m_strval.idx = 0;
            m_strval.v = "销售额";
            m_pieser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_pieser.tx.strRef.strCache.pt.Add(m_strval);


            //m_pieser.dLbls = new CT_DLbls();
            //m_pieser.dLbls.showLegendKey = new CT_Boolean() ;
            //m_pieser.dLbls.showLegendKey.val = 0;
            //m_pieser.dLbls.showVal = new CT_Boolean();
            //m_pieser.dLbls.showVal.val = 0; 
            //m_pieser.dLbls.showCatName = new CT_Boolean();
            //m_pieser.dLbls.showCatName.val = 0; 
            //m_pieser.dLbls.showSerName = new CT_Boolean();
            //m_pieser.dLbls.showSerName.val = 0; 
            //m_pieser.dLbls.showPercent = new CT_Boolean();
            //m_pieser.dLbls.showPercent.val = 1;
            //m_pieser.dLbls.showBubbleSize = new CT_Boolean();
            //m_pieser.dLbls.showBubbleSize.val = 0;
            //m_pieser.dLbls.showLeaderLines = new CT_Boolean();
            //m_pieser.dLbls.showLeaderLines.val = 1;

            //行标题
            m_pieser.cat = new CT_AxDataSource();
            m_pieser.cat.strRef = new CT_StrRef();
            m_pieser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_pieser.cat.strRef.strCache = new CT_StrData();
            m_pieser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_pieser.cat.strRef.strCache.ptCount.val = 4;
            m_pieser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_strval = new CT_StrVal();
            m_strval.idx = 0;
            m_strval.v = "第一季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 1;
            m_strval.v = "第二季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 2;
            m_strval.v = "第三季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 3;
            m_strval.v = "第四季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            //值
            m_pieser.val = new CT_NumDataSource();
            m_pieser.val.numRef = new CT_NumRef();
            m_pieser.val.numRef.f = "Sheet1!$B$2:$B$5";
            m_pieser.val.numRef.numCache = new CT_NumData();
            m_pieser.val.numRef.numCache.formatCode = "General";
            m_pieser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_pieser.val.numRef.numCache.ptCount.val = 4;
            m_pieser.val.numRef.numCache.pt = new List<CT_NumVal>();
            m_numval = new CT_NumVal();
            m_numval.idx = 0;
            m_numval.v = "8.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 1;
            m_numval.v = "3.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 2;
            m_numval.v = "1.4";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 3;
            m_numval.v = "1.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_piechart.ser.Add(m_pieser);

            //图表与数据格式显示设计（通过不同项目置的设置可以得到不同饼图的显示格式）
            m_piechart.dLbls = new CT_DLbls();
            m_piechart.dLbls.showLegendKey = new CT_Boolean();
            m_piechart.dLbls.showLegendKey.val = 0;
            m_piechart.dLbls.showVal = new CT_Boolean();//值显示设置
            m_piechart.dLbls.showVal.val = 0;
            m_piechart.dLbls.showCatName = new CT_Boolean();//分类显示设置
            m_piechart.dLbls.showCatName.val = 0;
            m_piechart.dLbls.showSerName = new CT_Boolean();//tx标题显示设置
            m_piechart.dLbls.showSerName.val = 0;
            m_piechart.dLbls.showPercent = new CT_Boolean();
            m_piechart.dLbls.showPercent.val = 1;
            m_piechart.dLbls.showBubbleSize = new CT_Boolean();
            m_piechart.dLbls.showBubbleSize.val = 0;
            m_piechart.dLbls.showLeaderLines = new CT_Boolean();
            m_piechart.dLbls.showLeaderLines.val = 1;

            m_piechart.firstSliceAng = new CT_FirstSliceAng();
            m_piechart.firstSliceAng.val = 0;

            m_chart.plotArea.pieChart.Add(m_piechart);

            //图例
            m_chart.legend = new CT_Legend();
            m_chart.legend.legendPos = new CT_LegendPos();
            m_chart.legend.legendPos.val = ST_LegendPos.t;//图例在上方
            m_chart.plotVisOnly = new CT_Boolean();
            m_chart.plotVisOnly.val = 1;

            gp = m_Docx.CreateParagraph();
            gr = gp.CreateRun();
            gp = m_Docx.CreateParagraph();
            gr = gp.CreateRun();
            gr.AddChartSpace(new XSSFWorkbook(new MemoryStream(bxlsxData)), ctpiechartspace, 5274310, 3076575);

            return m_Docx;
        }

        protected XWPFDocument CreatebarCharttoDocxwithinline()
        {
            //inline
            //不同柱状图设计
            XWPFDocument m_Docx = new XWPFDocument();
            //页面设置
            //A4:W=11906,h=16838
            //CT_SectPr m_SectPr = m_Docx.Document.body.AddNewSectPr();
            m_Docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_SectPr = m_Docx.Document.body.sectPr;
            //页面设置A4横向
            m_SectPr.pgSz.w = (ulong)16838;
            m_SectPr.pgSz.h = (ulong)11906;

            //创建页脚
            CT_Ftr m_ftr = new CT_Ftr();
            m_ftr.AddNewP().AddNewR().AddNewT().Value = "fff";//页脚内容
            //创建页脚关系（footern.xml）
            XWPFRelation Frelation = XWPFRelation.FOOTER;
            XWPFFooter m_f = (XWPFFooter)m_Docx.CreateRelationship(Frelation, XWPFFactory.GetInstance(), m_Docx.FooterList.Count + 1);
            //设置页脚
            m_f.SetHeaderFooter(m_ftr);
            CT_HdrFtrRef m_HdrFtr = m_SectPr.AddNewFooterReference();
            m_HdrFtr.type = ST_HdrFtr.@default;
            m_HdrFtr.id = m_f.GetPackageRelationship().Id;

            //创建页眉
            CT_Hdr m_Hdr = new CT_Hdr();
            m_Hdr.AddNewP().AddNewR().AddNewT().Value = "hhh";//页眉内容
            //创建页眉关系（headern.xml）
            XWPFRelation Hrelation = XWPFRelation.HEADER;
            XWPFHeader m_h = (XWPFHeader)m_Docx.CreateRelationship(Hrelation, XWPFFactory.GetInstance(), m_Docx.HeaderList.Count + 1);
            //设置页眉
            m_h.SetHeaderFooter(m_Hdr);
            m_HdrFtr = m_SectPr.AddNewHeaderReference();
            m_HdrFtr.type = ST_HdrFtr.@default;
            m_HdrFtr.id = m_h.GetPackageRelationship().Id;


            //插入柱状图图表
            //插入xlsx
            //创建xlsx
            XSSFWorkbook workbook = new XSSFWorkbook();

            //创建表单1（柱状图）
            ISheet sheet = workbook.CreateSheet("Sheet1");
            //5、表单1柱状图数据
            //     系列1	系列2 系列3
            //类别1 4.3  2.4   2
            //类别2 2.5  4.4   2
            //类别3 3.5  1.8   3
            //类别4 4.5  2.8   5

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell = row.CreateCell(1);
            cell.SetCellValue("系列 1");
            cell = row.CreateCell(2);
            cell.SetCellValue("系列 2");
            cell = row.CreateCell(3);
            cell.SetCellValue("系列 3");
            row = sheet.CreateRow(1);
            cell = row.CreateCell(0);
            cell.SetCellValue("类别 1");
            cell = row.CreateCell(1);
            cell.SetCellValue(4.3);
            cell = row.CreateCell(2);
            cell.SetCellValue(2.4);
            cell = row.CreateCell(3);
            cell.SetCellValue(2);
            row = sheet.CreateRow(2);
            cell = row.CreateCell(0);
            cell.SetCellValue("类别 2");
            cell = row.CreateCell(1);
            cell.SetCellValue(2.5);
            cell = row.CreateCell(2);
            cell.SetCellValue(4.4);
            cell = row.CreateCell(3);
            cell.SetCellValue(2);
            row = sheet.CreateRow(3);
            cell = row.CreateCell(0);
            cell.SetCellValue("类别 3");
            cell = row.CreateCell(1);
            cell.SetCellValue(3.5);
            cell = row.CreateCell(2);
            cell.SetCellValue(1.8);
            cell = row.CreateCell(3);
            cell.SetCellValue(3);
            row = sheet.CreateRow(4);
            cell = row.CreateCell(0);
            cell.SetCellValue("类别 4");
            cell = row.CreateCell(1);
            cell.SetCellValue(4.5);
            cell = row.CreateCell(2);
            cell.SetCellValue(2.8);
            cell = row.CreateCell(3);
            cell.SetCellValue(5);

            //将xlsx数据转为btye，因为在插入第一张图表后workbook流会关闭
            MemoryStream msxlsxData = new MemoryStream();
            workbook.Write(msxlsxData);
            msxlsxData.Flush();
            byte[] bxlsxData = msxlsxData.ToArray();

            //简单柱状图chartn.xml内容
            CT_ChartSpace ctbarchartspace = new CT_ChartSpace();

            ctbarchartspace.date1904 = new CT_Boolean();
            ctbarchartspace.date1904.val = 1;
            ctbarchartspace.lang = new CT_TextLanguageID();
            ctbarchartspace.lang.val = "zh-CN";

            CT_Chart m_chart = ctbarchartspace.AddNewChart();
            m_chart.plotArea = new CT_PlotArea();
            m_chart.plotArea.barChart = new List<CT_BarChart>();

            CT_BarChart m_barchart = new CT_BarChart();
            m_barchart.barDir = new CT_BarDir();
            m_barchart.barDir.val = ST_BarDir.col;
            m_barchart.grouping = new CT_BarGrouping();
            m_barchart.grouping.val = ST_BarGrouping.clustered;
            m_barchart.ser = new List<CT_BarSer>();

            CT_BarSer m_barser = new CT_BarSer();
            m_barser.idx = new CT_UnsignedInt();
            m_barser.idx.val = 0;
            m_barser.order = new CT_UnsignedInt();
            m_barser.order.val = 0;
            m_barser.tx = new CT_SerTx();
            m_barser.tx.strRef = new CT_StrRef();
            m_barser.tx.strRef.f = "Sheet1!$B$1";
            m_barser.tx.strRef.strCache = new CT_StrData();
            m_barser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.tx.strRef.strCache.ptCount.val = 1;
            m_barser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            CT_StrVal m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "系列 1";
            m_barser.tx.strRef.strCache.pt.Add(m_barpt);
            m_barser.cat = new CT_AxDataSource();
            m_barser.cat.strRef = new CT_StrRef();
            m_barser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_barser.cat.strRef.strCache = new CT_StrData();
            m_barser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.cat.strRef.strCache.ptCount.val = 4;
            m_barser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "类别 1";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 1;
            m_barpt.v = "类别 2";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 2;
            m_barpt.v = "类别 3";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 3;
            m_barpt.v = "类别 4";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barser.val = new CT_NumDataSource();
            m_barser.val.numRef = new CT_NumRef();
            m_barser.val.numRef.f = "Sheet1!$B$2:$B$5";
            m_barser.val.numRef.numCache = new CT_NumData();
            m_barser.val.numRef.numCache.formatCode = "General";
            m_barser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_barser.val.numRef.numCache.ptCount.val = 4;
            m_barser.val.numRef.numCache.pt = new List<CT_NumVal>();
            CT_NumVal m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 0;
            m_barvalpt.v = "4.3";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 1;
            m_barvalpt.v = "2.5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 2;
            m_barvalpt.v = "3.5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 3;
            m_barvalpt.v = "4.5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barchart.ser.Add(m_barser);

            m_barser = new CT_BarSer();
            m_barser.idx = new CT_UnsignedInt();
            m_barser.idx.val = 1;
            m_barser.order = new CT_UnsignedInt();
            m_barser.order.val = 1;
            m_barser.tx = new CT_SerTx();
            m_barser.tx.strRef = new CT_StrRef();
            m_barser.tx.strRef.f = "Sheet1!$C$1";
            m_barser.tx.strRef.strCache = new CT_StrData();
            m_barser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.tx.strRef.strCache.ptCount.val = 1;
            m_barser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "系列 2";
            m_barser.tx.strRef.strCache.pt.Add(m_barpt);
            m_barser.cat = new CT_AxDataSource();
            m_barser.cat.strRef = new CT_StrRef();
            m_barser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_barser.cat.strRef.strCache = new CT_StrData();
            m_barser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.cat.strRef.strCache.ptCount.val = 4;
            m_barser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "类别 1";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 1;
            m_barpt.v = "类别 2";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 2;
            m_barpt.v = "类别 3";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 3;
            m_barpt.v = "类别 4";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barser.val = new CT_NumDataSource();
            m_barser.val.numRef = new CT_NumRef();
            m_barser.val.numRef.f = "Sheet1!$C$2:$C$5";
            m_barser.val.numRef.numCache = new CT_NumData();
            m_barser.val.numRef.numCache.formatCode = "General";
            m_barser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_barser.val.numRef.numCache.ptCount.val = 4;
            m_barser.val.numRef.numCache.pt = new List<CT_NumVal>();
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 0;
            m_barvalpt.v = "2.4";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 1;
            m_barvalpt.v = "4.4";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 2;
            m_barvalpt.v = "1.8";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 3;
            m_barvalpt.v = "2.8";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barchart.ser.Add(m_barser);

            m_barser = new CT_BarSer();
            m_barser.idx = new CT_UnsignedInt();
            m_barser.idx.val = 2;
            m_barser.order = new CT_UnsignedInt();
            m_barser.order.val = 2;
            m_barser.tx = new CT_SerTx();
            m_barser.tx.strRef = new CT_StrRef();
            m_barser.tx.strRef.f = "Sheet1!$D$1";
            m_barser.tx.strRef.strCache = new CT_StrData();
            m_barser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.tx.strRef.strCache.ptCount.val = 1;
            m_barser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "系列 3";
            m_barser.tx.strRef.strCache.pt.Add(m_barpt);
            m_barser.cat = new CT_AxDataSource();
            m_barser.cat.strRef = new CT_StrRef();
            m_barser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_barser.cat.strRef.strCache = new CT_StrData();
            m_barser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.cat.strRef.strCache.ptCount.val = 4;
            m_barser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "类别 1";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 1;
            m_barpt.v = "类别 2";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 2;
            m_barpt.v = "类别 3";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 3;
            m_barpt.v = "类别 4";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barser.val = new CT_NumDataSource();
            m_barser.val.numRef = new CT_NumRef();
            m_barser.val.numRef.f = "Sheet1!$D$2:$D$5";
            m_barser.val.numRef.numCache = new CT_NumData();
            m_barser.val.numRef.numCache.formatCode = "General";
            m_barser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_barser.val.numRef.numCache.ptCount.val = 4;
            m_barser.val.numRef.numCache.pt = new List<CT_NumVal>();
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 0;
            m_barvalpt.v = "2";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 1;
            m_barvalpt.v = "2";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 2;
            m_barvalpt.v = "3";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 3;
            m_barvalpt.v = "5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barchart.ser.Add(m_barser);

            m_barchart.axId = new List<CT_UnsignedInt>();
            CT_UnsignedInt m_axId = new CT_UnsignedInt();
            m_axId.val = 57733120;
            m_barchart.axId.Add(m_axId);
            m_axId = new CT_UnsignedInt();
            m_axId.val = 57734656;
            m_barchart.axId.Add(m_axId);

            m_chart.plotArea.barChart.Add(m_barchart);

            m_chart.plotArea.catAx = new List<CT_CatAx>();
            CT_CatAx m_catAx = new CT_CatAx();
            m_catAx.axId = new CT_UnsignedInt();
            m_catAx.axId.val = 57733120;
            m_catAx.scaling = new CT_Scaling();
            m_catAx.scaling.orientation = new CT_Orientation();
            m_catAx.scaling.orientation.val = ST_Orientation.minMax;
            m_catAx.axPos = new CT_AxPos();
            m_catAx.axPos.val = ST_AxPos.b;
            m_catAx.tickLblPos = new CT_TickLblPos();
            m_catAx.tickLblPos.val = ST_TickLblPos.nextTo;
            m_catAx.crossAx = new CT_UnsignedInt();
            m_catAx.crossAx.val = 57733120;
            m_catAx.crosses = new CT_Crosses();
            m_catAx.crosses.val = ST_Crosses.autoZero;
            m_chart.plotArea.catAx.Add(m_catAx);

            m_chart.plotArea.valAx = new List<CT_ValAx>();
            CT_ValAx m_valAx = new CT_ValAx();
            m_valAx.axId = new CT_UnsignedInt();
            m_valAx.axId.val = 57734656;
            m_valAx.scaling = new CT_Scaling();
            m_valAx.scaling.orientation = new CT_Orientation();
            m_valAx.scaling.orientation.val = ST_Orientation.minMax;
            m_valAx.axPos = new CT_AxPos();
            m_valAx.axPos.val = ST_AxPos.l;
            m_valAx.majorGridlines = new CT_ChartLines();
            m_valAx.numFmt = new NPOI.OpenXmlFormats.Dml.Chart.CT_NumFmt();
            m_valAx.numFmt.formatCode = "General";
            m_valAx.numFmt.sourceLinked = true;
            m_valAx.tickLblPos = new CT_TickLblPos();
            m_valAx.tickLblPos.val = ST_TickLblPos.nextTo;
            m_valAx.crossAx = new CT_UnsignedInt();
            m_valAx.crossAx.val = 57733120;
            m_valAx.crosses = new CT_Crosses();
            m_valAx.crosses.val = ST_Crosses.autoZero;
            m_chart.plotArea.valAx.Add(m_valAx);

            m_chart.legend = new CT_Legend();
            m_chart.legend.legendPos = new CT_LegendPos();
            m_chart.legend.legendPos.val = ST_LegendPos.r;
            m_chart.plotVisOnly = new CT_Boolean();
            m_chart.plotVisOnly.val = 1;

            XWPFParagraph gp = m_Docx.CreateParagraph();
            XWPFRun gr = gp.CreateRun();
            gr.AddChartSpace(new XSSFWorkbook(new MemoryStream(bxlsxData)), ctbarchartspace, 5274310, 3076575);


            //创建\word\charts\chartn.xml内容（格式柱状图）
            ctbarchartspace = new CT_ChartSpace();

            ctbarchartspace.date1904 = new CT_Boolean();
            ctbarchartspace.date1904.val = 1;
            ctbarchartspace.lang = new CT_TextLanguageID();
            ctbarchartspace.lang.val = "zh-CN";

            m_chart = ctbarchartspace.AddNewChart();

            //图表标题
            m_chart.title = new CT_Title();
            m_chart.title.tx = new CT_Tx();
            m_chart.title.tx.rich = new CT_TextBody();
            m_chart.title.tx.rich.AddNewBodyPr();
            m_chart.title.tx.rich.AddNewLstStyle();
            m_chart.title.tx.rich.p = new List<NPOI.OpenXmlFormats.Dml.CT_TextParagraph>();
            NPOI.OpenXmlFormats.Dml.CT_TextParagraph charttp = new NPOI.OpenXmlFormats.Dml.CT_TextParagraph();
            charttp.AddNewPPr().defRPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            NPOI.OpenXmlFormats.Dml.CT_RegularTextRun r = charttp.AddNewR();
            r.rPr = new NPOI.OpenXmlFormats.Dml.CT_TextCharacterProperties();
            r.rPr.lang = "zh-CN";
            r.rPr.altLang = "en-US";
            r.t = "柱状图";
            m_chart.title.tx.rich.p.Add(charttp);

            m_chart.title.overlay = new CT_Boolean();
            m_chart.title.overlay.val = 0;
            m_chart.autoTitleDeleted = new CT_Boolean();
            m_chart.autoTitleDeleted.val = 0;

            m_chart.plotArea = new CT_PlotArea();
            m_chart.plotArea.AddNewLayout();
            m_chart.plotArea.barChart = new List<CT_BarChart>();

            m_barchart = new CT_BarChart();
            m_barchart.barDir = new CT_BarDir();
            m_barchart.barDir.val = ST_BarDir.col;
            m_barchart.grouping = new CT_BarGrouping();
            m_barchart.grouping.val = ST_BarGrouping.clustered;
            m_barchart.ser = new List<CT_BarSer>();

            m_barser = new CT_BarSer();
            m_barser.idx = new CT_UnsignedInt();
            m_barser.idx.val = 0;
            m_barser.order = new CT_UnsignedInt();
            m_barser.order.val = 0;
            m_barser.tx = new CT_SerTx();
            m_barser.tx.strRef = new CT_StrRef();
            m_barser.tx.strRef.f = "Sheet1!$B$1";
            m_barser.tx.strRef.strCache = new CT_StrData();
            m_barser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.tx.strRef.strCache.ptCount.val = 1;
            m_barser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "系列 1";
            m_barser.tx.strRef.strCache.pt.Add(m_barpt);

            //分类标题显示
            m_barser.invertIfNegative = new CT_Boolean();
            m_barser.invertIfNegative.val = 0; 

            m_barser.cat = new CT_AxDataSource();
            m_barser.cat.strRef = new CT_StrRef();
            m_barser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_barser.cat.strRef.strCache = new CT_StrData();
            m_barser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.cat.strRef.strCache.ptCount.val = 4;
            m_barser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "类别 1";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 1;
            m_barpt.v = "类别 2";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 2;
            m_barpt.v = "类别 3";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 3;
            m_barpt.v = "类别 4";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barser.val = new CT_NumDataSource();
            m_barser.val.numRef = new CT_NumRef();
            m_barser.val.numRef.f = "Sheet1!$B$2:$B$5";
            m_barser.val.numRef.numCache = new CT_NumData();
            m_barser.val.numRef.numCache.formatCode = "General";
            m_barser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_barser.val.numRef.numCache.ptCount.val = 4;
            m_barser.val.numRef.numCache.pt = new List<CT_NumVal>();
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 0;
            m_barvalpt.v = "4.3";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 1;
            m_barvalpt.v = "2.5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 2;
            m_barvalpt.v = "3.5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 3;
            m_barvalpt.v = "4.5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barchart.ser.Add(m_barser);

            m_barser = new CT_BarSer();
            m_barser.idx = new CT_UnsignedInt();
            m_barser.idx.val = 1;
            m_barser.order = new CT_UnsignedInt();
            m_barser.order.val = 1;
            m_barser.tx = new CT_SerTx();
            m_barser.tx.strRef = new CT_StrRef();
            m_barser.tx.strRef.f = "Sheet1!$C$1";
            m_barser.tx.strRef.strCache = new CT_StrData();
            m_barser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.tx.strRef.strCache.ptCount.val = 1;
            m_barser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "系列 2";
            m_barser.tx.strRef.strCache.pt.Add(m_barpt);

            //分类标题
            m_barser.invertIfNegative = new CT_Boolean();
            m_barser.invertIfNegative.val = 0; 

            m_barser.cat = new CT_AxDataSource();
            m_barser.cat.strRef = new CT_StrRef();
            m_barser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_barser.cat.strRef.strCache = new CT_StrData();
            m_barser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.cat.strRef.strCache.ptCount.val = 4;
            m_barser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "类别 1";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 1;
            m_barpt.v = "类别 2";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 2;
            m_barpt.v = "类别 3";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 3;
            m_barpt.v = "类别 4";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barser.val = new CT_NumDataSource();
            m_barser.val.numRef = new CT_NumRef();
            m_barser.val.numRef.f = "Sheet1!$C$2:$C$5";
            m_barser.val.numRef.numCache = new CT_NumData();
            m_barser.val.numRef.numCache.formatCode = "General";
            m_barser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_barser.val.numRef.numCache.ptCount.val = 4;
            m_barser.val.numRef.numCache.pt = new List<CT_NumVal>();
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 0;
            m_barvalpt.v = "2.4";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 1;
            m_barvalpt.v = "4.4";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 2;
            m_barvalpt.v = "1.8";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 3;
            m_barvalpt.v = "2.8";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barchart.ser.Add(m_barser);

            m_barser = new CT_BarSer();
            m_barser.idx = new CT_UnsignedInt();
            m_barser.idx.val = 2;
            m_barser.order = new CT_UnsignedInt();
            m_barser.order.val = 2;
            m_barser.tx = new CT_SerTx();
            m_barser.tx.strRef = new CT_StrRef();
            m_barser.tx.strRef.f = "Sheet1!$D$1";
            m_barser.tx.strRef.strCache = new CT_StrData();
            m_barser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.tx.strRef.strCache.ptCount.val = 1;
            m_barser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "系列 3";
            m_barser.tx.strRef.strCache.pt.Add(m_barpt);

            //分类标题
            m_barser.invertIfNegative = new CT_Boolean();
            m_barser.invertIfNegative.val = 0; 

            m_barser.cat = new CT_AxDataSource();
            m_barser.cat.strRef = new CT_StrRef();
            m_barser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_barser.cat.strRef.strCache = new CT_StrData();
            m_barser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_barser.cat.strRef.strCache.ptCount.val = 4;
            m_barser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_barpt = new CT_StrVal();
            m_barpt.idx = 0;
            m_barpt.v = "类别 1";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 1;
            m_barpt.v = "类别 2";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 2;
            m_barpt.v = "类别 3";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barpt = new CT_StrVal();
            m_barpt.idx = 3;
            m_barpt.v = "类别 4";
            m_barser.cat.strRef.strCache.pt.Add(m_barpt);
            m_barser.val = new CT_NumDataSource();
            m_barser.val.numRef = new CT_NumRef();
            m_barser.val.numRef.f = "Sheet1!$D$2:$D$5";
            m_barser.val.numRef.numCache = new CT_NumData();
            m_barser.val.numRef.numCache.formatCode = "General";
            m_barser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_barser.val.numRef.numCache.ptCount.val = 4;
            m_barser.val.numRef.numCache.pt = new List<CT_NumVal>();
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 0;
            m_barvalpt.v = "2";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 1;
            m_barvalpt.v = "2";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 2;
            m_barvalpt.v = "3";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barvalpt = new CT_NumVal();
            m_barvalpt.idx = 3;
            m_barvalpt.v = "5";
            m_barser.val.numRef.numCache.pt.Add(m_barvalpt);
            m_barchart.ser.Add(m_barser);

            //图表与数据显示设计
            m_barchart.dLbls = new CT_DLbls();
            m_barchart.dLbls.showLegendKey = new CT_Boolean();
            m_barchart.dLbls.showLegendKey.val = 0;
            m_barchart.dLbls.showVal = new CT_Boolean();//值显示设置
            m_barchart.dLbls.showVal.val = 1;
            m_barchart.dLbls.showCatName = new CT_Boolean();//分类显示设置
            m_barchart.dLbls.showCatName.val = 0;
            m_barchart.dLbls.showSerName = new CT_Boolean();//tx标题显示设置
            m_barchart.dLbls.showSerName.val = 0;
            m_barchart.dLbls.showPercent = new CT_Boolean();
            m_barchart.dLbls.showPercent.val = 0;
            m_barchart.dLbls.showBubbleSize = new CT_Boolean();
            m_barchart.dLbls.showBubbleSize.val = 0;
            m_barchart.dLbls.showLeaderLines = new CT_Boolean();
            m_barchart.dLbls.showLeaderLines.val = 0;
            m_barchart.gapWidth = new CT_GapAmount();
            m_barchart.gapWidth.val = 150;
            m_barchart.overlap = new CT_Overlap();
            m_barchart.overlap.val = (sbyte)-25; 

            m_barchart.axId = new List<CT_UnsignedInt>();
            m_axId = new CT_UnsignedInt();
            m_axId.val = 57733120;
            m_barchart.axId.Add(m_axId);
            m_axId = new CT_UnsignedInt();
            m_axId.val = 57734656;
            m_barchart.axId.Add(m_axId);

            m_chart.plotArea.barChart.Add(m_barchart);

            m_chart.plotArea.catAx = new List<CT_CatAx>();
            m_catAx = new CT_CatAx();
            m_catAx.axId = new CT_UnsignedInt();
            m_catAx.axId.val = 57733120;
            m_catAx.scaling = new CT_Scaling();
            m_catAx.scaling.orientation = new CT_Orientation();
            m_catAx.scaling.orientation.val = ST_Orientation.minMax;
            m_catAx.delete = new CT_Boolean();//分类标题
            m_catAx.delete.val = 0;
            m_catAx.axPos = new CT_AxPos();
            m_catAx.axPos.val = ST_AxPos.b;
            m_catAx.majorTickMark = new CT_TickMark();
            m_catAx.majorTickMark.val = ST_TickMark.none;
            m_catAx.minorTickMark = new CT_TickMark();
            m_catAx.minorTickMark.val = ST_TickMark.none;  
            m_catAx.tickLblPos = new CT_TickLblPos();
            m_catAx.tickLblPos.val = ST_TickLblPos.nextTo;
            m_catAx.crossAx = new CT_UnsignedInt();
            m_catAx.crossAx.val = 57733120;
            m_catAx.crosses = new CT_Crosses();
            m_catAx.crosses.val = ST_Crosses.autoZero;
            m_catAx.auto = new CT_Boolean();
            m_catAx.auto.val = 0;
            m_catAx.lblAlgn = new CT_LblAlgn();
            m_catAx.lblAlgn.val = ST_LblAlgn.ctr;
            m_catAx.lblOffset = new CT_LblOffset();
            m_catAx.lblOffset.val = 100;
            m_catAx.noMultiLvlLbl = new CT_Boolean();
            m_catAx.noMultiLvlLbl.val = 100;
            m_chart.plotArea.catAx.Add(m_catAx);

            m_chart.plotArea.valAx = new List<CT_ValAx>();
            m_valAx = new CT_ValAx();
            m_valAx.axId = new CT_UnsignedInt();
            m_valAx.axId.val = 57734656;
            m_valAx.scaling = new CT_Scaling();
            m_valAx.scaling.orientation = new CT_Orientation();
            m_valAx.scaling.orientation.val = ST_Orientation.minMax;
            m_valAx.delete = new CT_Boolean();
            m_valAx.delete.val = 1; 
            m_valAx.axPos = new CT_AxPos();
            m_valAx.axPos.val = ST_AxPos.l;
            m_valAx.numFmt = new NPOI.OpenXmlFormats.Dml.Chart.CT_NumFmt();
            m_valAx.numFmt.formatCode = "General";
            m_valAx.numFmt.sourceLinked = true;
            m_valAx.majorTickMark = new CT_TickMark();
            m_valAx.majorTickMark.val = ST_TickMark.none;
            m_valAx.minorTickMark = new CT_TickMark();
            m_valAx.minorTickMark.val = ST_TickMark.none;  
            m_valAx.tickLblPos = new CT_TickLblPos();
            m_valAx.tickLblPos.val = ST_TickLblPos.nextTo;
            m_valAx.crossAx = new CT_UnsignedInt();
            m_valAx.crossAx.val = 57733120;
            m_valAx.crosses = new CT_Crosses();
            m_valAx.crosses.val = ST_Crosses.autoZero;
            m_valAx.crossBetween = new CT_CrossBetween();
            m_valAx.crossBetween.val = ST_CrossBetween.between;  
            m_chart.plotArea.valAx.Add(m_valAx);

            //图例位置
            m_chart.legend = new CT_Legend();
            m_chart.legend.legendPos = new CT_LegendPos();
            m_chart.legend.legendPos.val = ST_LegendPos.t; //在上方
            m_chart.legend.overlay = new CT_Boolean();
            m_chart.legend.overlay.val = 0;
            m_chart.plotVisOnly = new CT_Boolean();
            m_chart.plotVisOnly.val = 1;
            m_chart.dispBlanksAs = new CT_DispBlanksAs();
            m_chart.dispBlanksAs.val = ST_DispBlanksAs.gap;
            m_chart.showDLblsOverMax = new CT_Boolean();
            m_chart.showDLblsOverMax.val = 0;

            gp = m_Docx.CreateParagraph();
            gr = gp.CreateRun();
            gr.AddChartSpace(new XSSFWorkbook(new MemoryStream(bxlsxData)) , ctbarchartspace, 5274310, 3076575);

            return m_Docx;
        }
        protected XWPFDocument CreateCharttoDocxwithAnchor()
        {
            //anchor
            XWPFDocument m_Docx = new XWPFDocument();
            //页面设置
            //A4:W=11906,h=16838
            //CT_SectPr m_SectPr = m_Docx.Document.body.AddNewSectPr();
            m_Docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_SectPr = m_Docx.Document.body.sectPr;
            //页面设置A4横向
            m_SectPr.pgSz.w = (ulong)16838;
            m_SectPr.pgSz.h = (ulong)11906;

            //插入饼图图表
            //插入xlsx
            //创建xlsx
            XSSFWorkbook workbook = new XSSFWorkbook();
            //创建表单1（饼图）
            ISheet sheet = workbook.CreateSheet("Sheet1");
            //表单1饼图数据
            //         销售额
            //第一季度 8.2
            //第二季度 3.2
            //第三季度 1.4
            //第四季度 1.2

            IRow row = sheet.CreateRow(0);
            ICell cell = row.CreateCell(0);
            cell = row.CreateCell(0);
            cell = row.CreateCell(1);
            cell.SetCellValue("销售额");
            row = sheet.CreateRow(1);
            cell = row.CreateCell(0);
            cell.SetCellValue("第一季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(8.2);
            row = sheet.CreateRow(2);
            cell = row.CreateCell(0);
            cell.SetCellValue("第二季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(3.2);
            row = sheet.CreateRow(3);
            cell = row.CreateCell(0);
            cell.SetCellValue("第三季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(1.4);
            row = sheet.CreateRow(4);
            cell = row.CreateCell(0);
            cell.SetCellValue("第四季度");
            cell = row.CreateCell(1);
            cell.SetCellValue(1.2);

            //把xlsx存入内存流并转为字节流
            MemoryStream msworkbook = new MemoryStream();
            workbook.Write(msworkbook);
            msworkbook.Flush();
            byte[] data = msworkbook.ToArray();
            msworkbook.Close();

            //饼图1
            //创建\word\charts\chartn.xml内容（饼图）

            CT_ChartSpace ctpiechartspace = new CT_ChartSpace();

            ctpiechartspace.date1904 = new CT_Boolean();
            ctpiechartspace.date1904.val = 1;
            ctpiechartspace.lang = new CT_TextLanguageID();
            ctpiechartspace.lang.val = "zh-CN";

            CT_Chart m_chart = ctpiechartspace.AddNewChart();
            //图表标题
            m_chart.title = new CT_Title();//标题采用tx的值
            m_chart.title.overlay = new CT_Boolean();
            m_chart.title.overlay.val = 0;
            m_chart.autoTitleDeleted = new CT_Boolean();
            m_chart.autoTitleDeleted.val = 0; 

            m_chart.plotArea = new CT_PlotArea();
            m_chart.plotArea.pieChart = new List<CT_PieChart>();
            //饼图
            CT_PieChart m_piechart = new CT_PieChart();
            m_piechart.varyColors = new CT_Boolean();
            m_piechart.varyColors.val = 1;
            m_piechart.ser = new List<CT_PieSer>();
            CT_PieSer m_pieser = new CT_PieSer();
            //m_piechart.ser.Add(m_pieser);
            //标题
            m_pieser.tx = new CT_SerTx();
            m_pieser.tx.strRef = new CT_StrRef();
            m_pieser.tx.strRef.f = "Sheet1!$B$1";
            m_pieser.tx.strRef.strCache = new CT_StrData();
            m_pieser.tx.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_pieser.tx.strRef.strCache.ptCount.val = 1;
            CT_StrVal m_strval = new CT_StrVal();
            m_strval.idx = 0;
            m_strval.v = "销售额";
            m_pieser.tx.strRef.strCache.pt = new List<CT_StrVal>();
            m_pieser.tx.strRef.strCache.pt.Add(m_strval);

            
            //m_pieser.dLbls = new CT_DLbls();
            //m_pieser.dLbls.showLegendKey = new CT_Boolean() ;
            //m_pieser.dLbls.showLegendKey.val = 0;
            //m_pieser.dLbls.showVal = new CT_Boolean();
            //m_pieser.dLbls.showVal.val = 0; 
            //m_pieser.dLbls.showCatName = new CT_Boolean();
            //m_pieser.dLbls.showCatName.val = 0; 
            //m_pieser.dLbls.showSerName = new CT_Boolean();
            //m_pieser.dLbls.showSerName.val = 0; 
            //m_pieser.dLbls.showPercent = new CT_Boolean();
            //m_pieser.dLbls.showPercent.val = 1;
            //m_pieser.dLbls.showBubbleSize = new CT_Boolean();
            //m_pieser.dLbls.showBubbleSize.val = 0;
            //m_pieser.dLbls.showLeaderLines = new CT_Boolean();
            //m_pieser.dLbls.showLeaderLines.val = 1;

            //行标题
            m_pieser.cat = new CT_AxDataSource();
            m_pieser.cat.strRef = new CT_StrRef();
            m_pieser.cat.strRef.f = "Sheet1!$A$2:$A$5";
            m_pieser.cat.strRef.strCache = new CT_StrData();
            m_pieser.cat.strRef.strCache.ptCount = new CT_UnsignedInt();
            m_pieser.cat.strRef.strCache.ptCount.val = 4;
            m_pieser.cat.strRef.strCache.pt = new List<CT_StrVal>();
            m_strval = new CT_StrVal();
            m_strval.idx = 0;
            m_strval.v = "第一季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 1;
            m_strval.v = "第二季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 2;
            m_strval.v = "第三季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            m_strval = new CT_StrVal();
            m_strval.idx = 3;
            m_strval.v = "第四季度";
            m_pieser.cat.strRef.strCache.pt.Add(m_strval);
            //值
            m_pieser.val = new CT_NumDataSource();
            m_pieser.val.numRef = new CT_NumRef();
            m_pieser.val.numRef.f = "Sheet1!$B$2:$B$5";
            m_pieser.val.numRef.numCache = new CT_NumData();
            m_pieser.val.numRef.numCache.formatCode = "General";
            m_pieser.val.numRef.numCache.ptCount = new CT_UnsignedInt();
            m_pieser.val.numRef.numCache.ptCount.val = 4;
            m_pieser.val.numRef.numCache.pt = new List<CT_NumVal>();
            CT_NumVal m_numval = new CT_NumVal();
            m_numval.idx = 0;
            m_numval.v = "8.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 1;
            m_numval.v = "3.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 2;
            m_numval.v = "1.4";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_numval = new CT_NumVal();
            m_numval.idx = 3;
            m_numval.v = "1.2";
            m_pieser.val.numRef.numCache.pt.Add(m_numval);
            m_piechart.ser.Add(m_pieser);

            //图表与数据显示设计
            m_piechart.dLbls = new CT_DLbls();
            m_piechart.dLbls.showLegendKey = new CT_Boolean();
            m_piechart.dLbls.showLegendKey.val = 0;
            m_piechart.dLbls.showVal = new CT_Boolean();//值显示设置
            m_piechart.dLbls.showVal.val = 0;
            m_piechart.dLbls.showCatName = new CT_Boolean();//分类显示设置
            m_piechart.dLbls.showCatName.val = 0;
            m_piechart.dLbls.showSerName = new CT_Boolean();//tx标题显示设置
            m_piechart.dLbls.showSerName.val = 0;
            m_piechart.dLbls.showPercent = new CT_Boolean();
            m_piechart.dLbls.showPercent.val = 1;
            m_piechart.dLbls.showBubbleSize = new CT_Boolean();
            m_piechart.dLbls.showBubbleSize.val = 0;
            m_piechart.dLbls.showLeaderLines = new CT_Boolean();
            m_piechart.dLbls.showLeaderLines.val = 1;

            m_piechart.firstSliceAng = new CT_FirstSliceAng();
            m_piechart.firstSliceAng.val = 0; 

            m_chart.plotArea.pieChart.Add(m_piechart);

            //图例
            m_chart.legend = new CT_Legend();
            m_chart.legend.legendPos = new CT_LegendPos();
            m_chart.legend.legendPos.val = ST_LegendPos.t;//图例在上方
            m_chart.plotVisOnly = new CT_Boolean();
            m_chart.plotVisOnly.val = 1;

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

            //图左上角坐标
            CT_PosH posH = new CT_PosH();
            posH.relativeFrom = ST_RelFromH.column;
            posH.posOffset = 2000000;//单位：EMUS,1CM=360000EMUS
            CT_PosV posV = new CT_PosV();
            posV.relativeFrom = ST_RelFromV.paragraph;
            posV.posOffset = 0;

            //四周-bisides
            CT_WrapSquare wrapSquare = new CT_WrapSquare();
            wrapSquare.wrapText = ST_WrapText.bothSides;

            XWPFParagraph gp = m_Docx.CreateParagraph();
            XWPFRun gr = gp.CreateRun();
            gp = m_Docx.CreateParagraph();
            gr = gp.CreateRun();
            gr.AddChartSpace(new XSSFWorkbook(new MemoryStream(data)), ctpiechartspace, 4274310, 2076575, posH, posV, wrapSquare,anchor, simplePos, effectExtent);
 //           gr.AddBreak();//分页

            //饼图2
            m_chart.legend.legendPos.val = ST_LegendPos.r; //图例在右侧 

            //anchor方式插图
            anchor = new CT_Anchor();
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

            simplePos = new CT_Positive2D();
            simplePos.x = 0;
            simplePos.y = 0;

            effectExtent = new CT_EffectExtent();
            effectExtent.l = 0;
            effectExtent.t = 0;
            effectExtent.r = 0;
            effectExtent.b = 0;

            //图左上角坐标
            posH = new CT_PosH();
            posH.relativeFrom = ST_RelFromH.column;
            posH.posOffset = 0;//单位：EMUS,1CM=360000EMUS
            posV = new CT_PosV();
            posV.relativeFrom = ST_RelFromV.paragraph;
            posV.posOffset = 2176575;


            //wrapTight（紧密）
            CT_WrapTight wrapTight = new CT_WrapTight();
            wrapTight.wrapText = ST_WrapText.bothSides;
            wrapTight.wrapPolygon = new CT_WrapPath();
            wrapTight.wrapPolygon.edited = false;
            wrapTight.wrapPolygon.start = new CT_Positive2D();
            wrapTight.wrapPolygon.start.x = 0;
            wrapTight.wrapPolygon.start.y = 0;
            CT_Positive2D lineTo = new CT_Positive2D();
            wrapTight.wrapPolygon.lineTo = new List<CT_Positive2D>();
            lineTo = new CT_Positive2D();
            lineTo.x = 0;
            lineTo.y = 1343;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Positive2D();
            lineTo.x = 21405;
            lineTo.y = 1343;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Positive2D();
            lineTo.x = 21405;
            lineTo.y = 0;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo.x = 0;
            lineTo.y = 0;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);

            gp = m_Docx.CreateParagraph();
            gr = gp.CreateRun();
            //修改标题
            workbook = new XSSFWorkbook(new MemoryStream(data));
            workbook.GetSheet("Sheet1").GetRow(0).GetCell(1).SetCellValue("销售金额");

            ctpiechartspace.chart.plotArea.pieChart[0].ser[0].tx.strRef.strCache.pt[0].v = "销售金额";

            gr.AddChartSpace(workbook, ctpiechartspace, 3274310, 2076575,posH,posV, wrapTight,anchor,simplePos,effectExtent);

            //gr.AddBreak();//分页

            //饼图3
            m_chart.legend.legendPos.val = ST_LegendPos.b; //图例在下方 

            //anchor方式插图
            anchor = new CT_Anchor();
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

            simplePos = new CT_Positive2D();
            simplePos.x = 0;
            simplePos.y = 0;

            effectExtent = new CT_EffectExtent();
            effectExtent.l = 0;
            effectExtent.t = 0;
            effectExtent.r = 0;
            effectExtent.b = 0;

            //图左上角坐标
            posH = new CT_PosH();
            posH.relativeFrom = ST_RelFromH.column;
            posH.posOffset = 3674310;//单位：EMUS,1CM=360000EMUS
            posV = new CT_PosV();
            posV.relativeFrom = ST_RelFromV.paragraph;
            posV.posOffset = 2176575;


            //wrapThrough(穿越)
            CT_WrapThrough wrapThrough = new CT_WrapThrough();
            wrapThrough.wrapText = ST_WrapText.bothSides;
            wrapThrough.wrapPolygon = new CT_WrapPath();
            wrapThrough.wrapPolygon.edited = false;
            wrapThrough.wrapPolygon.start = new CT_Positive2D();
            wrapThrough.wrapPolygon.start.x = 0;
            wrapThrough.wrapPolygon.start.y = 0;
            lineTo = new CT_Positive2D();
            wrapThrough.wrapPolygon.lineTo = new List<CT_Positive2D>();
            lineTo = new CT_Positive2D();
            lineTo.x = 0;
            lineTo.y = 1343;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Positive2D();
            lineTo.x = 21405;
            lineTo.y = 1343;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Positive2D();
            lineTo.x = 21405;
            lineTo.y = 0;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo.x = 0;
            lineTo.y = 0;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);

            //gp = m_Docx.CreateParagraph();
            //gr = gp.CreateRun();
            //修改标题
            workbook = new XSSFWorkbook(new MemoryStream(data));
            workbook.GetSheet("Sheet1").GetRow(0).GetCell(1).SetCellValue("销售金额数");
            ctpiechartspace.chart.plotArea.pieChart[0].ser[0].tx.strRef.strCache.pt[0].v = "销售金额数";

            gr.AddChartSpace(workbook, ctpiechartspace, 3274310, 2076575,posH,posV, wrapThrough,anchor,simplePos,effectExtent);



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


    }
}
