using Microsoft.Reporting.WinForms;
using SPJP.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SZ_PDFJsonPrint
{
    public partial class ReportForm : Form
    {
        DataTable FilterOrderResults;
        List<OnlineShow> OnlineShow_datas;
        List<PDF_checkdataDetail> PDFcheckdataDetail;
        List<Types> PDF_Types;
        int ongong_index = 0;
        double allpage_count = 0;
        bool have_yushu = false;

        public ReportForm()
        {
            InitializeComponent();

            //InitializeReportEvent_pl();
            InitializeReportEvent();

            //1
        }

        private void ReportForm_Load(object sender, EventArgs e)
        {
            //4

            ongong_index = 0;
            //  reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
            //NewMethod(); 
            this.reportViewer1.RefreshReport();
        }
        public void autoRefreshReport()
        {

            this.reportViewer1.RefreshReport();  
        }
        private void InitializeReportEvent()
        {
            //2
            //  this.reportViewer1.Reset();

            this.reportViewer1.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);
            this.reportViewer1.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);


            this.reportViewer1.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.Percent;
            this.reportViewer1.ZoomPercent = 100;
            PageSettings pageset = new PageSettings();
            pageset.Landscape = false;
            //var pageSettings = this.reportViewer1.GetPageSettings();
            pageset.PaperSize = new PaperSize()
            {
                //Width = 210,
                //Height = 297
                Width = 827,
                Height = 1169
            };
            pageset.Margins = new Margins() { Left = 10, Top = 10, Bottom = 10, Right = 10 };
            reportViewer1.SetPageSettings(pageset);
        }

        private void NewMethod()
        {
            ReportDataSource rds = new ReportDataSource("DataSet1", FilterOrderResults);
            this.reportViewer1.LocalReport.SubreportProcessing += LocalReport_SubreportProcessing;
            this.reportViewer1.LocalReport.DataSources.Clear();
            this.reportViewer1.LocalReport.DataSources.Add(rds);
            this.reportViewer1.LocalReport.Refresh();
            this.reportViewer1.RefreshReport();
        }
        public void InitializeDataSource(DataTable orders, List<OnlineShow> OnlineShow_datas1, List<PDF_checkdataDetail> PDFcheckdataDetail1, List<Types> PDF_Types1)
        {
         //   ongong_index = 0;

            //3
            have_yushu = false;
            PDFcheckdataDetail = new List<PDF_checkdataDetail>();
            PDF_Types = new List<Types>();
            FilterOrderResults = new DataTable();
            PDFcheckdataDetail = PDFcheckdataDetail1;
            FilterOrderResults = orders;
            OnlineShow_datas = OnlineShow_datas1;
            PDF_Types = PDF_Types1;

            //下方 显示总页数
            if (PDF_Types != null)
            {
                int add1 = 0;
                if (PDF_Types.Count % 2 != 0)
                    add1 = add1 + 1;

                int allc = PDFcheckdataDetail.Count / 2 + add1;

                OnlineShow_datas[0].all_count = allc.ToString();

                #region 确认要用多少个report
                int i = 0;

                foreach (Types item in PDF_Types)
                {
                    i++;
                    if (i == 2)
                    {
                        allpage_count++;
                        i = 0;
                    }
                    if (PDF_Types.Count % 2 == 0)
                    {

                    }
                    else
                        have_yushu = true;
                }
            }

                #endregion
            this.reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.EnableExternalImages = true;
            //this.reportViewer1.LocalReport.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", orders));
            //this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet3", orders));


            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", OnlineShow_datas));

            //this.reportViewer1.LocalReport.SetParameters(new ReportParameter("ReportParameter1", PDFcheckdataDetail.Count.ToString()));


        }

        void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {

            //5

            #region step 1
            //ReportDataSource rds = new ReportDataSource("DataSet1", FilterOrderResults);

            //e.DataSources.Clear();
            //e.DataSources.Add(rds);
            //return; 
            #endregion

            //3

            var s = e.ReportPath;
            e.DataSources.Clear();

            int not = 0;
            #region   //check  add  index


            if (PDF_Types != null)
            {
                int report_voulmn = PDF_Types.Count / 2;
                report_voulmn = report_voulmn + 1;

                if (s == "Report2" && allpage_count > 0)//
                {
                    string typecode = PDF_Types[0].typeCode;
                    string typecode2 = PDF_Types[1].typeCode;

                    List<PDF_checkdataDetail> findsapinfo = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode);
                    e.DataSources.Add(new ReportDataSource("DataSet_up", findsapinfo));
                    // OnlineShow_datas[0].showimage = true;


                    List<PDF_checkdataDetail> findsapinfo2 = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode2);
                    e.DataSources.Add(new ReportDataSource("DataSet_down", findsapinfo2));

                    ongong_index++;



                }
                else if (s == "Report3" && have_yushu == true)//如果有余数则 新建一个单页
                {
                    string typecode = PDF_Types[PDF_Types.Count - 1].typeCode;
                    List<PDF_checkdataDetail> findsapinfo = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode);
                    e.DataSources.Add(new ReportDataSource("DataSet_up", findsapinfo));
                    OnlineShow_datas[0].showimage = false;
                    //e.DataSources.Add(new ReportParameter("DeptNo", "02"));
                    //e.DataSources.Add(new ReportParameter("Parameter1", "This　is　a　parameter"));
                    //e.DataSources.Add(new ReportDataSource("showimage", true));
                    ongong_index++;


                }
                else if (s.Contains("Report") && !s.Contains("_") && Convert.ToDouble(s.Substring(6, s.Length - 6)) >= 5 && allpage_count > 1 && ongong_index < allpage_count)//
                {
                    string typecode = "";
                    string typecode2 = "";

                    int newindex = 0;
                    if (have_yushu == true)
                    {
                        newindex = ongong_index;
                        typecode = PDF_Types[newindex].typeCode;
                        typecode2 = PDF_Types[newindex + 1].typeCode;
                    }
                    else
                    {
                        newindex = ongong_index + 1;
                        if (newindex < PDF_Types.Count)
                        typecode = PDF_Types[newindex].typeCode;
                        if (newindex < PDF_Types.Count-1)
                            typecode2 = PDF_Types[newindex + 1].typeCode;
                   
                    }

                    List<PDF_checkdataDetail> findsapinfo = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode);
                    e.DataSources.Add(new ReportDataSource("DataSet_up", findsapinfo));
                    // OnlineShow_datas[0].showimage = true;

                    List<PDF_checkdataDetail> findsapinfo2 = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode2);
                    e.DataSources.Add(new ReportDataSource("DataSet_down", findsapinfo2));

                    ongong_index++;
                    string rname = s.Substring(6, s.Length - 6);

                    if (rname == "5")
                        OnlineShow_datas[0].showimage5 = true;


                }
                else if (ongong_index >= allpage_count)
                {


                    //OnlineShow_datas[0].showimage = true;
                    e.DataSources.Add(new ReportDataSource("DataSet_up", PDFcheckdataDetail));
                    e.DataSources.Add(new ReportDataSource("DataSet_down", PDFcheckdataDetail));

                }
                e.DataSources.Add(new ReportDataSource("DataSet1", OnlineShow_datas));

            }
            #endregion

            //e.DataSources.Add(new ReportDataSource("DataSet1", new List<v_pendingorder>() { orderFirst }));
            //  e.DataSources.Add(new ReportDataSource("DataSet3", FilterOrderResults));

            //e.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource(reportName, ds1.Tables[0]));
            //throw new NotImplementedException();
        }

        private void rz(object sender, EventArgs e)
        {

            //ReportPageSettings rt = reportViewer1.LocalReport.GetDefaultPageSettings();
            //if (reportViewer1.ParentForm.Width > rt.PaperSize.Width)
            //{
            //    int kk = (reportViewer1.ParentForm.Width - rt.PaperSize.Width) / 2;
            //    reportViewer1.Padding = new Padding(kk, 1, kk, 1);


            //}



        }


        #region A4

        private void Print(string printerName)
        {
            PrintDocument printDoc = new PrintDocument();
            if (printerName.Length > 0)
            {
                printDoc.PrinterSettings.PrinterName = printerName;
            }
            foreach (PaperSize ps in printDoc.PrinterSettings.PaperSizes)
            {
                if (ps.PaperName == "A4")
                {
                    printDoc.PrinterSettings.DefaultPageSettings.PaperSize = ps;
                    printDoc.DefaultPageSettings.PaperSize = ps;
                    // printDoc.PrinterSettings.IsDefaultPrinter;//知道是否是预设定的打印机
                }
            }
            if (!printDoc.PrinterSettings.IsValid)
            {
                string msg = String.Format("Can't find printer " + printerName);
                System.Diagnostics.Debug.WriteLine(msg);
                return;
            }
            printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
            printDoc.Print();
        }
        private void PrintPage(object sender, PrintPageEventArgs ev)
        {
            int ass = 0;

            {
                // var m_streams = ItemEnities.GroupBy(x => x.ジャンル).Select(y => y.First());

                Metafile pageImage = new Metafile("234");
                ev.Graphics.DrawImage(pageImage, 0, 0, 827, 1169);//設置打印尺寸 单位是像素
                ass++;
                ev.HasMorePages = (ass < 6);
            }
        }

        #endregion


        #region new

        private void InitializeReportEvent_pl()
        {
            reportViewer1.Reset();
            StreamReader mainstream = new StreamReader(Application.StartupPath + "\\Report_Home.rdlc");
            reportViewer1.LocalReport.LoadReportDefinition(mainstream);
            mainstream.Close();
            //if (reportViewer1.ShowReportBody == false)
            //{
            //    reportViewer1.ShowReportBody = true;
            //}

            List<string> _reportNameList = new List<string>();
            _reportNameList.Add("Report4");
            _reportNameList.Add("Report1");//这个名字为在插入子报表时候，需要输入的报表名称。这个名称可以有具体的文件也可以没有。不要用类似Subreport2名称去加载子报表，否则会出错
            _reportNameList.Add("Report2");
            _reportNameList.Add("Report3");
            _reportNameList.Add("Report5");
            foreach (string reportName in _reportNameList)
            {
                StreamReader subStream = new StreamReader(reportName + ".rdlc");
                reportViewer1.LocalReport.LoadSubreportDefinition(reportName, subStream);
                subStream.Close();
            }
            //设置主报表数据源和所有报表（主，子）报表需要的参数等逻辑
            // ReportViewer1.LocalReport.DataSources.Add(数据源);
            //设置子报表进行事件订阅            
            // reportViewer1.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);

        }
        #endregion
    }
}
