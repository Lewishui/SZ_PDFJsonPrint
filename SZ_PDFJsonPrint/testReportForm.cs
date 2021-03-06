﻿using Microsoft.Reporting.WinForms;
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
    public partial class testReportForm : Form
    {
        DataTable FilterOrderResults;
        List<OnlineShow> OnlineShow_datas;

        public testReportForm(DataTable orders, List<OnlineShow> OnlineShow_datas1)
        {
            InitializeComponent();
            FilterOrderResults = new DataTable();
            FilterOrderResults = orders;
            OnlineShow_datas = OnlineShow_datas1;
            InitializeDataSource();
          InitializeReportEvent_pl();
       // InitializeReportEvent();
        }
        public void InitializeDataSource( )
        {
           

            this.reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.EnableExternalImages = true;
            //this.reportViewer1.LocalReport.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", orders));
            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet3", FilterOrderResults));
            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", OnlineShow_datas));
            //  this.reportViewer1.Reset();

        }
        private void testReportForm_Load(object sender, EventArgs e)
        {
            //  reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
            //NewMethod(); 
            this.reportViewer1.RefreshReport();
        }
        private void InitializeReportEvent()
        {
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

  
     
        void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {
      



            var s = e.ReportPath;
            e.DataSources.Clear();
            //e.DataSources.Add(new ReportDataSource("DataSet1", new List<v_pendingorder>() { orderFirst }));
            e.DataSources.Add(new ReportDataSource("DataSet3", FilterOrderResults));
            e.DataSources.Add(new ReportDataSource("DataSet1", OnlineShow_datas));

        
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
            StreamReader mainstream = new StreamReader(Application.StartupPath + "\\Report1.rdlc");
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
            _reportNameList.Add("Report2");
            foreach (string reportName in _reportNameList)
            {
                StreamReader subStream = new StreamReader( reportName + ".rdlc");
                reportViewer1.LocalReport.LoadSubreportDefinition(reportName, subStream);
                subStream.Close();
            }
            //设置主报表数据源和所有报表（主，子）报表需要的参数等逻辑
            // ReportViewer1.LocalReport.DataSources.Add(数据源);
            //设置子报表进行事件订阅            
            reportViewer1.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);

        }
        #endregion
    }
}
