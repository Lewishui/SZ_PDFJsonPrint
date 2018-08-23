using Microsoft.Reporting.WinForms;
using SPJP.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SZ_PDFJsonPrint
{
    public partial class ReportForm : Form
    {
        List<clsOrderDatabaseinfo> FilterOrderResults;
        public ReportForm()
        {
            InitializeComponent();

            InitializeReportEvent();
        }

        private void ReportForm_Load(object sender, EventArgs e)
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

        private void NewMethod()
        {
            ReportDataSource rds = new ReportDataSource("DataSet1", FilterOrderResults);
            this.reportViewer1.LocalReport.SubreportProcessing += LocalReport_SubreportProcessing;
            this.reportViewer1.LocalReport.DataSources.Clear();
            this.reportViewer1.LocalReport.DataSources.Add(rds);
            this.reportViewer1.LocalReport.Refresh();
            this.reportViewer1.RefreshReport();
        }
        public void InitializeDataSource(List<clsOrderDatabaseinfo> orders)
        {
            FilterOrderResults = new List<clsOrderDatabaseinfo>();
            FilterOrderResults = orders;

            this.reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.EnableExternalImages = true;
            //this.reportViewer1.LocalReport.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", orders));
            this.reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", orders));

        }

        void LocalReport_SubreportProcessing(object sender, SubreportProcessingEventArgs e)
        {
            #region step 1
            //ReportDataSource rds = new ReportDataSource("DataSet1", FilterOrderResults);

            //e.DataSources.Clear();
            //e.DataSources.Add(rds);
            //return; 
            #endregion



            var s = e.ReportPath;
            e.DataSources.Clear();
            //e.DataSources.Add(new ReportDataSource("DataSet1", new List<v_pendingorder>() { orderFirst }));
            e.DataSources.Add(new ReportDataSource("DataSet1", FilterOrderResults));
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
    }
}
