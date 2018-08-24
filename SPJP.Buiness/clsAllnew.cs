using Microsoft.Reporting.WinForms;
using SPJP.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Net;
namespace SPJP.Buiness
{
    public class clsAllnew
    {

        string orderprint;
        string tisprint;
        #region print
        private List<Stream> m_streams;
        private int m_currentPageIndex;
        private BackgroundWorker bgWorker1;
        //private object missing = System.Reflection.Missing.Value;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        public log4net.ILog ProcessLogger { get; set; }
        public log4net.ILog ExceptionLogger { get; set; }

        #endregion

        public void Run(List<clsOrderDatabaseinfo> FilterOrderResults)
        {
            LocalReport report = new LocalReport();
            report.ReportPath = Application.StartupPath + "\\Report1.rdlc";
            report.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", FilterOrderResults));
            Export(report);
            m_currentPageIndex = 0;
            Print(orderprint, 0, 0);
        }
        public void Export(LocalReport report)
        {
            //A4    21*29.7厘米（210mm×297mm）
            // A5的尺寸:148毫米*210毫米
            //string deviceInfo =
            //  "<DeviceInfo>" +
            //  "  <OutputFormat>EMF</OutputFormat>" +
            //  "  <PageWidth>8.5in</PageWidth>" +
            //  "  <PageHeight>11in</PageHeight>" +
            //  "  <MarginTop>0.25in</MarginTop>" +
            //  "  <MarginLeft>0.25in</MarginLeft>" +
            //  "  <MarginRight>0.25in</MarginRight>" +
            //  "  <MarginBottom>0.25in</MarginBottom>" +
            //  "</DeviceInfo>";       

            //    string deviceInfo =
            //"<DeviceInfo>" +
            //"  <OutputFormat>EMF</OutputFormat>" +
            //"  <PageWidth>26cm</PageWidth>" +
            //"  <PageHeight>15.2cm</PageHeight>" +
            //"  <MarginTop>0.1cm</MarginTop>" +
            //"  <MarginLeft>0.1cm</MarginLeft>" +
            //"  <MarginRight>0.1cm</MarginRight>" +
            //"  <MarginBottom>0.1cm</MarginBottom>" +
            //"</DeviceInfo>";
            //           string deviceInfo =
            //"<DeviceInfo>" +
            //"  <OutputFormat>EMF</OutputFormat>" +
            //"  <PageWidth>9.44in</PageWidth>" +
            //"  <PageHeight>5.98in</PageHeight>" +
            //"  <MarginTop>0.0cm</MarginTop>" +
            //"  <MarginLeft>0.0cm</MarginLeft>" +
            //"  <MarginRight>0.0cm</MarginRight>" +
            //"  <MarginBottom>0.0cm</MarginBottom>" +
            //"</DeviceInfo>";

            //            string deviceInfo =
            //"<DeviceInfo>" +
            //"  <OutputFormat>EMF</OutputFormat>" +
            //"  <PageWidth>8.66in</PageWidth>" +
            //"  <PageHeight>5.2</PageHeight>" +
            //"  <MarginTop>0.0cm</MarginTop>" +
            //"  <MarginLeft>0.0cm</MarginLeft>" +
            //"  <MarginRight>0.0cm</MarginRight>" +
            //"  <MarginBottom>0.0cm</MarginBottom>" +
            //"</DeviceInfo>";

            string deviceInfo =
                            "<DeviceInfo>" +
                            "  <OutputFormat>EMF</OutputFormat>" +
                            "  <PageWidth>8.5in</PageWidth>" +
                            "  <PageHeight>11.1in</PageHeight>" +
                            "  <MarginTop>0in</MarginTop>" +
                            "  <MarginLeft>0.0cm</MarginLeft>" +
                            "  <MarginRight>0.0cm</MarginRight>" +
                            "  <MarginBottom>0in</MarginBottom>" +
                            "</DeviceInfo>";


            Warning[] warnings;
            m_streams = new List<Stream>();
            report.Render("Image", deviceInfo, CreateStream,
               out warnings);
            foreach (Stream stream in m_streams)
                stream.Position = 0;
        }
        private Stream CreateStream(string name, string fileNameExtension,

Encoding encoding, string mimeType, bool willSeek)
        {

            //如果需要将报表输出的数据保存为文件，请使用FileStream对象。
            Stream stream = new MemoryStream();

            m_streams.Add(stream);

            return stream;

        }
        public void Print(string defaultPrinterName, int lenpage, int withpage)
        {

            m_currentPageIndex = 0;
            if (m_streams == null || m_streams.Count == 0)
                return;
            //声明PrintDocument对象用于数据的打印

            PrintDocument printDoc = new PrintDocument();

            //指定需要使用的打印机的名称，使用空字符串""来指定默认打印机

            if (defaultPrinterName == "" || defaultPrinterName == null)
                defaultPrinterName = printDoc.PrinterSettings.PrinterName;

            printDoc.PrinterSettings.PrinterName = defaultPrinterName;

            //判断指定的打印机是否可用

            if (!printDoc.PrinterSettings.IsValid)
            {
                MessageBox.Show("Can't find printer");
                return;
            }
            //声明PrintDocument对象的PrintPage事件，具体的打印操作需要在这个事件中处理。

            printDoc.PrintPage += new PrintPageEventHandler(PrintPage);

            //执行打印操作，Print方法将触发PrintPage事件。
            printDoc.DefaultPageSettings.Landscape = false;
            //大小
            if (lenpage != 0)
                printDoc.DefaultPageSettings.PaperSize = new PaperSize("Custom", lenpage, withpage);


            printDoc.Print();

        }
        private void PrintPage(object sender, PrintPageEventArgs ev)
        {
            Metafile pageImage = new
               Metafile(m_streams[m_currentPageIndex]);
            //ev.PageSettings.Landscape = true;
            //
            StringFormat SF = new StringFormat();
            SF.LineAlignment = StringAlignment.Center;
            SF.Alignment = StringAlignment.Center;
            //RectangleF rect = new RectangleF(0, 0, ev.PageBounds.Width, ev.Graphics.MeasureString("Authors Informations", new Font("Times New Roman", 20)).Height);    //其中e.PageBounds属性表示页面全部区域的矩形区域
            //ev.Graphics.MeasureString(string,Font).Heighte.Graphics.DrawString("Authors Informations",new Font("Times New Roman",20),Brushes.Black,rect,SF);
            float left = ev.PageSettings.Margins.Left;//打印区域的左边界
            float top = ev.PageSettings.Margins.Top;//打印区域的上边界
            float width = ev.PageSettings.PaperSize.Width - left - ev.PageSettings.Margins.Right;//计算出有效打印区域的宽度
            float height = ev.PageSettings.PaperSize.Height - top - ev.PageSettings.Margins.Bottom;//计算出有效打印区域的高度
            ev.Graphics.DrawImage(pageImage, ev.PageBounds);
            m_currentPageIndex++;
            ev.HasMorePages = (m_currentPageIndex < m_streams.Count);
        }


        #region add data
        private void InitializeReportEvent()
        {
            LocalReport report = new LocalReport();
            report.ReportPath = Application.StartupPath + "\\Report1.rdlc";

            //report.LocalReport.SubreportProcessing += new SubreportProcessingEventHandler(LocalReport_SubreportProcessing);
            //  this.report.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout);
        }

        #endregion


        public void Run2(List<clsOrderDatabaseinfo> FilterOrderResults)
        {
            LocalReport report = new LocalReport();
            report.ReportPath = Application.StartupPath + "\\Report2.rdlc";
            report.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", FilterOrderResults));
            Export(report);
            m_currentPageIndex = 0;
            Print(orderprint, 0, 0);
        }
        public void Run3(List<clsOrderDatabaseinfo> FilterOrderResults)
        {
            LocalReport report = new LocalReport();
            report.ReportPath = Application.StartupPath + "\\Report3.rdlc";
            report.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", FilterOrderResults));
            Export(report);
            m_currentPageIndex = 0;
            Print(orderprint, 0, 0);
        }



        public List<clsOrderDatabaseinfo> ReadJSON_Report(ref BackgroundWorker bgWorker, string casetype)
        {

            List<clsOrderDatabaseinfo> ResutsReport = new List<clsOrderDatabaseinfo>();

            ReadJsonID();



            return ResutsReport;

        }
        #region 假数据
        protected void Page_Load(object sender, EventArgs e)
        {
            string inputJsonString = @"
                [
                    {StudentID:'100',Name:'aaa',Hometown:'china'},
                    {StudentID:'101',Name:'bbb',Hometown:'us'},
                    {StudentID:'102',Name:'ccc',Hometown:'england'}
                ]";
            JArray jsonObj = JArray.Parse(inputJsonString);
            string message = @"<table border='1'>
                    <tr><td width='80'>StudentID</td><td width='100'>Name</td><td width='100'>Hometown</td></tr>";
            string tpl = "<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>";
            foreach (JObject jObject in jsonObj)
            {
                message += String.Format(tpl, jObject["StudentID"], jObject["Name"], jObject["Hometown"]);
            }
            message += "</table>";
            string tx = message;
        }

        private string readtxtJsom()
        {

            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "json\\pdf.json";

            string[] fileText = File.ReadAllLines(A_Path);

            string mailbody = "";

            for (int i = 0; i < fileText.Length; i++)
            {
                string a = fileText[i].Trim();
                if (fileText[i].Contains("Regards"))
                {

                }
                else if (i != 0)
                    mailbody = mailbody + "\r\n" + a;
                else if (i == 0)
                    mailbody = a;
            }

            return mailbody;


        }

        #endregion
        private void ReadJsonID()
        {
            string result1 = DoPost("https://api.douban.com/v2/book/isbn/9787115212948", "10");

            //假数据
            result1 = readtxtJsom();

            //取大字符串,这个“100”是我随便填的！！！
            JObject obj = (JObject)JsonConvert.DeserializeObject(result1);
            //再讲字符串转成json格式
            string tx1 = obj["typeCode"].ToString();
            //获取summary内容


            Newtonsoft.Json.Linq.JArray userAarray1 = Newtonsoft.Json.Linq.JArray.Parse(result1) as Newtonsoft.Json.Linq.JArray;
            List<clsOrderDatabaseinfo> userListModel = userAarray1.ToObject<List<clsOrderDatabaseinfo>>();


            //new 
            #region MyRegion
//            JArray array = (JArray)obj["article"];
//            if (array != null)

//                foreach (var jObject in array)
//                {
//                    //赋值属性
//                }
//            string inputJsonString = @"
//                [
//                    {StudentID:'100',Name:'aaa',Hometown:'china'},
//                    {StudentID:'101',Name:'bbb',Hometown:'us'},
//                    {StudentID:'102',Name:'ccc',Hometown:'england'}
//                ]";

//            string message = @"<table border='1'>
//                    <tr><td width='80'>StudentID</td><td width='100'>Name</td><td width='100'>Hometown</td></tr>";
//            string tpl = "<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>";
//            List<clsOrderDatabaseinfo> studentList = JsonConvert.DeserializeObject<List<clsOrderDatabaseinfo>>(inputJsonString);//注意这里必须为List<Student>类型,因为客户端提交的是一个数组json
//            foreach (clsOrderDatabaseinfo student in studentList)
//            {
//                //message += String.Format(tpl, student.StudentID, student.Name, student.Hometown);
//            } 
            #endregion

        }
        public string DoPost(string url, string data)
        {
            HttpWebRequest req = GetWebRequest(url, "POST");
            byte[] postData = Encoding.UTF8.GetBytes(data);
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(postData, 0, postData.Length);
            reqStream.Close();
            HttpWebResponse rsp = (HttpWebResponse)req.GetResponse();
            Encoding encoding = Encoding.GetEncoding(rsp.CharacterSet);
            return GetResponseAsString(rsp, encoding);
        }
        public HttpWebRequest GetWebRequest(string url, string method)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.ServicePoint.Expect100Continue = false;
            req.ContentType = "application/x-www-form-urlencoded;charset=utf-8";
            req.ContentType = "text/json";
            req.Method = method;
            req.KeepAlive = true;
            req.UserAgent = "guanyisoft";
            req.Timeout = 1000000;
            req.Proxy = null;
            return req;
        }
        public string GetResponseAsString(HttpWebResponse rsp, Encoding encoding)
        {
            StringBuilder result = new StringBuilder();
            Stream stream = null;
            StreamReader reader = null;
            try
            {
                // 以字符流的方式读取HTTP响应
                stream = rsp.GetResponseStream();
                reader = new StreamReader(stream, encoding);
                // 每次读取不大于256个字符，并写入字符串
                char[] buffer = new char[256];
                int readBytes = 0;
                while ((readBytes = reader.Read(buffer, 0, buffer.Length)) > 0)
                {
                    result.Append(buffer, 0, readBytes);
                }
            }
            finally
            {
                // 释放资源
                if (reader != null) reader.Close();
                if (stream != null) stream.Close();
                if (rsp != null) rsp.Close();
            }
            return result.ToString();
        }
    }
}
