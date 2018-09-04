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
using System.Web.Script.Serialization;
using System.Data;
using SPJP.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
namespace SPJP.Buiness
{
    public class clsAllnew
    {

        string orderprint;
        string tisprint;
        //excel
        public List<clsExcelinfo> TBB;
        public List<Datas> tclass_datas;
        public List<Root> Root_datas;

        //客户查找
        public List<Online_Data> Online_datas;
        public List<MaGait> Online_MaGait;
        public List<Online_Root> Online_Root_datas;
        //PDF
        public List<PDF_Root> PDF_Rootdb;
        public List<Types> PDF_Types;
        public List<ChildType> PDF_ChildType;
        public DataTable Excel_body;

        private string interface_pdfinfo;
        private string interface_excelinfo;
        private string interface_onlineinfo;
        string readtypeA;


        string savepdfexcel_path;

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

        public void Run(List<Datas> FilterOrderResults)
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

            //string deviceInfo =
            //                "<DeviceInfo>" +
            //                "  <OutputFormat>EMF</OutputFormat>" +
            //                "  <PageWidth>8.5in</PageWidth>" +
            //                "  <PageHeight>11.1in</PageHeight>" +
            //                "  <MarginTop>0in</MarginTop>" +
            //                "  <MarginLeft>0.0cm</MarginLeft>" +
            //                "  <MarginRight>0.0cm</MarginRight>" +
            //                "  <MarginBottom>0in</MarginBottom>" +
            //                "</DeviceInfo>";

            string deviceInfo =
                      "<DeviceInfo>" +
                      "  <OutputFormat>EMF</OutputFormat>" +
                      "  <PageWidth>21cm</PageWidth>" +
                      "  <PageHeight>29.7cm</PageHeight>" +
                      "  <MarginTop>0cm</MarginTop>" +
                      "  <MarginLeft>1cm</MarginLeft>" +
                      "  <MarginRight>1cm</MarginRight>" +
                      "  <MarginBottom>0cm</MarginBottom>" +
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


        public void Run2(List<Datas> FilterOrderResults)
        {
            LocalReport report = new LocalReport();
            report.ReportPath = Application.StartupPath + "\\Report2.rdlc";
            report.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", FilterOrderResults));
            Export(report);
            m_currentPageIndex = 0;
            Print(orderprint, 0, 0);
        }
        public void Run3(List<Datas> FilterOrderResults)
        {
            LocalReport report = new LocalReport();
            report.ReportPath = Application.StartupPath + "\\Report3.rdlc";
            report.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", FilterOrderResults));
            Export(report);
            m_currentPageIndex = 0;
            Print(orderprint, 0, 0);
        }



        public List<clsOrderDatabaseinfo> ReadJSON_Report(ref BackgroundWorker bgWorker, string readtype, string pdf_json, string excel_json, string Online_json)
        {
            interface_pdfinfo = pdf_json;
            interface_excelinfo = excel_json;

            interface_onlineinfo = Online_json;
            readtypeA = readtype;

            List<clsOrderDatabaseinfo> ResutsReport = new List<clsOrderDatabaseinfo>();
            TBB = new List<clsExcelinfo>();
            tclass_datas = new List<Datas>();
            Online_datas = new List<Online_Data>();
            Online_MaGait = new List<MaGait>();
            PDF_Rootdb = new List<PDF_Root>();
            PDF_Types = new List<Types>();
            PDF_ChildType = new List<ChildType>();
            Root_datas = new List<Root>();
            Online_Root_datas = new List<Online_Root>();
            Excel_body = new DataTable();

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

        private string readtxtJsom(string A_Path)
        {


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


            #region MyRegion

            //            mailbody = @"{
            //                            'code': 0,
            //                             'msg': 'success',
            //	            'reqId':'40288a5f6518566601651861fd000004',
            //                'data': {
            //		            'header': {
            //			            'rows': [{
            //					            'cols': [{
            //							            'text': '基础信息',
            //							            'size': 20		//size表示跨列数目
            //						            },
            //						            {
            //							            'text': '站立过程SiSt',
            //							            'size': 3
            //						            },
            //						            {
            //							            'text': '行走Gait',
            //							            'size': 35
            //						            },
            //						            {
            //							            'text': '平衡Balance',
            //							            'size': 2
            //						            }
            //					            ]
            //				            },
            //				            {
            //					            'cols': [{
            //							            'text': '序号',
            //							            'size': 1
            //						            },
            //						            {
            //							            'text': '患者',
            //							            'size': 1
            //						            },
            //						            {
            //							            'text': '起立过程时间SiSt Duration（sec）',
            //							            'size': 1
            //						            },
            //						            {
            //							            'text': '右脚步长Step Length R.（cm）',
            //							            'size': 1
            //						            },
            //						            {
            //							            'text': '左脚步长Step Length L.（cm）',
            //							            'size': 1
            //						            },
            //						            {
            //							            'text': '躯干矢状面峰值角速度Trunk Sagittal PeakVelocity（degree/sec）',
            //							            'size': 1
            //						            }
            //					            ]
            //				            }
            //			            ]
            //		            },
            //		            'content': {
            //			            'datas': [
            //                                    {
            //                              'indexNumber': 1, //序号
            //                              'patientId': 7957, //患者ID
            //                              'patientName': '王二',//患者姓名
            //                              'sex': 1,// 患者性别（1男,2女）
            //                              'birthday': '2000-01-01 00:00:00',//出生年月
            //                              'diseaseType': '1',//诊断（1帕金森综合征,2.特发性震颤…）
            //                              'courseOfDisease': 2,// 病程（年）
            //                              'serialNumber': '201808062022447132156',//检查订单流水号
            //                              'inspectId': '25',//检查ID
            //                              'checkType': '1',//检查内容（1计时起身-步行,2窄道）
            //                              'dualTask': '1',//双任务（0否,1是）
            //                              'checkRemark': '检查项备注内容',//检查备注
            //                              'checkStartTime': '2018-08-01 17:19:00',//检查开始时间
            //                              'checkEndTime': '2018-08-06 17:19:00',//检查结束时间
            //                              'checkHospitalId': '8a9e2d385f1eb70c015f2e41a15000c5',//检查医院id
            //                              'checkHospitalName': '华侨医院',//检查医院
            //                              'checkDoctorId': 'biwei',//检查医生id
            //                              'checkDoctorName': '毕伟',//检查医生
            //                              'auditDoctorName': null,// 审核医生
            //                              'checkConclusion': '这是医生的结论',//检查结论
            //                              'sistDuration': '23±2',//起立过程时间SiSt Duration（sec）
            //                              'peakSiStAngularVelocity': '22±1',
            //                        //起立过程中躯干矢状面峰值角速度Peak SiSt Angular Velocity(degree/sec)
            //                              'sistTrunkRoM': '21±0.5',
            //                        //起立过程中躯干矢状面角度范围SiSt Trunk RoM（degree）
            //                              'stepLength_R': null,
            //                        //右脚步长Step Length R.（cm）
            //                              'stepLength_L': null,
            //                        //左脚步长Step Length L.（cm）
            //                              'stepLength_avg': null,
            //                        //左右脚平均步长Step Length（cm）
            //                              'strideVelocity_R': null,
            //                        //右脚速率Stride Velocity R.（m/s）
            //                              'strideVelocity_L': null,
            //                        //左脚速率Stride Velocity L.（m/s）
            //                              'strideVelocity_avg': null,
            //                        //左右脚平均速率Stride Velocity（m/s）
            //                              'strideLength_R': null,
            //                        //右脚跨步长Stride Length R.（cm）
            //                              'strideLength_L': null,
            //                        //左脚跨步长Stride Length L.（cm）
            //                              'strideLength_avg': null,
            //                        //左右脚平均跨步长Stride Length（cm）
            //                              'gaitCycle_R': null,
            //                        //右脚跨步时长Gait Cycle R.（sec）
            //                              'gaitCycle_L': null,
            //                        //左脚跨步时长Gait Cycle L.（sec）
            //                              'gaitCycle_avg': null,
            //                        //左右脚平均跨步时长Gait Cycle（sec）
            //                              'cadence_R': null,
            //                        //右脚步频Cadence R.（step/min）
            //                              'cadence_L': null,
            //                        //左脚步频Cadence L.（step/min）
            //                              'cadence_avg': null,
            //                        //左右脚平均步频Cadence（step/min）      
            //                        'support_R': null,
            //                        //右腿支撑相Right Support（%）
            //                              'support_L': null,
            //                        //左腿支撑相Left Support（%）
            //                              'support_avg': null,
            //                        //双腿支撑相Double Support（%）
            //                              'swing_R': null,
            //                        //右侧摆动相Swing R.（%）
            //                              'swing_L': null,
            //                        //左侧摆动相Swing L.（%）
            //                              'swing_avg': null,
            //                        //摆动相Swing（%）
            //
            //                              'stance_R': null,
            //                        //右侧支撑相Stance R.（%）
            //                              'stance_L': null,
            //                        //左侧支撑相Stance L.（%）
            //                              'stance_avg': null,
            //                        //支撑相Stance（%）
            //                              'trunkSagittalPeakVelocity': null,
            //                        //躯干矢状面运动峰值角速度Trunk Sagittal Peak Velocity（degree/sec）
            //                              'trunkHorizontalPeakVelocity': null,
            //                        //躯干横断面运动峰值角速度Trunk Horizontal Peak Velocity （degree/sec）
            //                              'strideVelocityAsymmetry': null,
            //                        //左右脚速率的相对偏差Stride Velocity Asymmetry（%）
            //                              'strideLengthAsymmetry': null,
            //                        //左右脚跨步长的相对偏差Stride Length Asymmetry（%）
            //                              'swingAsymmetry': null,
            //                        //左右腿摆动相相对偏差Swing Asymmetry（%）
            //                              'stanceAsymmetry': null,
            //                        //左右腿支撑相相对偏差Stance Asymmetry（%）
            //                              'shankAsymmetry': null,
            //                        //左右小腿角度范围的相对偏差Shank Asymmetry（%）
            //                              'peakShankVelocityAsymmetry': null,
            //                        //左右小腿速率的相对偏差Peak Shank Velocity Asymmetry（%）
            //                              'shank_SSI': null,
            //                        //小腿角度范围对称性指数Shank SSI  (Symbolic Symmetry Index)（%）
            //                              'meanPhaseDifference': null,
            //                        //小腿相对平均相位差Mean Phase Difference（%）
            //                              'phaseCoordinationIndex': null,
            //                        //小腿相对平均相位差Mean Phase Difference（%）
            //                              'balanceTrunkSagittalPeakVelocity': null,
            //                        //躯干矢状面峰值角速度Trunk Sagittal Peak Velocity（degree/sec）
            //                              'balanceTrunkHorizontalPeakVelocity': null
            //                        //躯干横断面峰值角速度Trunk Horizontal Peak Velocity（degree/sec）
            //                        }]
            //		                }
            //	                }
            //
            //            }";
            #endregion
            return mailbody;


        }

        #endregion

        private void ReadJsonID()
        {
            string result1 = DoPost("https://api.douban.com/v2/book/isbn/9787115212948", "10");

            //假数据
            //string A_Path = AppDomain.CurrentDomain.BaseDirectory + "json\\pdf.json";

            //result1 = readtxtJsom(A_Path);

            //JObject obj = (JObject)JsonConvert.DeserializeObject(result1);

            //string tx1 = obj["typegroupcode"].ToString();

            //new 
            #region MyRegion
            ////Newtonsoft.Json.Linq.JArray userAarray1 = Newtonsoft.Json.Linq.JArray.Parse(result1) as Newtonsoft.Json.Linq.JArray;
            ////List<clsOrderDatabaseinfo> userListModel = userAarray1.ToObject<List<clsOrderDatabaseinfo>>();
            //JavaScriptSerializer jss = new JavaScriptSerializer();
            ////1
            //result_Msg result = jss.Deserialize<result_Msg>(result1);
            //if (result.resultData.classinfo != null && result.resultData.classinfo.Count > 0)
            //{
            //    List<clsPDFchildTypeinfo> classinfos = result.resultData.classinfo;
            //}
            ////2
            //var jarr = jss.Deserialize<Dictionary<string, object>>(result1);
            //foreach (var j in jarr)
            //{
            //    Console.WriteLine(string.Format("{0}:{1}", j.Key, j.Value));
            //}

            ////            JArray array = (JArray)obj["article"];
            ////            if (array != null)

            ////                foreach (var jObject in array)
            ////                {
            ////                    //赋值属性
            ////                }
            ////            string inputJsonString = @"
            ////                [
            ////                    {StudentID:'100',Name:'aaa',Hometown:'china'},
            ////                    {StudentID:'101',Name:'bbb',Hometown:'us'},
            ////                    {StudentID:'102',Name:'ccc',Hometown:'england'}
            ////                ]";

            ////            string message = @"<table border='1'>
            ////                    <tr><td width='80'>StudentID</td><td width='100'>Name</td><td width='100'>Hometown</td></tr>";
            ////            string tpl = "<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>";
            ////            List<clsOrderDatabaseinfo> studentList = JsonConvert.DeserializeObject<List<clsOrderDatabaseinfo>>(inputJsonString);//注意这里必须为List<Student>类型,因为客户端提交的是一个数组json
            ////            foreach (clsOrderDatabaseinfo student in studentList)
            ////            {
            ////                //message += String.Format(tpl, student.StudentID, student.Name, student.Hometown);
            ////            } 
            #endregion

            //TEST_Tocovermethod2(result1);

            string A_Path = "";
            if (readtypeA != "" && readtypeA != "A")
            {
                if (interface_pdfinfo != null && interface_pdfinfo != "")
                    Tocovermethod_PDF(interface_pdfinfo);
                if (interface_excelinfo != null && interface_excelinfo != "")
                    Tocovermethod_excel(interface_excelinfo);
                if (interface_onlineinfo != null && interface_onlineinfo != "")
                    Tocovermethod_clientOnline(interface_onlineinfo);

            }
            else
            {
                A_Path = AppDomain.CurrentDomain.BaseDirectory + "json\\pdf.json";
                A_Path = "C:\\json\\pdf.json";
                result1 = readtxtJsom(A_Path);
                Tocovermethod_PDF(result1);

                A_Path = AppDomain.CurrentDomain.BaseDirectory + "json\\excel.json";
                A_Path = "C:\\json\\excel.json";
                result1 = readtxtJsom(A_Path);
                Tocovermethod_excel(result1);

                A_Path = AppDomain.CurrentDomain.BaseDirectory + "json\\Online.json";
                A_Path = "C:\\json\\Online.json";
                result1 = readtxtJsom(A_Path);
                Tocovermethod_clientOnline(result1);
            }
        }

        private void TEST_Tocovermethod2(string json)
        {
            string A_Path = AppDomain.CurrentDomain.BaseDirectory + "json\\1.json";
            json = readtxtJsom(A_Path);

            //第一次解析
            Dictionary<string, object> dic = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
            //获取具体数据部分5
            object obj = dic["resultData"];
            //将数据部分再次转换为json字符串
            string jsondata = Newtonsoft.Json.JsonConvert.SerializeObject(obj);
            //获取数据中的  不同类型的数据   
            Dictionary<string, object> dicc = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata);

            //chalssinfo 
            object objclass = dicc["classinfo"];
            string jsonclass = Newtonsoft.Json.JsonConvert.SerializeObject(objclass);
            List<clsPDFchildTypeinfo> tclass = Newtonsoft.Json.JsonConvert.DeserializeObject<List<clsPDFchildTypeinfo>>(jsonclass);
            //otherinfo 
            object objother = dicc["otherinfo"];
            string jsonother = Newtonsoft.Json.JsonConvert.SerializeObject(objother);
            DataTable tother = Newtonsoft.Json.JsonConvert.DeserializeObject<DataTable>(jsonother);


            // tclass 和  tother 里面分别存放 classinfo和otherinfo  然后可以操作datatale 或者转成list也行

        }
        private void Tocovermethod_PDF(string json)
        {


            //第一次解析
            Dictionary<string, object> dic = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            JObject mm = (JObject)JsonConvert.DeserializeObject(json.ToString());

            PDF_Root item = new PDF_Root();
            item.typegroupcode = mm["typegroupcode"].ToString();
            PDF_Rootdb.Add(item);

            object obj = dic["types"];
            string jsondata = Newtonsoft.Json.JsonConvert.SerializeObject(obj);
            //types
            JArray jArray = (JArray)JsonConvert.DeserializeObject(jsondata);//jsonArrayText必须是带[]数组格式字符串
            //
            PDF_Types = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Types>>(jsondata);

            //childType
            foreach (JToken jt in jArray)
            {

                string partentID = jt["typeCode"].ToString();

                JObject jo = (JObject)jt;

                JArray temp = (JArray)jo["childType"];

                foreach (JToken token in temp)
                {
                    ChildType itemTypes = new ChildType();

                    itemTypes.typeCode = token["typeCode"].ToString();
                    itemTypes.typeName = token["typeName"].ToString();
                    itemTypes.typeEnName = token["typeEnName"].ToString();
                    itemTypes.unit = token["unit"].ToString();
                    itemTypes.distinguish = token["distinguish"].ToString();

                    itemTypes.partentID = partentID;

                    PDF_ChildType.Add(itemTypes);
                }


            }


        }
        private void Tocovermethod_excel(string json)
        {

            //第一次解析
            Dictionary<string, object> dic = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
            //获取具体数据部分5
            object obj = dic["data"];
            //将数据部分再次转换为json字符串
            string jsondata = Newtonsoft.Json.JsonConvert.SerializeObject(obj);
            //childType
            Dictionary<string, object> dic1 = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata);

            object obj_header = dic1["header"];
            //获取数据中的  不同类型的数据   
            string jsondata_header = Newtonsoft.Json.JsonConvert.SerializeObject(obj_header);
            Dictionary<string, object> dicc = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata_header);

            object obj_rows = dicc["rows"];
            string jsondata_rows = Newtonsoft.Json.JsonConvert.SerializeObject(obj_rows);

            string jsonArrayText = jsondata_rows;
            JArray jArray = (JArray)JsonConvert.DeserializeObject(jsonArrayText);//jsonArrayText必须是带[]数组格式字符串

            string str = jArray[0]["cols"].ToString();
            //string jsondata_rows1 = Newtonsoft.Json.JsonConvert.SerializeObject(str);
            //List<clsExcelinfo> TBB = Newtonsoft.Json.JsonConvert.DeserializeObject<List<clsExcelinfo>>(jsondata_rows1);
            //JsonTextReader json111 = new JsonTextReader(new StringReader(json));
            //Dictionary<string, object> dicc_rows = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata_rows1);
            //while (json111.Read())
            //{

            //    Console.WriteLine(json111.Value + "--" + json111.TokenType + "--" + json111.ValueType);

            //}

            //
            JObject mm = (JObject)JsonConvert.DeserializeObject(json.ToString());

            Root item1 = new Root();
            item1.code = mm["code"].ToString();
            item1.msg = mm["msg"].ToString();
            item1.reqId = mm["reqId"].ToString();

            Root_datas.Add(item1);


            //cols
            int rowindex = 0;

            foreach (JToken jt in jArray)
            {

                JObject jo = (JObject)jt;

                JArray temp = (JArray)jo["cols"];
                rowindex++;


                foreach (JToken token in temp)
                {
                    clsExcelinfo item = new clsExcelinfo();
                    item.text = token["text"].ToString();
                    item.size = token["size"].ToString();
                    item.cols = rowindex.ToString();
                    TBB.Add(item);
                }
            }
            //datas
            object obj_content = dic1["content"];
            string jsondata_content = Newtonsoft.Json.JsonConvert.SerializeObject(obj_content);
            Dictionary<string, object> dic121 = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata_content);

            object obj_datas = dic121["datas"];
            string jsondata_datas = Newtonsoft.Json.JsonConvert.SerializeObject(obj_datas);
            tclass_datas = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Datas>>(jsondata_datas);
            tclass_datas[0].reqId = item1.reqId;





            Excel_body = Newtonsoft.Json.JsonConvert.DeserializeObject<DataTable>(jsondata_datas);

        }

        private void Tocovermethod_clientOnline(string json)
        {

            Dictionary<string, object> dic = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            object obj = dic["data"];

            string jsondata = Newtonsoft.Json.JsonConvert.SerializeObject(obj);

            Dictionary<string, object> dic1 = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(jsondata);


            JObject mm1 = (JObject)JsonConvert.DeserializeObject(json.ToString());

            Online_Root item1 = new Online_Root();
            item1.code = mm1["code"].ToString();
            item1.msg = mm1["msg"].ToString();
            item1.reqId = mm1["reqId"].ToString();

            Online_Root_datas.Add(item1);


            //
            Online_Data item = new Online_Data();

            JObject mm = (JObject)JsonConvert.DeserializeObject(obj.ToString());
            //再讲字符串转成json格式
            item.id = mm["id"].ToString();
            item.serialNumber = mm["serialNumber"].ToString();
            item.patientId = mm["patientId"].ToString();
            item.acquisitionStartTime = mm["acquisitionStartTime"].ToString();
            item.acquisitionEndTime = mm["acquisitionEndTime"].ToString();
            item.equipmentModel = mm["equipmentModel"].ToString();
            item.equipmentNumber = mm["equipmentNumber"].ToString();
            item.checkNumber = mm["checkNumber"].ToString();
            item.dataSources = mm["dataSources"].ToString();
            item.type = mm["type"].ToString();
            item.remark = mm["remark"].ToString();

            Online_datas.Add(item);

            object obj_maGait = dic1["maGait"];
            string json_maGait = Newtonsoft.Json.JsonConvert.SerializeObject(obj_maGait);
            Online_MaGait = Newtonsoft.Json.JsonConvert.DeserializeObject<List<MaGait>>(json_maGait);

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


        public void InitializeDataSource(List<clsExcelinfo> TBB1, List<Datas> tclass_datas1, List<Root> Root_datas1, List<Online_Data> Online_datas1, List<MaGait> Online_MaGait1, List<Online_Root> Online_Root_datas1, List<PDF_Root> PDF_Rootdb1, List<Types> PDF_Types1, List<ChildType> PDF_ChildType1)
        {
            //excel
            TBB = new List<clsExcelinfo>();
            tclass_datas = new List<Datas>();
            Root_datas = new List<Root>();

            Online_datas = new List<Online_Data>();
            Online_MaGait = new List<MaGait>();
            Online_Root_datas = new List<Online_Root>();

            PDF_Rootdb = new List<PDF_Root>();
            PDF_Types = new List<Types>();
            PDF_ChildType = new List<ChildType>();

            TBB = TBB1;
            tclass_datas = tclass_datas1;
            Root_datas = Root_datas1;

            Online_datas = Online_datas1;
            Online_MaGait = Online_MaGait1;
            Online_Root_datas = Online_Root_datas1;

            PDF_Rootdb = PDF_Rootdb1;
            PDF_Types = PDF_Types1;
            PDF_ChildType = PDF_ChildType1;
        }
        public void DownLoadExcel(ref BackgroundWorker bgWorker, string pathname)
        {
            bgWorker1 = bgWorker;

            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "Model.xlsx");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            string[] temp1 = System.Text.RegularExpressions.Regex.Split(pathname, " ");
            sfdDownFile.FileName = pathname;

            string strExcelFileName = string.Empty;
            #endregion

            #region 导出前校验模板信息
            if (string.IsNullOrEmpty(sfdDownFile.FileName))
            {
                MessageBox.Show("File name can't be empty, please confirm, thanks!");
                return;
            }
            if (!File.Exists(fullPath))
            {
                MessageBox.Show("Template file does not exist, please confirm, thanks!");
                return;
            }
            else
            {
                strExcelFileName = sfdDownFile.FileName;
                strExcelFileName = pathname;

            }
            #endregion
            #region Excel 初始化

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel._Workbook ExcelBook =
            ExcelApp.Workbooks.Open(fullPath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);
            #endregion
            #region Sheet 初始化
            try
            {
                Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[1];
                //打开时是否显示Excel
                //ExcelApp.Visible = true;
                //ExcelApp.ScreenUpdating = true;
            #endregion

                #region 填充数据
                int RowIndex = 2;
                int doing = 0;
                if (tclass_datas != null)
                    foreach (Datas item in tclass_datas)
                    {


                        doing++;

                        bgWorker1.ReportProgress(0, "导出进度  :  " + RowIndex.ToString() + "/" + tclass_datas.Count.ToString());

                        RowIndex++;
                        #region MyRegion
                        ExcelApp.Visible = true;
                        ExcelApp.ScreenUpdating = true;
                        ExcelSheet.Cells[RowIndex, 1] = RowIndex - 2;
                        ExcelSheet.Cells[RowIndex, 2] = item.patientId;
                        ExcelSheet.Cells[RowIndex, 3] = item.patientName;
                        ExcelSheet.Cells[RowIndex, 4] = item.sex;
                        ExcelSheet.Cells[RowIndex, 5] = item.birthday;
                        ExcelSheet.Cells[RowIndex, 6] = "'" + item.diseaseType;
                        ExcelSheet.Cells[RowIndex, 7] = item.courseOfDisease;
                        ExcelSheet.Cells[RowIndex, 8] = "'" + item.serialNumber;
                        ExcelSheet.Cells[RowIndex, 9] = "'" + item.inspectId;
                        ExcelSheet.Cells[RowIndex, 10] = "'" + item.checkType;
                        ExcelSheet.Cells[RowIndex, 11] = "'" + item.dualTask;
                        ExcelSheet.Cells[RowIndex, 12] = item.checkRemark;
                        ExcelSheet.Cells[RowIndex, 13] = item.checkStartTime;
                        ExcelSheet.Cells[RowIndex, 14] = item.checkEndTime;
                        ExcelSheet.Cells[RowIndex, 15] = "'" + item.checkHospitalId;
                        ExcelSheet.Cells[RowIndex, 16] = item.checkHospitalName;
                        ExcelSheet.Cells[RowIndex, 17] = item.checkDoctorId;
                        ExcelSheet.Cells[RowIndex, 18] = item.checkDoctorName;
                        ExcelSheet.Cells[RowIndex, 19] = item.auditDoctorName;
                        ExcelSheet.Cells[RowIndex, 20] = item.checkConclusion;
                        ExcelSheet.Cells[RowIndex, 21] = "'" + item.sistDuration;
                        ExcelSheet.Cells[RowIndex, 22] = "'" + item.peakSiStAngularVelocity;
                        ExcelSheet.Cells[RowIndex, 23] = "'" + item.sistTrunkRoM;
                        ExcelSheet.Cells[RowIndex, 24] = item.stepLength_R;
                        ExcelSheet.Cells[RowIndex, 25] = item.stepLength_L;
                        ExcelSheet.Cells[RowIndex, 26] = item.stepLength_avg;
                        ExcelSheet.Cells[RowIndex, 27] = item.strideVelocity_R;
                        ExcelSheet.Cells[RowIndex, 28] = item.strideVelocity_L;
                        ExcelSheet.Cells[RowIndex, 29] = item.strideVelocity_avg;
                        ExcelSheet.Cells[RowIndex, 30] = item.strideLength_R;
                        ExcelSheet.Cells[RowIndex, 31] = item.strideLength_L;
                        ExcelSheet.Cells[RowIndex, 32] = item.strideLength_avg;
                        ExcelSheet.Cells[RowIndex, 33] = item.gaitCycle_R;
                        ExcelSheet.Cells[RowIndex, 34] = item.gaitCycle_L;
                        ExcelSheet.Cells[RowIndex, 35] = item.gaitCycle_avg;
                        ExcelSheet.Cells[RowIndex, 36] = item.cadence_R;
                        ExcelSheet.Cells[RowIndex, 37] = item.cadence_L;
                        ExcelSheet.Cells[RowIndex, 38] = item.cadence_avg;
                        ExcelSheet.Cells[RowIndex, 39] = item.support_R;
                        ExcelSheet.Cells[RowIndex, 40] = item.support_L;
                        ExcelSheet.Cells[RowIndex, 41] = item.support_avg;
                        ExcelSheet.Cells[RowIndex, 42] = item.swing_R;
                        ExcelSheet.Cells[RowIndex, 43] = item.swing_L;
                        ExcelSheet.Cells[RowIndex, 44] = item.swing_avg;
                        ExcelSheet.Cells[RowIndex, 45] = item.stance_R;
                        ExcelSheet.Cells[RowIndex, 46] = item.stance_L;
                        ExcelSheet.Cells[RowIndex, 47] = item.stance_avg;
                        ExcelSheet.Cells[RowIndex, 48] = item.trunkSagittalPeakVelocity;
                        ExcelSheet.Cells[RowIndex, 49] = item.trunkHorizontalPeakVelocity;
                        ExcelSheet.Cells[RowIndex, 50] = item.strideVelocityAsymmetry;
                        ExcelSheet.Cells[RowIndex, 51] = item.strideLengthAsymmetry;
                        ExcelSheet.Cells[RowIndex, 52] = item.swingAsymmetry;
                        ExcelSheet.Cells[RowIndex, 53] = item.stanceAsymmetry;
                        ExcelSheet.Cells[RowIndex, 54] = item.shankAsymmetry;
                        ExcelSheet.Cells[RowIndex, 55] = item.peakShankVelocityAsymmetry;
                        ExcelSheet.Cells[RowIndex, 56] = item.shank_SSI;
                        ExcelSheet.Cells[RowIndex, 57] = item.meanPhaseDifference;
                        ExcelSheet.Cells[RowIndex, 58] = item.phaseCoordinationIndex;
                        ExcelSheet.Cells[RowIndex, 59] = item.balanceTrunkSagittalPeakVelocity;
                        ExcelSheet.Cells[RowIndex, 60] = item.balanceTrunkHorizontalPeakVelocity;
                        #endregion

                    }
                ExcelBook.RefreshAll();
                #region 写入文件
                ExcelApp.ScreenUpdating = true;
                if (doing != 0)
                    ExcelBook.SaveAs(strExcelFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                ExcelApp.DisplayAlerts = false;

                #endregion
            }

            #region 异常处理
            catch (Exception ex)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
                ExcelBook = null;
                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                throw ex;
            }
            #endregion

            #region Finally垃圾回收
            finally
            {
                ExcelBook.Close(false, missingValue, missingValue);
                ExcelBook = null;
                ExcelApp.DisplayAlerts = true;
                ExcelApp.Quit();
                clsKeyMyExcelProcess.Kill(ExcelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion

                #endregion
        }
        public string XLSSavesaCSV(string FilePath)
        {
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            QuertExcel();
            string _NewFilePath = "";

            Excel.Application excelApplication;
            Excel.Workbooks excelWorkBooks = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet excelWorkSheet = null;

            try
            {
                excelApplication = new Excel.Application();
                excelWorkBooks = excelApplication.Workbooks;
                excelWorkBook = ((Excel.Workbook)excelWorkBooks.Open(FilePath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue));
                excelWorkSheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];
                excelApplication.Visible = false;
                excelApplication.DisplayAlerts = false;
                _NewFilePath = FilePath.Replace(".xlsx", ".csv");

                // excelWorkSheet._SaveAs(FilePath, Excel.XlFileFormat.xlCSVWindows, missingValue, missingValue, missingValue,missingValue,missingValue, missingValue, missingValue);
                excelWorkBook.SaveAs(_NewFilePath, Excel.XlFileFormat.xlCSV, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                QuertExcel();
                //ExcelFormatHelper.DeleteFile(FilePath);
            }
            catch (Exception exc)
            {
                throw new Exception(exc.Message);
            }
            return _NewFilePath;
        }
        private static void QuertExcel()
        {
            Process[] excels = Process.GetProcessesByName("EXCEL");
            foreach (var item in excels)
            {
                item.Kill();
            }
        }



        public void DownLoadPDF(ref BackgroundWorker bgWorker, string pathname)
        {
            bgWorker1 = bgWorker;
            DownLoadPDFExcel();
            if (XLSConvertToPDF(savepdfexcel_path, savepdfexcel_path.Replace("xlsx", "pdf")))
            {
                var dir = System.IO.Path.GetDirectoryName(pathname);
                string namesave = System.IO.Path.GetFileName(savepdfexcel_path);
                File.Copy(savepdfexcel_path.Replace("xlsx", "pdf"), dir + "\\" + namesave.Replace("xlsx", "pdf"));

                File.Delete(savepdfexcel_path);
                File.Delete(savepdfexcel_path.Replace("xlsx", "pdf"));
            }



        }
        public void DownLoadPDFExcel()
        {


            #region 获取模板路径
            System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            string fullPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "System\\"), "TEL.xlsx");
            SaveFileDialog sfdDownFile = new SaveFileDialog();
            sfdDownFile.OverwritePrompt = false;
            string DesktopPath = Convert.ToString(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            sfdDownFile.Filter = "Excel files (*.xls,*.xlsx)|*.xls;*.xlsx";
            string file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results\\");
            string pathname = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") +
                ".xlsx";

            sfdDownFile.FileName = file + pathname;

            string strExcelFileName = string.Empty;
            #endregion

            #region 导出前校验模板信息
            if (string.IsNullOrEmpty(sfdDownFile.FileName))
            {
                MessageBox.Show("File name can't be empty, please confirm, thanks!");
                return;
            }
            if (!File.Exists(fullPath))
            {
                MessageBox.Show("Template file does not exist, please confirm, thanks!");
                return;
            }
            else
            {
                strExcelFileName = sfdDownFile.FileName;
                //  strExcelFileName = pathname;

            }
            #endregion
            #region Excel 初始化

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel._Workbook ExcelBook =
            ExcelApp.Workbooks.Open(fullPath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);
            #endregion
            #region Sheet 初始化
            try
            {
                Microsoft.Office.Interop.Excel._Worksheet ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[1];
                //打开时是否显示Excel
                //ExcelApp.Visible = true;
                //ExcelApp.ScreenUpdating = true;
            #endregion

                #region 填充数据
                int doing = 0;
                if (tclass_datas != null)
                    foreach (Datas item in tclass_datas)
                    {

                        int RowIndex = 2;
                        doing++;

                        bgWorker1.ReportProgress(0, "导出进度  :  " + RowIndex.ToString() + "/" + tclass_datas.Count.ToString());

                        RowIndex++;
                        #region MyRegion
                        //ExcelApp.Visible = true;
                        //ExcelApp.ScreenUpdating = true;
                        ExcelSheet.Cells[3, 3] = item.inspectId;
                        ExcelSheet.Cells[3, 7] = item.serialNumber;
                        ExcelSheet.Cells[6, 3] = item.patientName;
                        ExcelSheet.Cells[6, 5] = item.diseaseType;
                        ExcelSheet.Cells[8, 3] = item.checkType;
                        ExcelSheet.Cells[8, 5] = item.checkStartTime;
                        ExcelSheet.Cells[10, 3] = item.checkDoctorName;
                        ExcelSheet.Cells[10, 8] = item.auditDoctorName;
                        ExcelSheet.Cells[14, 2] = item.checkConclusion;
                        ExcelSheet.Cells[43, 3] = item.checkDoctorName;
                        ExcelSheet.Cells[43, 6] = item.auditDoctorName;
                        ExcelSheet.Cells[45, 3] = DateTime.Now.ToString("MM/dd/yyyy");
                        ExcelSheet.Cells[50, 5] = item.checkHospitalName;

                        #endregion
                        ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[2];

                        ExcelSheet.Cells[7, 6] = item.sistDuration;
                        ExcelSheet.Cells[8, 6] = item.peakSiStAngularVelocity;
                        ExcelSheet.Cells[9, 6] = item.sistTrunkRoM;
                        ExcelSheet.Cells[23, 6] = item.balanceTrunkSagittalPeakVelocity;
                        ExcelSheet.Cells[24, 6] = item.balanceTrunkHorizontalPeakVelocity;

                        ExcelSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelBook.Worksheets[3];

                        ExcelSheet.Cells[7, 6] = item.stepLength_avg;
                        ExcelSheet.Cells[7, 7] = item.stepLength_R;
                        ExcelSheet.Cells[7, 8] = item.stepLength_L;
                        ExcelSheet.Cells[8, 6] = item.strideVelocity_avg;
                        ExcelSheet.Cells[8, 7] = item.strideVelocity_R;
                        ExcelSheet.Cells[8, 8] = item.strideVelocity_L;
                        ExcelSheet.Cells[9, 6] = item.strideLength_avg;
                        ExcelSheet.Cells[9, 7] = item.strideLength_R;
                        ExcelSheet.Cells[9, 8] = item.strideLength_L;
                        ExcelSheet.Cells[10, 6] = item.gaitCycle_avg;
                        ExcelSheet.Cells[10, 7] = item.gaitCycle_R;
                        ExcelSheet.Cells[10, 8] = item.gaitCycle_L;
                        ExcelSheet.Cells[11, 6] = item.cadence_avg;
                        ExcelSheet.Cells[11, 7] = item.cadence_R;
                        ExcelSheet.Cells[11, 8] = item.cadence_L;
                        ExcelSheet.Cells[12, 6] = item.support_avg;
                        ExcelSheet.Cells[12, 7] = item.support_R;
                        ExcelSheet.Cells[12, 8] = item.support_L;
                        ExcelSheet.Cells[13, 6] = item.swing_avg;
                        ExcelSheet.Cells[13, 7] = item.swing_R;
                        ExcelSheet.Cells[13, 8] = item.swing_L;
                        ExcelSheet.Cells[14, 6] = item.stance_avg;
                        ExcelSheet.Cells[14, 7] = item.stance_R;
                        ExcelSheet.Cells[14, 8] = item.stance_L;
                        ExcelSheet.Cells[15, 6] = item.trunkSagittalPeakVelocity;
                        ExcelSheet.Cells[16, 6] = item.trunkHorizontalPeakVelocity;
                        ExcelSheet.Cells[17, 6] = item.strideVelocityAsymmetry;
                        ExcelSheet.Cells[18, 6] = item.strideLengthAsymmetry;
                        ExcelSheet.Cells[19, 6] = item.swingAsymmetry;
                        ExcelSheet.Cells[20, 6] = item.stanceAsymmetry;
                        ExcelSheet.Cells[21, 6] = item.shankAsymmetry;
                        ExcelSheet.Cells[22, 6] = item.peakShankVelocityAsymmetry;
                        ExcelSheet.Cells[23, 6] = item.shank_SSI;
                        ExcelSheet.Cells[24, 6] = item.meanPhaseDifference;
                        ExcelSheet.Cells[25, 6] = "";

                    }
                ExcelBook.RefreshAll();
                #region 写入文件
                ExcelApp.ScreenUpdating = true;

                ExcelBook.SaveAs(strExcelFileName, missingValue, missingValue, missingValue, missingValue, missingValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, missingValue, missingValue, missingValue, missingValue, missingValue);
                ExcelApp.DisplayAlerts = false;
                savepdfexcel_path = strExcelFileName;

                #endregion
            }

            #region 异常处理
            catch (Exception ex)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Quit();
                ExcelBook = null;
                ExcelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                throw ex;
            }
            #endregion

            #region Finally垃圾回收
            finally
            {
                ExcelBook.Close(false, missingValue, missingValue);
                ExcelBook = null;
                ExcelApp.DisplayAlerts = true;
                ExcelApp.Quit();
                clsKeyMyExcelProcess.Kill(ExcelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            #endregion

                #endregion
        }

        private bool XLSConvertToPDF(string sourcePath, string targetPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Application ExcelApp = null;
            Microsoft.Office.Interop.Excel._Workbook ExcelBook = null;
            try
            {

                object target = targetPath;
                object type = targetType;

                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                System.Reflection.Missing missingValue = System.Reflection.Missing.Value;
                ExcelBook = ExcelApp.Workbooks.Open(sourcePath, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue, missingValue);

                ExcelBook.ExportAsFixedFormat(targetType, target, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                result = true;


            }
            catch
            {
                result = false;
            }
            finally
            {
                if (ExcelBook != null)
                {
                    ExcelBook.Close(true, missing, missing);
                    ExcelBook = null;
                }
                if (ExcelApp != null)
                {
                    ExcelApp.Quit();
                    ExcelApp = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }


    }
}
