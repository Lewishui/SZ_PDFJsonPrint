﻿using Microsoft.Reporting.WinForms;
using SPJP.Buiness;
using SPJP.Common;
using SPJP.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SZ_PDFJsonPrint;

namespace Interface_Json
{
    public class MainPort
    {
        private delegate void InvokeHandler();
        public log4net.ILog ProcessLogger;
        public log4net.ILog ExceptionLogger;
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        private SortableBindingList<clsOrderDatabaseinfo> sortablePendingOrderList;
        List<clsOrderDatabaseinfo> FilterOrderResults;

        public ReportForm reportForm;

        public Microsoft.Reporting.WinForms.ReportViewer reportViewer1;

        string strFileName;
        DataTable Excel_body;

        //excel
        List<clsExcelinfo> TBB;
        List<Datas> tclass_datas;
        List<Root> Root_datas;
        List<Datas> printclass_datas;


        //客户查找
        List<Online_Data> Online_datas;
        List<MaGait> Online_MaGait;
        List<Online_Root> Online_Root_datas;
        List<OnlineShow> OnlineShow_datas;

        //PDF
        List<PDF_Root> PDF_Rootdb;
        List<Types> PDF_Types;
        List<ChildType> PDF_ChildType;
        List<PDF_checkdataDetail> PDFcheckdataDetail;

        bool issaveok = false;

        DataTable qtyTable_;


        private string interface_pdfinfo;
        private string interface_excelinfo;
        private string interface_onlineinfo;

        public MainPort()
        {
            FilterOrderResults = new List<clsOrderDatabaseinfo>();
            InitializeDataSource();


     
            reportForm = new ReportForm();
        }
        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }
        private void InitializeDataSource()
        {

            clsAllnew BusinessHelp = new clsAllnew();

        }
        public class SortableBindingList<T> : BindingList<T>
        {
            private bool isSortedCore = true;
            private ListSortDirection sortDirectionCore = ListSortDirection.Ascending;
            private PropertyDescriptor sortPropertyCore = null;
            private string defaultSortItem;

            public SortableBindingList() : base() { }

            public SortableBindingList(IList<T> list) : base(list) { }

            protected override bool SupportsSortingCore
            {
                get { return true; }
            }

            protected override bool SupportsSearchingCore
            {
                get { return true; }
            }

            protected override bool IsSortedCore
            {
                get { return isSortedCore; }
            }

            protected override ListSortDirection SortDirectionCore
            {
                get { return sortDirectionCore; }
            }

            protected override PropertyDescriptor SortPropertyCore
            {
                get { return sortPropertyCore; }
            }

            protected override int FindCore(PropertyDescriptor prop, object key)
            {
                for (int i = 0; i < this.Count; i++)
                {
                    if (Equals(prop.GetValue(this[i]), key)) return i;
                }
                return -1;
            }

            protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
            {
                isSortedCore = true;
                sortPropertyCore = prop;
                sortDirectionCore = direction;
                Sort();
            }

            protected override void RemoveSortCore()
            {
                if (isSortedCore)
                {
                    isSortedCore = false;
                    sortPropertyCore = null;
                    sortDirectionCore = ListSortDirection.Ascending;
                    Sort();
                }
            }

            public string DefaultSortItem
            {
                get { return defaultSortItem; }
                set
                {
                    if (defaultSortItem != value)
                    {
                        defaultSortItem = value;
                        Sort();
                    }
                }
            }

            private void Sort()
            {
                List<T> list = (this.Items as List<T>);
                list.Sort(CompareCore);
                ResetBindings();
            }

            private int CompareCore(T o1, T o2)
            {
                int ret = 0;
                if (SortPropertyCore != null)
                {
                    ret = CompareValue(SortPropertyCore.GetValue(o1), SortPropertyCore.GetValue(o2), SortPropertyCore.PropertyType);
                }
                if (ret == 0 && DefaultSortItem != null)
                {
                    PropertyInfo property = typeof(T).GetProperty(DefaultSortItem, BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.IgnoreCase, null, null, new Type[0], null);
                    if (property != null)
                    {
                        ret = CompareValue(property.GetValue(o1, null), property.GetValue(o2, null), property.PropertyType);
                    }
                }
                if (SortDirectionCore == ListSortDirection.Descending) ret = -ret;
                return ret;
            }

            private static int CompareValue(object o1, object o2, Type type)
            {
                if (o1 == null) return o2 == null ? 0 : -1;
                else if (o2 == null) return 1;
                else if (type.IsPrimitive || type.IsEnum) return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
                else if (type == typeof(DateTime)) return Convert.ToDateTime(o1).CompareTo(o2);
                else return String.Compare(o1.ToString().Trim(), o2.ToString().Trim());
            }
        }


        private void PrintReportForEDI()
        {

            if (PDF_Types == null)
                return;
            #region maintain
            Create_table(true);
            #endregion

        }

        public void Readform_interface(string pdfinfo, string excelinfo, string onlineinfo)
        {
            interface_pdfinfo = pdfinfo;
            interface_excelinfo = excelinfo;
            interface_onlineinfo = onlineinfo;

            #region Noway
            DateTime oldDate = DateTime.Now;
            DateTime dt3;
            string endday = DateTime.Now.ToString("yyyy/MM/dd");
            dt3 = Convert.ToDateTime(endday);
            DateTime dt2;
            dt2 = Convert.ToDateTime("2018/09/14");

            TimeSpan ts = dt2 - dt3;
            int timeTotal = ts.Days;
            if (timeTotal < 0)
            {
                MessageBox.Show("Error 23082:");
                Application.Exit();

                return;
            }
            #endregion

            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(ReadJSONfromServer);

                bgWorker.RunWorkerAsync();

                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();

                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    #region Excel
                    #region 构造 datatable
                    qtyTable_ = new DataTable();
                    var qtyTable = new DataTable();
                    //第一行
                    // qtyTable.Rows.Add(qtyTable.NewRow());
                    int cl = 0;
                    //foreach (clsExcelinfo k in TBB)
                    //{
                    //    if (k.cols == "1")
                    //    {
                    //        qtyTable.Columns.Add(k.text, System.Type.GetType("System.String"));
                    //        //qtyTable.Rows[0][cl] = k.text;
                    //        cl++;
                    //    }
                    //}
                    //第二行
                    qtyTable.Rows.Add(qtyTable.NewRow());
                    int cl2 = 0;
                    foreach (clsExcelinfo k in TBB)
                    {
                        if (k.cols == "2")
                        {
                            if (cl2 < cl)
                            {
                                qtyTable.Rows[0][cl2] = k.text;

                                cl2++;
                            }
                            else
                            {
                                qtyTable.Columns.Add(k.text, System.Type.GetType("System.String"));
                                cl2++;
                            }
                        }
                    }
                    foreach (Datas k in tclass_datas)
                    {
                        qtyTable.Rows.Add(qtyTable.NewRow());
                    }
                    int rowi = 1;

                    foreach (Datas k in tclass_datas)
                    {
                        //qtyTable.Rows[rowi][0] = k.Maichangdaima;
                    }
                    for (int row = 0; row < Excel_body.Rows.Count; row++)
                        for (int cloumn = 0; cloumn < Excel_body.Columns.Count; cloumn++)
                            //foreach (DataColumn dcn in Excel_body.Columns)
                            if (row < Excel_body.Rows.Count && cloumn < Excel_body.Columns.Count)
                            {
                                qtyTable.Rows[row][cloumn] = Excel_body.Rows[row][cloumn].ToString();
                            }

                    #endregion
                    qtyTable_ = qtyTable;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化信息失败，请检查数据格式！" + ex, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
                throw ex;
            }
        }
        private void ReadJSONfromServer(object sender, DoWorkEventArgs e)
        {
            PDF_Rootdb = new List<PDF_Root>();
            PDF_Types = new List<Types>();
            PDF_ChildType = new List<ChildType>();

            Online_datas = new List<Online_Data>();
            Online_MaGait = new List<MaGait>();
            Online_Root_datas = new List<Online_Root>();

            TBB = new List<clsExcelinfo>();
            tclass_datas = new List<Datas>();
            Root_datas = new List<Root>();
            Excel_body = new DataTable();
            OnlineShow_datas = new List<OnlineShow>();
            clsAllnew BusinessHelp = new clsAllnew();
            //导入程序集
            DateTime oldDate = DateTime.Now;
            BusinessHelp.ReadJSON_Report(ref this.bgWorker, "B", interface_pdfinfo, interface_excelinfo, interface_onlineinfo);

            PDF_Rootdb = BusinessHelp.PDF_Rootdb;
            PDF_Types = BusinessHelp.PDF_Types;
            PDF_ChildType = BusinessHelp.PDF_ChildType;

            TBB = BusinessHelp.TBB;
            tclass_datas = BusinessHelp.tclass_datas;
            Root_datas = BusinessHelp.Root_datas;
            Excel_body = BusinessHelp.Excel_body;

            Online_datas = BusinessHelp.Online_datas;
            Online_MaGait = BusinessHelp.Online_MaGait;
            Online_Root_datas = BusinessHelp.Online_Root_datas;
            OnlineShow_datas = BusinessHelp.OnlineShow_datas;


            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ": " + s.Seconds.ToString() + "  (分:秒)";
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

        }
        /// <summary>
        /// 下载CSV 
        /// </summary> 好用 替换成 Eexcel
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 

        //public void DownSystemFile_CSV()
        //{
        //    //ExportCSV(qtyTable_, "");
        //    //好用
        //    downcsv(qtyTable_);
        //    //

        //}
        /// <summary>
        /// 下载excel 
        /// </summary>
        public void DownSystemFile_CSV()
        {

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".xls";
            saveFileDialog.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb,*.csv)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv";
            strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(downreport);
                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results"), "");
                    System.Diagnostics.Process.Start("explorer.exe", strFileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ex" + ex);
                return;
                throw ex;
            }
        }
        private void downreport(object sender, DoWorkEventArgs e)
        {
            DateTime oldDate = DateTime.Now;
            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.InitializeDataSource(TBB, tclass_datas, Root_datas, Online_datas, Online_MaGait, Online_Root_datas, PDF_Rootdb, PDF_Types, PDF_ChildType, Excel_body);

            //BusinessHelp.pbStatus = pbStatus;
            //BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            BusinessHelp.DownLoadExcel(ref this.bgWorker, strFileName);

            //    BusinessHelp.XLSSavesaCSV(strFileName);
            //暂停
            //BusinessHelp.DownLoadPDF(ref this.bgWorker, strFileName);

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_015 + "\r\n" + Showtime);
        }
        private void downcsv(DataTable dataGridView)
        {
            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                strHeader += dataGridView.Columns[i].ColumnName + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < dataGridView.Rows.Count - 1; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView.Columns.Count; k++)
                {
                    if (dataGridView.Rows[j][k].ToString() != null)
                    {
                        // if (k == 7 || k == 5 || k == 8 || k == 9 || k == 10 || k == 5)
                        strRowValue += "'" + dataGridView.Rows[j][k].ToString().Replace("\r\n", " ").Replace("\n", "'") + delimiter;
                        //if (dataGridView.Rows[j].Cells[k].Value != null)
                        //    strRowValue +=   ((char)(9)).ToString() +dataGridView.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ") + delimiter;
                        //else
                        //    strRowValue +=  ((char)(9)).ToString()+  dataGridView.Rows[j].Cells[k].Value + delimiter;
                        // else
                        strRowValue += dataGridView.Rows[j][k].ToString().Replace("\r\n", " ").Replace("\n", "'") + delimiter;

                    }
                    else
                    {
                        strRowValue += dataGridView.Rows[j][k].ToString() + delimiter;
                    }
                }

                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成 ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        #region datatable  download csv
        public void ExportCSV(DataTable dataTable, string fileName)
        {
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();

            }
            else
            {
                return;

            }

            StreamWriter StreamWriter = new StreamWriter(strFileName, false, System.Text.Encoding.GetEncoding("gb2312"));
            StreamWriter.WriteLine(GetCSVFormatData(dataTable).ToString());
            StreamWriter.Flush();
            StreamWriter.Close();
        }

        /// <summary>
        /// 通过DataTable获得CSV格式数据
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <returns>CSV字符串数据</returns>
        public static StringBuilder GetCSVFormatData(DataTable dataTable)
        {
            StringBuilder StringBuilder = new StringBuilder();
            // 写出表头
            string title = "";

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                title += dataTable.Columns[i].ColumnName + ",";

            }
            StringBuilder.Append(title);
            StringBuilder.Append("\n");
            // 写出数据
            int count = 0;
            foreach (DataRowView dataRowView in dataTable.DefaultView)
            {
                count++;
                foreach (DataColumn DataColumn in dataTable.Columns)
                {
                    string field = dataRowView[DataColumn.ColumnName].ToString();

                    if (field.IndexOf('"') >= 0)
                    {
                        field = field.Replace("\"", "\"\"");
                    }
                    field = field.Replace("  ", " ");
                    if (field.IndexOf(',') >= 0 || field.IndexOf('"') >= 0 || field.IndexOf('<') >= 0 || field.IndexOf('>') >= 0 || field.IndexOf("'") >= 0)
                    {
                        field = "\"" + field + "\"";
                    }
                    if (DataColumn.ColumnName == "检查订单流水号" || DataColumn.ColumnName == "检查订单流水号" || DataColumn.ColumnName == "检查订单流水号" || DataColumn.ColumnName == "检查订单流水号" || DataColumn.ColumnName == "检查订单流水号")
                        field = "'" + field;

                    StringBuilder.Append(field + ",");
                    field = string.Empty;
                }
                if (count != dataTable.Rows.Count)
                {
                    StringBuilder.Append("\n");
                }
            }
            return StringBuilder;
        }

        #endregion


        public void ViewPDF()
        {
            //pdfExport();
            PrintReportForEDI();
        }

        public void PDFprint()
        {
            try
            {
                Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ex" + ex);
                return;
                throw;
            }
        }

        private void Print()
        {
            if (tclass_datas == null)
                return;
  
            clsAllnew BusinessHelp = new clsAllnew();
            //new

            #region new pirnt
            PDFcheckdataDetail = new List<PDF_checkdataDetail>();
            DataTable qtyTable = PDFcheckdaadetail_Method();


            int allpage_count;
            bool have_yushu;
            int ongong_index = 0;
            int vvv = 0;
            get_pagecount_andYushu(out allpage_count, out have_yushu);

            BusinessHelp.RunR1(OnlineShow_datas);

            foreach (Types item in PDF_Types)
            {
                string typecode = "";
                string typecode2 = "";
                if (ongong_index < allpage_count)
                {
                    typecode = PDF_Types[vvv].typeCode;
                    typecode2 = PDF_Types[vvv + 1].typeCode;

                    List<PDF_checkdataDetail> findsapinfo = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode);

                    List<PDF_checkdataDetail> findsapinfo2 = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode2);
                    BusinessHelp.RunR2(OnlineShow_datas, findsapinfo, findsapinfo2);
                }
                ongong_index++;
                vvv = ongong_index + 1;
            }
            if (have_yushu == true)
            {
                string typecode = PDF_Types[PDF_Types.Count - 1].typeCode;
                List<PDF_checkdataDetail> findsapinfo = PDFcheckdataDetail.FindAll(o => o.typeCode != null && o.typeCode == typecode);
                BusinessHelp.RunR3(OnlineShow_datas, findsapinfo);
            }
            #endregion
            MessageBox.Show("打印完成！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pdfExport(object sender, DoWorkEventArgs e)
        {
            DateTime oldDate = DateTime.Now;

            //回写测试，如果可以回写到数据集中，然后更新RDLC，并输出，那么就用这种模式

            //refresh report viewer
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string extension;

            reportViewer1 = new ReportViewer();

            //   reportForm.InitializeDataSource(printclass_datas);//tclass_datas

            //new 
            Create_table(false);

            //reportViewer1 = reportForm.reportViewer1;
            //  reportForm.ShowDialog();

            reportViewer1 = reportForm.reportViewer1;

            byte[] bytes = this.reportViewer1.LocalReport.Render(
               "pdf", null, out mimeType, out encoding, out extension,
               out streamids, out warnings);

            FileStream fs = new FileStream(strFileName, FileMode.Create);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
            issaveok = true;


            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_015 + "\r\n" + Showtime);


            //ExportRpt(0); 
        }

        public void SavePDF()
        {
            issaveok = false;

            {

                var saveFileDialog = new SaveFileDialog();
                saveFileDialog.DefaultExt = ".pdf";
                saveFileDialog.Filter = "PDF(*.pdf)|*.pdf";
                strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                saveFileDialog.FileName = strFileName;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    strFileName = saveFileDialog.FileName.ToString();
                }
                else
                {
                    return;
                }
                //  pdfExport(strFileName);

                try
                {
                    InitialBackGroundWorker();
                    bgWorker.DoWork += new DoWorkEventHandler(pdfExport);
                    bgWorker.RunWorkerAsync();
                    // 启动消息显示画面
                    frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                        clsShowMessage.MSG_007,
                                                        clsConstant.Dialog_Status_Disable);
                    frmMessageShow.ShowDialog();
                    // 数据读取成功后在画面显示
                    if (blnBackGroundWorkIsOK)
                    {
                        string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results"), "");
                        System.Diagnostics.Process.Start("explorer.exe", strFileName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ex" + ex);
                    return;
                    throw ex;
                }

            }
            //if (issaveok == true)
            //    MessageBox.Show("报表已经成功导出到桌面！", "Info");
            //else
            //    MessageBox.Show("报表导出失败，请查找原因！", "Info");
        }


        private void Create_table(bool ishow)
        {
            PDFcheckdataDetail = new List<PDF_checkdataDetail>();
            DataTable qtyTable = PDFcheckdaadetail_Method();

            #region check show page
            int allpage_count;
            bool have_yushu;
            get_pagecount_andYushu(out allpage_count, out have_yushu);


            if (allpage_count >= 1)
                OnlineShow_datas[0].showimage2 = true;
            if (allpage_count >= 2)
                OnlineShow_datas[0].showimage5 = true;
            if (allpage_count >= 3)
                OnlineShow_datas[0].showimage6 = true;
            if (allpage_count >= 4)
                OnlineShow_datas[0].showimage7 = true;
            if (allpage_count >= 5)
                OnlineShow_datas[0].showimage8 = true;
            if (allpage_count >= 6)
                OnlineShow_datas[0].showimage9 = true;
            if (allpage_count >= 7)
                OnlineShow_datas[0].showimage10 = true;
            if (allpage_count >= 8)
                OnlineShow_datas[0].showimage11 = true;
            if (allpage_count >= 9)
                OnlineShow_datas[0].showimage12 = true;
            if (allpage_count >= 10)
                OnlineShow_datas[0].showimage13 = true;
            if (allpage_count >= 11)
                OnlineShow_datas[0].showimage14 = true;
            if (allpage_count >= 12)
                OnlineShow_datas[0].showimage15 = true;
            if (have_yushu == true)
                OnlineShow_datas[0].showimage3 = true;
            #endregion

            reportForm = new ReportForm();
            reportForm.InitializeDataSource(qtyTable, OnlineShow_datas, PDFcheckdataDetail, PDF_Types);
            if (ishow == true)
                reportForm.ShowDialog();

            //  reportForm = null;
        }

        private DataTable PDFcheckdaadetail_Method()
        {
            DataTable qtyTable = new DataTable();
            foreach (Types item in PDF_Types)
            {
                List<ChildType> Child = item.childType;
                qtyTable = new DataTable();
                qtyTable.Columns.Add("name1", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name2", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name3", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name4", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name5", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name6", System.Type.GetType("System.String"));

                foreach (ChildType k in Child)
                {
                    qtyTable.Rows.Add(qtyTable.NewRow());
                }
                int row = 0;
                foreach (ChildType k in Child)
                {
                    if (k.typeName == null || k.typeName == "")
                        qtyTable.Rows[row][0] = "-";
                    else
                        qtyTable.Rows[row][0] = k.typeName;

                    if (k.typeEnName == null || k.typeEnName == "")
                        qtyTable.Rows[row][0] = "-";
                    else
                        qtyTable.Rows[row][0] = k.typeName + "\r\n" + k.typeEnName;

                    if (k.unit == null || k.unit == "")
                        qtyTable.Rows[row][1] = "-";
                    else
                        qtyTable.Rows[row][1] = k.unit;

                    List<MaGait> FindOnline_MaGait = Online_MaGait.FindAll(o => o.type != null && o.type == k.typeCode);

                    if (FindOnline_MaGait.Count == 1)
                    {
                        MaGait ik = FindOnline_MaGait[0];

                        if (ik.avgNormal == null || ik.avgNormal == "")
                            qtyTable.Rows[row][2] = "-";
                        else
                            qtyTable.Rows[row][2] = ik.avgNormal;

                        if (ik.rightNormal == null || ik.rightNormal == "")
                            qtyTable.Rows[row][3] = "-";
                        else
                            qtyTable.Rows[row][3] = ik.rightNormal;

                        qtyTable.Rows[row][4] = "/";

                        if (ik.leftNormal == null || ik.leftNormal == "")
                            qtyTable.Rows[row][5] = "-";
                        else
                            qtyTable.Rows[row][5] = ik.leftNormal;
                    }
                    PDF_checkdataDetail tempnote = new PDF_checkdataDetail();
                    tempnote.name1 = qtyTable.Rows[row][0].ToString();
                    tempnote.name2 = qtyTable.Rows[row][1].ToString();
                    tempnote.name3 = qtyTable.Rows[row][2].ToString();
                    tempnote.name4 = qtyTable.Rows[row][3].ToString();
                    tempnote.name5 = qtyTable.Rows[row][4].ToString();
                    tempnote.name6 = qtyTable.Rows[row][5].ToString();
                    tempnote.typeCode = item.typeCode;
                    tempnote.typeName = item.typeName;
                    tempnote.typeEnName = item.typeEnName;
                    PDFcheckdataDetail.Add(tempnote);
                    row++;
                }
            }
            return qtyTable;
        }

        private void get_pagecount_andYushu(out int allpage_count, out bool have_yushu)
        {
            int i = 0;
            allpage_count = 0;
            have_yushu = false;

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

    }
}
