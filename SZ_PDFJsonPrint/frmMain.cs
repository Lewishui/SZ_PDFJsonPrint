using Microsoft.Reporting.WinForms;
using SPJP.Buiness;
using SPJP.Common;
using SPJP.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using WeifenLuo.WinFormsUI.Docking;

namespace SZ_PDFJsonPrint
{
    public partial class frmMain : DockContent
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

        public frmMain()
        {
            InitializeComponent();
            FilterOrderResults = new List<clsOrderDatabaseinfo>();

            //clsOrderDatabaseinfo item = new clsOrderDatabaseinfo();
            //item.patientId = "2323";
            //item.dizhi = "jintian";

            //FilterOrderResults.Add(item);
            InitializeDataSource();

            //
            reportForm = new ReportForm();
            //reportForm.InitializeDataSource(null, OnlineShow_datas, PDFcheckdataDetail, PDF_Types);

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
            //string start_time = clsCommHelp.objToDateTime1(dateTimePicker1.Text);
            //string end_time = clsCommHelp.objToDateTime1(dateTimePicker2.Text);

            //OrderResults = BusinessHelp.findOrder_Server(this.keywordTextBox.Text, start_time, end_time);

            //if (Result)
            this.dataGridView3.AutoGenerateColumns = false;
            sortablePendingOrderList = new SortableBindingList<clsOrderDatabaseinfo>(FilterOrderResults);
            this.bindingSource1.DataSource = sortablePendingOrderList;
            this.dataGridView3.DataSource = this.bindingSource1;
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

        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //DataGridViewColumn column = dataGridView.Columns[e.ColumnIndex];
            //clsAllnew BusinessHelp = new clsAllnew();

            //if (column == editColumn1)
            //{
            //    var row = dataGridView.Rows[e.RowIndex];

            //    var model = row.DataBoundItem as clsOrderDatabaseinfo;
            //    //隐藏背景图
            //    //FilterOrderResults[0].showimage = true;

            //    PrintReportForEDI();

            //    return;

            //    Print();
            //}

        }
        private void PrintReportForEDI()
        {

            if (PDF_Types == null)
                return;
            #region maintain
            Create_table(true);
            #endregion
            //reportForm.InitializeDataSource(findsapinfo);
            //InitializeEdiData();
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

        private void filterButton_Click(object sender, EventArgs e)
        {
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

                    #region pdf
                    // this.dataGridView.DataSource = null;
                    //this.dataGridView2.AutoGenerateColumns = false;
                    //this.dataGridView2.DataSource = PDF_Rootdb;
                    this.dataGridView3.AutoGenerateColumns = false;
                    this.dataGridView3.DataSource = PDF_Types;
                    //this.dataGridView4.AutoGenerateColumns = false;
                    this.dataGridView4.DataSource = PDF_ChildType;

                    this.toolStripLabel1.Text = "Count : 0";

                    #endregion


                    #region Excel
                    #region 构造 datatable
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

                    this.dataGridView1.DataSource = null;
                    //this.dataGridView1.AutoGenerateColumns = false;
                    this.dataGridView1.DataSource = Root_datas;


                    this.dataGridView2.DataSource = null;
                    //this.dataGridView2.AutoGenerateColumns = false;
                    this.dataGridView2.DataSource = TBB;


                    //this.dataGridView5.DataSource = null;
                    //this.dataGridView5.AutoGenerateColumns = false;
                    //this.dataGridView5.DataSource = tclass_datas;
                    this.bindingSource7.DataSource = qtyTable;
                    this.dataGridView5.DataSource = this.bindingSource7;

                    #endregion


                    #region Online
                    this.dataGridView8.DataSource = null;
                    this.dataGridView8.AutoGenerateColumns = false;
                    this.dataGridView8.DataSource = Online_Root_datas;


                    this.dataGridView7.DataSource = null;
                    //this.dataGridView7.AutoGenerateColumns = false;
                    this.dataGridView7.DataSource = Online_datas;


                    this.dataGridView6.DataSource = null;
                    //this.dataGridView6.AutoGenerateColumns = false;
                    this.dataGridView6.DataSource = Online_MaGait;

                    #endregion


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化信息失败，请检查数据格式！" + ex, "异常", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
                throw ex;
            }

            clsAllnew BusinessHelp = new clsAllnew();


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

            BusinessHelp.ReadJSON_Report(ref this.bgWorker, "A", "", "", "");

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

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = dataGridView3.Columns[e.ColumnIndex];
            clsAllnew BusinessHelp = new clsAllnew();

            if (column == typeCode)
            {


            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = dataGridView3.Columns[e.ColumnIndex];
            clsAllnew BusinessHelp = new clsAllnew();
            if (e.RowIndex < 0)
                return;

            if (column == typeCode)
            {
                var row = dataGridView3.Rows[e.RowIndex];

                var model = row.DataBoundItem as Types;

                List<ChildType> findsapinfo = PDF_ChildType.FindAll(o => o.partentID != null && model != null && o.partentID == model.typeCode);
                this.dataGridView4.AutoGenerateColumns = false;
                this.dataGridView4.DataSource = findsapinfo;
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            int s = this.tabControl1.SelectedIndex;
            if (s == 1)
            {
                downcsv(dataGridView5);

            }

            return;


            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "Excel Files(*.xls,*.xlsx,*.xlsm,*.xlsb,*.csv)|*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv";
            strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
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

            BusinessHelp.InitializeDataSource(TBB, tclass_datas, Root_datas, Online_datas, Online_MaGait, Online_Root_datas, PDF_Rootdb, PDF_Types, PDF_ChildType);

            BusinessHelp.pbStatus = pbStatus;
            BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            BusinessHelp.DownLoadExcel(ref this.bgWorker, strFileName);

            BusinessHelp.XLSSavesaCSV(strFileName);
            //暂停
            //BusinessHelp.DownLoadPDF(ref this.bgWorker, strFileName);

            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_015 + "\r\n" + Showtime);


        }
        private void downcsv(DataGridView dataGridView)
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
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
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
                strHeader += dataGridView.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < dataGridView.Rows.Count - 1; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView.Columns.Count; k++)
                {
                    if (dataGridView.Rows[j].Cells[k].Value != null)
                    {
                        if (k == 7 || k == 5 || k == 8 || k == 9 || k == 10 || k == 5)
                            strRowValue += "'" + dataGridView.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "'") + delimiter;
                        //if (dataGridView.Rows[j].Cells[k].Value != null)
                        //    strRowValue +=   ((char)(9)).ToString() +dataGridView.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ") + delimiter;
                        //else
                        //    strRowValue +=  ((char)(9)).ToString()+  dataGridView.Rows[j].Cells[k].Value + delimiter;
                        else
                            strRowValue += dataGridView.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "'") + delimiter;


                    }
                    else
                    {
                        strRowValue += dataGridView.Rows[j].Cells[k].Value + delimiter;
                    }
                }

                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成 ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {



            //pdfExport();

            PrintReportForEDI();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
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
            //for (int i = 0; i < tclass_datas.Count; i++)
            //{
            //    printclass_datas = new List<Datas>();

            //    printclass_datas = tclass_datas.FindAll(o => o.serialNumber != null && o.serialNumber == tclass_datas[i].serialNumber);

            clsAllnew BusinessHelp = new clsAllnew();
            #region OLD
            //this.toolStripLabel1.Text = "打印中 1/3";
            //BusinessHelp.Run(printclass_datas);
            //this.toolStripLabel1.Text = "打印中 2/3";

            //BusinessHelp.Run2(printclass_datas);
            //this.toolStripLabel1.Text = "打印中 3/3";

            //BusinessHelp.Run3(printclass_datas);
            #endregion

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

            //}
            MessageBox.Show("打印完成！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pdfExport(string pathl)
        {
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

            reportViewer1 = reportForm.reportViewer1;
           reportForm.ShowDialog();

           reportViewer1 = reportForm.reportViewer1;
        
            byte[] bytes = this.reportViewer1.LocalReport.Render(
               "pdf", null, out mimeType, out encoding, out extension,
               out streamids, out warnings);

            FileStream fs = new FileStream(pathl, FileMode.Create);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
            issaveok = true;


            //ExportRpt(0); 
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            issaveok = false;

            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".pdf";
            saveFileDialog.Filter = "PDF(*.pdf)|*.pdf";
            strFileName = "System  Info" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            pdfExport(strFileName);

            if (issaveok == true)
                MessageBox.Show("报表已经成功导出到桌面！", "Info");
            else
                MessageBox.Show("报表导出失败，请查找原因！", "Info");
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = dataGridView5.Columns[e.ColumnIndex];
            clsAllnew BusinessHelp = new clsAllnew();

            if (column == typeCode)
            {
                var row = dataGridView3.Rows[e.RowIndex];

                var model = row.DataBoundItem as Datas;
                printclass_datas = new List<Datas>();

                printclass_datas = tclass_datas.FindAll(o => o.serialNumber != null && model != null && o.serialNumber == model.serialNumber);

            }
        }
    }
}
