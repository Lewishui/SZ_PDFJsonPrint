using SPJP.Buiness;
using SPJP.Common;
using SPJP.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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


        //excel
        List<clsExcelinfo> TBB;
        List<Datas> tclass_datas;
        List<Root> Root_datas;



        //客户查找
        List<Online_Data> Online_datas;
        List<MaGait> Online_MaGait;
     List<Online_Root> Online_Root_datas;
        //PDF
        List<PDF_Root> PDF_Rootdb;
        List<Types> PDF_Types;
        List<ChildType> PDF_ChildType;


        public frmMain()
        {
            InitializeComponent();
            FilterOrderResults = new List<clsOrderDatabaseinfo>();

            clsOrderDatabaseinfo item = new clsOrderDatabaseinfo();
            item.patientId = "2323";
            item.dizhi = "jintian";

            FilterOrderResults.Add(item);
            InitializeDataSource();

            //
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
            //string start_time = clsCommHelp.objToDateTime1(dateTimePicker1.Text);
            //string end_time = clsCommHelp.objToDateTime1(dateTimePicker2.Text);

            //OrderResults = BusinessHelp.findOrder_Server(this.keywordTextBox.Text, start_time, end_time);

            //if (Result)
            this.dataGridView.AutoGenerateColumns = false;
            sortablePendingOrderList = new SortableBindingList<clsOrderDatabaseinfo>(FilterOrderResults);
            this.bindingSource1.DataSource = sortablePendingOrderList;
            this.dataGridView.DataSource = this.bindingSource1;
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
            DataGridViewColumn column = dataGridView.Columns[e.ColumnIndex];
            clsAllnew BusinessHelp = new clsAllnew();

            if (column == editColumn1)
            {
                var row = dataGridView.Rows[e.RowIndex];

                var model = row.DataBoundItem as clsOrderDatabaseinfo;
                //隐藏背景图
                //FilterOrderResults[0].showimage = true;

                PrintReportForEDI();

                return;

                BusinessHelp.Run(FilterOrderResults);
                BusinessHelp.Run2(FilterOrderResults);
                BusinessHelp.Run3(FilterOrderResults);
            }
            MessageBox.Show("打印完成！", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void PrintReportForEDI()
        {
            reportForm.InitializeDataSource(FilterOrderResults);
            reportForm.ShowDialog();
            //InitializeEdiData();
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
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.AutoGenerateColumns = false;
                    this.dataGridView1.DataSource = Root_datas;


                    this.dataGridView2.DataSource = null;
                    this.dataGridView2.AutoGenerateColumns = false;
                    this.dataGridView2.DataSource = TBB;


                    this.dataGridView5.DataSource = null;
                    this.dataGridView5.AutoGenerateColumns = false;
                    this.dataGridView5.DataSource = tclass_datas;

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


            clsAllnew BusinessHelp = new clsAllnew();
            //导入程序集
            DateTime oldDate = DateTime.Now;

            BusinessHelp.ReadJSON_Report(ref this.bgWorker, "A");

            PDF_Rootdb = BusinessHelp.PDF_Rootdb;
            PDF_Types = BusinessHelp.PDF_Types;
            PDF_ChildType = BusinessHelp.PDF_ChildType;


            TBB = BusinessHelp.TBB;
            tclass_datas = BusinessHelp.tclass_datas;
            Root_datas = BusinessHelp.Root_datas;


            Online_datas = BusinessHelp.Online_datas;
            Online_MaGait = BusinessHelp.Online_MaGait;
            Online_Root_datas = BusinessHelp.Online_Root_datas;


            DateTime FinishTime = DateTime.Now;
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ": " + s.Seconds.ToString() + "  (分:秒)";
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = dataGridView.Columns[e.ColumnIndex];
            clsAllnew BusinessHelp = new clsAllnew();

            if (column == typeCode)
            {


            }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewColumn column = dataGridView3.Columns[e.ColumnIndex];
            clsAllnew BusinessHelp = new clsAllnew();

            if (column == typeCode)
            {
                var row = dataGridView3.Rows[e.RowIndex];

                var model = row.DataBoundItem as Types;

                List<ChildType> findsapinfo = PDF_ChildType.FindAll(o => o.partentID != null && model != null && o.partentID == model.typeCode);
                this.dataGridView4.AutoGenerateColumns = false;
                this.dataGridView4.DataSource = findsapinfo;
            }
        }
    }
}
