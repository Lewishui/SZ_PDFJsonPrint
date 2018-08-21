using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace SZ_PDFJsonPrint
{
    public partial class frmLogin : Form
    {
        public log4net.ILog ProcessLogger;
        public log4net.ILog ExceptionLogger;
        private TextBox txtSAPPassword;
        private CheckBox chkSaveInfo;
        Sunisoft.IrisSkin.SkinEngine se = null;
        frmAboutBox aboutbox;
        private frmMain frmMain;
        //存放要显示的信息
        List<string> messages;
        //要显示信息的下标索引
        int index = 0;
        public frmLogin()
        {
            InitializeComponent();
            aboutbox = new frmAboutBox();
        }

    

        private void btmain_Click(object sender, EventArgs e)
        {

            if (frmMain == null)
            {
                frmMain = new frmMain();
                frmMain.FormClosed += new FormClosedEventHandler(FrmOMS_FormClosed);
            }
            if (frmMain == null)
            {
                frmMain = new frmMain();
            }
            frmMain.Show(this.dockPanel2);
        }

        private void 关于系统ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aboutbox.ShowDialog();
        }
        void FrmOMS_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sender is frmMain)
            {
                frmMain = null;
            }
        }

    }
}
