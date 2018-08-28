using System;
using System.Collections.Generic;
using System.Linq;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace SZ_PDFJsonPrint
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            #region Noway
            DateTime oldDate = DateTime.Now;
            DateTime dt3;
            string endday = DateTime.Now.ToString("yyyy/MM/dd");
            dt3 = Convert.ToDateTime(endday);
            DateTime dt2;
            dt2 = Convert.ToDateTime("2018/09/10");

            TimeSpan ts = dt2 - dt3;
            int timeTotal = ts.Days;
            if (timeTotal < 0)
            {
                MessageBox.Show("Error 23082:");
                Application.Exit();

               return;
            }
            #endregion
            Application.Run(new frmLogin());
        }
    }
}
