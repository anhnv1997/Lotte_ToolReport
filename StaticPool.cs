using iParking.Database;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ToolReport.Objects;

namespace ToolReport
{
    public static class StaticPool
    {
        // CSDL
        public static MDB mdb = null; // new MDB(Application.StartupPath + "\\iParkingPGS.mdb", "17032008");
        public static MDB mdbEvent = null; // new MDB(Application.StartupPath + "\\iParkingPGS.mdb", "17032008");

        // Thong tin CSDL
        public static string SQLServerName  = Properties.Settings.Default.SQLServerName;
        public static string SQLDatabaseName = Properties.Settings.Default.SQLDatabaseName;
        public static string SQLDatabaseNameEvent = Properties.Settings.Default.SQLDatabaseNameEvent;
        public static string SQLAuthentication = Properties.Settings.Default.SQLAuthentication;
        public static string SQLUserName = Properties.Settings.Default.SQLUserName;
        public static string SQLPassword = Properties.Settings.Default.SQLPassword;

        public static DiscountCollection discountCollection = new DiscountCollection();
        public static PayTypeCollection payTypeCollection = new PayTypeCollection();
        public static ParkMaster_cdCollection parkMaster_cdCollection = new ParkMaster_cdCollection();
        public static ParkZone_cdCollection parkZone_cdCollection = new ParkZone_cdCollection();
        public static Biz_TypeCollection biz_TypeCollection = new Biz_TypeCollection();
        public static DailySumaryCollection dailySumaryCollection = new DailySumaryCollection();


        public static StringCollection receiveMails = Properties.Settings.Default.ListMailReceiveReport;

        public static string reportFileName = "";
        public static string Id(DataGridView dgv)
        {
            string _lcma = "";
            DataGridViewRow _drv = dgv.CurrentRow;
            try
            {
                _lcma = _drv.Cells["_ID"].Value.ToString();
            }
            catch
            {
                _lcma = "";
            }
            return _lcma;
        }

        public static void Logger_Error(string s)
        {
            try
            {
                string pathFile = Path.GetDirectoryName(Application.ExecutablePath) + @"\logs\ERROR_LOG_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";
                using (StreamWriter writer = new StreamWriter(pathFile, true))
                {
                    try
                    {
                        StackFrame callStack = new StackFrame(1, true);
                        writer.WriteLine("ERROR " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "-" + callStack.GetMethod().Name + " - " + s);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
                //MessageBox.Show(ex.Message);
            }
        }
        public static void Logger_Info(string s)
        {
            try
            {
                string pathFile = Path.GetDirectoryName(Application.ExecutablePath) + @"\logs\INFO_LOG_" + DateTime.Now.ToString("dd_MM_yyyy") + ".txt";
                using (StreamWriter writer = new StreamWriter(pathFile, true))
                {
                    try
                    {
                        writer.WriteLine("INFO " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " - " + s);
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Logger_Error(ex.Message);
            }
        }
    }
}
