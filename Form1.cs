using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Net.Sockets;
using System.Net.Cache;
using System.IO;
using iParking.Database;
using ToolReport.Databases;
using SpreadsheetLight;
using iParking.Databases;
using ToolReport.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using ToolReport.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using ConnectionConfig;

namespace ToolReport
{
    public partial class Form1 : Form
    {
        SQLConn[] sqls = null;

        int endPMS01 = 0;
        int endParkMaster_cd = 0;
        DateTime currentDay;
        bool isSaveToDB = true;
        bool isReported = true;
        bool isSaveFile = false;

        #region:Form
        public Form1()
        {
            InitializeComponent();
            // The path to the key where Windows looks for startup applications
            RegistryKey rkApp = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (!IsStartupItem())
                // Add the value in the registry so that the application runs at startup
                rkApp.SetValue("ReportManager", Application.ExecutablePath.ToString());
        }
        private bool IsStartupItem()
        {
            // The path to the key where Windows looks for startup applications
            RegistryKey rkApp = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            if (rkApp.GetValue("ReportManager") == null)
                // The value doesn't exist, the application is not set to run at startup
                return false;
            else
                // The value exists, the application is set to run at startup
                return true;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                if (File.Exists(Application.StartupPath + "\\SQLConn.xml"))
                {
                    FileXML.ReadXMLSQLConn(Application.StartupPath + "\\SQLConn.xml", ref sqls);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("frmConnectionConfig: " + ex.Message);
            }
            ConnectToSQLServer();
            currentDay = DateTime.Parse(Properties.Settings.Default.currentDay);
            txtRegTime.Text = "RegTime: " + DateTime.Parse(Properties.Settings.Default.TimeSendReport).ToString("HH:mm:ss");
        }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            bool cursorNotInBar = Screen.GetWorkingArea(this).Contains(Cursor.Position);
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                notifyIcon1.Visible = true;
                this.Hide();
            }
        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon1.Visible = false;
        }
        private void Form1_Shown(object sender, EventArgs e)
        {
            //this.Hide();
            //notifyIcon1.Visible = true;
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            var p = new Process();
            string path = Application.ExecutablePath;
            p.StartInfo.FileName = path;  // just for example, you can use yours.
            p.Start();
        }
        #region: Controls
        private void tsmSetting_Click(object sender, EventArgs e)
        {
            frmSetting frm = new frmSetting();
            frm.ShowDialog();
            if (frm.DialogResult == DialogResult.OK)
            {
                txtRegTime.Text = "RegTime: " + DateTime.Parse(Properties.Settings.Default.TimeSendReport).ToString("HH:mm:ss");
            }
        }
        #endregion
        #endregion

        #region:Event Sumary In Day
        //Get
        public static string GetTimeInDayCondition()
        {
            string startTime = DateTime.Now.AddDays(-1).ToString("yyyyMMdd 00:00:00");
            string endTime = DateTime.Now.AddDays(-1).ToString("yyyyMMdd 23:59:59");
            string timeCondition = $"cast(tblCardEvent.DatetimeIn as datetime) between  '{startTime}' AND '{endTime}'";
            return timeCondition;
        }
        private static DataTable GetEventInDayData()
        {
            string close_DT = $"CLOSE_DT='{DateTime.Now.AddDays(-1).ToString("yyyyMMdd")}'";
            string parkMasterCD = "PARKMASTER_CD = '01'";
            string parkZoneCD = "VehicleGroupID as PARK_ZONE_CD";
            string BizType = "BIZ_Type='010001'";
            string codeDiscountType = "DCTypeCode as DC_TYPE";
            string payType = "PAY_TYPE = 'PK0901'";
            string dcRate = "DiscountPercent as DC_RATE";
            string count = "count(Moneys) as COUNT";
            string orgAmount = "sum(Moneys) as ORG_AMOUNT";
            string dcAmount = "sum( ReducedMoney) as DC_AMOUNT";
            string RegMAN = "REG_MAN ='SYSTEM'";

            string cmd = $@"Select  {close_DT}, {parkMasterCD},{parkZoneCD},{BizType},{codeDiscountType},
                                    {payType},{dcRate},{count},{orgAmount},{dcAmount},{RegMAN}
                            from tblCardEvent
                            where {GetTimeInDayCondition()}                       
                            group by DCTypeCode,VehicleGroupID,DiscountPercent";

            DataTable dtEventData = StaticPool.mdbEvent.FillData(cmd);
            return dtEventData;
        }
        //Save to DB
        private void SaveDailyEventDB(DataTable dtEventData, int CLOSE_DATE, int PARKMASTER_CD, int PARK_ZONE_CD, int BIZ_Type, int DC_TYPE, int PAY_TYPE, int DC_RATE, int COUNT, int ORG_AMOUNT, int DC_AMOUNT, int REG_MAN)
        {
            foreach (DataRow row in dtEventData.Rows)
            {
                DailySumary dailySumary = new DailySumary();
                dailySumary.Id = Guid.NewGuid().ToString();
                dailySumary.PARKMASTER_CD = row[PARKMASTER_CD].ToString();
                dailySumary.PARK_ZONE_CD = row[PARK_ZONE_CD].ToString();
                dailySumary.BIZ_TYPE = row[BIZ_Type].ToString();
                dailySumary.DISCOUNT_TYPE = row[DC_TYPE].ToString();
                dailySumary.PAY_TYPE = row[PAY_TYPE].ToString();

                dailySumary.DISCOUNT_RATE = row[DC_RATE].ToString() != null && row[DC_RATE].ToString() != "" ? Convert.ToInt32(row[DC_RATE].ToString()) : 0;
                dailySumary.COUNT = row[COUNT].ToString() != null && row[COUNT].ToString() != "" ? Convert.ToInt32(row[COUNT].ToString()) : 0;
                dailySumary.ORG_AMOUNT = row[ORG_AMOUNT].ToString() != null && row[ORG_AMOUNT].ToString() != "" ? Convert.ToInt32(row[ORG_AMOUNT].ToString()) : 0;
                dailySumary.DISCOUNT_AMOUNT = row[DC_AMOUNT].ToString() != null && row[DC_AMOUNT].ToString() != "" ? Convert.ToInt32(row[DC_AMOUNT].ToString()) : 0;
                dailySumary.ADJUST_AMOUNT = dailySumary.ORG_AMOUNT - dailySumary.DISCOUNT_AMOUNT;
                dailySumary.VAT_AMOUNT = (int)(dailySumary.ADJUST_AMOUNT / 1.1);
                dailySumary.REG_MAN = row[REG_MAN].ToString();
                string regDTM = DateTime.Now.ToString("yyyy/MM/dd") + " " + DateTime.Parse(Properties.Settings.Default.TimeSendReport).ToString("HH:mm:ss");
                dailySumary.REG_DTM = DateTime.Parse(regDTM);
                string closeDate = row[CLOSE_DATE].ToString().Substring(0, 4) + "/" + row[CLOSE_DATE].ToString().Substring(4, 2) + "/" + row[CLOSE_DATE].ToString().Substring(6, 2) + " 00:00:00";
                dailySumary.CLOSE_DATE = DateTime.Parse(closeDate);
                if (!isSaveToDB)
                {
                    bool result = tblDailySumary.Insert(dailySumary);
                    if (result)
                    {
                        isSaveToDB = true;
                        dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Save to DB", "Success");
                        StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "Save to DB Success");
                    }
                    else
                    {
                        isSaveToDB = false;
                        dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Save to DB", "Error");
                        StaticPool.Logger_Error(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "Save to DB Error");
                    }
                }

            }
        }

        #endregion

        #region: Report Data to Excel File
        private void CreatReportFile(DataTable dtData)
        {
            SLDocument sl = new SLDocument();
            //Create
            CreateReportTittle(sl);
            CreateDiscountTable(sl);
            CreateParkMasterTable(sl);
            CreateParkZoneTable(sl);
            CreateBizTable(sl);
            CreateEventTable(dtData, sl);
            CreatePayType(sl);
            //Save
            SaveReportFile(sl);
        }
        #region: Create
        private void CreateReportTittle(SLDocument sl)
        {
            sl.MergeWorksheetCells("B1", "O1");

            sl.SetCellValue("B1", "BẢNG HIỂN THỊ TRÊN SAP_PMS_01");
            sl.SetCellStyle("B1", ExcelTools.CreateHeader1Style(sl));

            sl.SetCellValue("B" + (endPMS01 + 1), "PMS_02");
            sl.SetCellStyle("B" + (endPMS01 + 1), ExcelTools.CreateHeader2Style(sl));
        }
        private static void CreateEventTable(DataTable dtData, SLDocument sl)
        {
            SetEventTableHeader(sl);
            SetEventTableConten(dtData, sl);
        }
        private void CreateParkMasterTable(SLDocument sl)
        {
            SetParkMasterTableHeader(sl);
            SetParkMasterTableContent(sl);

        }
        private void CreateParkZoneTable(SLDocument sl)
        {
            SetParkZoneTableHeader(sl);
            SetParkZoneTableContent(sl);
        }
        private void CreatePayType(SLDocument sl)
        {
            SetPayTypeTableHeader(sl);
            SetPayTypeTableContent(sl);
        }
        private void CreateBizTable(SLDocument sl)
        {
            SetBizTableHeader(sl);
            SetBizTableContent(sl);
        }
        private void CreateDiscountTable(SLDocument sl)
        {
            SetDiscountTableHeader(sl);
            SetDiscountTableCOntent(sl);
        }
        //Daily Sumary
        private static void SetEventTableHeader(SLDocument sl)
        {
            sl.SetColumnWidth("B", 11);
            sl.SetColumnWidth("C", 17);
            sl.SetColumnWidth("D", 15.5);
            sl.SetColumnWidth("E", 8.9);
            sl.SetColumnWidth("F", 10);
            sl.SetColumnWidth("G", 9);
            sl.SetColumnWidth("H", 15);
            sl.SetColumnWidth("I", 6);
            sl.SetColumnWidth("J", 12);
            sl.SetColumnWidth("K", 18);
            sl.SetColumnWidth("L", 15);
            sl.SetColumnWidth("M", 12);
            sl.SetColumnWidth("N", 14);
            sl.SetColumnWidth("O", 21);

            sl.SetRowHeight(2, 57);

            SLStyle wrapStyle = sl.CreateStyle();
            wrapStyle.Font.FontSize = 9;
            wrapStyle.SetHorizontalAlignment(DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
            wrapStyle.SetVerticalAlignment(DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center);

            wrapStyle.SetWrapText(true);

            sl.SetCellStyle("B2", "O2", ExcelTools.CreateTableHeaderStyle(sl));
            sl.SetCellStyle("B2", "O2", ExcelTools.CreateAllBorderStyle(sl));
            sl.SetCellStyle("B2", "O2", wrapStyle);

            sl.SetCellValue("B2", "CLOSE_DT\n(CLOSEDATE)");
            sl.SetCellValue("C2", "PARKMASTER_CD\n(PARKMASTER CODE)");
            sl.SetCellValue("D2", "PARK_ZONE_CD\n(PARK_ZONE_CODE)");
            sl.SetCellValue("E2", "BIZ_Type\nBUSINESS");
            sl.SetCellValue("F2", "DC_TYPE\n(DISCOUNT)");
            sl.SetCellValue("G2", "PAY_TYPE\n(PAYMANT)");
            sl.SetCellValue("H2", "DC_RATE\n(DISCOUNT_RATE)");
            sl.SetCellValue("I2", "COUNT");
            sl.SetCellValue("J2", "ORG_AMOUNT\n(ORIGINAL)");
            sl.SetCellValue("K2", "DC_AMOUNT\n(DISCOUNT_AMOUNT)");
            sl.SetCellValue("L2", "AD_AMOUNT\n(ADJUST_AMOUNT)");
            sl.SetCellValue("M2", "VAT_AMOUNT");
            sl.SetCellValue("N2", "REG_MAN\n(REGISTER_MAN)");
            sl.SetCellValue("O2", "REG_DTM\nREGUSTER_DATE,TIME\nJobSchedule\n(03:40:00)");

            SLStyle style1 = sl.CreateStyle();
            style1.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Orange, System.Drawing.Color.Orange);
            sl.SetCellStyle("F2", style1);

            SLStyle style2 = sl.CreateStyle();
            style2.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(178, 255, 102), System.Drawing.Color.FromArgb(178, 255, 102));
            sl.SetCellStyle("G2", style2);

            SLStyle style3 = sl.CreateStyle();
            style3.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 153, 76), System.Drawing.Color.FromArgb(0, 153, 76));
            sl.SetCellStyle("C2", style3);

            SLStyle style4 = sl.CreateStyle();
            style4.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(102, 178, 255), System.Drawing.Color.FromArgb(102, 178, 255));
            sl.SetCellStyle("D2", style4);

            SLStyle style5 = sl.CreateStyle();
            style5.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.FromArgb(255, 160, 122), System.Drawing.Color.FromArgb(255, 160, 122));
            sl.SetCellStyle("E2", style5);
        }
        private static void SetEventTableConten(DataTable dtData, SLDocument sl)
        {
            for (int i = 3; i < dtData.Rows.Count + 3; i++)
            {
                SLStyle alignCenterStyle = ExcelTools.CreateAlignCenterStyle(sl);
                sl.SetCellStyle("B" + i, "O" + i, alignCenterStyle);
                sl.SetCellStyle("B" + i, "O" + i, ExcelTools.CreateAllBorderStyle(sl));
                DataRow row = dtData.Rows[i - 3];

                sl.SetCellStyle("F" + i, ExcelTools.SetColorStyle(sl, System.Drawing.Color.Orange));

                sl.SetCellStyle("G" + i, ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(178, 255, 102)));

                sl.SetCellStyle("C" + i, ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(0, 153, 76)));

                sl.SetCellStyle("D" + i, ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(102, 178, 255)));

                sl.SetCellStyle("E" + i, ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(255, 160, 122)));

                sl.SetCellValue("B" + i, row["CLOSE_DT"].ToString());
                sl.SetCellValue("C" + i, row["PARKMASTER_CD"].ToString());
                sl.SetCellValue("D" + i, row["PARK_ZONE_CD"].ToString());
                sl.SetCellValue("E" + i, row["BIZ_Type"].ToString());
                sl.SetCellValue("F" + i, row["DC_TYPE"].ToString());
                sl.SetCellValue("G" + i, row["PAY_TYPE"].ToString());
                sl.SetCellValue("H" + i, row["DC_RATE"].ToString());
                sl.SetCellValue("I" + i, row["COUNT"].ToString());

                sl.SetCellValue("J" + i, row["ORG_AMOUNT"].ToString());
                sl.SetCellValue("K" + i, row["DC_AMOUNT"].ToString());

                int orgAmount = Convert.ToInt32(row["ORG_AMOUNT"].ToString());
                int dc_amount = row["DC_AMOUNT"].ToString() != null && row["DC_AMOUNT"].ToString() != "" ? Convert.ToInt32(row["DC_AMOUNT"].ToString()) : 0;


                sl.SetCellValue("L" + i, orgAmount - dc_amount);
                sl.SetCellValue("M" + i, (orgAmount - dc_amount) / 1.1);
                sl.SetCellValue("N" + i, row["REG_MAN"].ToString());
                sl.SetCellValue("O" + i, DateTime.Parse(Properties.Settings.Default.TimeSendReport).ToString("yyyyMMddHHmmss"));
            }
        }
        //Discount
        private void SetDiscountTableHeader(SLDocument sl)
        {
            sl.MergeWorksheetCells("B" + (endPMS01 + 2), "D" + (endPMS01 + 2));
            sl.SetCellValue("B" + (endPMS01 + 2), "BẢNG 1: Discount_cd");
            sl.SetCellStyle("B" + (endPMS01 + 2), ExcelTools.CreateAlignCenterStyle(sl));

            sl.SetCellValue("B" + (endPMS01 + 3), "Code");
            sl.SetCellValue("C" + (endPMS01 + 3), "Value");
            sl.SetCellValue("D" + (endPMS01 + 3), "Description");
            sl.SetCellStyle("B" + (endPMS01 + 3), "D" + (endPMS01 + 3), ExcelTools.CreateAlignCenterStyle(sl));
            sl.SetCellStyle("B" + (endPMS01 + 3), "D" + (endPMS01 + 3), ExcelTools.CreateAllBorderStyle(sl));
        }
        private void SetDiscountTableCOntent(SLDocument sl)
        {
            tblDiscountParking.LoadDisCount(StaticPool.discountCollection);
            if (StaticPool.discountCollection != null)
            {
                if (StaticPool.discountCollection.Count > 0)
                {
                    for (int i = 0; i < StaticPool.discountCollection.Count; i++)
                    {
                        sl.SetCellStyle("B" + (endPMS01 + 3 + i + 1), ExcelTools.SetColorStyle(sl, System.Drawing.Color.Orange));

                        sl.SetCellStyle("B" + (endPMS01 + 3 + i + 1), "D" + (endPMS01 + 3 + i + 1), ExcelTools.CreateAllBorderStyle(sl));
                        sl.SetCellValue("B" + (endPMS01 + 3 + i + 1), StaticPool.discountCollection[i].Code);
                        sl.SetCellValue("C" + (endPMS01 + 3 + i + 1), StaticPool.discountCollection[i].Name);
                        sl.SetCellValue("D" + (endPMS01 + 3 + i + 1), StaticPool.discountCollection[i].Description);
                    }
                }
            }
        }
        //ParkMaster
        private void SetParkMasterTableHeader(SLDocument sl)
        {
            sl.MergeWorksheetCells("F" + (endPMS01 + 2), "H" + (endPMS01 + 2));
            sl.SetCellValue("F" + (endPMS01 + 2), "BẢNG 2: Parkmaster_cd");
            sl.SetCellStyle("F" + (endPMS01 + 2), ExcelTools.CreateAlignCenterStyle(sl));


            sl.SetCellValue("F" + (endPMS01 + 3), "Code");
            sl.SetCellValue("G" + (endPMS01 + 3), "Value");
            sl.SetCellValue("H" + (endPMS01 + 3), "Description");
            sl.SetCellStyle("F" + (endPMS01 + 3), "H" + (endPMS01 + 3), ExcelTools.CreateTableHeaderStyle(sl));
            sl.SetCellStyle("F" + (endPMS01 + 3), "H" + (endPMS01 + 3), ExcelTools.CreateAllBorderStyle(sl));
        }
        private void SetParkMasterTableContent(SLDocument sl)
        {
            tblParkMaster_cd.LoadParkMaster_cd(StaticPool.parkMaster_cdCollection);
            if (StaticPool.parkMaster_cdCollection != null)
            {
                if (StaticPool.parkMaster_cdCollection.Count > 0)
                {
                    for (int i = 0; i < StaticPool.parkMaster_cdCollection.Count; i++)
                    {
                        sl.SetCellStyle("F" + (endPMS01 + 3 + i + 1), ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(0, 153, 76)));

                        sl.SetCellStyle("G" + (endPMS01 + 3 + i + 1), "H" + (endPMS01 + 3 + i + 1), ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(204, 204, 255)));

                        sl.SetCellStyle("F" + (endPMS01 + 3 + i + 1), "H" + (endPMS01 + 3 + i + 1), ExcelTools.CreateAllBorderStyle(sl));
                        sl.SetCellValue("F" + (endPMS01 + 3 + i + 1), StaticPool.parkMaster_cdCollection[i].Code);
                        sl.SetCellValue("G" + (endPMS01 + 3 + i + 1), StaticPool.parkMaster_cdCollection[i].Name);
                        sl.SetCellValue("H" + (endPMS01 + 3 + i + 1), StaticPool.parkMaster_cdCollection[i].Description);
                    }
                    endParkMaster_cd = endPMS01 + 3 + StaticPool.parkMaster_cdCollection.Count + 3;

                }
                else
                {
                    endParkMaster_cd = endPMS01 + 3;
                }
            }
        }
        //Parkzone
        private void SetParkZoneTableHeader(SLDocument sl)
        {
            sl.MergeWorksheetCells("J" + (endPMS01 + 2), "L" + (endPMS01 + 2));
            sl.SetCellValue("J" + (endPMS01 + 2), "BẢNG 3: Park_zone_cd");
            sl.SetCellStyle("J" + (endPMS01 + 2), ExcelTools.CreateAlignCenterStyle(sl));

            sl.SetCellValue("J" + (endPMS01 + 3), "Code");
            sl.SetCellValue("K" + (endPMS01 + 3), "Value");
            sl.SetCellValue("L" + (endPMS01 + 3), "Description");
            sl.SetCellStyle("J" + (endPMS01 + 3), "L" + (endPMS01 + 3), ExcelTools.CreateTableHeaderStyle(sl));
            sl.SetCellStyle("J" + (endPMS01 + 3), "L" + (endPMS01 + 3), ExcelTools.CreateAllBorderStyle(sl));
        }
        private void SetParkZoneTableContent(SLDocument sl)
        {
            tblVehicleGroup.LoadParkZone_cd(StaticPool.parkZone_cdCollection);
            if (StaticPool.parkZone_cdCollection != null)
            {
                if (StaticPool.parkZone_cdCollection.Count > 0)
                {
                    for (int i = 0; i < StaticPool.parkZone_cdCollection.Count; i++)
                    {
                        sl.SetCellStyle("J" + (endPMS01 + 3 + i + 1), ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(102, 178, 255)));

                        sl.SetCellStyle("J" + (endPMS01 + 3 + i + 1), "L" + (endPMS01 + 3 + i + 1), ExcelTools.CreateAllBorderStyle(sl));
                        sl.SetCellValue("J" + (endPMS01 + 3 + i + 1), StaticPool.parkZone_cdCollection[i].Code);
                        sl.SetCellValue("K" + (endPMS01 + 3 + i + 1), StaticPool.parkZone_cdCollection[i].Name);
                        sl.SetCellValue("L" + (endPMS01 + 3 + i + 1), StaticPool.parkZone_cdCollection[i].Description);
                    }
                }
            }
        }
        //Biz
        private void SetBizTableHeader(SLDocument sl)
        {
            sl.MergeWorksheetCells("N" + (endPMS01 + 2), "O" + (endPMS01 + 2));
            sl.SetCellValue("N" + (endPMS01 + 2), "BẢNG 4: Biz_Type");
            sl.SetCellStyle("N" + (endPMS01 + 2), ExcelTools.CreateAlignCenterStyle(sl));


            sl.SetCellValue("N" + (endPMS01 + 3), "Code");
            sl.SetCellValue("O" + (endPMS01 + 3), "Value");
            sl.SetCellStyle("N" + (endPMS01 + 3), "O" + (endPMS01 + 3), ExcelTools.CreateTableHeaderStyle(sl));
            sl.SetCellStyle("N" + (endPMS01 + 3), "O" + (endPMS01 + 3), ExcelTools.CreateAllBorderStyle(sl));
        }
        private void SetBizTableContent(SLDocument sl)
        {
            tblBiz_Type.LoadBiz_Type(StaticPool.biz_TypeCollection);
            if (StaticPool.biz_TypeCollection != null)
            {
                if (StaticPool.biz_TypeCollection.Count > 0)
                {
                    for (int i = 0; i < StaticPool.biz_TypeCollection.Count; i++)
                    {
                        sl.SetCellStyle("N" + (endPMS01 + 3 + i + 1), ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(255, 160, 122)));

                        sl.SetCellStyle("N" + (endPMS01 + 3 + i + 1), "O" + (endPMS01 + 3 + i + 1), ExcelTools.CreateAllBorderStyle(sl));
                        sl.SetCellValue("N" + (endPMS01 + 3 + i + 1), StaticPool.biz_TypeCollection[i].Code);
                        sl.SetCellValue("O" + (endPMS01 + 3 + i + 1), StaticPool.biz_TypeCollection[i].Name);
                    }
                }
            }
        }
        //PayType
        private void SetPayTypeTableHeader(SLDocument sl)
        {
            sl.MergeWorksheetCells("F" + endParkMaster_cd, "H" + endParkMaster_cd);
            sl.SetCellValue("F" + endParkMaster_cd, "BẢNG 5 PayType_cd");
            sl.SetCellStyle("F" + endParkMaster_cd, ExcelTools.CreateAlignCenterStyle(sl));

            sl.SetCellValue("F" + (endParkMaster_cd + 1), "Code");
            sl.SetCellValue("G" + (endParkMaster_cd + 1), "Value");
            sl.SetCellValue("H" + (endParkMaster_cd + 1), "Description");
            sl.SetCellStyle("F" + (endParkMaster_cd + 1), "H" + (endParkMaster_cd + 1), ExcelTools.CreateTableHeaderStyle(sl));
            sl.SetCellStyle("F" + (endParkMaster_cd + 1), "H" + (endParkMaster_cd + 1), ExcelTools.CreateAllBorderStyle(sl));
        }
        private void SetPayTypeTableContent(SLDocument sl)
        {
            tblPayType.LoadPayType(StaticPool.payTypeCollection);
            if (StaticPool.payTypeCollection != null)
            {
                if (StaticPool.payTypeCollection.Count > 0)
                {
                    for (int i = 0; i < StaticPool.payTypeCollection.Count; i++)
                    {
                        sl.SetCellStyle("F" + (endParkMaster_cd + 1 + i + 1), ExcelTools.SetColorStyle(sl, System.Drawing.Color.FromArgb(178, 255, 102)));

                        sl.SetCellStyle("F" + (endParkMaster_cd + 1 + i + 1), "H" + (endParkMaster_cd + 1 + i + 1), ExcelTools.CreateAllBorderStyle(sl));
                        sl.SetCellValue("F" + (endParkMaster_cd + 1 + i + 1), StaticPool.payTypeCollection[i].Code);
                        sl.SetCellValue("G" + (endParkMaster_cd + 1 + i + 1), StaticPool.payTypeCollection[i].Name);
                        sl.SetCellValue("H" + (endParkMaster_cd + 1 + i + 1), StaticPool.payTypeCollection[i].Description);
                    }
                }
            }
        }
        #endregion

        #region: Save
        private void SaveReportFile(SLDocument sl)
        {
            try
            {
                if (!isSaveFile)
                {
                    StaticPool.reportFileName = @$"./Reports/Report{DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")}.xlsx";
                    sl.SaveAs(StaticPool.reportFileName);
                    dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Create Report File: " + StaticPool.reportFileName, "Success");
                    StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + "Create Report File: " + StaticPool.reportFileName + " Success");
                }
                isSaveFile = true;
            }
            catch (Exception ex)
            {
                StaticPool.Logger_Error(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ":" + "Create Report File: " + StaticPool.reportFileName + " error:" + ex.Message);
                dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Save Report File", "Error");
                isSaveFile = false;
            }
        }
        #endregion
        #endregion

        #region:Internal
        private void ConnectToSQLServer()
        {
            //string cbSQLServerName = StaticPool.SQLServerName;
            //string cbSQLDatabaseName = StaticPool.SQLDatabaseName;
            //string cbSQLAuthentication = StaticPool.SQLAuthentication;
            //string txtSQLUserName = StaticPool.SQLUserName;
            //string txtSQLPassword = StaticPool.SQLPassword;
            //string cbSQLDatabaseNameEvent = StaticPool.SQLDatabaseNameEvent;
            if (sqls != null && sqls.Length > 0)
            {
                string cbSQLServerName = sqls[0].SQLServerName;
                string cbSQLDatabaseName = sqls[0].SQLDatabase;
                string cbSQLAuthentication = sqls[0].SQLAuthentication;
                string txtSQLUserName = sqls[0].SQLUserName;
                string txtSQLPassword = CryptorEngine.Decrypt(sqls[0].SQLPassword, true);
                string cbSQLDatabaseNameEvent = sqls[0].SQLDatabaseEx;
                StaticPool.mdb = new MDB(cbSQLServerName, cbSQLDatabaseName, cbSQLAuthentication, txtSQLUserName, txtSQLPassword);
                StaticPool.mdbEvent = new MDB(cbSQLServerName, cbSQLDatabaseNameEvent, cbSQLAuthentication, txtSQLUserName, txtSQLPassword);
            }
            //StaticPool.mdb = new MDB(cbSQLServerName, cbSQLDatabaseName, cbSQLAuthentication, txtSQLUserName, txtSQLPassword);
            //StaticPool.mdbEvent = new MDB(cbSQLServerName, cbSQLDatabaseNameEvent, cbSQLAuthentication, txtSQLUserName, txtSQLPassword);
        }
        private void Client_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show("Error Happening \n" + e.Error.Message, "Error");
                return;
            }
            MessageBox.Show("Send Successfully", "Done");
        }
        #endregion

        private void timerCheckNewDay_Tick(object sender, EventArgs e)
        {
            if (isReported)
            {
                if (DateTime.Now.AddDays(-1).Day == currentDay.Day)
                {
                    StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM//dd HH:mm:ss") + "New Day Update");
                    isReported = false;
                    isSaveToDB = false;
                    isSaveFile = false;
                    Properties.Settings.Default.currentDay = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    currentDay = DateTime.Now;
                    string timeSendReport = DateTime.Parse(Properties.Settings.Default.TimeSendReport).ToString("HH:mm:ss");
                    Properties.Settings.Default.TimeSendReport = DateTime.Now.ToString("yyyy/MM/dd") + " " + timeSendReport;
                    Properties.Settings.Default.Save();
                }
                else if (DateTime.Now.Day != currentDay.Day)
                {
                    StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM//dd HH:mm:ss") + $"Update time from {currentDay.ToString()} To {DateTime.Now.ToString()}");
                    Properties.Settings.Default.currentDay = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    currentDay = DateTime.Now;
                    string timeSendReport = DateTime.Parse(Properties.Settings.Default.TimeSendReport).ToString("HH:mm:ss");
                    Properties.Settings.Default.TimeSendReport = DateTime.Now.ToString("yyyy/MM/dd") + " " + timeSendReport;
                    Properties.Settings.Default.Save();
                }
            }
        }
        private void timerSendReport_Tick(object sender, EventArgs e)
        {
            if (!isReported)
            {
                if (DateTime.Now.Hour == DateTime.Parse(Properties.Settings.Default.TimeSendReport).Hour && DateTime.Now.Minute == DateTime.Parse(Properties.Settings.Default.TimeSendReport).Minute)
                {
                    StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM//dd HH:mm:ss") + "Register Time");
                    ReportEventInDay();
                    SendReportToGmail();
                }
            }
        }
        private void ReportEventInDay()
        {
            DataTable dtEventData = GetEventInDayData();
            if (dtEventData != null)
            {
                if (dtEventData.Rows.Count > 0)
                {
                    int CLOSE_DATE = 0, PARKMASTER_CD = 1, PARK_ZONE_CD = 2, BIZ_Type = 3, DC_TYPE = 4, PAY_TYPE = 5, DC_RATE = 6, COUNT = 7, ORG_AMOUNT = 8, DC_AMOUNT = 9, REG_MAN = 10;
                    SaveDailyEventDB(dtEventData, CLOSE_DATE, PARKMASTER_CD, PARK_ZONE_CD, BIZ_Type, DC_TYPE, PAY_TYPE, DC_RATE, COUNT, ORG_AMOUNT, DC_AMOUNT, REG_MAN);
                }
                else
                {
                    dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Dont have new Event Data");
                    StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "Dont have new Event Data");
                }
                endPMS01 = dtEventData.Rows.Count + 2;
                CreatReportFile(dtEventData);
            }
        }
        private void SendReportToGmail()
        {
            try
            {
                using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587))
                {
                    client.EnableSsl = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    string password = "jmuvtjnuiskqeref";
                    client.Credentials = new NetworkCredential("nguyenvietanh09031997@gmail.com", password);
                    MailMessage msg = new MailMessage();
                    msg.To.Add("anhnv@kztek.net");
                    if (StaticPool.receiveMails != null)
                    {
                        foreach (string mailAddr in StaticPool.receiveMails)
                        {
                            msg.To.Add(mailAddr);
                        }
                    }
                    msg.From = new MailAddress("nguyenvietanh09031997@gmail.com", "ReportManager");
                    msg.Subject = "Report: " + DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd");
                    msg.Body = "";
                    Attachment dinhkem = new Attachment(StaticPool.reportFileName);
                    msg.Attachments.Add(dinhkem);
                    client.SendCompleted += Client_SendCompleted; ;
                    client.Send(msg);
                    isReported = true;
                    dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Send Report to Gmail", "Success");
                    StaticPool.Logger_Info(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "Send Report to Gmail Success");
                }
            }
            catch (Exception ex)
            {
                dgvMessage.Rows.Add(dgvMessage.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "Send Report to Gmail", "Error");
                StaticPool.Logger_Error(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "Send Report to Gmail error:" + ex.Message);
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            txtCurrentTime.Text = "Current Time: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        }
    }
}
