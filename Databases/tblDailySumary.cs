using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolReport;
using ToolReport.Objects;

namespace iParking.Databases
{
    public class tblDailySumary
    {
        public static string TBL_NAME = "tblDailySumary";

        public static string TBL_COL_ID = "id";

        public static string TBL_COL_CLOSE_DT = "CLOSE_DT";
        public static string TBL_COL_PARKMASTER_CD = "PARKMASTER_CD";
        public static string TBL_COL_PARK_ZONE_CD = "PARK_ZONE_CD";
        public static string TBL_COL_BIZ_TYPE = "BIZ_TYPE";
        public static string TBL_COL_PAY_TYPE = "PAY_TYPE";

        public static string TBL_COL_DISCOUNT_TYPE = "DISCOUNT_TYPE";
        public static string TBL_COL_DISCOUNT_RATE = "DISCOUNT_RATE";

        public static string TBL_COL_COUNT = "COUNT";

        public static string TBL_COL_ORG_AMOUNT = "ORG_AMOUNT";
        public static string TBL_COL_DISCOUNT_AMOUNT = "DISCOUNT_AMOUNT";
        public static string TBL_COL_ADJUST_AMOUNT = "ADJUST_AMOUNT";
        public static string TBL_COL_VAT_AMOUNT = "VAT_AMOUNT";

        public static string TBL_COL_REG_MAN = "REG_MAN";
        public static string TBL_COL_REG_DTM = "REG_DTM";

        public static DailySumaryCollection LoadDailySumary(DailySumaryCollection dailySumaryCollection, string startTime, string endTime)
        {
            string GetCMD = $@"Select {TBL_COL_ID},{TBL_COL_PARKMASTER_CD},{TBL_COL_PARK_ZONE_CD},{TBL_COL_BIZ_TYPE}
                                      {TBL_COL_DISCOUNT_TYPE},{TBL_COL_PAY_TYPE},{TBL_COL_DISCOUNT_RATE},{TBL_COL_COUNT},
                                      {TBL_COL_ORG_AMOUNT},{TBL_COL_DISCOUNT_AMOUNT},{TBL_COL_ADJUST_AMOUNT},
                                      {TBL_COL_VAT_AMOUNT},{TBL_COL_REG_MAN},{TBL_COL_REG_DTM},{TBL_COL_CLOSE_DT}
                               from {TBL_NAME} ";
            DataTable dtDailySumary = StaticPool.mdb.FillData(GetCMD);
            dailySumaryCollection.Clear();
            if (dtDailySumary != null && dtDailySumary.Rows.Count > 0)
            {
                foreach (DataRow row in dtDailySumary.Rows)
                {
                    DailySumary dailySumary = new DailySumary();
                    dailySumary.Id = row[TBL_COL_ID].ToString();
                    dailySumary.PARKMASTER_CD = row[TBL_COL_PARKMASTER_CD].ToString();
                    dailySumary.PARK_ZONE_CD = row[TBL_COL_PARK_ZONE_CD].ToString();
                    dailySumary.BIZ_TYPE = row[TBL_COL_BIZ_TYPE].ToString();
                    dailySumary.DISCOUNT_TYPE = row[TBL_COL_DISCOUNT_TYPE].ToString();
                    dailySumary.PAY_TYPE = row[TBL_COL_PAY_TYPE].ToString();
                    dailySumary.DISCOUNT_RATE = Convert.ToInt32(row[TBL_COL_DISCOUNT_RATE].ToString());
                    dailySumary.COUNT = Convert.ToInt32(row[TBL_COL_COUNT].ToString());
                    dailySumary.ORG_AMOUNT = Convert.ToInt32(row[TBL_COL_ORG_AMOUNT].ToString());
                    dailySumary.DISCOUNT_AMOUNT = Convert.ToInt32(row[TBL_COL_DISCOUNT_AMOUNT].ToString());
                    dailySumary.ADJUST_AMOUNT = Convert.ToInt32(row[TBL_COL_ADJUST_AMOUNT].ToString());
                    dailySumary.VAT_AMOUNT = Convert.ToInt32(row[TBL_COL_VAT_AMOUNT].ToString());
                    dailySumary.REG_MAN = row[TBL_COL_REG_MAN].ToString();
                    dailySumary.REG_DTM = DateTime.Parse(row[TBL_COL_REG_DTM].ToString());
                    dailySumary.CLOSE_DATE = DateTime.Parse(row[TBL_COL_CLOSE_DT].ToString());

                    dailySumaryCollection.Add(dailySumary);
                }
                dtDailySumary.Dispose();
                return dailySumaryCollection;

            }
            return null;
        }
        //Add
        public static bool Insert(DailySumary dailySumary)
        {
            string insertCMD = $@"Insert into {TBL_NAME}
                                        ({TBL_COL_ID},{TBL_COL_CLOSE_DT},{TBL_COL_PARKMASTER_CD},{TBL_COL_PARK_ZONE_CD},
                                         {TBL_COL_BIZ_TYPE},{TBL_COL_PAY_TYPE},{TBL_COL_DISCOUNT_TYPE},{TBL_COL_DISCOUNT_RATE},
                                         {TBL_COL_COUNT},{TBL_COL_DISCOUNT_AMOUNT},{TBL_COL_ADJUST_AMOUNT},{TBL_COL_VAT_AMOUNT},
                                         {TBL_COL_REG_MAN},{TBL_COL_REG_DTM})
                                  values('{dailySumary.Id}','{dailySumary.CLOSE_DATE}','{dailySumary.PARKMASTER_CD}','{dailySumary.PARK_ZONE_CD}',
                                         '{dailySumary.BIZ_TYPE}','{dailySumary.PAY_TYPE}','{dailySumary.DISCOUNT_TYPE}',{dailySumary.DISCOUNT_RATE},
                                          {dailySumary.COUNT},{dailySumary.DISCOUNT_AMOUNT},{dailySumary.ADJUST_AMOUNT},{dailySumary.VAT_AMOUNT},
                                         '{dailySumary.REG_MAN}','{dailySumary.REG_DTM}')";

            if (!StaticPool.mdb.ExecuteCommand(insertCMD))
            {
                if (!StaticPool.mdb.ExecuteCommand(insertCMD))
                {
                    return false;
                }
            }
            return true;
        }
        //Modify
        public static bool Modify(DailySumary dailySumary)
        {
            string modifyCMD = $@"update {TBL_NAME} 
                                  set {TBL_COL_CLOSE_DT}='{dailySumary.CLOSE_DATE}',
                                      {TBL_COL_PARKMASTER_CD}='{dailySumary.PARKMASTER_CD}',
                                      {TBL_COL_PARK_ZONE_CD}='{dailySumary.PARK_ZONE_CD}',
                                      {TBL_COL_BIZ_TYPE}='{dailySumary.BIZ_TYPE}',
                                      {TBL_COL_PAY_TYPE}='{dailySumary.PAY_TYPE}',
                                      {TBL_COL_DISCOUNT_TYPE}='{dailySumary.DISCOUNT_TYPE}',
                                      {TBL_COL_DISCOUNT_RATE}={dailySumary.DISCOUNT_RATE},
                                      {TBL_COL_COUNT}={dailySumary.COUNT},
                                      {TBL_COL_ORG_AMOUNT}={dailySumary.ORG_AMOUNT},
                                      {TBL_COL_DISCOUNT_AMOUNT}={dailySumary.DISCOUNT_AMOUNT},
                                      {TBL_COL_ADJUST_AMOUNT}={dailySumary.ADJUST_AMOUNT},
                                      {TBL_COL_VAT_AMOUNT}={dailySumary.VAT_AMOUNT},
                                      {TBL_COL_REG_MAN}='{dailySumary.REG_MAN}',
                                      {TBL_COL_REG_DTM}='{dailySumary.REG_DTM}'
                                  Where {TBL_COL_ID} = '{dailySumary.Id}'";
            if (!StaticPool.mdb.ExecuteCommand(modifyCMD))
            {
                if (!StaticPool.mdb.ExecuteCommand(modifyCMD))
                {
                    return false;
                }
            }
            return true;
        }
        //Delete
        public static bool Delete(string id)
        {
            string deleteCMD = $"Delete {TBL_NAME} where {TBL_COL_ID}='{id}'";
            if (!StaticPool.mdb.ExecuteCommand(deleteCMD))
            {
                if (!StaticPool.mdb.ExecuteCommand(deleteCMD))
                {
                    return false;
                }
            }
            return true;
        }
        public static bool DeleteAll()
        {
            string deleteAllCMD = $"Delete {TBL_NAME}";
            if (!StaticPool.mdb.ExecuteCommand(deleteAllCMD))
            {
                if (!StaticPool.mdb.ExecuteCommand(deleteAllCMD))
                {
                    return false;
                }
            }
            return true;
        }



    }
}
