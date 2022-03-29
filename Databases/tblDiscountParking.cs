using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolReport.Objects;

namespace ToolReport.Databases
{
    public class tblDiscountParking
    {
        public static string TBL_NAME = "tblDiscountParking";
        public static string TBL_COL_ID = "Id";
        public static string TBL_COL_NAME = "DCTypeName";
        public static string TBL_COL_CODE = "DCTypeCode";
        public static string TBL_COL_PERCENT = "DiscountPercent";
        public static string TBL_COL_AMOUNT = "DiscountAmount";
        public static string TBL_COL_DESCRIPTION = "Note";

        public static DiscountCollection LoadDisCount(DiscountCollection discountCollection)
        {
            DataTable dtDiscount = StaticPool.mdb.FillData($@"Select {TBL_COL_ID},{TBL_COL_CODE},{TBL_COL_DESCRIPTION},{TBL_COL_NAME},{TBL_COL_PERCENT},{TBL_COL_AMOUNT} 
                                                        from {TBL_NAME} ");
            discountCollection.Clear();
            if (dtDiscount != null && dtDiscount.Rows.Count > 0)
            {
                foreach (DataRow row in dtDiscount.Rows)
                {
                    Discount discount = new Discount();
                    discount.Id = row[TBL_COL_ID].ToString();
                    discount.Code = row[TBL_COL_CODE].ToString();
                    discount.Name = row[TBL_COL_NAME].ToString();
                    discount.PercentDiscount = row[TBL_COL_PERCENT].ToString();
                    discount.AmountReduced = Convert.ToInt32(row[TBL_COL_AMOUNT].ToString());
                    discount.Description = row[TBL_COL_DESCRIPTION].ToString();

                    discountCollection.Add(discount);
                }
                dtDiscount.Dispose();
                return discountCollection;

            }
            return null;
        }
        //Add
        public static bool Insert(string id,string code,string name, string percent, int amount, string description)
        {
            string insertCMD = $@"Insert into {TBL_NAME}({TBL_COL_ID},{TBL_COL_CODE},{TBL_COL_NAME},{TBL_COL_PERCENT},{TBL_COL_AMOUNT},{TBL_COL_DESCRIPTION})
                                  values('{id}',N'{code}',N'{name}',N'{percent}',{amount},N'{description}')";
            
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
        public static bool Modify(string id, string code, string name, string percent, int amount, string description)
        {
            string modifyCMD = $@"update {TBL_NAME} 
                                  set {TBL_COL_CODE}=N'{code}',
                                      {TBL_COL_DESCRIPTION}=N'{description}',
                                      {TBL_COL_NAME}=N'{name}',
                                      {TBL_COL_PERCENT}=N'{percent}' ,
                                      {TBL_COL_AMOUNT} = {amount}
                                  Where {TBL_COL_ID} = '{id}'";
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
        public static bool Delete(string discountID)
        {
            string deleteCMD = $"Delete {TBL_NAME} where {TBL_COL_ID}='{discountID}'";
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
