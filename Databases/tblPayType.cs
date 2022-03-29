using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolReport.Objects;

namespace ToolReport.Databases
{
    public class tblPayType
    {
        public static string TBL_NAME = "tblPayType";
        public static string TBL_COL_ID = "id";
        public static string TBL_COL_NAME = "name";
        public static string TBL_COL_CODE = "code";
        public static string TBL_COL_DESCRIPTION = "description";

        public static PayTypeCollection LoadPayType(PayTypeCollection payTypeCollection)
        {
            DataTable dtPayType = StaticPool.mdb.FillData($@"Select {TBL_COL_ID},{TBL_COL_CODE},{TBL_COL_DESCRIPTION},{TBL_COL_NAME}
                                                        from {TBL_NAME} ");
            payTypeCollection.Clear();
            if (dtPayType != null && dtPayType.Rows.Count > 0)
            {
                foreach (DataRow row in dtPayType.Rows)
                {
                    PayType paytype = new PayType();
                    paytype.ID = row[TBL_COL_ID].ToString();
                    paytype.Code = row[TBL_COL_CODE].ToString();
                    paytype.Name = row[TBL_COL_NAME].ToString();
                    paytype.Description = row[TBL_COL_DESCRIPTION].ToString();

                    payTypeCollection.Add(paytype);
                }
                dtPayType.Dispose();
                return payTypeCollection;

            }
            return null;
        }
        //Add
        public static bool Insert(string id,string code,string name, string description)
        {
            string insertCMD = $@"Insert into {TBL_NAME}({TBL_COL_ID},{TBL_COL_CODE},{TBL_COL_NAME},{TBL_COL_DESCRIPTION})
                                  values('{id}',N'{code}',N'{name}',N'{description}')";
            
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
        public static bool Modify(string id, string code, string name, string description)
        {
            string modifyCMD = $@"update {TBL_NAME} 
                                  set {TBL_COL_CODE}=N'{code}',
                                      {TBL_COL_DESCRIPTION}=N'{description}',
                                      {TBL_COL_NAME}=N'{name}'
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
        public static bool Delete(string payTypeID)
        {
            string deleteCMD = $"Delete {TBL_NAME} where {TBL_COL_ID}='{payTypeID}'";
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
