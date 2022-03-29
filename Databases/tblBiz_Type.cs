using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolReport.Objects;

namespace ToolReport.Databases
{
    public class tblBiz_Type
    {
        public static string TBL_NAME = "tblBiz_Type";
        public static string TBL_COL_ID = "id";
        public static string TBL_COL_NAME = "name";
        public static string TBL_COL_CODE = "code";
        public static string TBL_COL_DESCRIPTION = "description";

        public static Biz_TypeCollection LoadBiz_Type(Biz_TypeCollection biz_TypeCollection)
        {
            DataTable dtbiz_TypeCollection = StaticPool.mdb.FillData($@"Select {TBL_COL_ID},{TBL_COL_CODE},{TBL_COL_DESCRIPTION},{TBL_COL_NAME}
                                                        from {TBL_NAME} ");
            biz_TypeCollection.Clear();
            if (dtbiz_TypeCollection != null && dtbiz_TypeCollection.Rows.Count > 0)
            {
                foreach (DataRow row in dtbiz_TypeCollection.Rows)
                {
                    Biz_Type biz_Type = new Biz_Type();
                    biz_Type.ID = row[TBL_COL_ID].ToString();
                    biz_Type.Code = row[TBL_COL_CODE].ToString();
                    biz_Type.Name = row[TBL_COL_NAME].ToString();
                    biz_Type.Description = row[TBL_COL_DESCRIPTION].ToString();

                    biz_TypeCollection.Add(biz_Type);
                }
                dtbiz_TypeCollection.Dispose();
                return biz_TypeCollection;

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
