﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToolReport.Objects;

namespace ToolReport.Databases
{
    public class tblVehicleGroup
    {
        public static string TBL_NAME = "tblVehicleGroup";
        public static string TBL_COL_ID = "VehicleGroupID";
        public static string TBL_COL_NAME = "VehicleGroupName";
        public static string TBL_COL_CODE = "VehicleGroupID";
        public static string TBL_COL_DESCRIPTION = "VehicleType";

        public static ParkZone_cdCollection LoadParkZone_cd(ParkZone_cdCollection parkZone_cdCollection)
        {
            DataTable dtparkZone_cdCollection = StaticPool.mdb.FillData($@"Select {TBL_COL_ID},{TBL_COL_CODE},{TBL_COL_DESCRIPTION},{TBL_COL_NAME}
                                                        from {TBL_NAME} ");
            parkZone_cdCollection.Clear();
            if (dtparkZone_cdCollection != null && dtparkZone_cdCollection.Rows.Count > 0)
            {
                foreach (DataRow row in dtparkZone_cdCollection.Rows)
                {
                    ParkZone_cd parkZone_cd = new ParkZone_cd();
                    parkZone_cd.ID = row[TBL_COL_ID].ToString();
                    parkZone_cd.Code = row[TBL_COL_CODE].ToString();
                    parkZone_cd.Name = row[TBL_COL_NAME].ToString();
                    parkZone_cd.Description = row[TBL_COL_DESCRIPTION].ToString();

                    parkZone_cdCollection.Add(parkZone_cd);
                }
                dtparkZone_cdCollection.Dispose();
                return parkZone_cdCollection;

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
