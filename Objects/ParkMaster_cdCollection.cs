using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class ParkMaster_cdCollection : CollectionBase
    {
        // Constructor
        public ParkMaster_cdCollection()
        {

        }

        public ParkMaster_cd this[int index]
        {
            get { return (ParkMaster_cd)InnerList[index]; }
        }

        // Add
        public void Add(ParkMaster_cd parkMaster_cd)
        {
            InnerList.Add(parkMaster_cd);
        }

        // Remove
        public void Remove(ParkMaster_cd parkMaster_cd)
        {
            InnerList.Remove(parkMaster_cd);
        }

        // Get zone by it's zoneID
        public ParkMaster_cd GetParkMaster_cd(string payTypeID)
        {
            foreach (ParkMaster_cd parkMaster_cd in InnerList)
            {
                if (parkMaster_cd.ID == payTypeID)
                {
                    return parkMaster_cd;
                }
            }
            return null;
        }
    }
}


