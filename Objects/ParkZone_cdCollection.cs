using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class ParkZone_cdCollection : CollectionBase
    {
        // Constructor
        public ParkZone_cdCollection()
        {

        }

        public ParkZone_cd this[int index]
        {
            get { return (ParkZone_cd)InnerList[index]; }
        }

        // Add
        public void Add(ParkZone_cd parkZone_cd)
        {
            InnerList.Add(parkZone_cd);
        }

        // Remove
        public void Remove(ParkZone_cd parkZone_cd)
        {
            InnerList.Remove(parkZone_cd);
        }

        // Get zone by it's zoneID
        public ParkZone_cd GetPayType(string parkZone_cdID)
        {
            foreach (ParkZone_cd parkZone_cd in InnerList)
            {
                if (parkZone_cd.ID == parkZone_cdID)
                {
                    return parkZone_cd;
                }
            }
            return null;
        }
    }
}


