using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class DailySumaryCollection : CollectionBase
    {
        // Constructor
        public DailySumaryCollection()
        {

        }

        public DailySumary this[int index]
        {
            get { return (DailySumary)InnerList[index]; }
        }

        // Add
        public void Add(DailySumary dailySumary)
        {
            InnerList.Add(dailySumary);
        }

        // Remove
        public void Remove(DailySumary dailySumary)
        {
            InnerList.Remove(dailySumary);
        }

        // Get zone by it's zoneID
        public DailySumary GetPayType(string dailySumaryID)
        {
            foreach (DailySumary dailySumary in InnerList)
            {
                if (dailySumary.Id == dailySumaryID)
                {
                    return dailySumary;
                }
            }
            return null;
        }
    }
}


