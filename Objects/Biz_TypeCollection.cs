using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class Biz_TypeCollection : CollectionBase
    {
        // Constructor
        public Biz_TypeCollection()
        {

        }

        public Biz_Type this[int index]
        {
            get { return (Biz_Type)InnerList[index]; }
        }

        // Add
        public void Add(Biz_Type biz_type)
        {
            InnerList.Add(biz_type);
        }

        // Remove
        public void Remove(Biz_Type biz_type)
        {
            InnerList.Remove(biz_type);
        }

        // Get zone by it's zoneID
        public Biz_Type GetPayType(string biz_typeID)
        {
            foreach (Biz_Type biz_type in InnerList)
            {
                if (biz_type.ID == biz_typeID)
                {
                    return biz_type;
                }
            }
            return null;
        }
    }
}


