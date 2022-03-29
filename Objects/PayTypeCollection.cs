using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class PayTypeCollection : CollectionBase
    {
        // Constructor
        public PayTypeCollection()
        {

        }

        public PayType this[int index]
        {
            get { return (PayType)InnerList[index]; }
        }

        // Add
        public void Add(PayType payType)
        {
            InnerList.Add(payType);
        }

        // Remove
        public void Remove(PayType payType)
        {
            InnerList.Remove(payType);
        }

        // Get zone by it's zoneID
        public PayType GetPayType(string payTypeID)
        {
            foreach (PayType payType in InnerList)
            {
                if (payType.ID == payTypeID)
                {
                    return payType;
                }
            }
            return null;
        }
    }
}


