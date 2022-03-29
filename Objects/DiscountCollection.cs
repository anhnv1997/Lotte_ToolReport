using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class DiscountCollection : CollectionBase
    {
        // Constructor
        public DiscountCollection()
        {

        }

        public Discount this[int index]
        {
            get { return (Discount)InnerList[index]; }
        }

        // Add
        public void Add(Discount discount)
        {
            InnerList.Add(discount);
        }

        // Remove
        public void Remove(Discount discount)
        {
            InnerList.Remove(discount);
        }

        // Get zone by it's zoneID
        public Discount GetDiscount(string discountID)
        {
            foreach (Discount discount in InnerList)
            {
                if (discount.Id == discountID)
                {
                    return discount;
                }
            }
            return null;
        }
    }
}


