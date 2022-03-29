using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class Discount
    {
        private string id;
        private string description;
        private string code;
        private string name;
        private string percentDiscount;
        private int amountReduced;

        public string Id { get => id; set => id = value; }
        public string Description { get => description; set => description = value; }
        public string Code { get => code; set => code = value; }
        public string Name { get => name; set => name = value; }
        public int AmountReduced { get => amountReduced; set => amountReduced = value; }
        public string PercentDiscount { get => percentDiscount; set => percentDiscount = value; }
    }
}
