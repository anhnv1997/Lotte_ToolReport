using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolReport.Objects
{
    public class DailySumary
    {
        public string Name;
        public string Id = "id";
        public string PARKMASTER_CD = "";
        public string PARK_ZONE_CD = "";
        public string BIZ_TYPE = "";
        public string DISCOUNT_TYPE = "";
        public string PAY_TYPE = "";
        public int DISCOUNT_RATE = 0;
        public int COUNT = 0;
        public int ORG_AMOUNT = 0;
        public int DISCOUNT_AMOUNT = 0;
        public int ADJUST_AMOUNT = 0;
        public int VAT_AMOUNT = 0;
        public string REG_MAN = "";
        public DateTime REG_DTM ;
        public DateTime CLOSE_DATE ;
    }
}
