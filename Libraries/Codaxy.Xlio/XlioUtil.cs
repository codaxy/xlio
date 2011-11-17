using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class XlioUtil
    {
        static readonly DateTime baseDate = new DateTime(1899, 12, 31);

        public static DateTime ToDateTime(double excelDateTime)
        {
            if (excelDateTime > 59) excelDateTime -= 1; //Excel/Lotus 2/29/1900 bug   
            return baseDate.AddDays(excelDateTime);
        }

        public static double ToExcelDateTime(DateTime dateTime)
        {
            var res = (dateTime - baseDate).TotalDays;
            if (res > 58)
                res += 1;
            return res;
        }
    }
}
