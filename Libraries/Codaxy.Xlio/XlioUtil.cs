using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

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

        /// <summary>
        /// Returns the Excel name of the zero indexed column number 0->A, 1->B
        /// </summary>
        /// <param name="column">Zero indexed column number.</param>
        /// <returns></returns>
        public static string GetColumnName(int col)
        {
            Stack<char> stack = new Stack<char>();
            do
            {
                stack.Push((Char)('A' + (col % 26)));
                col = col / 26 - 1;
            } while (col >= 0);
            return new String(stack.ToArray());
        }
    }
}
