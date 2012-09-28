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
            //Excel/Lotus 2/29/1900 bug   
            if (excelDateTime > 59)
                excelDateTime -= 1;
            return baseDate.AddDays(excelDateTime);
        }

        public static double ToExcelDateTime(DateTime dateTime)
        {
            //Excel does not support milliseconds and causes rounding problems
            var trimmedMilliseconds = new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour, dateTime.Minute, dateTime.Second);
            var res = (trimmedMilliseconds - baseDate).TotalDays;
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

        public static Color GetDefaultIndexedColor(uint indexed)
        {
            switch (indexed)
            {
                case 00: return new Color(0x00000000);
                case 01: return new Color(0x00ffffff);
                case 02: return new Color(0x00ff0000);
                case 03: return new Color(0x0000ff00);
                case 04: return new Color(0x000000ff);
                case 05: return new Color(0x00ffff00);
                case 06: return new Color(0x00ff00ff);
                case 07: return new Color(0x0000ffff);
                case 08: return new Color(0x00000000);
                case 09: return new Color(0x00ffffff);
                case 10: return new Color(0x00ff0000);
                case 11: return new Color(0x0000ff00);
                case 12: return new Color(0x000000ff);
                case 13: return new Color(0x00ffff00);
                case 14: return new Color(0x00ff00ff);
                case 15: return new Color(0x0000ffff);
                case 16: return new Color(0x00800000);
                case 17: return new Color(0x00008000);
                case 18: return new Color(0x00000080);
                case 19: return new Color(0x00808000);
                case 20: return new Color(0x00800080);
                case 21: return new Color(0x00008080);
                case 22: return new Color(0x00C0C0C0);
                case 23: return new Color(0x00808080);
                case 24: return new Color(0x009999FF);
                case 25: return new Color(0x00993366);
                case 26: return new Color(0x00FFFFCC);
                case 27: return new Color(0x00CCFFFF);
                case 28: return new Color(0x00660066);
                case 29: return new Color(0x00FF8080);
                case 30: return new Color(0x000066CC);
                case 31: return new Color(0x00CCCCFF);
                case 32: return new Color(0x00000080);
                case 33: return new Color(0x00FF00FF);
                case 34: return new Color(0x00FFFF00);
                case 35: return new Color(0x0000FFFF);
                case 36: return new Color(0x00800080);
                case 37: return new Color(0x00800000);
                case 38: return new Color(0x00008080);
                case 39: return new Color(0x000000FF);
                case 40: return new Color(0x0000CCFF);
                case 41: return new Color(0x00CCFFFF);
                case 42: return new Color(0x00CCFFCC);
                case 43: return new Color(0x00FFFF99);
                case 44: return new Color(0x0099CCFF);
                case 45: return new Color(0x00FF99CC);
                case 46: return new Color(0x00CC99FF);
                case 47: return new Color(0x00FFCC99);
                case 48: return new Color(0x003366FF);
                case 49: return new Color(0x0033CCCC);
                case 50: return new Color(0x0099CC00);
                case 51: return new Color(0x00FFCC00);
                case 52: return new Color(0x00FF9900);
                case 53: return new Color(0x00FF6600);
                case 54: return new Color(0x00666699);
                case 55: return new Color(0x00969696);
                case 56: return new Color(0x00003366);
                case 57: return new Color(0x00339966);
                case 58: return new Color(0x00003300);
                case 59: return new Color(0x00333300);
                case 60: return new Color(0x00993300);
                case 61: return new Color(0x00993366);
                case 62: return new Color(0x00333399);
                case 63: return new Color(0x00333333);
                case 64: return new Color(0x00000000);
                case 65: return new Color(0x00FFFFFF);
                default : return null;
            }
        }

        public static Color GetDefaultThemeColor(uint index)
        {
            switch (index)
            {                
                case 00: return new Color(0x00ffffff);
                case 01: return new Color(0x00000000);                
                case 02: return new Color(0x00EEECE1);
                case 03: return new Color(0x001F497D);
                case 04: return new Color(0x004F81BD);
                case 05: return new Color(0x00C0504D);
                case 06: return new Color(0x009BBB59);
                case 07: return new Color(0x008064A2);
                case 08: return new Color(0x004BACC6);
                case 09: return new Color(0x00F79646);
                case 10: return new Color(0x000000FF);
                case 11: return new Color(0x00800080);
                default:
                    return null;
            }
        }
    }
}