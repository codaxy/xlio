using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    public class NumberFormat
    {
        public const String General = "General";

        internal static bool IsDateTimeFormat(String code)
        {
            if (String.IsNullOrEmpty(code))
                return false;
            return code.IndexOfAny(new[] { 'd', 'm', 'y', 'h', 's' }) != -1;
        }

        internal static bool TryGetPredefinedFormatId(String code, out uint formatId)
        {
            switch (code)
            {
                case "General": formatId = 0; return true;
                case "0": formatId = 1; return true;
                case "0.00": formatId = 2; return true;
                case "#,##0": formatId = 3; return true;
                case "#,##0.00": formatId = 4; return true;
                case "0%": formatId = 9; return true;
                case "0.00%": formatId = 10; return true;
                case "0.00E+00": formatId = 11; return true;
                case "# ?/?": formatId = 12; return true;
                case "# ??/??": formatId = 13; return true;
                case "mm-dd-yy": formatId = 14; return true;
                case "d-mmm-yy": formatId = 15; return true;
                case "d-mmm": formatId = 16; return true;
                case "mmm-yy": formatId = 17; return true;
                case "h:mm AM/PM": formatId = 18; return true;
                case "h:mm:ss AM/PM": formatId = 19; return true;
                case "h:mm": formatId = 20; return true;
                case "h:mm:ss": formatId = 21; return true;
                case "m/d/yy h:mm": formatId = 22; return true;
                case "#,##0 ;(#,##0)": formatId = 37; return true;
                case "#,##0 ;[Red](#,##0)": formatId = 38; return true;
                case "#,##0.00;(#,##0.00)": formatId = 39; return true;
                case "#,##0.00;[Red](#,##0.00)": formatId = 40; return true;
                case "mm:ss": formatId = 45; return true;
                case "[h]:mm:ss": formatId = 46; return true;
                case "mmss.0": formatId = 47; return true;
                case "##0.0E+0": formatId = 48; return true;
                case "@": formatId = 49; return true;
                default:
                    formatId = 0;
                    return false;
            }
        }

        internal static bool TryGetPredefinedFormatCode(uint numFmtId, out string formatCode)
        {
            switch (numFmtId)
            {
                case 0: formatCode = General; return true;
                case 1: formatCode = "0"; return true;
                case 2: formatCode = "0.00"; return true;
                case 3: formatCode = "#,##0"; return true;
                case 4: formatCode = "#,##0.00"; return true;
                case 9: formatCode = "0%"; return true;
                case 10: formatCode = "0.00%"; return true;
                case 11: formatCode = "0.00E+00"; return true;
                case 12: formatCode = "# ?/?"; return true;
                case 13: formatCode = "# ??/??"; return true;
                case 14: formatCode = "mm-dd-yy"; return true;
                case 15: formatCode = "d-mmm-yy"; return true;
                case 16: formatCode = "d-mmm"; return true;
                case 17: formatCode = "mmm-yy"; return true;
                case 18: formatCode = "h:mm AM/PM"; return true;
                case 19: formatCode = "h:mm:ss AM/PM"; return true;
                case 20: formatCode = "h:mm"; return true;
                case 21: formatCode = "h:mm:ss"; return true;
                case 22: formatCode = "m/d/yy h:mm"; return true;
                case 37: formatCode = "#,##0 ;(#,##0)"; return true;
                case 38: formatCode = "#,##0 ;[Red](#,##0)"; return true;
                case 39: formatCode = "#,##0.00;(#,##0.00)"; return true;
                case 40: formatCode = "#,##0.00;[Red](#,##0.00)"; return true;
                case 45: formatCode = "mm:ss"; return true;
                case 46: formatCode = "[h]:mm:ss"; return true;
                case 47: formatCode = "mmss.0"; return true;
                case 48: formatCode = "##0.0E+0"; return true;
                case 49: formatCode = "@"; return true;
                default:
                    formatCode = null;
                    return false;
            }
        }
    }
}
