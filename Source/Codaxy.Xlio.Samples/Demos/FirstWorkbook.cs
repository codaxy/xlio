using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio;

namespace Codaxy.Xlio.Samples
{
    class FirstWorkbook
    {
        public static void Run()
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.AddSheet("Demo");
            sheet["A1"].Value = "Hello World";
            sheet["A1"].Style.Font.Size = 20;
            workbook.Save(@"FirstWorkbook.xlsx");
        }
    }
}
