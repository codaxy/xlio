using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Samples.Usage
{
    public class Formulas
    {
        public static void Demo()
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.AddSheet("Sheet1");

            sheet["A1"].Value = "Fibonacci Numbers";
            sheet["A2"].Value = 1;
            sheet["A3"].Value = 1;
            for (var r = 3; r < 100; r++)
                sheet[r, 0].Formula = String.Format("={0}+{1}", Cell.Format(r - 1, 0), Cell.Format(r - 2, 0));

            //instead of setting formula for each cell it is easier to apply shared formula to a cell range

            sheet["B2"].Value = 1;
            sheet["B3"].Value = 1;
            sheet["B4", "B100"].SetFormula("=B2+B3");

            sheet["D1"].Value = "Roots";

            var roots = sheet["D2", "F100"];
            roots.GetColumns(1, 3).SetStyle(new CellStyle { Format = "0.000" });            
            
            foreach (var cdp in roots.GetColumn(0))
                cdp.Data.Value = cdp.Cell.Row + 1;

            roots.GetColumn(1).SetFormula("=D2^(1/2)");
            roots.GetColumn(2).SetFormula("=D2^(1/3)");
            roots.GetColumn(3).SetFormula("=D2^(1/4)");
            workbook.Save("Formulas");
        }
    }
}
