using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio;
using Codaxy.Xlio.IO;

namespace Codaxy.Xlio.Samples
{
    class Demo1
    {
        public static void Run()
        {
            var workbook = Workbook.Load(@"C:\Work\Test.xlsx");
            var sheet = workbook.Sheets[0];
            sheet["A1"].Value = "Hello";
            sheet["A1"].Style.Fill = new CellFill
            {
                Foreground = new Color(255, 255, 0, 0),
                Pattern = FillPattern.Solid
            };
            sheet[0, 1].Value = "World";
            sheet["A3"].Value = "This is a merged cell with border.";
            sheet["A3"].Style.Alignment.WrapText = true;
            sheet["A3", "B5"].Merge();
            sheet["A3", "B5"].SetOutsideBorder(new BorderEdge
            {
                Style = BorderStyle.Thick,
                Color = Color.Black
            });

            sheet["A6"].Value = "Formatted:";
            var b6 = sheet["B6"];
            b6.Value = 5;
            b6.Style.Format = "#,#.00";

            for (var c = 0; c < 5; c++)
            {
                var bgColor = new Color((byte)(100 + c * 30), 255, (byte)(100 + c * 30));
                for (var r = 0; r < 50; r++)
                    sheet[r, c].Style.Fill = CellFill.Solid(bgColor);
            }

            workbook.Save(@"C:\Work\Xlio.xlsx", XlsxFileWriterOptions.AutoFit);
        }
    }
}
