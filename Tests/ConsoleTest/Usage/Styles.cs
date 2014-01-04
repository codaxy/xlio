using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Samples.Usage
{
    class Styles
    {
        public static void Colors()
        {
            Color black = Color.Black; //used often
            Color gray = new Color(0xaaaaaa); //hex rrggbb
            Color yellow = new Color(255, 255, 0); //rgb
            Color red = new Color("#f00"); //hex string
            Color blue = new Color("#0000ff"); //hex string
        }

        public static void Borders()
        {
            var border = new CellBorder();
            border.Left = new BorderEdge { Style = BorderStyle.Double, Color = Color.Black }; //apply double black border to left edge
            border.All = new BorderEdge { Style = BorderStyle.Medium, Color = Color.Black }; //apply medium black border to left, right, top and bottom edge
            border.Diagonal = new BorderEdge { Style = BorderStyle.Dashed, Color = Color.Black }; // diagonal border is supported also
            border.DiagonalUp = true; //in this case you should set which diagonal to use
            border.DiagonalDown = false; //in this case you should set which diagonal to use

            var wb = new Workbook();
            var sheet = wb.Sheets.AddSheet("Demo Sheet");
            sheet["A1"].Style.Border = border; //apply prepared border to cell A1
            sheet["B1"].Style.Border = border; //apply prepared border to cell B1
            sheet["C1"].Style.Border = border; //apply prepared border to cell C1
            sheet["C1"].Style.Border.Bottom = null; //watch out, you have removed bottom border on A1, A2 and A3

            var range = sheet["A3", "E10"];

            range.SetBorder(new BorderEdge { Style = BorderStyle.Thin, Color = Color.Black }); //border is easier to set on ranges 

            range.GetRow(0).SetOutsideBorder(new BorderEdge { Style = BorderStyle.Medium, Color = Color.Black }); //header may need a heavier border

            range.SetOutsideBorder(new BorderEdge { Style = BorderStyle.Double, Color = Color.Black }); //use double border outside the table

            wb.Save("Borders.xlsx");
        }

        public static void Fills()
        {
            var whiteFill = CellFill.Solid(new Color("#fff")); //white fill
            var blackFill = new CellFill { Pattern = FillPattern.Solid, Foreground = Color.Black }; //make sure you set the foreground color

            var wb = new Workbook();
            var sheet = wb.Sheets.AddSheet("Chessboard");

            var cells = sheet[1, 1, 8, 8];
            cells.SetOutsideBorder(new BorderEdge { Color = Color.Black, Style = BorderStyle.Medium });
            
            for (var i = 0; i < cells.Range.Width; i++)
                for (var j = 0; j < cells.Range.Height; j++)
                cells[j, i].Style.Fill = (i + j) % 2 == 0 ? whiteFill : blackFill;

            var rows = cells.GetColumn(-1);
            for (var i = 0; i < rows.Range.Height; i++)
            {
                rows[i].Value = 8 - i;
                rows[i].Style.Alignment.Horizontal = HorizontalAlignment.Center;
                rows[i].Style.Alignment.Vertical = VerticalAlignment.Center;
            }

            var columns = cells.GetRow(8);
            for (var i = 0; i < columns.Range.Width; i++)
            {
                columns[i].Value = new string(new char[] { (Char)('A' + i) });
                columns[i].Style.Alignment.Horizontal = HorizontalAlignment.Center;
                columns[i].Style.Alignment.Vertical = VerticalAlignment.Center;
            }

            sheet.DefaultRowHeight = 40; //px
            wb.Save("Fills.xlsx");
        }

        public static void Fonts()
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.AddSheet("Fonts");

            workbook.DefaultFont.Name = "Arial"; //set default font on workbook level
            workbook.DefaultFont.Size = 10; //set default font on workbook level            

            var r = 0; //current row
            var c = 1; //current column

            ++r; 
            sheet[r, c].Value = "Normal";

            ++r;
            sheet[r, c].Value = "Bold";
            sheet[r, c].Style.Font.Bold = true;

            ++r;
            sheet[r, c].Value = "Italic";
            sheet[r, c].Style.Font.Italic = true;

            ++r;
            sheet[r, c].Value = "Strike";
            sheet[r, c].Style.Font.Strike = true;

            ++r;
            sheet[r, c].Value = "Underline";
            sheet[r, c].Style.Font.Underline = FontUnderline.Single;

            ++r;
            sheet[r, c].Value = "Red";
            sheet[r, c].Style.Font.Color = new Color("#f00");

            ++r;
            sheet[r, c].Value = "20pt";
            sheet[r, c].Style.Font.Size = 20;

            ++r;
            sheet[r, c].Value = "This text will wrap into multiple lines.";
            sheet[r, c].Style.Alignment.WrapText = true;

            ++r;
            sheet[r, c].Value = "This text shrink to fit the cell.";
            sheet[r, c].Style.Alignment.ShrinkToFit = true;

            r = 0;
            c = 3;

            ++r;
            sheet[r, c].Value = "Left";
            sheet[r, c].Style.Alignment.Horizontal = HorizontalAlignment.Left;

            ++r;
            sheet[r, c].Value = "Right";
            sheet[r, c].Style.Alignment.Horizontal = HorizontalAlignment.Right;

            ++r;
            sheet[r, c].Value = "Center";
            sheet[r, c].Style.Alignment.Horizontal = HorizontalAlignment.Center;

            ++r;
            ++r;
            sheet[r, c].Value = "Indent 1";
            sheet[r, c].Style.Alignment.Indent = 1; //3 spaces

            ++r;
            sheet[r, c].Value = "Indent 2";
            sheet[r, c].Style.Alignment.Indent = 2; //6 spaces

            ++r;
            sheet[r, c].Value = "Indent 3";
            sheet[r, c].Style.Alignment.Indent = 3; //9 spaces            

            workbook.Save("Fonts.xlsx");
        }

        public static void NumberFormat()
        {
            var wb = new Workbook();
            var sheet = wb.Sheets.AddSheet("Numbers");

            var r = 0; //current row
            var c = 1; //current column

            ++r;
            sheet[r, c].Value = 1;
            sheet[r, c].Style.Format = null; //default format

            ++r;
            sheet[r, c].Value = 1000;
            sheet[r, c].Style.Format = "#,#0"; //thousands separator

            ++r;
            sheet[r, c].Value = 1325.236;
            sheet[r, c].Style.Format = "#,#0.00"; //two decimal digits

            ++r;
            sheet[r, c].Value = DateTime.Today; //dates get date format set by default

            ++r;
            sheet[r, c].Value = DateTime.Today;
            sheet[r, c].Style.Format = "yyyy-mm-dd"; //custom date format

            ++r;
            sheet[r, c].Value = 0.25;
            sheet[r, c].Style.Format = "0.0%";  //percentage

            wb.Save("NumberFormat.xlsx");
        }
    }
}
