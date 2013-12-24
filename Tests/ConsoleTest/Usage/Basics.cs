using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio;

namespace Codaxy.Xlio.Samples
{
    class Basics
    {
        public static void LoadExistingWorkbook()
        {
            var path = @"C:\Files\MyWorkbook.xlsx";
            var workbook = Workbook.Load(path);
            var sheet1 = workbook.Sheets[0];
            sheet1["A1"].Value = "XLIO";
            workbook.Save(path);
        }

        public static void SheetManipulation()
        {
            var workbook = new Workbook();
            workbook.DefaultFont = new CellFont { Name = "Tahoma", Size = 10 };

            var sheet1 = workbook.Sheets.AddSheet("Sheet 1");

            var sheet2 = new Sheet("Sheet2");
            workbook.Sheets.AddSheet(sheet2);

            workbook.Sheets.RemoveSheet(0);
            workbook.Sheets.RemoveSheet("Sheet 2");

            var sheet3 = workbook.Sheets.AddSheet("Sheet 3");
            sheet3.DefaultRowHeight = 10;
            sheet3.Page.Orientation = PageOrientation.Landscape;
            sheet3.ShowGridLines = true;            
        }

        public static void CellsAndRanges()
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.AddSheet("Cells and Ranges");            

            sheet["A1"].Value = "A1";
            sheet[0, 1].Value = "B1"; //sheet[row, col]
            sheet[1, 0].Value = "A2"; 
            sheet["B2"].Value = "B2"; 

            var title = sheet[2, 0, 2, 3].Merge();
            title.Value = "Title";
            title.Style.Alignment.HAlign = HorizontalAlignment.Center;

            var headers = sheet["A4", "D4"];
            headers.SetBorder(new BorderEdge { Style = BorderStyle.Thin, Color = Color.Black }); //add border to header cells

            headers[0].Value = "Header 1"; //access a cell relative to range's top-left corner
            headers[1].Value = "Header 2";
            headers[2].Value = "Header 3";
            headers[3].Value = "Header 4";

            var data = sheet["A5", "D10"];
            for (var row = 0; row < 6; row++)
            {
                var rowRange = data.GetRow(row); //get row subrange
                for (var col = 0; col < 4; col++)
                    rowRange[col].Value = String.Format("{0}:{1}", row, col);
            }

            data.GetColumn(0).ApplyStyle(new CellStyle { Fill = new CellFill { Foreground = new Color(0xffff00), Pattern = FillPattern.Solid } }); //apply yellow background to first column

            workbook.Save("CellsAndRanges.xlsx");
        }
    }
}
