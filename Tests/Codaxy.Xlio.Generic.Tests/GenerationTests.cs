using PetaTest;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Generic.Tests
{
    [TestFixture]
    class GenerationTests
    {
         [Test]
         public void BlankSheetTest()
         {
             var wb = new Workbook();
             var sheet = wb.Sheets.AddSheet(new Sheet("Sheet 1"));
             sheet.ShowGridLines = true;

             using (var f = File.Open("Blank.xlsx", FileMode.Create))
                 wb.SaveToStream(f);
         }

         [Test]
         public void DoubleBorderMergedHVC()
         {
             var wb = new Workbook();
             var sheet1 = wb.Sheets.AddSheet(new Sheet("Sheet 1") { ShowGridLines = true });
             var sheet2 = wb.Sheets.AddSheet(new Sheet("Sheet 2") { ShowGridLines = true });
             var sheet3 = wb.Sheets.AddSheet(new Sheet("Sheet 3") { ShowGridLines = true });
             var range = sheet3[0, 0, 10, 10];
             range.SetOutsideBorder(new BorderEdge { Color = Color.Black, Style = BorderStyle.Double });
             var area = range.Merge();
             area.Value = "Hello";
             area.Style.Alignment.VAlign = VerticalAlignment.Center;
             area.Style.Alignment.HAlign = HorizontalAlignment.Center;

             using (var f = File.Open("DoubleBorderMergedHVC.xlsx", FileMode.Create))
                 wb.SaveToStream(f);
         }

         [Test(Active=false)]
         public void RowHeightTest()
         {
             var wb = new Workbook();
             var sheet1 = wb.Sheets.AddSheet(new Sheet("Sheet 1") { ShowGridLines = true });

             sheet1[0].Height = 200;

             using (var f = File.Open("RowHeightTest.xlsx", FileMode.Create))
                 wb.SaveToStream(f);
         }

         [Test(Active = true)]
         public void DefaultFont()
         {
             var wb = new Workbook();
             wb.DefaultFont = new CellFont { Name = "Arial", Size = 20 };

             var sheet1 = wb.Sheets.AddSheet(new Sheet("Sheet 1") { ShowGridLines = true });

             sheet1["A1"].Value = "Arial 20";
             sheet1["A2"].Value = "Bold Arial 20";
             sheet1["A2"].Style.Font.Bold = true;

             using (var f = File.Open("DefaultFont.xlsx", FileMode.Create))
                 wb.SaveToStream(f);
         }

        [Test]
         public void ProblematicTest()
         {
             var PhaseWidth = 13;
             var phaseStartCol = 0;
             var VOffset = 10;
             var TableHeight = 10;
             var wb = new Workbook();
             var sheet = wb.Sheets.AddSheet(new Sheet("Sheet") { ShowGridLines = true });

             for (var i = 0; i < 3; i++)
             {
                 var phaseArea = sheet[0, phaseStartCol, 100, phaseStartCol + PhaseWidth - 1];
                 phaseArea.SetOutsideBorder(new BorderEdge { Style = BorderStyle.Double, Color = Color.Black });

                 var title = phaseArea[0, 3, 2, PhaseWidth - 4].Merge();

                 title.Value = "Costs Budget";
                 title.Style.Alignment.HAlign = HorizontalAlignment.Center;
                 title.Style.Alignment.VAlign = VerticalAlignment.Center;
                 title.Style.Font.Size = 16;
                 title.Style.Font.Bold = true;

                 var tableArea = phaseArea[VOffset, 0, VOffset + TableHeight, PhaseWidth - 1];
                 var square = tableArea[0, 0, 4, 1].Merge();
                 var header = tableArea[0, 2, 0, PhaseWidth - 1].Merge();
                 header.Value = "Something";

                 //var iCosts = tableArea[1, 2, 1, 4].Merge();
                 //iCosts.Value = "Incurred Costs";

                 phaseStartCol += PhaseWidth;
             }

             using (var f = File.Open("ProblematicTest.xlsx", FileMode.Create))
                 wb.SaveToStream(f);

         }
    }
}
