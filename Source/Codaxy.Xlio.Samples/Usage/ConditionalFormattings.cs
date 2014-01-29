using PetaTest;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Samples.Usage
{
    [TestFixture]
    class ConditionalFormattings
    {

        [Test]
        public static void CreateConditionalFormatting()
        {
            var workbook = new Workbook(); //creating workbook and adding sheet
            var companiesSheet = new Sheet("CompaniesByRevenue");
            workbook.Sheets.AddSheet(companiesSheet);

            companiesSheet["A1"].Value = "Rank";
            companiesSheet["B1"].Value = "Company";
            companiesSheet["C1"].Value = "Revenue (USD billions)";
            companiesSheet["D1"].Value = "Capitalization (USD billions)";
            companiesSheet["E1"].Value = "Employees";

            CellStyle headingStyle = new CellStyle(); //adding header bold and centered header style
            headingStyle.Font.Bold = true;
            headingStyle.Alignment.Horizontal = HorizontalAlignment.Center;
            companiesSheet["A1", "E1"].SetStyle(headingStyle);

            companiesSheet.Columns[0].Width = 5;
            companiesSheet.Columns[1].Width = 30;
            companiesSheet.Columns[2].Width = 25;
            companiesSheet.Columns[3].Width = 25;
            companiesSheet.Columns[4].Width = 15;

            companiesSheet["A2"].Value = 1;
            companiesSheet["A3"].Value = 2;
            companiesSheet["A4"].Value = 3;
            companiesSheet["A5"].Value = 4;
            companiesSheet["A6"].Value = 5;

            companiesSheet["B2"].Value = "Wal-Mart Stores, Inc. (USA)";
            companiesSheet["B3"].Value = "Royal Dutch Shell (NLD)";
            companiesSheet["B4"].Value = "Exxon Mobil Corporation (USA)";
            companiesSheet["B5"].Value = "China National Petroleum (CHN)";
            companiesSheet["B6"].Value = "Sinopec Group (CHN)";

            companiesSheet["C2"].Value = 469;
            companiesSheet["C3"].Value = 467;
            companiesSheet["C4"].Value = 453;
            companiesSheet["C5"].Value = 425;
            companiesSheet["C6"].Value = 411;

            companiesSheet["D2"].Value = 248;
            companiesSheet["D3"].Value = 132;
            companiesSheet["D4"].Value = 406;
            companiesSheet["D5"].Value = null;
            companiesSheet["D6"].Value = 81;

            companiesSheet["E2"].Value = 2200000;
            companiesSheet["E3"].Value = 90000;
            companiesSheet["E4"].Value = 76900;
            companiesSheet["E5"].Value = 1668072;
            companiesSheet["E6"].Value = 401000;

            CellStyle thousandsFormatStyle = new CellStyle(); //adding thousands separator format on Employees column
            thousandsFormatStyle.Alignment.Horizontal = HorizontalAlignment.Center;
            thousandsFormatStyle.Format = "#,#0";
            companiesSheet["E2", "E6"].SetStyle(thousandsFormatStyle);

            CellStyle currencyFormatStyle = new CellStyle(); //adding $ (dollar) sign on Revenue and Capitalization columns
            currencyFormatStyle.Format = "$0";
            currencyFormatStyle.Alignment.Horizontal = HorizontalAlignment.Center;
            companiesSheet["C2", "D6"].SetStyle(currencyFormatStyle);


            ConditionalFormatting companyNameRules = new ConditionalFormatting(Range.Parse("B2:B6")); //creating ConditionalFormatting for specified range
            /*
                Adding conditional formatting rule for Company column. 
                If it contains string "USA", background color changes to red,
                for string "NLD" (Netherlands) it is blue, and for string 
                "CHN" (China) it is yellow. 
            */
            CellStyle redFillStyle = new CellStyle(); 
            CellFill redStyleFill = new CellFill();
            redStyleFill.Background = Color.Red;
            redStyleFill.Pattern = FillPattern.Solid;
            redFillStyle.Fill = redStyleFill;
            companyNameRules.AddContainsText("USA", redFillStyle); //adding ContainsText rule

            CellStyle blueFillStyle = new CellStyle();
            CellFill blueStyleFill = new CellFill();
            blueStyleFill.Background = new Color("#0096FF");
            blueStyleFill.Pattern = FillPattern.Solid;
            blueFillStyle.Fill = blueStyleFill;
            companyNameRules.AddEndsWithText("(NLD)", blueFillStyle);

            CellStyle yellowFillStyle = new CellStyle();
            CellFill yellowStyleFill = new CellFill();
            yellowStyleFill.Background = new Color("#FFFF96");
            yellowStyleFill.Pattern = FillPattern.Solid;
            yellowFillStyle.Fill = yellowStyleFill;
            companyNameRules.AddContainsText("CHN", yellowFillStyle);

            ConditionalFormatting revenueRules = new ConditionalFormatting();
            revenueRules.AddRange(Range.Parse("C2:C6")); //multiple ranges or cells can be added
            revenueRules.AddDefaultIconSet(IconSetType.Item5Arrows); //add IconSet with 5 arrows
            CellStyle underlineCellStyle = new CellStyle();
            underlineCellStyle.Font.Underline = FontUnderline.Single;
            revenueRules.AddGreaterThan("460", underlineCellStyle);

            ConditionalFormatting capitalizationRules = new ConditionalFormatting(Range.Parse("D1:D6")); //specifying range in constructor
            List<CFVO> cfvoList = new List<CFVO> // crating list of Conditional Formatting Value Objects
            {
                new CFVO
                {
                    Type = CFVOType.Min
                },
                new CFVO
                {
                    Type = CFVOType.Max
                }
            };
            List<Color> colorList = new List<Color>
            {
                Color.Yellow,
                Color.Orange
            };
            capitalizationRules.AddColorScale(cfvoList, colorList);
            capitalizationRules.AddDefaultIconSet();

            CellStyle blankStyle = new CellStyle(); //create red fill style for background
            CellFill blankStyleFill = new CellFill();
            blankStyleFill.Background = Color.Red;
            blankStyleFill.Pattern = FillPattern.Solid;
            blankStyle.Fill = blankStyleFill;

            capitalizationRules.AddIsBlank(blankStyle); //add style applied on blank cell

            ConditionalFormatting employeesRules = new ConditionalFormatting(Range.Parse("E1:E6"));
            employeesRules.AddDefaultDataBar(Color.Blue); //adding blue data bar conditional formatting rule

            //adding different conditional formattings to sheet
            companiesSheet.ConditionalFormatting.Add(companyNameRules);
            companiesSheet.ConditionalFormatting.Add(revenueRules);
            companiesSheet.ConditionalFormatting.Add(capitalizationRules);
            companiesSheet.ConditionalFormatting.Add(employeesRules);

            workbook.Save(@"ConditionalFormatting.xlsx");
        }
        


    }
}
