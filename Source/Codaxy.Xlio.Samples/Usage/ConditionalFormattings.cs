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
        [Test(Active=false)]
        public static void CreateConditionalFormatting()
        {
            var workbook = new Workbook(); //creating workbook and adding sheet
            var companiesSheet = new Sheet("CompaniesByRevenue");
            workbook.Sheets.AddSheet(companiesSheet);

            companiesSheet.Columns[0].Width = 5;
            companiesSheet.Columns[1].Width = 30;
            companiesSheet.Columns[2].Width = 25;
            companiesSheet.Columns[3].Width = 25;
            companiesSheet.Columns[4].Width = 15;

            var headingStyle = new CellStyle(); //define bold and centered header style
            headingStyle.Font.Bold = true;
            headingStyle.Alignment.Horizontal = HorizontalAlignment.Center;

            var thousandsFormatStyle = new CellStyle(); //adding thousands separator format on Employees column
            thousandsFormatStyle.Alignment.Horizontal = HorizontalAlignment.Center;
            thousandsFormatStyle.Format = "#,#0";

            var currencyFormatStyle = new CellStyle(); //adding $ (dollar) sign on Revenue and Capitalization columns
            currencyFormatStyle.Format = "$#,#0";
            currencyFormatStyle.Alignment.Horizontal = HorizontalAlignment.Center;

            var table = companiesSheet["A2", "E6"];
            table.GetRow(-1).SetValues("Rank", "Company", "Revenue (USD billions)", "Capitalization (USD billions)", "Employees").SetStyle(headingStyle);
            table.GetColumn(0).SetValues(1, 2, 3, 4, 5);
            table.GetColumn(1).SetValues("Wal-Mart Stores, Inc. (USA)", "Royal Dutch Shell (NLD)", "Exxon Mobil Corporation (USA)", "China National Petroleum (CHN)", "Sinopec Group (CHN)");
            table.GetColumn(2).SetValues(469, 467, 453, 425, 411).SetStyle(currencyFormatStyle);
            table.GetColumn(3).SetValues(248, 132, 406, null, 81).SetStyle(currencyFormatStyle);
            table.GetColumn(4).SetValues(2200000, 90000, 76900, 1668072, 401000).SetStyle(thousandsFormatStyle);
            
            var originFormating = new ConditionalFormatting(Range.Parse("B2:B6")); //creating ConditionalFormatting for specified range
            originFormating.AddContainsText("USA", new CellStyle { Fill = CellFill.BackColor(Color.Red) });
            originFormating.AddEndsWithText("(NLD)", new CellStyle { Fill = CellFill.BackColor(new Color("#0096FF")) });
            originFormating.AddContainsText("CHN", new CellStyle { Fill = CellFill.BackColor(new Color("#FFFF96")) });

            var revenueFormatting = new ConditionalFormatting();
            revenueFormatting.AddRange("C2:C6"); //multiple ranges or cells can be added
            revenueFormatting.AddDefaultIconSet(IconSetType.Item5Arrows); //add IconSet with 5 arrows            
            revenueFormatting.AddGreaterThan("460", new CellStyle { Font = new CellFont { Underline = FontUnderline.Single } });

            var capitalizationFormatting = new ConditionalFormatting(Range.Parse("D1:D6")); //specifying range in constructor            
            capitalizationFormatting.AddColorScale(Color.Yellow, Color.Orange);
            capitalizationFormatting.AddDefaultIconSet();
            capitalizationFormatting.AddIsBlank(new CellStyle { Fill = CellFill.BackColor(Color.Red) }); //add style applied on blank cell

            ConditionalFormatting employeesNumberFormating = new ConditionalFormatting(Range.Parse("E1:E6"));
            employeesNumberFormating.AddDefaultDataBar(Color.Blue); //adding blue data bar conditional formatting rule

            //adding different conditional formattings to sheet
            companiesSheet.ConditionalFormatting.Add(originFormating);
            companiesSheet.ConditionalFormatting.Add(revenueFormatting);
            companiesSheet.ConditionalFormatting.Add(capitalizationFormatting);
            companiesSheet.ConditionalFormatting.Add(employeesNumberFormating);

            workbook.Save(@"ConditionalFormatting.xlsx");
        }
    }
}
