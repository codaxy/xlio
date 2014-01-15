using PetaTest;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Samples.Usage
{
    [TestFixture]
    class DefinedNames
    {
        [Test]
        public static void CreatingDefinedNames()
        {
            //creating workbook with 2 worksheets
            var workbook = new Workbook();
            var exchangeRatesSheet = new Sheet("Exchange Rates");
            workbook.Sheets.AddSheet(exchangeRatesSheet);
            var pricesSheet = new Sheet("Prices");
            workbook.Sheets.AddSheet(pricesSheet);

            //defining currencies
            exchangeRatesSheet["A1"].Value = "EUR";
            exchangeRatesSheet["A2"].Value = "GBP";
            exchangeRatesSheet["A3"].Value = "USD";

            //defining exchange rates
            exchangeRatesSheet["B1"].Value = 1.0;
            exchangeRatesSheet["B2"].Value = 0.7;
            exchangeRatesSheet["B3"].Value = 1.43;

            //defining VAT percentage
            exchangeRatesSheet["E1"].Value = "VAT";
            exchangeRatesSheet["F1"].Value = 0.17;

            //set named cell for VAT
            workbook.DefinedNames.AddCell("VAT", Cell.Parse("F1"), exchangeRatesSheet);

            //set named range for exchange rates
            Range exchangeRatesRange = Range.Parse("A1:B3");
            workbook.DefinedNames.AddRange("ExchangeRate", exchangeRatesRange, exchangeRatesSheet);

            //defining headers in Prices sheet
            pricesSheet["A1"].Value = "Product";
            pricesSheet["B1"].Value = "Price";
            pricesSheet["C1"].Value = "EUR";
            pricesSheet["D1"].Value = "USD";
            pricesSheet["E1"].Value = "GBP";

            //setting headers style
            var labelsRange = pricesSheet["A1", "E1"];
            CellStyle headerStyle = new CellStyle();
            headerStyle.Font.Bold = true;
            headerStyle.Alignment.HAlign = HorizontalAlignment.Center;
            labelsRange.SetStyle(headerStyle);

            //defining products
            pricesSheet["A2"].Value = "Xlio";
            pricesSheet["B2"].Value = 50;
            pricesSheet["A3"].Value = "Azzura";
            pricesSheet["B3"].Value = 65;

            //set formula for calculating prices in different currencies
            string formulaString = "=LOOKUP(C$1, ExchangeRate)*$B2*(1+VAT)";
            var sheetRange = pricesSheet["C2", "E3"];
            sheetRange.SetFormula(formulaString);

            workbook.Save(@"DefinedNames.xlsx");
        }

    }
}
