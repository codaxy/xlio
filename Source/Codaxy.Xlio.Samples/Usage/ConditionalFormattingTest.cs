using PetaTest;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Samples.Usage
{
    [TestFixture]
    class ConditionalFormattingTest
    {
        [Test]
        public static void CreateConditionalFormatting()
        {
            /*
             //PODESAVNJE STILA ZA NEKI RANGE
             Sheet s = new Sheet("naziv");
             s[Range.Parse("A1:B2")].SetStyle(new CellStyle()); 
            */


            var workbook = new Workbook();
            var sheet = new Sheet("Sheet1");
            workbook.Sheets.AddSheet(sheet);
            
            sheet["A1"].Value = 4;
            sheet["A2"].Value = 1;
            sheet["B1"].Value = 5;
            sheet["B2"].Value = 2;

            sheet["D1"].Value = 6;//test
            sheet["D2"].Value = 1;//test
            sheet["E1"].Value = 2;//test
            sheet["E2"].Value = 3;//test

            sheet["E4"].Value = 4;//test

            sheet["A11"].Value = 6.1;//test
            sheet["A10"].Value = 2.2;//test
            sheet["B10"].Value = 3.8;//test
            sheet["B11"].Value = 4.4;//test

            sheet["A12"].Value = "simo";//test
            sheet["A13"].Value = "vaso";//test
            sheet["B12"].Value = "vidoje";//test
            sheet["B13"].Value = "gavro";//test

            ConditionalFormatting rules = new ConditionalFormatting();
            rules.AddRange(Range.Parse("A1:B2"));

            CellStyle cs = new CellStyle();
            CellFill cfill = new CellFill();
            cfill.Background = Color.Yellow;
            cs.Fill = cfill;
            cs.Font.Color = Color.Red;

            ConditionalFormattingCondition condition = ConditionalFormattingCondition.GreaterThanOrEqual("3", cs);
            rules.AddRule(condition);
            sheet.ConditionalFormatting.Add(rules);

            /*
            ConditionalFormatting rules = new ConditionalFormatting();
            rules.AddRange(Range.Parse("A1:B2"));
            rules.AddRange(Range.Parse("D1:E2"));
            rules.AddCell(Cell.Parse("E4"));

            ColorScale colorScale = new ColorScale(ConditionalFormattingColorScaleType.ThreeColorScale);
            rules.AddRule(colorScale);
            rules.AddRule(new IconSet(ConditionalFormattingTypeIconSetType.Item5Quarters));
            DataBar dataBar = new DataBar(Color.Yellow);
            rules.AddRule(dataBar);

            sheet.ConditionalFormatting.Add(rules);
            

            ConditionalFormatting r2 = new ConditionalFormatting(Range.Parse("A10:B11"));
            r2.AddRule(new DataBar(Color.Magenta));

            sheet.ConditionalFormatting.Add(r2);
            */
            
            /*

            //conditional formatting
            ConditionalFormatting cf = new ConditionalFormatting();
            cf.SequenceOfReferences.Add("A1:B2");
            cf.SequenceOfReferences.Add("D1:E2");//test
            //end of conditional formatting

            //conditional formatting rule
            ConditionalFormattingRule rule = new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.IconSet
            };
            //end of conditional formatting rule
           
            //icon set
            ConditionalFormattingIconSet iconSet = new ConditionalFormattingIconSet(ConditionalFormattingTypeIconSetType.Item4TrafficLights);
            iconSet.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.Percent,
                Value = "0"
            });
            iconSet.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.Percent,
                Value = "33"
            });
            iconSet.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.Percent,
                Value = "67"
            });
            iconSet.IconSetType = ConditionalFormattingTypeIconSetType.Item3Arrows;
            //end of icon set

            */

            /*
            //data bar
            ConditionalFormattingDataBar db = new ConditionalFormattingDataBar();
            db.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.min
            });
            db.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.max
            });
            db.Color = new Color("#00FF00");
            //end of data bar
            */

            /*
            //color scale
            ConditionalFormattingColorScale colorScale = new ConditionalFormattingColorScale();
            colorScale.CFVOList.Add(new ConditionalFormattingValueObject { 
                Type = CFVOType.min
            });
            colorScale.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.max
            });
            colorScale.Colors.Add(new Color("#FF0000"));
            colorScale.Colors.Add(new Color("#0070C0"));
            //end of color scale
            rule.ColorScale = colorScale;
            
            rule.IconSet = iconSet;
            cf.Rules.Add(rule);
            sheet.ConditionalFormatting.Add(cf);
            */

            workbook.Save(@"ConditionalFormatting.xlsx");
        }

    }
}
