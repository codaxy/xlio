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
            var workbook = new Workbook();
            var sheet = new Sheet("Sheet1");
            workbook.Sheets.AddSheet(sheet);
            sheet["A1"].Value = 4;
            sheet["A2"].Value = 1;
            sheet["B1"].Value = 5;
            sheet["B2"].Value = 2;

            //conditional formatting
            ConditionalFormatting cf = new ConditionalFormatting();
            cf.SequenceOfReferences.Add("A1:B2");
            //end of conditional formatting

            //conditional formatting rule
            ConditionalFormattingRule rule = new ConditionalFormattingRule
            {
                Type = ConditionalFormattingType.iconSet
            };
            //end of conditional formatting rule

            //icon set
            ConditionalFormattingIconSet iconSet = new ConditionalFormattingIconSet();
            iconSet.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.percent,
                Value = "0"
            });
            iconSet.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.percent,
                Value = "33"
            });
            iconSet.CFVOList.Add(new ConditionalFormattingValueObject
            {
                Type = CFVOType.percent,
                Value = "67"
            });
            iconSet.IconSetType = ConditionalFormattingTypeIconSetType.Item3Arrows;
            //end of icon set

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
             */

            rule.IconSet = iconSet;
            cf.Rules.Add(rule);
            sheet.ConditionalFormatting.Add(cf);

            workbook.Save(@"ConditionalFormatting.xlsx");
        }

    }
}
