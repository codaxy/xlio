using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PetaTest;

namespace Codaxy.Xlio.Generic.Tests.Bugreports
{
    [TestFixture]
    public class CamplannerExportTests
    {
        [Test(Active=true)]
        public void WrongStyleLoadedTest()
        {
            Action<Sheet, String> test = (sheet, mode) =>
            {                
                Assert.AreEqual(10.0, sheet.Columns["H"].DefaultStyle.Font.Size);
                Assert.AreEqual(BorderStyle.None, sheet[9].Style.Border.Top.Style);
                Assert.AreEqual(10.0, sheet[25].Style.Font.Size);
                Assert.IsNull(sheet["H26"].Style.Font.Size);
                Assert.AreEqual(BorderStyle.Thin, sheet["B9"].Style.Border.Top.Style);
            };

            var wb1 = Workbook.ReadFile(@"files\AdvancedSimExport.xlsx");
            var sheet1 = wb1.Sheets[0];

            test(sheet1, "Read");

            wb1.Save(@"files\AdvancedSimExport-out.xlsx");
            
            var wb2 = Workbook.ReadFile(@"files\AdvancedSimExport-out.xlsx");
            var sheet2 = wb2.Sheets[0];

            test(sheet2, "Write");
        }
    }
}
