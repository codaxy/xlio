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
        [Test]
        public void WrongStyleLoadedTest()
        {
            var wb = Workbook.ReadFile(@"files\AdvancedSimExport.xlsx");
            var sheet = wb.Sheets[0];

            Assert.AreEqual(10.0, sheet.Columns["H"].DefaultStyle.Font.Size);
            Assert.AreEqual(BorderStyle.None, sheet[9].Style.Border.Top.Style);
            Assert.AreEqual(10.0, sheet[25].Style.Font.Size);
            Assert.IsNull(sheet["H26"].Style.Font.Size);
        }
    }
}
