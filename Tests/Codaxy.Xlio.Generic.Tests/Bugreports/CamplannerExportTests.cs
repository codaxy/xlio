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

            Assert.AreEqual(10.0, sheet.Columns["H"].Style.Font.Size);
        }
    }
}
