using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PetaTest;

namespace Codaxy.Xlio.Generic.Tests
{
    [TestFixture]
    class CellTests
    {
        [Test]
        public void ColumnNameTest()
        {
            Assert.AreEqual("A", XlioUtil.GetColumnName(0));
            Assert.AreEqual("Z", XlioUtil.GetColumnName(25));
            Assert.AreEqual("AA", XlioUtil.GetColumnName(26));
            Assert.AreEqual("BA", XlioUtil.GetColumnName(52));
        }

        [Test(Active=false)]
        public void ParseCellTest()
        {
            Assert.AreEqual(0, Cell.Parse("A").Col);
            Assert.AreEqual(25, Cell.Parse("Z").Col);
            Assert.AreEqual(26, Cell.Parse("AA1").Col);
            Assert.AreEqual(52, Cell.Parse("BA1").Col);            
        }
    }
}
