using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PetaTest;

namespace Codaxy.Xlio
{
    [TestFixture]
    class StyleTests
    {
        [Test]
        public void WebColorParseTest()
        {
            Assert.AreEqual(Color.Black, new Color("#000"));
            Assert.AreEqual(new Color(0xffff0000), new Color("#f00"));
            Assert.AreEqual(new Color(0xff123456), new Color("#123456"));
        }

        [Test]
        public void TwoInstancesOfCellStyleAreEqual()
        {
            Assert.AreEqualAndHaveSameHash(new CellStyle(), new CellStyle());

            Assert.AreEqualAndHaveSameHash(new BorderEdge() { Style = BorderStyle.DashDot }, new BorderEdge() { Style = BorderStyle.DashDot });
            Assert.AreEqualAndHaveSameHash(new Color(1, 2, 3, 4), new Color(1, 2, 3, 4));
            Assert.AreEqualAndHaveSameHash(new CellFont { Name = "T", Size = 10 }, new CellFont { Name = "T", Size = 10 });            

            var c1 = new CellStyle
            {
                Alignment = new CellAlignment { HAlign = HorizontalAlignment.Center, Indent = 2 },
                Border = new CellBorder { Bottom = new BorderEdge { Color = new Color(1, 2, 3, 4), Style = BorderStyle.DashDot } },
                Fill = new CellFill { Background = new Color(2, 3, 4, 5), Pattern = FillPattern.Solid },
                Font = new CellFont { Bold = true, Name = "Tahoma" },
                Format = "{0}"
            };

            var c2 = new CellStyle
            {
                Alignment = new CellAlignment { HAlign = HorizontalAlignment.Center, Indent = 2 },
                Border = new CellBorder { Bottom = new BorderEdge { Color = new Color(1, 2, 3, 4), Style = BorderStyle.DashDot } },
                Fill = new CellFill { Background = new Color(2, 3, 4, 5), Pattern = FillPattern.Solid },
                Font = new CellFont { Bold = true, Name = "Tahoma" },
                Format = "{0}"
            };

        }

        [Test]
        public void DictionaryTest()
        {
            Dictionary<CellStyle, int> index = new Dictionary<CellStyle, int>();

            index.Add(new CellStyle(), 0);
            Assert.IsTrue(index.ContainsKey(new CellStyle()));

        }

        
    }
}
