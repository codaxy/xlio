using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PetaTest;
using System.IO;

namespace Codaxy.Xlio.Generic.Tests
{

    [TestFixture]
    class ImportExportTests
    {
        class C1
        {            
            public int? Int { get; set; }
            public double? Double { get; set; }
            public String String { get; set; }
            public DateTime? Time { get; set; }
            public Guid? Guid { get; set; }
            public Boolean? Bool { get; set; }
        }

        [Test]
        public void Test1()
        {
            List<C1> data = new List<C1>() {
                new C1 { Bool = true, Guid = Guid.NewGuid(), String = "Test", Double = 5.5, Int = 3, Time = DateTime.Now },
                new C1 { Bool = true }
            };

            var tmpFile = "output.xlsx";//Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
            XlioExporter.Export(data, tmpFile);
            var import = XlioImporter.Import<C1>(tmpFile);

            Assert.AreEqual(import.Count, data.Count);

            for (var i = 0; i < import.Count; i++)
            {
                Assert.AreEqual(data[i].Bool, import[i].Bool);
                Assert.AreEqual(data[i].Int, import[i].Int);
                Assert.AreEqual(data[i].String, import[i].String);                
                Assert.AreEqual(data[i].Time.HasValue, import[i].Time.HasValue);
                if (data[i].Time.HasValue)
                    Assert.IsTrue(Math.Abs((data[i].Time.Value - import[i].Time.Value).TotalMilliseconds) < 1000);                
                Assert.AreEqual(data[i].Guid, import[i].Guid);
                Assert.AreEqual(data[i].Double, import[i].Double);
            }
        }
    }
}
