using PetaTest;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.Samples.Usage
{
    [TestFixture]
    class ImportExport
    {
        class Model
        {
            public String Name { get; set; }
            public decimal Number { get; set; }
        }
        

        [Test(Active =true)]
        public void Export() {

            var data = new[] {
                new Model { Name = " N", Number = 10.25m },
                new Model { Name = "M  D", Number = 10.25m }
            };

            var table = Generic.TableInfo<Model>.Build();

            table.Columns[1].ExportColumnFormatter = (c) =>
            {
                c.DefaultStyle.Format = "#,##0.0000";
            };

            Generic.XlioExporter.Export(data, "export1.xlsx", table);
        }
    }
}
